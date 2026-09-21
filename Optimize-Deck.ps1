<#
.SYNOPSIS
  Diagnose and shrink bloated PowerPoint decks. Streams entries one at a time,
  so it handles decks far larger than a browser can hold.

.DESCRIPTION
  A .pptx is a ZIP of XML parts plus a media folder. Bloat is almost always
  media: images stored at far higher resolution than they display, PNGs holding
  photographs, orphaned media left behind by deleted slides, and embedded video.

  Three modes:

    -Report                Diagnose only. Writes nothing except an optional
                           CSV template for video links.
    (default)              Re-encode images, drop orphans. Video left embedded.
    -VideoLinks <csv>      Additionally hand the video to PowerPoint, which
                           relinks it as online media that plays in the slide.
                           Requires PowerPoint on the machine.

  The original file is never modified.

.PARAMETER Path
  A .pptx file, or a folder to scan for them.

.PARAMETER OutDir
  Where repaired copies go. Defaults to a "repaired" folder beside the input.

.PARAMETER MaxDim
  Longest edge kept, in pixels. 1920 conservative, 1600 balanced, 1280 aggressive.

.PARAMETER Quality
  JPEG quality 1-100. 85 conservative, 80 balanced, 75 aggressive.

.PARAMETER ExtractVideosTo
  Save embedded video out to this folder. Point it at a OneDrive-synced folder
  and sync performs the upload, which is what puts the files in Stream.

.PARAMETER VideoLinks
  CSV with columns Media,Url. Generate the template with -Report.

.EXAMPLE
  .\Optimize-Deck.ps1 -Path '.\deck.pptx' -Report

.EXAMPLE
  .\Optimize-Deck.ps1 -Path '.\deck.pptx' -ExtractVideosTo '.\videos-for-stream'

.EXAMPLE
  .\Optimize-Deck.ps1 -Path '.\deck.pptx' -VideoLinks '.\deck.video-links.csv'

.NOTES
  Verified against a real 311 MB / 93-slide deck: 311 MB -> 35 MB.
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory, Position=0)][string]$Path,
  [string]$OutDir,
  [ValidateRange(320,4096)][int]$MaxDim = 1600,
  [ValidateRange(1,100)][int]$Quality = 80,
  [switch]$Report,
  [string]$ExtractVideosTo,
  [string]$VideoLinks
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'DeckDoctor.psm1') -Force

# ========================================================== main ==============

$targets = @()
if (Test-Path $Path -PathType Container) {
  $targets = @(Get-ChildItem -LiteralPath $Path -Filter *.pptx -File | Select-Object -ExpandProperty FullName)
} else {
  $targets = @((Resolve-Path -LiteralPath $Path).Path)
}
if (-not $targets.Count) { Write-Error "No .pptx found at $Path"; return }

$urls = $null
if ($VideoLinks) {
  $urls = @{}
  foreach ($row in (Import-Csv -LiteralPath $VideoLinks)) {
    if ($row.Url -and $row.Url.Trim() -match '^https?://') { $urls[$row.Media.Trim()] = $row.Url.Trim() }
  }
  Write-Host ("  Loaded {0} video link(s)" -f $urls.Count) -ForegroundColor Gray
}

foreach ($file in $targets) {
  $model = Get-DeckModel $file
  Show-Report $model

  $vids = @($model.Items | Where-Object { $_.Kind -eq 'video' -and $_.Referenced })

  if ($ExtractVideosTo) {
    $ExtractVideosTo = (New-Item -ItemType Directory -Force -Path $ExtractVideosTo).FullName
    $zip = [System.IO.Compression.ZipFile]::OpenRead($file)
    try {
      foreach ($v in $vids) {
        $vsl = @($v.Slides)
        $sn = if ($vsl.Count) { $vsl[0] } else { 0 }
        $name = ('slide{0:d2}-{1}' -f $sn, (Split-Path $v.Part -Leaf))
        $dest = Join-Path $ExtractVideosTo $name
        $e = $zip.GetEntry($v.Part)
        $ins = $e.Open(); $outs = [System.IO.File]::Create($dest)
        try { $ins.CopyTo($outs, 1MB) } finally { $outs.Dispose(); $ins.Dispose() }
        Write-Host ("  saved {0}  {1}" -f $name, (Format-Size $v.Bytes)) -ForegroundColor Gray
      }
    } finally { $zip.Dispose() }
    Write-Host "  Videos written to $ExtractVideosTo - wait for OneDrive to sync, then copy share links." -ForegroundColor Green
  }

  if ($Report) {
    if ($vids.Count) {
      $tpl = [System.IO.Path]::ChangeExtension($file, $null).TrimEnd('.') + '.video-links.csv'
      $vids | Sort-Object Bytes -Descending | ForEach-Object {
        [pscustomobject]@{ Media=$_.Part; SizeMB=[math]::Round($_.Bytes/1MB,1)
                           Slide=(($_.Slides) -join ','); Url='' }
      } | Export-Csv -LiteralPath $tpl -NoTypeInformation
      Write-Host "  Video link template: $tpl" -ForegroundColor Green
      Write-Host "  Fill the Url column, then re-run with -VideoLinks" -ForegroundColor Gray
    }
    continue
  }

  $od = if ($OutDir) { $OutDir } else { Join-Path (Split-Path $file -Parent) 'repaired' }
  # .FullName is load-bearing: the .NET ZipFile APIs resolve relative paths
  # against the PROCESS working directory, which is not PowerShell's current
  # location. Everything handed to .NET below must be absolute.
  $od = (New-Item -ItemType Directory -Force -Path $od).FullName
  $out = Join-Path $od ((Split-Path $file -LeafBase) + ' (repaired).pptx')

  # Video first, via PowerPoint; images second, via XML.
  #
  # Hand-written video XML does not survive PowerPoint, whatever the schema
  # says - three different structures were rejected while python-pptx,
  # LibreOffice and a full relationship audit all passed them. So the video
  # step is delegated to PowerPoint through COM (Set-InlineVideo.ps1).
  #
  # It must run BEFORE the image pass: PowerPoint's save re-inflates images
  # that have already been compressed (15.7 MB of PNG measured back up to
  # 39.5 MB), so the other order throws the image work away.
  # captured before $model is swapped for the temp deck, so the report still
  # compares against the file the user actually started with
  $origSize = $model.Size
  $source = $file
  $tempDeck = $null
  $preNotes = @()
  if ($urls -and $urls.Count) {
    Write-Host '  Converting video via PowerPoint...' -ForegroundColor Gray
    $tempDeck = Join-Path $env:TEMP ('deckdoctor-inline-' + [guid]::NewGuid().ToString('N').Substring(0,8) + '.pptx')
    & (Join-Path $PSScriptRoot 'Set-InlineVideo.ps1') -Path $file -VideoLinks $VideoLinks -OutFile $tempDeck | Out-Null
    if (-not (Test-Path -LiteralPath $tempDeck)) {
      Write-Host '  Video conversion produced nothing - is PowerPoint installed? Continuing with images only.' -ForegroundColor Yellow
      $tempDeck = $null
    } else {
      $source = $tempDeck
      $preNotes += 'Video converted to linked online media by PowerPoint.'
      $model = Get-DeckModel $source
    }
  }

  # urls are deliberately NOT passed on - the XML video path is never used
  $notes = @($preNotes) + @(Repair-Deck $model $out @{} $MaxDim $Quality)
  if ($tempDeck) { Remove-Item -LiteralPath $tempDeck -Force -ErrorAction SilentlyContinue }

  # @() again: a function returning an empty collection yields $null in
  # PowerShell, and .Count on $null throws under Set-StrictMode Latest.
  $bad = @(Test-Package $out)
  if ($bad.Count) {
    Remove-Item $out -Force
    Write-Host ''
    Write-Host '  VERIFICATION FAILED - nothing was produced.' -ForegroundColor Red
    $bad | Select-Object -First 5 | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
    Write-Host '  PowerPoint would show the repair prompt on that file. Original untouched.' -ForegroundColor Red
    continue
  }

  $newSize = (Get-Item $out).Length
  Write-Host ''
  foreach ($n in ($notes | Select-Object -Unique)) { Write-Host "  note: $n" -ForegroundColor Yellow }
  Write-Host ('  {0} -> {1}   ({2}% smaller)' -f (Format-Size $origSize), (Format-Size $newSize),
              [int](100*($origSize-$newSize)/$origSize)) -ForegroundColor Green
  Write-Host "  $out" -ForegroundColor Green
  Write-Host '  Verified: every relationship resolves. Open it and click through before sending.' -ForegroundColor Gray
  Write-Host ''
}
