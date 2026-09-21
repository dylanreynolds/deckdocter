<#
.SYNOPSIS
  Replace embedded videos with linked online videos that play inside the slide,
  by driving PowerPoint itself rather than editing the XML.

.DESCRIPTION
  Hand-editing the video XML does not work. Three different structures were
  built by hand - poster+hyperlink, linked videoFile, and the p14:media link
  form - and PowerPoint rejected all three with "PowerPoint found a problem with
  content", even though python-pptx, LibreOffice and a full relationship audit
  all passed them. PowerPoint's tolerance for media shapes is narrower than the
  schema, and it is not documented.

  So this does not write XML. It opens the deck in PowerPoint, deletes each
  video shape, and calls Shapes.AddMediaObject2 with the URL and LinkToFile,
  which is the same code path Insert > Video > This Device (linked) uses. The
  file that comes out was written by PowerPoint, so it opens in PowerPoint.

  Run this AFTER Optimize-Deck.ps1 has done the image work. That stage is
  XML-based, fast, and verified - only the video step needs PowerPoint.

.PARAMETER Path
  The deck to convert - normally the "(repaired).pptx" from Optimize-Deck.ps1.

.PARAMETER VideoLinks
  CSV with Media,Url columns, as produced by -Report and filled by
  Get-StreamLinks.ps1.

.PARAMETER OutFile
  Where to write the result. Defaults to "<name> (inline video).pptx" alongside.

.PARAMETER KeepPosterFrame
  Re-apply each original video's poster image to the new shape, so the slide
  looks unchanged before playback. Without this PowerPoint shows its own first
  frame, which for a streamed file can be a black rectangle until it loads.

.EXAMPLE
  .\Set-InlineVideo.ps1 -Path '.\repaired\deck (repaired).pptx' -VideoLinks '.\links.csv'

.NOTES
  Needs PowerPoint installed. It runs hidden, but do not close PowerPoint while
  it works. The input file is not modified.
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory, Position=0)][string]$Path,
  [Parameter(Mandatory)][string]$VideoLinks,
  [string]$OutFile,
  [switch]$KeepPosterFrame
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'DeckDoctor.psm1') -Force

$deck = (Resolve-Path -LiteralPath $Path).Path
$csv  = (Resolve-Path -LiteralPath $VideoLinks).Path
if (-not $OutFile) {
  $OutFile = Join-Path (Split-Path $deck -Parent) ((Split-Path $deck -LeafBase) + ' (inline video).pptx')
}
$OutFile = [System.IO.Path]::GetFullPath($OutFile)

$urls = @{}
foreach ($r in (Import-Csv -LiteralPath $csv)) {
  if ($r.Url -and $r.Url.Trim() -match '^https?://') { $urls[(Split-Path $r.Media -Leaf)] = $r.Url.Trim() }
}
if (-not $urls.Count) { throw "No usable URLs in $csv - fill the Url column first (Get-StreamLinks.ps1)." }
Write-Host "  $($urls.Count) link(s) loaded" -ForegroundColor Gray

# Map slide -> shape name -> media file, read from the XML. Matching on the
# shape's own name is what keeps two videos on one slide from being swapped;
# relying on shape order silently mixes them up.
$map = Get-VideoShapeMap $deck
if (-not $map.Count) { throw 'No embedded video shapes found in that deck. Has the video already been removed?' }
Write-Host "  $($map.Count) video shape(s) found" -ForegroundColor Gray

$posters = @{}
if ($KeepPosterFrame) {
  $posterDir = Join-Path $env:TEMP ('deckdoctor-posters-' + [guid]::NewGuid().ToString('N').Substring(0,8))
  New-Item -ItemType Directory -Force -Path $posterDir | Out-Null
  $posters = Export-VideoPosters $deck $posterDir
  Write-Host "  $($posters.Count) poster frame(s) exported" -ForegroundColor Gray
}

Write-Host '  Opening in PowerPoint...' -ForegroundColor Gray
$pp = New-Object -ComObject PowerPoint.Application
$pp.DisplayAlerts = 1                     # ppAlertsNone: a bad file throws rather than prompting
$pres = $null
$done = 0; $missing = @()
try {
  $pres = $pp.Presentations.Open($deck, $false, $false, $false)   # not read-only, no window

  foreach ($entry in ($map | Sort-Object Slide)) {
    $url = $null
    if ($urls.ContainsKey($entry.Media)) { $url = $urls[$entry.Media] }
    if (-not $url) { $missing += "$($entry.Media) (slide $($entry.Slide))"; continue }
    if ($entry.Slide -gt $pres.Slides.Count) { $missing += "slide $($entry.Slide) does not exist"; continue }

    $slide = $pres.Slides.Item($entry.Slide)
    $shape = $null
    foreach ($s in $slide.Shapes) { if ($s.Name -eq $entry.Name) { $shape = $s; break } }
    if (-not $shape) {
      foreach ($s in $slide.Shapes) { if ($s.Type -eq 16) { $shape = $s; break } }   # fall back to first media shape
    }
    if (-not $shape) { $missing += "$($entry.Name) not found on slide $($entry.Slide)"; continue }

    $L = $shape.Left; $T = $shape.Top; $W = $shape.Width; $H = $shape.Height
    $Z = $shape.ZOrderPosition
    $rot = $shape.Rotation
    $lockAR = $shape.LockAspectRatio
    $cropL = 0; $cropR = 0; $cropT = 0; $cropB = 0
    try {
      $pf = $shape.PictureFormat
      $cropL = $pf.CropLeft; $cropR = $pf.CropRight; $cropT = $pf.CropTop; $cropB = $pf.CropBottom
    } catch { }
    $shape.Delete()

    # LinkToFile = true, SaveWithDocument = false -> the bytes stay in SharePoint
    $new = $slide.Shapes.AddMediaObject2($url, $true, $false, $L, $T, $W, $H)
    $new.Name = $entry.Name

    # Sizing has to come AFTER aspect lock is released. A new media shape is
    # created with LockAspectRatio on, so assigning height silently recomputes
    # width from the video's native ratio: a 151 x 268.4 pt portrait tile came
    # back 477.2 pt wide (268.4 x 16/9) and destroyed the slide layout.
    $new.LockAspectRatio = 0                 # msoFalse
    if ($cropL -or $cropR -or $cropT -or $cropB) {
      try {
        $pf2 = $new.PictureFormat
        $pf2.CropLeft = $cropL; $pf2.CropRight = $cropR
        $pf2.CropTop  = $cropT; $pf2.CropBottom = $cropB
      } catch { }
    }
    $new.Width = $W; $new.Height = $H         # width first, then height
    $new.Left  = $L; $new.Top    = $T         # cropping can shift origin, so set position last
    if ($rot) { $new.Rotation = $rot }
    $new.LockAspectRatio = $lockAR            # leave the shape as the original was

    if ($KeepPosterFrame -and $posters.ContainsKey($entry.Media)) {
      try { $new.MediaFormat.SetDisplayPicture($posters[$entry.Media]) } catch { }
    }
    # restore stacking so the video sits where the original did
    while ($new.ZOrderPosition -gt $Z) { $new.ZOrder(3) }   # msoSendBackwardOne
    while ($new.ZOrderPosition -lt $Z) { $new.ZOrder(2) }   # msoBringForwardOne

    $done++
    Write-Host ("  slide {0,-3} {1}" -f $entry.Slide, $entry.Media) -ForegroundColor Gray
  }

  if (Test-Path -LiteralPath $OutFile) { Remove-Item -LiteralPath $OutFile -Force }
  $pres.SaveAs($OutFile)
  $pres.Close(); $pres = $null
} finally {
  if ($pres) { try { $pres.Close() } catch { } }
  try { $pp.Quit() } catch { }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
}

# Reopen with PowerPoint as the final word on whether the file is good.
Write-Host '  Verifying with PowerPoint...' -ForegroundColor Gray
Start-Sleep -Seconds 2
$ok = $false; $mediaCount = 0
$pp2 = New-Object -ComObject PowerPoint.Application
$pp2.DisplayAlerts = 1
try {
  $d2 = $pp2.Presentations.Open($OutFile, $true, $false, $false)
  foreach ($s in $d2.Slides) { foreach ($sh in $s.Shapes) { if ($sh.Type -eq 16) { $mediaCount++ } } }
  $ok = $true
  $d2.Close()
} catch {
  Write-Host "  PowerPoint REJECTED the result: $($_.Exception.Message)" -ForegroundColor Red
} finally {
  try { $pp2.Quit() } catch { }
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp2) | Out-Null
}

Write-Host ''
if ($missing.Count) {
  Write-Host "  no link for: $($missing -join '; ')" -ForegroundColor Yellow
}
if ($ok) {
  $a = (Get-Item $deck).Length; $b = (Get-Item $OutFile).Length
  Write-Host ("  {0} video(s) converted   {1} -> {2}" -f $done, (Format-Size $a), (Format-Size $b)) -ForegroundColor Green
  Write-Host "  $mediaCount media shape(s) in the result" -ForegroundColor Green
  Write-Host "  $OutFile" -ForegroundColor Green
  Write-Host '  Opened and re-read by PowerPoint itself - no repair prompt.' -ForegroundColor Gray
} else {
  Write-Host '  The result did not survive verification. Nothing to trust here.' -ForegroundColor Red
}
Write-Host ''
