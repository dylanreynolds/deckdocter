<#
  DeckDoctor engine. Shared by Optimize-Deck.ps1 (command line) and
  Deck-Doctor-GUI.ps1 (window), so there is exactly one implementation of the
  logic that rewrites a .pptx - and one place where it has been verified.
#>
Set-StrictMode -Version Latest

$script:__added = @()

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
# GDI+ has no PNG compression control and writes noticeably larger files than
# WPF's PngBitmapEncoder - measured 30-40% larger on this deck's screenshots.
# WPF is optional: if it will not load, PNG falls back to GDI+.
$script:HasWpf = $true
try {
  Add-Type -AssemblyName PresentationCore
  Add-Type -AssemblyName PresentationFramework
  Add-Type -AssemblyName WindowsBase
} catch { $script:HasWpf = $false }

# A per-pixel scan in PowerShell is far too slow for hundreds of images, so the
# alpha test is compiled. It answers one question: does the alpha channel carry
# information, or is it 255 everywhere (in which case the PNG is a photograph
# wearing a transparency channel and belongs in JPEG).
if (-not ('DeckDoctorNative' -as [type])) {
  Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class DeckDoctorNative {
    public static bool AlphaUsed(byte[] b) {
        for (int i = 3; i < b.Length; i += 4) if (b[i] != 255) return true;
        return false;
    }
    // Bitmap.GetHbitmap() hands back a GDI object that the GC does not own.
    // Across hundreds of images an undeleted handle exhausts the GDI pool and
    // image creation starts failing, so every handle is released explicitly.
    [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr o);
    public static bool DeleteObjectSafe(IntPtr o) {
        if (o == IntPtr.Zero) return false;
        return DeleteObject(o);
    }
}
'@
}

$VIDEO_EXT  = @('.mp4','.mov','.m4v','.avi','.wmv','.mkv','.webm','.mpg','.mpeg')
$RASTER_EXT = @('.png','.jpg','.jpeg','.bmp','.gif','.tif','.tiff')
$MEDIA_EXT_URI = '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}'
$VIDEO_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video'
$HLINK_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'

function Format-Size([long]$b) {
  if ($b -ge 1GB) { return ('{0:N2} GB' -f ($b/1GB)) }
  if ($b -ge 1MB) { return ('{0:N1} MB' -f ($b/1MB)) }
  if ($b -ge 1KB) { return ('{0:N0} KB' -f ($b/1KB)) }
  return "$b B"
}

# Resolve a relationship Target against the folder holding the .rels file.
function Resolve-Part([string]$baseDir, [string]$target) {
  $parts = New-Object System.Collections.Generic.List[string]
  foreach ($seg in (($baseDir + '/' + $target) -split '/')) {
    if ($seg -eq '' -or $seg -eq '.') { continue }
    if ($seg -eq '..') { if ($parts.Count) { $parts.RemoveAt($parts.Count-1) } ; continue }
    $parts.Add($seg)
  }
  return ($parts -join '/')
}

function Get-RelAttrs([string]$tag) {
  $h = @{}
  foreach ($m in [regex]::Matches($tag, '([A-Za-z:]+)="([^"]*)"')) {
    $h[$m.Groups[1].Value] = $m.Groups[2].Value
  }
  return $h
}

function Read-EntryText($zip, [string]$name) {
  $e = $zip.GetEntry($name); if (-not $e) { return $null }
  $s = $e.Open(); try { $r = New-Object System.IO.StreamReader($s); $r.ReadToEnd() } finally { $s.Dispose() }
}

function Read-EntryBytes($zip, [string]$name) {
  $e = $zip.GetEntry($name)
  $s = $e.Open()
  try {
    $ms = New-Object System.IO.MemoryStream
    $s.CopyTo($ms)
    return $ms.ToArray()
  } finally { $s.Dispose() }
}

# ---------------------------------------------------------------- analysis ----
function Get-DeckModel([string]$file) {
  $zip = [System.IO.Compression.ZipFile]::OpenRead($file)
  try {
    $referenced = New-Object System.Collections.Generic.HashSet[string]
    $slidesOf   = @{}      # media part -> hashset of slide numbers
    $videoRels  = @{}       # slide rels path -> list of @{Id;Target;Abs;Bogus}

    foreach ($e in $zip.Entries) {
      if ($e.FullName -notlike '*.rels') { continue }
      $baseDir = Split-Path (Split-Path $e.FullName -Parent) -Parent
      $baseDir = $baseDir -replace '\\','/'
      $xml = Read-EntryText $zip $e.FullName
      $slideNo = $null
      if ($e.FullName -match 'ppt/slides/_rels/slide(\d+)\.xml\.rels$') { $slideNo = [int]$Matches[1] }

      foreach ($m in [regex]::Matches($xml, '<Relationship[^>]*/>')) {
        $a = Get-RelAttrs $m.Value
        $tgt = if ($a.ContainsKey('Target')) { $a['Target'] } else { '' }
        $typ = if ($a.ContainsKey('Type'))   { $a['Type'] }   else { '' }
        $mode= if ($a.ContainsKey('TargetMode')) { $a['TargetMode'] } else { '' }
        if ($tgt -match '^(https?:|#|mailto:)') { continue }
        if ($mode -eq 'External') {
          # A video rel with Target="NULL" is a genuine defect some PowerPoint
          # versions write. The media is then only reachable via p14:media.
          if ($typ -match '2006/relationships/video' -and $slideNo) {
            if (-not $videoRels.ContainsKey($e.FullName)) { $videoRels[$e.FullName] = @() }
            $videoRels[$e.FullName] += @{ Id=$a['Id']; Target=$tgt; Abs=$null; Bogus=$true }
          }
          continue
        }
        $abs = Resolve-Part $baseDir $tgt
        [void]$referenced.Add($abs)
        if ($slideNo) {
          if (-not $slidesOf.ContainsKey($abs)) { $slidesOf[$abs] = New-Object System.Collections.Generic.HashSet[int] }
          [void]$slidesOf[$abs].Add($slideNo)
        }
        if ($typ -match '2006/relationships/video' -or $typ -match '2007/relationships/media') {
          if (-not $videoRels.ContainsKey($e.FullName)) { $videoRels[$e.FullName] = @() }
          $videoRels[$e.FullName] += @{ Id=$a['Id']; Target=$tgt; Abs=$abs; Bogus=$false }
        }
      }
    }

    $items = @()
    foreach ($e in $zip.Entries) {
      if ($e.FullName -notlike 'ppt/media/*') { continue }
      $ext = [System.IO.Path]::GetExtension($e.FullName).ToLower()
      $kind = if ($VIDEO_EXT -contains $ext) { 'video' } elseif ($RASTER_EXT -contains $ext) { 'raster' } else { 'other' }
      # @() is load-bearing: Sort-Object returns a bare Int32 for a single-slide
      # video, and Set-StrictMode Latest throws on .Count against a scalar.
      $sl = if ($slidesOf.ContainsKey($e.FullName)) { @($slidesOf[$e.FullName] | Sort-Object) } else { @() }
      $items += [pscustomobject]@{
        Part = $e.FullName; Ext = $ext; Kind = $kind
        Bytes = $e.CompressedLength
        Referenced = $referenced.Contains($e.FullName)
        Slides = $sl
      }
    }

    $slideCount = ($zip.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide\d+\.xml$' }).Count
    return [pscustomobject]@{
      File = $file; Size = (Get-Item $file).Length; Slides = $slideCount
      Items = $items; VideoRels = $videoRels
      HasThumb = [bool]$zip.GetEntry('docProps/thumbnail.jpeg')
    }
  } finally { $zip.Dispose() }
}

# ------------------------------------------------------------ image re-encode --
# Returns @{Bytes; Ext; W; H; OW; OH} or $null to keep the original untouched.
function Convert-Raster([byte[]]$bytes, [string]$ext, [int]$maxDim, [int]$quality) {
  $ms = New-Object System.IO.MemoryStream($bytes, $false)
  $img = $null; $bmp = $null; $g = $null
  try {
    try { $img = [System.Drawing.Image]::FromStream($ms) } catch { return $null }  # EMF/WMF etc
    $ow = $img.Width; $oh = $img.Height
    $scale = if ([Math]::Max($ow,$oh) -gt $maxDim) { $maxDim / [Math]::Max($ow,$oh) } else { 1.0 }
    $nw = [Math]::Max(1, [int][Math]::Round($ow * $scale))
    $nh = [Math]::Max(1, [int][Math]::Round($oh * $scale))

    # Alpha is tested on the ORIGINAL, not the resized copy. Resizing with
    # bicubic interpolation onto a transparent bitmap bleeds partial alpha into
    # the border pixels, so a fully opaque source comes back looking as though
    # it uses transparency - which wrongly keeps photographs as PNG. Measured:
    # a 4.1 MB opaque PNG stayed a 3.2 MB PNG instead of becoming a 120 KB JPEG.
    $alphaUsed = $false
    if ([System.Drawing.Image]::IsAlphaPixelFormat($img.PixelFormat)) {
      $probe = New-Object System.Drawing.Bitmap($img)
      try {
        $rect = New-Object System.Drawing.Rectangle(0,0,$probe.Width,$probe.Height)
        $data = $probe.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly,
                                [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        try {
          $len = [Math]::Abs($data.Stride) * $probe.Height
          $buf = New-Object byte[] $len
          [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $buf, 0, $len)
          $alphaUsed = [DeckDoctorNative]::AlphaUsed($buf)
        } finally { $probe.UnlockBits($data) }
      } finally { $probe.Dispose() }
    }

    $bmp = New-Object System.Drawing.Bitmap($nw, $nh, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode  = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    if (-not $alphaUsed) { $g.Clear([System.Drawing.Color]::White) }
    $g.DrawImage($img, 0, 0, $nw, $nh)

    $out = New-Object System.IO.MemoryStream
    if ($alphaUsed) {
      if ($script:HasWpf) {
        $hb = $bmp.GetHbitmap()
        try {
          $bsrc = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHBitmap(
                    $hb, [IntPtr]::Zero, [System.Windows.Int32Rect]::Empty,
                    [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions())
          $penc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
          $penc.Interlace = [System.Windows.Media.Imaging.PngInterlaceOption]::Off
          $penc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bsrc))
          $penc.Save($out)
        } finally {
          [void][DeckDoctorNative]::DeleteObjectSafe($hb)
        }
      } else {
        $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
      }
      $newExt = '.png'
    } else {
      $flat = New-Object System.Drawing.Bitmap($nw, $nh, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
      $fg = [System.Drawing.Graphics]::FromImage($flat)
      try {
        $fg.Clear([System.Drawing.Color]::White)
        $fg.DrawImage($bmp, 0, 0, $nw, $nh)
      } finally { $fg.Dispose() }
      $enc = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() |
             Where-Object { $_.MimeType -eq 'image/jpeg' } | Select-Object -First 1
      $ep = New-Object System.Drawing.Imaging.EncoderParameters(1)
      $ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter(
                        [System.Drawing.Imaging.Encoder]::Quality, [int64]$quality)
      $flat.Save($out, $enc, $ep)
      $ep.Dispose(); $flat.Dispose()
      $newExt = '.jpeg'
    }
    $res = $out.ToArray(); $out.Dispose()
    if ($res.Length -ge $bytes.Length) { return $null }
    return @{ Bytes=$res; Ext=$newExt; W=$nw; H=$nh; OW=$ow; OH=$oh }
  } finally {
    if ($g)   { $g.Dispose() }
    if ($bmp) { $bmp.Dispose() }
    if ($img) { $img.Dispose() }
    $ms.Dispose()
  }
}

# ------------------------------------------------------------------- report ----
function Show-Report($model) {
  $items = $model.Items
  $vid  = @($items | Where-Object { $_.Kind -eq 'video'  -and $_.Referenced })
  $ras  = @($items | Where-Object { $_.Kind -eq 'raster' -and $_.Referenced })
  $orph = @($items | Where-Object { -not $_.Referenced })
  $sum = { param($a) if ($a.Count) { ($a | Measure-Object -Property Bytes -Sum).Sum } else { 0 } }
  $vB = & $sum $vid; $rB = & $sum $ras; $oB = & $sum $orph

  Write-Host ''
  Write-Host ("  {0}" -f (Split-Path $model.File -Leaf)) -ForegroundColor White
  Write-Host ("  {0}   {1} slides   {2} media parts" -f (Format-Size $model.Size), $model.Slides, $items.Count) -ForegroundColor Gray
  Write-Host ''
  $pct = { param($b) if ($model.Size) { [int](100*$b/$model.Size) } else { 0 } }
  Write-Host ("  Video     {0,10}  {1,3}%   {2} file(s)" -f (Format-Size $vB), (& $pct $vB), $vid.Count)
  Write-Host ("  Images    {0,10}  {1,3}%   {2} file(s)" -f (Format-Size $rB), (& $pct $rB), $ras.Count)
  Write-Host ("  Orphaned  {0,10}  {1,3}%   {2} file(s)" -f (Format-Size $oB), (& $pct $oB), $orph.Count)
  Write-Host ''
  $top = @($items | Where-Object { $_.Bytes -gt 300KB } | Sort-Object Bytes -Descending | Select-Object -First 15)
  if ($top.Count) {
    Write-Host '  Largest parts' -ForegroundColor White
    foreach ($t in $top) {
      $sl = @($t.Slides)
      $s = if ($sl.Count) { ($sl -join ',') } else { '-' }
      Write-Host ("   {0,10}  slide {1,-10} {2}" -f (Format-Size $t.Bytes), $s, (Split-Path $t.Part -Leaf))
    }
    Write-Host ''
  }
}

# ------------------------------------------------------------------- repair ----
function Repair-Deck($model, [string]$outFile, [hashtable]$urls, [int]$maxDim, [int]$quality,
                     [ValidateSet('hyperlink','inline')][string]$videoMode = 'hyperlink') {
  $src = [System.IO.Compression.ZipFile]::OpenRead($model.File)
  $notes = @()
  try {
    # --- plan raster re-encodes (streamed, one image resident at a time) ---
    $newBytes = @{}; $rename = @{}
    $rasters = @($model.Items | Where-Object { $_.Kind -eq 'raster' -and $_.Referenced })
    $i = 0
    foreach ($it in $rasters) {
      $i++
      Write-Progress -Activity 'Re-encoding images' -Status "$i of $($rasters.Count)" -PercentComplete (100*$i/[Math]::Max(1,$rasters.Count))
      $b = Read-EntryBytes $src $it.Part
      $r = Convert-Raster $b $it.Ext $maxDim $quality
      if ($r) {
        $newBytes[$it.Part] = $r.Bytes
        $tgt = if ($r.Ext -eq $it.Ext) { $it.Part } else { [System.IO.Path]::ChangeExtension($it.Part, $r.Ext) }
        if ($tgt -ne $it.Part) { $rename[$it.Part] = $tgt }
      }
    }
    Write-Progress -Activity 'Re-encoding images' -Completed

    # --- plan de-link ---
    $removed = New-Object System.Collections.Generic.HashSet[string]
    foreach ($it in ($model.Items | Where-Object { -not $_.Referenced })) { [void]$removed.Add($it.Part) }
    if ($model.HasThumb) { [void]$removed.Add('docProps/thumbnail.jpeg') }

    $newSlideXml = @{}; $newRelXml = @{}
    if ($urls -and $urls.Count) {
      foreach ($relPath in $model.VideoRels.Keys) {
        if ($relPath -notmatch 'ppt/slides/_rels/slide(\d+)\.xml\.rels$') { continue }
        $sn = [int]$Matches[1]
        $slidePath = $relPath -replace '_rels/','' -replace '\.rels$',''
        $relXml = Read-EntryText $src $relPath
        $sldXml = Read-EntryText $src $slidePath

        $kill = @{}; $bogus = New-Object System.Collections.Generic.HashSet[string]
        foreach ($r in $model.VideoRels[$relPath]) {
          if ($r.Bogus) {
            [void]$bogus.Add($r.Id)
            $notes += "slide$sn`: malformed video relationship $($r.Id) (Target=`"$($r.Target)`") - pre-existing defect"
            continue
          }
          if ($urls.ContainsKey($r.Abs)) { $kill[$r.Id] = $r.Abs }
        }
        if (-not $kill.Count -and -not $bogus.Count) { continue }

        $used = New-Object System.Collections.Generic.HashSet[string]
        foreach ($m in [regex]::Matches($relXml + $sldXml, 'r:(?:id|embed|link)="(rId\d+)"')) { [void]$used.Add($m.Groups[1].Value) }
        $next = 1
        foreach ($m in [regex]::Matches($relXml, 'Id="rId(\d+)"')) {
          $n = [int]$m.Groups[1].Value; if ($n -ge $next) { $next = $n + 1 }
        }
        $added = @()

        if ($videoMode -eq 'inline') {
          # KEEP the video shape and re-point its relationship at the URL rather
          # than swapping the shape for a still image. PowerPoint's linked-video
          # form is the same <p:pic> with:
          #   <a:videoFile r:link="rIdV"/>  rIdV = EXTERNAL rel of type .../video
          #   <p14:media   r:link="rIdV"/>  r:link, not r:embed (embed = a part
          #                                 inside the package)
          # The poster frame in <p:blipFill> is untouched, so the slide looks the
          # same but PowerPoint streams the file.
          #
          # Resolution is per SHAPE, not per slide. A shape's true media is named
          # by its own p14:media r:embed; the videoFile rel can be malformed
          # (Target="NULL"). Picking "the first video rel on the slide" instead
          # silently gives two shapes the same URL and loses one video.
          $relById = @{}
          foreach ($r in $model.VideoRels[$relPath]) { $relById[$r.Id] = $r }
          $ridUrl = @{}      # videoRid -> url
          $dropRid = @{}     # rels to delete once their shape no longer needs them

          foreach ($pm in [regex]::Matches($sldXml, '<p:pic>.*?</p:pic>',
                            [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
            $blk = $pm.Value
            $vm = [regex]::Match($blk, '<a:videoFile r:link="(rId\d+)"')
            $mm = [regex]::Match($blk, '<p14:media[^>]*r:embed="(rId\d+)"')
            if (-not $vm.Success -and -not $mm.Success) { continue }
            $videoRid = if ($vm.Success) { $vm.Groups[1].Value } else { $null }
            $mediaRid = if ($mm.Success) { $mm.Groups[1].Value } else { $null }

            # the media part this shape actually plays
            $abs = $null
            if ($mediaRid -and $relById.ContainsKey($mediaRid)) { $abs = $relById[$mediaRid].Abs }
            if (-not $abs -and $videoRid -and $relById.ContainsKey($videoRid)) { $abs = $relById[$videoRid].Abs }
            if (-not $abs -or -not $urls.ContainsKey($abs)) { continue }

            if (-not $videoRid) { $videoRid = $mediaRid }   # no videoFile: reuse the media rel
            $ridUrl[$videoRid] = $urls[$abs]
            [void]$removed.Add($abs)
            if ($mediaRid -and $mediaRid -ne $videoRid) { $dropRid[$mediaRid] = $true }
          }
          if (-not $ridUrl.Count) { continue }

          foreach ($rid in $ridUrl.Keys) {
            $esc = $ridUrl[$rid] -replace '&','&amp;' -replace '"','&quot;'
            $newRel = '<Relationship Id="' + $rid + '" Type="' + $VIDEO_REL_TYPE +
                      '" Target="' + $esc + '" TargetMode="External"/>'
            if ([regex]::IsMatch($relXml, '<Relationship[^>]*Id="' + $rid + '"[^>]*/>')) {
              $relXml = [regex]::Replace($relXml, '<Relationship[^>]*Id="' + $rid + '"[^>]*/>', $newRel)
            } else {
              $relXml = $relXml -replace '</Relationships>', ($newRel + '</Relationships>')
            }
          }
          foreach ($rid in $dropRid.Keys) {
            $relXml = [regex]::Replace($relXml, '<Relationship[^>]*Id="' + $rid + '"[^>]*/>', '')
          }

          # point each shape's p14:media at its own videoFile rel
          $sldXml = [regex]::Replace($sldXml, '<p:pic>.*?</p:pic>', {
            param($mo)
            $blk = $mo.Value
            $vm = [regex]::Match($blk, '<a:videoFile r:link="(rId\d+)"')
            if (-not $vm.Success) { return $blk }
            $v = $vm.Groups[1].Value
            if (-not $ridUrl.ContainsKey($v)) { return $blk }
            return [regex]::Replace($blk, '(<p14:media[^>]*)r:embed="rId\d+"', ('${1}r:link="' + $v + '"'))
          }, [System.Text.RegularExpressions.RegexOptions]::Singleline)
        }
        else {
        $sldXml = [regex]::Replace($sldXml, '<p:pic>.*?</p:pic>', {
          param($mo)
          $blk = $mo.Value
          $refs = New-Object System.Collections.Generic.HashSet[string]
          foreach ($x in [regex]::Matches($blk, '<a:videoFile r:link="(rId\d+)"')) { [void]$refs.Add($x.Groups[1].Value) }
          foreach ($x in [regex]::Matches($blk, '<p14:media[^>]*r:embed="(rId\d+)"')) { [void]$refs.Add($x.Groups[1].Value) }
          $hits = @($refs | Where-Object { $kill.ContainsKey($_) })
          $hasBogus = @($refs | Where-Object { $bogus.Contains($_) }).Count -gt 0
          if (-not $hits.Count -and -not $hasBogus) { return $blk }

          $url = if ($hits.Count) { $urls[$kill[$hits[0]]] } else { ($urls.Values | Select-Object -First 1) }
          $b = [regex]::Replace($blk, '<a:videoFile r:link="rId\d+"\s*/>', '')
          $b = [regex]::Replace($b, ('<p:extLst><p:ext uri="' + [regex]::Escape($MEDIA_EXT_URI) + '">.*?</p:ext></p:extLst>'), '')
          while ($used.Contains("rId$next")) { $next++ }
          $rid = "rId$next"; [void]$used.Add($rid); $next++
          $script:__added += ,@($rid, $url)
          $b = [regex]::Replace($b, '<a:hlinkClick r:id="" action="ppaction://media"\s*/>', "<a:hlinkClick r:id=`"$rid`"/>")
          return $b
        }, [System.Text.RegularExpressions.RegexOptions]::Singleline)

        $added = $script:__added; $script:__added = @()

        foreach ($rid in (@($kill.Keys) + @($bogus))) {
          $relXml = [regex]::Replace($relXml, ('<Relationship[^>]*Id="' + $rid + '"[^>]*/>'), '')
          if ($kill.ContainsKey($rid)) { [void]$removed.Add($kill[$rid]) }
        }
        $inject = ''
        foreach ($pair in $added) {
          $u = $pair[1] -replace '&','&amp;' -replace '"','&quot;'
          $inject += "<Relationship Id=`"$($pair[0])`" Type=`"$HLINK_TYPE`" Target=`"$u`" TargetMode=`"External`"/>"
        }
        $relXml = $relXml -replace '</Relationships>', ($inject + '</Relationships>')
        }
        $newRelXml[$relPath] = $relXml
        $newSlideXml[$slidePath] = $sldXml
      }
    }

    # --- write the new package, streaming ---
    if (Test-Path $outFile) { Remove-Item $outFile -Force }
    $dst = [System.IO.Compression.ZipFile]::Open($outFile, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
      $j = 0; $total = $src.Entries.Count
      foreach ($e in $src.Entries) {
        $j++
        if ($j % 25 -eq 0) { Write-Progress -Activity 'Writing repaired deck' -Status "$j of $total" -PercentComplete (100*$j/$total) }
        $n = $e.FullName
        if ($n.EndsWith('/')) { continue }
        if ($removed.Contains($n)) { continue }

        # de-linked slides / rels: already rewritten strings
        if ($newSlideXml.ContainsKey($n)) { Write-ZipText $dst $n $newSlideXml[$n]; continue }
        if ($newRelXml.ContainsKey($n)) {
          Write-ZipText $dst $n (Update-RelTargets $newRelXml[$n] $n $rename); continue
        }
        if ($n.EndsWith('.rels')) {
          $x = Read-EntryText $src $n
          $x = Update-RelTargets $x $n $rename
          if ($model.HasThumb) { $x = [regex]::Replace($x, '<Relationship[^>]*Target="[^"]*thumbnail\.jpeg"[^>]*/>', '') }
          Write-ZipText $dst $n $x; continue
        }
        if ($n -eq '[Content_Types].xml') {
          $x = Read-EntryText $src $n
          if ($x -notmatch 'Extension="jpeg"') {
            $x = [regex]::Replace($x, '(<Types[^>]*>)', '$1<Default Extension="jpeg" ContentType="image/jpeg"/>')
          }
          $x = [regex]::Replace($x, '<Override PartName="/docProps/thumbnail\.jpeg"[^>]*/>', '')
          Write-ZipText $dst $n $x; continue
        }
        if ($newBytes.ContainsKey($n)) {
          $tgt = if ($rename.ContainsKey($n)) { $rename[$n] } else { $n }
          $lvl = if ($tgt -match '\.(jpe?g)$') { [System.IO.Compression.CompressionLevel]::NoCompression }
                 else { [System.IO.Compression.CompressionLevel]::Optimal }
          Write-ZipBytes $dst $tgt $newBytes[$n] $lvl
          continue
        }
        # copy through, streamed
        $lvl = if ($n -like 'ppt/media/*') { [System.IO.Compression.CompressionLevel]::NoCompression }
               else { [System.IO.Compression.CompressionLevel]::Optimal }
        $ne = $dst.CreateEntry($n, $lvl)
        $ins = $e.Open(); $outs = $ne.Open()
        try { $ins.CopyTo($outs, 1MB) } finally { $outs.Dispose(); $ins.Dispose() }
      }
      Write-Progress -Activity 'Writing repaired deck' -Completed
    } finally { $dst.Dispose() }
  } finally { $src.Dispose() }

  return $notes
}

function Write-ZipText($zip, [string]$name, [string]$text) {
  $e = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
  $s = $e.Open()
  try {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
    $s.Write($bytes, 0, $bytes.Length)
  } finally { $s.Dispose() }
}

function Write-ZipBytes($zip, [string]$name, [byte[]]$bytes, $level) {
  $e = $zip.CreateEntry($name, $level)
  $s = $e.Open()
  try { $s.Write($bytes, 0, $bytes.Length) } finally { $s.Dispose() }
}

function Update-RelTargets([string]$xml, [string]$relPath, [hashtable]$rename) {
  if (-not $rename.Count) { return $xml }
  $baseDir = (Split-Path (Split-Path $relPath -Parent) -Parent) -replace '\\','/'
  return [regex]::Replace($xml, 'Target="([^"]+)"', {
    param($mo)
    $t = $mo.Groups[1].Value
    if ($t -match '^(https?:|#|mailto:)') { return $mo.Value }
    $abs = Resolve-Part $baseDir $t
    if (-not $rename.ContainsKey($abs)) { return $mo.Value }
    # relative path from the rels' base dir to the renamed part
    $to = $rename[$abs] -split '/'
    $from = if ($baseDir) { $baseDir -split '/' } else { @() }
    $k = 0
    while ($k -lt $from.Count -and $k -lt ($to.Count-1) -and $from[$k] -eq $to[$k]) { $k++ }
    $up = @(); for ($z=$k; $z -lt $from.Count; $z++) { $up += '..' }
    $rel = (($up + $to[$k..($to.Count-1)]) -join '/')
    return "Target=`"$rel`""
  })
}

# A slide referencing a relationship that is not present is exactly what makes
# PowerPoint show the repair prompt - and neither python-pptx nor LibreOffice
# will report it. So the output is checked before it is handed over.
function Test-Package([string]$file) {
  $zip = [System.IO.Compression.ZipFile]::OpenRead($file)
  try {
    $bad = @()
    foreach ($e in $zip.Entries) {
      if ($e.FullName -notmatch '^ppt/slides/slide(\d+)\.xml$') { continue }
      $rp = "ppt/slides/_rels/slide$($Matches[1]).xml.rels"
      if (-not $zip.GetEntry($rp)) { continue }
      $have = New-Object System.Collections.Generic.HashSet[string]
      foreach ($m in [regex]::Matches((Read-EntryText $zip $rp), 'Id="(rId\d+)"')) { [void]$have.Add($m.Groups[1].Value) }
      $sx = Read-EntryText $zip $e.FullName
      $miss = @()
      foreach ($m in [regex]::Matches($sx, 'r:(?:id|embed|link)="(rId\d+)"')) {
        if (-not $have.Contains($m.Groups[1].Value)) { $miss += $m.Groups[1].Value }
      }
      if ($miss.Count) { $bad += ("{0} -> {1}" -f $e.FullName, (($miss | Select-Object -Unique) -join ', ')) }
    }
    # every internal target must resolve to a real part
    $names = New-Object System.Collections.Generic.HashSet[string]
    foreach ($e in $zip.Entries) { [void]$names.Add($e.FullName) }
    foreach ($e in $zip.Entries) {
      if ($e.FullName -notlike '*.rels') { continue }
      $baseDir = (Split-Path (Split-Path $e.FullName -Parent) -Parent) -replace '\\','/'
      foreach ($m in [regex]::Matches((Read-EntryText $zip $e.FullName), '<Relationship[^>]*/>')) {
        $a = Get-RelAttrs $m.Value
        $t = if ($a.ContainsKey('Target')) { $a['Target'] } else { '' }
        if (($a.ContainsKey('TargetMode') -and $a['TargetMode'] -eq 'External') -or $t -match '^(https?:|#|mailto:)') { continue }
        $abs = Resolve-Part $baseDir $t
        if (-not $names.Contains($abs)) { $bad += ("{0} -> missing part {1}" -f $e.FullName, $abs) }
      }
    }
    return $bad
  } finally { $zip.Dispose() }
}


# ------------------------------------------------------- OneDrive discovery ---
function Get-OneDriveAccount {
  $root = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
  if (-not (Test-Path $root)) { throw 'No OneDrive accounts found in the registry. Is OneDrive signed in?' }
  foreach ($k in (Get-ChildItem $root -EA SilentlyContinue)) {
    if ($k.PSChildName -eq 'Personal') { continue }   # consumer OneDrive has no SPO host
    $p = Get-ItemProperty $k.PSPath
    if (-not $p.PSObject.Properties['UserFolder'])    { continue }
    if (-not $p.PSObject.Properties['SPOResourceId']) { continue }
    [pscustomobject]@{
      Account = $k.PSChildName
      Email   = if ($p.PSObject.Properties['UserEmail']) { $p.UserEmail } else { $null }
      Root    = $p.UserFolder
      Host    = $p.SPOResourceId.TrimEnd('/')
      Tenant  = if ($p.PSObject.Properties['ConfiguredTenantId']) { $p.ConfiguredTenantId } else { $null }
    }
    break
  }
}

# first.last@contoso.com -> first_last_contoso_com  (how SPO names personal sites)
function Get-PersonalSegment([string]$email) {
  return ($email -replace '[@.]', '_')
}

function Get-ConstructedUrl($acct, [string]$fullPath) {
  $rootTrim = $acct.Root.TrimEnd('\')
  if (-not $fullPath.StartsWith($rootTrim, [StringComparison]::OrdinalIgnoreCase)) {
    return $null   # not inside the sync root, so no web location exists
  }
  $rel = $fullPath.Substring($rootTrim.Length).TrimStart('\')
  $segs = $rel -split '\\' | ForEach-Object { [uri]::EscapeDataString($_) }
  $seg  = Get-PersonalSegment $acct.Email
  # The local sync root maps to the personal "Documents" library, so that
  # segment is always present in addition to any Documents folder on disk.
  return "$($acct.Host)/personal/$seg/Documents/" + ($segs -join '/')
}


# Slide number + shape NAME + the media file that shape plays, read straight
# from the XML. The name is the join key when driving PowerPoint through COM:
# two videos on one slide are indistinguishable by position or order, and
# getting them the wrong way round swaps the videos silently.
function Get-VideoShapeMap([string]$file) {
  $VID = @('.mp4','.mov','.m4v','.avi','.wmv','.mkv','.webm','.mpg','.mpeg')
  $zip = [System.IO.Compression.ZipFile]::OpenRead($file)
  $out = @()
  try {
    foreach ($e in $zip.Entries) {
      $mm = [regex]::Match($e.FullName, '^ppt/slides/slide(\d+)\.xml$')
      if (-not $mm.Success) { continue }
      $sn = [int]$mm.Groups[1].Value
      $rp = "ppt/slides/_rels/slide$sn.xml.rels"
      if (-not $zip.GetEntry($rp)) { continue }
      $rel = Read-EntryText $zip $rp
      $sld = Read-EntryText $zip $e.FullName

      $target = @{}
      foreach ($t in [regex]::Matches($rel, '<Relationship[^>]*/>')) {
        $a = Get-RelAttrs $t.Value
        if (-not $a.ContainsKey('Id')) { continue }
        $tgt = if ($a.ContainsKey('Target')) { $a['Target'] } else { '' }
        $ext = [System.IO.Path]::GetExtension($tgt).ToLower()
        if ($VID -contains $ext) { $target[$a['Id']] = (Split-Path $tgt -Leaf) }
      }
      if (-not $target.Count) { continue }

      foreach ($pm in [regex]::Matches($sld, '<p:pic>.*?</p:pic>',
                        [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
        $blk = $pm.Value
        $nm = [regex]::Match($blk, '<p:cNvPr[^>]*name="([^"]*)"')
        if (-not $nm.Success) { continue }
        $media = $null
        foreach ($rx in @('<p14:media[^>]*r:embed="(rId\d+)"', '<a:videoFile r:link="(rId\d+)"')) {
          foreach ($h in [regex]::Matches($blk, $rx)) {
            $rid = $h.Groups[1].Value
            if ($target.ContainsKey($rid)) { $media = $target[$rid]; break }
          }
          if ($media) { break }
        }
        if (-not $media) { continue }
        $out += [pscustomobject]@{ Slide = $sn; Name = $nm.Groups[1].Value; Media = $media }
      }
    }
  } finally { $zip.Dispose() }
  return $out
}

# Pull each video shape's existing poster frame out to disk, so it can be put
# back on the replacement shape. Without it PowerPoint shows its own first
# frame, which for a streamed file is often black until it loads.
function Export-VideoPosters([string]$file, [string]$destDir) {
  $VID = @('.mp4','.mov','.m4v','.avi','.wmv','.mkv','.webm','.mpg','.mpeg')
  $zip = [System.IO.Compression.ZipFile]::OpenRead($file)
  $map = @{}
  try {
    foreach ($e in $zip.Entries) {
      $mm = [regex]::Match($e.FullName, '^ppt/slides/slide(\d+)\.xml$')
      if (-not $mm.Success) { continue }
      $sn = [int]$mm.Groups[1].Value
      $rp = "ppt/slides/_rels/slide$sn.xml.rels"
      if (-not $zip.GetEntry($rp)) { continue }
      $rel = Read-EntryText $zip $rp
      $sld = Read-EntryText $zip $e.FullName

      $vidOf = @{}; $imgOf = @{}
      foreach ($t in [regex]::Matches($rel, '<Relationship[^>]*/>')) {
        $a = Get-RelAttrs $t.Value
        if (-not $a.ContainsKey('Id')) { continue }
        $tgt = if ($a.ContainsKey('Target')) { $a['Target'] } else { '' }
        $ext = [System.IO.Path]::GetExtension($tgt).ToLower()
        if ($VID -contains $ext) { $vidOf[$a['Id']] = (Split-Path $tgt -Leaf) }
        elseif ($ext -match '^\.(png|jpe?g|gif|bmp|tiff?)$') {
          $imgOf[$a['Id']] = (Resolve-Part 'ppt/slides' $tgt)
        }
      }
      if (-not $vidOf.Count) { continue }

      foreach ($pm in [regex]::Matches($sld, '<p:pic>.*?</p:pic>',
                        [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
        $blk = $pm.Value
        $media = $null
        foreach ($h in [regex]::Matches($blk, 'r:(?:embed|link)="(rId\d+)"')) {
          $rid = $h.Groups[1].Value
          if ($vidOf.ContainsKey($rid)) { $media = $vidOf[$rid]; break }
        }
        if (-not $media) { continue }
        $bm = [regex]::Match($blk, '<a:blip r:embed="(rId\d+)"')
        if (-not $bm.Success) { continue }
        $imgRid = $bm.Groups[1].Value
        if (-not $imgOf.ContainsKey($imgRid)) { continue }
        $part = $imgOf[$imgRid]
        $entry = $zip.GetEntry($part)
        if (-not $entry) { continue }
        $dest = Join-Path $destDir ($media + [System.IO.Path]::GetExtension($part))
        $ins = $entry.Open(); $outs = [System.IO.File]::Create($dest)
        try { $ins.CopyTo($outs) } finally { $outs.Dispose(); $ins.Dispose() }
        $map[$media] = $dest
      }
    }
  } finally { $zip.Dispose() }
  return $map
}


# Every synced location - personal OneDrive AND any SharePoint team library -
# and the SharePoint URL it maps to, read from OneDrive's own local settings.
#
# This is the whole reason no API is needed. OneDrive writes ClientPolicy*.ini
# next to its settings with a DavUrlNamespace line giving each library's web
# address; pairing that with the sync roots in the registry turns any local file
# path into its SharePoint URL offline.
#
# It also unlocks the better answer to sharing: put the videos in a team library
# the audience already has access to and there is nothing to grant at all.

# DavUrlNamespace comes back with raw spaces ("Shared Documents"), while path
# segments built here are escaped. Mixing the two yields a URL that is invalid
# and inconsistent, so the whole path is normalised the same way. Host and
# scheme are left alone.

# Turn whatever a user pastes for a SharePoint folder into a base URL that
# filenames can be appended to. This is what lets the tool work with a team
# site that is not synced locally: upload in the browser, copy the address,
# paste it in.
#
# The forms that carry a real path, and so work:
#   .../Forms/AllItems.aspx?id=%2Fteams%2FX%2FShared+Documents%2Fvideos   (address bar)
#   .../Forms/AllItems.aspx?RootFolder=%2Fteams%2F...                     (older address bar)
#   .../:f:/r/teams/X/Shared Documents/videos?csf=1&web=1                 (Copy link, path form)
#   https://host/teams/X/Shared Documents/videos                          (typed or trimmed)
#
# The form that does NOT work:
#   .../:f:/g/EiXXXXXXXXXXXXXXXXXX                                        (Copy link, token form)
# That is an opaque share token, not a path - there is no way to append a
# filename to it, and no way to expand it without an API call. Callers get
# $null and a reason so they can tell the user which link to fetch instead.
function ConvertTo-FolderBaseUrl([string]$pasted) {
  $t = ($pasted + '').Trim().Trim('"').Trim("'")
  if (-not $t) { return @{ Url = $null; Reason = 'No URL given.' } }
  if ($t -notmatch '^https?://') { return @{ Url = $null; Reason = 'That is not a URL.' } }

  $u = $null
  try { $u = [uri]$t } catch { return @{ Url = $null; Reason = 'That URL could not be parsed.' } }
  $root = "$($u.Scheme)://$($u.Authority)"

  # query-string forms carry the server-relative path
  if ($u.Query) {
    $q = [System.Web.HttpUtility]::ParseQueryString($u.Query)
    foreach ($key in @('id','RootFolder')) {
      $v = $q[$key]
      if ($v -and $v.StartsWith('/')) {
        $segs = ($v.TrimStart('/') -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
        return @{ Url = "$root/$segs"; Reason = $null }
      }
    }
  }

  $path = [uri]::UnescapeDataString($u.AbsolutePath)

  # opaque share token - cannot be extended with a filename
  if ($path -match '^/:[a-z]:/g/') {
    return @{ Url = $null; Reason = 'That is a share-token link, which cannot have filenames appended. Open the folder in the browser and copy the address bar instead, or use Copy link and pick the option that shows the folder path.' }
  }
  # path-carrying Copy link: strip the /:x:/r prefix
  if ($path -match '^/:[a-z]:/r/(.+)$') { $path = '/' + $Matches[1] }

  # drop a trailing page (AllItems.aspx etc) so only the folder remains
  $path = $path -replace '/Forms/[^/]+\.aspx$',''
  $path = $path -replace '/[^/]+\.aspx$',''
  $path = $path.TrimEnd('/')
  if (-not $path) { return @{ Url = $null; Reason = 'That URL has no folder path in it.' } }

  $segs = ($path.TrimStart('/') -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
  return @{ Url = "$root/$segs"; Reason = $null }
}

# Base folder URL + a filename -> the file's URL.
function Join-FolderUrl([string]$baseUrl, [string]$fileName) {
  if (-not $baseUrl) { return $null }
  return $baseUrl.TrimEnd('/') + '/' + [uri]::EscapeDataString($fileName)
}

function ConvertTo-EscapedUrl([string]$url) {
  $m = [regex]::Match($url, '^(https?://[^/]+)(/.*)?$')
  if (-not $m.Success) { return $url }
  $host_ = $m.Groups[1].Value
  $path  = $m.Groups[2].Value
  if (-not $path) { return $host_ }
  $segs = $path.TrimStart('/') -split '/' | ForEach-Object {
    if ($_ -eq '') { '' } else { [uri]::EscapeDataString([uri]::UnescapeDataString($_)) }
  }
  return $host_ + '/' + ($segs -join '/')
}

function Get-SyncRoots {
  $out = @()
  $settings = Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\settings'
  if (-not (Test-Path $settings)) { return $out }

  foreach ($acctDir in (Get-ChildItem $settings -Directory -EA SilentlyContinue)) {
    if ($acctDir.Name -notmatch '^Business\d+$') { continue }

    # mount points for this account, from the registry
    $mounts = @()
    $cache = "HKCU:\Software\Microsoft\OneDrive\Accounts\$($acctDir.Name)\ScopeIdToMountPointPathCache"
    if (Test-Path $cache) {
      $p = Get-ItemProperty $cache
      foreach ($n in $p.PSObject.Properties.Name) {
        if ($n -like 'PS*') { continue }
        $v = $p.$n
        if ($v -and (Test-Path -LiteralPath $v)) { $mounts += $v }
      }
    }
    if (-not $mounts.Count) { continue }

    foreach ($ini in (Get-ChildItem $acctDir.FullName -Filter 'ClientPolicy*.ini' -EA SilentlyContinue)) {
      $dav = $null; $title = $null
      foreach ($line in (Get-Content $ini.FullName -Encoding Unicode -EA SilentlyContinue)) {
        if ($line -match '^DavUrlNamespace\s*=\s*(.+?)\s*$') { $dav = $Matches[1] }
        elseif ($line -match '^SiteTitle\s*=\s*(.+?)\s*$')   { $title = $Matches[1] }
      }
      if (-not $dav) { continue }
      $dav = $dav.TrimEnd('/')

      if ($dav -match '/personal/') {
        # personal OneDrive: the sync root maps straight onto the library root
        $root = $mounts | Where-Object { (Split-Path $_ -Leaf) -like 'OneDrive*' } | Select-Object -First 1
        if ($root) { $out += [pscustomobject]@{ Root = $root; UrlBase = $dav; Title = $title; Personal = $true } }
        continue
      }

      # Team library. OneDrive names the local folder "<SiteTitle> - <folder>",
      # and DavUrlNamespace stops at the library root, so the folder segment has
      # to be added back or every URL lands one level too high.
      foreach ($mp in $mounts) {
        $leaf = Split-Path $mp -Leaf
        if ($title -and $leaf -eq $title) {
          $out += [pscustomobject]@{ Root = $mp; UrlBase = $dav; Title = $title; Personal = $false }
        } elseif ($title -and $leaf.StartsWith("$title - ")) {
          $sub = $leaf.Substring($title.Length + 3)
          $segs = ($sub -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
          $out += [pscustomobject]@{ Root = $mp; UrlBase = "$dav/$segs"; Title = $title; Personal = $false }
        }
      }
    }
  }
  return $out
}

# Local path -> SharePoint URL, using whichever sync root actually contains it.
# Longest match wins, so a team library nested under another folder still
# resolves to the right library rather than the first one that looks close.
function Get-SyncedFileUrl([string]$fullPath) {
  $best = $null
  foreach ($r in (Get-SyncRoots)) {
    $root = $r.Root.TrimEnd('\')
    if ($fullPath.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
      if (-not $best -or $root.Length -gt $best.Root.TrimEnd('\').Length) { $best = $r }
    }
  }
  if (-not $best) { return $null }
  $rel = $fullPath.Substring($best.Root.TrimEnd('\').Length).TrimStart('\')
  if (-not $rel) { return (ConvertTo-EscapedUrl $best.UrlBase) }
  $segs = ($rel -split '\\' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
  return (ConvertTo-EscapedUrl "$($best.UrlBase)/$segs")
}

Export-ModuleMember -Function ConvertTo-FolderBaseUrl, Join-FolderUrl, ConvertTo-EscapedUrl, Get-SyncRoots, Get-SyncedFileUrl, Get-VideoShapeMap, Export-VideoPosters, Get-OneDriveAccount, Get-PersonalSegment, Get-ConstructedUrl, Format-Size, Resolve-Part, Get-RelAttrs, Read-EntryText, Read-EntryBytes, Get-DeckModel, Convert-Raster, Show-Report, Repair-Deck, Write-ZipText, Write-ZipBytes, Update-RelTargets, Test-Package
