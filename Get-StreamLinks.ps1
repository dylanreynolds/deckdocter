<#
.SYNOPSIS
  Fill in the Url column of a Deck Doctor video-links CSV automatically, from
  files already sitting in a OneDrive-synced folder.

.DESCRIPTION
  No sign-in, no API, nothing to get approved. Two sources for the URLs:

  Synced locations (default)
    OneDrive writes its own settings to disk - ClientPolicy*.ini gives every
    synced library's SharePoint address - so a local path can be turned into a
    web URL entirely offline. Covers personal OneDrive and any synced team
    library. The usual trap is folder nesting: with Known Folder Move the sync
    root already contains a "Documents" folder, so a correct URL often reads
    "Documents/Documents/..." - which looks wrong and is easy to "fix" into
    something broken.

  -FolderUrl
    For a team site that is not synced here. Open the folder in the browser,
    copy the address, pass it in; filenames are appended to it.

  Either way the result is derived, not confirmed by Microsoft, so open one
  link in a browser before building a deck around them.

.PARAMETER VideoFolder
  Local folder holding the extracted videos.

.PARAMETER Csv
  The video-links CSV to fill in. Matching is by filename.

.PARAMETER FolderUrl
  SharePoint folder address for videos in a site that is not synced to this PC.
  Filenames are appended to it. Takes precedence over the synced locations.

.EXAMPLE
  .\Get-StreamLinks.ps1 -VideoFolder .\videos-for-stream -Csv .\video-links-TEMPLATE.csv

.EXAMPLE
  .\Get-StreamLinks.ps1 -VideoFolder .\videos -Csv .\deck.video-links.csv -FolderUrl 'https://contoso.sharepoint.com/teams/Team/Shared Documents/videos'

.NOTES
  Requires the videos to have finished syncing. A link to a file OneDrive has
  not uploaded yet resolves to nothing.
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory)][string]$VideoFolder,
  [Parameter(Mandatory)][string]$Csv,

  # A SharePoint folder URL to build the links from, for videos in a team site
  # that is not synced to this PC. Open the folder in the browser and copy the
  # address. Takes precedence over the synced locations.
  [string]$FolderUrl
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'DeckDoctor.psm1') -Force

# ==================================================================== main ====
$baseUrl = $null
if ($FolderUrl) {
  $parsed = ConvertTo-FolderBaseUrl $FolderUrl
  if (-not $parsed.Url) { throw $parsed.Reason }
  $baseUrl = $parsed.Url
}
$roots = @(Get-SyncRoots)
if (-not $baseUrl -and -not $roots.Count) {
  throw 'No synced OneDrive or SharePoint library found. Pass -FolderUrl with the folder''s address instead.'
}

Write-Host ''
if ($baseUrl) { Write-Host "  Building links from the folder URL you gave:" -ForegroundColor Gray
                Write-Host "    $baseUrl" -ForegroundColor DarkGray; Write-Host '' }
Write-Host '  Synced locations found (no sign-in needed):' -ForegroundColor Gray
foreach ($r in $roots) {
  $kind = if ($r.Personal) { 'personal OneDrive' } else { 'team library' }
  Write-Host ("    [{0}] {1}" -f $kind, $r.Root) -ForegroundColor Gray
  Write-Host ("        -> {0}" -f (ConvertTo-EscapedUrl $r.UrlBase)) -ForegroundColor DarkGray
}
$anyTeam = @($roots | Where-Object { -not $_.Personal }).Count
Write-Host ''

$folder = (Resolve-Path -LiteralPath $VideoFolder).Path
$csvPath = (Resolve-Path -LiteralPath $Csv).Path
$rows = @(Import-Csv -LiteralPath $csvPath)
if (-not $rows.Count) { throw "No rows in $csvPath" }
if (-not ($rows[0].PSObject.Properties['Media'] -and $rows[0].PSObject.Properties['Url'])) {
  throw "$csvPath does not look like a Deck Doctor video-links CSV (needs Media and Url columns)."
}


$files = @(Get-ChildItem -LiteralPath $folder -File |
           Where-Object { $_.Extension -match '^\.(mp4|mov|m4v|avi|wmv|mkv|webm|mpg|mpeg)$' })
Write-Host "  $($files.Count) video file(s) in $folder" -ForegroundColor Gray

# The extractor names files slideNN-<original>, so match on the original tail.
function Find-File([string]$mediaPart) {
  $leaf = Split-Path $mediaPart -Leaf
  $hit = $files | Where-Object { $_.Name -eq $leaf }
  if (-not $hit) { $hit = $files | Where-Object { $_.Name -like "*-$leaf" } }
  if (-not $hit) { $hit = $files | Where-Object { $_.Name -like "*$([System.IO.Path]::GetFileNameWithoutExtension($leaf))*" } }
  return @($hit)[0]
}

$filled = 0; $skipped = 0; $notSynced = @()
foreach ($row in $rows) {
  $f = Find-File $row.Media
  if (-not $f) {
    Write-Host ("  {0,-22} no matching file in the folder" -f (Split-Path $row.Media -Leaf)) -ForegroundColor Yellow
    $skipped++; continue
  }

  # A cloud-only placeholder means OneDrive has the file; a file still pending
  # UPLOAD is the dangerous case, and shows as present locally with no cloud
  # copy. There is no supported API for sync state, so this only warns on the
  # obvious signal: a very recent write time.
  if ((Get-Date) - $f.LastWriteTime -lt [TimeSpan]::FromMinutes(2)) { $notSynced += $f.Name }

  $constructed = if ($baseUrl) { Join-FolderUrl $baseUrl $f.Name } else { Get-SyncedFileUrl $f.FullName }
  $url = $constructed

  if (-not $url) {
    Write-Host ("  {0,-22} not inside the OneDrive sync root - no web location" -f $f.Name) -ForegroundColor Yellow
    $skipped++; continue
  }
  $row.Url = $url
  $filled++
  Write-Host ("  {0,-22} -> {1}" -f $f.Name, $url) -ForegroundColor Gray
}

$rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation
Write-Host ''
Write-Host "  $filled link(s) written to $csvPath" -ForegroundColor Green
if ($skipped) { Write-Host "  $skipped row(s) left blank" -ForegroundColor Yellow }

# Derived either way, so say so.
if ($true) {
  Write-Host ''
  $src = if ($baseUrl) { 'the folder URL you gave' } else { "OneDrive's own local settings" }
  Write-Host "  These URLs are derived from $src, not confirmed by Microsoft." -ForegroundColor Yellow
  Write-Host '  Open one in a browser before building the deck.' -ForegroundColor Yellow
  if ($anyTeam) {
    Write-Host ''
    Write-Host '  Tip: videos kept in a TEAM library inherit that site''s permissions, so' -ForegroundColor Cyan
    Write-Host '  colleagues can already play them and no sharing step is needed at all.' -ForegroundColor Cyan
    Write-Host '  Files in your personal OneDrive start private and must be shared first.' -ForegroundColor Cyan
  }
}

if ($notSynced.Count) {
  Write-Host ''
  Write-Host "  Written in the last two minutes, so possibly not uploaded yet: $($notSynced -join ', ')" -ForegroundColor Yellow
  Write-Host '  Wait for the OneDrive sync tick before trusting those links.' -ForegroundColor Yellow
}
Write-Host ''
Write-Host ("  Next: .\Optimize-Deck.ps1 -Path <deck.pptx> -VideoLinks `"$csvPath`"") -ForegroundColor Gray
Write-Host ''
