<#
.SYNOPSIS
  Window for Deck Doctor. Same engine as Optimize-Deck.ps1 (DeckDoctor.psm1),
  wrapped in WPF so it can be handed to someone who will never open a terminal.

.DESCRIPTION
  Drop a .pptx in, read the diagnosis, optionally save the videos out to a
  OneDrive folder and paste their Stream links back, then build. The original
  file is never modified.

  Launch:  powershell -ExecutionPolicy Bypass -STA -File .\Deck-Doctor-GUI.ps1

  -STA matters: WPF's PNG encoder requires a single-threaded apartment.
#>
[CmdletBinding()]
param(
  # Optional deck to preload, so the script can be set as the "Open with"
  # handler for .pptx. $args does not exist in an advanced function, and
  # touching it under StrictMode is fatal - hence a real parameter.
  [Parameter(Position=0)][string]$Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# A WPF event handler that throws takes the whole window down with it, which
# looks exactly like a crash and leaves nothing behind to diagnose. Everything
# wired to a control goes through Invoke-Safe, and anything that still escapes
# is caught at the dispatcher and written to this log.
$script:LogPath = Join-Path $PSScriptRoot 'deck-doctor.log'

function Write-Log([string]$msg) {
  try {
    Add-Content -LiteralPath $script:LogPath -Encoding utf8 -Value (
      '{0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date), $msg)
  } catch { }
}

function Invoke-Safe([string]$what, [scriptblock]$body) {
  try { & $body }
  catch {
    $e = $_
    Write-Log "ERROR in $what : $($e.Exception.GetType().Name): $($e.Exception.Message)"
    Write-Log "  at $($e.InvocationInfo.ScriptLineNumber): $($e.InvocationInfo.Line.Trim())"
    Write-Log "  $($e.ScriptStackTrace -replace "`r?`n", ' | ')"
    try {
      if ($script:ui) {
        foreach ($b in @('BtnBrowse','BtnAnalyse','BtnBuild','BtnSaveVideos','BtnFillLinks')) {
          if ($script:ui[$b]) { $script:ui[$b].IsEnabled = $true }
        }
        if ($script:ui.Bar) { $script:ui.Bar.Visibility = 'Collapsed' }
        if ($script:ui.TxtStatus) { $script:ui.TxtStatus.Text = '' }
      }
    } catch { }
    [System.Windows.MessageBox]::Show(
      "$what failed. The window is still usable and your original file is untouched.`n`n" +
      "$($e.Exception.Message)`n`nFull detail written to:`n$($script:LogPath)",
      'Deck Doctor', 'OK', 'Error') | Out-Null
  }
}

Write-Log "--- session start (PID $PID, PowerShell $($PSVersionTable.PSVersion)) ---"

$ModulePath = Join-Path $PSScriptRoot 'DeckDoctor.psm1'
if (-not (Test-Path $ModulePath)) {
  [System.Windows.MessageBox]::Show("DeckDoctor.psm1 not found beside this script.`n`nExpected: $ModulePath",
    'Deck Doctor', 'OK', 'Error') | Out-Null
  return
}
Import-Module $ModulePath -Force

# --------------------------------------------------------------------- XAML ---
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Deck Doctor - Subway" Height="820" Width="1000"
        WindowStartupLocation="CenterScreen" Background="#F9FAFB"
        FontFamily="Segoe UI">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="14,8"/>
      <Setter Property="Margin" Value="0,0,10,0"/>
      <Setter Property="Background" Value="White"/>
      <Setter Property="BorderBrush" Value="#E5E7EB"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Cursor" Value="Hand"/>
    </Style>
    <Style x:Key="Primary" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="#008C15"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderBrush" Value="#008C15"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="White"/>
      <Setter Property="BorderBrush" Value="#E5E7EB"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="6"/>
      <Setter Property="Padding" Value="16"/>
      <Setter Property="Margin" Value="0,0,0,14"/>
    </Style>
    <Style x:Key="H" TargetType="TextBlock">
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#111827"/>
      <Setter Property="Margin" Value="0,0,0,2"/>
    </Style>
    <Style x:Key="Sub" TargetType="TextBlock">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Foreground" Value="#6B7280"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Margin" Value="0,0,0,10"/>
    </Style>
    <Style x:Key="KpiLabel" TargetType="TextBlock">
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Foreground" Value="#6B7280"/>
    </Style>
    <Style x:Key="KpiValue" TargetType="TextBlock">
      <Setter Property="FontSize" Value="21"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#111827"/>
    </Style>
  </Window.Resources>

  <DockPanel>
    <Border DockPanel.Dock="Top" Background="White" BorderBrush="#008C15" BorderThickness="0,0,0,3">
      <StackPanel Orientation="Horizontal" Margin="24,14">
        <TextBlock Text="Subway" FontSize="17" FontWeight="Bold" Foreground="#111827"/>
        <TextBlock Text="  |  " FontSize="13" Foreground="#9CA3AF" VerticalAlignment="Center"/>
        <TextBlock Text="Deck Doctor - PowerPoint size diagnosis and repair"
                   FontSize="13" Foreground="#374151" VerticalAlignment="Center"/>
      </StackPanel>
    </Border>

    <Border DockPanel.Dock="Bottom" Background="White" BorderBrush="#E5E7EB" BorderThickness="0,1,0,0">
      <TextBlock Margin="24,10" FontSize="11" Foreground="#6B7280"
                 Text="Runs entirely on this machine - nothing is uploaded. Your original file is never modified."/>
    </Border>

    <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="24,18">
      <StackPanel>

        <Border Style="{StaticResource Card}" AllowDrop="True" Name="DropZone">
          <StackPanel>
            <TextBlock Style="{StaticResource H}" Text="1. Choose a deck"/>
            <TextBlock Style="{StaticResource Sub}"
                       Text="Drop a .pptx anywhere on this panel, or browse. Large decks are fine - parts are streamed one at a time."/>
            <StackPanel Orientation="Horizontal">
              <TextBox Name="TxtPath" Width="700" Height="30" VerticalContentAlignment="Center"
                       IsReadOnly="True" Background="#F9FAFB" BorderBrush="#E5E7EB" Padding="8,0"
                       FontSize="12" Margin="0,0,10,0"/>
              <Button Name="BtnBrowse" Content="Browse..."/>
              <Button Name="BtnAnalyse" Content="Analyse" Style="{StaticResource Primary}" IsEnabled="False"/>
            </StackPanel>
          </StackPanel>
        </Border>

        <Border Style="{StaticResource Card}" Name="CardDiag" Visibility="Collapsed">
          <StackPanel>
            <TextBlock Style="{StaticResource H}" Text="2. Diagnosis"/>
            <TextBlock Style="{StaticResource Sub}" Name="TxtDeckLine" Text=""/>
            <UniformGrid Rows="1" Columns="4" Margin="0,4,0,12">
              <StackPanel><TextBlock Style="{StaticResource KpiLabel}" Text="Current size"/><TextBlock Style="{StaticResource KpiValue}" Name="KpiSize"/></StackPanel>
              <StackPanel><TextBlock Style="{StaticResource KpiLabel}" Text="Video"/><TextBlock Style="{StaticResource KpiValue}" Name="KpiVideo"/></StackPanel>
              <StackPanel><TextBlock Style="{StaticResource KpiLabel}" Text="Images"/><TextBlock Style="{StaticResource KpiValue}" Name="KpiImages"/></StackPanel>
              <StackPanel><TextBlock Style="{StaticResource KpiLabel}" Text="Orphaned"/><TextBlock Style="{StaticResource KpiValue}" Name="KpiOrphan"/></StackPanel>
            </UniformGrid>
            <TextBlock Name="TxtFindings" FontSize="12" Foreground="#374151" TextWrapping="Wrap"/>
          </StackPanel>
        </Border>

        <Border Style="{StaticResource Card}" Name="CardVideo" Visibility="Collapsed">
          <StackPanel>
            <TextBlock Style="{StaticResource H}" Text="3. Video - the largest single win"/>
            <TextBlock Style="{StaticResource Sub}"
                       Text="Save the videos into a OneDrive folder; sync uploads them, which is what puts them in Stream. Then paste each share link below. Fill in every row and the build removes the video outright and repoints the existing thumbnail at Stream. Leave them blank and video stays embedded."/>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
              <Button Name="BtnSaveVideos" Content="1. Save videos to a folder..."/>
              <Button Name="BtnFillLinks" Content="2. Fill links from OneDrive" IsEnabled="False"/>
            </StackPanel>
            <TextBlock Name="TxtVideoStatus" FontSize="12" Foreground="#6B7280" TextWrapping="Wrap" Margin="0,0,0,12"/>

            <TextBlock Style="{StaticResource H}" Text="Sharing" Margin="0,4,0,2"/>
            <TextBlock Style="{StaticResource Sub}" Margin="0,0,0,6"
                       Text="The deck streams these files from wherever they live; it does not carry them. You will always see them play because you own them, so test with a colleague rather than yourself."/>
            <TextBlock Style="{StaticResource Sub}" Margin="0,0,0,6"
                       Text="Simplest approach: save the videos into a synced SharePoint team library rather than your personal OneDrive. Files there inherit the site's permissions, so anyone with access to the site can already play them and there is nothing to share. A personal OneDrive folder starts private - if you use one, share that folder yourself in OneDrive (right-click the folder, Share, then choose everyone in your organisation)."/>
            <TextBlock Name="TxtSyncRoots" FontSize="10" Foreground="#6B7280" TextWrapping="Wrap"
                       FontFamily="Consolas" Margin="0,0,0,12"/>
            <TextBlock Style="{StaticResource H}" Text="Team site folder (optional)" Margin="0,4,0,2"/>
            <TextBlock Style="{StaticResource Sub}" Margin="0,0,0,6"
                       Text="If the videos live in a SharePoint folder that is not synced to this PC, open that folder in the browser and paste its address here - the links below are then built from it. Leave blank to use the synced locations above. A Copy link that looks like /:f:/g/... will not work; it is a share token with no folder path in it, so use the browser address bar."/>
            <TextBox Name="TxtFolderUrl" Height="28" VerticalContentAlignment="Center"
                     BorderBrush="#E5E7EB" Padding="8,0" FontSize="11" Margin="0,0,0,12"
                     ToolTip="e.g. https://contoso.sharepoint.com/teams/YourTeam/Shared Documents/videos"/>
            <DataGrid Name="GridVideos" AutoGenerateColumns="False" HeadersVisibility="Column"
                      CanUserAddRows="False" CanUserDeleteRows="False" GridLinesVisibility="Horizontal"
                      HorizontalGridLinesBrush="#E5E7EB" BorderBrush="#E5E7EB" BorderThickness="1"
                      RowBackground="White" AlternatingRowBackground="#F9FAFB" FontSize="12"
                      MaxHeight="260" Background="White">
              <DataGrid.Columns>
                <DataGridTextColumn Header="File"  Binding="{Binding Name}"   Width="150" IsReadOnly="True"/>
                <DataGridTextColumn Header="Size"  Binding="{Binding SizeText}" Width="80" IsReadOnly="True"/>
                <DataGridTextColumn Header="Slide" Binding="{Binding SlideText}" Width="60" IsReadOnly="True"/>
                <DataGridTextColumn Header="Stream share link (paste here)" Binding="{Binding Url, UpdateSourceTrigger=PropertyChanged}" Width="*"/>
              </DataGrid.Columns>
            </DataGrid>
          </StackPanel>
        </Border>

        <Border Style="{StaticResource Card}" Name="CardBuild" Visibility="Collapsed">
          <StackPanel>
            <TextBlock Style="{StaticResource H}" Text="4. Build"/>
            <TextBlock Style="{StaticResource Sub}" Text="Produces a new copy. Verified before it is handed over: if any slide would reference a missing relationship, nothing is written."/>
            <TextBlock Name="TxtLevel" FontSize="12" Foreground="#111827" FontWeight="SemiBold" Margin="0,0,0,4"/>
            <Slider Name="SldLevel" Minimum="1" Maximum="3" Value="2" TickFrequency="1"
                    IsSnapToTickEnabled="True" TickPlacement="BottomRight" Width="420"
                    HorizontalAlignment="Left" Margin="0,0,0,2"/>
            <Grid Width="420" HorizontalAlignment="Left" Margin="0,0,0,14">
              <TextBlock Text="Conservative" FontSize="10" Foreground="#6B7280" HorizontalAlignment="Left"/>
              <TextBlock Text="Balanced" FontSize="10" Foreground="#6B7280" HorizontalAlignment="Center"/>
              <TextBlock Text="Aggressive" FontSize="10" Foreground="#6B7280" HorizontalAlignment="Right"/>
            </Grid>
            <StackPanel Orientation="Horizontal">
              <Button Name="BtnBuild" Content="Build repaired copy" Style="{StaticResource Primary}"/>
              <Button Name="BtnOpenFolder" Content="Open output folder" IsEnabled="False"/>
            </StackPanel>
            <ProgressBar Name="Bar" Height="6" Margin="0,14,0,6" Foreground="#008C15"
                         Background="#F3F4F6" BorderThickness="0" Visibility="Collapsed"/>
            <TextBlock Name="TxtStatus" FontSize="12" Foreground="#374151" TextWrapping="Wrap"/>
            <Border Name="ResultBox" Background="#F3FBF5" BorderBrush="#A7E3B4" BorderThickness="1"
                    CornerRadius="4" Padding="12" Margin="0,12,0,0" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Name="TxtResult" FontSize="13" FontWeight="SemiBold" Foreground="#006B10" TextWrapping="Wrap"/>
                <TextBlock Name="TxtResultPath" FontSize="11" Foreground="#374151" TextWrapping="Wrap" Margin="0,4,0,0"/>
              </StackPanel>
            </Border>
            <Border Name="NoteBox" Background="#FEF3C7" BorderBrush="#FCD34D" BorderThickness="1"
                    CornerRadius="4" Padding="12" Margin="0,10,0,0" Visibility="Collapsed">
              <StackPanel>
                <TextBlock Text="Pre-existing problems found in the source deck" FontSize="12" FontWeight="SemiBold" Foreground="#92400E"/>
                <TextBlock Name="TxtNotes" FontSize="11" Foreground="#92400E" TextWrapping="Wrap" Margin="0,3,0,0"/>
              </StackPanel>
            </Border>
          </StackPanel>
        </Border>

        <Border Style="{StaticResource Card}">
          <StackPanel>
            <TextBlock Style="{StaticResource H}" Text="What only PowerPoint can fix"/>
            <TextBlock Style="{StaticResource Sub}" Margin="0"
                       Text="Cropped-away image data: a cropped photo still stores every original pixel - select all images, then Picture Format &gt; Compress Pictures &gt; Delete cropped areas of pictures. Editing history: File &gt; Options &gt; Advanced &gt; Discard editing data. Embedded fonts: File &gt; Options &gt; Save. Embedded workbooks: paste charts as pictures instead. Presenting tip: a heavy deck presents far better through PowerPoint Live in Teams than local Slide Show, and always carry a PDF export."/>
          </StackPanel>
        </Border>

      </StackPanel>
    </ScrollViewer>
  </DockPanel>
</Window>
'@

$win = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))

$script:ui = @{}
$ui = $script:ui
foreach ($n in @('DropZone','TxtPath','BtnBrowse','BtnAnalyse','CardDiag','TxtDeckLine','KpiSize','KpiVideo',
                 'KpiImages','KpiOrphan','TxtFindings','CardVideo','BtnSaveVideos','BtnFillLinks','TxtFolderUrl','TxtVideoStatus','GridVideos',
                 'TxtSyncRoots',
                 'CardBuild','TxtLevel','SldLevel','BtnBuild','BtnOpenFolder','Bar','TxtStatus','ResultBox',
                 'TxtResult','TxtResultPath','NoteBox','TxtNotes')) {
  $ui[$n] = $win.FindName($n)
}

$PROFILES = @{
  1 = @{ Name='Conservative'; Max=1920; Q=85 }
  2 = @{ Name='Balanced';     Max=1600; Q=80 }
  3 = @{ Name='Aggressive';   Max=1280; Q=75 }
}

$state = [hashtable]::Synchronized(@{
  File = $null; Model = $null; OutFile = $null; OutDir = $null
  Rows = New-Object System.Collections.ObjectModel.ObservableCollection[object]
  Worker = $null; Handle = $null; Timer = $null; Phase = ''; VideoFolder = $null
})

# Folder picking differs by runtime, so it is probed rather than assumed.
# .NET 10 (PowerShell 7.6) has Microsoft.Win32.OpenFolderDialog - the modern
# native picker. Older hosts fall back to WinForms FolderBrowserDialog, whose
# writable properties also differ between .NET Framework and .NET Core: the
# title override is UseDescriptionForTitle, NOT UseDescriptionForSelection.
# Setting a property that does not exist is a terminating error under
# StrictMode, which is what took the window down.
function Select-Folder([string]$prompt) {
  $t = 'Microsoft.Win32.OpenFolderDialog' -as [type]
  if ($t) {
    $d = New-Object Microsoft.Win32.OpenFolderDialog
    $d.Title = $prompt
    $d.Multiselect = $false
    if ($d.ShowDialog()) { return $d.FolderName }
    return $null
  }
  $d = New-Object System.Windows.Forms.FolderBrowserDialog
  $d.Description = $prompt
  if ($d.PSObject.Properties['UseDescriptionForTitle']) { $d.UseDescriptionForTitle = $true }
  if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $d.SelectedPath }
  return $null
}

function Set-Level {
  $p = $PROFILES[[int]$ui.SldLevel.Value]
  $ui.TxtLevel.Text = "Compression level - $($p.Name.ToLower()) ($($p.Max)px, quality $($p.Q))"
}

function Set-Busy([bool]$busy) {
  foreach ($b in @('BtnBrowse','BtnAnalyse','BtnBuild','BtnSaveVideos','BtnFillLinks')) { $ui[$b].IsEnabled = -not $busy }
  $ui.Bar.Visibility = if ($busy) { 'Visible' } else { 'Collapsed' }
  $ui.Bar.IsIndeterminate = $busy
  if (-not $busy -and $state.File) { $ui.BtnAnalyse.IsEnabled = $true }
}

function Select-Deck([string]$path) {
  if (-not $path -or -not (Test-Path -LiteralPath $path)) { return }
  if ($path -notmatch '\.pptm?x$') {
    [System.Windows.MessageBox]::Show('That is not a .pptx file. Deck Doctor works on the Open XML format only.',
      'Deck Doctor','OK','Warning') | Out-Null
    return
  }
  $state.File = (Resolve-Path -LiteralPath $path).Path
  $ui.TxtPath.Text = $state.File
  $ui.BtnAnalyse.IsEnabled = $true
  foreach ($c in @('CardDiag','CardVideo','CardBuild')) { $ui[$c].Visibility = 'Collapsed' }
  $ui.ResultBox.Visibility = 'Collapsed'; $ui.NoteBox.Visibility = 'Collapsed'
}

# ------------------------------------------------------------------ analysis --
# Analysis is quick enough to run inline; only the repair needs a worker.
function Invoke-Analyse {
  if (-not $state.File) { return }
  Set-Busy $true
  $ui.TxtStatus.Text = 'Reading the package and resolving every relationship...'
  $win.Dispatcher.Invoke([action]{}, 'Background')
  try {
    $model = Get-DeckModel $state.File
  } catch {
    Set-Busy $false
    $ui.TxtStatus.Text = ''
    [System.Windows.MessageBox]::Show("Could not read that deck.`n`n$($_.Exception.Message)",
      'Deck Doctor','OK','Error') | Out-Null
    return
  }
  $state.Model = $model

  $items = $model.Items
  $vid  = @($items | Where-Object { $_.Kind -eq 'video'  -and $_.Referenced })
  $ras  = @($items | Where-Object { $_.Kind -eq 'raster' -and $_.Referenced })
  $orph = @($items | Where-Object { -not $_.Referenced })
  $sum = { param($a) if ($a.Count) { ($a | Measure-Object Bytes -Sum).Sum } else { 0 } }
  $vB = & $sum $vid; $rB = & $sum $ras; $oB = & $sum $orph

  $ui.TxtDeckLine.Text = "$(Split-Path $model.File -Leaf)   -   $($model.Slides) slides, $($items.Count) media parts"
  $ui.KpiSize.Text   = Format-Size $model.Size
  $ui.KpiVideo.Text  = Format-Size $vB
  $ui.KpiImages.Text = Format-Size $rB
  $ui.KpiOrphan.Text = Format-Size $oB

  $f = @()
  if ($model.Size -gt 0) {
    if ($vB -gt 0.15 * $model.Size) {
      $f += "Embedded video is the dominant cost: $($vid.Count) file(s), $([int](100*$vB/$model.Size))% of the deck. Compression cannot fix this - the video has to leave the file."
    }
    if ($ras.Count) {
      $f += "$($ras.Count) images totalling $(Format-Size $rB). Most decks store images far larger than they display; re-encoding usually recovers the majority of this with no visible change."
    }
    if ($orph.Count) {
      $f += "$($orph.Count) orphaned media part(s), $(Format-Size $oB), left behind by deleted slides and referenced by nothing. PowerPoint never removes these."
    }
  }
  if (-not $f.Count) { $f += 'No major bloat found - this deck is already reasonably sized.' }
  $ui.TxtFindings.Text = ($f -join "`n`n")

  $state.Rows.Clear()
  foreach ($v in ($vid | Sort-Object Bytes -Descending)) {
    $sl = @($v.Slides)
    $state.Rows.Add([pscustomobject]@{
      Name      = Split-Path $v.Part -Leaf
      Part      = $v.Part
      SizeText  = Format-Size $v.Bytes
      SlideText = if ($sl.Count) { $sl -join ',' } else { '-' }
      FirstSlide= if ($sl.Count) { $sl[0] } else { 0 }
      Url       = ''
    })
  }
  $ui.GridVideos.ItemsSource = $state.Rows

  # Show the synced locations, so the team-library option is a visible fact
  # rather than advice. This is read from OneDrive's own settings - no API.
  try {
    $lines = @()
    foreach ($r in (Get-SyncRoots)) {
      $kind = if ($r.Personal) { 'personal - private by default' } else { 'team library - already shared with the site' }
      $lines += ("{0}  -  {1}" -f (Split-Path $r.Root -Leaf), $kind)
    }
    $ui.TxtSyncRoots.Text = if ($lines.Count) { "Synced here:`n  " + ($lines -join "`n  ") } else { '' }
  } catch { $ui.TxtSyncRoots.Text = '' }
  $ui.CardVideo.Visibility = if ($vid.Count) { 'Visible' } else { 'Collapsed' }
  if ($vid.Count) { $ui.BtnFillLinks.IsEnabled = $true }
  $ui.CardDiag.Visibility = 'Visible'
  $ui.CardBuild.Visibility = 'Visible'
  $ui.TxtStatus.Text = ''
  Set-Busy $false
}

# ------------------------------------------------------------ save videos out --
function Save-Videos {
  if (-not $state.Model) { return }
  $dest = Select-Folder 'Choose a folder inside OneDrive - sync will upload the videos to Stream'
  if (-not $dest) { return }
  Set-Busy $true
  $n = 0
  try {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($state.Model.File)
    try {
      foreach ($r in $state.Rows) {
        $name = ('slide{0:d2}-{1}' -f $r.FirstSlide, $r.Name)
        $ui.TxtVideoStatus.Text = "Writing $name ..."
        $win.Dispatcher.Invoke([action]{}, 'Background')
        $e = $zip.GetEntry($r.Part)
        $ins = $e.Open(); $outs = [System.IO.File]::Create((Join-Path $dest $name))
        try { $ins.CopyTo($outs, 1MB) } finally { $outs.Dispose(); $ins.Dispose() }
        $n++
      }
    } finally { $zip.Dispose() }
    $ui.TxtVideoStatus.Text = "$n file(s) written to $dest - wait for the OneDrive sync tick, then use 'Fill links from OneDrive'."
    $state.VideoFolder = $dest
    $ui.BtnFillLinks.IsEnabled = $true
  } catch {
    $ui.TxtVideoStatus.Text = "Stopped after $n file(s): $($_.Exception.Message)"
  }
  Set-Busy $false
}

# ------------------------------------------------------------- fill links ----
# The links are derivable: OneDrive records the tenant host and sync root, so a
# local path maps to a web URL without any sign-in. That mapping is a
# PREDICTION though, so the checkbox offers to have Microsoft confirm it.
function Fill-Links {
  $dest = $state.VideoFolder
  if (-not $dest -or -not (Test-Path -LiteralPath $dest)) {
    $dest = Select-Folder 'Where did you save the videos?'
    if (-not $dest) { return }
    $state.VideoFolder = $dest
  }

  # A pasted folder URL wins: it also covers team sites that are not synced.
  $baseUrl = $null
  $pasted = ($ui.TxtFolderUrl.Text + '').Trim()
  if ($pasted) {
    $parsed = ConvertTo-FolderBaseUrl $pasted
    if (-not $parsed.Url) {
      [System.Windows.MessageBox]::Show(
        "That folder URL cannot be used.`n`n$($parsed.Reason)",
        'Deck Doctor','OK','Warning') | Out-Null
      return
    }
    $baseUrl = $parsed.Url
  } elseif (-not @(Get-SyncRoots).Count) {
    [System.Windows.MessageBox]::Show(
      'No synced OneDrive or SharePoint library was found, so links cannot be derived from the file paths. Paste the folder URL above instead.',
      'Deck Doctor','OK','Warning') | Out-Null
    return
  }

  Set-Busy $true
  $files = @(Get-ChildItem -LiteralPath $dest -File |
             Where-Object { $_.Extension -match '^\.(mp4|mov|m4v|avi|wmv|mkv|webm|mpg|mpeg)$' })
  $filled = 0; $missing = 0
  foreach ($r in $state.Rows) {
    $leaf = $r.Name
    $f = $files | Where-Object { $_.Name -eq $leaf -or $_.Name -like "*-$leaf" } | Select-Object -First 1
    if (-not $f) { $missing++; continue }
    $url = if ($baseUrl) { Join-FolderUrl $baseUrl $f.Name } else { Get-SyncedFileUrl $f.FullName }
    if ($url) { $r.Url = $url; $filled++ }
  }
  $ui.GridVideos.Items.Refresh()

  $msg = "$filled of $($state.Rows.Count) link(s) filled in."
  if ($missing) { $msg += " $missing had no matching file in that folder." }
  $msg += if ($baseUrl) { ' Built from the folder URL you pasted' } else { ' Derived from the local sync settings' }
  $msg += ' - open one in a browser before you build, to be sure.'
  $ui.TxtVideoStatus.Text = $msg
  Write-Log "fill-links: filled=$filled missing=$missing pastedUrl=$([bool]$baseUrl)"
  Set-Busy $false
}

# ------------------------------------------------------------------- build ----
function Invoke-Build {
  if (-not $state.Model) { return }

  # All-or-nothing: a partial set of links would strip some video while leaving
  # other slides pointing at media that is no longer in the package.
  $urls = @{}
  $filled = 0
  foreach ($r in $state.Rows) {
    if ($r.Url -and $r.Url.Trim() -match '^https?://') { $urls[$r.Part] = $r.Url.Trim(); $filled++ }
  }
  if ($state.Rows.Count -and $filled -ne $state.Rows.Count) {
    if ($filled -gt 0) {
      $ans = [System.Windows.MessageBox]::Show(
        "$filled of $($state.Rows.Count) links are filled in.`n`nVideo can only be removed when every row has a link. Build anyway, fixing images only and leaving video embedded?",
        'Deck Doctor','YesNo','Question')
      if ($ans -ne [System.Windows.MessageBoxResult]::Yes) { return }
    }
    $urls = @{}
  }

  $p = $PROFILES[[int]$ui.SldLevel.Value]
  $state.OutDir = Join-Path (Split-Path $state.Model.File -Parent) 'repaired'
  $state.OutDir = (New-Item -ItemType Directory -Force -Path $state.OutDir).FullName
  $state.OutFile = Join-Path $state.OutDir ((Split-Path $state.Model.File -LeafBase) + ' (repaired).pptx')

  $ui.ResultBox.Visibility = 'Collapsed'; $ui.NoteBox.Visibility = 'Collapsed'
  Set-Busy $true
  $ui.Bar.IsIndeterminate = $false
  $ui.Bar.Value = 0
  $ui.TxtStatus.Text = 'Starting...'

  # STA runspace: the WPF PNG encoder the engine uses requires one.
  $rs = [runspacefactory]::CreateRunspace()
  $rs.ApartmentState = 'STA'
  $rs.ThreadOptions  = 'ReuseThread'
  $rs.Open()
  $rs.SessionStateProxy.SetVariable('modulePath', $ModulePath)
  $rs.SessionStateProxy.SetVariable('deckFile',   $state.Model.File)
  $rs.SessionStateProxy.SetVariable('outFile',    $state.OutFile)
  $rs.SessionStateProxy.SetVariable('urls',       $urls)
  $rs.SessionStateProxy.SetVariable('maxDim',     $p.Max)
  $rs.SessionStateProxy.SetVariable('quality',    $p.Q)

  $rs.SessionStateProxy.SetVariable('inlineScript', (Join-Path $PSScriptRoot 'Set-InlineVideo.ps1'))

  $worker = [powershell]::Create()
  $worker.Runspace = $rs
  [void]$worker.AddScript({
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'
    Import-Module $modulePath -Force

    $origSize = (Get-Item -LiteralPath $deckFile).Length
    $source = $deckFile
    $temp = $null
    $notes = @()

    # Video first, images second. Hand-written video XML is rejected by
    # PowerPoint no matter how well-formed it is, so the video step is done BY
    # PowerPoint through COM. And it has to run first: PowerPoint's save
    # re-inflates already-compressed images (measured 15.7 MB of PNG back up to
    # 39.5 MB), so doing images first throws that work away.
    if ($urls.Count) {
      Write-Progress -Activity 'Converting video' -Status 'PowerPoint is rewriting the video shapes' -PercentComplete 5
      $csvTmp = Join-Path $env:TEMP ('deckdoctor-links-' + [guid]::NewGuid().ToString('N').Substring(0,8) + '.csv')
      $urls.GetEnumerator() | ForEach-Object {
        [pscustomobject]@{ Media = $_.Key; Url = $_.Value }
      } | Export-Csv -LiteralPath $csvTmp -NoTypeInformation

      $temp = Join-Path $env:TEMP ('deckdoctor-inline-' + [guid]::NewGuid().ToString('N').Substring(0,8) + '.pptx')
      & $inlineScript -Path $deckFile -VideoLinks $csvTmp -OutFile $temp | Out-Null
      Remove-Item -LiteralPath $csvTmp -Force -ErrorAction SilentlyContinue
      if (-not (Test-Path -LiteralPath $temp)) {
        throw 'The video conversion produced nothing. Is PowerPoint installed and closed?'
      }
      $source = $temp
      $notes += 'Video converted to linked online media by PowerPoint.'
    }

    $model = Get-DeckModel $source
    # no urls passed on: the XML video path is never used
    $notes += @(Repair-Deck $model $outFile @{} $maxDim $quality)
    $bad = @(Test-Package $outFile)
    if ($temp) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
    [pscustomobject]@{ Notes = $notes; Bad = $bad; Size = $origSize }
  })

  $state.Worker = $worker
  $state.Handle = $worker.BeginInvoke()

  # Poll the worker's progress stream - Write-Progress inside the engine lands
  # there, so the bar reflects real work without the engine knowing about a UI.
  $timer = New-Object System.Windows.Threading.DispatcherTimer
  $timer.Interval = [TimeSpan]::FromMilliseconds(250)
  $state.Timer = $timer
  $timer.Add_Tick({ Invoke-Safe 'Repair' {
    # A tick can still fire after Stop() is queued, and the handler nulls
    # $state.Worker when it finishes - so re-entry must be survivable.
    $w = $state.Worker
    if (-not $w) { $state.Timer.Stop(); return }

    if ($w.Streams.Progress.Count) {
      $rec = $w.Streams.Progress[$w.Streams.Progress.Count-1]
      if ($rec.PercentComplete -ge 0) { $ui.Bar.Value = $rec.PercentComplete }
      $ui.TxtStatus.Text = "$($rec.Activity) - $($rec.StatusDescription)"
    }
    if (-not $state.Handle.IsCompleted) { return }

    $state.Timer.Stop()
    $res = $null; $err = $null
    try { $res = $w.EndInvoke($state.Handle) } catch { $err = $_ }
    $streamErr = if ($w.Streams.Error.Count) { $w.Streams.Error[0].ToString() } else { $null }
    try { $w.Runspace.Close(); $w.Dispose() } catch { }
    $state.Worker = $null

    Set-Busy $false
    $ui.Bar.Visibility = 'Collapsed'

    if ($err -or $streamErr) {
      $ui.TxtStatus.Text = ''
      $msg = if ($err) { $err.Exception.Message } else { $streamErr }
      Write-Log "worker failed: $msg"
      [System.Windows.MessageBox]::Show("The repair failed and nothing was written.`n`n$msg",
        'Deck Doctor','OK','Error') | Out-Null
      return
    }

    # An empty result would make @($res)[-1] throw under StrictMode, which
    # inside an event handler would take the window down with no message.
    $all = @($res)
    if (-not $all.Count) {
      $ui.TxtStatus.Text = ''
      Write-Log 'worker returned no result object'
      [System.Windows.MessageBox]::Show(
        "The repair returned nothing. Detail in:`n$($script:LogPath)",
        'Deck Doctor','OK','Error') | Out-Null
      return
    }
    $out = $all[-1]
    if ($out.Bad.Count) {
      if (Test-Path -LiteralPath $state.OutFile) { Remove-Item -LiteralPath $state.OutFile -Force }
      $ui.TxtStatus.Text = ''
      [System.Windows.MessageBox]::Show(
        "Verification failed - nothing was produced.`n`n" +
        (($out.Bad | Select-Object -First 4) -join "`n") +
        "`n`nPowerPoint would show the repair prompt on that file. Your original is untouched.",
        'Deck Doctor','OK','Error') | Out-Null
      return
    }

    $newSize = (Get-Item -LiteralPath $state.OutFile).Length
    $pctSaved = if ($out.Size) { [int](100*($out.Size-$newSize)/$out.Size) } else { 0 }
    $ui.TxtResult.Text = "$(Format-Size $out.Size) -> $(Format-Size $newSize)   ($pctSaved% smaller)"
    $ui.TxtResultPath.Text = "$($state.OutFile)`n`nVerified: every relationship resolves. Open it and click through the slides that had video before you send it."
    $ui.ResultBox.Visibility = 'Visible'
    $ui.BtnOpenFolder.IsEnabled = $true
    $ui.TxtStatus.Text = ''
    $notes = @($out.Notes | Select-Object -Unique)
    if ($notes.Count) {
      $ui.TxtNotes.Text = ($notes -join "`n")
      $ui.NoteBox.Visibility = 'Visible'
    }
    Write-Log "repair OK: $($state.OutFile) ($newSize bytes)"
  }})
  $timer.Start()
}

# ------------------------------------------------------------------ wiring ----
$ui.BtnBrowse.Add_Click({ Invoke-Safe 'Browse' {
  $d = New-Object System.Windows.Forms.OpenFileDialog
  $d.Filter = 'PowerPoint decks (*.pptx;*.pptm)|*.pptx;*.pptm'
  if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Select-Deck $d.FileName }
}})
$ui.BtnAnalyse.Add_Click({ Invoke-Safe 'Analyse' { Invoke-Analyse } })
$ui.BtnSaveVideos.Add_Click({ Invoke-Safe 'Saving videos' { Save-Videos } })
$ui.BtnFillLinks.Add_Click({ Invoke-Safe 'Filling links' { Fill-Links } })
$ui.BtnBuild.Add_Click({ Invoke-Safe 'Starting the repair' { Invoke-Build } })
$ui.BtnOpenFolder.Add_Click({ Invoke-Safe 'Open folder' {
  if ($state.OutDir -and (Test-Path -LiteralPath $state.OutDir)) { Start-Process explorer.exe $state.OutDir }
}})
$ui.SldLevel.Add_ValueChanged({ Invoke-Safe 'Compression level' { Set-Level } })
$ui.DropZone.Add_Drop({
  param($sender, $e)
  Invoke-Safe 'Drop' {
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
      Select-Deck (@($e.Data.GetData([System.Windows.DataFormats]::FileDrop)))[0]
    }
  }
})
# Last line of defence: keep the window alive and record anything that got past
# Invoke-Safe, instead of the process disappearing silently.
$win.Dispatcher.Add_UnhandledException({
  param($sender, $e)
  Write-Log "UNHANDLED: $($e.Exception.GetType().Name): $($e.Exception.Message)"
  Write-Log "  $($e.Exception.StackTrace -replace "`r?`n", ' | ')"
  [System.Windows.MessageBox]::Show(
    "Something unexpected happened, but the window is still open and your original file is untouched.`n`n" +
    "$($e.Exception.Message)`n`nDetail written to:`n$($script:LogPath)",
    'Deck Doctor','OK','Error') | Out-Null
  $e.Handled = $true
})
$ui.DropZone.Add_DragOver({
  param($sender, $e)
  $e.Effects = if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { 'Copy' } else { 'None' }
  $e.Handled = $true
})
$win.Add_Closing({
  if ($state.Timer) { $state.Timer.Stop() }
  if ($state.Worker) { try { $state.Worker.Stop() } catch {} }
})

Set-Level
if ($Path) { Select-Deck $Path }
[void]$win.ShowDialog()
