<#
.SYNOPSIS
    Renders the images of the guide and the readme: the window (light and dark, after a search, during a search,
    after a cancellation, after a restore, in rooms mode, the occurrences of a series, after a transfer) and the HTML report, from fictitious data.

.DESCRIPTION
    No tenant and no real data: the meetings come from the simulated tenant of the tests
    (tests\MeetingCleanup.FakeGraph.ps1), loaded inside the module, with contoso.com names. The window is
    rendered off screen (RenderTargetBitmap); the report is opened by Microsoft Edge headless.

    Writes docs\images\gui-search-light.png, gui-search-dark.png, gui-progress-light.png, gui-done-light.png, gui-restore-light.png,
    gui-rooms-light.png, gui-occurrences-light.png, gui-transfer-light.png, report-overview.png, report-dark.png, report-transfers.png. Needs an interactive session (WPF) and Microsoft Edge.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.3.0
#>
#Requires -Version 7.4
[CmdletBinding()]
param([string]$Destination = (Join-Path $PSScriptRoot '..\docs\images'))

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$Destination = [IO.Path]::GetFullPath($Destination)
[void][IO.Directory]::CreateDirectory($Destination)
# The reports of the images go where a real installation would put them (the path shows in the window),
# only when that folder does not exist (it is removed at the end); otherwise under artifacts\.
$neutral = Join-Path $env:SystemDrive 'Tools\MeetingCleanup'
$ownNeutral = -not (Test-Path -LiteralPath $neutral)
$ownParent = -not (Test-Path -LiteralPath (Split-Path $neutral -Parent))
$work = if ($ownNeutral) { Join-Path $neutral 'reports' } else { Join-Path $root 'artifacts\doc-images' }
if (-not $ownNeutral -and (Test-Path $work)) { Remove-Item $work -Recurse -Force }
try { [void][IO.Directory]::CreateDirectory($work) }
catch { $ownNeutral = $false; $work = Join-Path $root 'artifacts\doc-images'; if (Test-Path $work) { Remove-Item $work -Recurse -Force }; [void][IO.Directory]::CreateDirectory($work) }
# The neutral folder goes at the end, after a failure too (left behind, the next run would write under artifacts\).
$cleanup = {
    if ($ownNeutral) {
        Start-Sleep -Seconds 1
        Remove-Item -LiteralPath ($(if ($ownParent) { Split-Path $neutral -Parent } else { $neutral })) -Recurse -Force -ErrorAction SilentlyContinue
    }
}
trap { & $cleanup; break }
Import-Module (Join-Path $root 'MeetingCleanup.psd1') -Force
$module = Get-Module MeetingCleanup

& $module {
    param($FakePath, $Destination, $Work)
    Set-StrictMode -Off
    . $FakePath
    function script:Start-MclGraphSend { param($Method, $Url, $Body) [pscustomobject]@{ Task = $null; Request = $null; Response = (Invoke-FakeGraphHttp -Method $Method -Url $Url -Body $Body) } }
    $script:Quiet = $true
    # The simulated tenant lives in this runspace: the window runs its work here (not in its background runspace).
    $script:GuiInline = $true

    # ---- fictitious tenant: Megan Bowen has left, her mailbox is kept as a shared mailbox ------------------
    $seed = {
        Reset-FakeTenant
        $d = 'contoso.com'
        Add-FakeMailbox "megan.bowen@$d" -Name 'Megan Bowen' -Aliases "mbowen@$d"
        foreach ($u in 'alex.wilber', 'lidia.holloway', 'adele.vance', 'joni.sherman', 'lee.gu', 'nestor.wilke') { Add-FakeMailbox "$u@$d" -Name ((Get-Culture).TextInfo.ToTitleCase($u.Replace('.', ' '))) }
        foreach ($r in 'paris-01', 'paris-02', 'lyon-01') { Add-FakeMailbox "room-$r@$d" -Name "Room $r" -Kind Room }
        Add-FakeGroup "sales-team@$d" -Name 'Sales team' -Members "joni.sherman@$d", "lee.gu@$d"
        $year = (Get-Date).Year + 1
        $weekly = @{ pattern = @{ type = 'weekly'; interval = 1; daysOfWeek = @('monday') }; range = @{ type = 'noEnd'; startDate = "$year-01-05" } }
        Add-FakeMeeting -Organizer "megan.bowen@$d" -OrganizerName 'Megan Bowen' -Subject 'Weekly sales review' -Start "$year-01-05T08:30:00" -Attendees "alex.wilber@$d", "sales-team@$d" -Rooms "room-paris-01@$d" -Recurrence $weekly | Out-Null
        Add-FakeMeeting -Organizer "megan.bowen@$d" -OrganizerName 'Megan Bowen' -Subject 'Q1 budget workshop' -Start "$year-01-14T13:00:00" -Minutes 120 -Attendees "lidia.holloway@$d", "adele.vance@$d", "partner@fabrikam.com" -Rooms "room-paris-02@$d", "room-lyon-01@$d" | Out-Null
        Add-FakeMeeting -Organizer "megan.bowen@$d" -OrganizerName 'Megan Bowen' -Subject 'Project Atlas kick-off' -Start "$year-01-20T09:00:00" -Minutes 60 -Attendees "nestor.wilke@$d", "adele.vance@$d" -Rooms "room-lyon-01@$d" -NoOrganizerCopy | Out-Null
        Add-FakeMeeting -Organizer "megan.bowen@$d" -OrganizerName 'Megan Bowen' -Subject '1:1 Alex / Megan' -Start "$year-01-08T16:00:00" -Attendees "alex.wilber@$d" | Out-Null
        # Lynne Robbins has left too, her mailbox is deleted: her meeting is found in a room.
        Add-FakeMeeting -Organizer "lynne.robbins@$d" -OrganizerName 'Lynne Robbins' -Subject 'Supplier quarterly review' -Start "$year-02-03T14:00:00" -Minutes 60 -Attendees "adele.vance@$d", "lee.gu@$d" -Rooms "room-paris-02@$d" | Out-Null
        Add-FakeMeeting -Organizer "alex.wilber@$d" -OrganizerName 'Alex Wilber' -Subject 'Not Megan''s meeting' -Start "$year-01-08T10:00:00" -Attendees "megan.bowen@$d" -Rooms "room-paris-01@$d" | Out-Null
    }
    . $seed

    $settings = Get-MclDefaultConfiguration
    $settings.TenantId = 'contoso.onmicrosoft.com'; $settings.Organization = 'contoso.onmicrosoft.com'; $settings.AppId = '6b8e1f0a-3c52-4b8e-9a51-0f2d7c3e4a19'
    $settings.CertificateThumbprint = '3F2A9C7B1E6D4A8F0B5C2E9D7A1F4B6C8E0D2A5B'; $settings.TimeZone = 'Europe/Paris'
    $settings.OutputPath = $Work; $settings.LogPath = Join-Path $Work 'logs'; $settings.ConfigPath = 'config\MeetingCleanup.config.psd1'
    $token = New-FakeToken -TenantId '0b6c7d1e-2f3a-4b5c-8d9e-0f1a2b3c4d5e'
    $script:Graph = @{ Settings = $settings; Token = $token; ExpiresUtc = [datetime]::UtcNow.AddHours(1); Certificate = $null; Secret = $null; Roles = @('Calendars.ReadWrite', 'User.Read.All', 'Place.Read.All', 'GroupMember.Read.All')
        CanWrite = $true; CanRead = $true; CanReadUsers = $true; CanReadPlaces = $true; CanReadGroups = $true; TenantGuid = '0b6c7d1e-2f3a-4b5c-8d9e-0f1a2b3c4d5e'; AppName = 'Meeting Cleanup'; Renew = { @{ Token = $token; ExpiresUtc = [datetime]::UtcNow.AddHours(1) } } }
    function Connect-MclGraph { param($Settings, $Secret, $Action) [pscustomobject]$script:Graph }
    # Exchange Online PowerShell of the restore: the Recoverable Items of the simulated tenant.
    function script:Connect-MclExchange { param($Settings, $Secret) }
    function script:Disconnect-MclExchange { }
    function script:Get-MclPurgedItems { param([string[]]$Mailbox, [datetime]$StartUtc, [datetime]$EndUtc) Get-FakeRecoverableItems -Mailbox $Mailbox -StartUtc $StartUtc -EndUtc $EndUtc }
    function script:Restore-MclPurgedItem { param([string]$Mailbox, [string]$EntryId) Restore-FakeRecoverableItem -Mailbox $Mailbox -EntryId $EntryId }
    function script:Invoke-MclChangeMeetingOrganizer { param([string]$Mailbox, [string]$EventId, [string]$NewOrganizer, $From) Invoke-FakeChangeMeetingOrganizer -Mailbox $Mailbox -EventId $EventId -NewOrganizer $NewOrganizer }

    $render = {
        param($Form, [string]$Path)
        $w = $Form.Form
        $w.WindowStartupLocation = 'Manual'; $w.Left = -4000; $w.Top = 0; $w.Width = 1320; $w.Height = 900; $w.ShowInTaskbar = $false
        if (-not $w.IsVisible) { $w.Show() }
        [Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] {}, [Windows.Threading.DispatcherPriority]::Background)
        $w.UpdateLayout()
        $rtb = [Windows.Media.Imaging.RenderTargetBitmap]::new([int]$w.ActualWidth, [int]$w.ActualHeight, 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
        $rtb.Render($w.Content)
        $enc = [Windows.Media.Imaging.PngBitmapEncoder]::new(); $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($rtb))
        $fs = [IO.File]::Create($Path); try { $enc.Save($fs) } finally { $fs.Dispose() }
    }
    $search = {
        param([string]$Theme)
        $f = New-MclForm -Configuration $settings -Theme $Theme
        $f.Form.WindowStartupLocation = 'Manual'; $f.Form.Left = -4000; $f.Form.ShowInTaskbar = $false; $f.Form.Show()
        $f.Controls.Organizer.Text = "megan.bowen@$d" + [Environment]::NewLine + "lynne.robbins@$d"
        $f.Controls.StartDate.SelectedDate = [datetime]"$year-01-01"
        $f.Controls.EndDate.SelectedDate = [datetime]"$year-03-31"
        $f.Controls.ConnectionExpander.IsExpanded = $false
        Invoke-MclGuiSearch
        $f
    }

    $light = & $search 'Light'
    & $render $light (Join-Path $Destination 'gui-search-light.png')
    $reportFolder = $script:Gui.LastFolder
    $light.Form.Close()
    $dark = & $search 'Dark'
    & $render $dark (Join-Path $Destination 'gui-search-dark.png')
    $dark.Form.Close()

    # ---- a search in progress: the lines of a lab run (1,860 rooms), with the names of the images -----------
    $running = New-MclForm -Configuration $settings -Theme Light
    $running.Form.WindowStartupLocation = 'Manual'; $running.Form.Left = -4000; $running.Form.ShowInTaskbar = $false; $running.Form.Show()
    $running.Controls.Organizer.Text = "megan.bowen@$d"
    $running.Controls.StartDate.SelectedDate = [datetime]"$year-01-01"
    $running.Controls.EndDate.SelectedDate = [datetime]"$year-03-31"
    $running.Controls.ConnectionExpander.IsExpanded = $false
    Start-MclGuiRun 'Searching...'
    foreach ($line in @(
            @('Step', '[1/6] Microsoft Graph'), @('Info', 'Access token obtained, valid until 09:42:10 UTC.'), @('Ok', 'Application Meeting Cleanup · tenant contoso.onmicrosoft.com'),
            @('Step', '[2/6] Organizer'), @('Ok', "Megan Bowen <megan.bowen@$d> · mailbox present · 2 addresses compared"),
            @('Step', '[3/6] Mailboxes to search'), @('Info', "organizer's calendar: 1 mailbox(es)"), @('Info', 'room mailboxes: 1,860 mailbox(es)'),
            @('Step', '[4/6] Search'), @('Progress', '0.674|1,254/1,861 mailboxes searched|about 20 s left'))) { Add-MclGuiLine $line[0] $line[1] }
    & $render $running (Join-Path $Destination 'gui-progress-light.png')
    Stop-MclGuiRun
    $running.Form.Close()

    # ---- after "Cancel and clean" on three meetings -----------------------------------------------------
    $done = & $search 'Light'
    foreach ($row in @($done.Rows)) { if ($row.Subject -like '1:1*') { $row.Selected = $false } }
    $done.Controls.ActionCancel.IsChecked = $true
    $done.Controls.Comment.Text = 'Megan Bowen has left Contoso: this meeting is cancelled. Contact Alex Wilber for the follow-up.'
    Update-MclGuiState
    $script:GuiAnswers = [Collections.Generic.Queue[string]]::new()
    $script:GuiAnswers.Enqueue('Yes')
    Invoke-MclGuiApply
    $script:GuiAnswers = $null
    & $render $done (Join-Path $Destination 'gui-done-light.png')
    $doneFolder = $script:Gui.LastFolder
    $done.Form.Close()

    # ---- Remove, then Restore... of that run (the tenant reset) -------------------------------------------
    . $seed
    $restore = & $search 'Light'
    foreach ($row in @($restore.Rows)) { if ($row.Subject -like '1:1*') { $row.Selected = $false } }
    Update-MclGuiState
    $script:GuiAnswers = [Collections.Generic.Queue[string]]::new()
    $script:GuiAnswers.Enqueue('Yes')
    Invoke-MclGuiApply
    $script:GuiAnswers.Enqueue('Yes'); $script:GuiAnswers.Enqueue('Yes')
    Invoke-MclGuiRestore
    $script:GuiAnswers = $null
    $restore.Controls.Meetings.SelectedIndex = 1
    & $render $restore (Join-Path $Destination 'gui-restore-light.png')
    $restore.Form.Close()

    # ---- rooms mode: every meeting of a room in January, a series limited to its occurrences ---------------
    . $seed
    $rooms = New-MclForm -Configuration $settings -Theme 'Light'
    $rooms.Form.WindowStartupLocation = 'Manual'; $rooms.Form.Left = -4000; $rooms.Form.ShowInTaskbar = $false; $rooms.Form.Show()
    $rooms.Controls.ModeRooms.IsChecked = $true
    $rooms.Controls.Organizer.Text = "room-paris-01@$d" + [Environment]::NewLine + "room-paris-02@$d"
    $rooms.Controls.StartDate.SelectedDate = [datetime]"$year-01-05"
    $rooms.Controls.EndDate.SelectedDate = [datetime]"$year-01-16"
    $rooms.Controls.ConnectionExpander.IsExpanded = $false
    $rooms.Controls.ActionCancel.IsChecked = $true
    $rooms.Controls.Comment.Text = 'The rooms of the 2nd floor are closed for works from 5 to 16 January.'
    Invoke-MclGuiSearch
    $rooms.Controls.Meetings.SelectedIndex = 0
    & $render $rooms (Join-Path $Destination 'gui-rooms-light.png')
    $rooms.Form.Close()

    # ---- series by occurrences: the Mondays of January of the weekly review, two of them ticked ----------
    . $seed
    $series = New-MclForm -Configuration $settings -Theme 'Light'
    $series.Form.WindowStartupLocation = 'Manual'; $series.Form.Left = -4000; $series.Form.ShowInTaskbar = $false; $series.Form.Show()
    $series.Controls.Organizer.Text = "megan.bowen@$d"
    $series.Controls.Subject.Text = 'Weekly sales review'
    $series.Controls.StartDate.SelectedDate = [datetime]"$year-01-01"
    $series.Controls.EndDate.SelectedDate = [datetime]"$year-01-31"
    $series.Controls.ConnectionExpander.IsExpanded = $false
    $series.Controls.SeriesOccurrences.IsChecked = $true
    $series.Controls.ActionCancel.IsChecked = $true
    $series.Controls.Comment.Text = 'No sales review on these Mondays (inventory).'
    Update-MclGuiState
    Invoke-MclGuiSearch
    $series.Controls.Meetings.SelectedIndex = 0
    Show-MclGuiOccurrences
    $occurrences = @($script:Gui.OccurrenceRows)
    for ($i = 0; $i -lt $occurrences.Count; $i++) { $occurrences[$i].Selected = $i -in 1, 2 }
    Update-MclGuiOccurrenceCount
    & $render $series (Join-Path $Destination 'gui-occurrences-light.png')
    Hide-MclGuiOccurrences
    $series.Form.Close()

    # ---- transfer: Megan (mailbox kept) moved by Exchange Online, Lynne (deleted) re-created ---------------
    . $seed
    $transfer = & $search 'Light'
    foreach ($row in @($transfer.Rows)) { if ($row.Subject -notin 'Q1 budget workshop', 'Supplier quarterly review') { $row.Selected = $false } }
    $transfer.Controls.ActionTransfer.IsChecked = $true
    $transfer.Controls.NewOrganizer.Text = "alex.wilber@$d"
    Update-MclGuiState
    $script:GuiAnswers = [Collections.Generic.Queue[string]]::new()
    $script:GuiAnswers.Enqueue('Yes')
    Invoke-MclGuiApply
    $script:GuiAnswers = $null
    $transfer.Controls.Meetings.SelectedIndex = 4
    & $render $transfer (Join-Path $Destination 'gui-transfer-light.png')
    $transferFolder = $script:Gui.LastFolder
    $transfer.Form.Close()
    [pscustomobject]@{ Report = (Join-Path $reportFolder 'MeetingCleanup.html'); Done = (Join-Path $doneFolder 'MeetingCleanup.html'); Transfer = (Join-Path $transferFolder 'MeetingCleanup.html') }
} (Join-Path $root 'tests\MeetingCleanup.FakeGraph.ps1') $Destination $work | Set-Variable reports

# ---- the HTML report, opened by Microsoft Edge headless ------------------------------------------------
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge is needed for the images of the report.' }
# The Transfer report opens on its Transfers tab: its tables only (scoutFocus=tables).
foreach ($shot in @(@{ Html = $reports.Report; Theme = 'light'; File = 'report-overview.png' }, @{ Html = $reports.Done; Theme = 'dark'; File = 'report-dark.png' },
        @{ Html = $reports.Transfer; Theme = 'light'; File = 'report-transfers.png'; Query = '&scoutFocus=tables'; Size = '1360,560' })) {
    $profile = Join-Path $work "edge-$([guid]::NewGuid().ToString('N'))"
    $url = ([Uri]$shot.Html).AbsoluteUri + "?scoutTheme=$($shot.Theme)" + $(if ($shot.Query) { $shot.Query } else { '' })
    $size = if ($shot.Size) { $shot.Size } else { '1360,1180' }
    $png = Join-Path $Destination $shot.File
    $before = if (Test-Path -LiteralPath $png) { (Get-Item -LiteralPath $png).LastWriteTimeUtc } else { [datetime]::MinValue }
    # Edge headless sometimes stays open after writing its screenshot: waited for 60 s at most, then stopped.
    $p = Start-Process -FilePath $edge -PassThru -WindowStyle Hidden -ArgumentList @('--headless=new', '--disable-gpu', '--hide-scrollbars', "--user-data-dir=`"$profile`"", "--window-size=$size", "--screenshot=`"$png`"", "`"$url`"")
    if (-not $p.WaitForExit(60000)) {
        if (-not (Test-Path -LiteralPath $png) -or (Get-Item -LiteralPath $png).LastWriteTimeUtc -le $before) { Write-Warning "Edge did not write $($shot.File) within 60 s." }
        Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Milliseconds 500
}
& $cleanup
Get-ChildItem -LiteralPath $Destination -Filter '*.png' | Sort-Object Name | ForEach-Object { '{0,10:N0}  {1}' -f $_.Length, $_.Name }
