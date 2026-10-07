#Requires -Version 7.4
<#
.SYNOPSIS
    Measures the time of Meeting Cleanup on a large volume, without any tenant: the steps that run on the
    computer once Microsoft Graph has answered, a whole search on the simulated tenant of the tests, and the
    window (list filled, scrolled, every meeting ticked).

.DESCRIPTION
    The engine is the one of the tool folder; the data are synthetic (meetings of 40 organizers, attendees,
    rooms). Steps measured:
      objects       the meeting and copy objects built from the Graph answers (New-MclMeeting, New-MclCopy)
      counts        the totals of a result (Update-MclResultCounts)
      plan          the plan of a Remove (Get-MclCleanupPlan)
      report        the CSV, JSON and HTML files (Export-MclReport)
      search        -Search: a whole search on the simulated tenant of the tests (tests\MeetingCleanup.FakeGraph.ps1)
      window        -Gui: list filled (Update-MclGuiRows), scrolled from top to bottom, Tick all / Untick all
    The window needs a single-threaded apartment: pwsh -STA -File .\tools\Measure-MeetingCleanup.ps1 -Gui

.PARAMETER Meetings
    Number of meetings (default 600: about 4,200 copies with 5 attendees).

.PARAMETER Attendees
    Attendees per meeting (default 5); each meeting also has its organizer and a room.

.EXAMPLE
    pwsh -STA -File .\tools\Measure-MeetingCleanup.ps1 -Meetings 600 -Search -Gui

.NOTES
    Author  : Nicolas Fabert
    Version : 1.2.2
    Part of : Meeting Cleanup (repository tool, not in the package)
#>
[CmdletBinding()]
param([int]$Meetings = 600, [int]$Attendees = 5, [switch]$Search, [switch]$Gui)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if ($Gui -and [Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') { throw 'The window needs an STA thread: pwsh -STA -File .\tools\Measure-MeetingCleanup.ps1 -Gui' }
$load = [Diagnostics.Stopwatch]::StartNew()
Import-Module (Join-Path $root 'MeetingCleanup.psd1') -Force
$load.Stop()
$module = Get-Module MeetingCleanup

$times = & $module {
    param($N, $A, $DoSearch, $DoGui, $FakePath, $LoadMs)
    $script:Quiet = $true
    $rows = [Collections.Generic.List[object]]::new()
    $time = {
        param([string]$Step, [scriptblock]$Code)
        $sw = [Diagnostics.Stopwatch]::StartNew(); $out = & $Code; $sw.Stop()
        $rows.Add([pscustomobject]@{ Step = $Step; Ms = [Math]::Round($sw.Elapsed.TotalMilliseconds) })
        $out
    }
    $rows.Add([pscustomobject]@{ Step = 'module import'; Ms = $LoadMs })
    $settings = Get-MclDefaultConfiguration
    $settings.TimeZone = 'UTC'; $settings.OutputPath = Join-Path ([IO.Path]::GetTempPath()) ('mcl-measure-' + [guid]::NewGuid().ToString('N').Substring(0, 8))

    # ---- a synthetic result: N meetings of 40 organizers, A attendees out of 2,000, one room out of 30 -------
    $list = [Collections.Generic.List[object]]::new()
    & $time "objects ($N meetings)" {
        for ($i = 0; $i -lt $N; $i++) {
            $org = "org$($i % 40)@contoso.com"
            $ev = [pscustomobject]@{ id = "E$i"; iCalUId = ('{0:X40}' -f $i); subject = "Meeting $i"; type = 'singleInstance'; organizer = [pscustomobject]@{ emailAddress = [pscustomobject]@{ address = $org; name = "Organizer $($i % 40)" } }; isOrganizer = $true
                start = [pscustomobject]@{ dateTime = '2027-01-05T09:00:00.0000000'; timeZone = 'UTC' }; end = [pscustomobject]@{ dateTime = '2027-01-05T09:30:00.0000000'; timeZone = 'UTC' }; isCancelled = $false; recurrence = $null; responseStatus = [pscustomobject]@{ response = 'organizer' }; showAs = 'busy' }
            $m = New-MclMeeting -Key $ev.iCalUId -Event $ev -Mailbox $org -Settings $settings
            $m.OrganizerKey = $org; $m.OrganizerCopy = 'Present'
            $m.Copies.Add((New-MclCopy -Key $m.MeetingId -Mailbox $org -Role 'Organizer' -Via 'Organizer calendar' -Event $ev))
            for ($j = 0; $j -lt $A; $j++) { $m.Copies.Add((New-MclCopy -Key $m.MeetingId -Mailbox "user$(($i * 7 + $j) % 2000)@contoso.com" -Role 'Attendee' -Via 'Attendee list' -Event ([pscustomobject]@{ id = "A$i-$j"; subject = $ev.subject; responseStatus = [pscustomobject]@{ response = 'accepted' }; showAs = 'busy'; isCancelled = $false }))) }
            $m.Copies.Add((New-MclCopy -Key $m.MeetingId -Mailbox "room$($i % 30)@contoso.com" -Role 'Room' -Via 'Room search' -Event ([pscustomobject]@{ id = "R$i"; subject = $ev.subject; responseStatus = [pscustomobject]@{ response = 'accepted' }; showAs = 'busy'; isCancelled = $false })))
            $list.Add($m)
        }
    }
    $organizers = @(0..39 | ForEach-Object { [pscustomobject]@{ Input = "org$_@contoso.com"; DisplayName = "Organizer $_"; PrimaryAddress = "org$_@contoso.com"; Addresses = @("org$_@contoso.com"); UserId = "u$_"; Account = 'Present'; State = 'Mailbox'; Detail = 'mailbox present' } })
    $result = [pscustomobject]@{ Tool = 'Meeting Cleanup'; Version = $script:ToolVersion; Action = 'Report'; Status = 'Completed'; Error = ''; StartedUtc = [datetime]::UtcNow.ToString('o'); CompletedUtc = ''; DurationSeconds = 0.0
        Request = [pscustomobject]@{ Mode = 'Organizers'; Room = @(); RoomFile = ''; Organizer = @($organizers.Input); Start = '2027-01-01T00:00:00Z'; End = '2027-12-31T00:00:00Z'; StartText = '2027-01-01'; EndText = '2027-12-31'; Subject = ''; MeetingId = @(); SearchIn = @('Organizer', 'Rooms'); Mailboxes = 0; MailboxFile = ''; TimeZone = 'UTC' }
        Tenant = 't'; Organization = 'contoso.onmicrosoft.com'; AppId = 'a'; AppName = 'Meeting Cleanup'; Organizers = $organizers; Searched = [pscustomobject]@{ Mailboxes = 70; Read = 70; Events = $N; NoMailbox = 0; Denied = 0; Errors = 0 }
        Meetings = $list; Warnings = [Collections.Generic.List[string]]::new(); Counts = $null }
    $copies = 0; foreach ($m in $list) { $copies += $m.Copies.Count }
    $rows[$rows.Count - 1].Step = "objects ($N meetings, $copies copies)"

    & $time 'counts (Update-MclResultCounts)' { Update-MclResultCounts $result } | Out-Null
    & $time 'plan of a Remove (Get-MclCleanupPlan)' { $null = Get-MclCleanupPlan -Result $result -Action Remove } | Out-Null
    & $time 'report (CSV, JSON, HTML)' { $null = Export-MclReport -Result $result -OutputPath $settings.OutputPath -Directory (Join-Path $settings.OutputPath 'report') } | Out-Null

    # ---- a whole search on the simulated tenant of the tests --------------------------------------------------
    if ($DoSearch) {
        Set-StrictMode -Off
        . $FakePath
        function script:Start-MclGraphSend { param($Method, $Url, $Body) [pscustomobject]@{ Task = $null; Request = $null; Response = (Invoke-FakeGraphHttp -Method $Method -Url $Url -Body $Body) } }
        Reset-FakeTenant
        for ($o = 0; $o -lt 10; $o++) { Add-FakeMailbox "org$o@contoso.test" -Name "Organizer $o" }
        for ($u = 0; $u -lt 200; $u++) { Add-FakeMailbox "user$u@contoso.test" -Name "User $u" }
        for ($r = 0; $r -lt 10; $r++) { Add-FakeMailbox "room$r@contoso.test" -Name "Room $r" -Kind Room }
        $fakeMeetings = [Math]::Max(10, [int]($N / 4))
        for ($i = 0; $i -lt $fakeMeetings; $i++) {
            $att = @(for ($j = 0; $j -lt $A; $j++) { "user$(($i * 7 + $j) % 200)@contoso.test" })
            Add-FakeMeeting -Organizer "org$($i % 10)@contoso.test" -OrganizerName "Organizer $($i % 10)" -Subject "Meeting $i" -Start ([datetime]'2030-03-01').AddHours($i).ToString('yyyy-MM-ddTHH:mm:ss') -Attendees $att -Rooms "room$($i % 10)@contoso.test" | Out-Null
        }
        $token = New-FakeToken -TenantId '11111111-2222-3333-4444-555555555555'
        $settings.TenantId = '11111111-2222-3333-4444-555555555555'
        $script:Graph = @{ Settings = $settings; Token = $token; ExpiresUtc = [datetime]::UtcNow.AddHours(1); Certificate = $null; Secret = $null; Roles = @('Calendars.ReadWrite', 'User.Read.All', 'Place.Read.All', 'GroupMember.Read.All')
            CanWrite = $true; CanRead = $true; CanReadUsers = $true; CanReadPlaces = $true; CanReadGroups = $true; TenantGuid = $settings.TenantId; AppName = 'Meeting Cleanup'; Renew = { @{ Token = $token; ExpiresUtc = [datetime]::UtcNow.AddHours(1) } } }
        Set-StrictMode -Version Latest
        $request = New-MclRequest -Settings $settings -Organizer @(0..9 | ForEach-Object { "org$_@contoso.test" }) -SearchIn Organizer, Rooms -Start ([datetime]'2030-01-01') -End ([datetime]'2030-12-31')
        Initialize-MclSteps -Total 6
        $found = & $time "search, simulated tenant ($fakeMeetings meetings)" { Find-MclMeetings -Settings $settings -Request $request }
        $rows[$rows.Count - 1].Step += " -> $($found.Counts.Copies) copies"
    }

    # ---- the window --------------------------------------------------------------------------------------------
    if ($DoGui) {
        $form = New-MclForm -Configuration $settings -Theme Light
        $w = $form.Form
        $w.WindowStartupLocation = 'Manual'; $w.Left = -4000; $w.Top = 0; $w.Width = 1320; $w.Height = 900; $w.ShowInTaskbar = $false
        $w.Show()
        $flush = { [Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] {}, [Windows.Threading.DispatcherPriority]::Background); $w.UpdateLayout() }
        & $flush
        $script:Gui.Result = $result
        & $time "window: list filled ($N rows)" { Update-MclGuiRows; & $flush } | Out-Null
        $grid = $form.Controls.Meetings
        $find = { param($v) if ($v -is [Windows.Controls.ScrollViewer]) { return $v }; for ($k = 0; $k -lt [Windows.Media.VisualTreeHelper]::GetChildrenCount($v); $k++) { $s = & $find ([Windows.Media.VisualTreeHelper]::GetChild($v, $k)); if ($s) { return $s } }; $null }
        $viewer = & $find $grid
        $scroll = @{ Pages = 0 }
        & $time 'window: list scrolled top to bottom' {
            $viewer.ScrollToTop(); & $flush
            while ($viewer.VerticalOffset + $viewer.ViewportHeight -lt $viewer.ExtentHeight - 0.5 -and $scroll.Pages -lt 2000) { $viewer.PageDown(); & $flush; $scroll.Pages++ }
        } | Out-Null
        $rows[$rows.Count - 1].Step += " ($($scroll.Pages) pages, $([int]($rows[$rows.Count - 1].Ms / [Math]::Max(1, $scroll.Pages))) ms each)"
        & $time 'window: Untick all then Tick all' { Set-MclGuiSelection $false; & $flush; Set-MclGuiSelection $true; & $flush } | Out-Null
        & $time 'window: one tick (Update-MclGuiState)' { Update-MclGuiState } | Out-Null
        $w.Close()
    }
    Remove-Item $settings.OutputPath -Recurse -Force -ErrorAction SilentlyContinue
    $rows
} $Meetings $Attendees ([bool]$Search) ([bool]$Gui) (Join-Path $root 'tests\MeetingCleanup.FakeGraph.ps1') ([Math]::Round($load.Elapsed.TotalMilliseconds))

Write-Host ("Meeting Cleanup {0} - {1} meetings x {2} attendees - PowerShell {3}" -f (Get-Module MeetingCleanup).Version, $Meetings, $Attendees, $PSVersionTable.PSVersion)
$times | ForEach-Object { '  {0,-58} {1,9:N0} ms' -f $_.Step, $_.Ms } | Write-Host
$times
