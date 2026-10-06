#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.1.0' }
<#
    Meeting Cleanup - automated tests (Pester 6.1 or later).
    Author  : Nicolas Fabert
    Version : 1.2.0

    Run:  .\Run-Tests.ps1      (or Invoke-Pester -Path .\tests -Output Detailed)

    No tenant is needed: Start-MclGraphSend is replaced by a simulated Exchange Online tenant
    (tests\MeetingCleanup.FakeGraph.ps1) that answers like Graph did in the lab: same iCalUId in every copy,
    room copies with the organizer's name as subject, the attendee list in every copy, a filter on organizer
    refused, a cancellation sent whenever the organizer's meeting is removed or cancelled.
#>

BeforeAll {
    $script:Root = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $script:Root 'MeetingCleanup.psd1') -Force
    $script:Module = Get-Module MeetingCleanup
    . (Join-Path $PSScriptRoot 'MeetingCleanup.FakeGraph.ps1')
    & $script:Module { $script:Quiet = $true }

    $script:Tenant = '11111111-2222-3333-4444-555555555555'
    function New-TestSettings([hashtable]$Overrides = @{}) {
        $s = & $script:Module { Get-MclDefaultConfiguration }
        $s.TenantId = $script:Tenant; $s.AppId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'; $s.CertificateThumbprint = ('AB' * 20)
        $s.TimeZone = 'UTC'; $s.MaxRetries = 3; $s.MaxConcurrency = 4
        $s.OutputPath = Join-Path $script:Root 'artifacts\test-reports'; $s.LogPath = Join-Path $script:Root 'artifacts\test-logs'
        foreach ($k in $Overrides.Keys) { $s[$k] = $Overrides[$k] }
        return $s
    }
    function Connect-Test([hashtable]$Settings, [string[]]$Roles, [string]$Action = 'Report') {
        $token = if ($Roles) { New-FakeToken -TenantId $script:Tenant -Roles $Roles } else { New-FakeToken -TenantId $script:Tenant }
        Mock -ModuleName MeetingCleanup Get-MclCertificate { [pscustomobject]@{ Thumbprint = 'AB'; NotAfter = (Get-Date).AddYears(1) } }
        Mock -ModuleName MeetingCleanup Get-MclAppToken { @{ Token = $token; ExpiresUtc = [datetime]::UtcNow.AddHours(1) } }.GetNewClosure()
        Connect-MclGraph -Settings $Settings -Action $Action
    }
    function New-TestRequest([hashtable]$Settings, [string[]]$Organizer, [string[]]$SearchIn, [hashtable]$More = @{}) {
        $a = @{ Settings = $Settings; Organizer = $Organizer; SearchIn = $SearchIn; Start = [datetime]'2030-01-01'; End = [datetime]'2030-12-31' }
        foreach ($k in $More.Keys) { $a[$k] = $More[$k] }
        New-MclRequest @a
    }
    function Find-Test([string[]]$Organizer, [string[]]$SearchIn, [hashtable]$More = @{}, [hashtable]$Overrides = @{}) {
        $s = New-TestSettings $Overrides
        $null = Connect-Test $s
        & $script:Module { Initialize-MclSteps -Total 6 }
        Find-MclMeetings -Settings $s -Request (New-TestRequest $s $Organizer $SearchIn $More)
    }
    function Get-Meeting($Result, [string]$Subject) { $Result.Meetings | Where-Object { $_.Subject -like "$Subject*" } | Select-Object -First 1 }
    function Get-CopyOf($Meeting, [string]$Mailbox) { $Meeting.Copies | Where-Object Mailbox -eq $Mailbox | Select-Object -First 1 }

    # Every Graph request goes to the simulated tenant, and so do the cmdlets of Exchange Online PowerShell.
    Mock -ModuleName MeetingCleanup Start-MclGraphSend { [pscustomobject]@{ Task = $null; Request = $null; Response = (Invoke-FakeGraphHttp -Method $Method -Url $Url -Body $Body) } }
    Mock -ModuleName MeetingCleanup Connect-MclExchange { }
    Mock -ModuleName MeetingCleanup Disconnect-MclExchange { }
    Mock -ModuleName MeetingCleanup Get-MclPurgedItems { Get-FakeRecoverableItems -Mailbox $Mailbox -StartUtc $StartUtc -EndUtc $EndUtc }
    Mock -ModuleName MeetingCleanup Restore-MclPurgedItem { Restore-FakeRecoverableItem -Mailbox $Mailbox -EntryId $EntryId }
    function Invoke-TestCleanup([pscustomobject]$Result, [string]$Action = 'Remove', [string]$Comment = 'x') {
        $s = New-TestSettings
        $dir = Join-Path $s.OutputPath ('run-' + [guid]::NewGuid().ToString('N'))
        & $script:Module { Initialize-MclSteps -Total 2 }
        $done = Invoke-MclCleanup -Settings $s -Result $Result -Action $Action -Comment $Comment -BackupPath (Join-Path $dir 'MeetingCleanup-Backup.json')
        $report = Export-MclReport -Result $done -OutputPath $s.OutputPath -Directory $dir
        [pscustomobject]@{ Result = $done; Folder = $report.Directory }
    }
    function Invoke-TestRestore([string]$Folder, [string[]]$MeetingId) {
        $s = New-TestSettings
        $src = Import-MclRestoreSource -Path $Folder -MeetingId $MeetingId
        & $script:Module { Initialize-MclSteps -Total 3 }
        Invoke-MclRestore -Settings $s -Result $src
    }

    Mock -ModuleName MeetingCleanup Invoke-MclChangeMeetingOrganizer {
        if ($script:Fake.ContainsKey('NativeError') -and $script:Fake.NativeError) { return @{ Ok = $false; Error = $script:Fake.NativeError } }
        Invoke-FakeChangeMeetingOrganizer -Mailbox $Mailbox -EventId $EventId -NewOrganizer $NewOrganizer
    }
    function Find-Rooms([string[]]$Room, [datetime]$Start, [datetime]$End, [hashtable]$More = @{}) {
        $s = New-TestSettings
        $null = Connect-Test $s
        & $script:Module { Initialize-MclSteps -Total 6 }
        $a = @{ Settings = $s; Room = $Room; Start = $Start; End = $End }
        foreach ($k in $More.Keys) { $a[$k] = $More[$k] }
        Find-MclMeetings -Settings $s -Request (New-MclRequest @a)
    }
    function Invoke-TestTransfer([pscustomobject]$Result, [string]$NewOrganizer, [string]$Method = 'Auto', [datetime]$From = [datetime]'2030-01-01') {
        $s = New-TestSettings
        $dir = Join-Path $s.OutputPath ('transfer-' + [guid]::NewGuid().ToString('N'))
        & $script:Module { Initialize-MclSteps -Total 3 }
        $new = Resolve-MclNewOrganizer -Address $NewOrganizer
        $plan = Get-MclTransferPlan -Result $Result -NewOrganizer $new -Method $Method -From $From -Comment $s.TransferComment
        $done = Invoke-MclTransfer -Settings $s -Result $Result -Plan $plan -From $From -Comment $s.TransferComment -BackupPath (Join-Path $dir 'MeetingCleanup-Backup.json')
        $report = Export-MclReport -Result $done -OutputPath $s.OutputPath -Directory $dir
        [pscustomobject]@{ Result = $done; Folder = $report.Directory; Plan = $plan }
    }
    function Get-FakeEvents([string]$Mailbox, [string]$ICalUId) { @($script:Fake.Mailboxes[$Mailbox].Events | Where-Object iCalUId -eq $ICalUId) }
    # The tenant of most tests: an organizer with an alias, three attendees (one through a group), two
    # rooms, an external attendee, a departed organizer (no mailbox, no user) with and without room.
    function New-TestTenant {
        Reset-FakeTenant
        Add-FakeMailbox 'org@contoso.test' -Name 'Olivia Organizer' -Aliases 'olivia@contoso.test'
        Add-FakeMailbox 'att1@contoso.test' -Name 'Attendee 1' -Aliases 'att1.alias@contoso.test'
        Add-FakeMailbox 'att2@contoso.test' -Name 'Attendee 2'
        Add-FakeMailbox 'att3@contoso.test' -Name 'Attendee 3'
        Add-FakeMailbox 'room1@contoso.test' -Name 'Room 1' -Kind Room
        Add-FakeMailbox 'room2@contoso.test' -Name 'Room 2' -Kind Room
        Add-FakeGroup 'dl@contoso.test' -Name 'Team DL' -Members 'att3@contoso.test'
        $script:Ids = @{}
        $script:Ids.M1 = Add-FakeMeeting -Organizer 'org@contoso.test' -OrganizerName 'Olivia Organizer' -Subject 'M1 Budget review' -Start '2030-03-10T10:00:00' -Attendees 'att1@contoso.test', 'att2@contoso.test', 'ext@fabrikam.test', 'dl@contoso.test' -Rooms 'room1@contoso.test'
        $script:Ids.S1 = Add-FakeMeeting -Organizer 'org@contoso.test' -OrganizerName 'Olivia Organizer' -Subject 'S1 Weekly' -Start '2030-03-11T09:00:00' -Attendees 'att1@contoso.test' -Rooms 'room2@contoso.test' -Recurrence @{ pattern = @{ type = 'weekly'; interval = 1; daysOfWeek = @('monday') }; range = @{ type = 'numbered'; numberOfOccurrences = 4; startDate = '2030-03-11' } }
        $script:Ids.O1 = Add-FakeMeeting -Organizer 'org@contoso.test' -OrganizerName 'Olivia Organizer' -Subject 'O1 Orphan in a room' -Start '2030-04-01T08:00:00' -Attendees 'att2@contoso.test' -Rooms 'room1@contoso.test' -NoOrganizerCopy
        $script:Ids.Old = Add-FakeMeeting -Organizer 'org@contoso.test' -Subject 'Old series' -Start '2029-01-01T08:00:00' -Attendees 'att1@contoso.test' -Recurrence @{ pattern = @{ type = 'weekly'; interval = 1; daysOfWeek = @('monday') }; range = @{ type = 'endDate'; startDate = '2029-01-01'; endDate = '2029-06-30' } }
        $script:Ids.Past = Add-FakeMeeting -Organizer 'org@contoso.test' -Subject 'Past numbered series' -Start '2029-12-01T08:00:00' -Attendees 'att2@contoso.test' -Recurrence @{ pattern = @{ type = 'daily'; interval = 1 }; range = @{ type = 'numbered'; numberOfOccurrences = 3; startDate = '2029-12-01' } }
        $script:Ids.Appt = Add-FakeMeeting -Organizer 'org@contoso.test' -Subject 'Dentist' -Start '2030-03-15T08:00:00'
        $script:Ids.D1 = Add-FakeMeeting -Organizer 'gone@contoso.test' -OrganizerName 'Gary Gone' -Subject 'D1 Departed with room' -Start '2030-05-01T09:00:00' -Attendees 'att1@contoso.test' -Rooms 'room2@contoso.test'
        $script:Ids.D2 = Add-FakeMeeting -Organizer 'gone@contoso.test' -OrganizerName 'Gary Gone' -Subject 'D2 Departed without room' -Start '2030-05-02T09:00:00' -Attendees 'att2@contoso.test'
        $script:Ids.Other = Add-FakeMeeting -Organizer 'att2@contoso.test' -OrganizerName 'Attendee 2' -Subject 'Someone else' -Start '2030-03-10T14:00:00' -Attendees 'att1@contoso.test' -Rooms 'room1@contoso.test'
    }
}

Describe 'Configuration and request' {
    It 'reads the delivered configuration and resolves the paths from the tool folder' {
        $c = Import-MclConfiguration
        $c.OutputPath | Should -Be (Join-Path $script:Root 'reports')
        $c.LogPath | Should -Be (Join-Path $script:Root 'logs')
        $c.SearchIn | Should -Be @('Organizer', 'Rooms')
        $c.MaxConcurrency | Should -Be 16
        $c.Verify | Should -BeTrue
    }

    It 'lists unknown sections, unknown keys and invalid values together' {
        $path = Join-Path $script:Root 'artifacts\bad.config.psd1'
        [void][IO.Directory]::CreateDirectory((Split-Path $path))
        "@{ Tenant = @{ TenantId = 'not a tenant' }; Search = @{ SearchIn = @('Rooms', 'Nowhere'); Futur = 1 }; Extra = @{}; Graph = @{ MaxConcurrency = 99 }; Report = @{ CsvDelimiter = '|' } }" | Set-Content $path
        $text = ({ Import-MclConfiguration -Path $path } | Should -Throw -PassThru).Exception.Message
        $text | Should -Match "Unknown key 'Search.Futur'"
        $text | Should -Match "Unknown section 'Extra'"
        $text | Should -Match 'Tenant.TenantId must be'
        $text | Should -Match 'Search.SearchIn must contain'
        $text | Should -Match 'Graph.MaxConcurrency must be'
        $text | Should -Match 'Report.CsvDelimiter must be'
        Remove-Item $path
    }

    It 'requires the tenant and the application only to connect' {
        $s = New-TestSettings @{ TenantId = ''; AppId = '' }
        (Test-MclConfiguration -Configuration $s).IsValid | Should -BeTrue
        $p = (Test-MclConfiguration -Configuration $s -ForConnection).Problems
        $p | Should -Contain 'Tenant.TenantId is required.'
        $p | Should -Contain 'Authentication.AppId is required.'
    }

    It 'checks the request: organizer, X500, period, list of mailboxes, meeting IDs, replay' {
        $s = New-TestSettings
        (Test-MclRequest (New-TestRequest $s @() @('Rooms'))).Problems -join ' ' | Should -Match 'Give the organizer'
        (Test-MclRequest (New-TestRequest $s @('not-an-address') @('Rooms'))).Problems -join ' ' | Should -Match 'neither an SMTP address'
        (Test-MclRequest (New-TestRequest $s @('/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=abc-Gary') @('Rooms'))).IsValid | Should -BeTrue
        (Test-MclRequest (New-TestRequest $s @('a@b.test') @('Rooms') @{ Start = [datetime]'2030-02-01'; End = [datetime]'2030-01-01' })).Problems -join ' ' | Should -Match 'end of the period'
        (Test-MclRequest (New-TestRequest $s @('a@b.test') @('Mailboxes'))).Problems -join ' ' | Should -Match 'Mailboxes: give the list'
        (Test-MclRequest (New-TestRequest $s @('a@b.test') @('Rooms') @{ MeetingId = @('xyz') })).Problems -join ' ' | Should -Match 'is not an iCalUId'
        (Test-MclRequest (New-MclRequest -Settings $s -FromReport $script:Root -Action Report)).Problems -join ' ' | Should -Match 'choose -Action Remove, Cancel, Transfer or Restore'
        (Test-MclRequest (New-MclRequest -Settings $s -Action Restore)).Problems -join ' ' | Should -Match 'Restore needs -FromReport'
    }

    It 'splits the organizers typed in one text, keeping an X500 address whole' {
        $r = New-MclRequest -Settings (New-TestSettings) -Organizer 'a@contoso.test; b@contoso.test,c@contoso.test', '/o=Org/ou=Group (X)/cn=Recipients/cn=abc-Name with space'
        $r.Organizer | Should -Be @('a@contoso.test', 'b@contoso.test', 'c@contoso.test', '/o=Org/ou=Group (X)/cn=Recipients/cn=abc-Name with space')
    }

    It 'includes the end date and reads the dates in the time zone of the report' {
        $s = New-TestSettings @{ TimeZone = 'Europe/Paris' }
        $r = New-TestRequest $s @('a@b.test') @('Rooms') @{ Start = [datetime]'2030-07-01'; End = [datetime]'2030-07-31' }
        $r.Start | Should -Be ([datetime]::SpecifyKind([datetime]'2030-06-30T22:00:00', 'Utc'))
        $r.End | Should -Be ([datetime]::SpecifyKind([datetime]'2030-07-31T22:00:00', 'Utc'))
        & $script:Module { param($e) Format-MclDate $e 'Europe/Paris' -PeriodEnd } $r.End | Should -Be '2030-07-31'
    }

    It 'reads a text file or a CSV file of addresses' {
        $dir = Join-Path $script:Root 'artifacts'; [void][IO.Directory]::CreateDirectory($dir)
        $txt = Join-Path $dir 'mailboxes.txt'; "# list`r`nA@contoso.test`r`n`r`nb@contoso.test ; comment" | Set-Content $txt
        $csv = Join-Path $dir 'mailboxes.csv'; "DisplayName;PrimarySmtpAddress`r`nA;a@contoso.test`r`nC;c@contoso.test" | Set-Content $csv
        & $script:Module { param($p) Read-MclAddressFile $p } $txt | Should -Be @('A@contoso.test', 'b@contoso.test')
        & $script:Module { param($p) Read-MclAddressFile $p } $csv | Should -Be @('a@contoso.test', 'c@contoso.test')
    }

    It 'describes a recurrence in words' {
        $text = & $script:Module { Format-MclRecurrence @{ pattern = @{ type = 'weekly'; interval = 1; daysOfWeek = @('tuesday') }; range = @{ type = 'numbered'; numberOfOccurrences = 6; startDate = '2030-10-13' } } }
        $text | Should -Be 'Weekly (Tuesday), 6 occurrences from 2030-10-13'
        & $script:Module { Format-MclRecurrence @{ pattern = @{ type = 'daily'; interval = 2 }; range = @{ type = 'endDate'; startDate = '2030-01-01'; endDate = '2030-02-01' } } } | Should -Be 'every 2 days, from 2030-01-01 until 2030-02-01'
    }
}

Describe 'Microsoft Graph connection and requests' {
    It 'stops when the token belongs to another tenant' {
        $s = New-TestSettings
        Mock -ModuleName MeetingCleanup Get-MclCertificate { [pscustomobject]@{ Thumbprint = 'AB'; NotAfter = (Get-Date).AddYears(1) } }
        Mock -ModuleName MeetingCleanup Get-MclAppToken { @{ Token = (New-FakeToken -TenantId '99999999-2222-3333-4444-555555555555'); ExpiresUtc = [datetime]::UtcNow.AddHours(1) } }
        { Connect-MclGraph -Settings $s } | Should -Throw '*Connected to tenant 99999999*Nothing was read*'
    }

    It 'requires Calendars.ReadWrite to change anything, Calendars.Read to report' {
        $s = New-TestSettings
        { Connect-Test $s @('User.Read.All') } | Should -Throw '*no application permission Calendars.ReadWrite*'
        $c = Connect-Test $s @('Calendars.Read')
        $c.CanRead | Should -BeTrue
        $c.CanWrite | Should -BeFalse
        { Connect-Test $s @('Calendars.Read') 'Remove' } | Should -Throw '*Calendars.Read only*'
        $c = Connect-Test $s @('Calendars.ReadWrite')
        $c.CanReadUsers, $c.CanReadPlaces, $c.CanReadGroups | Should -Be @($false, $false, $false)
    }

    It 'sends 20 requests per $batch and never more than 4 for one mailbox' {
        New-TestTenant
        $null = Connect-Test (New-TestSettings)
        $requests = @(1..10 | ForEach-Object { & $script:Module { param($i) New-MclGraphRequest -Id "a$i" -Url '/users/att1%40contoso.test/calendar' } $_ }) +
            @(1..30 | ForEach-Object { & $script:Module { param($i) New-MclGraphRequest -Id "r$i" -Url "/users/room$($i % 2 + 1)%40contoso.test/calendar" } $_ })
        $res = & $script:Module { param($r) Invoke-MclGraphBatch -Requests $r } $requests
        $res.Count | Should -Be 40
        @($res.Values | Where-Object Status -eq 200).Count | Should -Be 40
        foreach ($b in $script:Fake.Batches) {
            $b.Count | Should -BeLessOrEqual 20
            ($b.Mailboxes | Group-Object | Measure-Object Count -Maximum).Maximum | Should -BeLessOrEqual 4
        }
    }

    It 'retries a throttled request after Retry-After and follows the pages of a list' {
        New-TestTenant
        $mb = $script:Fake.Mailboxes['att3@contoso.test']
        1..25 | ForEach-Object { Add-FakeMeeting -Organizer 'att2@contoso.test' -Subject "Paged $_" -Start '2030-06-01T08:00:00' -Attendees 'att3@contoso.test' | Out-Null }
        $script:Fake.Throttle['*/users/att3@contoso.test/events?*'] = 2
        $null = Connect-Test (New-TestSettings)
        $res = & $script:Module { Invoke-MclGraphBatch -Requests @(New-MclGraphRequest -Id 'p' -Url '/users/att3%40contoso.test/events?$top=10') -FollowPages }
        $res['p'].Status | Should -Be 200
        $res['p'].Values.Count | Should -Be ($mb.Events.Count)
        @($script:Fake.Calls | Where-Object Url -like '*att3@contoso.test/events*').Count | Should -Be 5   # 2 throttled + 3 pages
    }
}

Describe 'Search' {
    BeforeEach { New-TestTenant }

    It 'organizer present: its meetings, the copy of every attendee, room and group member' {
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms'
        @($r.Meetings.Subject) | Should -Be @('M1 Budget review', 'S1 Weekly', 'O1 Orphan in a room')
        $m1 = Get-Meeting $r 'M1'
        $m1.OrganizerCopy | Should -Be 'Present'
        @($m1.Copies | Where-Object EventId | Sort-Object Mailbox | ForEach-Object { "$($_.Mailbox)=$($_.Role)" }) | Should -Be @('att1@contoso.test=Attendee', 'att2@contoso.test=Attendee', 'att3@contoso.test=Attendee', 'org@contoso.test=Organizer', 'room1@contoso.test=Room')
        (Get-CopyOf $m1 'att3@contoso.test').Via | Should -Be 'Group Team DL'
        (Get-CopyOf $m1 'dl@contoso.test').Result | Should -Be 'Expanded'
        (Get-CopyOf $m1 'ext@fabrikam.test').Result | Should -Be 'Not processed'
        (Get-CopyOf $m1 'ext@fabrikam.test').Detail | Should -Match 'not a mailbox of this tenant'
        $r.Counts.Meetings | Should -Be 3
        $r.Status | Should -Be 'Completed'
    }

    It 'keeps a series with an occurrence in the period, leaves out ended series and appointments' {
        $r = Find-Test 'org@contoso.test' 'Organizer'
        @($r.Meetings.Subject) | Should -Not -Contain 'Old series'
        @($r.Meetings.Subject) | Should -Not -Contain 'Past numbered series'
        @($r.Meetings.Subject) | Should -Not -Contain 'Dentist'
        $s1 = Get-Meeting $r 'S1'
        $s1.Kind | Should -Be 'Series'
        $s1.NextInPeriod | Should -Be '2030-03-11 09:00'
        $s1.Recurrence | Should -Be 'Weekly (Monday), 4 occurrences from 2030-03-11'
    }

    It 'a meeting removed from the organizer calendar is found in the room: organizer copy Absent' {
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms'
        $o1 = Get-Meeting $r 'O1'
        $o1.OrganizerCopy | Should -Be 'Absent'
        $o1.Subject | Should -Be 'O1 Orphan in a room'   # the room shows the organizer's name, the attendee copy the subject
        @($o1.Copies | Where-Object EventId).Mailbox | Sort-Object | Should -Be @('att2@contoso.test', 'room1@contoso.test')
    }

    It 'departed organizer: found from the rooms, every copy found from the attendee list' {
        $r = Find-Test 'gone@contoso.test' 'Organizer', 'Rooms'
        $r.Organizers[0].State | Should -Be 'NotInDirectory'
        @($r.Meetings.Subject) | Should -Be @('D1 Departed with room')
        $d1 = $r.Meetings[0]
        $d1.OrganizerCopy | Should -Be 'Mailbox deleted'
        @($d1.Copies | Where-Object EventId).Mailbox | Sort-Object | Should -Be @('att1@contoso.test', 'room2@contoso.test')
    }

    It 'departed organizer without a room: found by searching every mailbox' {
        $r = Find-Test 'gone@contoso.test' 'AllMailboxes'
        @($r.Meetings.Subject) | Should -Be @('D1 Departed with room', 'D2 Departed without room')
        (Get-Meeting $r 'D1').Copies | Where-Object { $_.Role -eq 'Room' -and $_.EventId } | Should -Not -BeNullOrEmpty
    }

    It 'searches a list of mailboxes' {
        $r = Find-Test 'gone@contoso.test' 'Mailboxes' @{ Mailbox = @('att2@contoso.test') }
        @($r.Meetings.Subject) | Should -Be @('D2 Departed without room')
    }

    It 'matches the X500 address shown by the copies of a deleted mailbox' {
        $x500 = '/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=0f1e2d-Gary Gone'
        Add-FakeMeeting -Organizer 'nobody@contoso.test' -OrganizerName 'Gary Gone' -OrganizerAddressInCopies $x500.ToLowerInvariant() -Subject 'X1 Shown with X500' -Start '2030-06-01T10:00:00' -Attendees 'att1@contoso.test' -Rooms 'room1@contoso.test' | Out-Null
        $r = Find-Test $x500 'Rooms'
        @($r.Meetings.Subject) | Should -Be @('X1 Shown with X500')
        $r.Meetings[0].OrganizerCopy | Should -Be 'Not checked'
    }

    It 'compares every alias of the organizer, and counts once a copy reached through an alias' {
        Add-FakeMeeting -Organizer 'org@contoso.test' -OrganizerName 'Olivia Organizer' -OrganizerAddressInCopies 'olivia@contoso.test' -Subject 'A1 Sent from the alias' -Start '2030-07-01T10:00:00' -Attendees 'att1.alias@contoso.test' -Rooms 'room1@contoso.test' -NoOrganizerCopy | Out-Null
        $r = Find-Test 'org@contoso.test' 'Rooms', 'AllMailboxes' @{ Subject = 'A1' }
        $a1 = Get-Meeting $r 'A1'
        @($a1.Copies | Where-Object EventId).Mailbox | Sort-Object | Should -Be @('att1@contoso.test', 'room1@contoso.test')
    }

    It 'one meeting: by subject (the real one, not the room name) or by meeting ID' {
        $r = Find-Test 'org@contoso.test' 'Rooms' @{ Subject = 'budget' }
        @($r.Meetings.Subject) | Should -Be @('M1 Budget review')
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms' @{ MeetingId = @($script:Ids.S1) }
        @($r.Meetings.Subject) | Should -Be @('S1 Weekly')
    }

    It 'says what it cannot do without the optional permissions' {
        $s = New-TestSettings
        $null = Connect-Test $s @('Calendars.ReadWrite')
        & $script:Module { Initialize-MclSteps -Total 6 }
        $r = Find-MclMeetings -Settings $s -Request (New-TestRequest $s 'org@contoso.test' 'Organizer', 'Rooms')
        $r.Status | Should -Be 'Warning'
        ($r.Warnings -join ' ') | Should -Match 'Place.Read.All'
        ($r.Warnings -join ' ') | Should -Match 'GroupMember.Read.All'
        (Get-Meeting $r 'M1').Copies | Where-Object Mailbox -eq 'att3@contoso.test' | Should -BeNullOrEmpty
    }
}

Describe 'Remove and cancel' {
    BeforeEach { New-TestTenant }

    It 'Remove: the copies of the attendees and the rooms go, the organizer keeps the meeting, no message' {
        $s = New-TestSettings
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms'
        & $script:Module { Initialize-MclSteps -Total 2 }
        $r = Invoke-MclCleanup -Settings $s -Result $r -Action Remove
        $script:Fake.Messages.Count | Should -Be 0
        $m1 = Get-Meeting $r 'M1'
        $m1.Status | Should -Be 'Removed'
        (Get-CopyOf $m1 'org@contoso.test').Result | Should -Be 'Kept'
        @($m1.Copies | Where-Object { $_.Role -in 'Attendee', 'Room' -and $_.EventId } | ForEach-Object Result | Select-Object -Unique) | Should -Be @('Removed')
        @($m1.Copies | Where-Object { $_.Result -eq 'Removed' } | ForEach-Object Verified | Select-Object -Unique) | Should -Be @('Yes')
        @($script:Fake.Mailboxes['room1@contoso.test'].Events | Where-Object iCalUId -eq $script:Ids.M1).Count | Should -Be 0
        @($script:Fake.Mailboxes['org@contoso.test'].Events | Where-Object iCalUId -eq $script:Ids.M1).Count | Should -Be 1
        @($script:Fake.Mailboxes['room1@contoso.test'].Events | Where-Object iCalUId -eq $script:Ids.Other).Count | Should -Be 1
        $r.Counts.Kept | Should -Be 2
        $r.Status | Should -Be 'Completed'
    }

    It 'Cancel: the organizer cancels with the message, then the copies left are removed' {
        $s = New-TestSettings
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms'
        & $script:Module { Initialize-MclSteps -Total 2 }
        $r = Invoke-MclCleanup -Settings $s -Result $r -Action Cancel -Comment 'Olivia has left.'
        @($script:Fake.Messages | Where-Object Kind -eq 'Cancellation').Count | Should -Be 2   # M1 and S1; O1 has no organizer copy
        $script:Fake.Messages[0].Comment | Should -Be 'Olivia has left.'
        (Get-Meeting $r 'M1').Status | Should -Be 'Cancelled'
        (Get-CopyOf (Get-Meeting $r 'M1') 'org@contoso.test').Result | Should -Be 'Cancelled'
        (Get-CopyOf (Get-Meeting $r 'M1') 'room1@contoso.test').Result | Should -Be 'Already gone'   # the room processed the cancellation
        (Get-CopyOf (Get-Meeting $r 'M1') 'att1@contoso.test').Result | Should -Be 'Removed'
        $o1 = Get-Meeting $r 'O1'
        $o1.Status | Should -Be 'Removed'
        ($o1.Notes -join ' ') | Should -Match 'without a message'
        foreach ($mb in 'org@contoso.test', 'att1@contoso.test', 'att2@contoso.test', 'att3@contoso.test', 'room1@contoso.test', 'room2@contoso.test') {
            @($script:Fake.Mailboxes[$mb].Events | Where-Object { $_.iCalUId -in $script:Ids.M1, $script:Ids.S1, $script:Ids.O1 }).Count | Should -Be 0
        }
    }

    It 'Cancel refused for a meeting: its copies are left as they were' {
        $s = New-TestSettings
        $r = Find-Test 'org@contoso.test' 'Organizer' @{ Subject = 'M1' }
        $script:Fake.Fail['*/cancel'] = @{ Status = 403; Code = 'ErrorAccessDenied' }
        & $script:Module { Initialize-MclSteps -Total 2 }
        $r = Invoke-MclCleanup -Settings $s -Result $r -Action Cancel -Comment 'x'
        $m1 = $r.Meetings[0]
        $m1.Status | Should -Be 'Failed'
        (Get-CopyOf $m1 'att1@contoso.test').Result | Should -Be 'Not done'
        @($script:Fake.Mailboxes['att1@contoso.test'].Events | Where-Object iCalUId -eq $script:Ids.M1).Count | Should -Be 1
        $r.Status | Should -Be 'Failed'
    }

    It 'a copy already gone is not an error; an unticked meeting is not touched' {
        $s = New-TestSettings
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms'
        $script:Fake.Mailboxes['att1@contoso.test'].Events.RemoveAll({ param($e) $e.iCalUId -eq $script:Ids.M1 }) | Out-Null
        (Get-Meeting $r 'S1').Selected = $false
        & $script:Module { Initialize-MclSteps -Total 2 }
        $r = Invoke-MclCleanup -Settings $s -Result $r -Action Remove
        (Get-CopyOf (Get-Meeting $r 'M1') 'att1@contoso.test').Result | Should -Be 'Already gone'
        (Get-Meeting $r 'M1').Status | Should -Be 'Removed'
        (Get-Meeting $r 'S1').Status | Should -Be 'Skipped'
        @($script:Fake.Mailboxes['room2@contoso.test'].Events | Where-Object iCalUId -eq $script:Ids.S1).Count | Should -Be 1
    }

    It 'describes the plan before acting' {
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms'
        $plan = Get-MclCleanupPlan -Result $r -Action Remove
        $plan.Remove.Count | Should -Be 8   # M1: 3 attendees (one through the group) and a room, S1: 1 + room, O1: 1 + room
        $plan.Keep.Count | Should -Be 2
        $plan.Text | Should -Match 'stay in the calendar of their organizer'
        (Get-MclCleanupPlan -Result $r -Action Cancel).Cancel.Count | Should -Be 2
    }
}

Describe 'Report and replay' {
    BeforeEach { New-TestTenant }

    It 'writes the CSV, JSON and HTML files, with no formula injection' {
        Add-FakeMeeting -Organizer 'org@contoso.test' -Subject '=HYPERLINK("x")' -Start '2030-08-01T10:00:00' -Attendees 'att1@contoso.test' | Out-Null
        $s = New-TestSettings
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms'
        $report = Export-MclReport -Result $r -OutputPath $s.OutputPath
        foreach ($f in 'Meetings', 'Copies', 'Summary', 'Html') { Test-Path -LiteralPath $report.Files[$f] | Should -BeTrue }
        $csv = [IO.File]::ReadAllText($report.Files.Meetings)
        $csv | Should -Match ([regex]::Escape(';"''=HYPERLINK(""x"")";'))
        $html = [IO.File]::ReadAllText($report.Files.Html)
        $html | Should -Not -Match '\{\{[A-Z_]+\}\}'
        $html | Should -Match 'M1 Budget review'
        $html | Should -Not -Match '<script>alert'
        (Split-Path $report.Directory -Leaf) | Should -Match '^MeetingCleanup_Report_\d{8}-\d{6}'
    }

    It 'replays a reviewed report: the same meetings, by their IDs' {
        $s = New-TestSettings
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms'
        $report = Export-MclReport -Result $r -OutputPath $s.OutputPath
        $replay = Import-MclReport -Path $report.Directory -MeetingId $script:Ids.M1
        $replay.Meetings.Count | Should -Be 1
        $replay.Meetings[0].Subject | Should -Be 'M1 Budget review'
        $replay.Meetings[0].Copies.GetType().Name | Should -Be 'List`1'
        & $script:Module { Initialize-MclSteps -Total 2 }
        $done = Invoke-MclCleanup -Settings $s -Result $replay -Action Remove
        $done.Meetings[0].Status | Should -Be 'Removed'
        @($script:Fake.Mailboxes['att2@contoso.test'].Events | Where-Object iCalUId -eq $script:Ids.M1).Count | Should -Be 0
        @($script:Fake.Mailboxes['att2@contoso.test'].Events | Where-Object iCalUId -eq $script:Ids.O1).Count | Should -Be 1
        { Import-MclReport -Path $script:Root } | Should -Throw '*Summary.json*'
    }
}

Describe 'List of organizers' {
    BeforeEach { New-TestTenant }

    It 'reads a text file and a CSV file of organizers, X500 addresses included' {
        $dir = Join-Path $script:Root 'artifacts'; [void][IO.Directory]::CreateDirectory($dir)
        $txt = Join-Path $dir 'organizers.txt'
        "# leavers`r`norg@contoso.test`r`n/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=0a1b-Gary Gone`r`nnot-an-address`r`n" | Set-Content $txt
        $csv = Join-Path $dir 'organizers.csv'
        "DisplayName;PrimarySmtpAddress;LegacyExchangeDN`r`nOlivia;org@contoso.test;`r`nGary;;/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=0a1b-Gary Gone" | Set-Content $csv
        $a = & $script:Module { param($p) Read-MclAddressFile $p -AllowX500 } $txt
        $a.Count | Should -Be 2
        $a | Should -Contain 'org@contoso.test'
        $b = & $script:Module { param($p) Read-MclAddressFile $p -AllowX500 } $csv
        $b.Count | Should -Be 2
        ($b -join ' ') | Should -Match 'cn=0a1b-Gary Gone'
        $r = New-MclRequest -Settings (New-TestSettings) -Organizer 'gone@contoso.test' -OrganizerFile $txt
        $r.Organizer.Count | Should -Be 3
        (Test-MclRequest $r).IsValid | Should -BeTrue
        (Test-MclRequest (New-MclRequest -Settings (New-TestSettings) -OrganizerFile (Join-Path $dir 'missing.txt'))).Problems -join ' ' | Should -Match 'Organizer file not found'
    }

    It 'resolves many organizers in a few $batch calls, and searches them in one pass' {
        foreach ($i in 1..30) { Add-FakeMailbox "user$i@contoso.test" -Name "User $i" }
        $before = $script:Fake.Batches.Count
        $list = @('org@contoso.test', 'gone@contoso.test', 'olivia@contoso.test') + @(1..30 | ForEach-Object { "user$_@contoso.test" })
        $r = Find-Test $list 'Organizer', 'Rooms'
        $r.Organizers.Count | Should -Be 32   # the alias olivia@ is the same mailbox as org@
        ($script:Fake.Batches.Count - $before) | Should -BeLessThan 20
        @($r.Meetings.Subject) | Should -Be @('M1 Budget review', 'S1 Weekly', 'O1 Orphan in a room', 'D1 Departed with room')
        (Get-Meeting $r 'D1').OrganizerKey | Should -Be 'gone@contoso.test'
        (Get-Meeting $r 'M1').OrganizerKey | Should -Be 'org@contoso.test'
        $rows = @(& $script:Module { param($r) Get-MclOrganizerRows $r } $r)
        ($rows | Where-Object PrimaryAddress -eq 'org@contoso.test').Meetings | Should -Be 3
        ($rows | Where-Object PrimaryAddress -eq 'gone@contoso.test').Meetings | Should -Be 1
        ($rows | Where-Object PrimaryAddress -eq 'user1@contoso.test').Meetings | Should -Be 0
    }

    It 'finds the meeting of an organizer in the calendar of another organizer of the list' {
        # att2 organizes 'Someone else' with att1 and room1; both are in the list: found once, from either side.
        $r = Find-Test 'att1@contoso.test', 'att2@contoso.test' 'Organizer'
        @($r.Meetings | Where-Object Subject -eq 'Someone else').Count | Should -Be 1
        (Get-Meeting $r 'Someone else').OrganizerCopy | Should -Be 'Present'
    }
}

Describe 'Backup and restore' {
    BeforeEach { New-TestTenant }

    It 'writes the backup before any change, with the content of each meeting' {
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms'
        $run = Invoke-TestCleanup $r 'Remove'
        $file = Join-Path $run.Folder 'MeetingCleanup-Backup.json'
        Test-Path $file | Should -BeTrue
        $b = Get-Content $file -Raw | ConvertFrom-Json -Depth 32
        $b.Kind | Should -Be 'Backup'
        $b.Meetings.Count | Should -Be 3
        ($b.Meetings | Where-Object Subject -eq 'M1 Budget review').Event.subject | Should -Be 'M1 Budget review'
        @(($b.Meetings | Where-Object Subject -eq 'M1 Budget review').Copies).Count | Should -Be 5
        Test-Path (Join-Path $run.Folder 'MeetingCleanup-Organizers.csv') | Should -BeTrue
        @($run.Result.Meetings | ForEach-Object { $_.Copies } | Where-Object Result -eq 'Removed' | Where-Object { -not $_.ActionUtc }).Count | Should -Be 0
    }

    It 'Remove leaves the copies Declined / Free; Restore puts them back, answered again, without any message' {
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms'
        $run = Invoke-TestCleanup $r 'Remove'
        $orgCopy = $script:Fake.Mailboxes['org@contoso.test'].Events | Where-Object iCalUId -eq $script:Ids.M1
        ($orgCopy.attendees | Where-Object { $_.emailAddress.address -eq 'room1@contoso.test' }).status.response | Should -Be 'declined'
        $restored = Invoke-TestRestore $run.Folder
        $restored.Action | Should -Be 'Restore'
        $restored.Status | Should -Be 'Completed'
        $restored.Counts.Restored | Should -Be 8
        (Get-Meeting $restored 'M1').Status | Should -Be 'Restored'
        $room = $script:Fake.Mailboxes['room1@contoso.test'].Events | Where-Object iCalUId -eq $script:Ids.M1
        $room.responseStatus.response | Should -Be 'accepted'
        $room.showAs | Should -Be 'busy'
        $att = $script:Fake.Mailboxes['att1@contoso.test'].Events | Where-Object iCalUId -eq $script:Ids.M1
        $att.responseStatus.response | Should -Be 'tentativelyAccepted'
        ($orgCopy.attendees | Where-Object { $_.emailAddress.address -eq 'room1@contoso.test' }).status.response | Should -Be 'accepted'
        $script:Fake.Messages.Count | Should -Be 0
        @($restored.Meetings | ForEach-Object { $_.Copies } | Where-Object Result -eq 'Restored' | Where-Object Verified -ne 'Yes').Count | Should -Be 0
    }

    It 'restores exactly the meeting chosen in a room holding several copies with the same subject' {
        Add-FakeMeeting -Organizer 'org@contoso.test' -OrganizerName 'Olivia Organizer' -Subject 'M2 Second in room1' -Start '2030-03-12T10:00:00' -Attendees 'att2@contoso.test' -Rooms 'room1@contoso.test' | Out-Null
        Add-FakeMeeting -Organizer 'org@contoso.test' -OrganizerName 'Olivia Organizer' -Subject 'M3 Third in room1' -Start '2030-03-13T10:00:00' -Attendees 'att2@contoso.test' -Rooms 'room1@contoso.test' | Out-Null
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms'
        $run = Invoke-TestCleanup $r 'Remove'
        @($script:Fake.Mailboxes['room1@contoso.test'].Purges).Count | Should -Be 4   # M1, O1, M2, M3: all named 'Olivia Organizer'
        $m2 = (Get-Meeting $run.Result 'M2').MeetingId
        $restored = Invoke-TestRestore $run.Folder $m2
        @($script:Fake.Mailboxes['room1@contoso.test'].Events | ForEach-Object iCalUId) | Should -Contain $m2
        @($script:Fake.Mailboxes['room1@contoso.test'].Events | Where-Object { $_.iCalUId -in $script:Ids.M1, $script:Ids.O1 }).Count | Should -Be 0
        @($script:Fake.Mailboxes['room1@contoso.test'].Purges).Count | Should -Be 3
        $restored.Warnings.Count | Should -Be 0
        $restored.Counts.Restored | Should -Be 2
    }

    It 'Cancel: a meeting cancelled by its organizer is not restorable, a meeting removed silently is' {
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms'
        $run = Invoke-TestCleanup $r 'Cancel' 'Bye'
        $restored = Invoke-TestRestore $run.Folder
        (Get-Meeting $restored 'M1').Status | Should -Be 'Not restorable'
        (Get-Meeting $restored 'O1').Status | Should -Be 'Restored'
        @($script:Fake.Mailboxes['att2@contoso.test'].Events | Where-Object iCalUId -eq $script:Ids.O1).Count | Should -Be 1
        $restored.Status | Should -Be 'Warning'
    }

    It 'a copy already back is left as is; a copy no longer in Recoverable Items is reported Not found' {
        $r = Find-Test 'org@contoso.test' 'Organizer' @{ Subject = 'M1' }
        $run = Invoke-TestCleanup $r 'Remove'
        $first = Invoke-TestRestore $run.Folder
        $first.Counts.Restored | Should -Be 4
        $script:Fake.Mailboxes['room1@contoso.test'].Events.RemoveAll({ param($e) $e.iCalUId -eq $script:Ids.M1 }) | Out-Null
        $again = Invoke-TestRestore $run.Folder
        $again.Counts.AlreadyPresent | Should -Be 3
        (Get-CopyOf (Get-Meeting $again 'M1') 'room1@contoso.test').Result | Should -Be 'Not found'
        $again.Status | Should -Be 'Warning'
    }

    It 'restores a report of version 1.0.0 (no time per copy) with the time of the run' {
        $r = Find-Test 'gone@contoso.test' 'Rooms', 'AllMailboxes'
        $run = Invoke-TestCleanup $r 'Remove'
        $summary = Join-Path $run.Folder 'MeetingCleanup-Summary.json'
        $json = Get-Content $summary -Raw | ConvertFrom-Json -Depth 32
        foreach ($m in $json.Meetings) { foreach ($c in $m.Copies) { $c.PSObject.Properties.Remove('ActionUtc'); $c.PSObject.Properties.Remove('ShowAs') } }
        [IO.File]::WriteAllText($summary, ($json | ConvertTo-Json -Depth 32))
        $restored = Invoke-TestRestore $run.Folder
        $restored.Counts.Restored | Should -Be 3
        @($script:Fake.Mailboxes['att2@contoso.test'].Events | Where-Object iCalUId -eq $script:Ids.D2).Count | Should -Be 1
    }

    It 'never removes anything: a copy that cannot be verified stays, a meeting with the same subject is not touched' {
        # att1 organizes its own meeting with the same subject: a restore must never touch it.
        Add-FakeMeeting -Organizer 'att1@contoso.test' -OrganizerName 'Attendee 1' -Subject 'M1 Budget review' -Start '2030-03-20T10:00:00' -Attendees 'att3@contoso.test' | Out-Null
        $r = Find-Test 'org@contoso.test' 'Organizer' @{ Subject = 'M1' }
        $run = Invoke-TestCleanup $r 'Remove'
        $script:Fake.Fail["/users/att2@contoso.test/events?*select=id,responseStatus,showAs*"] = @{ Status = 403; Code = 'ErrorAccessDenied' }
        $from = $script:Fake.Calls.Count
        $restored = Invoke-TestRestore $run.Folder
        $script:Fake.Fail.Clear()
        @($script:Fake.Calls | Select-Object -Skip $from | Where-Object Url -like '*permanentDelete*').Count | Should -Be 0
        $att2 = Get-CopyOf (Get-Meeting $restored 'M1') 'att2@contoso.test'
        $att2.Result | Should -Be 'Restored'
        $att2.Verified | Should -BeLike 'Unknown*'
        @($script:Fake.Mailboxes['att2@contoso.test'].Events | Where-Object iCalUId -eq $script:Ids.M1).Count | Should -Be 1
        @($script:Fake.Mailboxes['att1@contoso.test'].Events | Where-Object { $_.subject -eq 'M1 Budget review' -and $_.isOrganizer }).Count | Should -Be 1
        $script:Fake.Messages.Count | Should -Be 0
    }

    It 'does not guess: copies with the same subject removed in the same second, one asked, nothing restored there' {
        $m2id = Add-FakeMeeting -Organizer 'org@contoso.test' -OrganizerName 'Olivia Organizer' -Subject 'M2 Second in room1' -Start '2030-03-12T10:00:00' -Attendees 'att2@contoso.test' -Rooms 'room1@contoso.test'
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms' @{ MeetingId = @($script:Ids.M1, $m2id) }
        $run = Invoke-TestCleanup $r 'Remove'
        # The removal times of Exchange are to the second: two items of one second cannot be told apart.
        $purges = @($script:Fake.Mailboxes['room1@contoso.test'].Purges)
        $purges.Count | Should -Be 2
        $purges[1].LastModifiedUtc = $purges[0].LastModifiedUtc
        $m2 = (Get-Meeting $run.Result 'M2').MeetingId
        $restored = Invoke-TestRestore $run.Folder $m2
        $room = Get-CopyOf (Get-Meeting $restored 'M2') 'room1@contoso.test'
        $room.Result | Should -Be 'Failed'
        $room.Detail | Should -BeLike 'ambiguous*Get-RecoverableItems -Identity room1@contoso.test*'
        @($script:Fake.Mailboxes['room1@contoso.test'].Purges).Count | Should -Be 2
        (Get-CopyOf (Get-Meeting $restored 'M2') 'att2@contoso.test').Result | Should -Be 'Restored'
        # The whole group asked: every item comes back, whatever its rank.
        $all = Invoke-TestRestore $run.Folder
        @($script:Fake.Mailboxes['room1@contoso.test'].Purges).Count | Should -Be 0
        @($all.Meetings | ForEach-Object { $_.Copies } | Where-Object { $_.Mailbox -eq 'room1@contoso.test' -and $_.Result -eq 'Restored' -and $_.Verified -eq 'Yes' }).Count | Should -Be 2
    }

    It 'a copy put back but not verified is a warning; the same restore again answers it, without a message' {
        $r = Find-Test 'org@contoso.test' 'Organizer' @{ Subject = 'M1' }
        $run = Invoke-TestCleanup $r 'Remove'
        $script:Fake.Fail["/users/room1@contoso.test/events?*select=id,responseStatus,showAs*"] = @{ Status = 403; Code = 'ErrorAccessDenied' }
        $first = Invoke-TestRestore $run.Folder
        $script:Fake.Fail.Clear()
        $first.Status | Should -Be 'Warning'
        (Get-CopyOf (Get-Meeting $first 'M1') 'room1@contoso.test').Verified | Should -BeLike 'Unknown*'
        $room = $script:Fake.Mailboxes['room1@contoso.test'].Events | Where-Object iCalUId -eq $script:Ids.M1
        $room.showAs | Should -Be 'free'
        $again = Invoke-TestRestore $run.Folder
        $again.Status | Should -Be 'Completed'
        $copy = Get-CopyOf (Get-Meeting $again 'M1') 'room1@contoso.test'
        $copy.Result | Should -Be 'Already present'
        $copy.Detail | Should -BeLike '*answered accepted again*'
        $room.showAs | Should -Be 'busy'
        $room.responseStatus.response | Should -Be 'accepted'
        (Get-CopyOf (Get-Meeting $again 'M1') 'att1@contoso.test').Detail | Should -Be 'already in the calendar: nothing to restore'
        $script:Fake.Messages.Count | Should -Be 0
    }
    It 'replays a report of version 1.0.0 with Remove' {
        $s = New-TestSettings
        $r = Find-Test 'org@contoso.test' 'Organizer' @{ Subject = 'M1' }
        $report = Export-MclReport -Result $r -OutputPath $s.OutputPath
        $summary = $report.Files.Summary
        $json = Get-Content $summary -Raw | ConvertFrom-Json -Depth 32
        foreach ($m in $json.Meetings) { $m.PSObject.Properties.Remove('OrganizerKey'); foreach ($c in $m.Copies) { $c.PSObject.Properties.Remove('ActionUtc'); $c.PSObject.Properties.Remove('ShowAs') } }
        [IO.File]::WriteAllText($summary, ($json | ConvertTo-Json -Depth 32))
        $replay = Import-MclReport -Path $report.Directory
        $run = Invoke-TestCleanup $replay 'Remove'
        $run.Result.Counts.Removed | Should -Be 4
        @($run.Result.Meetings | ForEach-Object { $_.Copies } | Where-Object Result -eq 'Removed' | Where-Object { -not $_.ActionUtc }).Count | Should -Be 0
    }
    It 'refuses the report of a search' {
        $s = New-TestSettings
        $r = Find-Test 'org@contoso.test' 'Organizer'
        $report = Export-MclReport -Result $r -OutputPath $s.OutputPath
        { Import-MclRestoreSource -Path $report.Directory } | Should -Throw '*Remove, Cancel or Transfer run*'
    }
}
Describe 'Rooms' {
    BeforeEach { New-TestTenant }

    It 'finds every meeting of a room, whatever its organizer, and every copy of it' {
        $r = Find-Rooms 'room1@contoso.test' '2030-01-01' '2030-12-31'
        $r.Request.Mode | Should -Be 'Rooms'
        @($r.Meetings | ForEach-Object Subject) | Should -Be @('M1 Budget review', 'Someone else', 'O1 Orphan in a room')
        @($r.Organizers | ForEach-Object PrimaryAddress | Sort-Object) | Should -Be @('att2@contoso.test', 'org@contoso.test')
        (Get-CopyOf (Get-Meeting $r 'M1') 'att1@contoso.test').EventId | Should -Not -BeNullOrEmpty
        (Get-Meeting $r 'O1').OrganizerCopy | Should -Be 'Absent'
    }

    It 'cancels only the occurrences of a series that fall in the period; the series goes on' {
        $r = Find-Rooms 'room2@contoso.test' '2030-03-15' '2030-03-26'
        $s1 = Get-Meeting $r 'S1'
        $s1.Scope | Should -Be 'Occurrences'
        $s1.Occurrences | Should -Be 2
        @($s1.Copies | Where-Object Occurrence).Count | Should -Be 6
        $run = Invoke-TestCleanup $r 'Cancel' 'Room closed for works'
        $cancels = @($script:Fake.Messages | Where-Object Kind -eq 'Cancellation')
        @($cancels | ForEach-Object Occurrence | Sort-Object) | Should -Be @('2030-03-18T09:00', '2030-03-25T09:00')
        @($cancels | Where-Object Comment -ne 'Room closed for works').Count | Should -Be 0
        foreach ($mb in 'room2@contoso.test', 'att1@contoso.test', 'org@contoso.test') {
            $copy = Get-FakeEvents $mb $script:Ids.S1
            $copy.Count | Should -Be 1
            @($copy[0].DeletedOccurrences | Sort-Object) | Should -Be @('2030-03-18T09:00', '2030-03-25T09:00')
        }
        @($script:Fake.Mailboxes['room2@contoso.test'].Purges).Count | Should -Be 0
        $run.Result.Status | Should -Be 'Completed'
        (Get-Meeting (Invoke-TestRestore $run.Folder) 'S1').Status | Should -Be 'Not restorable'
    }

    It 'removes occurrences silently (not restorable) and single meetings (restorable)' {
        $r = Find-Rooms 'room2@contoso.test' '2030-03-15' '2030-05-31'
        (Get-Meeting $r 'S1').Occurrences | Should -Be 3
        $run = Invoke-TestCleanup $r 'Remove'
        $script:Fake.Messages.Count | Should -Be 0
        @((Get-FakeEvents 'att1@contoso.test' $script:Ids.S1)[0].DeletedOccurrences).Count | Should -Be 3
        @((Get-FakeEvents 'org@contoso.test' $script:Ids.S1)[0].DeletedOccurrences).Count | Should -Be 0
        (Get-FakeEvents 'room2@contoso.test' $script:Ids.D1).Count | Should -Be 0
        $restored = Invoke-TestRestore $run.Folder
        $restored.Counts.Restored | Should -Be 2
        $restored.Counts.NotRestorable | Should -Be 6
        (Get-FakeEvents 'room2@contoso.test' $script:Ids.D1).Count | Should -Be 1
    }

    It 'handles a series whole when every occurrence is in the period' {
        $r = Find-Rooms 'room2@contoso.test' '2030-03-01' '2030-04-30'
        (Get-Meeting $r 'S1').Scope | Should -Be 'Whole'
        @((Get-Meeting $r 'S1').Copies | Where-Object Occurrence).Count | Should -Be 0
    }

    It 'whole only when the period holds every occurrence: a period starting after the first one splits it' {
        $uid = Add-FakeMeeting -Organizer 'org@contoso.test' -OrganizerName 'Olivia Organizer' -Subject 'E1 Daily' -Start '2030-07-01T09:00:00' -Attendees 'att1@contoso.test' -Rooms 'room1@contoso.test' -Recurrence @{ pattern = @{ type = 'daily'; interval = 1 }; range = @{ type = 'endDate'; startDate = '2030-07-01'; endDate = '2030-07-04' } }
        $r = Find-Rooms 'room1@contoso.test' ([datetime]'2030-07-01T14:00:00') '2030-07-10'
        $e1 = Get-Meeting $r 'E1'
        $e1.Scope | Should -Be 'Occurrences'
        $e1.Occurrences | Should -Be 3
        (Get-Meeting (Find-Rooms 'room1@contoso.test' '2030-06-30' '2030-07-10') 'E1').Scope | Should -Be 'Whole'
        $null = Invoke-TestCleanup $r 'Cancel'
        @((Get-FakeEvents 'att1@contoso.test' $uid)[0].CancelledOccurrences | Sort-Object) | Should -Be @('2030-07-02T09:00', '2030-07-03T09:00', '2030-07-04T09:00')
    }

    It 'an occurrence moved to another room is not acted on' {
        # The occurrence of 18/03 moved from room 2 to room 1: room 2 no longer holds it.
        (Get-FakeEvents 'room2@contoso.test' $script:Ids.S1)[0].DeletedOccurrences.Add('2030-03-18T09:00')
        $r = Find-Rooms 'room2@contoso.test' '2030-03-15' '2030-03-26'
        $s1 = Get-Meeting $r 'S1'
        $s1.Occurrences | Should -Be 1
        @($s1.Copies | Where-Object Occurrence | ForEach-Object Mailbox | Sort-Object) | Should -Be @('att1@contoso.test', 'org@contoso.test', 'room2@contoso.test')
        $null = Invoke-TestCleanup $r 'Cancel'
        @($script:Fake.Messages | Where-Object Kind -eq 'Cancellation' | ForEach-Object Occurrence) | Should -Be @('2030-03-25T09:00')
        @((Get-FakeEvents 'att1@contoso.test' $script:Ids.S1)[0].CancelledOccurrences) | Should -Be @('2030-03-25T09:00')
    }

    It 'the occurrences of the organizer cannot be read: the series is left as it is, attendees included' {
        $script:Fake.Fail['/users/org@contoso.test/events/*/instances*'] = @{ Status = 403; Code = 'ErrorAccessDenied' }
        $r = Find-Rooms 'room2@contoso.test' '2030-03-15' '2030-03-26'
        $script:Fake.Fail.Clear()
        $s1 = Get-Meeting $r 'S1'
        @($s1.Copies | Where-Object EventId).Count | Should -Be 0
        ($s1.Notes -join ' ') | Should -Match 'Not acted on'
        $null = Invoke-TestCleanup $r 'Cancel'
        $script:Fake.Messages.Count | Should -Be 0
        foreach ($mb in 'att1@contoso.test', 'room2@contoso.test') { @((Get-FakeEvents $mb $script:Ids.S1)[0].DeletedOccurrences).Count | Should -Be 0 }
    }

    It 'Cancel: an organizer copy that cannot be read (denied) is not a deleted mailbox; the meeting is left as it is' {
        $script:Fake.Fail['/users/org@contoso.test/events?*'] = @{ Status = 403; Code = 'ErrorAccessDenied' }
        $r = Find-Rooms 'room1@contoso.test' '2030-03-10' '2030-03-11'
        $script:Fake.Fail.Clear()
        (Get-Meeting $r 'M1').OrganizerCopy | Should -Be 'Not read'
        (Get-Meeting $r 'Someone').OrganizerCopy | Should -Be 'Present'
        $run = Invoke-TestCleanup $r 'Cancel' 'Room closed'
        (Get-Meeting $run.Result 'M1').Status | Should -Be 'Skipped'
        (Get-Meeting $run.Result 'Someone').Status | Should -Be 'Cancelled'
        $run.Result.Status | Should -Be 'Warning'
        foreach ($mb in 'att1@contoso.test', 'att2@contoso.test', 'room1@contoso.test') { (Get-FakeEvents $mb $script:Ids.M1).Count | Should -Be 1 }
        @($script:Fake.Messages | Where-Object ICalUId -eq $script:Ids.M1).Count | Should -Be 0
    }

    It 'checks the request: a period for an action, rooms or organizers, no transfer of rooms' {
        $s = New-TestSettings
        (Test-MclRequest (New-MclRequest -Settings $s -Room 'room1@contoso.test' -Action Cancel)).Problems -join ' ' | Should -Match 'give the period'
        (Test-MclRequest (New-MclRequest -Settings $s -Room 'room1@contoso.test' -Start '2030-01-01' -End '2030-01-31' -Action Cancel)).IsValid | Should -BeTrue
        (Test-MclRequest (New-MclRequest -Settings $s -Room 'room1@contoso.test' -Organizer 'org@contoso.test')).Problems -join ' ' | Should -Match 'not both'
        (Test-MclRequest (New-MclRequest -Settings $s -Room 'room1@contoso.test' -Start '2030-01-01' -End '2030-01-31' -Action Transfer -NewOrganizer 'att3@contoso.test')).Problems -join ' ' | Should -Match 'not rooms'
    }
}

Describe 'Transfer' {
    BeforeEach { New-TestTenant }

    It 'deleted organizer: re-created by the new organizer, one invitation, the old copies removed silently' {
        $r = Find-Test 'gone@contoso.test' 'Rooms', 'AllMailboxes'
        $t = Invoke-TestTransfer $r 'att3@contoso.test'
        $t.Plan.Recreate.Count | Should -Be 2
        $t.Plan.Native.Count | Should -Be 0
        $t.Result.Status | Should -Be 'Completed'
        $d1 = Get-Meeting $t.Result 'D1'
        $d1.Status | Should -Be 'Transferred'
        $d1.TransferMethod | Should -Be 'Recreate'
        $d1.NewOrganizer | Should -Be 'att3@contoso.test'
        $new = @($script:Fake.Mailboxes['att3@contoso.test'].Events | Where-Object { $_.isOrganizer -and $_.subject -eq 'D1 Departed with room' })
        $new.Count | Should -Be 1
        @($new[0].attendees | ForEach-Object { $_.emailAddress.address } | Sort-Object) | Should -Be @('att1@contoso.test', 'room2@contoso.test')
        $d1.NewMeetingId | Should -Be $new[0].iCalUId
        (Get-FakeEvents 'room2@contoso.test' $new[0].iCalUId)[0].showAs | Should -Be 'busy'
        (Get-FakeEvents 'att1@contoso.test' $new[0].iCalUId).Count | Should -Be 1
        foreach ($mb in 'att1@contoso.test', 'room2@contoso.test') { (Get-FakeEvents $mb $script:Ids.D1).Count | Should -Be 0 }
        (Get-FakeEvents 'att2@contoso.test' $script:Ids.D2).Count | Should -Be 0
        @($script:Fake.Messages | Where-Object Kind -eq 'Invitation').Count | Should -Be 2
        @($script:Fake.Messages | Where-Object Kind -eq 'Cancellation').Count | Should -Be 0
        Test-Path (Join-Path $t.Folder 'MeetingCleanup-Backup.json') | Should -BeTrue
        (Import-Csv (Join-Path $t.Folder 'MeetingCleanup-Meetings.csv') -Delimiter ';' | Where-Object Subject -like 'D1*').NewOrganizer | Should -Be 'att3@contoso.test'
    }

    It 're-creates a series from now, with its removed and moved occurrences, in one invitation' {
        foreach ($mb in 'org@contoso.test', 'att1@contoso.test', 'room2@contoso.test') {
            $copy = (Get-FakeEvents $mb $script:Ids.S1)[0]
            $copy.DeletedOccurrences.Add('2030-03-25T09:00')
            $copy.Exceptions['2030-04-01T09:00'] = @{ Start = [datetime]'2030-04-01T11:00:00'; End = [datetime]'2030-04-01T11:30:00'; Subject = ''; Location = '' }
        }
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms' @{ Subject = 'S1' }
        $t = Invoke-TestTransfer $r 'att3@contoso.test' 'Recreate' ([datetime]'2030-03-12')
        (Get-Meeting $t.Result 'S1').Status | Should -Be 'Transferred'
        $new = @($script:Fake.Mailboxes['att3@contoso.test'].Events | Where-Object { $_.isOrganizer -and $_.subject -eq 'S1 Weekly' })[0]
        $new.start.dateTime | Should -BeLike '2030-03-11T09:00*'
        $new.recurrence.range.type | Should -Be 'endDate'
        [string]$new.recurrence.range.endDate | Should -Be '2030-04-01'
        $att = (Get-FakeEvents 'att1@contoso.test' $new.iCalUId)[0]
        $starts = @(Get-FakeOccurrences $att ([datetime]'2030-03-01') ([datetime]'2030-05-01') | ForEach-Object { New-FakeInstance $att $_ } | Where-Object { $_ } | ForEach-Object { $_.start.dateTime.Substring(0, 16) })
        $starts | Should -Be @('2030-03-11T09:00', '2030-03-18T09:00', '2030-04-01T11:00')
        ($t.Plan.Lines -join ' ') | Should -Match 'move from now'
        @($script:Fake.Messages | Where-Object Kind -eq 'Invitation').Count | Should -Be 1
        $cancel = @($script:Fake.Messages | Where-Object Kind -eq 'Cancellation')
        $cancel.Count | Should -Be 1
        $cancel[0].Comment | Should -BeLike '*att3@contoso.test*'
        foreach ($mb in 'org@contoso.test', 'att1@contoso.test', 'room2@contoso.test') { (Get-FakeEvents $mb $script:Ids.S1).Count | Should -Be 0 }
    }

    It 'user deleted but mailbox still reachable (soft-deleted): re-created, his copy left as it is, no cancellation' {
        Add-FakeMailbox 'left@contoso.test' -Name 'Lena Left' -NoDirectory
        $uid = Add-FakeMeeting -Organizer 'left@contoso.test' -OrganizerName 'Lena Left' -Subject 'L1 Left with room' -Start '2030-06-03T10:00:00' -Attendees 'att2@contoso.test' -Rooms 'room1@contoso.test'
        $r = Find-Test 'left@contoso.test' 'Organizer', 'Rooms'
        (Get-Meeting $r 'L1').OrganizerCopy | Should -Be 'Present'
        $t = Invoke-TestTransfer $r 'att3@contoso.test'
        $t.Plan.Native.Count | Should -Be 0
        $l1 = Get-Meeting $t.Result 'L1'
        $l1.Status | Should -Be 'Transferred'
        (Get-CopyOf $l1 'left@contoso.test').Result | Should -Be 'Kept'
        @($script:Fake.Messages | Where-Object Kind -eq 'Cancellation').Count | Should -Be 0
        foreach ($mb in 'att2@contoso.test', 'room1@contoso.test') { (Get-FakeEvents $mb $uid).Count | Should -Be 0 }
        (Get-FakeEvents 'room1@contoso.test' $l1.NewMeetingId).Count | Should -Be 1
    }
    It 'organizer mailbox gone: the occurrences come from the attendees and the rooms, a cancelled one does not come back' {
        $uid = Add-FakeMeeting -Organizer 'gone@contoso.test' -OrganizerName 'Gary Gone' -Subject 'G1 Weekly of a leaver' -Start '2030-06-03T09:00:00' -Attendees 'att1@contoso.test', 'att2@contoso.test' -Rooms 'room1@contoso.test' -Recurrence @{ pattern = @{ type = 'weekly'; interval = 1; daysOfWeek = @('monday') }; range = @{ type = 'numbered'; numberOfOccurrences = 4; startDate = '2030-06-03' } }
        (Get-FakeEvents 'att1@contoso.test' $uid)[0].CancelledOccurrences.Add('2030-06-10T09:00')
        (Get-FakeEvents 'att2@contoso.test' $uid)[0].CancelledOccurrences.Add('2030-06-10T09:00')
        (Get-FakeEvents 'room1@contoso.test' $uid)[0].DeletedOccurrences.Add('2030-06-10T09:00')
        (Get-FakeEvents 'att1@contoso.test' $uid)[0].DeletedOccurrences.Add('2030-06-17T09:00')
        $r = Find-Test 'gone@contoso.test' 'Rooms' @{ Subject = 'G1' }
        $t = Invoke-TestTransfer $r 'att3@contoso.test'
        $g1 = Get-Meeting $t.Result 'G1'
        $g1.Status | Should -Be 'Transferred'
        $att = (Get-FakeEvents 'att2@contoso.test' $g1.NewMeetingId)[0]
        $starts = @(Get-FakeOccurrences $att ([datetime]'2030-06-01') ([datetime]'2030-07-01') | ForEach-Object { New-FakeInstance $att $_ } | Where-Object { $_ } | ForEach-Object { $_.start.dateTime.Substring(0, 10) })
        $starts | Should -Be @('2030-06-03', '2030-06-17', '2030-06-24')
    }
    It 'organizer present, Auto: moved by Exchange Online, nothing re-created' {
        $r = Find-Test 'org@contoso.test' 'Organizer' @{ Subject = 'M1' }
        $t = Invoke-TestTransfer $r 'att3@contoso.test'
        $m1 = Get-Meeting $t.Result 'M1'
        $m1.Status | Should -Be 'Transferred'
        $m1.TransferMethod | Should -Be 'Native'
        $m1.NewMeetingId | Should -Not -BeNullOrEmpty
        $script:Fake.Messages.Count | Should -Be 0
        $script:Fake.Created.Count | Should -Be 0
        (Get-FakeEvents 'att1@contoso.test' $script:Ids.M1)[0].organizer.emailAddress.address | Should -Be 'att3@contoso.test'
    }

    It 'Exchange Online refuses: the meeting is Failed, nothing changed, re-creation suggested' {
        $script:Fake.NativeError = 'A server side error has occurred because of which the operation could not be completed.'
        $r = Find-Test 'org@contoso.test' 'Organizer' @{ Subject = 'M1' }
        $t = Invoke-TestTransfer $r 'att3@contoso.test'
        $m1 = Get-Meeting $t.Result 'M1'
        $m1.Status | Should -Be 'Failed'
        ($m1.Notes -join ' ') | Should -Match 'TransferMethod Recreate'
        $t.Result.Status | Should -Be 'Failed'
        (Get-FakeEvents 'org@contoso.test' $script:Ids.M1).Count | Should -Be 1
    }

    It 'the invitation fails: the new meeting goes, nothing is sent, the old room copy can be restored' {
        $script:Fake.Fail['/users/att3@contoso.test/events/*'] = @{ Status = 403; Code = 'ErrorAccessDenied'; Method = 'PATCH' }
        $r = Find-Test 'gone@contoso.test' 'Rooms', 'AllMailboxes' @{ Subject = 'D1' }
        $t = Invoke-TestTransfer $r 'att3@contoso.test'
        $script:Fake.Fail.Clear()
        $d1 = Get-Meeting $t.Result 'D1'
        $d1.Status | Should -Be 'Failed'
        @($script:Fake.Mailboxes['att3@contoso.test'].Events | Where-Object subject -eq 'D1 Departed with room').Count | Should -Be 0
        (Get-FakeEvents 'att1@contoso.test' $script:Ids.D1).Count | Should -Be 1
        (Get-FakeEvents 'room2@contoso.test' $script:Ids.D1).Count | Should -Be 0
        @($script:Fake.Messages | Where-Object Kind -eq 'Invitation').Count | Should -Be 0
        $restored = Invoke-TestRestore $t.Folder
        $restored.Counts.Restored | Should -Be 1
        (Get-FakeEvents 'room2@contoso.test' $script:Ids.D1).Count | Should -Be 1
    }

    It 'plan: a meeting already organized by the new organizer, Native impossible for a deleted mailbox' {
        $r = Find-Test 'org@contoso.test' 'Organizer' @{ Subject = 'M1' }
        $new = Resolve-MclNewOrganizer -Address 'olivia@contoso.test'
        (Get-MclTransferPlan -Result $r -NewOrganizer $new).Skipped[0].Reason | Should -Match 'already organized'
        $g = Find-Test 'gone@contoso.test' 'Rooms', 'AllMailboxes'
        $plan = Get-MclTransferPlan -Result $g -NewOrganizer (Resolve-MclNewOrganizer -Address 'att3@contoso.test') -Method Native
        $plan.Native.Count + $plan.Recreate.Count | Should -Be 0
        $plan.Skipped[0].Reason | Should -Match 'mailbox gone'
        { Resolve-MclNewOrganizer -Address 'nobody@contoso.test' } | Should -Throw '*not a mailbox*'
    }

    It 'refuses the meetings of a rooms search' {
        $r = Find-Rooms 'room1@contoso.test' '2030-01-01' '2030-12-31'
        { Get-MclTransferPlan -Result $r -NewOrganizer (Resolve-MclNewOrganizer -Address 'att3@contoso.test') } | Should -Throw '*organizers*'
    }

    It 'account not looked up (no User.Read.All): the organizer counts as active, his copy is cancelled' {
        $s = New-TestSettings
        $null = Connect-Test $s @('Calendars.ReadWrite', 'Place.Read.All', 'GroupMember.Read.All')
        & $script:Module { Initialize-MclSteps -Total 6 }
        $r = Find-MclMeetings -Settings $s -Request (New-TestRequest $s 'org@contoso.test' 'Organizer' @{ Subject = 'M1' })
        $r.Organizers[0].Account | Should -Be 'Unknown'
        $plan = Get-MclTransferPlan -Result $r -NewOrganizer (Resolve-MclNewOrganizer -Address 'att3@contoso.test')
        $plan.Native.Count | Should -Be 1
        $t = Invoke-TestTransfer $r 'att3@contoso.test' 'Recreate'
        (Get-CopyOf (Get-Meeting $t.Result 'M1') 'org@contoso.test').Result | Should -Be 'Cancelled'
    }

    It 'a Transfer report replayed: the meetings transferred are skipped, the new meeting is never removed' {
        $r = Find-Test 'gone@contoso.test' 'Rooms', 'AllMailboxes' @{ Subject = 'D1' }
        $t = Invoke-TestTransfer $r 'att3@contoso.test'
        $again = Import-MclReport -Path $t.Folder
        $plan = Get-MclTransferPlan -Result $again -NewOrganizer (Resolve-MclNewOrganizer -Address 'att2@contoso.test')
        $plan.Recreate.Count | Should -Be 0
        $plan.Skipped[0].Reason | Should -Match 'already transferred'
        @((Get-MclCleanupPlan -Result $again -Action Remove).Remove | Where-Object Role -eq 'New organizer').Count | Should -Be 0
    }

    It 'a Transfer report replayed with Remove: a meeting moved by Exchange Online is left as it is' {
        $r = Find-Test 'org@contoso.test' 'Organizer' @{ Subject = 'M1' }
        $t = Invoke-TestTransfer $r 'att3@contoso.test'
        (Get-Meeting $t.Result 'M1').TransferMethod | Should -Be 'Native'
        $again = Import-MclReport -Path $t.Folder
        $plan = Get-MclCleanupPlan -Result $again -Action Remove
        $plan.Remove.Count | Should -Be 0
        $plan.Held.Count | Should -Be 1
        $run = Invoke-TestCleanup $again 'Remove'
        (Get-Meeting $run.Result 'M1').Status | Should -Be 'Skipped'
        foreach ($mb in 'att1@contoso.test', 'att2@contoso.test', 'room1@contoso.test') { (Get-FakeEvents $mb $script:Ids.M1).Count | Should -Be 1 }
    }

    It 'an occurrence moved from a past slot to a later date is carried over' {
        $today = [datetime]::UtcNow.Date
        $first = $today.AddDays(-8).AddHours(9)
        $uid = Add-FakeMeeting -Organizer 'org@contoso.test' -OrganizerName 'Olivia Organizer' -Subject 'P1 Weekly moved' -Start $first.ToString('yyyy-MM-ddTHH:mm:ss') -Attendees 'att1@contoso.test' -Rooms 'room1@contoso.test' -Recurrence @{ pattern = @{ type = 'weekly'; interval = 1; daysOfWeek = @($first.DayOfWeek.ToString().ToLowerInvariant()) }; range = @{ type = 'numbered'; numberOfOccurrences = 4; startDate = $first.ToString('yyyy-MM-dd') } }
        $moved = $today.AddDays(2).AddHours(11)
        foreach ($mb in 'org@contoso.test', 'att1@contoso.test', 'room1@contoso.test') { (Get-FakeEvents $mb $uid)[0].Exceptions[$first.ToString('yyyy-MM-ddTHH:mm')] = @{ Start = $moved; End = $moved.AddMinutes(30); Subject = ''; Location = '' } }
        $s = New-TestSettings
        $null = Connect-Test $s
        & $script:Module { Initialize-MclSteps -Total 6 }
        $r = Find-MclMeetings -Settings $s -Request (New-TestRequest $s 'org@contoso.test' 'Organizer' @{ Subject = 'P1'; Start = $today.AddDays(-30); End = $today.AddDays(60) })
        $t = Invoke-TestTransfer $r 'att3@contoso.test' 'Recreate'
        $p1 = Get-Meeting $t.Result 'P1'
        $p1.Status | Should -Be 'Transferred'
        $att = (Get-FakeEvents 'att1@contoso.test' $p1.NewMeetingId)[0]
        $starts = @(Get-FakeOccurrences $att ([datetime]::MinValue) $today.AddDays(60) | ForEach-Object { New-FakeInstance $att $_ } | Where-Object { $_ } | ForEach-Object { $_.start.dateTime.Substring(0, 16) })
        $starts | Should -Be @($moved.ToString('yyyy-MM-ddTHH:mm'), $first.AddDays(14).ToString('yyyy-MM-ddTHH:mm'), $first.AddDays(21).ToString('yyyy-MM-ddTHH:mm'))
    }

    It 'the new series does not match the old one: not transferred, the new meeting removed, nothing sent' {
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms' @{ Subject = 'S1' }
        # A time zone that does not match the times the copies give: every slot of the new series is shifted.
        (Get-Meeting $r 'S1').TimeZone = 'Dateline Standard Time'
        $t = Invoke-TestTransfer $r 'att3@contoso.test' 'Recreate'
        $s1 = Get-Meeting $t.Result 'S1'
        $s1.Status | Should -Be 'Failed'
        ($s1.Notes -join ' ') | Should -Match 'without a slot in the new one'
        @($script:Fake.Mailboxes['att3@contoso.test'].Events | Where-Object subject -eq 'S1 Weekly').Count | Should -Be 0
        $script:Fake.Messages.Count | Should -Be 0
        foreach ($mb in 'org@contoso.test', 'att1@contoso.test', 'room2@contoso.test') { (Get-FakeEvents $mb $script:Ids.S1).Count | Should -Be 1 }
    }

    It 'Stop waits for the end of a re-creation (hold), then stops' {
        & $script:Module {
            $script:Ui = @{ Cancel = $true; Hold = $true }
            try { { Assert-MclNotCancelled } | Should -Not -Throw; $script:Ui.Hold = $false; { Assert-MclNotCancelled } | Should -Throw '*Stopped*' }
            finally { $script:Ui = $null }
        }
    }
}

Describe 'Window' {    It 'builds the window from the configuration' {
        $s = New-TestSettings @{ SearchIn = @('Rooms', 'AllMailboxes'); CancelComment = 'Bye' }
        $f = New-MclForm -Configuration $s -Theme Light
        $c = $f.Controls
        $c.ScopeOrganizer.IsChecked | Should -BeFalse
        $c.ScopeRooms.IsChecked | Should -BeTrue
        $c.ScopeAll.IsChecked | Should -BeTrue
        $c.Comment.Text | Should -Be 'Bye'
        $c.ActionRemove.IsChecked | Should -BeTrue
        $c.Apply.IsEnabled | Should -BeFalse
        $c.TenantId.Text | Should -Be $script:Tenant
        $f.Form.Close()
    }

    It 'counts the meetings ticked on the action button' {
        New-TestTenant
        $s = New-TestSettings
        $r = Find-Test 'org@contoso.test' 'Organizer', 'Rooms'
        $f = New-MclForm -Configuration $s -Theme Dark
        & $script:Module { param($r) $script:Gui.Result = $r; Update-MclGuiRows } $r
        $f.Rows.Count | Should -Be 3
        $f.Controls.ApplyText.Text | Should -Be 'Remove 3 meetings'
        $f.Controls.Apply.IsEnabled | Should -BeTrue
        $f.Rows[0].Selected = $false
        $f.Controls.ActionCancel.IsChecked = $true
        $f.Controls.ApplyText.Text | Should -Be 'Cancel 2 meetings'
        $f.Controls.Counts.Text | Should -Match '3 found'
        $f.Controls.ActionTransfer.IsChecked = $true
        $f.Controls.ApplyText.Text | Should -Be 'Transfer 2 meetings'
        $f.Controls.NewOrganizer.IsEnabled | Should -BeTrue
        $f.Controls.ModeRooms.IsChecked = $true
        $f.Controls.ActionTransfer.IsEnabled | Should -BeFalse
        $f.Controls.ActionRemove.IsChecked | Should -BeTrue
        $f.Controls.SearchCard.IsEnabled | Should -BeFalse
        $f.Controls.WhoTitle.Text | Should -Be 'Rooms'
        $f.Form.Close()
    }
}

Describe 'Command line' {
    It 'stops with exit code 1 and the list of the problems on a wrong configuration' {
        $path = Join-Path $script:Root 'artifacts\cli-bad.config.psd1'
        "@{ Graph = @{ MaxConcurrency = 0 } }" | Set-Content $path
        $out = & pwsh -NoProfile -File (Join-Path $script:Root 'Invoke-MeetingCleanup.ps1') -ConfigPath $path -Organizer 'a@contoso.test' 2>&1
        $LASTEXITCODE | Should -Be 1
        ($out -join ' ') | Should -Match 'Graph.MaxConcurrency must be'
        Remove-Item $path
    }

    It 'refuses to run without an organizer' {
        $path = Join-Path $script:Root 'artifacts\cli-ok.config.psd1'
        "@{ Logging = @{ Path = '$((Join-Path $script:Root 'artifacts\test-logs'))' } }" | Set-Content $path
        $out = & pwsh -NoProfile -File (Join-Path $script:Root 'Invoke-MeetingCleanup.ps1') -ConfigPath $path 2>&1
        $LASTEXITCODE | Should -Be 1
        ($out -join ' ') | Should -Match 'Give the organizer'
        Remove-Item $path
    }
}
