<#
.SYNOPSIS
    Meeting Cleanup - organizer, mailboxes to search, meetings and their copies (dot-sourced by MeetingCleanup.psm1).

.DESCRIPTION
    How a meeting is found, whatever the case (organizer present or deleted, one meeting, a series or a period):

      1. Organizer   its addresses (SMTP, aliases, X500) from the directory, and whether its mailbox still exists.
      2. Mailboxes   where to search: the organizer's calendar, the rooms (places API, Search.Rooms, RoomFile),
                     a list of mailboxes, or every mailbox of the tenant.
      3. Search      in each of them, the meetings of the period whose organizer is the one searched:
                     $filter "type eq 'seriesMaster' or (end ge start-of-period and start lt end-of-period)",
                     then the organizer is compared on this side (Graph refuses a filter on organizer, 501).
                     A series is kept when one of its occurrences falls in the period (instances).
      4. Attendees   one copy holds the whole attendee list, the organizer included: every internal attendee,
                     room and member of an invited group is then asked for its own copy, by iCalUId (the same
                     in every copy of a meeting). So a meeting found in one room is found everywhere.

    Every copy is kept with its mailbox, its role (Organizer, Attendee, Room) and how it was found; the
    attendees that cannot be processed (external, deleted, on-premises) are listed too, with the reason.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.2.0
#>

$script:EventSelect = 'id,iCalUId,subject,type,organizer,isOrganizer,start,end,isCancelled,recurrence,responseStatus,showAs'
$script:StepIndex = 0
$script:StepTotal = 0

function Initialize-MclSteps { param([int]$Total) $script:StepIndex = 0; $script:StepTotal = $Total }

function Write-MclNextStep {
    param([Parameter(Mandatory = $true)][string]$Title, [string]$Icon = 'Info')
    $script:StepIndex++
    Write-MclStep -Number $script:StepIndex -Total ([Math]::Max($script:StepIndex, $script:StepTotal)) -Title $Title -Icon $Icon
}

function Get-MclProperty {
    <# A property of an object or a dictionary, or $null (Set-StrictMode safe). #>
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name)) { return $Object[$Name] } return $null }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function ConvertTo-MclDateUtc {
    <# A Graph dateTimeTimeZone (UTC by default) to a UTC DateTime. #>
    param($Value)
    if ($null -eq $Value) { return $null }
    $text = [string](Get-MclProperty $Value 'dateTime')
    if (-not $text) { return $null }
    $zone = [string](Get-MclProperty $Value 'timeZone')
    $d = [datetime]::Parse($text, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
    if ($d.Kind -eq [DateTimeKind]::Utc) { return $d }
    if (-not $zone -or $zone -eq 'UTC') { return [datetime]::SpecifyKind($d, [DateTimeKind]::Utc) }
    try { return [TimeZoneInfo]::ConvertTimeToUtc([datetime]::SpecifyKind($d, [DateTimeKind]::Unspecified), [TimeZoneInfo]::FindSystemTimeZoneById($zone)) }
    catch { return [datetime]::SpecifyKind($d, [DateTimeKind]::Utc) }
}

function Format-MclRecurrence {
    <# A Graph patternedRecurrence in words: Weekly (Tuesday), 6 occurrences from 2026-10-13. #>
    param($Recurrence)
    if (-not $Recurrence) { return '' }
    $p = Get-MclProperty $Recurrence 'pattern'; $r = Get-MclProperty $Recurrence 'range'
    $type = [string](Get-MclProperty $p 'type')
    $interval = [int](Get-MclProperty $p 'interval')
    $days = @(Get-MclProperty $p 'daysOfWeek' | Where-Object { $_ } | ForEach-Object { (Get-Culture -Name 'en-US').TextInfo.ToTitleCase([string]$_) })
    $every = if ($interval -gt 1) { "every $interval " } else { '' }
    $text = switch ($type) {
        'daily' { if ($every) { "${every}days" } else { 'Daily' } }
        'weekly' { "$(if ($every) { "${every}weeks" } else { 'Weekly' }) ($($days -join ', '))" }
        'absoluteMonthly' { "$(if ($every) { "${every}months" } else { 'Monthly' }) (day $(Get-MclProperty $p 'dayOfMonth'))" }
        'relativeMonthly' { "$(if ($every) { "${every}months" } else { 'Monthly' }) ($(Get-MclProperty $p 'index') $($days -join ', '))" }
        'absoluteYearly' { "Yearly (day $(Get-MclProperty $p 'dayOfMonth') of month $(Get-MclProperty $p 'month'))" }
        'relativeYearly' { "Yearly ($(Get-MclProperty $p 'index') $($days -join ', ') of month $(Get-MclProperty $p 'month'))" }
        default { $type }
    }
    $from = [string](Get-MclProperty $r 'startDate')
    switch ([string](Get-MclProperty $r 'type')) {
        'endDate' { $text += ", from $from until $(Get-MclProperty $r 'endDate')" }
        'numbered' { $text += ", $(Get-MclProperty $r 'numberOfOccurrences') occurrences from $from" }
        default { $text += ", from $from, no end" }
    }
    return $text
}

function Test-MclSeriesInPeriod {
    <#
        Coarse check of a series against the period, from its range only (no request): $false when it surely
        ends before or starts after the period. A series that may cross the period is checked with its instances.
    #>
    param($Recurrence, [datetime]$Start, [datetime]$End)
    $r = Get-MclProperty $Recurrence 'range'
    if (-not $r) { return $true }
    $first = [string](Get-MclProperty $r 'startDate')
    if ($first -and [datetime]::ParseExact($first, 'yyyy-MM-dd', $null) -gt $End.Date) { return $false }
    if ([string](Get-MclProperty $r 'type') -eq 'endDate') {
        $last = [string](Get-MclProperty $r 'endDate')
        if ($last -and $last -ne '0001-01-01' -and [datetime]::ParseExact($last, 'yyyy-MM-dd', $null).AddDays(1) -lt $Start.Date) { return $false }
    }
    return $true
}

function Get-MclUserPath {
    <# /users/<address> with the address escaped. #>
    param([Parameter(Mandatory = $true)][string]$Address)
    return "/users/$([Uri]::EscapeDataString($Address))"
}

function Test-MclNoMailbox {
    <# The answer of Exchange Online when an address is not a mailbox it can open. #>
    param($Result)
    return $Result.Status -eq 404 -and $Result.ErrorCode -in 'ErrorInvalidUser', 'MailboxNotEnabledForRESTAPI', 'ResourceNotFound', 'Request_ResourceNotFound', 'ErrorNonExistentMailbox'
}

function Get-MclMailboxProblem {
    <# Why a mailbox cannot be read, in words. #>
    param($Result)
    switch ($Result.ErrorCode) {
        'ErrorInvalidUser' { 'not a mailbox of this tenant (external, deleted, group or contact)' }
        'MailboxNotEnabledForRESTAPI' { 'mailbox inactive, soft-deleted or on-premises' }
        'ErrorAccessDenied' { 'access denied (application access policy or RBAC for Applications scope)' }
        default { "$($Result.Status) $($Result.ErrorCode): $($Result.ErrorMessage)" }
    }
}

function Resolve-MclOrganizer {
    <#
    .SYNOPSIS
        The organizer(s) searched: addresses (primary, aliases, X500) and the state of the mailbox.
    .DESCRIPTION
        Every organizer of the list at once, in $batch calls: the directory (by any SMTP address, then by UPN),
        then the calendar of each one (a deleted user may keep a reachable mailbox for a while).
    .OUTPUTS
        One object per organizer: Input, DisplayName, PrimaryAddress, Addresses, UserId, State
        (Mailbox | NoMailbox | NotInDirectory), Detail.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string[]]$Identity)

    $g = $script:Graph
    $organizers = [Collections.Generic.List[object]]::new()
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($typed in $Identity) {
        $value = $typed.Trim()
        if (-not $value) { continue }
        if ($value -match $script:X500Pattern) {
            $x500 = ($value -replace '^(?i)x500:', '').ToLowerInvariant()
            if (-not $seen.Add($x500)) { continue }
            $organizers.Add([pscustomobject]@{ Input = $value; DisplayName = ''; PrimaryAddress = $x500; Addresses = @($x500); UserId = ''; Account = 'Unknown'; State = 'NotInDirectory'
                    Detail = 'X500 address (legacyExchangeDN): matched as is in the copies' })
            continue
        }
        $address = $value.ToLowerInvariant()
        if (-not $seen.Add($address)) { continue }
        $organizers.Add([pscustomobject]@{ Input = $value; DisplayName = ''; PrimaryAddress = $address; Addresses = @($address); UserId = ''; Account = 'Unknown'; State = 'NotInDirectory'; Detail = '' })
    }
    $smtp = @($organizers | Where-Object { $_.PrimaryAddress -notmatch '^/o=' })
    $progress = { param($done, $total) Write-MclProgress ($done / [Math]::Max(1, $total)) ('{0:N0}/{1:N0} organizers looked up' -f $done, $total) }

    # ---- directory: any SMTP address, then the UPN for those not found -----------------------------
    if ($g.CanReadUsers -and $smtp.Count) {
        $select = 'id,displayName,mail,userPrincipalName,proxyAddresses'
        $pending = $smtp
        foreach ($filter in "proxyAddresses/any(p:p eq 'smtp:{0}')", "userPrincipalName eq '{0}'") {
            if (-not $pending.Count) { break }
            $requests = for ($i = 0; $i -lt $pending.Count; $i++) {
                $f = [Uri]::EscapeDataString(($filter -f $pending[$i].PrimaryAddress.Replace("'", "''")))
                New-MclGraphRequest -Id "u$i" -Url "/users?`$filter=$f&`$select=$select"
            }
            $res = Invoke-MclGraphBatch -Requests @($requests) -OnProgress $progress
            $next = [Collections.Generic.List[object]]::new()
            for ($i = 0; $i -lt $pending.Count; $i++) {
                $org = $pending[$i]
                $user = if ($res["u$i"].Status -eq 200) { $res["u$i"].Values | Select-Object -First 1 } else { $null }
                # Not in the directory only when every lookup answered (a lookup in error leaves it unknown).
                if ($res["u$i"].Status -ne 200) { $org.Account = 'Lookup failed' }
                if (-not $user) { $next.Add($org); continue }
                $all = [Collections.Generic.List[string]]::new()
                $all.Add($org.PrimaryAddress)
                if ($user.mail) { $all.Add(([string]$user.mail).ToLowerInvariant()) }
                foreach ($p in @($user.proxyAddresses)) { if ([string]$p -match '^(?i)(smtp|x500):(.+)$') { $all.Add($Matches[2].ToLowerInvariant()) } }
                $org.DisplayName = [string]$user.displayName
                if ($user.mail) { $org.PrimaryAddress = ([string]$user.mail).ToLowerInvariant() }
                $org.Addresses = @($all | Select-Object -Unique)
                $org.UserId = [string]$user.id
                $org.Account = 'Present'
                $org.State = 'NoMailbox'
            }
            $pending = $next.ToArray()
        }
        foreach ($org in $pending) { $org.Account = if ($org.Account -eq 'Lookup failed') { 'Unknown' } else { 'Deleted' } }
    }

    # ---- the mailbox of each organizer: its calendar can be opened or not ---------------------------
    if ($smtp.Count) {
        $requests = for ($i = 0; $i -lt $smtp.Count; $i++) {
            $o = $smtp[$i]
            New-MclGraphRequest -Id "c$i" -Url "$(Get-MclUserPath $(if ($o.UserId) { $o.UserId } else { $o.PrimaryAddress }))/calendar?`$select=id"
        }
        $res = Invoke-MclGraphBatch -Requests @($requests) -OnProgress $progress
        for ($i = 0; $i -lt $smtp.Count; $i++) {
            $org = $smtp[$i]; $r = $res["c$i"]
            if ($r.Status -eq 200) {
                $org.State = 'Mailbox'
                $org.Detail = switch ($org.Account) { 'Deleted' { 'mailbox present, account not in the directory any more (user deleted recently)' } 'Unknown' { if ($g.CanReadUsers) { 'mailbox present, account not looked up (directory error)' } else { 'mailbox present' } } default { 'mailbox present' } }
            }
            elseif ($org.UserId) { $org.Detail = "user found, $(Get-MclMailboxProblem $r)" }
            else { $org.Detail = if ($g.CanReadUsers) { 'not in the directory: deleted mailbox or external address' } else { "not a reachable mailbox ($(Get-MclMailboxProblem $r))" } }
        }
    }
    # An organizer found twice (two aliases of one mailbox in the list) is kept once.
    $unique = [Collections.Generic.List[object]]::new()
    foreach ($o in $organizers) {
        if ($o.UserId -and @($unique | Where-Object UserId -eq $o.UserId).Count) { continue }
        $unique.Add($o)
    }
    return , $unique.ToArray()
}

function Get-MclRoomAddresses {
    <# Room mailboxes: places API (Place.Read.All), plus Search.Rooms and Search.RoomFile. #>
    param([Parameter(Mandatory = $true)][hashtable]$Settings)
    $rooms = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if ($script:Graph.CanReadPlaces) {
        # The places API pages with $skip only (no nextLink beyond $top).
        $skip = 0
        do {
            $r = Invoke-MclGraph -Path "places/microsoft.graph.room?`$select=emailAddress&`$top=999&`$skip=$skip"
            if ($r.Status -ne 200) { throw "Graph places -> $($r.Status) $($r.ErrorCode): $($r.ErrorMessage)" }
            $page = @($r.Body.value)
            foreach ($p in $page) { if ($p.emailAddress) { [void]$rooms.Add(([string]$p.emailAddress).ToLowerInvariant()) } }
            $skip += $page.Count
        } while ($page.Count -ge 999)
    }
    foreach ($a in @($Settings.Rooms)) { if ($a) { [void]$rooms.Add(([string]$a).ToLowerInvariant()) } }
    if ($Settings.RoomFile) { foreach ($a in (Read-MclAddressFile $Settings.RoomFile)) { [void]$rooms.Add($a.ToLowerInvariant()) } }
    return , @($rooms)
}

function Get-MclSearchMailboxes {
    <#
    .SYNOPSIS
        The mailboxes to search, each once, with the scope that brought it (Organizer, Rooms, Mailboxes, AllMailboxes).
    .OUTPUTS
        @{ Mailboxes = list of @{ Address; Scope; IsRoom }; Rooms = set of the room addresses known; Warnings }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Settings, [Parameter(Mandatory = $true)][pscustomobject]$Request, [object[]]$Organizers = @())

    $g = $script:Graph
    $warnings = [Collections.Generic.List[string]]::new()
    $list = [Collections.Generic.List[object]]::new()
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $add = { param([string]$Address, [string]$Scope) if ($Address -and $seen.Add($Address.ToLowerInvariant())) { $list.Add([pscustomobject]@{ Address = $Address.ToLowerInvariant(); Scope = $Scope; IsRoom = $false }) } }
    $scopes = @($Request.SearchIn)
    if ($Request.PSObject.Properties['Mode'] -and $Request.Mode -eq 'Rooms') {
        # Rooms mode: exactly the rooms given, whatever the organizer of their meetings.
        foreach ($a in @($Request.Room)) { & $add $a 'Rooms' }
        foreach ($m in $list) { $m.IsRoom = $true }
        $given = [Collections.Generic.HashSet[string]]::new([string[]]@($list | ForEach-Object Address), [StringComparer]::OrdinalIgnoreCase)
        foreach ($a in @($Settings.Rooms)) { if ($a) { [void]$given.Add([string]$a) } }
        return [pscustomobject]@{ Mailboxes = $list.ToArray(); Rooms = $given; Warnings = $warnings.ToArray() }
    }
    $rooms = @()
    if ($scopes -contains 'Rooms' -or $scopes -contains 'AllMailboxes') {
        if (-not $g.CanReadPlaces -and -not @($Settings.Rooms).Count -and -not $Settings.RoomFile) {
            $warnings.Add('Rooms: the application has no Place.Read.All and Search.Rooms / Search.RoomFile are empty: no room list. Rooms invited to a meeting are still found from its attendee list.')
        }
        $rooms = Get-MclRoomAddresses -Settings $Settings
    }
    $roomSet = [Collections.Generic.HashSet[string]]::new([string[]]@($rooms), [StringComparer]::OrdinalIgnoreCase)
    if ($scopes -contains 'Organizer') {
        foreach ($o in $Organizers) {
            if ($o.State -eq 'Mailbox') { & $add $(if ($o.UserId) { [string]$o.PrimaryAddress } else { [string]$o.Input }) 'Organizer' }
        }
    }
    if ($scopes -contains 'Rooms') { foreach ($a in $rooms) { & $add $a 'Rooms' } }
    if ($scopes -contains 'Mailboxes') {
        foreach ($a in @($Request.Mailboxes)) { & $add $a 'Mailboxes' }
        if ($Request.MailboxFile) { foreach ($a in (Read-MclAddressFile $Request.MailboxFile)) { & $add $a 'Mailboxes' } }
    }
    if ($scopes -contains 'AllMailboxes') {
        if (-not $g.CanReadUsers) { $warnings.Add('Every mailbox: the application has no User.Read.All, the list of the mailboxes cannot be read.') }
        else {
            foreach ($u in (Get-MclGraphAll -Path "users?`$select=mail,userPrincipalName&`$top=999")) { if ($u.mail) { & $add ([string]$u.mail) 'AllMailboxes' } }
        }
    }
    foreach ($m in $list) { $m.IsRoom = $roomSet.Contains($m.Address) }
    [pscustomobject]@{ Mailboxes = $list.ToArray(); Rooms = $roomSet; Warnings = $warnings.ToArray() }
}

function New-MclCopy {
    param([string]$Key, [string]$Mailbox, [string]$Role, [string]$Via, $Event, [string]$Result = '', [string]$Detail = '')
    [pscustomobject]@{
        MeetingId  = $Key
        Mailbox    = $Mailbox.ToLowerInvariant()
        Role       = $Role
        Via        = $Via
        EventId    = [string](Get-MclProperty $Event 'id')
        Subject    = [string](Get-MclProperty $Event 'subject')
        Response   = [string](Get-MclProperty (Get-MclProperty $Event 'responseStatus') 'response')
        ShowAs     = [string](Get-MclProperty $Event 'showAs')
        Cancelled  = [bool](Get-MclProperty $Event 'isCancelled')
        Action     = ''
        Result     = $Result
        HttpStatus = 0
        Detail     = $Detail
        Verified   = ''
        ActionUtc  = ''
        Occurrence = ''
        OccurrenceStart = ''
        SeriesId   = ''
    }
}

function New-MclMeeting {
    param([string]$Key, $Event, [string]$Mailbox, [hashtable]$Settings)
    $start = ConvertTo-MclDateUtc (Get-MclProperty $Event 'start')
    $end = ConvertTo-MclDateUtc (Get-MclProperty $Event 'end')
    $organizer = Get-MclProperty (Get-MclProperty $Event 'organizer') 'emailAddress'
    [pscustomobject]@{
        MeetingId     = $Key
        Subject       = [string](Get-MclProperty $Event 'subject')
        Organizer     = ([string](Get-MclProperty $organizer 'address')).ToLowerInvariant()
        OrganizerName = [string](Get-MclProperty $organizer 'name')
        OrganizerKey  = ''
        Kind          = if ([string](Get-MclProperty $Event 'type') -eq 'seriesMaster') { 'Series' } else { 'Single' }
        Start         = $start
        End           = $end
        StartText     = Format-MclDate $start $Settings.TimeZone
        EndText       = Format-MclDate $end $Settings.TimeZone
        NextInPeriod  = ''
        Recurrence    = Format-MclRecurrence (Get-MclProperty $Event 'recurrence')
        Location      = ''
        Cancelled     = [bool](Get-MclProperty $Event 'isCancelled')
        OrganizerCopy = 'Not checked'
        Attendees     = @()
        Copies        = [Collections.Generic.List[object]]::new()
        Selected      = $true
        Status        = 'Found'
        Notes         = [Collections.Generic.List[string]]::new()
        SubjectFromRoom = $false
        Scope         = 'Whole'
        Occurrences   = 0
        RecurrenceData = $null
        TimeZone      = ''
        NewOrganizer  = ''
        NewMeetingId  = ''
        TransferMethod = ''
    }
}

function Find-MclMeetings {
    <#
    .SYNOPSIS
        Finds the meetings of the request and every copy of them. Needs Connect-MclGraph first.
    .OUTPUTS
        The result: Meetings (each with its Copies), Organizers, Searched, Warnings, Counts.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Settings, [Parameter(Mandatory = $true)][pscustomobject]$Request)

    $started = [datetime]::UtcNow
    $g = $script:Graph
    $dot = $script:Dot
    $warnings = [Collections.Generic.List[string]]::new()
    $S = $Request.Start.ToString('yyyy-MM-ddTHH:mm:ss'); $E = $Request.End.ToString('yyyy-MM-ddTHH:mm:ss')
    # PowerShell variables ignore the case: no other $s / $e in this function.

    $roomsMode = $Request.PSObject.Properties['Mode'] -and $Request.Mode -eq 'Rooms'

    # ---- 1. Organizer (or the rooms) ----------------------------------------------------------------
    if ($roomsMode) {
        Write-MclNextStep "Rooms ($(@($Request.Room).Count))" 'Room'
        if (@($Request.Room).Count -le 10) { foreach ($a in @($Request.Room)) { Write-MclItem Info $a -Icon Room } }
        else { Write-MclItem Info ('{0:N0} rooms{1}' -f @($Request.Room).Count, $(if ($Request.RoomFile) { " from $([IO.Path]::GetFileName($Request.RoomFile))" })) -Icon Room }
        Write-MclItem Info 'Every meeting of these rooms in the period, whatever its organizer; a series is limited to its occurrences in the period.'
        $organizers = @()
    }
    else {
        Write-MclNextStep $(if (@($Request.Organizer).Count -gt 1) { "Organizers ($(@($Request.Organizer).Count))" } else { 'Organizer' }) 'User'
        $organizers = Resolve-MclOrganizer -Identity $Request.Organizer
        if ($organizers.Count -le 10) {
            foreach ($o in $organizers) {
                $name = if ($o.DisplayName) { "$($o.DisplayName) <$($o.PrimaryAddress)>" } else { $o.Input }
                $status = if ($o.State -eq 'Mailbox') { 'Ok' } else { 'Warn' }
                Write-MclItem $status "$name $dot $($o.Detail)$(if ($o.Addresses.Count -gt 1) { " $dot $($o.Addresses.Count) addresses compared" })" -Icon User
            }
        }
        else {
            # A long list: the counts here, every organizer in the report (Organizers.csv) and the log.
            foreach ($o in $organizers) { Write-MclLog 'INFO' "Organizer $($o.Input): $($o.State), $($o.Detail)" }
            $count = { param([string]$State) @($organizers | Where-Object State -eq $State).Count }
            Write-MclItem Ok ('{0:N0} organizers: {1:N0} with a mailbox {2} {3:N0} in the directory without a mailbox {2} {4:N0} not in the directory (deleted or X500)' -f $organizers.Count, (& $count 'Mailbox'), $dot, (& $count 'NoMailbox'), (& $count 'NotInDirectory')) -Icon People
        }
        if (-not $g.CanReadUsers) { $warnings.Add('No User.Read.All: the aliases of the organizer are not known, only the address typed is compared.'); Write-MclItem Warn $warnings[-1] }
        if (@($Request.SearchIn) -contains 'Organizer' -and -not @($organizers | Where-Object State -eq 'Mailbox').Count) {
            Write-MclItem Info "No organizer mailbox to search: the meetings are searched in the other mailboxes ($((@($Request.SearchIn) | Where-Object { $_ -ne 'Organizer' } | ForEach-Object { Get-MclScopeText $_ }) -join ', '))."
        }
    }
    $orgAddress = @{}
    foreach ($o in $organizers) { foreach ($a in $o.Addresses) { $orgAddress[$a] = $o } }

    # ---- 2. Mailboxes to search --------------------------------------------------------------------
    Write-MclNextStep 'Mailboxes to search' 'Search'
    $plan = Get-MclSearchMailboxes -Settings $Settings -Request $Request -Organizers $organizers
    foreach ($w in $plan.Warnings) { $warnings.Add($w); Write-MclItem Warn $w }
    $byScope = $plan.Mailboxes | Group-Object Scope
    foreach ($grp in $byScope) { Write-MclItem Info ("{0}: {1:N0} mailbox(es)" -f (Get-MclScopeText $grp.Name), $grp.Count) -Icon $(if ($grp.Name -eq 'Rooms') { 'Room' } else { 'Mail' }) }
    if (-not $plan.Mailboxes.Count) {
        $warnings.Add('No mailbox to search.')
        Write-MclItem Warn 'No mailbox to search: add a scope (rooms, a list of mailboxes, every mailbox).'
    }

    # ---- 3. Search ------------------------------------------------------------------------------------
    Write-MclNextStep 'Search' 'Calendar'
    $filter = [Uri]::EscapeDataString("type eq 'seriesMaster' or (end/dateTime ge '$S' and start/dateTime lt '$E')")
    $requests = foreach ($m in $plan.Mailboxes) {
        New-MclGraphRequest -Id $m.Address -Url "$(Get-MclUserPath $m.Address)/events?`$filter=$filter&`$select=$($script:EventSelect)&`$top=$($Settings.PageSize)"
    }
    $mailboxCount = $plan.Mailboxes.Count
    $results = Invoke-MclGraphBatch -Requests @($requests) -FollowPages -OnProgress {
        param($done, $total) Write-MclProgress ($done / [Math]::Max(1, $total)) ('{0:N0}/{1:N0} mailboxes searched' -f $done, $total)
    }
    $meetings = [ordered]@{}
    $searched = [ordered]@{ Mailboxes = $mailboxCount; Read = 0; Events = 0; NoMailbox = 0; Denied = 0; Errors = 0 }
    $ids = [Collections.Generic.HashSet[string]]::new([string[]]@($Request.MeetingId), [StringComparer]::OrdinalIgnoreCase)
    # The mailbox of each organizer searched: there, its own meetings are recognised by isOrganizer.
    $orgByMailbox = @{}
    foreach ($o in @($organizers | Where-Object State -eq 'Mailbox')) { foreach ($a in @($o.Addresses) + [string]$o.Input) { $orgByMailbox[$a.ToLowerInvariant()] = $o } }
    $problems = [Collections.Generic.List[string]]::new()
    foreach ($m in $plan.Mailboxes) {
        $r = $results[$m.Address]
        if ($r.Status -ne 200) {
            if (Test-MclNoMailbox $r) { $searched.NoMailbox++ }
            elseif ($r.Status -eq 403) { $searched.Denied++; $problems.Add("$($m.Address): $(Get-MclMailboxProblem $r)") }
            else { $searched.Errors++; $problems.Add("$($m.Address): $(Get-MclMailboxProblem $r)") }
            Write-MclLog 'WARN' "Search $($m.Address): $(Get-MclMailboxProblem $r)"
            continue
        }
        $searched.Read++
        $mailboxOrg = $orgByMailbox[$m.Address]
        foreach ($ev in $r.Values) {
            $searched.Events++
            $address = ([string](Get-MclProperty (Get-MclProperty (Get-MclProperty $ev 'organizer') 'emailAddress') 'address')).ToLowerInvariant()
            $own = $mailboxOrg -and [bool](Get-MclProperty $ev 'isOrganizer')
            $org = if ($own) { $mailboxOrg } else { $orgAddress[$address] }
            if ($roomsMode) {
                # Every meeting of the room: its organizer is the one shown in the copy.
                $own = [bool](Get-MclProperty $ev 'isOrganizer')
                $org = [pscustomobject]@{ PrimaryAddress = $(if ($own) { $m.Address } else { $address }) }
                if (-not $org.PrimaryAddress) { continue }
            }
            if (-not $org) { continue }
            $key = ([string](Get-MclProperty $ev 'iCalUId')).ToUpperInvariant()
            if (-not $key) { continue }
            if ($ids.Count -and -not $ids.Contains($key)) { continue }
            if ([string](Get-MclProperty $ev 'type') -eq 'seriesMaster' -and -not (Test-MclSeriesInPeriod (Get-MclProperty $ev 'recurrence') $Request.Start $Request.End)) { continue }
            if (-not $meetings.Contains($key)) {
                $meetings[$key] = New-MclMeeting -Key $key -Event $ev -Mailbox $m.Address -Settings $Settings
                $meetings[$key].OrganizerKey = [string]$org.PrimaryAddress
            }
            $role = if ($own) { 'Organizer' } elseif ($m.IsRoom) { 'Room' } else { 'Attendee' }
            $meetings[$key].Copies.Add((New-MclCopy -Key $key -Mailbox $m.Address -Role $role -Via $(switch ($m.Scope) { 'Organizer' { 'Organizer calendar' } 'Rooms' { 'Room search' } 'Mailboxes' { 'List search' } default { 'All mailboxes' } }) -Event $ev))
        }
    }
    Write-MclItem Ok ('{0:N0} mailbox(es) read {1} {2:N0} calendar item(s) {1} {3:N0} meeting(s) of {4}' -f $searched.Read, $dot, $searched.Events, $meetings.Count, $(if ($roomsMode) { 'the rooms' } elseif ($organizers.Count -gt 1) { "the $($organizers.Count) organizers" } else { 'the organizer' })) -Icon Calendar
    if ($searched.NoMailbox) { Write-MclItem Skip ('{0:N0} address(es) without a mailbox in Exchange Online (no licence, on-premises, or not a mailbox)' -f $searched.NoMailbox) }
    if ($searched.Denied + $searched.Errors) {
        $text = '{0:N0} mailbox(es) could not be read: {1}' -f ($searched.Denied + $searched.Errors), (($problems | Select-Object -First 3) -join ' | ')
        $warnings.Add($text); Write-MclItem Warn $text
    }

    # ---- series: keep those with an occurrence in the period ----------------------------------------
    $series = @($meetings.Values | Where-Object Kind -eq 'Series')
    if ($series.Count) {
        $requests = foreach ($mt in $series) {
            $c = $mt.Copies[0]
            New-MclGraphRequest -Id $mt.MeetingId -Url "$(Get-MclUserPath $c.Mailbox)/events/$([Uri]::EscapeDataString($c.EventId))/instances?startDateTime=${S}Z&endDateTime=${E}Z&`$select=start&`$top=50"
        }
        $inst = Invoke-MclGraphBatch -Requests @($requests)
        $dropped = 0
        foreach ($mt in $series) {
            $r = $inst[$mt.MeetingId]
            if ($r.Status -eq 200 -and -not $r.Values.Count) { $meetings.Remove($mt.MeetingId); $dropped++; continue }
            if ($r.Status -eq 200) {
                $next = $r.Values | ForEach-Object { ConvertTo-MclDateUtc $_.start } | Sort-Object | Select-Object -First 1
                $mt.NextInPeriod = Format-MclDate $next $Settings.TimeZone
            }
            else { $mt.Notes.Add("Occurrences in the period not checked ($($r.Status) $($r.ErrorCode)): the series is kept.") }
        }
        if ($dropped) { Write-MclLog 'INFO' "$dropped series without an occurrence in the period left out." }
    }

    # ---- 4. Attendees -------------------------------------------------------------------------------
    Write-MclNextStep 'Attendees, rooms and groups' 'People'
    $list = @($meetings.Values)
    if ($list.Count) { Complete-MclMeetings -Meetings $list -Settings $Settings -Organizers $organizers -Rooms $plan.Rooms -Warnings $warnings }
    $list = @($list | Where-Object Status -ne 'Appointment')

    # Subject filter, once the real subject is known (a room may show the organizer's name instead).
    if ($Request.Subject) {
        $pattern = if ($Request.Subject -match '[*?]') { $Request.Subject } else { "*$($Request.Subject)*" }
        $before = $list.Count
        $list = @($list | Where-Object { $_.Subject -like $pattern })
        Write-MclItem Info ("Subject '{0}': {1} of {2} meeting(s) kept" -f $Request.Subject, $list.Count, $before)
    }
    $list = @($list | Sort-Object Start, Subject)

    # Rooms: a series is acted on only by its occurrences in the period (a room closed for two weeks does
    # not end a series of a year), unless every occurrence of the series is in the period.
    if ($roomsMode -and $list.Count) { Split-MclSeriesOccurrences -Meetings $list -Settings $Settings -Start $Request.Start -End $Request.End }
    if ($roomsMode) {
        # The organizers of the meetings found, for the report (Organizers tab): who is concerned.
        $organizers = @($list | Group-Object OrganizerKey | ForEach-Object {
                $first = $_.Group[0]
                $state = switch ($first.OrganizerCopy) { 'Present' { 'Mailbox' } 'Absent' { 'Mailbox' } 'Mailbox deleted' { 'NotInDirectory' } default { 'Unknown' } }
                [pscustomobject]@{ Input = $_.Name; DisplayName = [string]$first.OrganizerName; PrimaryAddress = $_.Name; Addresses = @($_.Name); UserId = ''; Account = 'Unknown'; State = $state
                    Detail = '{0} meeting(s) in the rooms {1} organizer copy {2}' -f $_.Count, $dot, ($first.OrganizerCopy).ToLowerInvariant() }
            })
    }

    $result = [pscustomobject]@{
        Tool            = 'Meeting Cleanup'
        Version         = $script:ToolVersion
        Action          = 'Report'
        Status          = 'Completed'
        Error           = ''
        StartedUtc      = $started.ToString('o')
        CompletedUtc    = ''
        DurationSeconds = 0.0
        Request         = [pscustomobject]@{
            Mode = $(if ($roomsMode) { 'Rooms' } else { 'Organizers' }); Room = @($Request.Room); RoomFile = [string](Get-MclProperty $Request 'RoomFile')
            Organizer = @($Request.Organizer); Start = $Request.Start.ToString('o'); End = $Request.End.ToString('o')
            StartText = (Format-MclDate $Request.Start $Settings.TimeZone) -replace ' 00:00$', ''; EndText = Format-MclDate $Request.End $Settings.TimeZone -PeriodEnd
            Subject = $Request.Subject; MeetingId = @($Request.MeetingId); SearchIn = @($Request.SearchIn); Mailboxes = @($Request.Mailboxes).Count; MailboxFile = $Request.MailboxFile
            TimeZone = (Get-MclTimeZone $Settings.TimeZone).Id
        }
        Tenant          = [string]$g.TenantGuid
        Organization    = [string]$Settings.Organization
        AppId           = [string]$Settings.AppId
        AppName         = [string]$g.AppName
        Organizers      = @($organizers)
        Searched        = [pscustomobject]$searched
        Meetings        = [Collections.Generic.List[object]]::new()
        Warnings        = $warnings
        Counts          = $null
    }
    foreach ($m in $list) { $result.Meetings.Add($m) }
    Update-MclResultCounts $result
    $result.CompletedUtc = [datetime]::UtcNow.ToString('o')
    $result.DurationSeconds = [Math]::Round(([datetime]::UtcNow - $started).TotalSeconds, 1)
    if ($warnings.Count -and $result.Status -eq 'Completed') { $result.Status = 'Warning' }
    return $result
}

function Split-MclSeriesOccurrences {
    <#
        Rooms mode: each series is acted on by its occurrences in the period that the rooms searched hold (an
        occurrence moved to another room is not one of them). Every copy of the series (organizer, attendees,
        rooms) is replaced by these occurrences, each a copy of its own (Occurrence, the ID of the occurrence in
        that mailbox, SeriesId). A series whose every occurrence is held by the rooms in the period is kept whole.
        When the occurrences of the organizer or of a room searched cannot be read, the meeting is left as it is
        (Not processed): never the whole series, never the attendees without their organizer.
    #>
    param([Parameter(Mandatory = $true)][object[]]$Meetings, [Parameter(Mandatory = $true)][hashtable]$Settings, [datetime]$Start, [datetime]$End)

    $series = @($Meetings | Where-Object { $_.Kind -eq 'Series' -and @($_.Copies | Where-Object EventId).Count })
    if (-not $series.Count) { return }
    $zone = Get-MclTimeZone $Settings.TimeZone
    $firstDay = [TimeZoneInfo]::ConvertTimeFromUtc($Start, $zone).Date
    $lastDay = [TimeZoneInfo]::ConvertTimeFromUtc($End.AddTicks(-1), $zone).Date
    $S = $Start.ToString('yyyy-MM-ddTHH:mm:ss'); $E = $End.ToString('yyyy-MM-ddTHH:mm:ss')
    $select = 'id,subject,start,originalStart,showAs,responseStatus,isCancelled'
    $requests = [Collections.Generic.List[object]]::new()
    $copies = @{}
    $byMeeting = @{}
    for ($i = 0; $i -lt $series.Count; $i++) {
        $byMeeting[$i] = [Collections.Generic.List[string]]::new()
        foreach ($c in @($series[$i].Copies | Where-Object EventId)) {
            $id = "o$($requests.Count)"
            $copies[$id] = $c
            $byMeeting[$i].Add($id)
            $requests.Add((New-MclGraphRequest -Id $id -Url "$(Get-MclUserPath $c.Mailbox)/events/$([Uri]::EscapeDataString($c.EventId))/instances?startDateTime=${S}Z&endDateTime=${E}Z&`$select=$select&`$top=200"))
        }
    }
    $res = Invoke-MclGraphBatch -Requests $requests.ToArray() -FollowPages -OnProgress { param($done, $total) Write-MclProgress ($done / [Math]::Max(1, $total)) ('{0:N0}/{1:N0} series copies: occurrences read' -f $done, $total) }

    # The occurrences held by the rooms searched; the meetings that cannot be split safely.
    $keys = @{}; $untouched = @{}
    for ($i = 0; $i -lt $series.Count; $i++) {
        $roomIds = @($byMeeting[$i] | Where-Object { $copies[$_].Via -eq 'Room search' })
        $unread = @(@($roomIds) + @($byMeeting[$i] | Where-Object { $copies[$_].Role -eq 'Organizer' }) | Where-Object { $res[$_].Status -ne 200 })
        if ($series[$i].OrganizerCopy -eq 'Not read') { $untouched[$i] = "the copy of its organizer could not be read"; continue }
        if (-not $roomIds.Count) { $untouched[$i] = 'no copy of the rooms searched to read its occurrences from'; continue }
        if ($unread.Count) { $untouched[$i] = "occurrences of the period not read in $($copies[$unread[0]].Mailbox) ($(Get-MclMailboxProblem $res[$unread[0]]))"; continue }
        $set = [Collections.Generic.HashSet[string]]::new()
        foreach ($id in $roomIds) { foreach ($o in @($res[$id].Values)) { if (-not [bool](Get-MclProperty $o 'isCancelled')) { [void]$set.Add((Get-MclTimeKey $o)) } } }
        $keys[$i] = $set
    }

    # The whole series when the rooms hold every one of its occurrences in the period: a numbered series by its
    # count, a series with an end date against every occurrence of the organizer (read around its range).
    $whole = @{}
    $wide = [Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $series.Count; $i++) {
        if (-not $keys.Contains($i) -or -not $keys[$i].Count) { continue }
        $range = Get-MclProperty $series[$i].RecurrenceData 'range'
        $type = [string](Get-MclProperty $range 'type')
        if ($type -eq 'numbered') { if ($keys[$i].Count -ge [int](Get-MclProperty $range 'numberOfOccurrences')) { $whole[$i] = $true } }
        elseif ($type -eq 'endDate') {
            $from = [string](Get-MclProperty $range 'startDate'); $to = [string](Get-MclProperty $range 'endDate')
            $org = $byMeeting[$i] | Where-Object { $copies[$_].Role -eq 'Organizer' } | Select-Object -First 1
            if (-not $from -or -not $to -or -not $org) { continue }
            $d0 = [datetime]::ParseExact($from, 'yyyy-MM-dd', $null); $d1 = [datetime]::ParseExact($to, 'yyyy-MM-dd', $null)
            if ($d0 -lt $firstDay.AddDays(-1) -or $d1 -gt $lastDay.AddDays(1)) { continue }
            $c = $copies[$org]
            $wide.Add((New-MclGraphRequest -Id "w$i" -Url "$(Get-MclUserPath $c.Mailbox)/events/$([Uri]::EscapeDataString($c.EventId))/instances?startDateTime=$($d0.AddDays(-31).ToString('yyyy-MM-dd'))T00:00:00Z&endDateTime=$($d1.AddDays(32).ToString('yyyy-MM-dd'))T00:00:00Z&`$select=id,start,originalStart,isCancelled&`$top=500"))
        }
    }
    if ($wide.Count) {
        $all = Invoke-MclGraphBatch -Requests $wide.ToArray() -FollowPages
        foreach ($id in $all.Keys) {
            if ($all[$id].Status -ne 200) { continue }
            $i = [int]$id.Substring(1)
            $every = [Collections.Generic.HashSet[string]]::new()
            foreach ($o in @($all[$id].Values)) { if (-not [bool](Get-MclProperty $o 'isCancelled')) { [void]$every.Add((Get-MclTimeKey $o)) } }
            if ($every.Count -and $every.SetEquals($keys[$i])) { $whole[$i] = $true }
        }
    }

    $split = 0; $occurrences = 0; $left = 0
    for ($i = 0; $i -lt $series.Count; $i++) {
        $m = $series[$i]
        if ($whole.Contains($i)) { $m.Notes.Add('Every occurrence of the series is in the period, in the rooms searched: the series is handled as a whole.'); continue }
        $m.Scope = 'Occurrences'
        if ($untouched.Contains($i)) {
            $left++
            foreach ($c in @($m.Copies | Where-Object EventId)) { $c.EventId = ''; $c.Result = 'Not processed'; $c.Detail = "the series is left as it is: $($untouched[$i])" }
            $m.Occurrences = 0
            $m.Notes.Add("Not acted on: $($untouched[$i]). Acting on the attendees only, or on the whole series, would not be right.")
            continue
        }
        $split++
        $list = [Collections.Generic.List[object]]::new()
        foreach ($c in @($m.Copies)) {
            $id = $byMeeting[$i] | Where-Object { [object]::ReferenceEquals($copies[$_], $c) } | Select-Object -First 1
            if (-not $id) { $list.Add($c); continue }
            $r = $res[$id]
            if ($r.Status -ne 200) {
                $c.EventId = ''; $c.Result = 'Not processed'; $c.Detail = "occurrences of the period not read ($(Get-MclMailboxProblem $r)): the copy is left as it is"
                $list.Add($c); continue
            }
            # The occurrences of the rooms searched only (one moved to another room stays as it is).
            $mine = @($r.Values | Where-Object { $keys[$i].Contains((Get-MclTimeKey $_)) })
            if (-not $mine.Count) {
                $c.EventId = ''; $c.Result = 'No copy'; $c.Detail = 'no occurrence of the rooms searched in the period in this mailbox (declined, removed or moved)'
                $list.Add($c); continue
            }
            foreach ($o in ($mine | Sort-Object { ConvertTo-MclDateUtc $_.start })) {
                $occ = New-MclCopy -Key $m.MeetingId -Mailbox $c.Mailbox -Role $c.Role -Via $c.Via -Event $o
                if (-not $occ.Subject) { $occ.Subject = $c.Subject }
                $when = ConvertTo-MclDateUtc $o.start
                $occ.Occurrence = Format-MclDate $when $Settings.TimeZone
                $occ.OccurrenceStart = $when.ToString('o')
                $occ.SeriesId = $c.EventId
                $list.Add($occ)
                $occurrences++
            }
        }
        $m.Copies.Clear()
        foreach ($c in $list) { $m.Copies.Add($c) }
        $m.Occurrences = $keys[$i].Count
        $m.Notes.Add(('{0} occurrence(s) in the period in the rooms searched: only they are acted on, the series goes on outside the period.' -f $keys[$i].Count))
    }
    Write-MclItem Info ('{0} series limited to their occurrences in the period ({1:N0} occurrence copies){2}{3}' -f $split, $occurrences, $(if ($whole.Count) { " $($script:Dot) $($whole.Count) entirely in the period, handled whole" }), $(if ($left) { " $($script:Dot) $left left as they are (occurrences not read)" })) -Icon Calendar
    if ($left) { Write-MclItem Warn ('{0} series left as they are: the occurrences of their organizer or of a room could not be read (see the notes).' -f $left) }
}

function Complete-MclMeetings {
    <#
        Details of each meeting (subject, attendees, location) from its best copy, then the copy of every
        attendee, room and group member not found yet, looked up by iCalUId.
    #>
    param(
        [Parameter(Mandatory = $true)][object[]]$Meetings,
        [Parameter(Mandatory = $true)][hashtable]$Settings,
        [object[]]$Organizers = @(),
        $Rooms,
        [Collections.Generic.List[string]]$Warnings
    )

    $g = $script:Graph
    $dot = $script:Dot
    $roomSet = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($a in @($Rooms)) { if ($a) { [void]$roomSet.Add([string]$a) } }
    $orgAddress = @{}
    foreach ($o in $Organizers) { foreach ($a in $o.Addresses) { $orgAddress[$a] = $o } }
    $rank = @{ Organizer = 0; Attendee = 1; Room = 2 }

    # ---- details from the best copy --------------------------------------------------------------
    $requests = foreach ($m in $Meetings) {
        $ref = $m.Copies | Where-Object EventId | Sort-Object { $rank[$_.Role] } | Select-Object -First 1
        New-MclGraphRequest -Id $m.MeetingId -Url "$(Get-MclUserPath $ref.Mailbox)/events/$([Uri]::EscapeDataString($ref.EventId))?`$select=subject,attendees,organizer,location,start,end,type,recurrence,isCancelled,originalStartTimeZone"
    }
    $details = Invoke-MclGraphBatch -Requests @($requests) -OnProgress { param($done, $total) Write-MclProgress ($done / [Math]::Max(1, $total)) ('{0:N0}/{1:N0} meetings read' -f $done, $total) }
    foreach ($m in $Meetings) {
        $r = $details[$m.MeetingId]
        $ref = $m.Copies | Where-Object EventId | Sort-Object { $rank[$_.Role] } | Select-Object -First 1
        if ($r.Status -ne 200) { $m.Notes.Add("Details not read from $($ref.Mailbox): $($r.Status) $($r.ErrorCode)."); continue }
        $ev = $r.Body
        $m.Subject = [string]$ev.subject
        $m.SubjectFromRoom = $ref.Role -eq 'Room'
        $m.Location = [string](Get-MclProperty (Get-MclProperty $ev 'location') 'displayName')
        $m.Attendees = @(@($ev.attendees) | Where-Object { $_ } | ForEach-Object {
                [pscustomobject]@{ Address = ([string]$_.emailAddress.address).ToLowerInvariant(); Name = [string]$_.emailAddress.name; Type = [string]$_.type }
            })
        $org = Get-MclProperty (Get-MclProperty $ev 'organizer') 'emailAddress'
        if ($org -and $org.address) { $m.Organizer = ([string]$org.address).ToLowerInvariant(); $m.OrganizerName = [string]$org.name }
        if ($ev.recurrence) { $m.Recurrence = Format-MclRecurrence $ev.recurrence; $m.RecurrenceData = $ev.recurrence }
        $m.TimeZone = [string](Get-MclProperty $ev 'originalStartTimeZone')
        if ($ref.Role -eq 'Organizer' -and -not @($m.Attendees).Count) { $m.Notes.Add('Appointment without attendees: not a meeting.') }
    }
    # An appointment of the organizer (no attendee) is not a meeting: left out.
    foreach ($m in @($Meetings)) {
        if (@($m.Copies | Where-Object Role -eq 'Organizer').Count -and -not @($m.Attendees).Count -and $details[$m.MeetingId].Status -eq 200) { $m.Status = 'Appointment' }
    }

    # ---- copies of the attendees, by iCalUId ---------------------------------------------------------
    $lookups = [ordered]@{}
    # Meetings whose organizer copy could not be read (denied, throttled, error): not a deleted mailbox.
    $orgNotRead = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $want = {
        param($m, [string]$Address, [string]$Type, [string]$Via)
        if (-not $Address -or $Address -notmatch '@') { return }
        if (@($m.Copies | Where-Object { $_.Mailbox -eq $Address }).Count) { return }
        if (($orgAddress.ContainsKey($Address) -or $Address -eq $m.Organizer) -and @($m.Copies | Where-Object { $_.Role -eq 'Organizer' -and $_.EventId }).Count) { return }
        $id = "$($m.MeetingId)|$Address"
        if (-not $lookups.Contains($id)) { $lookups[$id] = [pscustomobject]@{ Meeting = $m; Address = $Address; Type = $Type; Via = $Via } }
    }
    foreach ($m in $Meetings) {
        if ($m.Status -eq 'Appointment') { continue }
        & $want $m $m.Organizer 'organizer' 'Attendee list'
        foreach ($a in @($m.Attendees)) { & $want $m $a.Address $a.Type 'Attendee list' }
    }
    $groupFor = @{}
    $pass = 0
    while ($lookups.Count -and $pass -lt 2) {
        $pass++
        $requests = foreach ($id in $lookups.Keys) {
            $l = $lookups[$id]
            $f = [Uri]::EscapeDataString("iCalUId eq '$($l.Meeting.MeetingId)'")
            New-MclGraphRequest -Id $id -Url "$(Get-MclUserPath $l.Address)/events?`$filter=$f&`$select=$($script:EventSelect)"
        }
        $found = Invoke-MclGraphBatch -Requests @($requests) -OnProgress { param($done, $total) Write-MclProgress ($done / [Math]::Max(1, $total)) ('{0:N0}/{1:N0} attendee copies looked up' -f $done, $total) }
        $unknown = [ordered]@{}
        foreach ($id in $lookups.Keys) {
            $l = $lookups[$id]; $r = $found[$id]; $m = $l.Meeting
            $isOrganizer = $orgAddress.ContainsKey($l.Address) -or $l.Address -eq $m.Organizer
            $role = if ($isOrganizer) { 'Organizer' } elseif ($l.Type -eq 'resource' -or $roomSet.Contains($l.Address)) { 'Room' } else { 'Attendee' }
            if ($r.Status -eq 200) {
                if ($r.Values.Count) {
                    foreach ($ev in $r.Values) {
                        # The same item reached through another address of the mailbox (alias) is one copy.
                        if (@($m.Copies | Where-Object { $_.EventId -and $_.EventId -eq [string]$ev.id }).Count) { continue }
                        $copyRole = if ($role -eq 'Organizer' -and -not (Get-MclProperty $ev 'isOrganizer')) { 'Attendee' } else { $role }
                        $m.Copies.Add((New-MclCopy -Key $m.MeetingId -Mailbox $l.Address -Role $copyRole -Via $l.Via -Event $ev))
                    }
                }
                else { $m.Copies.Add((New-MclCopy -Key $m.MeetingId -Mailbox $l.Address -Role $role -Via $l.Via -Event $null -Result 'No copy' -Detail 'no copy in this mailbox (declined and removed, or never received)')) }
                continue
            }
            if ($r.Status -eq 404 -and $r.ErrorCode -eq 'ErrorInvalidUser' -and $role -ne 'Organizer' -and $l.Via -eq 'Attendee list' -and $g.CanReadGroups) {
                if (-not $unknown.Contains($l.Address)) { $unknown[$l.Address] = [Collections.Generic.List[object]]::new() }
                $unknown[$l.Address].Add($l)
                continue
            }
            $notRead = $role -eq 'Organizer' -and -not (Test-MclNoMailbox $r)
            if ($notRead) { [void]$orgNotRead.Add($m.MeetingId) }
            $text = if ($notRead) { "organizer copy not read ($(Get-MclMailboxProblem $r)): the meeting cannot be cancelled" } elseif ($role -eq 'Organizer') { 'organizer mailbox deleted or not reachable' } else { Get-MclMailboxProblem $r }
            $m.Copies.Add((New-MclCopy -Key $m.MeetingId -Mailbox $l.Address -Role $role -Via $l.Via -Event $null -Result 'Not processed' -Detail $text))
        }
        $lookups = [ordered]@{}
        if (-not $unknown.Count) { break }

        # ---- groups invited: their members ---------------------------------------------------------------
        $requests = foreach ($address in $unknown.Keys) {
            $f = [Uri]::EscapeDataString("proxyAddresses/any(p:p eq 'smtp:$address')")
            New-MclGraphRequest -Id $address -Url "/groups?`$filter=$f&`$select=id,displayName,mail"
        }
        $groups = Invoke-MclGraphBatch -Requests @($requests)
        $memberRequests = [Collections.Generic.List[object]]::new()
        foreach ($address in $unknown.Keys) {
            $grp = if ($groups[$address].Status -eq 200) { $groups[$address].Values | Select-Object -First 1 } else { $null }
            if (-not $grp) {
                foreach ($l in $unknown[$address]) { $l.Meeting.Copies.Add((New-MclCopy -Key $l.Meeting.MeetingId -Mailbox $address -Role 'Attendee' -Via $l.Via -Event $null -Result 'Not processed' -Detail 'not a mailbox of this tenant (external, deleted or contact)')) }
                continue
            }
            $groupFor[$address] = @{ Name = [string]$grp.displayName; Id = [string]$grp.id; Meetings = $unknown[$address] }
            $memberRequests.Add((New-MclGraphRequest -Id $address -Url "/groups/$($grp.id)/transitiveMembers/microsoft.graph.user?`$select=mail&`$top=999"))
        }
        if ($memberRequests.Count) {
            $members = Invoke-MclGraphBatch -Requests $memberRequests.ToArray() -FollowPages
            foreach ($address in $groupFor.Keys) {
                $info = $groupFor[$address]
                $mails = @($members[$address].Values | ForEach-Object { ([string]$_.mail).ToLowerInvariant() } | Where-Object { $_ })
                foreach ($l in $info.Meetings) {
                    $l.Meeting.Copies.Add((New-MclCopy -Key $l.Meeting.MeetingId -Mailbox $address -Role 'Group' -Via $l.Via -Event $null -Result 'Expanded' -Detail ("group '{0}': {1} member(s) with a mailbox" -f $info.Name, $mails.Count)))
                    foreach ($mail in $mails) { & $want $l.Meeting $mail 'required' "Group $($info.Name)" }
                }
            }
            $groupFor = @{}
        }
    }

    # ---- summary per meeting ---------------------------------------------------------------------------
    foreach ($m in $Meetings) {
        $orgCopy = @($m.Copies | Where-Object { $_.Role -eq 'Organizer' -and $_.EventId })
        $orgRow = @($m.Copies | Where-Object { $_.Role -eq 'Organizer' -and -not $_.EventId })
        $m.OrganizerCopy = if ($orgCopy.Count) { 'Present' } elseif ($orgNotRead.Contains($m.MeetingId)) { 'Not read' } elseif ($orgRow.Count -and $orgRow[0].Result -eq 'No copy') { 'Absent' } elseif ($orgRow.Count) { 'Mailbox deleted' } else { 'Not checked' }
        $best = $m.Copies | Where-Object { $_.EventId -and $_.Role -ne 'Room' -and $_.Subject } | Sort-Object { $rank[$_.Role] } | Select-Object -First 1
        if ($best -and ($m.SubjectFromRoom -or -not $m.Subject)) { $m.Subject = $best.Subject; $m.SubjectFromRoom = $false }
        $m.Cancelled = [bool]@($m.Copies | Where-Object Cancelled).Count
    }
    $total = @($Meetings | ForEach-Object { @($_.Copies | Where-Object EventId) }).Count
    $rooms = @($Meetings | ForEach-Object { @($_.Copies | Where-Object { $_.EventId -and $_.Role -eq 'Room' }) }).Count
    $skipped = @($Meetings | ForEach-Object { @($_.Copies | Where-Object Result -eq 'Not processed') }).Count
    $appointments = @($Meetings | Where-Object Status -eq 'Appointment').Count
    Write-MclItem Ok ('{0:N0} cop{1} of {2:N0} meeting(s) {3} {4:N0} in rooms {3} organizer copy present for {5:N0}' -f $total, $(if ($total -eq 1) { 'y' } else { 'ies' }), ($Meetings.Count - $appointments), $dot, $rooms, @($Meetings | Where-Object OrganizerCopy -eq 'Present').Count) -Icon People
    if ($skipped) { Write-MclItem Skip ('{0:N0} attendee(s) not processed: external, deleted, on-premises or not reachable (listed in the report)' -f $skipped) }
    if ($appointments) { Write-MclItem Skip ('{0:N0} appointment(s) of the organizer without attendees left out' -f $appointments) }
    if (-not $g.CanReadGroups) {
        $text = 'No GroupMember.Read.All: the members of a group invited to a meeting are not looked up.'
        if ($Warnings) { $Warnings.Add($text) }
        Write-MclItem Warn $text
    }
}

function Update-MclResultCounts {
    <# Totals of a result: meetings, copies, mailboxes, actions. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result)
    $meetings = @($Result.Meetings | Where-Object { $_.Status -ne 'Appointment' })
    $copies = @($meetings | ForEach-Object { @($_.Copies) })
    $real = @($copies | Where-Object EventId)
    $Result.Counts = [pscustomobject]@{
        Meetings     = $meetings.Count
        Series       = @($meetings | Where-Object Kind -eq 'Series').Count
        Selected     = @($meetings | Where-Object Selected).Count
        Copies       = $real.Count
        Mailboxes    = @($real | ForEach-Object Mailbox | Select-Object -Unique).Count
        RoomCopies   = @($real | Where-Object Role -eq 'Room').Count
        OrganizerCopies = @($real | Where-Object Role -eq 'Organizer').Count
        NotProcessed = @($copies | Where-Object Result -eq 'Not processed').Count
        Removed      = @($copies | Where-Object Result -eq 'Removed').Count
        Cancelled    = @($copies | Where-Object Result -eq 'Cancelled').Count
        AlreadyGone  = @($copies | Where-Object Result -eq 'Already gone').Count
        Kept         = @($copies | Where-Object Result -eq 'Kept').Count
        Failed       = @($copies | Where-Object Result -eq 'Failed').Count
        Restored     = @($copies | Where-Object Result -eq 'Restored').Count
        NotFound     = @($copies | Where-Object Result -eq 'Not found').Count
        AlreadyPresent = @($copies | Where-Object Result -eq 'Already present').Count
        NotRestorable = @($copies | Where-Object Result -eq 'Not restorable').Count
        Transferred  = @($Result.Meetings | Where-Object Status -eq 'Transferred').Count
        OccurrenceCopies = @($real | Where-Object Occurrence).Count
        Organizers   = @($Result.Organizers).Count
    }
}
