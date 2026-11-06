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
    Version : 1.3.1
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
    <# A Graph dateTimeTimeZone (UTC by default) to a UTC DateTime (compiled: MeetingCleanupNative.Fast.ToUtc). #>
    param($Value)
    return [MeetingCleanupNative.Fast]::ToUtc($Value)
}

function Format-MclRecurrence {
    <# A Graph patternedRecurrence in words: Weekly (Tuesday), 6 occurrences from 2026-10-13 (compiled). #>
    param($Recurrence)
    return [MeetingCleanupNative.Fast]::Recurrence($Recurrence)
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
    <# A copy of a meeting in one mailbox (compiled: MeetingCleanupNative.Fast.NewCopy, the same properties). #>
    param([string]$Key, [string]$Mailbox, [string]$Role, [string]$Via, $Event, [string]$Result = '', [string]$Detail = '')
    [MeetingCleanupNative.Fast]::NewCopy($Key, $Mailbox, $Role, $Via, $Event, $Result, $Detail)
}

function New-MclMeeting {
    <# A meeting found in a calendar (compiled: MeetingCleanupNative.Fast.NewMeeting). #>
    param([string]$Key, $Event, [string]$Mailbox, [hashtable]$Settings)
    [MeetingCleanupNative.Fast]::NewMeeting($Key, $Event, (Get-MclTimeZone $Settings.TimeZone))
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
        if ([string](Get-MclProperty $Request 'SeriesScope') -eq 'Occurrences') { Write-MclItem Info 'A series: only its occurrences in the period are acted on (they can be chosen one by one in the window); the series goes on outside the period.' }
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
    $fast = [MeetingCleanupNative.Fast]
    $zone = Get-MclTimeZone $Settings.TimeZone
    $via = @{ Organizer = 'Organizer calendar'; Rooms = 'Room search'; Mailboxes = 'List search' }
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
        $mailboxVia = if ($via.ContainsKey($m.Scope)) { $via[$m.Scope] } else { 'All mailboxes' }
        # Every calendar item of the mailbox, filtered in one compiled call (most are not the organizer's): the
        # organizer compared (own meetings by isOrganizer), rooms mode every meeting, the meeting IDs asked for.
        $searched.Events += $r.Values.Count
        foreach ($x in $fast::MatchEvents($r.Values, $orgAddress, $mailboxOrg, [bool]$roomsMode, $m.Address, $ids)) {
            if ($x.IsSeriesMaster -and -not (Test-MclSeriesInPeriod $x.Recurrence $Request.Start $Request.End)) { continue }
            $key = $x.Key
            if (-not $meetings.Contains($key)) {
                $meetings[$key] = $fast::NewMeeting($key, $x.Event, $zone)
                $meetings[$key].OrganizerKey = $x.OrganizerKey
            }
            $role = if ($x.Own) { 'Organizer' } elseif ($m.IsRoom) { 'Room' } else { 'Attendee' }
            $meetings[$key].Copies.Add($fast::NewCopy($key, $m.Address, $role, $mailboxVia, $x.Event, '', ''))
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
                $next = $null
                foreach ($o in $r.Values) { $d = $fast::ToUtc($fast::Prop($o, 'start')); if ($null -ne $d -and ($null -eq $next -or $d -lt $next)) { $next = $d } }
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
    # not end a series of a year), unless every occurrence of the series is in the period. Organizers with
    # -SeriesScope Occurrences: the occurrences of the period of each series, to be chosen in the window.
    $byOccurrence = -not $roomsMode -and [string](Get-MclProperty $Request 'SeriesScope') -eq 'Occurrences'
    if ($roomsMode -and $list.Count) { Split-MclSeriesOccurrences -Meetings $list -Settings $Settings -Start $Request.Start -End $Request.End -Source Rooms }
    elseif ($byOccurrence -and @($list | Where-Object Kind -eq 'Series').Count) { Split-MclSeriesOccurrences -Meetings $list -Settings $Settings -Start $Request.Start -End $Request.End -Source Organizer }
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
            SeriesScope = $(if ($roomsMode -or $byOccurrence) { 'Occurrences' } else { 'Whole' })
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
        A series acted on by its occurrences in the period, not whole. Every copy of the series (organizer,
        attendees, rooms) is replaced by these occurrences, each a copy of its own (Occurrence, the ID of the
        occurrence in that mailbox, SeriesId); the window can then leave some of them out (SkippedOccurrences).
          -Source Rooms      rooms mode: the occurrences the rooms searched hold (one moved to another room is not
                             one of them); a series whose every occurrence is there is kept whole.
          -Source Organizer  -SeriesScope Occurrences: the occurrences of the organizer's calendar, or of the copies
                             of the attendees and the rooms when the organizer has none (deleted mailbox).
        When the occurrences of the organizer (or of a room searched) cannot be read, the meeting is left as it is
        (Not processed): never the whole series, never the attendees without their organizer.
    #>
    param([Parameter(Mandatory = $true)][object[]]$Meetings, [Parameter(Mandatory = $true)][hashtable]$Settings, [datetime]$Start, [datetime]$End,
        [ValidateSet('Rooms', 'Organizer')][string]$Source = 'Rooms')

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
    # The request of each copy, by the copy itself (a copy object is not a key by value).
    $idOf = [Collections.Generic.Dictionary[object, string]]::new([Collections.Generic.ReferenceEqualityComparer]::Instance)
    for ($i = 0; $i -lt $series.Count; $i++) {
        $byMeeting[$i] = [Collections.Generic.List[string]]::new()
        foreach ($c in @($series[$i].Copies | Where-Object EventId)) {
            $id = "o$($requests.Count)"
            $copies[$id] = $c
            $idOf[$c] = $id
            $byMeeting[$i].Add($id)
            $requests.Add((New-MclGraphRequest -Id $id -Url "$(Get-MclUserPath $c.Mailbox)/events/$([Uri]::EscapeDataString($c.EventId))/instances?startDateTime=${S}Z&endDateTime=${E}Z&`$select=$select&`$top=200"))
        }
    }
    $res = Invoke-MclGraphBatch -Requests $requests.ToArray() -FollowPages -OnProgress { param($done, $total) Write-MclProgress ($done / [Math]::Max(1, $total)) ('{0:N0}/{1:N0} series copies: occurrences read' -f $done, $total) }

    # The occurrences that count (held by the rooms searched, or of the organizer); the meetings that cannot be
    # split safely.
    $keys = @{}; $untouched = @{}
    for ($i = 0; $i -lt $series.Count; $i++) {
        if ($series[$i].OrganizerCopy -eq 'Not read') { $untouched[$i] = "the copy of its organizer could not be read"; continue }
        $orgIds = @($byMeeting[$i] | Where-Object { $copies[$_].Role -eq 'Organizer' })
        if ($Source -eq 'Rooms') {
            $sourceIds = @($byMeeting[$i] | Where-Object { $copies[$_].Via -eq 'Room search' })
            if (-not $sourceIds.Count) { $untouched[$i] = 'no copy of the rooms searched to read its occurrences from'; continue }
            $unread = @(@($sourceIds) + @($orgIds) | Where-Object { $res[$_].Status -ne 200 })
        }
        else {
            # The organizer's calendar is the reference; without it (deleted mailbox), every copy read.
            $sourceIds = @(if ($orgIds.Count) { $orgIds } else { $byMeeting[$i] | Where-Object { $res[$_].Status -eq 200 } })
            if (-not $sourceIds.Count) { $untouched[$i] = 'no copy whose occurrences could be read'; continue }
            $unread = @($orgIds | Where-Object { $res[$_].Status -ne 200 })
        }
        if ($unread.Count) { $untouched[$i] = "occurrences of the period not read in $($copies[$unread[0]].Mailbox) ($(Get-MclMailboxProblem $res[$unread[0]]))"; continue }
        $set = [Collections.Generic.HashSet[string]]::new()
        foreach ($id in $sourceIds) { foreach ($o in @($res[$id].Values)) { if (-not [bool](Get-MclProperty $o 'isCancelled')) { [void]$set.Add((Get-MclTimeKey $o)) } } }
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
        if ($whole.Contains($i)) {
            if ($Source -eq 'Rooms') { $m.Notes.Add('Every occurrence of the series is in the period, in the rooms searched: the series is handled as a whole.'); continue }
            # Asked for by occurrence: each one acted on (Cancel sends one cancellation each); said, not changed.
            $m.Notes.Add('Every occurrence of the series is in the period: each one is acted on separately (one cancellation each with Cancel); -SeriesScope Whole acts on the series at once.')
        }
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
            $id = $null
            if (-not $idOf.TryGetValue($c, [ref]$id)) { $list.Add($c); continue }
            $r = $res[$id]
            if ($r.Status -ne 200) {
                $c.EventId = ''; $c.Result = 'Not processed'; $c.Detail = "occurrences of the period not read ($(Get-MclMailboxProblem $r)): the copy is left as it is"
                $list.Add($c); continue
            }
            # The occurrences that count only (rooms: one moved to another room stays as it is).
            $mine = @($r.Values | Where-Object { $keys[$i].Contains((Get-MclTimeKey $_)) })
            if (-not $mine.Count) {
                $c.EventId = ''; $c.Result = 'No copy'
                $c.Detail = if ($Source -eq 'Rooms') { 'no occurrence of the rooms searched in the period in this mailbox (declined, removed or moved)' } else { 'no occurrence of the period in this mailbox (declined, removed or moved)' }
                $list.Add($c); continue
            }
            foreach ($o in ($mine | Sort-Object { [MeetingCleanupNative.Fast]::ToUtc($_.start) })) {
                $occ = [MeetingCleanupNative.Fast]::NewCopy($m.MeetingId, $c.Mailbox, $c.Role, $c.Via, $o, '', '')
                if (-not $occ.Subject) { $occ.Subject = $c.Subject }
                $when = ConvertTo-MclDateUtc $o.start
                $occ.Occurrence = [MeetingCleanupNative.Fast]::FormatDate($when, $zone, $false, $false)
                $occ.OccurrenceStart = $when.ToString('o')
                $occ.OccurrenceKey = Get-MclTimeKey $o
                $occ.SeriesId = $c.EventId
                $list.Add($occ)
                $occurrences++
            }
        }
        $m.Copies.Clear()
        foreach ($c in $list) { $m.Copies.Add($c) }
        $m.Occurrences = $keys[$i].Count
        $m.Notes.Add(('{0} occurrence(s) in the period{1}: only they are acted on, the series goes on outside the period.' -f $keys[$i].Count, $(if ($Source -eq 'Rooms') { ' in the rooms searched' } else { '' })))
    }
    $wholeText = if ($whole.Count -and $Source -eq 'Rooms') { " $($script:Dot) $($whole.Count) entirely in the period, handled whole" } elseif ($whole.Count) { " $($script:Dot) $($whole.Count) entirely in the period" } else { '' }
    Write-MclItem Info ('{0} series limited to their occurrences in the period ({1:N0} occurrence copies){2}{3}' -f $split, $occurrences, $wholeText, $(if ($left) { " $($script:Dot) $left left as they are (occurrences not read)" })) -Icon Calendar
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
    $fast = [MeetingCleanupNative.Fast]

    # ---- details from the best copy --------------------------------------------------------------
    $refs = @{}
    $requests = foreach ($m in $Meetings) {
        $ref = Get-MclBestCopy $m
        $refs[$m.MeetingId] = $ref
        New-MclGraphRequest -Id $m.MeetingId -Url "$(Get-MclUserPath $ref.Mailbox)/events/$([Uri]::EscapeDataString($ref.EventId))?`$select=subject,attendees,organizer,location,start,end,type,recurrence,isCancelled,originalStartTimeZone"
    }
    $details = Invoke-MclGraphBatch -Requests @($requests) -OnProgress { param($done, $total) Write-MclProgress ($done / [Math]::Max(1, $total)) ('{0:N0}/{1:N0} meetings read' -f $done, $total) }
    foreach ($m in $Meetings) {
        $r = $details[$m.MeetingId]
        $ref = $refs[$m.MeetingId]
        if ($r.Status -ne 200) { $m.Notes.Add("Details not read from $($ref.Mailbox): $($r.Status) $($r.ErrorCode)."); continue }
        $ev = $r.Body
        $m.Subject = $fast::Text($ev, 'subject')
        $m.SubjectFromRoom = $ref.Role -eq 'Room'
        $m.Location = $fast::Text($ev, 'location', 'displayName')
        $attendees = [Collections.Generic.List[object]]::new()
        foreach ($a in @($fast::Prop($ev, 'attendees'))) {
            if ($a) { $attendees.Add([pscustomobject]@{ Address = $fast::Text($a, 'emailAddress', 'address').ToLowerInvariant(); Name = $fast::Text($a, 'emailAddress', 'name'); Type = $fast::Text($a, 'type') }) }
        }
        $m.Attendees = $attendees.ToArray()
        $orgAddressOfEvent = $fast::Text($ev, 'organizer', 'emailAddress', 'address')
        if ($orgAddressOfEvent) { $m.Organizer = $orgAddressOfEvent.ToLowerInvariant(); $m.OrganizerName = $fast::Text($ev, 'organizer', 'emailAddress', 'name') }
        $recurrence = $fast::Prop($ev, 'recurrence')
        if ($recurrence) { $m.Recurrence = Format-MclRecurrence $recurrence; $m.RecurrenceData = $recurrence }
        $m.TimeZone = $fast::Text($ev, 'originalStartTimeZone')
        if ($ref.Role -eq 'Organizer' -and -not $attendees.Count) { $m.Notes.Add('Appointment without attendees: not a meeting.') }
    }
    # An appointment of the organizer (no attendee) is not a meeting: left out.
    foreach ($m in @($Meetings)) {
        $isOrganizer = $false
        foreach ($c in $m.Copies) { if ($c.Role -eq 'Organizer') { $isOrganizer = $true; break } }
        if ($isOrganizer -and -not @($m.Attendees).Count -and $details[$m.MeetingId].Status -eq 200) { $m.Status = 'Appointment' }
    }

    # ---- copies of the attendees, by iCalUId ---------------------------------------------------------
    $lookups = [ordered]@{}
    # Meetings whose organizer copy could not be read (denied, throttled, error): not a deleted mailbox.
    $orgNotRead = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    # Per meeting: the mailboxes already listed, the items already known, and whether the organizer's copy is
    # there - kept up to date as copies are added (no new read of the copies for each attendee).
    $mailboxesOf = @{}
    $itemsOf = @{}
    $withOrganizerCopy = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($m in $Meetings) {
        $mailboxesOf[$m.MeetingId] = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $itemsOf[$m.MeetingId] = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($c in $m.Copies) {
            [void]$mailboxesOf[$m.MeetingId].Add($c.Mailbox)
            if ($c.EventId) { [void]$itemsOf[$m.MeetingId].Add($c.EventId); if ($c.Role -eq 'Organizer') { [void]$withOrganizerCopy.Add($m.MeetingId) } }
        }
    }
    $addCopy = {
        param($m, $c)
        $m.Copies.Add($c)
        [void]$mailboxesOf[$m.MeetingId].Add($c.Mailbox)
        if ($c.EventId) { [void]$itemsOf[$m.MeetingId].Add($c.EventId); if ($c.Role -eq 'Organizer') { [void]$withOrganizerCopy.Add($m.MeetingId) } }
    }
    $want = {
        param($m, [string]$Address, [string]$Type, [string]$Via)
        if (-not $Address -or $Address -notmatch '@') { return }
        if ($mailboxesOf[$m.MeetingId].Contains($Address)) { return }
        if (($orgAddress.ContainsKey($Address) -or $Address -eq $m.Organizer) -and $withOrganizerCopy.Contains($m.MeetingId)) { return }
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
                        if ($itemsOf[$m.MeetingId].Contains($fast::Text($ev, 'id'))) { continue }
                        $copyRole = if ($role -eq 'Organizer' -and -not $fast::Flag($ev, 'isOrganizer')) { 'Attendee' } else { $role }
                        & $addCopy $m $fast::NewCopy($m.MeetingId, $l.Address, $copyRole, $l.Via, $ev, '', '')
                    }
                }
                else { & $addCopy $m $fast::NewCopy($m.MeetingId, $l.Address, $role, $l.Via, $null, 'No copy', 'no copy in this mailbox (declined and removed, or never received)') }
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
            & $addCopy $m $fast::NewCopy($m.MeetingId, $l.Address, $role, $l.Via, $null, 'Not processed', $text)
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
                foreach ($l in $unknown[$address]) { & $addCopy $l.Meeting $fast::NewCopy($l.Meeting.MeetingId, $address, 'Attendee', $l.Via, $null, 'Not processed', 'not a mailbox of this tenant (external, deleted or contact)') }
                continue
            }
            $groupFor[$address] = @{ Name = [string]$grp.displayName; Id = [string]$grp.id; Meetings = $unknown[$address] }
            $memberRequests.Add((New-MclGraphRequest -Id $address -Url "/groups/$($grp.id)/transitiveMembers/microsoft.graph.user?`$select=mail&`$top=999"))
        }
        if ($memberRequests.Count) {
            $members = Invoke-MclGraphBatch -Requests $memberRequests.ToArray() -FollowPages
            foreach ($address in $groupFor.Keys) {
                $info = $groupFor[$address]
                $mails = [Collections.Generic.List[string]]::new()
                foreach ($u in $members[$address].Values) { $mail = $fast::Text($u, 'mail').ToLowerInvariant(); if ($mail) { $mails.Add($mail) } }
                foreach ($l in $info.Meetings) {
                    & $addCopy $l.Meeting $fast::NewCopy($l.Meeting.MeetingId, $address, 'Group', $l.Via, $null, 'Expanded', ("group '{0}': {1} member(s) with a mailbox" -f $info.Name, $mails.Count))
                    foreach ($mail in $mails) { & $want $l.Meeting $mail 'required' "Group $($info.Name)" }
                }
            }
            $groupFor = @{}
        }
    }

    # ---- summary per meeting (one pass over its copies) ----------------------------------------------------
    $total = 0; $roomCopies = 0; $skipped = 0; $appointments = 0; $present = 0
    foreach ($m in $Meetings) {
        $orgCopy = $false; $orgRow = $null; $best = $null; $bestRank = 9; $cancelled = $false
        foreach ($c in $m.Copies) {
            if ($c.Role -eq 'Organizer') { if ($c.EventId) { $orgCopy = $true } elseif (-not $orgRow) { $orgRow = $c } }
            if ($c.EventId) {
                $total++
                if ($c.Role -eq 'Room') { $roomCopies++ }
                elseif ($c.Subject) { $rk = $rank[$c.Role]; if ($null -eq $rk) { $rk = 3 }; if ($rk -lt $bestRank) { $bestRank = $rk; $best = $c } }
            }
            if ($c.Result -eq 'Not processed') { $skipped++ }
            if ($c.Cancelled) { $cancelled = $true }
        }
        $m.OrganizerCopy = if ($orgCopy) { 'Present' } elseif ($orgNotRead.Contains($m.MeetingId)) { 'Not read' } elseif ($orgRow -and $orgRow.Result -eq 'No copy') { 'Absent' } elseif ($orgRow) { 'Mailbox deleted' } else { 'Not checked' }
        if ($best -and ($m.SubjectFromRoom -or -not $m.Subject)) { $m.Subject = $best.Subject; $m.SubjectFromRoom = $false }
        $m.Cancelled = $cancelled
        if ($m.Status -eq 'Appointment') { $appointments++ }
        if ($m.OrganizerCopy -eq 'Present') { $present++ }
    }
    Write-MclItem Ok ('{0:N0} cop{1} of {2:N0} meeting(s) {3} {4:N0} in rooms {3} organizer copy present for {5:N0}' -f $total, $(if ($total -eq 1) { 'y' } else { 'ies' }), ($Meetings.Count - $appointments), $dot, $roomCopies, $present) -Icon People
    if ($skipped) { Write-MclItem Skip ('{0:N0} attendee(s) not processed: external, deleted, on-premises or not reachable (listed in the report)' -f $skipped) }
    if ($appointments) { Write-MclItem Skip ('{0:N0} appointment(s) of the organizer without attendees left out' -f $appointments) }
    if (-not $g.CanReadGroups) {
        $text = 'No GroupMember.Read.All: the members of a group invited to a meeting are not looked up.'
        if ($Warnings) { $Warnings.Add($text) }
        Write-MclItem Warn $text
    }
}

function Get-MclBestCopy {
    <# The copy a meeting is read from: the organizer's, else an attendee's, else a room's (one with an item ID). #>
    param([Parameter(Mandatory = $true)]$Meeting, [switch]$OrganizerAttendeeRoomOnly)
    $best = $null; $bestRank = 9
    foreach ($c in $Meeting.Copies) {
        if (-not $c.EventId) { continue }
        $rk = switch ($c.Role) { 'Organizer' { 0 } 'Attendee' { 1 } 'Room' { 2 } default { 3 } }
        if ($OrganizerAttendeeRoomOnly -and $rk -eq 3) { continue }
        if ($rk -lt $bestRank) { $bestRank = $rk; $best = $c }
    }
    return $best
}

function Update-MclResultCounts {
    <# Totals of a result: meetings, copies, mailboxes, actions (compiled, one pass over the copies). #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result)
    $Result.Counts = [MeetingCleanupNative.Fast]::Counts($Result.Meetings, @($Result.Organizers).Count)
}

# SIG # Begin signature block
# MIImIQYJKoZIhvcNAQcCoIImEjCCJg4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDAEzesTwC3ONSJ
# uRO1/djJe6u9LWYL+fhX6azY+3j9wKCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
# CDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAe
# Fw0xNjA2MjQyMDQ0MzBaFw0zMTA2MjQyMDQ0MzBaMHgxCzAJBgNVBAYTAlVTMQ4w
# DAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENv
# cnAxNDAyBgNVBAMMK1NTTC5jb20gQ29kZSBTaWduaW5nIEludGVybWVkaWF0ZSBD
# QSBSU0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCfgxNzqrDG
# bSHL24t6h3TQcdyOl3Ka5LuINLTdgAPGL0WkdJq/Hg9Q6p5tePOf+lEmqT2d0bKU
# Vz77OYkbkStW72fL5gvjDjmMxjX0jD3dJekBrBdCfVgWQNz51ShEHZVkMGE6ZPKX
# 13NMfXsjAm3zdetVPW+qLcSvvnSsXf5qtvzqXHnpD0OctVIFD+8+sbGP0EmtpuNC
# GVQ/8y8Ooct8/hP5IznaJRy4PgBKOm8yMDdkHseudQfYVdIYyQ6KvKNc8HwKp4WB
# wg6vj5lc02AlvINaaRwlE81y9eucgJvcLGfE3ckJmNVz68Qho+Uyjj4vUpjGYDdk
# jLJvSlRyGMwnh/rNdaJjIUy1PWT9K6abVa8mTGC0uVz+q0O9rdATZlAfC9KJpv/X
# gAbxwxECMzNhF/dWH44vO2jnFfF3VkopngPawismYTJboFblSSmNNqf1x1KiVgMg
# Lzh4gL32Bq5BNMuURb2bx4kYHwu6/6muakCZE93vUN8BuvIE1tAx3zQ4XldbyDge
# VtSsSKbt//m4wTvtwiS+RGCnd83VPZhZtEPqqmB9zcLlL/Hr9dQg1Zc0bl0EawUR
# 0tOSjAknRO1PNTFGfnQZBWLsiePqI3CY5NEv1IoTGEaTZeVYc9NMPSd6Ij/D+KNV
# t/nmh4LsRR7Fbjp8sU65q2j3m2PVkUG8qQIDAQABo4H7MIH4MA8GA1UdEwEB/wQF
# MAMBAf8wHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKSle44gCUNplkwMAYIKwYBBQUH
# AQEEJDAiMCAGCCsGAQUFBzABhhRodHRwOi8vb2NzcHMuc3NsLmNvbTARBgNVHSAE
# CjAIMAYGBFUdIAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwOwYDVR0fBDQwMjAwoC6g
# LIYqaHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0G
# A1UdDgQWBBRUwv4QlQCTzWr158DX2bJLuI8M4zAOBgNVHQ8BAf8EBAMCAYYwDQYJ
# KoZIhvcNAQELBQADggIBAPUPJodwr5miyvXWyfCNZj05gtOII9iCv49UhCe204MH
# 154niU2EjlTRIO5gQ9tXQjzHsJX2vszqoz2OTwbGK1mGf+tzG8rlQCbgPW/M9r1x
# xs19DiBAOdYF0q+UCL9/wlG3K7V7gyHwY9rlnOFpLnUdTsthHvWlM98CnRXZ7WmT
# V7pGRS6AvGW+5xI+3kf/kJwQrfZWsqTU+tb8LryXIbN2g9KR+gZQ0bGAKID+260P
# Z+34fdzZcFt6umi1s0pmF4/n8OdX3Wn+vF7h1YyfE7uVmhX7eSuF1W0+Z0duGwdc
# +1RFDxYRLhHDsLy1bhwzV5Qe/kI0Ro4xUE7bM1eV+jjk5hLbq1guRbfZIsr0WkdJ
# LCjoT4xCPGRo6eZDrBmRqccTgl/8cQo3t51Qezxd96JSgjXktefTCm9r/o35pNfV
# HUvnfWII+NnXrJlJ27WEQRQu9i5gl1NLmv7xiHp0up516eDap8nMLDt7TAp4z5T3
# NmC2gzyKVMtODWgqlBF1JhTqIDfM63kXdlV4cW3iSTgzN9vkbFnHI2LmvM4uVEv9
# XgMqyN0eS3FE0HU+MWJliymm7STheh2ENH+kF3y0rH0/NVjLw78a3Z9UVm1F5VPz
# iIorMaPKPlDRADTsJwjDZ8Zc6Gi/zy4WZbg8Zv87spWrmo2dzJTw7XhQf+xkR6Od
# MIIG9TCCBN2gAwIBAgIQBof4ljXihJpmCqqt2mhAWzANBgkqhkiG9w0BAQsFADB4
# MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24x
# ETAPBgNVBAoMCFNTTCBDb3JwMTQwMgYDVQQDDCtTU0wuY29tIENvZGUgU2lnbmlu
# ZyBJbnRlcm1lZGlhdGUgQ0EgUlNBIFIxMB4XDTI2MTAxMzAyNDgwMloXDTI3MTAx
# MzAyNDgwMlowXDELMAkGA1UEBhMCRlIxGzAZBgNVBAcMEkNoYW50ZWxvdXAtRW4t
# QnJpZTEXMBUGA1UECgwOTmljb2xhcyBGYWJlcnQxFzAVBgNVBAMMDk5pY29sYXMg
# RmFiZXJ0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsg6CT0mFQ4jB
# V2plIS4uWFSoVqRHd82Y7XYZ4CUTe/7nPBRUlXXdKTZMMHzGgiG8dEXNE8K1Xe90
# Dgl/FnorzAFlrNWSwf6t27UAUw9xkqtRDVVpVgpixpWkAwjcQeLG+YMqBkGbFi35
# TGvS23lPFWvLp4DxARrv15MtZ/NjNTCAAGUlMP2hFVBknEmrnYjEGyeJ93sl20RS
# WXQQLNqlenBdkGCrQW4XjHtRjJAEqAC68hN6nWevopSQ8CLUV7A1RFfHFarDrzrL
# zr0/NjIdSyIfM9hShln31NDrktB6Ups2XIyrtPDNB7fjkW4jpLlOqQGfYKsU9ErE
# 4SnFsh0yOOjitegyh1RvZ9wH++3DF4CmJfTli5l/gEIm8xCCryxI+K5EogW3wu1y
# u8NT2hcvHjdh7OpRyPyFAuLZ8MfX7kPAJpnuCeOuktzY/20eVD8bDglj+4zgqzsn
# cR/t6T/7w1/HxpRyrVoKvEyRx+mphYh5jMdEa0BjxOL0vAJpdGZSfXGmPm/yPqRn
# 5STTkMtzkVQXPbcvgKM0luzFcQKPHk3qdHLEXa8gf93ktYDbLE5Wd2UHrXz3SZxN
# URHPdIXCr1qgrqQve7041kMpzrkDI5DnG2lvRG4108gp02bHG0muKyZlDdCOra7h
# Zonu2nO79sFallruzzNi0o994PoWksUCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQC
# MAAwHwYDVR0jBBgwFoAUVML+EJUAk81q9efA19myS7iPDOMwegYIKwYBBQUHAQEE
# bjBsMEgGCCsGAQUFBzAChjxodHRwOi8vY2VydC5zc2wuY29tL1NTTGNvbS1TdWJD
# QS1Db2RlU2lnbmluZy1SU0EtNDA5Ni1SMS5jZXIwIAYIKwYBBQUHMAGGFGh0dHA6
# Ly9vY3Nwcy5zc2wuY29tMFEGA1UdIARKMEgwCAYGZ4EMAQQBMDwGDCsGAQQBgqkw
# AQMDATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3JlcG9zaXRv
# cnkwEwYDVR0lBAwwCgYIKwYBBQUHAwMwTQYDVR0fBEYwRDBCoECgPoY8aHR0cDov
# L2NybHMuc3NsLmNvbS9TU0xjb20tU3ViQ0EtQ29kZVNpZ25pbmctUlNBLTQwOTYt
# UjEuY3JsMB0GA1UdDgQWBBRoUx3630l9wA1fUpkKkw2GUIxl/jAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE3oj6cNuwlDToi1BwjgsA5djfRN4DIi
# BYDFZIRYaIRe/rEdnItz3aqQQdFuwXnJA/MQEmi2WAGek8GeZF+R41s8FbcDxBUE
# BoSt98TuBE+0xgxQWDiPhhXaX1Aw9V0GF4ZnlfCSw1eu2KiASaKn7Bo0PLbEAXff
# vftRXof/kyqtw8iID6I4Rhrf5DI3WVZQodakUUmWnbnziWvEZOWbNOfzwP+7BwEt
# 7jQANfbmEfb4p75rfy5dQVwniTDSvdo6xkrbe1Cr93GS7cJta18IG5CoIVUmdK+s
# W38B1sLuy8Mi0MstMXV3004yaRuzLdNaAM0f9WAFerxemiQC0kdmA44/lWV3lVkM
# WJSwkHBkkSg8Eh0hxcGzTH2xmtDhLJHNNLRwkI9u2Ur+HbLCt50oBgSV8BcCs/1i
# qi884iC/yuX+HIv9YgTUNZ8c4lLaq4VuliJpXLu2buHIsTkO0v0+5oCM91zjTVZm
# ULeH38N0Vr83psiJVuBSxQyt5nRVrpYz7FACq5gtXELkyViq/rGOlJKaRrgQ6WWq
# eoIXqKhVFAdP/jAEm2M8ezoVYsUgPyKY2OcuecV9coulAcxUIUydx6qYiOfsGj5u
# Dy7l4aYjvB771wkUMzUQwxzcDYXJtMMTaWRe6n9EemR67jc73u5CJyFg3PfA5WR1
# wZislcdCZuS6MYIYCDCCGAQCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDvb3KkBnMB
# R/4noZnH8KbOsYYsYx8W3OD1QEU3hVVmMTANBgkqhkiG9w0BAQEFAASCAgACPBcT
# 4z4/SOkboII/t8G7xxorLj3uIoDOLurPHTA4CaIfLSjOxgtfhLeXqFTBZZIw4GOz
# xaaWKy8wEZWdU90YsVDi0vRh9QpS+iOVjJmwnypdVXnlp308mghgP8/6Q8A1GgSK
# 03RsgwhIcb2QLAeScdZbWrCdxm4+PdjY3VFtt0uZxyCCw1QdJgK1+Si0h0uo7XRK
# yuQ+Is+l4U/FAOd3gaRUw8uh8GlhIPxfKR3oYtRM6OVwrrLI6MfRG1XprYDmE+/1
# l4rDUyRi9fgDXPCMMoghnVO3t/jsJHgUb1tCcujutseXPriKFKqjwJtkGijea5SB
# 674m5TJmJ85+IELpDgVxFrNgiaCDHbvx/IEOUMrXuDMvTaj/gGlOhGJHfC8kNAxi
# 6YgAWKdJq1XAHjn3n+xig1rjCd41IMrtD3O9V3VFoAynof4wDcNHTB3tm+M1Dyd6
# S7ziisEyxhAmSBVrPv7+/NshodzWWPxGr05Et6yXTuehUPGyJJPSARL0zSzdmVAA
# 0LSp2pvh9/v273iyCXJMZwmy05n7+hPTeK6xgbExzJXj7xHOQCF4EM58h4bmbSgW
# 4GcpiSdezErWA3opj2FI5rVM62GJURu+h22pvzvWzu/VW9MsqD7GO16we5894c4u
# h1X1DkhLetyzQf3y+BORx/IPmg4YK0LwPxOipqGCFMUwghTBBgorBgEEAYI3AwMB
# MYIUsTCCFK0GCSqGSIb3DQEHAqCCFJ4wghSaAgEDMQ0wCwYJYIZIAWUDBAIBMIGB
# BgsqhkiG9w0BCRABBKByBHAwbgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEICzGUQ/TygYdi46iJHCpSPrcC4ZamZQea5xdJGOXCk3VAgglL2pxiJN7
# VRgPMjAyNjExMDYwODQyMjVaMAMCAQECCEsOGwg96j9aoIIRpzCCBPMwggLboAMC
# AQICEFRD+De9Wu4YAYRyuFzajuMwDQYJKoZIhvcNAQELBQAwcjELMAkGA1UEBhMC
# VVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhT
# U0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNzdWluZyBS
# U0EgQ0EgMjAeFw0yNjA2MTIxNzIyNDZaFw0zNjA2MDkxNzIyNDVaMG4xCzAJBgNV
# BAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UE
# CgwIU1NMIENvcnAxKjAoBgNVBAMMIVNTTC5jb20gVGltZXN0YW1waW5nIFVuaXQg
# MjAyNiBFMTBZMBMGByqGSM49AgEGCCqGSM49AwEHA0IABOgli7hmEy5IkjjzCzbJ
# DgbYV7CIUUgOOn3jkML0i3A2jYer8VMwShMdSnRQURbOzh65x/ytFQOjGw0CBaVI
# ND2jggFSMIIBTjAfBgNVHSMEGDAWgBRwmmXBQ4Cj/OmxZIlu8TCcCSHJfDBNBggr
# BgEFBQcBAQRBMD8wPQYIKwYBBQUHMAKGMWh0dHA6Ly9jZXJ0LnNzbC5jb20vU1NM
# LmNvbS10aW1lU3RhbXBpbmctSS1SMi5jZXIwUQYDVR0gBEowSDA8BgwrBgEEAYKp
# MAEDBgEwLDAqBggrBgEFBQcCARYeaHR0cHM6Ly93d3cuc3NsLmNvbS9yZXBvc2l0
# b3J5MAgGBmeBDAEEAjAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDBCBgNVHR8EOzA5
# MDegNaAzhjFodHRwOi8vY3Jscy5zc2wuY29tL1NTTC5jb20tdGltZVN0YW1waW5n
# LUktUjIuY3JsMB0GA1UdDgQWBBQnrANWuihquhc7gz1Af3moJOzDPDAOBgNVHQ8B
# Af8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE92bs78A57CGU1wZ/oY5rgV4Mwd
# dn22MU/bmHumb4b6W2VinMg0IGnBE5D8PJOsnqp2AmMhWdOt5TDxKgIQA9pzBute
# GYez7R1vQ1qn7IIr8Wz2AXjfre1MYJGE93saW1jHTxKgbVlchzry3CrzMYCVnyYw
# DcFZMFQ35Mv4+3H9vPoopjfDynO0v7vPy3p0UZ1dltmH53EZ5M/Qn7OZ0Q82mYLo
# h666dhscIJzUEh+BgYxzpdz8y7J1WxqGSouGRI4F7qbOJhhjiu/hgdhwJUJ/jkve
# OHB6pZB9IBJbdlns6AjUaRC6YfamsbgZ40hKl1ln4BOshgcDtLsSjaqVlUx5gMK+
# krwT6H8bm1yLhJ/Yr0eAizrgMVEqyeNWQYP6qhtKCYxehpESm8au3MlkpBGWCYZi
# goNh4OX1x2oDsYsG5p+ST2db4Vgq25N7/3U8doL0yufnqwuR/f603JY1CO++sndq
# 34Y35Kb8gLG7V7OA+5658PIvcKK+fn3bUUiQVcasy5spBX7aC9FKdkPvrciglHa4
# m3XJ6weBykt7PrdwI3XGwNOgwi/Bon0IvXy9MkQUTZQOKOgT/kq++eTGgjY1EWGQ
# 7w8sSRwQU+3Ko2SMLIkfLPuvOtJ6BWwQbXYHtiMOwjwv9UN9SNu4VaJUuwdMmqN7
# /F9jLVst5rGZV03BMIIF3TCCA8WgAwIBAgIIeyyb0xaAMpkwDQYJKoZIhvcNAQEL
# BQAwfDELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3Vz
# dG9uMRgwFgYDVQQKDA9TU0wgQ29ycG9yYXRpb24xMTAvBgNVBAMMKFNTTC5jb20g
# Um9vdCBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eSBSU0EwHhcNMTYwMjEyMTczOTM5
# WhcNNDEwMjEyMTczOTM5WjB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAPkP3aMrfcvQKv7sZ4Wm5y4b
# unfh4/WvpOz6Sl2RxFdHaxh3a3by/ZPkPQ/CFp4LZsNWlJ4Xg4XOVu/yFv0AYvUi
# CVToZRdOQbngT0aXqhvIuG5iXmmxX9sqAn78bMrzQdjt0Oj8P2FI7bADFB0QDksZ
# 4LtO7IZl/zbzXmcCC52GVWH9ejjt/uIZALdvoVBidXQ8oPrIJZK0bnoix/geoeOy
# 3ZExqysdBP+lSgQ36YWkMyv94tZVNHwZpEpox7Ko07fKoZOI68GXvIz5HdkihCR0
# xwQ9aqkpk8zruFvh/l8lqjRYyMEjVJ0bmBHDOJx+PYZspQ9AhnwC9FwCTyjLrnGf
# DzrIM/4RJTXq/LrFYD3ZfBjVsqnTdXgDciLKOsMf7yzlLqn6niy2UUb9rwPW6mBo
# 6oUWNmuF6R7As93EJNyAKoFBbZQ+yODJgUEAnl6/f8UImKIYLEJAs/lvOCdLToD0
# PYFH4Ih86hzOtXVcUS4cK38acijnALXRdMbX5J+tB5O2UzU1/Dfkw/ZdFr4hc96S
# CvigY2q8lpJqPvi8ZVWb3vUNiSYE/CUapiVpy8JtynziWV+XrOvvLsi81xtZPCvM
# 8hnIk2snYxnP/Okm+Mpxm3+T/jRnhE6Z6/yzeAkzcLpmpnbtG3PrGqUNxCITIJRW
# Ck4sbE6x/c+cCbqiM+2HAgMBAAGjYzBhMB0GA1UdDgQWBBTdBAkHovV6fVJTEpKV
# 7jiAJQ2mWTAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFN0ECQei9Xp9UlMS
# kpXuOIAlDaZZMA4GA1UdDwEB/wQEAwIBhjANBgkqhkiG9w0BAQsFAAOCAgEAIBgR
# lCn7Jp0cHh5wYfGVcpNxJK1ok1iOMq8bs3AD/CUrdIWQPXhq9LmLpZc7tRiRux6n
# +UBbkflVma8eEdBcHadm47GUBwwyOabqG7B52B2ccETjit3E+ZUfijhDPwGFpUen
# PUayvOUiaPd7nNgsPgohyC0zrL/FgZkxdMF1ccW+sfAjRfSda/wZY52jvATGGAsl
# u1OJD7OAUN5F7kR/q5R4ZJjT9ijdh9hwZXT7DrkT66cPYakylszeu+1jTBi7qUD3
# oFRuIIhxdRjqerQ0cuAjJ3dctpDqhiVAq+8zD8ufgr6iIPv2tS0a5sKFsXQP+8hl
# AqRSAUfdSSLBv9jra6x+3uxjMxW3IwiPxg+NQVrdjsW5j+VFP3jbutIbQLH+cU0/
# 4IGiul607BXgk90IH37hVZkLId6Tngr75qNJvTYw/ud3sqB1l7UtgYgXZSD32pAA
# n8lSzDLKNXz1PQ/YK9f1JmzJBjSWFupwWRoyeXkLtoh/D1JIPb9s2KJELtFOt3JY
# 04kTlf5Eq/jXixtunLwsoFvVagCvXzfh1foQC5ichucmj87w7G6KVwuA406ywKBj
# YZC6VWg3dGq2ktufoYYitmUnDuy2n0Jg5GfCtdpBC8TTi2EbvPofkSvXRAdeuims
# 2cXp71NIWuuA8ShYIc2wBlX7Jz9TkHCpBB5XJ7kwggbLMIIEs6ADAgECAhBSnbvC
# uImH9YOA4LDrogY2MA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMQ4wDAYD
# VQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBv
# cmF0aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRo
# b3JpdHkgUlNBMB4XDTI2MDYxMjE2NDIxNVoXDTQxMDIxMjE3MzkzOVowcjELMAkG
# A1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYD
# VQQKDAhTU0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNz
# dWluZyBSU0EgQ0EgMjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAJ8v
# q+ZzTZC8MUabNgd5WlBjVPcu71oiedNd8n9Il8+uW0HcGvW44/Jf0pHl6hxqn0od
# haGnvvNBPJgxH/d6CMI7Mtgl9nF5isarRfByRmAeH+skYShA+dMJGzsrC9NiF3/v
# xGZ2uOI334YVc+DlP5OdsUPz5yIP7hYXNlpPknPwh6gK/pWFZD8+y8uUagfX9USh
# q/SQt+1tS4FXRwudHTSHbGt0Fz4ZbWGbm0d56ROAW34tI6FFuZ+TtGssHjeOUf2F
# S8QNGBpadTh1X1Z7Kl6vG4Z8NS3OJIyelzoF5TpcP2vM3ue0723bmZd3oXsE5rzt
# AJMv56jlg0vSDz8XcXRur6awW0rh5cTQDQQGx04hYJH2Zn9HWWJU/wJ2SKBcQKtB
# yCz6JX+DoFoBE++kolkpZtl2ESF8FY6XqhtKZb18ESEAc0ohCY+0HkBZwrx8eIED
# CDbzd/shv9a8YRG2oxr3tLUz95V+Dzzh1QBNECaeGczyNWGCZIVUa4lVcVUVOBbI
# OAT8D+iPmC5h29bWeES8DhTS4f4CWA4D5QoQqMTqIBcV4MD2Waw8hEWAdGnjnsVt
# DimnDeEUbvXKhdiirvwG+QC+3mAy1RlD6xBMmGyDAKVSqqXm3NP4w0J4gvt65Grk
# QI4FXPJtYqfwkIBtQ8OGfbpVUQGPVUemWtui7E21AgMBAAGjggFRMIIBTTASBgNV
# HRMBAf8ECDAGAQH/AgEAMB8GA1UdIwQYMBaAFN0ECQei9Xp9UlMSkpXuOIAlDaZZ
# ME0GCCsGAQUFBwEBBEEwPzA9BggrBgEFBQcwAoYxaHR0cDovL2NlcnQuc3NsLmNv
# bS9TU0xjb20tUm9vdENBLVJTQS00MDk2LVIxLmNlcjA/BgNVHSAEODA2MDQGBFUd
# IAAwLDAqBggrBgEFBQcCARYeaHR0cHM6Ly93d3cuc3NsLmNvbS9yZXBvc2l0b3J5
# MBMGA1UdJQQMMAoGCCsGAQUFBwMIMEIGA1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9j
# cmxzLnNzbC5jb20vU1NMY29tLVJvb3RDQS1SU0EtNDA5Ni1SMS5jcmwwHQYDVR0O
# BBYEFHCaZcFDgKP86bFkiW7xMJwJIcl8MA4GA1UdDwEB/wQEAwIBhjANBgkqhkiG
# 9w0BAQsFAAOCAgEAZRRbzTHpgOogYgWKOeiCkI1oVrkGmCrqOKqX9CiTREQLdYM/
# h7ANki98QP1fhAPEkTzT/Ur19IrPuZg7KaWL6SsmBqmD9+mBVtJmxdRdGbV8bpgl
# FDv2VeZgTSg3+UOD2O0G9HlgHqVHNKzSa/zL5+WYl3yT8Q29sZnIPGkQ+udv56/j
# q5NlAT9BON4iXSOCxe1bNvUd0hkS2GTsTNeQyOEHm7CkZdCg5X+DpmBginO7pk8Y
# m5GbTdQtSloJ3/Y2nVPycJ8kmmBAYMev7/hVQukhjjc4mO03LyPTFfqmMR1D5Tpu
# gGkD9nUuG41l55vOpXQm1JMbTgArpkn0/e9O+yTYlWuz2dSAZHf9jxNQAvILO1Pv
# 2JNQTfV0TyIteZP7gNcRfaHnnuSho8ohipKiE+1vEJIFE+c3gopxSpHCmn2THETk
# lWKV7ZrUKhhkdoOLj2KidCtYZrOoCV6x7f9+KfDPT61fKDVntGZdRKwzLV53VIN6
# oSS+KQN2VJaO8OZTOZpfgO4o6/9xpNr+EEaA3hNkoZk13R4LaqPC/309dfZmHBE8
# nBNjWBkV7AZdzdK8w7B0a/D9066a6oUa19OIktD+3kPpQ8YEt8oRp9zs1hoQ7fRn
# 9sdantP6nYRQFTpF8czGc+FGresbAH819Pv5vBhCluTo8Px2PpC71cM0jzsxggJV
# MIICUQIBATCBhjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29t
# IFRpbWVzdGFtcGluZyBJc3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7j
# MAsGCWCGSAFlAwQCAaCCAWAwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMBwG
# CSqGSIb3DQEJBTEPFw0yNjExMDYwODQyMjVaMCgGCSqGSIb3DQEJNDEbMBkwCwYJ
# YIZIAWUDBAIBoQoGCCqGSM49BAMCMC8GCSqGSIb3DQEJBDEiBCC4GXxq7TGa+vo+
# xqNmL/HyTyzMzkhIXnqaZIS1JzC1HTCByAYLKoZIhvcNAQkQAi8xgbgwgbUwgbIw
# ga8EIFdP6heq+l58aiubCMmYsXNeRQs4c8pE5qJMoJMOY/OwMIGKMHakdDByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7jMAoGCCqGSM49BAMCBEYw
# RAIgRGyBXxXGMAs80BAIXS0hUeQyMHA1TQxAsTa2peHyTTUCIAa26FZWsgZGFSLL
# xIKwgE9ZWfD64/nE4CI82UcEIRfV
# SIG # End signature block
