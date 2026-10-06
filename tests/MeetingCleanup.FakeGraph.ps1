<#
    Meeting Cleanup - simulated Exchange Online tenant behind Microsoft Graph, for the tests.
    Author  : Nicolas Fabert
    Version : 1.2.0

    Start-MclGraphSend (the only function of the tool that touches the network for Graph) is replaced by a
    mock that answers from this tenant: users, groups, rooms (places), calendars with meetings and their
    copies, $batch. It reproduces the behaviour measured in the lab:
      - the same iCalUId in every copy of a meeting; a room copy has the organizer's name as subject;
      - a copy holds the whole attendee list, the organizer included;
      - $filter on organizer is refused (501); the filters of the tool are evaluated;
      - permanentDelete of a copy of an attendee or a room is silent; in the organizer's calendar it sends a
        cancellation (recorded in $script:Fake.Messages); cancel sends one too, and the rooms remove their copy;
      - 404 ErrorInvalidUser for an address that is not a mailbox, MailboxNotEnabledForRESTAPI for a mailbox
        that cannot be opened, ErrorItemNotFound for an item that is gone;
      - permanentDelete of an attendee's or a room's copy moves it to Purges (new EntryID, LastModifiedUtc = now),
        marks it Declined / Free and the organizer's tracking Declined, silently; Get-FakeRecoverableItems and
        Restore-FakeRecoverableItem play Get-RecoverableItems and Restore-RecoverableItems; accept and
        tentativelyAccept with sendResponse false answer silently.
      - 1.2: each occurrence of a series has its own ID in each copy (instances); an occurrence removed in an
        attendee's or a room's copy is silent and NOT kept in Recoverable Items; cancelled by the organizer, it
        is removed from the rooms and marked cancelled at the attendees; a meeting created by POST has no
        attendee (no message) until a PATCH of its attendees sends the invitation (copies delivered, rooms
        accept); an occurrence of the organizer can be removed or moved (exception) before that.
#>

function Reset-FakeTenant {
    $script:Fake = @{
        Mailboxes = @{}            # address -> @{ Address; Name; Kind (User|Room); Events (List); Aliases; Reachable }
        Users     = [Collections.Generic.List[object]]::new()
        Groups    = [Collections.Generic.List[object]]::new()
        Messages  = [Collections.Generic.List[object]]::new()
        Calls     = [Collections.Generic.List[object]]::new()   # every request (method, url)
        Created   = [Collections.Generic.List[object]]::new()   # events created by POST (transfer)
        Batches   = [Collections.Generic.List[object]]::new()   # every $batch: the mailbox of each sub-request
        Throttle  = @{}            # url pattern -> number of 429 still to answer
        Fail      = @{}            # url pattern -> @{ Status; Code }
        Next      = 0
        Purged    = 0
    }
}

function New-FakeId { param([string]$Prefix = 'AAMk') $script:Fake.Next++; return '{0}{1:D8}=' -f $Prefix, $script:Fake.Next }

function Add-FakeMailbox {
    param([Parameter(Mandatory)][string]$Address, [string]$Name, [ValidateSet('User', 'Room')][string]$Kind = 'User', [string[]]$Aliases = @(), [switch]$InDirectoryOnly, [switch]$Unreachable, [switch]$NoDirectory)
    $a = $Address.ToLowerInvariant()
    if (-not $InDirectoryOnly) {
        $script:Fake.Mailboxes[$a] = @{ Address = $a; Name = $(if ($Name) { $Name } else { $a }); Kind = $Kind; Events = [Collections.Generic.List[object]]::new(); Purges = [Collections.Generic.List[object]]::new(); Aliases = @($Aliases | ForEach-Object { $_.ToLowerInvariant() }); Reachable = -not $Unreachable }
        foreach ($alias in $Aliases) { $script:Fake.Mailboxes[$alias.ToLowerInvariant()] = $script:Fake.Mailboxes[$a] }
    }
    # -NoDirectory: the user was deleted, Graph still opens the soft-deleted mailbox.
    if ($NoDirectory) { return }
    $script:Fake.Users.Add([pscustomobject]@{ id = [guid]::NewGuid().ToString(); displayName = $(if ($Name) { $Name } else { $a }); mail = $a; userPrincipalName = $a; proxyAddresses = @("SMTP:$a") + @($Aliases | ForEach-Object { "smtp:$_" }); Kind = $Kind })
}

function Add-FakeGroup {
    param([Parameter(Mandatory)][string]$Address, [string]$Name = 'Group', [string[]]$Members = @())
    $script:Fake.Groups.Add([pscustomobject]@{ id = [guid]::NewGuid().ToString(); displayName = $Name; mail = $Address.ToLowerInvariant(); Members = @($Members | ForEach-Object { $_.ToLowerInvariant() }) })
}

function New-FakeICalUId { ('040000008200E00074C5B7101A82E00800000000' + (-join (1..48 | ForEach-Object { '{0:X}' -f (Get-Random -Maximum 16) }))) }

function Add-FakeMeeting {
    <#
        A meeting sent by an organizer: a copy in the organizer's calendar (unless -NoOrganizerCopy or the
        organizer has no mailbox) and in each attendee mailbox (unless listed in -NoCopy).
        Attendees: addresses; rooms: -Rooms. -Recurrence: Graph patternedRecurrence (hashtable).
    #>
    param(
        [Parameter(Mandatory)][string]$Organizer, [string]$OrganizerName = 'Organizer', [Parameter(Mandatory)][string]$Subject,
        [Parameter(Mandatory)][datetime]$Start, [int]$Minutes = 30, [string[]]$Attendees = @(), [string[]]$Rooms = @(),
        [hashtable]$Recurrence, [string[]]$NoCopy = @(), [switch]$NoOrganizerCopy, [string]$OrganizerAddressInCopies, [switch]$Cancelled
    )
    $uid = New-FakeICalUId
    $orgShown = if ($OrganizerAddressInCopies) { $OrganizerAddressInCopies } else { $Organizer.ToLowerInvariant() }
    $list = @(@($Attendees | ForEach-Object { @{ emailAddress = @{ address = $_; name = $_ }; type = 'required'; status = @{ response = 'none' } } }) +
        @($Rooms | ForEach-Object { @{ emailAddress = @{ address = $_; name = "Room $_" }; type = 'resource'; status = @{ response = 'accepted' } } }))
    $utc = [datetime]::SpecifyKind($Start, [DateTimeKind]::Utc)
    $make = {
        param([string]$Mailbox, [bool]$IsOrganizer, [bool]$IsRoom)
        $attendees = if ($IsOrganizer) { $list } else { @(@{ emailAddress = @{ address = $orgShown; name = $OrganizerName }; type = 'required'; status = @{ response = 'none' } }) + $list }
        [pscustomobject]@{
            id = (New-FakeId); iCalUId = $uid; subject = $(if ($IsRoom) { $OrganizerName } elseif ($Cancelled -and -not $IsOrganizer) { "Canceled: $Subject" } else { $Subject })
            type = $(if ($Recurrence) { 'seriesMaster' } else { 'singleInstance' })
            organizer = @{ emailAddress = @{ address = $orgShown; name = $OrganizerName } }; isOrganizer = $IsOrganizer
            start = @{ dateTime = $utc.ToString('yyyy-MM-ddTHH:mm:ss.0000000'); timeZone = 'UTC' }; end = @{ dateTime = $utc.AddMinutes($Minutes).ToString('yyyy-MM-ddTHH:mm:ss.0000000'); timeZone = 'UTC' }
            isCancelled = [bool]($Cancelled -and -not $IsOrganizer); recurrence = $Recurrence; responseStatus = @{ response = $(if ($IsOrganizer) { 'organizer' } elseif ($IsRoom) { 'accepted' } else { 'notResponded' }) }; showAs = $(if ($IsOrganizer -or $IsRoom) { 'busy' } else { 'tentative' })
            attendees = $attendees; location = @{ displayName = $(if ($Rooms) { (@($Rooms | ForEach-Object { $_.Split('@')[0] }) -join '; ') } else { 'Microsoft Teams Meeting' }) }; Mailbox = $Mailbox
            body = @{ contentType = 'html'; content = "<p>$Subject</p>" }; importance = 'normal'; sensitivity = 'normal'; isAllDay = $false; isOnlineMeeting = $false; allowNewTimeProposals = $true; responseRequested = $true
            DeletedOccurrences = [Collections.Generic.List[string]]::new(); CancelledOccurrences = [Collections.Generic.List[string]]::new(); Exceptions = @{}
        }
    }
    $org = $script:Fake.Mailboxes[$Organizer.ToLowerInvariant()]
    if ($org -and -not $NoOrganizerCopy) { $org.Events.Add((& $make $org.Address $true $false)) }
    foreach ($a in @($Attendees) + @($Rooms)) {
        if ($NoCopy -contains $a) { continue }
        # Exchange delivers the invitation of a group to its members.
        $grp = $script:Fake.Groups | Where-Object mail -eq $a.ToLowerInvariant() | Select-Object -First 1
        $targets = if ($grp) { @($grp.Members) } else { @($a) }
        foreach ($t in $targets) {
            if ($NoCopy -contains $t) { continue }
            $mb = $script:Fake.Mailboxes[$t.ToLowerInvariant()]
            if ($mb) { $mb.Events.Add((& $make $mb.Address $false ($mb.Kind -eq 'Room'))) }
        }
    }
    return $uid
}

function Get-FakeOccurrences {
    <# Start of each occurrence of a series (daily or weekly on one day, numbered / endDate / noEnd). #>
    param($Event, [datetime]$From, [datetime]$To)
    $r = $Event.recurrence
    $first = [datetime]::Parse($Event.start.dateTime, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal)
    $step = if ($r.pattern.type -eq 'weekly') { 7 * [int]$r.pattern.interval } else { [int]$r.pattern.interval }
    $max = if ($r.range.type -eq 'numbered') { [int]$r.range.numberOfOccurrences } else { 2000 }
    $last = if ($r.range.type -eq 'endDate') { [datetime]::ParseExact($r.range.endDate, 'yyyy-MM-dd', $null).AddDays(1) } else { [datetime]::MaxValue }
    $d = $first
    for ($i = 0; $i -lt $max -and $d -lt $last -and $d -lt $To; $i++) {
        if ($d -ge $From) { $d }
        $d = $d.AddDays($step)
    }
}

function ConvertTo-FakeJson { param($Value) if ($null -eq $Value) { return '' } return ($Value | ConvertTo-Json -Depth 12 -Compress) }

function Get-FakeQuery {
    param([string]$Query)
    $q = @{}
    foreach ($pair in ($Query.TrimStart('?') -split '&')) {
        if (-not $pair) { continue }
        $k, $v = $pair -split '=', 2
        $q[[Uri]::UnescapeDataString($k)] = [Uri]::UnescapeDataString([string]$v)
    }
    return $q
}

function Select-FakeEvent {
    param($Event)
    $copy = $Event | Select-Object * -ExcludeProperty Mailbox, DeletedOccurrences, CancelledOccurrences, Exceptions
    return $copy
}

function Get-FakeMailboxList {
    <# Every mailbox once (an alias is a second key to the same mailbox). Select-Object -Unique would compare hashtables as equal. #>
    $seen = @{}
    foreach ($mb in $script:Fake.Mailboxes.Values) { if (-not $seen.ContainsKey($mb.Address)) { $seen[$mb.Address] = $true; $mb } }
}

function Get-FakeOccurrenceKey { param([datetime]$Start) $Start.ToString('yyyy-MM-ddTHH:mm', [Globalization.CultureInfo]::InvariantCulture) }

function New-FakeInstance {
    <# One occurrence of a series copy, as /instances gives it ($null when it was removed from that copy). #>
    param($Master, [datetime]$Slot)
    $key = Get-FakeOccurrenceKey $Slot
    if ($Master.DeletedOccurrences -contains $key) { return $null }
    $s0 = [datetime]::Parse($Master.start.dateTime, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal)
    $e0 = [datetime]::Parse($Master.end.dateTime, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal)
    $x = $Master.Exceptions[$key]
    $start = if ($x) { $x.Start } else { $Slot }
    $end = if ($x) { $x.End } else { $Slot.Add($e0 - $s0) }
    [pscustomobject]@{
        id = "$($Master.id)~$key"; iCalUId = $Master.iCalUId; subject = $(if ($x -and $x.Subject) { $x.Subject } else { $Master.subject }); type = $(if ($x) { 'exception' } else { 'occurrence' })
        seriesMasterId = $Master.id; originalStart = $Slot.ToString('yyyy-MM-ddTHH:mm:ssZ')
        start = @{ dateTime = $start.ToString('yyyy-MM-ddTHH:mm:ss.0000000'); timeZone = 'UTC' }; end = @{ dateTime = $end.ToString('yyyy-MM-ddTHH:mm:ss.0000000'); timeZone = 'UTC' }
        showAs = $Master.showAs; responseStatus = $Master.responseStatus; isCancelled = [bool]($Master.CancelledOccurrences -contains $key); isOrganizer = $Master.isOrganizer
        organizer = $Master.organizer; location = $(if ($x -and $x.Location) { @{ displayName = $x.Location } } else { $Master.location })
    }
}

function Invoke-FakeChangeMeetingOrganizer {
    <# Invoke-ChangeMeetingOrganizer as documented: the meeting moves to the new organizer, attendees updated silently. #>
    param([string]$Mailbox, [string]$EventId, [string]$NewOrganizer)
    $old = $script:Fake.Mailboxes[$Mailbox.ToLowerInvariant()]
    $new = $script:Fake.Mailboxes[$NewOrganizer.ToLowerInvariant()]
    $ev = if ($old) { $old.Events | Where-Object id -eq $EventId | Select-Object -First 1 } else { $null }
    if (-not $ev -or -not $new) { return @{ Ok = $false; Error = 'The meeting or the new organizer was not found.' } }
    [void]$old.Events.Remove($ev)
    $moved = $ev | Select-Object *
    $moved.id = New-FakeId; $moved.Mailbox = $new.Address
    $moved.organizer = @{ emailAddress = @{ address = $new.Address; name = $new.Name } }
    $moved.attendees = @($ev.attendees | Where-Object { ([string]$_.emailAddress.address).ToLowerInvariant() -ne $new.Address })
    $new.Events.Add($moved)
    foreach ($mb in @(Get-FakeMailboxList)) {
        foreach ($c in @($mb.Events | Where-Object { $_.iCalUId -eq $ev.iCalUId -and -not $_.isOrganizer })) { $c.organizer = $moved.organizer }
    }
    return @{ Ok = $true; Error = '' }
}

function Invoke-FakeGraphRequest {
    <# One Graph request against the simulated tenant: @{ status; body; headers }. Url relative to /v1.0 or absolute. #>
    param([string]$Method, [string]$Url, $Body)

    $relative = ($Url -replace '^https://graph\.microsoft\.com/v1\.0', '')
    $script:Fake.Calls.Add([pscustomobject]@{ Method = $Method; Url = [Uri]::UnescapeDataString($relative) })
    $path, $query = $relative -split '\?', 2
    $q = Get-FakeQuery $query
    $segments = @($path.Trim('/') -split '/' | ForEach-Object { [Uri]::UnescapeDataString($_) })
    $decoded = [Uri]::UnescapeDataString($relative)
    foreach ($pattern in @($script:Fake.Throttle.Keys)) {
        if ($decoded -like $pattern -and $script:Fake.Throttle[$pattern] -gt 0) {
            $script:Fake.Throttle[$pattern]--
            return @{ status = 429; headers = @{ 'Retry-After' = '0.05' }; body = @{ error = @{ code = 'ApplicationThrottled'; message = 'Too many requests' } } }
        }
    }
    foreach ($pattern in @($script:Fake.Fail.Keys)) {
        $f = $script:Fake.Fail[$pattern]
        if ($decoded -like $pattern -and (-not $f.ContainsKey('Method') -or $f.Method -eq $Method)) { return @{ status = $f.Status; body = @{ error = @{ code = $f.Code; message = 'simulated failure' } } } }
    }
    $notFound = { param([string]$Code = 'ErrorInvalidUser') @{ status = 404; body = @{ error = @{ code = $Code; message = "The requested user is invalid." } } } }

    # ---- directory ------------------------------------------------------------------------------------
    if ($segments[0] -eq 'users' -and $segments.Count -eq 1) {
        $filter = [string]$q['$filter']
        $users = $script:Fake.Users
        if ($filter -match "proxyAddresses/any\(p:p eq 'smtp:(.+?)'\)") { $a = $Matches[1]; $users = @($users | Where-Object { @($_.proxyAddresses | ForEach-Object { $_.Substring(5).ToLowerInvariant() }) -contains $a.ToLowerInvariant() }) }
        elseif ($filter -match "userPrincipalName eq '(.+?)'") { $a = $Matches[1]; $users = @($users | Where-Object { $_.userPrincipalName -eq $a.ToLowerInvariant() }) }
        return @{ status = 200; body = @{ value = @($users | Select-Object id, displayName, mail, userPrincipalName, proxyAddresses) } }
    }
    if ($segments[0] -eq 'groups' -and $segments.Count -eq 1) {
        $filter = [string]$q['$filter']
        $a = if ($filter -match "smtp:(.+?)'") { $Matches[1].ToLowerInvariant() } else { '' }
        return @{ status = 200; body = @{ value = @($script:Fake.Groups | Where-Object mail -eq $a | Select-Object id, displayName, mail) } }
    }
    if ($segments[0] -eq 'groups' -and $segments.Count -ge 3) {
        $grp = $script:Fake.Groups | Where-Object id -eq $segments[1]
        return @{ status = 200; body = @{ value = @($grp.Members | ForEach-Object { @{ mail = $_ } }) } }
    }
    if ($segments[0] -eq 'places') {
        $rooms = @($script:Fake.Mailboxes.Values | Where-Object Kind -eq 'Room' | Select-Object -ExpandProperty Address -Unique | Sort-Object)
        $skip = [int]$q['$skip']; $top = [int]$q['$top']
        return @{ status = 200; body = @{ value = @($rooms | Select-Object -Skip $skip -First $top | ForEach-Object { @{ emailAddress = $_ } }) } }
    }

    # ---- calendars --------------------------------------------------------------------------------------
    if ($segments[0] -ne 'users') { return @{ status = 400; body = @{ error = @{ code = 'BadRequest'; message = "Unknown path $path" } } } }
    $mb = $script:Fake.Mailboxes[$segments[1].ToLowerInvariant()]
    if (-not $mb) {
        # Graph also opens a mailbox by the ID of its user.
        $user = $script:Fake.Users | Where-Object id -eq $segments[1] | Select-Object -First 1
        if ($user) { $mb = $script:Fake.Mailboxes[$user.mail] }
    }
    if (-not $mb) { return & $notFound }
    if (-not $mb.Reachable) { return & $notFound 'MailboxNotEnabledForRESTAPI' }
    if ($segments.Count -eq 2) {
        $u = $script:Fake.Users | Where-Object mail -eq $mb.Address | Select-Object -First 1
        return @{ status = 200; body = @{ id = $(if ($u) { $u.id } else { $mb.Address }); displayName = $mb.Name; mail = $mb.Address; userPrincipalName = $mb.Address } }
    }
    if ($segments[2] -eq 'calendar') { return @{ status = 200; body = @{ id = 'calendar' } } }
    if ($segments[2] -ne 'events') { return @{ status = 400; body = @{ error = @{ code = 'BadRequest'; message = "Unknown path $path" } } } }

    if ($segments.Count -eq 3 -and $Method -eq 'POST') {
        # A new event: no attendee yet (an appointment, nothing is sent).
        $id = New-FakeId
        $new = [pscustomobject]@{
            id = $id; iCalUId = (New-FakeICalUId); subject = [string]$Body.subject; type = $(if ($Body.recurrence) { 'seriesMaster' } else { 'singleInstance' })
            organizer = @{ emailAddress = @{ address = $mb.Address; name = $mb.Name } }; isOrganizer = $true
            start = @{ dateTime = ([datetime]::Parse([string]$Body.start.dateTime, [Globalization.CultureInfo]::InvariantCulture)).ToString('yyyy-MM-ddTHH:mm:ss.0000000'); timeZone = 'UTC' }
            end = @{ dateTime = ([datetime]::Parse([string]$Body.end.dateTime, [Globalization.CultureInfo]::InvariantCulture)).ToString('yyyy-MM-ddTHH:mm:ss.0000000'); timeZone = 'UTC' }
            isCancelled = $false; recurrence = $Body.recurrence; responseStatus = @{ response = 'organizer' }; showAs = 'busy'; attendees = @()
            location = $Body.location; Mailbox = $mb.Address; body = $Body.body; importance = $Body.importance; sensitivity = $Body.sensitivity; isAllDay = [bool]$Body.isAllDay
            isOnlineMeeting = [bool]($Body.PSObject.Properties['isOnlineMeeting'] -and $Body.isOnlineMeeting); allowNewTimeProposals = $true; responseRequested = $true
            DeletedOccurrences = [Collections.Generic.List[string]]::new(); CancelledOccurrences = [Collections.Generic.List[string]]::new(); Exceptions = @{}
        }
        $mb.Events.Add($new)
        $script:Fake.Created.Add($new)
        return @{ status = 201; body = (Select-FakeEvent $new) }
    }
    if ($segments.Count -eq 3) {
        $filter = [string]$q['$filter']
        if ($filter -match 'organizer') { return @{ status = 501; body = @{ error = @{ code = 'ErrorInvalidUrlQueryFilter'; message = 'is not a supported filter expression.' } } } }
        $events = @($mb.Events)
        if ($filter -match "iCalUId eq '(.+?)'") { $uid = $Matches[1]; $events = @($events | Where-Object iCalUId -eq $uid) }
        elseif ($filter -match "^subject eq '(.*)'$") { $s = $Matches[1].Replace("''", "'"); $events = @($events | Where-Object { $_.subject -ceq $s }) }
        elseif ($filter -match "end/dateTime ge '(.+?)' and start/dateTime lt '(.+?)'") {
            $from = [datetime]$Matches[1]; $to = [datetime]$Matches[2]
            $events = @($events | Where-Object { $_.type -eq 'seriesMaster' -or ([datetime]$_.end.dateTime -ge $from -and [datetime]$_.start.dateTime -lt $to) })
        }
        $top = if ($q['$top']) { [int]$q['$top'] } else { 10 }
        $skip = [int]$q['$skip']
        $page = @($events | Select-Object -Skip $skip -First $top | ForEach-Object { Select-FakeEvent $_ })
        $body = @{ value = $page }
        if ($events.Count -gt $skip + $top) {
            $q['$skip'] = $skip + $top
            $body['@odata.nextLink'] = "https://graph.microsoft.com/v1.0${path}?" + (($q.Keys | ForEach-Object { "$([Uri]::EscapeDataString($_))=$([Uri]::EscapeDataString([string]$q[$_]))" }) -join '&')
        }
        return @{ status = 200; body = $body }
    }
    $notItem = @{ status = 404; body = @{ error = @{ code = 'ErrorItemNotFound'; message = 'The specified object was not found in the store.' } } }
    # An occurrence: <id of the series copy>~<slot>.
    if ($segments[3] -match '^(.+)~(\d{4}-\d\d-\d\dT\d\d:\d\d)$') {
        $master = $mb.Events | Where-Object id -eq $Matches[1] | Select-Object -First 1
        $key = $Matches[2]
        if (-not $master -or $master.DeletedOccurrences -contains $key) { return $notItem }
        $slot = [datetime]::ParseExact($key, 'yyyy-MM-ddTHH:mm', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal)
        if ($segments.Count -eq 4 -and $Method -eq 'GET') { return @{ status = 200; body = (New-FakeInstance $master $slot) } }
        if (($segments.Count -eq 4 -and $Method -eq 'DELETE') -or ($segments[4] -eq 'permanentDelete' -and $Method -eq 'POST')) {
            # Removed from this copy only: silent for an attendee or a room; never in Recoverable Items.
            $master.DeletedOccurrences.Add($key)
            if ($master.isOrganizer -and @($master.attendees).Count) { $script:Fake.Messages.Add([pscustomobject]@{ From = $mb.Address; Kind = 'Cancellation'; ICalUId = $master.iCalUId; Occurrence = $key; Comment = '' }) }
            return @{ status = 204 }
        }
        if ($segments.Count -eq 4 -and $Method -eq 'PATCH') {
            $x = @{ Start = $slot; End = $null; Subject = ''; Location = '' }
            $inst = New-FakeInstance $master $slot
            if ($Body.start) { $x.Start = [datetime]::SpecifyKind([datetime]::Parse([string]$Body.start.dateTime, [Globalization.CultureInfo]::InvariantCulture), [DateTimeKind]::Utc) }
            $x.End = if ($Body.end) { [datetime]::SpecifyKind([datetime]::Parse([string]$Body.end.dateTime, [Globalization.CultureInfo]::InvariantCulture), [DateTimeKind]::Utc) } else { [datetime]::Parse($inst.end.dateTime, [Globalization.CultureInfo]::InvariantCulture) }
            if ($Body.subject) { $x.Subject = [string]$Body.subject }
            if ($Body.location) { $x.Location = [string]$Body.location.displayName }
            $master.Exceptions[$key] = $x
            return @{ status = 200; body = (New-FakeInstance $master $slot) }
        }
        if ($segments[4] -eq 'cancel' -and $Method -eq 'POST') {
            if (-not $master.isOrganizer) { return @{ status = 400; body = @{ error = @{ code = 'ErrorInvalidRequest'; message = 'You need to be an organizer to cancel a meeting.' } } } }
            $master.DeletedOccurrences.Add($key)
            $script:Fake.Messages.Add([pscustomobject]@{ From = $mb.Address; Kind = 'Cancellation'; ICalUId = $master.iCalUId; Occurrence = $key; Comment = [string]$Body.Comment })
            foreach ($target in @(Get-FakeMailboxList)) {
                foreach ($copy in @($target.Events | Where-Object { $_.iCalUId -eq $master.iCalUId -and -not $_.isOrganizer })) {
                    if ($target.Kind -eq 'Room') { $copy.DeletedOccurrences.Add($key) } else { $copy.CancelledOccurrences.Add($key) }
                }
            }
            return @{ status = 202 }
        }
        return @{ status = 400; body = @{ error = @{ code = 'BadRequest'; message = "Unknown request on an occurrence $Method $path" } } }
    }
    $ev = $mb.Events | Where-Object id -eq $segments[3] | Select-Object -First 1
    if (-not $ev) { return $notItem }
    if ($segments.Count -eq 4 -and $Method -eq 'GET') { return @{ status = 200; body = (Select-FakeEvent $ev) } }
    if ($segments.Count -eq 4 -and $Method -eq 'PATCH') {
        if ($Body.attendees) {
            # The invitation: the attendees of the new meeting get their copy (groups: their members), rooms accept.
            $list = @($Body.attendees | ForEach-Object { @{ emailAddress = @{ address = ([string]$_.emailAddress.address).ToLowerInvariant(); name = [string]$_.emailAddress.name }; type = [string]$_.type; status = @{ response = $(if ($_.type -eq 'resource') { 'accepted' } else { 'none' }) } } })
            $ev.attendees = $list
            $script:Fake.Messages.Add([pscustomobject]@{ From = $mb.Address; Kind = 'Invitation'; ICalUId = $ev.iCalUId; Comment = ''; To = @($list | ForEach-Object { $_.emailAddress.address }) })
            foreach ($a in $list) {
                $address = $a.emailAddress.address
                $grp = $script:Fake.Groups | Where-Object mail -eq $address | Select-Object -First 1
                foreach ($member in $(if ($grp) { @($grp.Members) } else { @($address) })) {
                    $target = $script:Fake.Mailboxes[$member]
                    if (-not $target -or $target.Address -eq $mb.Address) { continue }
                    $isRoom = $target.Kind -eq 'Room'
                    $copy = [pscustomobject]@{
                        id = (New-FakeId); iCalUId = $ev.iCalUId; subject = $(if ($isRoom) { $mb.Name } else { $ev.subject }); type = $ev.type
                        organizer = $ev.organizer; isOrganizer = $false; start = $ev.start; end = $ev.end; isCancelled = $false; recurrence = $ev.recurrence
                        responseStatus = @{ response = $(if ($isRoom) { 'accepted' } else { 'notResponded' }) }; showAs = $(if ($isRoom) { 'busy' } else { 'tentative' })
                        attendees = @(@{ emailAddress = $ev.organizer.emailAddress; type = 'required'; status = @{ response = 'none' } }) + $list; location = $ev.location; Mailbox = $target.Address
                        body = $ev.body; importance = $ev.importance; sensitivity = $ev.sensitivity; isAllDay = $ev.isAllDay; isOnlineMeeting = $ev.isOnlineMeeting; allowNewTimeProposals = $true; responseRequested = $true
                        DeletedOccurrences = [Collections.Generic.List[string]]::new([string[]]@($ev.DeletedOccurrences)); CancelledOccurrences = [Collections.Generic.List[string]]::new(); Exceptions = $ev.Exceptions.Clone()
                    }
                    $target.Events.Add($copy)
                }
            }
        }
        foreach ($name in 'subject', 'location') { if ($Body.PSObject.Properties[$name]) { $ev.$name = $Body.$name } }
        return @{ status = 200; body = (Select-FakeEvent $ev) }
    }
    if ($segments.Count -eq 4 -and $Method -eq 'DELETE') {
        [void]$mb.Events.Remove($ev)
        if ($ev.isOrganizer -and @($ev.attendees).Count) { $script:Fake.Messages.Add([pscustomobject]@{ From = $mb.Address; Kind = 'Cancellation'; ICalUId = $ev.iCalUId; Comment = '' }) }
        return @{ status = 204 }
    }
    if ($segments[4] -eq 'instances') {
        # As Graph: the occurrences whose actual time (an exception may be moved) overlaps the window.
        $from = [datetime]::Parse($q['startDateTime'], [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal)
        $to = [datetime]::Parse($q['endDateTime'], [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal)
        $occ = @(Get-FakeOccurrences $ev ([datetime]::MinValue) $to.AddDays(366))
        $parse = { param([string]$Text) [datetime]::Parse($Text, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal) }
        $items = @($occ | ForEach-Object { New-FakeInstance $ev $_ } | Where-Object { $_ } | Where-Object { (& $parse $_.end.dateTime) -gt $from -and (& $parse $_.start.dateTime) -lt $to })
        return @{ status = 200; body = @{ value = $items } }
    }
    $sendCancellation = {
        param($Organizer, [string]$Comment)
        $script:Fake.Messages.Add([pscustomobject]@{ From = $Organizer.Address; Kind = 'Cancellation'; ICalUId = $ev.iCalUId; Comment = $Comment; To = @($ev.attendees | ForEach-Object { $_.emailAddress.address }) })
        foreach ($a in @($ev.attendees)) {
            $address = ([string]$a.emailAddress.address).ToLowerInvariant()
            $grp = $script:Fake.Groups | Where-Object mail -eq $address | Select-Object -First 1
            foreach ($member in $(if ($grp) { @($grp.Members) } else { @($address) })) {
                $target = $script:Fake.Mailboxes[$member]
                if (-not $target) { continue }
                foreach ($copy in @($target.Events | Where-Object iCalUId -eq $ev.iCalUId)) {
                    if ($target.Kind -eq 'Room') { [void]$target.Events.Remove($copy) } else { $copy.isCancelled = $true; $copy.subject = "Canceled: $($copy.subject)" }
                }
            }
        }
    }
    $tracking = {
        param([string]$Response)
        # The organizer's copy records the answer of this mailbox (silently, as Exchange does).
        $org = $script:Fake.Mailboxes[([string]$ev.organizer.emailAddress.address).ToLowerInvariant()]
        if (-not $org) { return }
        foreach ($oc in @($org.Events | Where-Object { $_.iCalUId -eq $ev.iCalUId -and $_.isOrganizer })) {
            foreach ($a in @($oc.attendees)) { if (([string]$a.emailAddress.address).ToLowerInvariant() -eq $mb.Address) { $a.status = @{ response = $Response } } }
        }
    }
    if ($segments[4] -eq 'permanentDelete' -and $Method -eq 'POST') {
        [void]$mb.Events.Remove($ev)
        # Exchange: removing the meeting in the organizer's calendar sends a cancellation.
        if ($ev.isOrganizer) { & $sendCancellation $mb '' }
        else {
            # An attendee's or a room's copy: Declined / Free, the organizer's tracking Declined, no message.
            $ev.responseStatus = @{ response = 'declined' }; $ev.showAs = 'free'
            & $tracking 'declined'
        }
        $script:Fake.Purged++
        # Get-RecoverableItems gives the time of the removal to the second.
        $now = [datetime]::UtcNow
        $mb.Purges.Add([pscustomobject]@{ Event = $ev; EntryID = ('00000000ABCDEF{0:X12}0000' -f $script:Fake.Purged); LastModifiedUtc = $now.AddTicks(-($now.Ticks % [TimeSpan]::TicksPerSecond)) })
        return @{ status = 204 }
    }
    if ($segments[4] -in 'accept', 'tentativelyAccept' -and $Method -eq 'POST') {
        $answer = if ($segments[4] -eq 'accept') { 'accepted' } else { 'tentativelyAccepted' }
        $ev.responseStatus = @{ response = $answer }; $ev.showAs = if ($answer -eq 'accepted') { 'busy' } else { 'tentative' }
        & $tracking $answer
        if ($Body -and $Body.sendResponse) { $script:Fake.Messages.Add([pscustomobject]@{ From = $mb.Address; Kind = 'Response'; ICalUId = $ev.iCalUId }) }
        return @{ status = 202 }
    }
    if ($segments[4] -eq 'cancel' -and $Method -eq 'POST') {
        if (-not $ev.isOrganizer) { return @{ status = 400; body = @{ error = @{ code = 'ErrorInvalidRequest'; message = 'Your request can''t be completed. You need to be an organizer to cancel a meeting.' } } } }
        [void]$mb.Events.Remove($ev)
        & $sendCancellation $mb ([string]$Body.Comment)
        return @{ status = 202 }
    }
    return @{ status = 400; body = @{ error = @{ code = 'BadRequest'; message = "Unknown request $Method $path" } } }
}

function Invoke-FakeGraphHttp {
    <# What Start-MclGraphSend would get back: @{ Status; RetryAfter; Content }. Handles $batch. #>
    param([string]$Method, [string]$Url, [string]$Body)
    if ($Url -match '/\$batch$') {
        $payload = $Body | ConvertFrom-Json -Depth 20
        $mailboxes = @($payload.requests | ForEach-Object { if ($_.url -match '^/users/([^/?]+)') { [Uri]::UnescapeDataString($Matches[1]).ToLowerInvariant() } else { '(directory)' } })
        $script:Fake.Batches.Add([pscustomobject]@{ Count = @($payload.requests).Count; Mailboxes = $mailboxes })
        $responses = foreach ($sub in @($payload.requests)) {
            $r = Invoke-FakeGraphRequest -Method $sub.method -Url $sub.url -Body $(if ($sub.PSObject.Properties['body']) { $sub.body } else { $null })
            $item = [ordered]@{ id = $sub.id; status = $r.status }
            if ($r.ContainsKey('headers')) { $item.headers = $r.headers }
            if ($r.ContainsKey('body')) { $item.body = $r.body }
            $item
        }
        return @{ Status = 200; RetryAfter = 0; Content = (ConvertTo-FakeJson @{ responses = @($responses) }) }
    }
    $parsed = if ($Body) { $Body | ConvertFrom-Json -Depth 20 } else { $null }
    $r = Invoke-FakeGraphRequest -Method $Method -Url $Url -Body $parsed
    $retry = if ($r.ContainsKey('headers') -and $r.headers['Retry-After']) { [double]$r.headers['Retry-After'] } else { 0 }
    return @{ Status = $r.status; RetryAfter = $retry; Content = $(if ($r.ContainsKey('body')) { ConvertTo-FakeJson $r.body } else { '' }) }
}

function New-FakeToken {
    <# Unsigned app-only token with the roles given. #>
    param([string]$TenantId = '11111111-2222-3333-4444-555555555555', [string[]]$Roles = @('Calendars.ReadWrite', 'User.Read.All', 'Place.Read.All', 'GroupMember.Read.All'))
    $b64 = { param($o) [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($o | ConvertTo-Json -Compress))).TrimEnd('=').Replace('+', '-').Replace('/', '_') }
    '{0}.{1}.sig' -f (& $b64 @{ alg = 'none'; typ = 'JWT' }), (& $b64 @{ tid = $TenantId; roles = $Roles; app_displayname = 'Meeting Cleanup (test)'; appid = 'app' })
}

function Get-FakeRecoverableItems {
    <# Get-RecoverableItems -SourceFolder PurgedItems, as Get-MclPurgedItems returns it. #>
    param([string[]]$Mailbox, [datetime]$StartUtc, [datetime]$EndUtc)
    foreach ($a in $Mailbox) {
        $mb = $script:Fake.Mailboxes[$a.ToLowerInvariant()]
        if (-not $mb) { continue }
        foreach ($p in @($mb.Purges | Where-Object { $_.LastModifiedUtc -ge $StartUtc -and $_.LastModifiedUtc -le $EndUtc })) {
            [pscustomobject]@{ Mailbox = $a.ToLowerInvariant(); Subject = [string]$p.Event.subject; EntryID = $p.EntryID; LastModifiedUtc = $p.LastModifiedUtc; LastParentPath = 'Calendar' }
        }
    }
}

function Restore-FakeRecoverableItem {
    <# Restore-RecoverableItems -EntryID: the same item back in the calendar (Declined / Free, as removed). #>
    param([string]$Mailbox, [string]$EntryId)
    $mb = $script:Fake.Mailboxes[$Mailbox.ToLowerInvariant()]
    $p = if ($mb) { $mb.Purges | Where-Object EntryID -eq $EntryId | Select-Object -First 1 } else { $null }
    if (-not $p) { return @{ Ok = $false; Folder = ''; Error = 'item not found' } }
    [void]$mb.Purges.Remove($p)
    $mb.Events.Add($p.Event)
    $script:Fake.Calls.Add([pscustomobject]@{ Method = 'EXO'; Url = "Restore-RecoverableItems $Mailbox $EntryId" })
    return @{ Ok = $true; Folder = 'Calendar'; Error = '' }
}