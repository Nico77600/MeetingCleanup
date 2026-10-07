<#
.SYNOPSIS
    Meeting Cleanup - transfer of meetings to a new organizer (dot-sourced by MeetingCleanup.psm1).

.DESCRIPTION
    Two ways, chosen for each meeting (Transfer.Method, -TransferMethod):

      Native     Exchange Online moves the meeting itself: Invoke-ChangeMeetingOrganizer (Exchange Online
                 PowerShell). Only when the organizer's mailbox still exists. Attendees of the organization are
                 updated silently (no new answer), external attendees receive a cancellation and an invitation.
                 The occurrences before the transfer date stay with the old organizer.
      Recreate   The meeting is created again in the new organizer's calendar from one of its copies (subject,
                 body, attendees, rooms, recurrence, occurrences removed or moved), then the old copies go.
                 Works whether the old mailbox exists or not. Measured in the lab (2026-10-06):
                   1. created without attendees (a plain appointment: no message), then the series shaped:
                      occurrences removed and moved, still without any message;
                   2. the old copies of the rooms removed (silent): their slots are free;
                   3. the attendees and the rooms added (PATCH): ONE invitation from the new organizer, with the
                      occurrences moved; the rooms accept without conflict;
                   4. the old copies of the attendees removed (silent) - or, when the old organizer still has a
                      mailbox, the old meeting cancelled by him with a message, then the copies left removed.
                 The attendees answer again. If a step fails before the invitation, the new meeting is removed
                 (nothing was sent) and the old room copies already removed can be restored (-Action Restore).
      Auto       Native when the organizer's mailbox exists, Recreate otherwise.

    Native: a series moves from the transfer date (-TransferFrom, default now); the occurrences before stay with the
    old organizer. Recreate: always from now - the next occurrence is the first one of the new series, and the old
    meeting goes whole (cancelled by an active old organizer, or removed from the attendees and the rooms), its
    past occurrences included (Backup.json keeps them). A single meeting already over is not transferred.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.2.2
#>

function Resolve-MclNewOrganizer {
    <# The new organizer: a mailbox of the tenant (its calendar can be opened). Returns @{ Address; Name; UserId }. #>
    param([Parameter(Mandatory = $true)][string]$Address)
    # The user by any of its addresses (the UPN may differ from the mail address).
    $f = [Uri]::EscapeDataString("proxyAddresses/any(p:p eq 'smtp:$($Address.ToLowerInvariant())')")
    $requests = @(
        (New-MclGraphRequest -Id 'user' -Url "$(Get-MclUserPath $Address)?`$select=id,displayName,mail,userPrincipalName"),
        (New-MclGraphRequest -Id 'proxy' -Url "/users?`$filter=$f&`$select=id,displayName,mail,userPrincipalName"),
        (New-MclGraphRequest -Id 'calendar' -Url "$(Get-MclUserPath $Address)/calendar?`$select=id")
    )
    $res = Invoke-MclGraphBatch -Requests $requests
    if ($res['calendar'].Status -ne 200) { throw "New organizer ${Address}: not a mailbox the application can open ($(Get-MclMailboxProblem $res['calendar']))." }
    $user = if ($res['user'].Status -eq 200) { $res['user'].Body } elseif ($res['proxy'].Status -eq 200) { $res['proxy'].Values | Select-Object -First 1 } else { $null }
    $mail = if ($user -and $user.mail) { ([string]$user.mail).ToLowerInvariant() } else { $Address.ToLowerInvariant() }
    [pscustomobject]@{ Address = $mail; Name = $(if ($user) { [string]$user.displayName } else { '' }); UserId = $(if ($user) { [string]$user.id } else { '' }); Input = $Address.ToLowerInvariant() }
}

function Get-MclOrganizerAccount {
    <# The account of an organizer: Present, Deleted (not in the directory any more) or Unknown (not looked up). #>
    param($Organizer)
    $account = [string](Get-MclProperty $Organizer 'Account')
    if ($account) { return $account }
    # Reports of 1.0 and 1.1 (and organizers looked up without the property): from what the search wrote.
    if ([string](Get-MclProperty $Organizer 'UserId')) { return 'Present' }
    if ([string](Get-MclProperty $Organizer 'Detail') -like '*account not in the directory*') { return 'Deleted' }
    return 'Unknown'
}

function Get-MclTransferPlan {
    <# What a transfer would do with the selected meetings, without doing it. #>
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Result,
        [Parameter(Mandatory = $true)][pscustomobject]$NewOrganizer,
        [ValidateSet('Auto', 'Native', 'Recreate')][string]$Method = 'Auto',
        [datetime]$From = [datetime]::UtcNow,
        [string]$Comment = 'This meeting is now organized by {0}.'
    )
    # A rooms search keeps occurrences and does not look the organizers up: a transfer starts from organizers.
    if ([string](Get-MclProperty $Result.Request 'Mode') -eq 'Rooms') { throw 'Transfer moves the meetings of organizers: search them with -Organizer / -OrganizerFile (a rooms search keeps occurrences and does not look the organizers up).' }
    $native = [Collections.Generic.List[object]]::new()
    $recreate = [Collections.Generic.List[object]]::new()
    $skipped = [Collections.Generic.List[object]]::new()
    # An organizer is active when his mailbox has the meeting and his account is not known to be deleted. A deleted
    # user whose mailbox Graph can still open (soft-deleted) is not: Exchange cannot move his meetings, his rooms do
    # not accept. An account not looked up (no User.Read.All) counts as active: his copy is cancelled, never left.
    $active = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $byAddress = @{}
    foreach ($o in @($Result.Organizers)) { foreach ($a in @(@($o.Addresses) + [string]$o.PrimaryAddress + [string]$o.Input)) { if ($a) { $byAddress[([string]$a).ToLowerInvariant()] = $o } } }
    $new = @($NewOrganizer.Address, $NewOrganizer.Input) | Where-Object { $_ }
    foreach ($m in @($Result.Meetings | Where-Object Selected)) {
        $reason = ''
        $org = $byAddress[([string]$m.OrganizerKey).ToLowerInvariant()]
        if (-not $org) { $org = $byAddress[([string]$m.Organizer).ToLowerInvariant()] }
        $present = $m.OrganizerCopy -eq 'Present' -and $org -and [string](Get-MclProperty $org 'State') -eq 'Mailbox' -and (Get-MclOrganizerAccount $org) -ne 'Deleted'
        $way = switch ($Method) { 'Native' { 'Native' } 'Recreate' { 'Recreate' } default { if ($present) { 'Native' } else { 'Recreate' } } }
        # The date the meeting moves from: -TransferFrom for Exchange Online, now for a re-creation.
        $since = if ($way -eq 'Native') { $From } else { [datetime]::UtcNow }
        $before = [string](Get-MclProperty $m 'ReportStatus')
        if ($new -contains [string]$m.Organizer -or $new -contains [string]$m.OrganizerKey) { $reason = "already organized by $($NewOrganizer.Address)" }
        elseif ([string](Get-MclProperty $m 'NewMeetingId') -or $before -in 'Transferred', 'Partial' -or @($m.Copies | Where-Object Role -eq 'New organizer').Count) { $reason = "already transferred by the run of this report (to $(Get-MclProperty $m 'NewOrganizer'))" }
        elseif ((Get-MclProperty $m 'Scope') -eq 'Occurrences') { $reason = 'occurrences of a series (rooms): transfer the whole series from the search of its organizer' }
        elseif ($m.Kind -ne 'Series' -and $m.End -and ([datetime]$m.End).ToUniversalTime() -le $since) { $reason = 'already over at the transfer date' }
        elseif ($m.Cancelled) { $reason = 'cancelled meeting' }
        elseif (-not @($m.Copies | Where-Object { $_.EventId -and $_.Role -in 'Organizer', 'Attendee', 'Room' }).Count) { $reason = 'no copy to read the meeting from' }
        if ($reason) { $skipped.Add([pscustomobject]@{ Meeting = $m; Reason = $reason }); continue }
        if ($present) { [void]$active.Add([string]$m.MeetingId) }
        switch ($Method) {
            'Native' { if ($present) { $native.Add($m) } else { $skipped.Add([pscustomobject]@{ Meeting = $m; Reason = 'organizer mailbox gone or account deleted: Exchange cannot move it (use -TransferMethod Auto or Recreate)' }) } }
            'Recreate' { $recreate.Add($m) }
            default { if ($present) { $native.Add($m) } else { $recreate.Add($m) } }
        }
    }
    $lines = [Collections.Generic.List[string]]::new()
    $who = if ($NewOrganizer.Name) { "$($NewOrganizer.Name) <$($NewOrganizer.Address)>" } else { $NewOrganizer.Address }
    $lines.Add(('{0} meeting(s) transferred to {1}{2}' -f ($native.Count + $recreate.Count), $who, $(if ($From -gt [datetime]::UtcNow.AddMinutes(5)) { ", from $(Format-MclDate $From ([string](Get-MclProperty $Result.Request 'TimeZone')))" } else { '' })))
    if ($native.Count) { $lines.Add(('{0} moved by Exchange Online (organizer mailbox present): the attendees of the organization are updated silently, external attendees receive a cancellation and a new invitation' -f $native.Count)) }
    if ($recreate.Count) {
        if ($From -gt [datetime]::UtcNow.AddMinutes(5)) { $lines.Add('Re-created meetings move from now: the transfer date applies to the meetings moved by Exchange Online only') }
        $withOrganizer = @($recreate | Where-Object { $active.Contains([string]$_.MeetingId) }).Count
        $leftWith = @($recreate | Where-Object { $_.OrganizerCopy -eq 'Present' -and -not $active.Contains([string]$_.MeetingId) }).Count
        if ($leftWith) { $lines.Add(('{0} of them still in the mailbox of their deleted organizer: that copy is left as it is (no cancellation from a deleted account)' -f $leftWith)) }
        $lines.Add(('{0} re-created in the calendar of the new organizer: every attendee and room receives ONE invitation from him and answers again; the old copies are removed without a message' -f $recreate.Count))
        if ($withOrganizer) { $lines.Add(('{0} of them still in the calendar of their old organizer: he cancels it, with the message "{1}"' -f $withOrganizer, ($Comment -f $who))) }
    }
    if ($skipped.Count) { $lines.Add(('{0} meeting(s) not transferred: {1}' -f $skipped.Count, ((@($skipped | Group-Object Reason | ForEach-Object { "$($_.Count) $($_.Name)" })) -join '; '))) }
    [pscustomobject]@{ Native = $native.ToArray(); Recreate = $recreate.ToArray(); Skipped = $skipped.ToArray(); Active = $active; NewOrganizer = $NewOrganizer; Lines = $lines.ToArray(); Text = ($lines -join " $($script:Dot) ") }
}

function Invoke-MclChangeMeetingOrganizer {
    <# Exchange Online moves one meeting (Invoke-ChangeMeetingOrganizer). Returns @{ Ok; Error }. #>
    param([Parameter(Mandatory = $true)][string]$Mailbox, [Parameter(Mandatory = $true)][string]$EventId, [Parameter(Mandatory = $true)][string]$NewOrganizer, [Nullable[datetime]]$From)
    $call = @{ Identity = $Mailbox; EventId = $EventId; NewOrganizer = $NewOrganizer; Confirm = $false; ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
    # The date must be in the future (else the next occurrence is the first one transferred).
    if ($null -ne $From -and ([datetime]$From).ToUniversalTime().Date -gt [datetime]::UtcNow.Date) { $call.TransferSeriesStartDate = ([datetime]$From).ToLocalTime() }
    try { $null = Invoke-ChangeMeetingOrganizer @call; return @{ Ok = $true; Error = '' } }
    catch { return @{ Ok = $false; Error = $_.Exception.Message } }
}

function Get-MclTimeKey {
    <# An occurrence of a series by its original start (the slot of the pattern), to the minute, in UTC. #>
    param($Occurrence)
    $o = Get-MclProperty $Occurrence 'originalStart'
    $d = if ($o -is [datetime]) { if ($o.Kind -eq [DateTimeKind]::Unspecified) { [datetime]::SpecifyKind($o, [DateTimeKind]::Utc) } else { $o.ToUniversalTime() } }
        elseif ($o -is [datetimeoffset]) { $o.UtcDateTime }
        elseif ($o) { ([datetimeoffset]::Parse([string]$o, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal)).UtcDateTime }
        else { ConvertTo-MclDateUtc (Get-MclProperty $Occurrence 'start') }
    return $d.ToString('yyyy-MM-ddTHH:mm', [Globalization.CultureInfo]::InvariantCulture)
}

function Format-MclGraphTime {
    <# A dateTime read from Graph (the JSON reader makes it a DateTime) as Graph expects it back. #>
    param($Value)
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-ddTHH:mm:ss.fffffff', [Globalization.CultureInfo]::InvariantCulture) }
    return [string]$Value
}

function ConvertTo-MclWallTime {
    <# A dateTime read from Graph, in the time zone asked for (Prefer outlook.timezone), as a DateTime. #>
    param($Value)
    if ($Value -is [datetime]) { return [datetime]::SpecifyKind($Value, [DateTimeKind]::Unspecified) }
    return [datetime]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture)
}

function Invoke-MclTransfer {
    <#
    .SYNOPSIS
        Transfers the selected meetings of a result to a new organizer (Native and / or Recreate, see the top of
        this file), writes the backup first, then verifies. Needs Connect-MclGraph, and Connect-MclExchange -For
        Transfer when some meetings are moved natively.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Settings,
        [Parameter(Mandatory = $true)][pscustomobject]$Result,
        [Parameter(Mandatory = $true)][pscustomobject]$Plan,
        [datetime]$From = [datetime]::UtcNow,
        [string]$Comment,
        [string]$BackupPath
    )

    $started = [datetime]::UtcNow
    $dot = $script:Dot
    $new = $Plan.NewOrganizer
    $who = if ($new.Name) { "$($new.Name) <$($new.Address)>" } else { $new.Address }
    $Result.Action = 'Transfer'
    foreach ($m in @($Result.Meetings | Where-Object { -not $_.Selected })) { $m.Status = 'Skipped' }
    foreach ($s in $Plan.Skipped) { $s.Meeting.Status = 'Skipped'; $s.Meeting.Notes.Add("Not transferred: $($s.Reason).") }

    Write-MclNextStep "Transfer to $who" 'User'
    foreach ($line in $Plan.Lines) { Write-MclItem Info $line }
    $todo = @($Plan.Native) + @($Plan.Recreate)
    if ($BackupPath -and $todo.Count) {
        $file = Save-MclBackup -Result $Result -Plan ([pscustomobject]@{ Meetings = $todo; Action = 'Transfer' }) -Path $BackupPath
        $Result | Add-Member -NotePropertyName BackupFile -NotePropertyValue $file -Force
    }

    # ---- Native: Exchange Online moves the meeting ---------------------------------------------------------
    foreach ($m in $Plan.Native) {
        Assert-MclNotCancelled
        $org = $m.Copies | Where-Object { $_.Role -eq 'Organizer' -and $_.EventId } | Select-Object -First 1
        $r = Invoke-MclChangeMeetingOrganizer -Mailbox $org.Mailbox -EventId $org.EventId -NewOrganizer $new.Address -From $From
        $org.Action = 'Transfer'; $org.ActionUtc = [datetime]::UtcNow.ToString('o')
        $m.TransferMethod = 'Native'; $m.NewOrganizer = $new.Address
        if ($r.Ok) { $org.Result = 'Transferred'; $org.Detail = "moved by Exchange Online to $($new.Address) (Invoke-ChangeMeetingOrganizer)"; $m.Status = 'Transferred' }
        else {
            $org.Result = 'Failed'; $org.Detail = "Invoke-ChangeMeetingOrganizer: $($r.Error)"; $m.Status = 'Failed'
            $m.Notes.Add('Exchange Online could not move it: -TransferMethod Recreate re-creates it with the new organizer instead.')
        }
    }
    if ($Plan.Native.Count) {
        $ok = @($Plan.Native | Where-Object Status -eq 'Transferred')
        Write-MclItem $(if ($ok.Count -lt $Plan.Native.Count) { 'Warn' } else { 'Ok' }) ('{0} of {1} moved by Exchange Online' -f $ok.Count, $Plan.Native.Count) -Icon Server
        # The new organizer's copy, to give its ID in the report (Exchange may take a moment to create it).
        if ($ok.Count) {
            $requests = for ($i = 0; $i -lt $ok.Count; $i++) {
                $f = [Uri]::EscapeDataString("subject eq '$(([string]$ok[$i].Subject).Replace("'", "''"))'")
                New-MclGraphRequest -Id "n$i" -Url "$(Get-MclUserPath $new.Address)/events?`$filter=$f&`$select=id,iCalUId,isOrganizer,type&`$top=20"
            }
            $res = Invoke-MclGraphBatch -Requests @($requests)
            for ($i = 0; $i -lt $ok.Count; $i++) {
                $hit = @($res["n$i"].Values | Where-Object isOrganizer) | Select-Object -First 1
                if ($hit) { $ok[$i].NewMeetingId = ([string]$hit.iCalUId).ToUpperInvariant(); Add-MclNewOrganizerCopy $ok[$i] $new.Address ([string]$hit.id) 'Transferred' 'the meeting in the calendar of the new organizer' }
                else { $ok[$i].Notes.Add("Not seen yet in the calendar of $($new.Address): Exchange Online may take a moment.") }
            }
        }
    }

    # ---- Recreate ---------------------------------------------------------------------------------------------
    # The time zone of each meeting: the one it was created in (read again in step 1), else the one of the report.
    $fallbackZone = Get-MclWindowsZone (Get-MclTimeZone $Settings.TimeZone).Id
    if (-not $fallbackZone) { $fallbackZone = 'UTC' }
    $rank = @{ Organizer = 0; Attendee = 1; Room = 2 }
    $work = @(foreach ($m in $Plan.Recreate) {
            $ref = $m.Copies | Where-Object { $_.EventId -and $_.Role -in 'Organizer', 'Attendee', 'Room' } | Sort-Object { $rank[$_.Role] } | Select-Object -First 1
            $tz = Get-MclWindowsZone ([string](Get-MclProperty $m 'TimeZone'))
            [pscustomobject]@{ Meeting = $m; Ref = $ref; TimeZone = $(if ($tz) { $tz } else { $fallbackZone }); Event = $null; NewId = ''; NewICalUId = ''; Failed = ''; Old = @{}; Horizon = $null }
        })
    if ($work.Count) {
        $m0 = [datetime]::UtcNow
        # By groups of 20: Stop (window) waits for the end of a group, never between a creation and its invitation.
        for ($k = 0; $k -lt $work.Count; $k += 20) {
            Assert-MclNotCancelled
            $group = @($work[$k..([Math]::Min($work.Count, $k + 20) - 1)])
            if ($script:Ui) { $script:Ui.Hold = $true }
            try { Invoke-MclRecreate -Work $group -NewOrganizer $new -Active $Plan.Active -Comment $Comment -Who $who }
            finally { if ($script:Ui) { $script:Ui.Hold = $false } }
        }
        $done = @($work | Where-Object { $_.Meeting.Status -in 'Transferred', 'Partial' })
        Write-MclItem $(if ($done.Count -lt $work.Count) { 'Warn' } else { 'Ok' }) ('{0} of {1} re-created with {2} and sent {3} {4:N1} s' -f $done.Count, $work.Count, $new.Address, $dot, ([datetime]::UtcNow - $m0).TotalSeconds) -Icon Calendar
        foreach ($w in @($work | Where-Object Failed | Select-Object -First 5)) { Write-MclItem Fail "$($w.Meeting.Subject): $($w.Failed)" }
    }

    # ---- verify: old copies gone ----------------------------------------------------------------------------------
    $gone = @($Result.Meetings | ForEach-Object { @($_.Copies | Where-Object { $_.Result -in 'Removed', 'Cancelled' -and $_.EventId }) })
    if ($Settings.Verify -and $gone.Count) {
        Write-MclNextStep 'Verify' 'Search'
        $requests = for ($i = 0; $i -lt $gone.Count; $i++) { New-MclGraphRequest -Id "v$i" -Url "$(Get-MclUserPath $gone[$i].Mailbox)/events/$([Uri]::EscapeDataString($gone[$i].EventId))?`$select=id" }
        $res = Invoke-MclGraphBatch -Requests @($requests)
        $still = 0
        for ($i = 0; $i -lt $gone.Count; $i++) {
            $r = $res["v$i"]
            if ($r.Status -eq 404) { $gone[$i].Verified = 'Yes' } elseif ($r.Status -eq 200) { $gone[$i].Verified = 'No'; $still++ } else { $gone[$i].Verified = "Unknown ($($r.Status))" }
        }
        Write-MclItem $(if ($still) { 'Warn' } else { 'Ok' }) ('{0} of {1} old copies verified gone' -f ($gone.Count - $still), $gone.Count)
    }

    Update-MclResultCounts $Result
    $selected = @($Result.Meetings | Where-Object { $_.Selected -and $_.Status -ne 'Skipped' })
    $Result.Status = if ($selected.Count -and -not @($selected | Where-Object Status -ne 'Failed').Count) { 'Failed' }
        elseif (@($selected | Where-Object { $_.Status -in 'Partial', 'Failed' }).Count -or $Plan.Skipped.Count) { 'Warning' } else { 'Completed' }
    $Result.CompletedUtc = [datetime]::UtcNow.ToString('o')
    $Result.DurationSeconds = [Math]::Round($Result.DurationSeconds + ([datetime]::UtcNow - $started).TotalSeconds, 1)
    $Result | Add-Member -NotePropertyName NewOrganizer -NotePropertyValue $new.Address -Force
    return $Result
}

function Get-MclWindowsZone {
    <# A time zone as Graph takes it (Windows ID), from a Windows or IANA ID; '' when unknown (tzone://..., custom). #>
    param([AllowEmptyString()][AllowNull()][string]$Id)
    if ([string]::IsNullOrWhiteSpace($Id) -or $Id -like 'tzone://*') { return '' }
    try { $zone = [TimeZoneInfo]::FindSystemTimeZoneById($Id) } catch { return '' }
    if ($zone.HasIanaId) { $windows = $null; if ([TimeZoneInfo]::TryConvertIanaIdToWindowsId($zone.Id, [ref]$windows)) { return $windows } else { return '' } }
    return $zone.Id
}

function Add-MclNewOrganizerCopy {
    <# The meeting in the calendar of the new organizer, as a copy of the report (role 'New organizer'). #>
    param($Meeting, [string]$Mailbox, [string]$EventId, [string]$Result, [string]$Detail)
    $c = New-MclCopy -Key $Meeting.MeetingId -Mailbox $Mailbox -Role 'New organizer' -Via 'Transfer' -Event $null -Result $Result -Detail $Detail
    $c.EventId = $EventId; $c.Action = 'Transfer'; $c.ActionUtc = [datetime]::UtcNow.ToString('o')
    $Meeting.Copies.Add($c)
}

function Invoke-MclRecreate {
    <#
        Re-creates a group of meetings with the new organizer (Invoke-MclTransfer, steps at the top of this file).
        Each item of -Work: Meeting, Ref (the copy read), TimeZone, and what the steps fill in. A meeting whose new
        series does not match the old one, or that fails before its invitation, is removed again (nothing sent).
    #>
    param([Parameter(Mandatory = $true)][object[]]$Work, [Parameter(Mandatory = $true)]$NewOrganizer, [Parameter(Mandatory = $true)]$Active, [string]$Comment, [string]$Who)
    $new = $NewOrganizer
    $invariant = [Globalization.CultureInfo]::InvariantCulture
    $utcStyle = [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal

    # 1. each meeting in full, in the time zone it was created in (read again when the copy gives another one)
    $read = {
        param([object[]]$Items)
        $requests = for ($i = 0; $i -lt $Items.Count; $i++) {
            $w = $Items[$i]
            New-MclGraphRequest -Id "f$i" -Url "$(Get-MclUserPath $w.Ref.Mailbox)/events/$([Uri]::EscapeDataString($w.Ref.EventId))?`$select=subject,body,start,end,location,attendees,recurrence,importance,sensitivity,isOnlineMeeting,onlineMeetingProvider,allowNewTimeProposals,isAllDay,responseRequested,type,organizer,originalStartTimeZone" -Headers @{ Prefer = "outlook.timezone=`"$($w.TimeZone)`", outlook.body-content-type=`"html`"" }
        }
        $res = Invoke-MclGraphBatch -Requests @($requests) -OnProgress { param($done, $total) Write-MclProgress ($done / [Math]::Max(1, $total)) ('{0:N0}/{1:N0} meetings read' -f $done, $total) }
        for ($i = 0; $i -lt $Items.Count; $i++) {
            if ($res["f$i"].Status -eq 200) { $Items[$i].Event = $res["f$i"].Body } else { $Items[$i].Failed = "not read from $($Items[$i].Ref.Mailbox): $($res["f$i"].Status) $($res["f$i"].ErrorCode)" }
        }
    }
    & $read $Work
    $again = @(foreach ($w in @($Work | Where-Object { -not $_.Failed })) {
            $zone = if ([string]$w.Event.type -eq 'seriesMaster') { Get-MclWindowsZone ([string](Get-MclProperty (Get-MclProperty $w.Event.recurrence 'range') 'recurrenceTimeZone')) } else { '' }
            if (-not $zone) { $zone = Get-MclWindowsZone ([string](Get-MclProperty $w.Event 'originalStartTimeZone')) }
            if ($zone -and $zone -ne $w.TimeZone) { $w.TimeZone = $zone; $w }
        })
    if ($again.Count) { & $read $again }

    # 2. series: the occurrences still to come, from up to three copies (one removed by an attendee is in another)
    $now = [datetime]::UtcNow
    $nowText = $now.ToString('yyyy-MM-ddTHH:mm:ss') + 'Z'
    $requests = [Collections.Generic.List[object]]::new()
    $seriesWork = @($Work | Where-Object { -not $_.Failed -and [string]$_.Event.type -eq 'seriesMaster' })
    foreach ($w in $seriesWork) {
        $range = Get-MclProperty $w.Event.recurrence 'range'
        $w.Horizon = switch ([string](Get-MclProperty $range 'type')) {
            'endDate' { [datetime]::ParseExact([string]$range.endDate, 'yyyy-MM-dd', $invariant).AddDays(2) }
            'numbered' { $now.AddYears(5) }
            default { $now.AddYears(2) }
        }
        # The organizer's copy knows every occurrence; without it, an attendee and a room (one may have removed
        # an occurrence the other still has).
        $orgCopy = @($w.Meeting.Copies | Where-Object { $_.EventId -and $_.Role -eq 'Organizer' } | Select-Object -First 1)
        $sources = if ($orgCopy.Count) { $orgCopy } else { @(@($w.Meeting.Copies | Where-Object { $_.EventId -and $_.Role -eq 'Attendee' } | Select-Object -First 1) + @($w.Meeting.Copies | Where-Object { $_.EventId -and $_.Role -eq 'Room' } | Select-Object -First 1)) | Where-Object { $_ } }
        foreach ($c in $sources) {
            $requests.Add((New-MclGraphRequest -Id "i$($requests.Count)|$([array]::IndexOf($Work, $w))|$($c.Role)" -Url "$(Get-MclUserPath $c.Mailbox)/events/$([Uri]::EscapeDataString($c.EventId))/instances?startDateTime=$nowText&endDateTime=$($w.Horizon.ToString('yyyy-MM-ddTHH:mm:ss'))Z&`$select=id,start,end,originalStart,type,location,subject,isCancelled&`$top=500" -Headers @{ Prefer = "outlook.timezone=`"$($w.TimeZone)`"" }))
        }
    }
    if ($requests.Count) {
        $res = Invoke-MclGraphBatch -Requests $requests.ToArray() -FollowPages -OnProgress { param($done, $total) Write-MclProgress ($done / [Math]::Max(1, $total)) ('{0:N0}/{1:N0} series read' -f $done, $total) }
        $priority = @{ Organizer = 0; Attendee = 1; Room = 2 }
        foreach ($id in $res.Keys) {
            $parts = $id.Split('|'); $w = $Work[[int]$parts[1]]; $role = $parts[2]
            if ($res[$id].Status -ne 200) { continue }
            foreach ($o in $res[$id].Values) {
                # An occurrence cancelled by the organizer stays in an attendee's calendar, marked cancelled.
                if ([bool](Get-MclProperty $o 'isCancelled')) { continue }
                # Still to come: its start (an occurrence moved from a past slot to a later date is one of them).
                if ((ConvertTo-MclDateUtc $o.start) -lt $now) { continue }
                $key = Get-MclTimeKey $o
                $known = $w.Old[$key]
                if (-not $known -or $priority[$role] -lt $priority[$known.Role]) { $w.Old[$key] = [pscustomobject]@{ Role = $role; Occurrence = $o } }
            }
        }
        foreach ($w in $seriesWork) { if (-not $w.Old.Count) { $w.Failed = 'no occurrence of the series after the transfer date: nothing to transfer' } }
    }

    # 3. the new meeting, without attendees (nothing is sent)
    $creates = [Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $Work.Count; $i++) {
        $w = $Work[$i]
        if ($w.Failed) { continue }
        $ev = $w.Event
        $subject = if ($w.Ref.Role -eq 'Room') { [string]$w.Meeting.Subject } else { [string]$ev.subject }
        $body = [ordered]@{
            subject = $subject; body = @{ contentType = 'html'; content = [string](Get-MclProperty $ev.body 'content') }
            start = @{ dateTime = (Format-MclGraphTime $ev.start.dateTime); timeZone = $w.TimeZone }; end = @{ dateTime = (Format-MclGraphTime $ev.end.dateTime); timeZone = $w.TimeZone }
            location = @{ displayName = [string](Get-MclProperty $ev.location 'displayName') }
            importance = [string]$ev.importance; sensitivity = [string]$ev.sensitivity; isAllDay = [bool]$ev.isAllDay
            allowNewTimeProposals = [bool]$ev.allowNewTimeProposals; responseRequested = [bool](Get-MclProperty $ev 'responseRequested')
            showAs = 'busy'; isReminderOn = $true; transactionId = [guid]::NewGuid().ToString()
        }
        if ($ev.isOnlineMeeting) {
            $body.isOnlineMeeting = $true; $body.onlineMeetingProvider = 'teamsForBusiness'
            $w.Meeting.Notes.Add('Online meeting: a new Teams link is created for the new organizer; the text of the invitation may still show the old link.')
        }
        if ([string]$ev.type -eq 'seriesMaster') {
            # The new series starts with the first occurrence still to come (its pattern slot), at the time of the series.
            $first = $w.Old.Keys | Sort-Object | Select-Object -First 1
            $zone = try { [TimeZoneInfo]::FindSystemTimeZoneById($w.TimeZone) } catch { [TimeZoneInfo]::Utc }
            $slot = [TimeZoneInfo]::ConvertTimeFromUtc([datetime]::ParseExact($first, 'yyyy-MM-ddTHH:mm', $invariant, $utcStyle), $zone)
            $s0 = ConvertTo-MclWallTime $ev.start.dateTime
            $e0 = ConvertTo-MclWallTime $ev.end.dateTime
            $start = $slot.Date.Add($s0.TimeOfDay)
            $body.start = @{ dateTime = $start.ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = $w.TimeZone }
            $body.end = @{ dateTime = $start.Add($e0 - $s0).ToString('yyyy-MM-ddTHH:mm:ss'); timeZone = $w.TimeZone }
            $range = $ev.recurrence.range
            $newRange = [ordered]@{ type = [string]$range.type; startDate = $start.ToString('yyyy-MM-dd'); recurrenceTimeZone = $w.TimeZone }
            switch ([string]$range.type) {
                'endDate' { $newRange.endDate = [string]$range.endDate }
                'numbered' {
                    # The last occurrence of the old series is the end of the new one.
                    $last = $w.Old.Keys | Sort-Object | Select-Object -Last 1
                    $lastLocal = [TimeZoneInfo]::ConvertTimeFromUtc([datetime]::ParseExact($last, 'yyyy-MM-ddTHH:mm', $invariant, $utcStyle), $zone)
                    $newRange.type = 'endDate'; $newRange.endDate = $lastLocal.ToString('yyyy-MM-dd')
                }
            }
            $body.recurrence = @{ pattern = $ev.recurrence.pattern; range = $newRange }
        }
        $creates.Add((New-MclGraphRequest -Id "c$i" -Method 'POST' -Url "$(Get-MclUserPath $new.Address)/events" -Body $body))
    }
    if ($creates.Count) {
        Write-MclItem Info ('{0} meeting(s) created again in the calendar of {1}, without attendees for now (nothing is sent)' -f $creates.Count, $new.Address) -Icon Calendar
        $res = Invoke-MclGraphBatch -Requests $creates.ToArray()
        for ($i = 0; $i -lt $Work.Count; $i++) {
            $r = $res["c$i"]
            if (-not $r) { continue }
            if ($r.Status -in 200, 201) { $Work[$i].NewId = [string]$r.Body.id; $Work[$i].NewICalUId = ([string]$r.Body.iCalUId).ToUpperInvariant() }
            else { $Work[$i].Failed = "not created in the calendar of $($new.Address): $($r.Status) $($r.ErrorCode) $($r.ErrorMessage)" }
        }
    }

    # 4. series shaped like the old one: occurrences removed or moved (still no attendee: no message). Every
    #    occurrence to come of the old series must have its slot in the new one, else the meeting is not transferred.
    $shape = @($seriesWork | Where-Object { $_.NewId -and -not $_.Failed })
    if ($shape.Count) {
        $requests = for ($i = 0; $i -lt $shape.Count; $i++) {
            $w = $shape[$i]
            $from = [datetime]::ParseExact(($w.Old.Keys | Sort-Object | Select-Object -First 1), 'yyyy-MM-ddTHH:mm', $invariant, $utcStyle).AddDays(-1)
            New-MclGraphRequest -Id "s$i" -Url "$(Get-MclUserPath $new.Address)/events/$([Uri]::EscapeDataString($w.NewId))/instances?startDateTime=$($from.ToString('yyyy-MM-ddTHH:mm:ss'))Z&endDateTime=$($w.Horizon.ToString('yyyy-MM-ddTHH:mm:ss'))Z&`$select=id,start,end,originalStart,location,subject&`$top=500" -Headers @{ Prefer = "outlook.timezone=`"$($w.TimeZone)`"" }
        }
        $res = Invoke-MclGraphBatch -Requests @($requests) -FollowPages
        $changes = [Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $shape.Count; $i++) {
            $w = $shape[$i]
            if ($res["s$i"].Status -ne 200) { $w.Failed = "occurrences of the new series not read: $($res["s$i"].Status) $($res["s$i"].ErrorCode)"; continue }
            $mine = [Collections.Generic.List[object]]::new()
            $slots = [Collections.Generic.HashSet[string]]::new()
            $removed = 0; $moved = 0
            foreach ($o in $res["s$i"].Values) {
                $key = Get-MclTimeKey $o
                [void]$slots.Add($key)
                $old = $w.Old[$key]
                if (-not $old) { $mine.Add((New-MclGraphRequest -Id "x$($changes.Count + $mine.Count)" -Method 'DELETE' -Url "$(Get-MclUserPath $new.Address)/events/$([Uri]::EscapeDataString([string]$o.id))")); $removed++; continue }
                $oo = $old.Occurrence
                if ([string]$oo.type -ne 'exception') { continue }
                $patch = [ordered]@{}
                if ((Format-MclGraphTime $oo.start.dateTime) -ne (Format-MclGraphTime $o.start.dateTime) -or (Format-MclGraphTime $oo.end.dateTime) -ne (Format-MclGraphTime $o.end.dateTime)) {
                    $patch.start = @{ dateTime = (Format-MclGraphTime $oo.start.dateTime); timeZone = $w.TimeZone }; $patch.end = @{ dateTime = (Format-MclGraphTime $oo.end.dateTime); timeZone = $w.TimeZone }
                }
                # The time and the subject of an occurrence are carried over. Not its location: Exchange takes the
                # room out of it when the room declines that occurrence, and the new series books the room again.
                if ($old.Role -ne 'Room' -and [string]$oo.subject -and [string]$oo.subject -ne [string]$o.subject) { $patch.subject = [string]$oo.subject }
                if ($patch.Count) { $mine.Add((New-MclGraphRequest -Id "x$($changes.Count + $mine.Count)" -Method 'PATCH' -Url "$(Get-MclUserPath $new.Address)/events/$([Uri]::EscapeDataString([string]$o.id))" -Body $patch)); $moved++ }
            }
            $missing = @($w.Old.Keys | Where-Object { -not $slots.Contains($_) } | Sort-Object)
            if ($missing.Count) {
                $w.Failed = '{0} occurrence(s) of the old series without a slot in the new one (first: {1} UTC; time zone of the series taken as {2}): not transferred' -f $missing.Count, $missing[0].Replace('T', ' '), $w.TimeZone
                continue
            }
            foreach ($rq in $mine) { $changes.Add([pscustomobject]@{ Work = $w; Request = $rq }) }
            if ($removed -or $moved) { $w.Meeting.Notes.Add(('New series shaped like the old one: {0} occurrence(s) removed, {1} moved or renamed.' -f $removed, $moved)) }
        }
        if ($changes.Count) {
            $res = Invoke-MclGraphBatch -Requests @($changes | ForEach-Object Request)
            foreach ($ch in $changes) {
                $r = $res[$ch.Request.Id]
                if ($r.Status -notin 200, 204 -and -not $ch.Work.Failed) { $ch.Work.Failed = "occurrence of the new series not changed: $($r.Status) $($r.ErrorCode) $($r.ErrorMessage)" }
            }
        }
    }

    # Failed after the creation and before any invitation: the new meeting (no attendee yet) is removed, silently.
    $rollback = {
        param([object[]]$Items)
        $list = @($Items | Where-Object { $_.NewId })
        if (-not $list.Count) { return }
        $requests = for ($i = 0; $i -lt $list.Count; $i++) { New-MclGraphRequest -Id "d$i" -Method 'DELETE' -Url "$(Get-MclUserPath $new.Address)/events/$([Uri]::EscapeDataString($list[$i].NewId))" }
        $res = Invoke-MclGraphBatch -Requests @($requests)
        for ($i = 0; $i -lt $list.Count; $i++) {
            $w = $list[$i]
            if ($res["d$i"].Status -in 200, 204, 404) { $w.NewId = ''; $w.NewICalUId = ''; $w.Meeting.Notes.Add('The meeting created for the new organizer was removed (nothing had been sent).') }
            else { $w.Meeting.Notes.Add("The meeting created for the new organizer (no attendee, nothing sent) could not be removed ($($res["d$i"].Status)): delete it from the calendar of $($new.Address).") }
        }
    }
    & $rollback @($Work | Where-Object Failed)

    # 5. the old copies of the rooms removed (silent): their slots are free for the new invitation
    $ready = @($Work | Where-Object { $_.NewId -and -not $_.Failed })
    $roomCopies = @($ready | ForEach-Object { @($_.Meeting.Copies | Where-Object { $_.EventId -and $_.Role -eq 'Room' }) })
    if ($roomCopies.Count) {
        Remove-MclCopyWaves -Copies $roomCopies -ProgressText 'old room copies removed'
        foreach ($w in $ready) {
            $bad = @($w.Meeting.Copies | Where-Object { $_.Role -eq 'Room' -and $_.Action -eq 'Remove' -and $_.Result -eq 'Failed' })
            if ($bad.Count) { $w.Failed = "old copy of the room $($bad[0].Mailbox) not removed: $($bad[0].Detail)" }
        }
        & $rollback @($ready | Where-Object Failed)
        $ready = @($ready | Where-Object { -not $_.Failed })
    }

    # 6. the invitation: attendees and rooms added to the new meeting
    if ($ready.Count) {
        $requests = for ($i = 0; $i -lt $ready.Count; $i++) {
            $w = $ready[$i]
            $oldOrganizer = @([string]$w.Meeting.Organizer, [string]$w.Meeting.OrganizerKey) | Where-Object { $_ }
            $attendees = @(@($w.Event.attendees) | Where-Object { $_ -and $_.emailAddress.address } | Where-Object {
                    $a = ([string]$_.emailAddress.address).ToLowerInvariant()
                    $oldOrganizer -notcontains $a -and $a -ne $new.Address -and $a -ne $new.Input
                } | ForEach-Object { @{ emailAddress = @{ address = [string]$_.emailAddress.address; name = [string]$_.emailAddress.name }; type = $(if ([string]$_.type) { [string]$_.type } else { 'required' }) } })
            if (-not $attendees.Count) { $w.Meeting.Notes.Add('No attendee left to invite: the meeting is an appointment of the new organizer.') }
            New-MclGraphRequest -Id "a$i" -Method 'PATCH' -Url "$(Get-MclUserPath $new.Address)/events/$([Uri]::EscapeDataString($w.NewId))" -Body @{ attendees = $attendees }
        }
        $res = Invoke-MclGraphBatch -Requests @($requests) -OnProgress { param($done, $total) Write-MclProgress ($done / [Math]::Max(1, $total)) ('{0:N0}/{1:N0} invitations sent' -f $done, $total) }
        for ($i = 0; $i -lt $ready.Count; $i++) {
            $r = $res["a$i"]
            if ($r.Status -ne 200) { $ready[$i].Failed = "attendees not invited: $($r.Status) $($r.ErrorCode) $($r.ErrorMessage)" }
            elseif ($r.Body.iCalUId) { $ready[$i].NewICalUId = ([string]$r.Body.iCalUId).ToUpperInvariant() }
        }
        & $rollback @($ready | Where-Object Failed)
        $ready = @($ready | Where-Object { -not $_.Failed })
    }

    # 7. the old meeting: cancelled by its old organizer when he is still active, then the old copies removed
    $cancels = @($ready | Where-Object { $Active.Contains([string]$_.Meeting.MeetingId) } | ForEach-Object { @($_.Meeting.Copies | Where-Object { $_.EventId -and $_.Role -eq 'Organizer' }) })
    foreach ($w in @($ready | Where-Object { -not $Active.Contains([string]$_.Meeting.MeetingId) })) {
        foreach ($c in @($w.Meeting.Copies | Where-Object { $_.EventId -and $_.Role -eq 'Organizer' })) { $c.Action = 'Keep'; $c.Result = 'Kept'; $c.Detail = 'left in the mailbox of the deleted organizer: no cancellation is sent from a deleted account' }
    }
    if ($cancels.Count) {
        $text = if ($Comment) { $Comment -f $Who } else { "This meeting is now organized by $Who." }
        $requests = for ($i = 0; $i -lt $cancels.Count; $i++) { New-MclGraphRequest -Id "k$i" -Method 'POST' -Url "$(Get-MclUserPath $cancels[$i].Mailbox)/events/$([Uri]::EscapeDataString($cancels[$i].EventId))/cancel" -Body @{ Comment = $text } }
        $res = Invoke-MclGraphBatch -Requests @($requests)
        for ($i = 0; $i -lt $cancels.Count; $i++) {
            Set-MclCopyResult $cancels[$i] 'Cancel' $res["k$i"] 'Cancelled'
            $cancels[$i].ActionUtc = $res["k$i"].DoneUtc.ToString('o')
            if ($cancels[$i].Result -eq 'Cancelled') { $cancels[$i].Detail = "cancelled by the old organizer: the meeting is now organized by $($new.Address)" }
        }
    }
    $attendeeCopies = @($ready | ForEach-Object { @($_.Meeting.Copies | Where-Object { $_.EventId -and $_.Role -eq 'Attendee' }) })
    if ($attendeeCopies.Count) { Remove-MclCopyWaves -Copies $attendeeCopies -ProgressText 'old attendee copies removed' }

    # 8. results
    foreach ($w in $Work) {
        $m = $w.Meeting
        $m.TransferMethod = 'Recreate'; $m.NewOrganizer = $new.Address
        if ($w.Failed) {
            $m.Status = 'Failed'
            $m.Notes.Add("Not transferred: $($w.Failed).")
            if ($w.NewId) { Add-MclNewOrganizerCopy $m $new.Address $w.NewId 'Failed' 'created without attendees and not removed again: delete it from the calendar of the new organizer' }
            if (@($m.Copies | Where-Object { $_.Role -eq 'Room' -and $_.Result -eq 'Removed' }).Count) { $m.Notes.Add('Old room copies were removed before the failure: -Action Restore -FromReport <this report> puts them back.') }
            continue
        }
        $m.NewMeetingId = $w.NewICalUId
        Add-MclNewOrganizerCopy $m $new.Address $w.NewId 'Created' ('re-created by the tool; invitation sent to {0} attendee(s) and room(s)' -f @($m.Copies | Where-Object { $_.Role -in 'Attendee', 'Room' -and $_.EventId }).Count)
        $left = @($m.Copies | Where-Object { $_.Role -in 'Organizer', 'Attendee', 'Room' -and $_.Action -in 'Remove', 'Cancel' -and $_.Result -eq 'Failed' })
        $m.Status = if ($left.Count) { 'Partial' } else { 'Transferred' }
        if ($left.Count) { $m.Notes.Add(('{0} old cop{1} not removed: attendees may see the meeting twice (see the copies).' -f $left.Count, $(if ($left.Count -eq 1) { 'y' } else { 'ies' }))) }
    }
}
