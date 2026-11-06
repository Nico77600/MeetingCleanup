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
    Version : 1.3.1
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
        elseif ((Get-MclProperty $m 'Scope') -eq 'Occurrences') { $reason = 'occurrences of a series: transfer the whole series (search its organizer with -SeriesScope Whole)' }
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

# SIG # Begin signature block
# MIImIgYJKoZIhvcNAQcCoIImEzCCJg8CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCR1mbHwtHQ0XcX
# b7PtCa+62LvoMyzGxvNyLeolH/mGLKCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYCTCCGAUCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBZ2hpvGPYR
# wvwgcA3IXYrXuYxxaM4/YwS/ZaBVlPBFJzANBgkqhkiG9w0BAQEFAASCAgBs/8gZ
# 5fCTJCJ/hlQVn86ypXl9busCMch6QA6ZuZbxZsgRD+4e9UpRZRJaR5gtp1OuFXb6
# wH0YoZFNkf4luorb1fJxYdNnNah10fJwRD8Bj0m/NWYJEMvAoSoN8R15ViyIuLZ+
# NxQL+HZc6ppcJEAyoOi+uHBY1Y34DAVhkRvgIJPnKCEzuZCMnWZMirgGou+j3kSb
# 1h4/XcLn8foTWYt2R7WZBTh0AWLkFjDtTf+P/5szLuG0chrUjcpTWnShhBivtBYR
# vpRJbdhrI99U5qbeKlVNqE3S5xGU55EgIPM81YlGNKMzbzftDpGGgnuxClbcupye
# HJ/MtrkUx//e2jdBb7gHg0qSEIVJAq7Ydii+8v8yKo5AKUSphsA7UZgU/yegAZnS
# Fpq1xv94C2ZEBF5St2cyo6MibYmsaurM/v6KLQ54NM1JsFlpvE6Y9ddaghkiN3BR
# P36QO+u7aAcOoaG5jMmcZr3ZKKljoeIegGoC6kZ9j4bsIp2PZTuoaqa4LetifQi3
# 0Uu26sXTkAhy07S+oSVzBVWJJgkNLXNRSPC05U3wbPUJGo/sHdJrPXg8UGU06FlD
# If5133033so1ZpZseCoJAuwxe5CCNtCpmVGNZ8SW5NM7jmhi7iBrLWBJSTSNlLD3
# hUsNVU2XuJfYTZOaAnOkEgX+z5kLZ6Klp5QrxKGCFMYwghTCBgorBgEEAYI3AwMB
# MYIUsjCCFK4GCSqGSIb3DQEHAqCCFJ8wghSbAgEDMQ0wCwYJYIZIAWUDBAIBMIGB
# BgsqhkiG9w0BCRABBKByBHAwbgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEII0QcjWN0YgFgy81x9LxajWz1ZveHOuOpcXizX8m0xerAghIDi+Ljkfc
# 6RgPMjAyNjExMDYwODQyMjZaMAMCAQECCGVsiZXBpMXsoIIRpzCCBPMwggLboAMC
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
# 9sdantP6nYRQFTpF8czGc+FGresbAH819Pv5vBhCluTo8Px2PpC71cM0jzsxggJW
# MIICUgIBATCBhjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29t
# IFRpbWVzdGFtcGluZyBJc3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7j
# MAsGCWCGSAFlAwQCAaCCAWAwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMBwG
# CSqGSIb3DQEJBTEPFw0yNjExMDYwODQyMjZaMCgGCSqGSIb3DQEJNDEbMBkwCwYJ
# YIZIAWUDBAIBoQoGCCqGSM49BAMCMC8GCSqGSIb3DQEJBDEiBCCipX1ybCUPN9I0
# e/hV5guPHlG26OV5ZmeoyJQee4EygDCByAYLKoZIhvcNAQkQAi8xgbgwgbUwgbIw
# ga8EIFdP6heq+l58aiubCMmYsXNeRQs4c8pE5qJMoJMOY/OwMIGKMHakdDByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7jMAoGCCqGSM49BAMCBEcw
# RQIhAPZkVQwZS8qgQD4gkjUBUdTEOOrNC4Fx0AhXdicpj4kzAiB6mAietBZCUD2X
# yQEQUWfHzW7QOz4LWpttEKBYGngYfg==
# SIG # End signature block
