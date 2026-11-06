<#
.SYNOPSIS
    Meeting Cleanup - remove or cancel, then verify (dot-sourced by MeetingCleanup.psm1).

.DESCRIPTION
    Behaviour of Exchange Online, measured in the lab (Graph v1.0, 2026-10-05):

      - Removing the copy of an ATTENDEE or a ROOM (permanentDelete) is silent: no message to the organizer
        or to anyone. The item goes to Recoverable Items\Purges (restorable by an administrator for the
        retention of deleted items, 14 days by default: Restore-RecoverableItems).
      - Removing the meeting from the ORGANIZER's calendar, with DELETE or permanentDelete, always sends a
        cancellation to every attendee still invited, external ones included. Graph has no silent way.
      - cancel (organizer) sends the cancellation with a message; the rooms free the slot themselves.

    So the two actions are:
      Remove   silent: the copies of the attendees and the rooms are removed; the meeting stays in the
               organizer's calendar when the mailbox still exists (deleting it would send a cancellation).
      Cancel   the organizer cancels the meeting (message to every attendee, rooms released), then the
               copies left in the attendees' and rooms' calendars are removed. A meeting no longer in the
               organizer's calendar (or with a deleted organizer) cannot be cancelled: its copies are removed.

    A series is handled as a whole (the series master: every occurrence and exception), unless it was limited to
    its occurrences in the period (rooms mode, -SeriesScope Occurrences): then each occurrence is acted on, and
    the occurrences left out in the window (SkippedOccurrences) are left as they are.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.3.1
#>

function Get-MclCleanupPlan {
    <# What an action would do on the selected meetings, without doing it (confirmation, window, banner). #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result, [Parameter(Mandatory = $true)][ValidateSet('Remove', 'Cancel')][string]$Action)

    $fast = [MeetingCleanupNative.Fast]
    $cancel = [Collections.Generic.List[object]]::new()
    $remove = [Collections.Generic.List[object]]::new()
    $keep = [Collections.Generic.List[object]]::new()
    $held = [Collections.Generic.List[object]]::new()
    $acted = [Collections.Generic.List[object]]::new()
    $notSelected = [Collections.Generic.List[object]]::new()
    $skippedOccurrences = 0
    $series = 0; $attendees = 0; $invited = 0; $occMeetings = 0; $occurrences = 0; $cancelOccurrences = 0; $removeRooms = 0
    $removeMailboxes = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $keptMeetings = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($m in $Result.Meetings) {
        if (-not $m.Selected) { continue }
        # Left as they are: a meeting moved by Exchange Online (Transfer report replayed: its copies ARE the moved
        # meeting), and for Cancel a meeting whose organizer copy could not be read (removing the other copies
        # would leave his meeting live, and his next update would send it again).
        $why = if ($fast::Text($m, 'TransferMethod') -eq 'Native' -and $fast::Text($m, 'ReportStatus') -in 'Transferred', 'Partial') { "moved by Exchange Online to $($fast::Text($m, 'NewOrganizer')): its copies are the moved meeting" }
            elseif ($Action -eq 'Cancel' -and [string]$m.OrganizerCopy -eq 'Not read') { "the copy of its organizer could not be read: it cannot be cancelled" }
            else { '' }
        if ($why) { $held.Add([pscustomobject]@{ Meeting = $m; Reason = $why }); continue }
        $acted.Add($m)
        if ($m.Kind -eq 'Series') { $series++ }
        # The occurrences left out in the window: their copies are left as they are.
        $skip = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($k in @(Get-MclProperty $m 'SkippedOccurrences')) { if ($k) { [void]$skip.Add([string]$k) } }
        if ($fast::Text($m, 'Scope') -eq 'Occurrences') { $occMeetings++; $occurrences += [Math]::Max(0, [int]$fast::Prop($m, 'Occurrences') - $skip.Count); $skippedOccurrences += $skip.Count }
        $list = @($m.Attendees)
        foreach ($a in $list) { if ($a.Type -ne 'resource') { $attendees++ } }
        $withOrganizer = $false
        # The meeting of a new organizer (Transfer report) is not an old copy: never removed by a replay.
        foreach ($c in $m.Copies) {
            if (-not $c.EventId -or $c.Role -eq 'New organizer') { continue }
            if ($skip.Count -and $skip.Contains([MeetingCleanupNative.GuiRows]::OccurrenceKeyOf($c))) { $notSelected.Add($c); continue }
            if ($c.Role -eq 'Organizer') {
                $withOrganizer = $true
                if ($Action -eq 'Cancel') { $cancel.Add($c); if ($fast::Text($c, 'Occurrence')) { $cancelOccurrences++ } }
                else { $keep.Add($c); [void]$keptMeetings.Add($c.MeetingId) }
            }
            else { $remove.Add($c); [void]$removeMailboxes.Add($c.Mailbox); if ($c.Role -eq 'Room') { $removeRooms++ } }
        }
        if ($withOrganizer) { $invited += $list.Count }
    }
    $selected = $acted.ToArray()
    $dot = $script:Dot
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add(('{0} meeting(s) selected ({1} series)' -f $selected.Count, $series))
    if ($cancel.Count) { $lines.Add(('{0} cancelled by the organizer: Exchange sends the cancellation message to their attendees ({1} invitation(s), external ones included)' -f $(if ($cancelOccurrences) { '{0} meeting(s) and {1} occurrence(s)' -f ($cancel.Count - $cancelOccurrences), $cancelOccurrences } else { $cancel.Count }), $invited)) }
    $lines.Add(('{0} cop{1} removed without any message in {2} mailbox(es): {3} attendee(s), {4} room(s)' -f $remove.Count, $(if ($remove.Count -eq 1) { 'y' } else { 'ies' }), $removeMailboxes.Count, ($remove.Count - $removeRooms), $removeRooms))
    if ($keep.Count) { $lines.Add(('{0} meeting(s) stay in the calendar of their organizer (removing them there would send a cancellation: choose Cancel to do it)' -f $keptMeetings.Count)) }
    if ($occMeetings) { $lines.Add(('{0} series limited to their occurrences in the period ({1} occurrence(s)): an occurrence removed is not kept in Recoverable Items, it cannot be restored' -f $occMeetings, $occurrences)) }
    if ($skippedOccurrences) { $lines.Add(('{0} occurrence(s) not ticked: left as they are' -f $skippedOccurrences)) }
    if ($held.Count) {
        $reasons = [ordered]@{}
        foreach ($h in $held) { $reasons[$h.Reason] = 1 + [int]$reasons[$h.Reason] }
        $lines.Add(('{0} meeting(s) left as they are: {1}' -f $held.Count, ((@($reasons.Keys | ForEach-Object { "$($reasons[$_]) $_" })) -join '; ')))
    }
    [pscustomobject]@{ Action = $Action; Meetings = $selected; Cancel = $cancel.ToArray(); Remove = $remove.ToArray(); Keep = $keep.ToArray(); Held = $held.ToArray(); NotSelected = $notSelected.ToArray(); Lines = $lines.ToArray(); Attendees = $attendees; Text = ($lines -join " $dot ") }
}
function Set-MclCopyResult {
    param($Copy, [string]$Action, $Response, [string]$Success)
    $Copy.Action = $Action
    $Copy.HttpStatus = [int]$Response.Status
    if ($Response.Status -in 200, 202, 204) { $Copy.Result = $Success; $Copy.Detail = '' }
    elseif ($Response.Status -eq 404 -and $Response.ErrorCode -in 'ErrorItemNotFound', 'ResourceNotFound') { $Copy.Result = 'Already gone'; $Copy.Detail = 'not in the calendar any more' }
    else { $Copy.Result = 'Failed'; $Copy.Detail = ("{0} {1}: {2}" -f $Response.Status, $Response.ErrorCode, $Response.ErrorMessage).Trim() }
}

function Save-MclBackup {
    <#
        Before any change: the full content of each meeting acted on (organizer copy when there is one, else the
        copy of an attendee: subject, body, attendees, recurrence, location...) and the state of each copy
        (mailbox, role, item ID, subject, response, free/busy). Written to Path; nothing is changed when the file
        cannot be written. The restore itself uses Recoverable Items (Invoke-MclRestore); this file keeps what
        was there, to check it or to send a meeting again by hand once the retention of deleted items is over.
    #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result, [Parameter(Mandatory = $true)][pscustomobject]$Plan, [Parameter(Mandatory = $true)][string]$Path)

    $meetings = @($Plan.Meetings)
    $refs = @{}
    $requests = foreach ($m in $meetings) {
        $ref = Get-MclBestCopy $m -OrganizerAttendeeRoomOnly
        if (-not $ref) { continue }
        $refs[$m.MeetingId] = $ref
        # An occurrence: the series is saved (its master), with every occurrence of the copies listed below.
        $itemId = if ((Get-MclProperty $ref 'SeriesId')) { $ref.SeriesId } else { $ref.EventId }
        New-MclGraphRequest -Id $m.MeetingId -Url "$(Get-MclUserPath $ref.Mailbox)/events/$([Uri]::EscapeDataString($itemId))"
    }
    $res = Invoke-MclGraphBatch -Requests @($requests) -OnProgress { param($done, $total) Write-MclProgress ($done / [Math]::Max(1, $total)) ('{0:N0}/{1:N0} meetings saved' -f $done, $total) }
    $unread = 0
    $saved = 0
    $fast = [MeetingCleanupNative.Fast]
    $items = foreach ($m in $meetings) {
        $ref = $refs[$m.MeetingId]
        $r = if ($ref) { $res[$m.MeetingId] } else { $null }
        if (-not $r -or $r.Status -ne 200) { $unread++ }
        $copies = [Collections.Generic.List[object]]::new()
        foreach ($c in $m.Copies) {
            if (-not $c.EventId) { continue }
            $copies.Add([ordered]@{ Mailbox = $c.Mailbox; Role = $c.Role; EventId = $c.EventId; Subject = $c.Subject; Response = $c.Response; ShowAs = $c.ShowAs; Occurrence = $fast::Text($c, 'Occurrence'); SeriesId = $fast::Text($c, 'SeriesId') })
        }
        $saved += $copies.Count
        [ordered]@{
            MeetingId = $m.MeetingId; Subject = $m.Subject; Organizer = $m.Organizer; OrganizerName = $m.OrganizerName; Kind = $m.Kind; StartText = $m.StartText; EndText = $m.EndText
            ReferenceMailbox = if ($ref) { $ref.Mailbox } else { '' }
            Event = if ($r -and $r.Status -eq 200) { $r.Body } else { $null }
            Scope = $fast::Text($m, 'Scope')
            Copies = $copies.ToArray()
        }
    }
    $backup = [ordered]@{ Tool = 'Meeting Cleanup'; Version = $script:ToolVersion; Kind = 'Backup'; CreatedUtc = [datetime]::UtcNow.ToString('o'); Action = $Plan.Action; Tenant = $Result.Tenant; Meetings = @($items) }
    [void][IO.Directory]::CreateDirectory((Split-Path $Path -Parent))
    [IO.File]::WriteAllText($Path, (ConvertTo-Json -InputObject $backup -Depth 32), [Text.UTF8Encoding]::new($false))
    Write-MclItem Ok ('Backup of {0} meeting(s) and {1} cop{2} before any change: {3}' -f $meetings.Count, $saved, $(if ($saved -eq 1) { 'y' } else { 'ies' }), $Path) -Icon File
    if ($unread) { Write-MclItem Warn "$unread meeting(s) could not be read for the backup (already gone, or the mailbox is not reachable): their state is in the report." }
    return $Path
}

function Remove-MclCopyWaves {
    <#
        permanentDelete of copies of attendees and rooms (silent), with the time of each removal (ActionUtc, the end
        of its own $batch call). Copies with the same subject in one mailbox (a room shows the organizer's name) are
        removed one after the other, at least 3 seconds apart: their order is then certain in Recoverable Items,
        for the restore. An occurrence never goes there: always in the first wave. On Stop, the copies already
        removed keep their result.
    #>
    param([Parameter(Mandatory = $true)][object[]]$Copies, [string]$ProgressText = 'copies removed')
    # Waves: the n-th copy of each (mailbox, subject) group in wave n. A copy alone in its group is in wave 0.
    # An occurrence never goes to Recoverable Items (nothing to restore by order): always in wave 0.
    $groups = [ordered]@{}
    foreach ($c in $Copies) {
        $key = if ([MeetingCleanupNative.Fast]::Text($c, 'Occurrence')) { "occ|$($c.Mailbox)|$($c.EventId)" } else { '{0}|{1}' -f $c.Mailbox, ([string]$c.Subject).Trim().ToLowerInvariant() }
        $list = $groups[$key]
        if (-not $list) { $list = [Collections.Generic.List[object]]::new(); $groups[$key] = $list }
        $list.Add($c)
    }
    $waves = [Collections.Generic.List[object]]::new()
    foreach ($list in $groups.Values) {
        for ($k = 0; $k -lt $list.Count; $k++) {
            while ($waves.Count -le $k) { $waves.Add([Collections.Generic.List[object]]::new()) }
            $waves[$k].Add($list[$k])
        }
    }
    $removeTotal = $Copies.Count; $before = 0
    $lastDone = [datetime]::MinValue
    foreach ($wave in $waves) {
        # At least 3 seconds between two waves: in Recoverable Items (time to the second) the copies with the
        # same subject in one mailbox are then in the order of the waves, without a tie.
        while (([datetime]::UtcNow - $lastDone).TotalSeconds -lt 3) { Wait-MclUi 200 }
        $items = $wave.ToArray()
        $requests = for ($i = 0; $i -lt $items.Count; $i++) {
            New-MclGraphRequest -Id "r$i" -Method 'POST' -Url "$(Get-MclUserPath $items[$i].Mailbox)/events/$([Uri]::EscapeDataString($items[$i].EventId))/permanentDelete"
        }
        $removeOffset = $before
        try { $res = Invoke-MclGraphBatch -Requests @($requests) -OnProgress { param($done, $all) Write-MclProgress (($removeOffset + $done) / [Math]::Max(1, $removeTotal)) ('{0:N0}/{1:N0} {2}' -f ($removeOffset + $done), $removeTotal, $ProgressText) } }
        catch [OperationCanceledException] {
            # Stopped: the copies already removed are in the report, with their time (they can be restored).
            $partial = $_.Exception.Data['Results']
            if ($partial) { for ($i = 0; $i -lt $items.Count; $i++) { $r = $partial["r$i"]; if ($r -and $r.Done) { Set-MclCopyResult $items[$i] 'Remove' $r 'Removed'; $items[$i].ActionUtc = $r.DoneUtc.ToString('o') } } }
            throw
        }
        # The time of each removal: the end of its $batch call (Graph runs the 20 requests of a call in turn).
        for ($i = 0; $i -lt $items.Count; $i++) {
            Set-MclCopyResult $items[$i] 'Remove' $res["r$i"] 'Removed'
            $items[$i].ActionUtc = $res["r$i"].DoneUtc.ToString('o')
            if ($res["r$i"].DoneUtc -gt $lastDone) { $lastDone = $res["r$i"].DoneUtc }
        }
        $before += $items.Count
    }
    if ($waves.Count -gt 1) { Write-MclLog 'INFO' "Removal in $($waves.Count) waves: copies with the same subject in one mailbox removed one after the other." }
}

function Invoke-MclCleanup {
    <#
    .SYNOPSIS
        Removes or cancels the selected meetings of a result (Find-MclMeetings or Import-MclReport), then verifies.
    .PARAMETER Comment
        Cancel: the message of the cancellation (plain text).
    .PARAMETER BackupPath
        File of the backup written before any change (Save-MclBackup). The command line and the window write it
        in the folder of the report.
    .NOTES
        The time of each removal is kept (ActionUtc): the restore finds the copy in Recoverable Items by
        mailbox, subject and that time. Copies with the same subject in the same mailbox (a room shows the name
        of the organizer as subject) are removed one after the other, so that their order is known.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Settings,
        [Parameter(Mandatory = $true)][pscustomobject]$Result,
        [Parameter(Mandatory = $true)][ValidateSet('Remove', 'Cancel')][string]$Action,
        [string]$Comment,
        [string]$BackupPath
    )

    $started = [datetime]::UtcNow
    $dot = $script:Dot
    $plan = Get-MclCleanupPlan -Result $Result -Action $Action
    $Result.Action = $Action
    foreach ($m in @($Result.Meetings | Where-Object { -not $_.Selected })) { $m.Status = 'Skipped' }
    foreach ($h in $plan.Held) {
        $h.Meeting.Status = 'Skipped'
        $h.Meeting.Notes.Add("Not acted on: $($h.Reason).")
        foreach ($c in @($h.Meeting.Copies | Where-Object { $_.EventId -and $_.Role -ne 'New organizer' })) { $c.Action = 'None'; $c.Result = 'Not processed'; $c.Detail = "left as it is: $($h.Reason)" }
    }
    foreach ($c in $plan.NotSelected) { $c.Action = 'None'; $c.Result = 'Skipped'; $c.Detail = 'occurrence not ticked: left as it is' }

    Write-MclNextStep $(if ($Action -eq 'Cancel') { 'Cancel and clean' } else { 'Remove silently' }) $(if ($Action -eq 'Cancel') { 'Cancel' } else { 'Trash' })
    foreach ($line in $plan.Lines) { Write-MclItem Info $line }
    if ($BackupPath -and $plan.Meetings.Count) {
        $file = Save-MclBackup -Result $Result -Plan $plan -Path $BackupPath
        $Result | Add-Member -NotePropertyName BackupFile -NotePropertyValue $file -Force
    }

    # ---- organizer: cancel (Cancel) or keep (Remove) -------------------------------------------------
    $failedMeetings = [Collections.Generic.HashSet[string]]::new()
    if ($plan.Cancel.Count) {
        $body = @{ Comment = [string]$Comment }
        $requests = for ($i = 0; $i -lt $plan.Cancel.Count; $i++) {
            $c = $plan.Cancel[$i]
            New-MclGraphRequest -Id "c$i" -Method 'POST' -Url "$(Get-MclUserPath $c.Mailbox)/events/$([Uri]::EscapeDataString($c.EventId))/cancel" -Body $body
        }
        try { $res = Invoke-MclGraphBatch -Requests @($requests) -OnProgress { param($done, $total) Write-MclProgress ($done / [Math]::Max(1, $total)) ('{0:N0}/{1:N0} meetings cancelled' -f $done, $total) } }
        catch [OperationCanceledException] {
            # Stopped: the cancellations already sent are in the report.
            $partial = $_.Exception.Data['Results']
            if ($partial) { for ($i = 0; $i -lt $plan.Cancel.Count; $i++) { $r = $partial["c$i"]; if ($r -and $r.Done) { Set-MclCopyResult $plan.Cancel[$i] 'Cancel' $r 'Cancelled'; $plan.Cancel[$i].ActionUtc = $r.DoneUtc.ToString('o') } } }
            throw
        }
        for ($i = 0; $i -lt $plan.Cancel.Count; $i++) {
            $c = $plan.Cancel[$i]
            Set-MclCopyResult $c 'Cancel' $res["c$i"] 'Cancelled'
            $c.ActionUtc = $res["c$i"].DoneUtc.ToString('o')
            if ($c.Result -eq 'Failed') { [void]$failedMeetings.Add($c.MeetingId); Write-MclItem Fail "Cancel $($c.Mailbox): $($c.Detail)" }
        }
        $sent = @($plan.Cancel | Where-Object Result -eq 'Cancelled')
        $occ = 0; foreach ($c in $sent) { if ([MeetingCleanupNative.Fast]::Text($c, 'Occurrence')) { $occ++ } }
        Write-MclItem $(if ($failedMeetings.Count) { 'Warn' } else { 'Ok' }) ('{0} cancelled by the organizer {1} {2} failed' -f $(if ($occ) { '{0} meeting(s) and {1} occurrence(s)' -f ($sent.Count - $occ), $occ } else { '{0} meeting(s)' -f $sent.Count }), $dot, $failedMeetings.Count) -Icon Cancel
    }
    foreach ($c in $plan.Keep) {
        $c.Action = 'Keep'; $c.Result = 'Kept'
        $c.Detail = "left in the organizer's calendar: removing it there sends a cancellation to the attendees (use Cancel)"
    }

    # ---- copies of the attendees and the rooms: permanentDelete (silent) -----------------------------
    $toRemove = [Collections.Generic.List[object]]::new()
    foreach ($c in $plan.Remove) {
        if ($failedMeetings.Contains($c.MeetingId)) { $c.Action = 'None'; $c.Result = 'Not done'; $c.Detail = 'the cancellation by the organizer failed: copy left as it was' }
        else { $toRemove.Add($c) }
    }
    if ($toRemove.Count) {
        Remove-MclCopyWaves -Copies $toRemove.ToArray() -ProgressText 'copies removed'
        $removed = 0; $gone = 0; $failed = [Collections.Generic.List[object]]::new()
        foreach ($c in $toRemove) { switch ($c.Result) { 'Removed' { $removed++ } 'Already gone' { $gone++ } 'Failed' { $failed.Add($c) } } }
        Write-MclItem $(if ($failed.Count) { 'Warn' } else { 'Ok' }) ('{0} cop{1} removed {2} {3} already gone {2} {4} failed' -f $removed, $(if ($removed -eq 1) { 'y' } else { 'ies' }), $dot, $gone, $failed.Count) -Icon Trash
        foreach ($c in ($failed | Select-Object -First 5)) { Write-MclItem Fail "$($c.Mailbox): $($c.Detail)" }
    }
    elseif (-not $plan.Cancel.Count) { Write-MclItem Skip 'No copy to remove in the attendees and the rooms.' }

    # ---- verify ------------------------------------------------------------------------------------
    $done = [Collections.Generic.List[object]]::new()
    foreach ($c in @($plan.Cancel) + @($plan.Remove)) { if ($c.Result -in 'Removed', 'Cancelled') { $done.Add($c) } }
    if ($Settings.Verify -and $done.Count) {
        Write-MclNextStep 'Verify' 'Search'
        $requests = for ($i = 0; $i -lt $done.Count; $i++) {
            New-MclGraphRequest -Id "v$i" -Url "$(Get-MclUserPath $done[$i].Mailbox)/events/$([Uri]::EscapeDataString($done[$i].EventId))?`$select=id,isCancelled"
        }
        $res = Invoke-MclGraphBatch -Requests @($requests)
        $still = 0
        for ($i = 0; $i -lt $done.Count; $i++) {
            $c = $done[$i]; $r = $res["v$i"]
            if ($r.Status -eq 404) { $c.Verified = 'Yes' }
            elseif ($r.Status -eq 200) { $c.Verified = 'No'; $c.Result = 'Failed'; $c.Detail = 'still in the calendar after the request'; $still++ }
            else { $c.Verified = "Unknown ($($r.Status))" }
        }
        Write-MclItem $(if ($still) { 'Warn' } else { 'Ok' }) ('{0} of {1} verified gone{2}' -f ($done.Count - $still), $done.Count, $(if ($still) { " $dot $still still present" } else { '' }))
    }

    # ---- status of each meeting and of the run -------------------------------------------------------
    foreach ($m in @($plan.Meetings)) {
        $ok = 0; $bad = 0; $cancelled = $false; $cancelTried = $false; $kept = $false
        foreach ($c in $m.Copies) {
            if ($c.Action -in 'Remove', 'Cancel' -and $c.Result -in 'Removed', 'Cancelled', 'Already gone') { $ok++ }
            if ($c.Result -in 'Failed', 'Not done') { $bad++ }
            if ($c.Action -eq 'Cancel') { $cancelTried = $true; if ($c.Result -eq 'Cancelled') { $cancelled = $true } }
            if ($c.Result -eq 'Kept') { $kept = $true }
        }
        $m.Status = if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } elseif ($cancelled) { 'Cancelled' } elseif ($ok) { 'Removed' } elseif ($kept) { 'Kept' } else { 'Nothing to do' }
        if ($Action -eq 'Cancel' -and -not $cancelTried -and $ok) { $m.Notes.Add('No meeting in the organizer''s calendar to cancel: the copies were removed without a message.') }
    }
    Update-MclResultCounts $Result
    $selectedCount = 0; $failedCount = 0; $warned = $false
    foreach ($m in $Result.Meetings) {
        if (-not $m.Selected) { continue }
        $selectedCount++
        if ($m.Status -eq 'Failed') { $failedCount++ }
        if ($m.Status -in 'Partial', 'Failed') { $warned = $true }
    }
    $Result.Status = if ($selectedCount -and $failedCount -eq $selectedCount) { 'Failed' } elseif ($warned -or $plan.Held.Count) { 'Warning' } else { 'Completed' }
    $Result.CompletedUtc = [datetime]::UtcNow.ToString('o')
    $Result.DurationSeconds = [Math]::Round($Result.DurationSeconds + ([datetime]::UtcNow - $started).TotalSeconds, 1)
    $Result | Add-Member -NotePropertyName CleanupComment -NotePropertyValue $(if ($Action -eq 'Cancel') { [string]$Comment } else { '' }) -Force
    return $Result
}

function Add-MclMeetingDefaults {
    <# A meeting read from a report of an older version: the properties added since, with their default. #>
    param($Meeting)
    $defaults = [ordered]@{ OrganizerKey = [string]$Meeting.Organizer; Scope = 'Whole'; Occurrences = 0; RecurrenceData = $null; TimeZone = ''; NewOrganizer = ''; NewMeetingId = ''; TransferMethod = ''; SkippedOccurrences = @() }
    foreach ($name in $defaults.Keys) { if (-not $Meeting.PSObject.Properties[$name]) { $Meeting | Add-Member -NotePropertyName $name -NotePropertyValue $defaults[$name] } }
}

function Import-MclReport {
    <#
    .SYNOPSIS
        Reads a report written by Export-MclReport (its folder or its Summary.json) to replay it: the same
        meetings and copies, by their IDs, without searching again.
    .PARAMETER MeetingId
        Only these meetings of the report (column MeetingId).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path, [string[]]$MeetingId)

    $file = if (Test-Path -LiteralPath $Path -PathType Container) { Get-ChildItem -LiteralPath $Path -Filter '*-Summary.json' -File | Select-Object -First 1 -ExpandProperty FullName } else { $Path }
    if (-not $file -or -not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "No *-Summary.json in ${Path}: give the folder of a Meeting Cleanup report." }
    $data = [IO.File]::ReadAllText($file) | ConvertFrom-Json -Depth 32
    if ($data.Tool -ne 'Meeting Cleanup' -or -not $data.PSObject.Properties['Meetings']) { throw "$file is not a Meeting Cleanup report." }
    $ids = [Collections.Generic.HashSet[string]]::new([string[]]@($MeetingId | Where-Object { $_ } | ForEach-Object { ([string]$_).ToUpperInvariant() }), [StringComparer]::OrdinalIgnoreCase)
    $meetings = [Collections.Generic.List[object]]::new()
    foreach ($m in @($data.Meetings)) {
        if ($ids.Count -and -not $ids.Contains([string]$m.MeetingId)) { continue }
        $copies = [Collections.Generic.List[object]]::new()
        foreach ($c in @($m.Copies)) {
            # Reports of 1.0.0 and 1.1.0: the properties added since.
            foreach ($name in 'ShowAs', 'ActionUtc', 'Occurrence', 'OccurrenceStart', 'OccurrenceKey', 'SeriesId') { if (-not $c.PSObject.Properties[$name]) { $c | Add-Member -NotePropertyName $name -NotePropertyValue '' } }
            if ($c.EventId) { $c.Action = ''; $c.Result = ''; $c.HttpStatus = 0; $c.Detail = ''; $c.Verified = ''; $c.ActionUtc = '' }
            $copies.Add($c)
        }
        Add-MclMeetingDefaults $m
        $notes = [Collections.Generic.List[string]]::new()
        foreach ($n in @($m.Notes)) { if ($n) { $notes.Add([string]$n) } }
        $m.Copies = $copies
        $m.Notes = $notes
        # The status in the report (a Transfer replayed skips the meetings it transferred).
        $m | Add-Member -NotePropertyName ReportStatus -NotePropertyValue ([string]$m.Status) -Force
        $m.Selected = $true
        $m.Status = 'Found'
        $meetings.Add($m)
    }
    if ($ids.Count -and $meetings.Count -lt $ids.Count) { Write-MclItem Warn ('{0} meeting ID(s) not in the report.' -f ($ids.Count - $meetings.Count)) }
    $result = [pscustomobject]@{
        Tool = 'Meeting Cleanup'; Version = $script:ToolVersion; Action = 'Report'; Status = 'Completed'; Error = ''
        StartedUtc = [datetime]::UtcNow.ToString('o'); CompletedUtc = ''; DurationSeconds = 0.0
        Request = $data.Request; Tenant = $data.Tenant; Organization = $data.Organization; AppId = $data.AppId; AppName = $data.AppName
        Organizers = @($data.Organizers); Searched = $data.Searched; Meetings = $meetings; Warnings = [Collections.Generic.List[string]]::new(); Counts = $null
        FromReport = $file
    }
    Update-MclResultCounts $result
    return $result
}

# SIG # Begin signature block
# MIImIgYJKoZIhvcNAQcCoIImEzCCJg8CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBQgjfTHLbXJMJi
# 362kJu9rtYJIcSoK8rX82OThvQ2CCaCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCACDA47/c5/
# 3SFOCjUJVvYMvBH/m1sE7ngsu8W6LcIAJTANBgkqhkiG9w0BAQEFAASCAgBs16PC
# 0tsj1yEjVKV88E4JIDBE+214/nCY9SGVY5/oRk8cwM1a2derbqgBhd4Tf0sNAMvg
# K5jpEdb6a0NpubiUxd7YGsT/0HoFbwrFIwZIgWnYSInP7jYzPwU8Ozn7Wm2vqWp4
# AZw/G/M+ugBcA/KohieNk7HTNzzfy5gVP/cw5b/FMagTAuiJlYaPykBZfalL9cSm
# 6bsW16ZN4DwF532UxW37O587wNYzWpscbYFUPNFSgNI9XewjNx6O0xytxkWd+a/v
# rIZsnDHXprq6XeBHwIFIr1rS1lj7vJ1N8gZ90xDYd4u09hg86pzJdXJPwb9AuAC3
# U4R90fs6J3jtcWBuNfbsVdf08ZMnhRmEeyj+HOahnITWILi8Kzu5SMa9EcYSsbeu
# y+bJCurlsbgTVgwxfAW2FtZxYFl9n5xECQ7CV68A8myy1FXffILd+mFfq3jDQQvs
# nXdQMxyUMCOJUmbZBQY+l+WQ+noq18mYKf+0a5B+zPU+vA1rxghW8UJaATgIk35s
# IHZQq5PaqHVL8V/rbxkAX+0/tIwBARncLHydIwEJkaMMrI5InRa6dUIKBOrxr3Cg
# tDIInkZLSTH+G+5t0mjsyLtHuJeJAeGu5oa01dYtoh1LLZnaJA/83fGucR/8Hf4G
# c+JEAXQTmSRLN/PQOuEov/HFy5y/QQRIugT/AqGCFMYwghTCBgorBgEEAYI3AwMB
# MYIUsjCCFK4GCSqGSIb3DQEHAqCCFJ8wghSbAgEDMQ0wCwYJYIZIAWUDBAIBMIGB
# BgsqhkiG9w0BCRABBKByBHAwbgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIEppq9bMuBIvuX6TVgz1CvCYjs0tH6REpSHHLlkpRw1/AghjZaTKyxQ/
# kRgPMjAyNjExMDYwODQyMjJaMAMCAQECCEzxxwCX2HTdoIIRpzCCBPMwggLboAMC
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
# CSqGSIb3DQEJBTEPFw0yNjExMDYwODQyMjJaMCgGCSqGSIb3DQEJNDEbMBkwCwYJ
# YIZIAWUDBAIBoQoGCCqGSM49BAMCMC8GCSqGSIb3DQEJBDEiBCCm9ZGdrXMNcvnw
# TOFCsikR4ImWI86b7nO+CA5ETFK4LDCByAYLKoZIhvcNAQkQAi8xgbgwgbUwgbIw
# ga8EIFdP6heq+l58aiubCMmYsXNeRQs4c8pE5qJMoJMOY/OwMIGKMHakdDByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7jMAoGCCqGSM49BAMCBEcw
# RQIhAJt8M3od3CvhjKa9c0ORccAEWMOeKhuB5A8PEMl+j2eOAiBHwrxqAzZNgQLa
# 6QUwTqf4phBFUXSJtv3UhMf6xeM+uA==
# SIG # End signature block
