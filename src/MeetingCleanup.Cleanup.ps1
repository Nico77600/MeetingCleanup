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

    A series is handled as a whole (the series master: every occurrence and exception).

.NOTES
    Author  : Nicolas Fabert
    Version : 1.2.3
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
        if ($fast::Text($m, 'Scope') -eq 'Occurrences') { $occMeetings++; $occurrences += [int]$fast::Prop($m, 'Occurrences') }
        $list = @($m.Attendees)
        foreach ($a in $list) { if ($a.Type -ne 'resource') { $attendees++ } }
        $withOrganizer = $false
        # The meeting of a new organizer (Transfer report) is not an old copy: never removed by a replay.
        foreach ($c in $m.Copies) {
            if (-not $c.EventId -or $c.Role -eq 'New organizer') { continue }
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
    if ($held.Count) {
        $reasons = [ordered]@{}
        foreach ($h in $held) { $reasons[$h.Reason] = 1 + [int]$reasons[$h.Reason] }
        $lines.Add(('{0} meeting(s) left as they are: {1}' -f $held.Count, ((@($reasons.Keys | ForEach-Object { "$($reasons[$_]) $_" })) -join '; ')))
    }
    [pscustomobject]@{ Action = $Action; Meetings = $selected; Cancel = $cancel.ToArray(); Remove = $remove.ToArray(); Keep = $keep.ToArray(); Held = $held.ToArray(); Lines = $lines.ToArray(); Attendees = $attendees; Text = ($lines -join " $dot ") }
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
    $defaults = [ordered]@{ OrganizerKey = [string]$Meeting.Organizer; Scope = 'Whole'; Occurrences = 0; RecurrenceData = $null; TimeZone = ''; NewOrganizer = ''; NewMeetingId = ''; TransferMethod = '' }
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
            foreach ($name in 'ShowAs', 'ActionUtc', 'Occurrence', 'OccurrenceStart', 'SeriesId') { if (-not $c.PSObject.Properties[$name]) { $c | Add-Member -NotePropertyName $name -NotePropertyValue '' } }
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
