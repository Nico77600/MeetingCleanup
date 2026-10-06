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
    Version : 1.2.0
#>

function Get-MclCleanupPlan {
    <# What an action would do on the selected meetings, without doing it (confirmation, window, banner). #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result, [Parameter(Mandatory = $true)][ValidateSet('Remove', 'Cancel')][string]$Action)

    $selected = @($Result.Meetings | Where-Object Selected)
    $cancel = [Collections.Generic.List[object]]::new()
    $remove = [Collections.Generic.List[object]]::new()
    $keep = [Collections.Generic.List[object]]::new()
    $held = [Collections.Generic.List[object]]::new()
    $acted = [Collections.Generic.List[object]]::new()
    foreach ($m in $selected) {
        # Left as they are: a meeting moved by Exchange Online (Transfer report replayed: its copies ARE the moved
        # meeting), and for Cancel a meeting whose organizer copy could not be read (removing the other copies
        # would leave his meeting live, and his next update would send it again).
        $why = if ([string](Get-MclProperty $m 'TransferMethod') -eq 'Native' -and [string](Get-MclProperty $m 'ReportStatus') -in 'Transferred', 'Partial') { "moved by Exchange Online to $(Get-MclProperty $m 'NewOrganizer'): its copies are the moved meeting" }
            elseif ($Action -eq 'Cancel' -and [string]$m.OrganizerCopy -eq 'Not read') { "the copy of its organizer could not be read: it cannot be cancelled" }
            else { '' }
        if ($why) { $held.Add([pscustomobject]@{ Meeting = $m; Reason = $why }); continue }
        $acted.Add($m)
        # The meeting of a new organizer (Transfer report) is not an old copy: never removed by a replay.
        foreach ($c in @($m.Copies | Where-Object { $_.EventId -and $_.Role -ne 'New organizer' })) {
            if ($c.Role -eq 'Organizer') { if ($Action -eq 'Cancel') { $cancel.Add($c) } else { $keep.Add($c) } }
            else { $remove.Add($c) }
        }
    }
    $selected = $acted.ToArray()
    $attendees = @($selected | ForEach-Object { @($_.Attendees) } | Where-Object { $_.Type -ne 'resource' })
    $dot = $script:Dot
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add(('{0} meeting(s) selected ({1} series)' -f $selected.Count, @($selected | Where-Object Kind -eq 'Series').Count))
    if ($cancel.Count) { $lines.Add(('{0} cancelled by the organizer: Exchange sends the cancellation message to their attendees ({1} invitation(s), external ones included)' -f $($n = @($cancel | Where-Object { (Get-MclProperty $_ 'Occurrence') }).Count; if ($n) { '{0} meeting(s) and {1} occurrence(s)' -f ($cancel.Count - $n), $n } else { $cancel.Count }), @($selected | Where-Object { @($_.Copies | Where-Object { $_.Role -eq 'Organizer' -and $_.EventId }).Count } | ForEach-Object { @($_.Attendees).Count } | Measure-Object -Sum).Sum)) }
    $lines.Add(('{0} cop{1} removed without any message in {2} mailbox(es): {3} attendee(s), {4} room(s)' -f $remove.Count, $(if ($remove.Count -eq 1) { 'y' } else { 'ies' }), @($remove | ForEach-Object Mailbox | Select-Object -Unique).Count, @($remove | Where-Object Role -ne 'Room').Count, @($remove | Where-Object Role -eq 'Room').Count))
    if ($keep.Count) { $lines.Add(('{0} meeting(s) stay in the calendar of their organizer (removing them there would send a cancellation: choose Cancel to do it)' -f @($keep | ForEach-Object MeetingId | Select-Object -Unique).Count)) }
    $occ = @($selected | Where-Object Scope -eq 'Occurrences')
    if ($occ.Count) { $lines.Add(('{0} series limited to their occurrences in the period ({1} occurrence(s)): an occurrence removed is not kept in Recoverable Items, it cannot be restored' -f $occ.Count, (@($occ | ForEach-Object Occurrences) | Measure-Object -Sum).Sum)) }
    if ($held.Count) { $lines.Add(('{0} meeting(s) left as they are: {1}' -f $held.Count, ((@($held | Group-Object Reason | ForEach-Object { "$($_.Count) $($_.Name)" })) -join '; '))) }
    [pscustomobject]@{ Action = $Action; Meetings = $selected; Cancel = $cancel.ToArray(); Remove = $remove.ToArray(); Keep = $keep.ToArray(); Held = $held.ToArray(); Lines = $lines.ToArray(); Attendees = $attendees.Count; Text = ($lines -join " $dot ") }
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

    $rank = @{ Organizer = 0; Attendee = 1; Room = 2 }
    $meetings = @($Plan.Meetings)
    $refs = @{}
    $requests = foreach ($m in $meetings) {
        $ref = $m.Copies | Where-Object { $_.EventId -and $_.Role -in 'Organizer', 'Attendee', 'Room' } | Sort-Object { $rank[$_.Role] } | Select-Object -First 1
        if (-not $ref) { continue }
        $refs[$m.MeetingId] = $ref
        # An occurrence: the series is saved (its master), with every occurrence of the copies listed below.
        $itemId = if ((Get-MclProperty $ref 'SeriesId')) { $ref.SeriesId } else { $ref.EventId }
        New-MclGraphRequest -Id $m.MeetingId -Url "$(Get-MclUserPath $ref.Mailbox)/events/$([Uri]::EscapeDataString($itemId))"
    }
    $res = Invoke-MclGraphBatch -Requests @($requests) -OnProgress { param($done, $total) Write-MclProgress ($done / [Math]::Max(1, $total)) ('{0:N0}/{1:N0} meetings saved' -f $done, $total) }
    $unread = 0
    $items = foreach ($m in $meetings) {
        $ref = $refs[$m.MeetingId]
        $r = if ($ref) { $res[$m.MeetingId] } else { $null }
        if (-not $r -or $r.Status -ne 200) { $unread++ }
        [ordered]@{
            MeetingId = $m.MeetingId; Subject = $m.Subject; Organizer = $m.Organizer; OrganizerName = $m.OrganizerName; Kind = $m.Kind; StartText = $m.StartText; EndText = $m.EndText
            ReferenceMailbox = if ($ref) { $ref.Mailbox } else { '' }
            Event = if ($r -and $r.Status -eq 200) { $r.Body } else { $null }
            Scope = [string](Get-MclProperty $m 'Scope')
            Copies = @($m.Copies | Where-Object EventId | ForEach-Object { [ordered]@{ Mailbox = $_.Mailbox; Role = $_.Role; EventId = $_.EventId; Subject = $_.Subject; Response = $_.Response; ShowAs = $_.ShowAs; Occurrence = [string](Get-MclProperty $_ 'Occurrence'); SeriesId = [string](Get-MclProperty $_ 'SeriesId') } })
        }
    }
    $backup = [ordered]@{ Tool = 'Meeting Cleanup'; Version = $script:ToolVersion; Kind = 'Backup'; CreatedUtc = [datetime]::UtcNow.ToString('o'); Action = $Plan.Action; Tenant = $Result.Tenant; Meetings = @($items) }
    [void][IO.Directory]::CreateDirectory((Split-Path $Path -Parent))
    [IO.File]::WriteAllText($Path, (ConvertTo-Json -InputObject $backup -Depth 32), [Text.UTF8Encoding]::new($false))
    Write-MclItem Ok ('Backup of {0} meeting(s) and {1} cop{2} before any change: {3}' -f $meetings.Count, @($items | ForEach-Object { $_.Copies }).Count, $(if (@($items | ForEach-Object { $_.Copies }).Count -eq 1) { 'y' } else { 'ies' }), $Path) -Icon File
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
    $waves = [Collections.Generic.List[object]]::new()
    foreach ($grp in ($Copies | Group-Object { if ((Get-MclProperty $_ 'Occurrence')) { "occ|$($_.Mailbox)|$($_.EventId)" } else { '{0}|{1}' -f $_.Mailbox, ([string]$_.Subject).Trim().ToLowerInvariant() } })) {
        for ($k = 0; $k -lt $grp.Count; $k++) {
            while ($waves.Count -le $k) { $waves.Add([Collections.Generic.List[object]]::new()) }
            $waves[$k].Add($grp.Group[$k])
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
        $occ = @($sent | Where-Object { (Get-MclProperty $_ 'Occurrence') }).Count
        Write-MclItem $(if ($failedMeetings.Count) { 'Warn' } else { 'Ok' }) ('{0} cancelled by the organizer {1} {2} failed' -f $(if ($occ) { '{0} meeting(s) and {1} occurrence(s)' -f ($sent.Count - $occ), $occ } else { '{0} meeting(s)' -f $sent.Count }), $dot, $failedMeetings.Count) -Icon Cancel
    }
    foreach ($c in $plan.Keep) {
        $c.Action = 'Keep'; $c.Result = 'Kept'
        $c.Detail = "left in the organizer's calendar: removing it there sends a cancellation to the attendees (use Cancel)"
    }

    # ---- copies of the attendees and the rooms: permanentDelete (silent) -----------------------------
    $toRemove = @($plan.Remove | Where-Object { -not $failedMeetings.Contains($_.MeetingId) })
    foreach ($c in @($plan.Remove | Where-Object { $failedMeetings.Contains($_.MeetingId) })) { $c.Action = 'None'; $c.Result = 'Not done'; $c.Detail = 'the cancellation by the organizer failed: copy left as it was' }
    if ($toRemove.Count) {
        Remove-MclCopyWaves -Copies $toRemove -ProgressText 'copies removed'
        $failed = @($toRemove | Where-Object Result -eq 'Failed')
        Write-MclItem $(if ($failed.Count) { 'Warn' } else { 'Ok' }) ('{0} cop{1} removed {2} {3} already gone {2} {4} failed' -f @($toRemove | Where-Object Result -eq 'Removed').Count, $(if (@($toRemove | Where-Object Result -eq 'Removed').Count -eq 1) { 'y' } else { 'ies' }), $dot, @($toRemove | Where-Object Result -eq 'Already gone').Count, $failed.Count) -Icon Trash
        foreach ($c in ($failed | Select-Object -First 5)) { Write-MclItem Fail "$($c.Mailbox): $($c.Detail)" }
    }
    elseif (-not $plan.Cancel.Count) { Write-MclItem Skip 'No copy to remove in the attendees and the rooms.' }

    # ---- verify ------------------------------------------------------------------------------------
    $done = @($plan.Cancel + $plan.Remove | Where-Object { $_.Result -in 'Removed', 'Cancelled' })
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
        $acted = @($m.Copies | Where-Object { $_.Action -in 'Remove', 'Cancel' })
        $ok = @($acted | Where-Object { $_.Result -in 'Removed', 'Cancelled', 'Already gone' })
        $bad = @($m.Copies | Where-Object { $_.Result -in 'Failed', 'Not done' })
        $m.Status = if ($bad.Count -and $ok.Count) { 'Partial' } elseif ($bad.Count) { 'Failed' }
            elseif (@($m.Copies | Where-Object { $_.Action -eq 'Cancel' -and $_.Result -eq 'Cancelled' }).Count) { 'Cancelled' }
            elseif ($ok.Count) { 'Removed' } elseif (@($m.Copies | Where-Object Result -eq 'Kept').Count) { 'Kept' } else { 'Nothing to do' }
        if ($Action -eq 'Cancel' -and -not @($m.Copies | Where-Object Action -eq 'Cancel').Count -and $ok.Count) { $m.Notes.Add('No meeting in the organizer''s calendar to cancel: the copies were removed without a message.') }
    }
    Update-MclResultCounts $Result
    $selected = @($Result.Meetings | Where-Object Selected)
    $Result.Status = if ($selected.Count -and -not @($selected | Where-Object { $_.Status -notin 'Failed' }).Count) { 'Failed' }
        elseif (@($selected | Where-Object { $_.Status -in 'Partial', 'Failed' }).Count -or $plan.Held.Count) { 'Warning' } else { 'Completed' }
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
