<#
.SYNOPSIS
    Meeting Cleanup - restore of the copies removed by a run (dot-sourced by MeetingCleanup.psm1).

.DESCRIPTION
    A copy removed by the tool (permanentDelete) is in Recoverable Items\Purges of its mailbox for the
    retention of deleted items (14 days by default, up to 30; longer with a hold). Measured in a lab tenant
    (Exchange Online, 2026-10-05):
      - Get-RecoverableItems -SourceFolder PurgedItems lists it: subject of the copy, item class
        IPM.Appointment, LastModifiedTime = time of the removal (UTC, text MM/dd/yyyy HH:mm:ss).
      - Its EntryID is a new one (the item moved): it is found by mailbox, subject and time of the removal.
        A room shows the name of the organizer as subject: the tool removed such copies one after the other,
        so their order is known (ActionUtc).
      - Restore-RecoverableItems -EntryID puts back the same item (same immutable ID, same series) in the
        calendar, in about 2 seconds, without any message.
      - On removal Exchange marked the copy Declined and Free, and the organizer's tracking shows the
        attendee Declined (no message). After the restore the tool answers again silently (sendResponse
        false): Accepted, or Tentative for an invitation that had no answer - the room is busy again and the
        organizer's tracking is corrected.

    What cannot be restored: a meeting cancelled by its organizer (the attendees received the cancellation:
    the organizer has to send it again), a copy whose retention is over, a copy already back.

    Exchange Online PowerShell (module ExchangeOnlineManagement 3.2+) runs Get-RecoverableItems and
    Restore-RecoverableItems: role Mailbox Import Export, which no role group has by default (developer guide,
    chapter 5). Graph checks that each copy is back and answers the invitation again.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.3.1
#>

$script:Exchange = $null

function Connect-MclExchange {
    <#
    .SYNOPSIS
        Connects to Exchange Online PowerShell: as the application (certificate, or the client secret through an
        access token) or as an administrator (Restore.Connection Interactive).
    .PARAMETER For
        Restore: Get-RecoverableItems and Restore-RecoverableItems (role Mailbox Import Export).
        Transfer: Invoke-ChangeMeetingOrganizer with its parameters EventId and NewOrganizer (custom role from User
        Options, or Mail Recipients).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Settings, [Security.SecureString]$Secret, [ValidateSet('Restore', 'Transfer')][string]$For = 'Restore')

    $module = Get-Module -ListAvailable ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $module -or $module.Version -lt [version]'3.2.0') {
        throw "The $($For.ToLowerInvariant()) needs the module ExchangeOnlineManagement 3.2 or later: Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force"
    }
    Import-Module $module.Path -ErrorAction Stop
    $commands = if ($For -eq 'Transfer') { @('Invoke-ChangeMeetingOrganizer') } else { @('Get-RecoverableItems', 'Restore-RecoverableItems') }
    $common = @{ ShowBanner = $false; CommandName = $commands; ErrorAction = 'Stop' }
    $how = ''
    if ($Settings.RestoreConnection -eq 'Interactive') {
        if ($Settings.RestoreUser) { $common.UserPrincipalName = $Settings.RestoreUser }
        $how = "administrator $(if ($Settings.RestoreUser) { $Settings.RestoreUser } else { '(sign-in window)' })"
    }
    else {
        $organization = if ([string]$Settings.Organization -match '\.onmicrosoft\.(com|us|de|cn)$') { [string]$Settings.Organization }
            elseif ([string]$Settings.TenantId -match '\.onmicrosoft\.(com|us|de|cn)$') { [string]$Settings.TenantId }
            else { throw 'Exchange Online PowerShell as the application: set Tenant.Organization to the initial domain of the tenant (contoso.onmicrosoft.com), or Restore.Connection = ''Interactive''.' }
        $common.AppId = $Settings.AppId
        $common.Organization = $organization
        if ($Settings.AuthMode -eq 'Certificate') { $common.Certificate = Get-MclCertificate $Settings.CertificateThumbprint }
        else {
            if (-not $Secret) {
                $value = [Environment]::GetEnvironmentVariable($Settings.ClientSecretVariable)
                if ($value) { $Secret = ConvertTo-SecureString $value -AsPlainText -Force; $value = $null }
                else { throw "ClientSecret mode: the environment variable $($Settings.ClientSecretVariable) is empty." }
            }
            $common.AccessToken = (Get-MclAppToken -Settings $Settings -Secret $Secret -Scope 'https://outlook.office365.com/.default').Token
        }
        $how = "application $($Settings.AppId) on $organization"
    }
    Write-MclItem Info "Exchange Online PowerShell: $how (about 10 seconds)..." -Icon Server
    Wait-MclUi 10 -NoCancel
    $before = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | ForEach-Object ConnectionId)
    try { Connect-ExchangeOnline @common | Out-Null }
    catch { throw "Exchange Online PowerShell: $($_.Exception.Message)" }
    $connection = Get-ConnectionInformation | Where-Object { $_.ConnectionId -notin $before } | Select-Object -Last 1
    $script:Exchange = @{ ConnectionId = $(if ($connection) { $connection.ConnectionId } else { $null }); How = $how }
    foreach ($cmd in $commands) {
        $found = Get-Command $cmd -ErrorAction SilentlyContinue
        $missing = if (-not $found) { $true } elseif ($For -eq 'Transfer') { -not ($found.Parameters.ContainsKey('EventId') -and $found.Parameters.ContainsKey('NewOrganizer')) } else { $false }
        if ($missing) {
            Disconnect-MclExchange
            if ($For -eq 'Transfer') { throw "Invoke-ChangeMeetingOrganizer (with -EventId and -NewOrganizer) is not available to the $how`: it needs a role with these parameters (developer guide, chapter 5: role 'Meeting Organizer Transfer' from User Options). A change of role takes up to an hour. Or -TransferMethod Recreate (Microsoft Graph only)." }
            throw "$cmd is not available to the $how`: it needs the role Mailbox Import Export (not in any role group by default). Developer guide, chapter 5: role group 'Meeting Cleanup Restore'. A change of role takes up to an hour."
        }
    }
    Write-MclItem Ok "Exchange Online PowerShell connected: $how" -Icon Server
}

function Disconnect-MclExchange {
    if ($script:Exchange) {
        try {
            if ($script:Exchange.ConnectionId) { Disconnect-ExchangeOnline -ConnectionId $script:Exchange.ConnectionId -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
            else { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }
        }
        catch { Write-MclLog 'WARN' "Exchange Online disconnection: $($_.Exception.Message)" }
        $script:Exchange = $null
    }
}

function ConvertFrom-MclExchangeTime {
    <# LastModifiedTime of Get-RecoverableItems: UTC, as text MM/dd/yyyy HH:mm:ss (or a DateTime). #>
    param($Value)
    if ($Value -is [datetime]) { return [datetime]::SpecifyKind($Value.ToUniversalTime(), [DateTimeKind]::Utc) }
    $inv = [Globalization.CultureInfo]::InvariantCulture
    $styles = [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact([string]$Value, [string[]]@('MM/dd/yyyy HH:mm:ss', 'M/d/yyyy h:mm:ss tt', 'yyyy-MM-ddTHH:mm:ss', 'o'), $inv, $styles, [ref]$d)) { return $d }
    if ([datetime]::TryParse([string]$Value, $inv, $styles, [ref]$d)) { return $d }
    return $null
}

function Get-MclPurgedItems {
    <# Calendar items in Recoverable Items\Purges of some mailboxes, removed between Start and End (UTC). #>
    param([Parameter(Mandatory = $true)][string[]]$Mailbox, [Parameter(Mandatory = $true)][datetime]$StartUtc, [Parameter(Mandatory = $true)][datetime]$EndUtc)

    $items = Get-RecoverableItems -Identity $Mailbox -SourceFolder PurgedItems -FilterItemType IPM.Appointment -ResultSize Unlimited `
        -FilterStartTime $StartUtc.ToLocalTime() -FilterEndTime $EndUtc.ToLocalTime() -ErrorAction Stop -WarningAction SilentlyContinue
    foreach ($i in @($items)) {
        if (-not $i) { continue }
        [pscustomobject]@{
            Mailbox         = ([string]$i.Identity).ToLowerInvariant()
            Subject         = [string]$i.Subject
            EntryID         = [string]$i.EntryID
            LastModifiedUtc = ConvertFrom-MclExchangeTime $i.LastModifiedTime
            LastParentPath  = [string]$i.LastParentPath
        }
    }
}

function Restore-MclPurgedItem {
    <# One item back to its folder. Returns @{ Ok; Folder; Error }. #>
    param([Parameter(Mandatory = $true)][string]$Mailbox, [Parameter(Mandatory = $true)][string]$EntryId)
    try {
        $r = Restore-RecoverableItems -Identity $Mailbox -EntryID $EntryId -SourceFolder PurgedItems -ErrorAction Stop -WarningAction SilentlyContinue | Select-Object -First 1
        $ok = [bool]($r -and $r.WasRestoredSuccessfully)
        return @{ Ok = $ok; Folder = [string]$(if ($r) { $r.RestoredToFolderPath }); Error = $(if ($ok) { '' } else { 'Restore-RecoverableItems did not restore the item' }) }
    }
    catch { return @{ Ok = $false; Folder = ''; Error = $_.Exception.Message } }
}

function Get-MclRestorePlan {
    <#
    .SYNOPSIS
        What a restore would do with a report of a Remove, Cancel or Transfer run, without doing it.
    .DESCRIPTION
        Restorable: the copies removed (Result Removed) of meetings not cancelled by their organizer. Not
        restorable: a meeting cancelled by its organizer (the attendees received the cancellation), a meeting
        transferred to a new organizer (it is there now), an occurrence of a series (Exchange does not keep a
        removed occurrence in Recoverable Items).
    #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result)

    $restore = [Collections.Generic.List[object]]::new()
    $cancelled = [Collections.Generic.List[object]]::new()
    $transferred = [Collections.Generic.List[object]]::new()
    $occurrences = [Collections.Generic.List[object]]::new()
    foreach ($m in @($Result.Meetings | Where-Object Selected)) {
        $removed = @($m.Copies | Where-Object { $_.EventId -and $_.Action -eq 'Remove' -and $_.Result -eq 'Removed' })
        if ([string](Get-MclProperty $m 'NewMeetingId') -or $m.Status -eq 'Transferred') { if ($removed.Count) { $transferred.Add($m) }; continue }
        if (@($m.Copies | Where-Object { $_.Action -eq 'Cancel' -and $_.Result -eq 'Cancelled' }).Count) { if ($removed.Count) { $cancelled.Add($m) }; continue }
        foreach ($c in $removed) { if ((Get-MclProperty $c 'Occurrence')) { $occurrences.Add($c) } else { $restore.Add($c) } }
    }
    $mailboxes = @($restore | ForEach-Object Mailbox | Select-Object -Unique).Count
    $lines = [Collections.Generic.List[string]]::new()
    if ($restore.Count) {
        $lines.Add(('{0} cop{1} to put back in {2} mailbox(es): {3} attendee(s), {4} room(s), from Recoverable Items (retention of deleted items: 14 days by default)' -f $restore.Count, $(if ($restore.Count -eq 1) { 'y' } else { 'ies' }), $mailboxes, @($restore | Where-Object Role -ne 'Room').Count, @($restore | Where-Object Role -eq 'Room').Count))
        $lines.Add('No message is sent: each copy comes back as it was, then is answered again silently (accepted, or tentative when there was no answer)')
    }
    else { $lines.Add('No copy removed silently by this run') }
    if ($cancelled.Count) { $lines.Add(('{0} meeting(s) cancelled by their organizer are not restorable: the attendees received the cancellation (the organizer has to send them again)' -f $cancelled.Count)) }
    if ($transferred.Count) { $lines.Add(('{0} meeting(s) transferred to a new organizer are not restorable: they are in the calendars again, with the new organizer' -f $transferred.Count)) }
    if ($occurrences.Count) { $lines.Add(('{0} occurrence(s) of series are not restorable: Exchange does not keep a removed occurrence in Recoverable Items (Backup.json lists them)' -f $occurrences.Count)) }
    [pscustomobject]@{ Restore = $restore.ToArray(); Cancelled = $cancelled.ToArray(); Transferred = $transferred.ToArray(); Occurrences = $occurrences.ToArray(); Mailboxes = $mailboxes; Lines = $lines.ToArray(); Text = ($lines -join " $($script:Dot) ") }
}

function Import-MclRestoreSource {
    <#
    .SYNOPSIS
        Reads the report of a Remove, Cancel or Transfer run for a restore: the copies removed, with the time of
        each removal (ActionUtc; reports of 1.0.0: the time of the run).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path, [string[]]$MeetingId)

    $file = if (Test-Path -LiteralPath $Path -PathType Container) { Get-ChildItem -LiteralPath $Path -Filter '*-Summary.json' -File | Select-Object -First 1 -ExpandProperty FullName } else { $Path }
    if (-not $file -or -not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "No *-Summary.json in ${Path}: give the folder of the report of a Remove, Cancel or Transfer run." }
    $data = [IO.File]::ReadAllText($file) | ConvertFrom-Json -Depth 32
    if ($data.Tool -ne 'Meeting Cleanup' -or -not $data.PSObject.Properties['Meetings']) { throw "$file is not a Meeting Cleanup report." }
    if ([string]$data.Action -notin 'Remove', 'Cancel', 'Transfer') { throw "$file is the report of a '$($data.Action)' run: give the report of a Remove, Cancel or Transfer run (folder MeetingCleanup_Remove_..., _Cancel_... or _Transfer_...)." }
    $ids = [Collections.Generic.HashSet[string]]::new([string[]]@($MeetingId | Where-Object { $_ } | ForEach-Object { ([string]$_).ToUpperInvariant() }), [StringComparer]::OrdinalIgnoreCase)
    $runStart = ConvertFrom-MclExchangeTime $data.StartedUtc
    $runEnd = ConvertFrom-MclExchangeTime $data.CompletedUtc
    $meetings = [Collections.Generic.List[object]]::new()
    # Every copy the run removed, of every meeting (also those not restored now): the restore matches the items
    # of Recoverable Items against all of them.
    $runRemoved = [Collections.Generic.List[object]]::new()
    foreach ($m in @($data.Meetings)) {
        $copies = [Collections.Generic.List[object]]::new()
        foreach ($c in @($m.Copies)) {
            foreach ($name in 'ShowAs', 'ActionUtc', 'Occurrence', 'OccurrenceStart', 'OccurrenceKey', 'SeriesId') { if (-not $c.PSObject.Properties[$name]) { $c | Add-Member -NotePropertyName $name -NotePropertyValue '' } }
            $c | Add-Member -NotePropertyName RemovedUtc -NotePropertyValue $(if ($c.ActionUtc) { ConvertFrom-MclExchangeTime $c.ActionUtc } else { $null }) -Force
            $c | Add-Member -NotePropertyName PreviousResult -NotePropertyValue ([string]$c.Result) -Force
            $c | Add-Member -NotePropertyName RestoredUtc -NotePropertyValue '' -Force
            # An occurrence is not in Recoverable Items: never counted among the items to match.
            if ($c.EventId -and [string]$c.Action -eq 'Remove' -and [string]$c.Result -eq 'Removed' -and -not $c.Occurrence) { $runRemoved.Add($c) }
            $copies.Add($c)
        }
        if ($ids.Count -and -not $ids.Contains([string]$m.MeetingId)) { continue }
        if (-not $m.Selected) { continue }
        $notes = [Collections.Generic.List[string]]::new()
        foreach ($n in @($m.Notes)) { if ($n) { $notes.Add([string]$n) } }
        $m.Copies = $copies
        $m.Notes = $notes
        Add-MclMeetingDefaults $m
        $meetings.Add($m)
    }
    if ($ids.Count -and $meetings.Count -lt $ids.Count) { Write-MclItem Warn ('{0} meeting ID(s) not in the report.' -f ($ids.Count - $meetings.Count)) }
    $result = [pscustomobject]@{
        Tool = 'Meeting Cleanup'; Version = $script:ToolVersion; Action = 'Restore'; Status = 'Completed'; Error = ''
        StartedUtc = [datetime]::UtcNow.ToString('o'); CompletedUtc = ''; DurationSeconds = 0.0
        Request = $data.Request; Tenant = $data.Tenant; Organization = $data.Organization; AppId = $data.AppId; AppName = $data.AppName
        Organizers = @($data.Organizers); Searched = $data.Searched; Meetings = $meetings; Warnings = [Collections.Generic.List[string]]::new(); Counts = $null
        FromReport = $file; SourceAction = [string]$data.Action; RunRemoved = $runRemoved.ToArray()
        SourceRunUtc = [pscustomobject]@{ Start = $runStart; End = $runEnd }
    }
    Update-MclResultCounts $result
    return $result
}

function Get-MclRestoreKey {
    <# The group of a copy for the restore: its mailbox and its subject (a room shows the organizer's name). #>
    param($Copy)
    '{0}|{1}' -f $Copy.Mailbox, ([string]$Copy.Subject).Trim().ToLowerInvariant()
}

function Invoke-MclRestore {
    <#
    .SYNOPSIS
        Puts back the copies removed by a run (Import-MclRestoreSource), checks them and answers again silently.
        Needs Connect-MclGraph and Connect-MclExchange first.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Settings, [Parameter(Mandatory = $true)][pscustomobject]$Result)

    $started = [datetime]::UtcNow
    $dot = $script:Dot
    $plan = Get-MclRestorePlan -Result $Result
    $window = [TimeSpan]::FromMinutes([int]$Settings.RestoreWindowMinutes)
    $immutable = @{ Prefer = 'IdType="ImmutableId"' }
    foreach ($m in $Result.Meetings) {
        foreach ($c in $m.Copies) {
            if ($plan.Restore -contains $c) { $c.Action = 'Restore'; $c.Result = ''; $c.Detail = ''; $c.Verified = ''; $c.HttpStatus = 0 }
            else {
                $c.Action = ''
                if ($c.PreviousResult) { $c.Detail = "not restored: $($c.PreviousResult.ToLowerInvariant()) by the run" }
                $c.Result = ''
            }
        }
    }
    foreach ($m in $plan.Cancelled) {
        $m.Status = 'Not restorable'
        $m.Notes.Add('Cancelled by its organizer: the attendees received the cancellation. The organizer has to send the meeting again.')
        foreach ($c in @($m.Copies | Where-Object { $_.PreviousResult -eq 'Removed' })) { $c.Result = 'Not restorable'; $c.Detail = 'meeting cancelled by its organizer' }
    }
    foreach ($m in $plan.Transferred) {
        $m.Status = 'Not restorable'
        $m.Notes.Add("Transferred to $($m.NewOrganizer): the meeting is in the calendars with its new organizer.")
        foreach ($c in @($m.Copies | Where-Object { $_.PreviousResult -eq 'Removed' })) { $c.Result = 'Not restorable'; $c.Detail = "meeting transferred to $($m.NewOrganizer)" }
    }
    foreach ($c in $plan.Occurrences) { $c.Result = 'Not restorable'; $c.Detail = 'occurrence of a series: not kept in Recoverable Items (see Backup.json)' }
    foreach ($m in @($Result.Meetings | Where-Object { $_.Status -ne 'Not restorable' -and @($_.Copies | Where-Object Result -eq 'Not restorable').Count -and -not @($_.Copies | Where-Object { $plan.Restore -contains $_ }).Count })) { $m.Status = 'Not restorable' }

    Write-MclNextStep 'Restore' 'Refresh'
    foreach ($line in $plan.Lines) { Write-MclItem Info $line }
    $todo = @($plan.Restore)
    if (-not $todo.Count) { Write-MclItem Skip 'No copy to restore.' }

    # ---- 1. which copies of the run are back in their calendar ------------------------------------------
    # Every copy the run removed with the same subject in the same mailbox (selected or not, of a cancelled
    # meeting too) is looked up: the items of Recoverable Items are matched against the copies that are not back.
    $keys = [Collections.Generic.HashSet[string]]::new([string[]]@($todo | ForEach-Object { Get-MclRestoreKey $_ }))
    $members = @($Result.RunRemoved | Where-Object { $keys.Contains((Get-MclRestoreKey $_)) })
    $todoSet = [Collections.Generic.HashSet[object]]::new([Collections.Generic.ReferenceEqualityComparer]::Instance)
    foreach ($c in $todo) { [void]$todoSet.Add($c) }
    $membersByKey = @{}
    foreach ($c in $members) {
        $key = Get-MclRestoreKey $c
        if (-not $membersByKey.ContainsKey($key)) { $membersByKey[$key] = [Collections.Generic.List[object]]::new() }
        $membersByKey[$key].Add($c)
    }
    # By reference: every [pscustomobject] has the same hash code, a hashtable would be slow with many copies.
    $present = [Collections.Generic.Dictionary[object, object]]::new([Collections.Generic.ReferenceEqualityComparer]::Instance)
    $presentEvent = [Collections.Generic.Dictionary[object, object]]::new([Collections.Generic.ReferenceEqualityComparer]::Instance)
    if ($members.Count) {
        $requests = for ($i = 0; $i -lt $members.Count; $i++) {
            $f = [Uri]::EscapeDataString("iCalUId eq '$($members[$i].MeetingId)'")
            New-MclGraphRequest -Id "p$i" -Url "$(Get-MclUserPath $members[$i].Mailbox)/events?`$filter=$f&`$select=id,showAs,responseStatus" -Headers $immutable
        }
        $res = Invoke-MclGraphBatch -Requests @($requests)
        for ($i = 0; $i -lt $members.Count; $i++) {
            $r = $res["p$i"]
            $present[$members[$i]] = if ($r.Status -eq 200) { [bool]$r.Values.Count } else { $null }
            if ($r.Status -eq 200 -and $r.Values.Count) { $presentEvent[$members[$i]] = $r.Values | Select-Object -First 1 }
            if ($todoSet.Contains($members[$i]) -and $r.Status -ne 200) { $members[$i].Result = 'Failed'; $members[$i].Detail = "mailbox not readable: $(Get-MclMailboxProblem $r)" }
        }
    }
    $pending = [Collections.Generic.List[object]]::new()
    # A copy already back but still Declined / Free as Exchange left it at the removal (a restore stopped, or not
    # verified): it is answered again with the copies restored now. Not a copy the attendee had declined.
    $stale = [Collections.Generic.List[object]]::new()
    foreach ($c in $todo) {
        if ($c.Result) { continue }
        if ($present.ContainsKey($c) -and $present[$c]) {
            $c.Result = 'Already present'; $c.Detail = 'already in the calendar: nothing to restore'
            $ev = if ($presentEvent.ContainsKey($c)) { $presentEvent[$c] } else { $null }
            if ($Settings.RestoreReAccept -and $ev -and [string](Get-MclProperty (Get-MclProperty $ev 'responseStatus') 'response') -eq 'declined' -and [string](Get-MclProperty $ev 'showAs') -eq 'free' -and [string]$c.Response -ne 'declined') {
                $c.EventId = [string]$ev.id; $c.Verified = 'Yes'; $c.Detail = 'already in the calendar, still Declined / Free from the removal'
                $stale.Add($c)
            }
        }
        else { $pending.Add($c) }
    }

    # ---- 2. each mailbox: what is in Purges, which item is which copy, restore -----------------------------
    # A copy is restored only when its item is certain: the items with its subject removed around the time of
    # the run are exactly the copies of the run that are not back, in the order of the removals (Remove leaves
    # 3 seconds between the copies with the same subject in one mailbox). Otherwise nothing is restored for that
    # subject in that mailbox (Failed, with the command to do it by hand): a wrong item is never put back.
    $tolerance = [TimeSpan]::FromMinutes(2)
    $todoCount = $pending.Count
    $done = 0
    $restoredItems = [Collections.Generic.List[object]]::new()
    $groups = @($pending | Group-Object Mailbox)
    $chunks = [Math]::Ceiling($groups.Count / 10)
    for ($k = 0; $k -lt $chunks; $k++) {
        Assert-MclNotCancelled
        $chunk = @($groups | Select-Object -Skip ($k * 10) -First 10)
        $mailboxes = @($chunk | ForEach-Object Name)
        # The copies of the run expected in Purges for the subjects of this chunk (to know the window to read).
        $chunkKeys = @($chunk | ForEach-Object Group | ForEach-Object { Get-MclRestoreKey $_ } | Select-Object -Unique)
        $absent = @($chunkKeys | ForEach-Object { $membersByKey[$_] } | Where-Object { $present.ContainsKey($_) -and $present[$_] -eq $false })
        $times = @($absent | ForEach-Object { if ($_.RemovedUtc) { $_.RemovedUtc } })
        $from = if ($times.Count) { ($times | Measure-Object -Minimum).Minimum } else { $Result.SourceRunUtc.Start }
        $to = if ($times.Count) { ($times | Measure-Object -Maximum).Maximum } else { $Result.SourceRunUtc.End }
        $copies = @($chunk | ForEach-Object Group)
        if (-not $from -or -not $to) { foreach ($c in $copies) { $c.Result = 'Failed'; $c.Detail = 'time of the removal unknown in the report' }; continue }
        $candidates = @()
        try { $candidates = @(Get-MclPurgedItems -Mailbox $mailboxes -StartUtc ($from - $window) -EndUtc ($to + $window)) }
        catch { foreach ($c in $copies) { $c.Result = 'Failed'; $c.Detail = "Get-RecoverableItems: $($_.Exception.Message)" }; continue }
        foreach ($group in ($copies | Group-Object { Get-MclRestoreKey $_ })) {
            $pend = @($group.Group)
            $subject = ([string]$pend[0].Subject).Trim()
            $mailbox = $pend[0].Mailbox
            $all = @($membersByKey[$group.Name])
            $unknown = @($all | Where-Object { -not $present.ContainsKey($_) -or $null -eq $present[$_] })
            $want = @($all | Where-Object { $present.ContainsKey($_) -and $present[$_] -eq $false } | Sort-Object { if ($_.RemovedUtc) { $_.RemovedUtc } else { [datetime]::MinValue } })
            $known = @($want | Where-Object RemovedUtc)
            $lo = if ($want.Count -and $known.Count -eq $want.Count) { ($known | ForEach-Object RemovedUtc | Measure-Object -Minimum).Minimum } else { $Result.SourceRunUtc.Start }
            $hi = if ($want.Count -and $known.Count -eq $want.Count) { ($known | ForEach-Object RemovedUtc | Measure-Object -Maximum).Maximum } else { $Result.SourceRunUtc.End }
            $found = @($candidates | Where-Object { $_.Mailbox -eq $mailbox -and ([string]$_.Subject).Trim() -eq $subject -and $_.LastModifiedUtc -ge ($lo - $tolerance) -and $_.LastModifiedUtc -le ($hi + $tolerance) } | Sort-Object LastModifiedUtc, EntryID)
            if (-not $found.Count) {
                foreach ($c in $pend) { $c.Result = 'Not found'; $c.Detail = 'not in Recoverable Items: retention of deleted items over, restored before, or removed by someone else' }
                continue
            }
            # Ranks are certain when every removal time is known and no two items or copies share a second.
            $distinct = { param($values) $list = @($values); @($list | Select-Object -Unique).Count -eq $list.Count }
            $ordered = $known.Count -eq $want.Count -and (& $distinct @($found | ForEach-Object { $_.LastModifiedUtc.ToString('yyyyMMddHHmmss') })) -and (& $distinct @($want | ForEach-Object { $_.RemovedUtc.ToString('yyyyMMddHHmmss') }))
            $whole = $pend.Count -eq $want.Count
            if ($unknown.Count -or $found.Count -ne $want.Count -or -not ($ordered -or $whole)) {
                $why = if ($unknown.Count) { "$($unknown.Count) other cop$(if ($unknown.Count -eq 1) { 'y' } else { 'ies' }) with this subject could not be checked" }
                       else { '{0} item(s) in Recoverable Items for {1} cop{2} removed by the run' -f $found.Count, $want.Count, $(if ($want.Count -eq 1) { 'y' } else { 'ies' }) }
                foreach ($c in $pend) {
                    $c.Result = 'Failed'
                    $c.Detail = "ambiguous, nothing restored ($why). By hand: Get-RecoverableItems -Identity $mailbox -SourceFolder PurgedItems -FilterItemType IPM.Appointment -SubjectContains '$($subject.Replace("'", "''"))', then Restore-RecoverableItems -EntryID <the item>"
                }
                continue
            }
            for ($j = 0; $j -lt $want.Count; $j++) {
                $c = $want[$j]
                if ($pend -notcontains $c) { continue }
                $r = Restore-MclPurgedItem -Mailbox $c.Mailbox -EntryId $found[$j].EntryID
                $c.RestoredUtc = [datetime]::UtcNow.ToString('o')
                if ($r.Ok) { $c.Result = 'Restored'; $c.Detail = "back in $($r.Folder)"; $restoredItems.Add($c) }
                else { $c.Result = 'Failed'; $c.Detail = "Restore-RecoverableItems: $($r.Error)" }
                $done++
                Write-MclProgress ($done / [Math]::Max(1, $todoCount)) ('{0:N0}/{1:N0} copies restored' -f $done, $todoCount)
            }
        }
    }
    if ($todo.Count) {
        Write-MclItem $(if (@($todo | Where-Object { $_.Result -notin 'Restored', 'Already present' }).Count) { 'Warn' } else { 'Ok' }) ('{0} cop{1} put back {2} {3} already present {2} {4} not found in Recoverable Items {2} {5} failed' -f $restoredItems.Count, $(if ($restoredItems.Count -eq 1) { 'y' } else { 'ies' }), $dot, @($todo | Where-Object Result -eq 'Already present').Count, @($todo | Where-Object Result -eq 'Not found').Count, @($todo | Where-Object Result -eq 'Failed').Count) -Icon Refresh
    }

    # ---- 3. check: each copy back in its calendar (nothing is ever removed by a restore) -------------------
    $back = [Collections.Generic.List[object]]::new()
    if ($restoredItems.Count -or $stale.Count) { Write-MclNextStep 'Verify and answer again' 'Search' }
    if ($restoredItems.Count) {
        $requests = for ($i = 0; $i -lt $restoredItems.Count; $i++) {
            $f = [Uri]::EscapeDataString("iCalUId eq '$($restoredItems[$i].MeetingId)'")
            New-MclGraphRequest -Id "v$i" -Url "$(Get-MclUserPath $restoredItems[$i].Mailbox)/events?`$filter=$f&`$select=id,responseStatus,showAs" -Headers $immutable
        }
        $res = Invoke-MclGraphBatch -Requests @($requests)
        $wrong = 0; $unchecked = 0
        for ($i = 0; $i -lt $restoredItems.Count; $i++) {
            $c = $restoredItems[$i]; $r = $res["v$i"]
            if ($r.Status -ne 200) { $unchecked++; $c.Verified = "Unknown ($($r.Status))"; $c.Detail = "$($c.Detail), not verified ($(Get-MclMailboxProblem $r)) nor answered again: it may show Declined / Free"; continue }
            $ev = $r.Values | Select-Object -First 1
            if ($ev) { $c.Verified = 'Yes'; $c.EventId = [string]$ev.id; $back.Add($c) }
            else { $wrong++; $c.Verified = 'No'; $c.Result = 'Failed'; $c.Detail = 'an item came back, but not this meeting: it is left in the calendar (Declined / Free), check it' }
        }
        if ($wrong) { $text = "$wrong item(s) put back are not the meeting expected: left in the calendar (Declined / Free), see the report."; $Result.Warnings.Add($text); Write-MclItem Warn $text }
        if ($unchecked) { $text = "$unchecked cop$(if ($unchecked -eq 1) { 'y' } else { 'ies' }) put back but not verified, nor answered again (they may show Declined / Free): the same restore again finishes them."; $Result.Warnings.Add($text) }
        Write-MclItem $(if ($back.Count -lt $restoredItems.Count) { 'Warn' } else { 'Ok' }) ('{0} of {1} verified back in their calendar{2}' -f $back.Count, $restoredItems.Count, $(if ($unchecked) { " $dot $unchecked not checked" } else { '' }))
    }
    if ($stale.Count) { Write-MclItem Info ('{0} cop{1} already back but still Declined / Free from the removal: answered again too' -f $stale.Count, $(if ($stale.Count -eq 1) { 'y' } else { 'ies' })) }

    # ---- 4. answer again, silently: the room busy again, the organizer's tracking corrected -----------------
    $toAnswer = @($back) + @($stale)
    if ($Settings.RestoreReAccept -and $toAnswer.Count) {
        $answers = @($toAnswer | ForEach-Object {
                $verb = switch ([string]$_.Response) { 'accepted' { 'accept' } 'organizer' { '' } 'declined' { '' } default { 'tentativelyAccept' } }
                if ($_.Role -eq 'Room' -and [string]$_.Response -in '', 'none', 'notResponded') { $verb = 'accept' }
                if ($verb) { [pscustomobject]@{ Copy = $_; Verb = $verb } }
            })
        if ($answers.Count) {
            $requests = for ($i = 0; $i -lt $answers.Count; $i++) {
                New-MclGraphRequest -Id "r$i" -Method 'POST' -Url "$(Get-MclUserPath $answers[$i].Copy.Mailbox)/events/$([Uri]::EscapeDataString($answers[$i].Copy.EventId))/$($answers[$i].Verb)" -Body @{ sendResponse = $false } -Headers $immutable
            }
            $res = Invoke-MclGraphBatch -Requests @($requests)
            $okAnswers = 0
            for ($i = 0; $i -lt $answers.Count; $i++) {
                $c = $answers[$i].Copy; $r = $res["r$i"]
                $word = if ($answers[$i].Verb -eq 'accept') { 'accepted' } else { 'tentative' }
                if ($r.Status -in 200, 202) { $okAnswers++; $c.Detail = "$($c.Detail), answered $word again (no message)" }
                else { $c.Detail = "$($c.Detail), not answered again ($($r.Status) $($r.ErrorCode)): it shows Declined / Free" }
            }
            if ($okAnswers -lt $answers.Count) { $missed = $answers.Count - $okAnswers; $Result.Warnings.Add("$missed cop$(if ($missed -eq 1) { 'y' } else { 'ies' }) not answered again (they show Declined / Free): the same restore again finishes them.") }
            Write-MclItem $(if ($okAnswers -lt $answers.Count) { 'Warn' } else { 'Ok' }) ('{0} of {1} answered again without a message: rooms busy again, the organizer''s tracking corrected' -f $okAnswers, $answers.Count)
        }
    }
    elseif ($back.Count) { foreach ($c in $back) { $c.Detail = "$($c.Detail), left Declined / Free (Restore.ReAccept is `$false)" } }
    # ---- status ------------------------------------------------------------------------------------------
    foreach ($m in @($Result.Meetings | Where-Object { $_.Status -ne 'Not restorable' })) {
        $acted = @($m.Copies | Where-Object Action -eq 'Restore')
        if (-not $acted.Count) { $m.Status = 'Nothing to do'; continue }
        $ok = @($acted | Where-Object { $_.Result -in 'Restored', 'Already present' })
        $m.Status = if ($ok.Count -eq $acted.Count) { 'Restored' } elseif ($ok.Count) { 'Partial' } else { 'Failed' }
    }
    Update-MclResultCounts $Result
    $all = @($Result.Meetings)
    $Result.Status = if ($all.Count -and -not @($all | Where-Object { $_.Status -ne 'Failed' }).Count) { 'Failed' }
        elseif (@($all | Where-Object { $_.Status -in 'Partial', 'Failed', 'Not restorable' }).Count -or $Result.Warnings.Count) { 'Warning' } else { 'Completed' }
    $Result.CompletedUtc = [datetime]::UtcNow.ToString('o')
    $Result.DurationSeconds = [Math]::Round(([datetime]::UtcNow - $started).TotalSeconds, 1)
    return $Result
}

# SIG # Begin signature block
# MIImIwYJKoZIhvcNAQcCoIImFDCCJhACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBhVezS+GU7OZG9
# L/lD5bEjg8IT66S2Q/EeOTZqYkXZKKCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYCjCCGAYCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCYtfEBOtNg
# udeXn2/7DGoeS1DQTH0Mhplx6pi+pMkThDANBgkqhkiG9w0BAQEFAASCAgCeir4c
# zq9TOZyY3mhs3fIpndMdt03eHW9092m2LM5l0bxyYJ4jZc80eRHZGkO8ZmcgW8SV
# DPFfwUKg8NfTpWNTcZCL3d995i2sLiv1tkkt+46nbK9IG05vERBuWjpS+tk8M52a
# SnlhNuRQe+Ey3vkqfeA8l5k9adkDPt4qh1Xd5FhquACTHK36+zaYKLZEXba0Et2q
# 0zQ+/gxrYo53MVvKSowQqA273nmvLWcIKJSgx2MqHOyOOlRj9/PhjOPRvQz7vkim
# DDua08nzXEXQRhUtfkF8RS93G5rN6IHiUmlVwrMpcnZXmhZ83jFNeEw3P3F0lvA0
# Rd0Ug8LIKfVcTWcklrrvD7iVHrmPIcCNC4wYULv9qzlpNvpCGUtBWWwHQoHml7jN
# lW8yBWZYY9bfiorFngtnySsHTkh2GuJVDhlqUejZkrT6ypk8Yv3JVQ6DCNAl0C+1
# yqhS8+nb9S+njNNv1DmrMlj4xOLUgdZhbw0LwJ6ttW0L6mdq7/sUodkB+i9bK9PS
# Ku038Wm0XWOyKEAld6fzwHmB1npXzW8i2/+71sZ0KePCHlLZMustAKfXlus8ewZD
# OjBkLjZxI/zTAQWdzozZyPK4UJjH8NH6KdwKQ+JneAd+WCpYYoCvtA5X+c+4uLol
# wBfdrD11Gloq8U5j4U2pMAdqYKZluaXnl2VAh6GCFMcwghTDBgorBgEEAYI3AwMB
# MYIUszCCFK8GCSqGSIb3DQEHAqCCFKAwghScAgEDMQ0wCwYJYIZIAWUDBAIBMIGB
# BgsqhkiG9w0BCRABBKByBHAwbgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIPTpe+aLqpfjGvtqDGqqik6x/9NH24QPzrJoz6NJ5I7UAghMKiGn1nBa
# dxgPMjAyNjExMDYwODQyMjVaMAMCAQECCEKKUuAFFBM5oIIRpzCCBPMwggLboAMC
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
# 9sdantP6nYRQFTpF8czGc+FGresbAH819Pv5vBhCluTo8Px2PpC71cM0jzsxggJX
# MIICUwIBATCBhjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29t
# IFRpbWVzdGFtcGluZyBJc3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7j
# MAsGCWCGSAFlAwQCAaCCAWAwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMBwG
# CSqGSIb3DQEJBTEPFw0yNjExMDYwODQyMjVaMCgGCSqGSIb3DQEJNDEbMBkwCwYJ
# YIZIAWUDBAIBoQoGCCqGSM49BAMCMC8GCSqGSIb3DQEJBDEiBCBHCr4ZhUHmq80k
# AAIbJcdDKDXFWj043mIiuP33IYPcMTCByAYLKoZIhvcNAQkQAi8xgbgwgbUwgbIw
# ga8EIFdP6heq+l58aiubCMmYsXNeRQs4c8pE5qJMoJMOY/OwMIGKMHakdDByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7jMAoGCCqGSM49BAMCBEgw
# RgIhAMSsMs9+Ii7V0PQIEFhVrWCzzxEYP+mGouP3RhYhQ8KMAiEAwF1IDcMIbmN2
# +p+GYKLCW/BvoBmybUC67Prxseumn5Q=
# SIG # End signature block
