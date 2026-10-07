<#
.SYNOPSIS
    Meeting Cleanup - finds the meetings of one or many organizers, or of rooms, in Exchange Online and removes
    them from the attendees and the rooms, cancels them, or transfers them to a new organizer; restores what was
    removed.

.DESCRIPTION
    One tool for every case:
      - one meeting (-Subject or -MeetingId), a series (handled as a whole) or every meeting of a period;
      - one organizer, several, or a list in a file (-OrganizerFile);
      - or rooms (-Room, -RoomFile): every meeting of these rooms in the period, whatever its organizer; a series
        is then limited to its occurrences in the period (a room closed for two weeks does not end a series);
      - organizer mailbox present (its calendar is searched) or deleted (the meetings are searched in the
        rooms, a list of mailboxes or every mailbox, by the organizer's address);
      - every copy is then looked up in the calendar of each internal attendee, room and member of an invited
        group, wherever the meeting was found.

    Actions:
      Report   (default) lists the meetings and their copies: nothing is changed.
      Remove   removes the copies of the attendees and the rooms without any message. The meeting stays in the
               organizer's calendar when the mailbox exists: Exchange sends a cancellation to the attendees
               whenever a meeting is removed there, so that is done by Cancel only.
      Cancel   the organizer cancels the meeting (message to every attendee, rooms released), then the copies
               left in the attendees' and rooms' calendars are removed.
      Transfer the meetings move to -NewOrganizer: moved by Exchange Online (Invoke-ChangeMeetingOrganizer) when
               the organizer's mailbox exists, re-created in the new organizer's calendar when it does not (the
               attendees receive one invitation from him); -TransferMethod Auto, Native or Recreate.
      Restore  with -FromReport <report of a Remove, Cancel or Transfer run>: the copies removed come back from
               Recoverable Items (retention of deleted items, 14 days by default), then are answered again
               silently. Needs Exchange Online PowerShell and the role Mailbox Import Export (developer guide, chapter 5).
    Remove, Cancel and Transfer write a backup of the meetings first, show what they will do and ask to type YES
    (-Force skips it). -FromReport replays the meetings of a reviewed report, without searching again.

    Microsoft Graph, application permissions (user guide, chapter 1). Writes CSV, JSON and HTML report files in a
    new folder, and a daily log file. Everything is set in config\MeetingCleanup.config.psd1; the parameters
    below override it.

.PARAMETER Organizer
    The organizer: SMTP address (any alias), or the X500 address (legacyExchangeDN) of a deleted mailbox.
    Several organizers may be given.

.PARAMETER OrganizerFile
    A list of organizers: one address per line (SMTP or X500, # = comment), or a CSV file with a column
    PrimarySmtpAddress, EmailAddress, Mail, UserPrincipalName, Organizer or LegacyExchangeDN.

.PARAMETER Room
    Rooms (or equipment) mode: every meeting of these rooms in the period, whatever its organizer. With an
    action, the period must be given (-Start and -End).

.PARAMETER RoomFile
    A list of rooms: one address per line, or a CSV file (PrimarySmtpAddress, Mail...).

.PARAMETER Start
    Start of the period (date, or date and time), in the time zone of Report.TimeZone. Default: today minus
    Search.PastDays.

.PARAMETER End
    End of the period; a date without a time is included. Default: today plus Search.FutureDays.

.PARAMETER Subject
    Only the meetings whose subject contains this text (* and ? are wildcards).

.PARAMETER MeetingId
    Only these meetings: column MeetingId of a report (iCalUId, the same in every copy of a meeting).

.PARAMETER SearchIn
    Where to search: Organizer, Rooms, Mailboxes, AllMailboxes (one or more). Default: Search.SearchIn.

.PARAMETER Mailbox
    Mailboxes of the 'Mailboxes' scope.

.PARAMETER MailboxFile
    File of the 'Mailboxes' scope: one address per line, or a CSV file (PrimarySmtpAddress, Mail...).

.PARAMETER Action
    Report (default), Remove, Cancel, Transfer or Restore.

.PARAMETER Comment
    Cancel: the message of the cancellation. Default: Cleanup.CancelComment.
    Transfer: the message of the old organizer when he still has a mailbox and the meeting is re-created
    ({0} = the new organizer). Default: Transfer.Comment.

.PARAMETER NewOrganizer
    Transfer: the new organizer, SMTP address of a mailbox of the tenant.

.PARAMETER TransferMethod
    Transfer: Auto (default, Transfer.Method), Native (Exchange Online only) or Recreate (Microsoft Graph only).

.PARAMETER TransferFrom
    Transfer: the date the series move from (default now): the occurrences before stay with the old organizer.

.PARAMETER FromReport
    Folder (or Summary.json) of a report. With Remove or Cancel: act on exactly its meetings. With Restore: the
    report of the Remove (or Cancel) run to undo. With -MeetingId: some of its meetings only.

.PARAMETER Force
    Remove, Cancel or Restore without asking to type YES (scheduled task, script).

.PARAMETER Gui
    Opens the window: same search, same actions, same report, and the restore.

.PARAMETER TenantId
    Overrides Tenant.TenantId.

.PARAMETER AppId
    Overrides Authentication.AppId.

.PARAMETER CertificateThumbprint
    Overrides Authentication.CertificateThumbprint.

.PARAMETER OutputPath
    Overrides Report.OutputPath.

.PARAMETER NoReport
    No CSV or HTML file. An action still writes its backup and its Summary.json (needed by the restore).

.PARAMETER ConfigPath
    Configuration file. Default: config\MeetingCleanup.config.psd1.

.EXAMPLE
    .\Invoke-MeetingCleanup.ps1 -Organizer john.doe@contoso.com
    The meetings of John Doe for the coming year, in his calendar and in the rooms: report only.

.EXAMPLE
    .\Invoke-MeetingCleanup.ps1 -OrganizerFile .\leavers.csv -SearchIn Organizer, Rooms, AllMailboxes
    The meetings of every person of the list (column PrimarySmtpAddress), present or deleted: report only.

.EXAMPLE
    .\Invoke-MeetingCleanup.ps1 -Organizer john.doe@contoso.com -Action Cancel -Comment 'John Doe has left the company.'
    John Doe's meetings are cancelled by his mailbox (message to the attendees, rooms released), then the
    copies left in the attendees' calendars are removed.

.EXAMPLE
    .\Invoke-MeetingCleanup.ps1 -Organizer john.doe@contoso.com -Subject 'Weekly review' -Action Remove
    One meeting (or series): removed from the attendees and the rooms, without any message.

.EXAMPLE
    .\Invoke-MeetingCleanup.ps1 -Room room-paris-01@contoso.com, room-paris-02@contoso.com -Start 2026-11-02 -End 2026-11-13 -Action Cancel -Comment 'The rooms of the 1st floor are closed for works.'
    Every meeting of the two rooms between 2 and 13 November is cancelled by its organizer (an occurrence for a
    series); the meetings whose organizer has left are removed from the calendars, without a message.

.EXAMPLE
    .\Invoke-MeetingCleanup.ps1 -Organizer john.doe@contoso.com -SearchIn Rooms, AllMailboxes -Action Transfer -NewOrganizer jane.roe@contoso.com
    John Doe has left (mailbox deleted): his meetings to come are re-created by Jane Roe, who sends the invitations.

.EXAMPLE
    .\Invoke-MeetingCleanup.ps1 -Action Restore -FromReport .\reports\MeetingCleanup_Remove_20261005-201500
    Undoes that Remove: the copies come back in the calendars of the attendees and the rooms.

.EXAMPLE
    .\Invoke-MeetingCleanup.ps1 -Gui

.NOTES
    Author  : Nicolas Fabert
    Version : 1.2.2
    Exit codes : 0 = completed, 1 = failed, 2 = finished with warnings (a copy not removed or restored, a mailbox not read...).
    Documentation : docs\MeetingCleanup-UserGuide.html (user guide: prerequisites, everyday commands) and
                    docs\MeetingCleanup-Guide.html (developer guide); sources: docs\*.md
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [string[]]$Organizer,
    [string]$OrganizerFile,
    [string[]]$Room,
    [string]$RoomFile,
    [datetime]$Start,
    [datetime]$End,
    [string]$Subject,
    [string[]]$MeetingId,
    [ValidateSet('Organizer', 'Rooms', 'Mailboxes', 'AllMailboxes')]
    [string[]]$SearchIn,
    [string[]]$Mailbox,
    [string]$MailboxFile,
    [ValidateSet('Report', 'Remove', 'Cancel', 'Transfer', 'Restore')]
    [string]$Action = 'Report',
    [string]$Comment,
    [string]$NewOrganizer,
    [ValidateSet('Auto', 'Native', 'Recreate')]
    [string]$TransferMethod,
    [datetime]$TransferFrom,
    [string]$FromReport,
    [switch]$Force,
    [switch]$Gui,
    [string]$TenantId,
    [string]$AppId,
    [string]$CertificateThumbprint,
    [string]$OutputPath,
    [switch]$NoReport,
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config\MeetingCleanup.config.psd1')
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$previousCulture = [Threading.Thread]::CurrentThread.CurrentCulture
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$exitCode = 1
$moduleLoaded = $false

function Confirm-MclAction {
    <# Shows the plan and asks to type YES; throws when the answer is not YES or the console cannot ask. #>
    param([string]$Title, [string[]]$Lines)
    Write-Host ''
    Write-Host "  $Title - $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Yellow
    foreach ($line in $Lines) { Write-Host "    - $line" -ForegroundColor Yellow }
    if ([Console]::IsInputRedirected -or -not [Environment]::UserInteractive) { throw 'Confirmation needed: run interactively, or add -Force.' }
    $answer = Read-Host '  Type YES to continue'
    if ($answer -cne 'YES') { throw 'Cancelled: nothing was changed.' }
    Write-MclLog 'INFO' "Confirmed by $([Environment]::UserName): $($Lines -join ' | ')"
}

try {
    Import-Module (Join-Path $PSScriptRoot 'MeetingCleanup.psd1') -Force
    $moduleLoaded = $true

    # ---- configuration, then command-line overrides ---------------------------------------------------
    $settings = Import-MclConfiguration -Path $ConfigPath -Root $PSScriptRoot
    if ($TenantId) { $settings.TenantId = $TenantId }
    if ($AppId) { $settings.AppId = $AppId }
    if ($CertificateThumbprint) { $settings.CertificateThumbprint = $CertificateThumbprint; $settings.AuthMode = 'Certificate' }
    if ($OutputPath) { $settings.OutputPath = [IO.Path]::GetFullPath($OutputPath, (Get-Location).Path) }
    $check = Test-MclConfiguration -Configuration $settings
    if (-not $check.IsValid) { throw ("Invalid value:`n - " + ($check.Problems -join "`n - ")) }

    $logPath = Start-MclLog -Directory $settings.LogPath -RetentionDays $settings.LogRetentionDays

    if ($Gui) {
        Write-MclLog 'STEP' '=== Meeting Cleanup - window opened ==='
        Show-MclGui -Configuration $settings
        $exitCode = 0
    }
    else {
        $requestArgs = @{ Settings = $settings; Organizer = $Organizer; OrganizerFile = $OrganizerFile; Room = $Room; RoomFile = $RoomFile; Subject = $Subject; MeetingId = $MeetingId; SearchIn = $SearchIn; Mailbox = $Mailbox; MailboxFile = $MailboxFile; Action = $Action; NewOrganizer = $NewOrganizer; TransferMethod = $TransferMethod }
        if ($PSBoundParameters.ContainsKey('TransferFrom')) { $requestArgs.TransferFrom = $TransferFrom }
        if ($PSBoundParameters.ContainsKey('Start')) { $requestArgs.Start = $Start }
        if ($PSBoundParameters.ContainsKey('End')) { $requestArgs.End = $End }
        if ($PSBoundParameters.ContainsKey('Comment')) { $requestArgs.Comment = $Comment }
        if ($FromReport) { $requestArgs.FromReport = [IO.Path]::GetFullPath($FromReport, (Get-Location).Path) }
        $request = New-MclRequest @requestArgs
        $requestCheck = Test-MclRequest -Request $request
        if (-not $requestCheck.IsValid) { throw ("Cannot run:`n - " + ($requestCheck.Problems -join "`n - ")) }

        Write-MclRunBanner -Settings $settings -Request $request -LogPath $logPath -NoReport:$NoReport
        $steps = switch ($request.Action) {
            'Restore' { 6 }
            'Report' { 6 }
            'Transfer' { $(if ($request.FromReport) { 3 } else { 6 }) + 2 + [int][bool]$settings.Verify }
            default { $(if ($request.FromReport) { 3 } else { 6 }) + 1 + [int][bool]$settings.Verify }
        }
        Initialize-MclSteps -Total $steps

        # ---- Microsoft Graph --------------------------------------------------------------------------
        Write-MclNextStep 'Microsoft Graph' 'Key'
        $connection = Connect-MclGraph -Settings $settings -Action $(if ($request.Action -in 'Restore', 'Transfer') { 'Remove' } else { $request.Action })
        $dot = [char]0x00B7
        Write-MclItem Ok ("Application {0} {1} tenant {2}" -f $(if ($connection.AppName) { "$($connection.AppName) ($($settings.AppId))" } else { $settings.AppId }), $dot, $connection.TenantGuid) -Icon Key
        Write-MclItem Info ("Permissions: {0}" -f (@($connection.Roles) -join ', ')) -Icon Shield
        if (-not $connection.CanReadPlaces -and @($request.SearchIn) -contains 'Rooms' -and $request.Action -ne 'Restore' -and $request.Mode -ne 'Rooms') { Write-MclItem Warn 'No Place.Read.All: the rooms come only from Search.Rooms and Search.RoomFile.' }

        if ($request.Action -eq 'Restore') {
            # ---- restore of a Remove (or Cancel) run --------------------------------------------------
            Write-MclNextStep 'Copies removed by the run' 'File'
            $result = Import-MclRestoreSource -Path $request.FromReport -MeetingId $request.MeetingId
            if ($result.Tenant -and $connection.TenantGuid -and $result.Tenant -ne $connection.TenantGuid) { throw "The report belongs to tenant $($result.Tenant), the application signs in to $($connection.TenantGuid)." }
            $plan = Get-MclRestorePlan -Result $result
            Write-MclItem Ok ('{0} meeting(s) of the {1} run {2} {3} cop{4} to restore' -f $result.Meetings.Count, $result.SourceAction, $dot, $plan.Restore.Count, $(if ($plan.Restore.Count -eq 1) { 'y' } else { 'ies' })) -Icon File
            if ($result.Meetings.Count) { Write-MclMeetingTable -Meetings @($result.Meetings) }
            if ($plan.Restore.Count) {
                if (-not $Force) { Confirm-MclAction -Title 'Restore' -Lines $plan.Lines }
                Write-MclNextStep 'Exchange Online PowerShell' 'Server'
                Connect-MclExchange -Settings $settings
                try { $result = Invoke-MclRestore -Settings $settings -Result $result }
                finally { Disconnect-MclExchange }
                Write-MclMeetingTable -Meetings @($result.Meetings)
            }
            else { $result = Invoke-MclRestore -Settings $settings -Result $result }
            $runPath = New-MclRunFolder -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Action 'Restore'
        }
        else {
            # ---- meetings: search, or the reviewed report ----------------------------------------------
            if ($request.FromReport) {
                Write-MclNextStep 'Meetings of the report' 'File'
                $result = Import-MclReport -Path $request.FromReport -MeetingId $request.MeetingId
                if ($result.Tenant -and $connection.TenantGuid -and $result.Tenant -ne $connection.TenantGuid) { throw "The report belongs to tenant $($result.Tenant), the application signs in to $($connection.TenantGuid)." }
                Write-MclItem Ok ("{0} meeting(s), {1} cop{2} from {3}" -f $result.Counts.Meetings, $result.Counts.Copies, $(if ($result.Counts.Copies -eq 1) { 'y' } else { 'ies' }), $result.FromReport) -Icon File
            }
            else {
                $result = Find-MclMeetings -Settings $settings -Request $request
            }
            if ($result.Meetings.Count) { Write-MclMeetingTable -Meetings @($result.Meetings) }

            # ---- action: backup in the folder of the report, then remove, cancel or transfer -------------
            $runPath = $null
            if ($request.Action -eq 'Transfer' -and $result.Meetings.Count) {
                $target = Resolve-MclNewOrganizer -Address $request.NewOrganizer
                $plan = Get-MclTransferPlan -Result $result -NewOrganizer $target -Method $request.TransferMethod -From $request.TransferFrom -Comment $request.Comment
                if (-not $Force) { Confirm-MclAction -Title 'Transfer' -Lines (@($plan.Lines) + 'A backup of the meetings is written first.') }
                $runPath = New-MclRunFolder -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Action 'Transfer'
                Write-MclNextStep 'Exchange Online PowerShell' 'Server'
                if ($plan.Native.Count) { Connect-MclExchange -Settings $settings -For Transfer }
                else { Write-MclItem Skip 'Not needed: every meeting is re-created with Microsoft Graph.' }
                try { $result = Invoke-MclTransfer -Settings $settings -Result $result -Plan $plan -From $request.TransferFrom -Comment $request.Comment -BackupPath (Join-Path $runPath "$($settings.ReportPrefix)-Backup.json") }
                finally { if ($plan.Native.Count) { Disconnect-MclExchange } }
                Write-MclMeetingTable -Meetings @($result.Meetings)
            }
            elseif ($request.Action -ne 'Report' -and $result.Meetings.Count) {
                $plan = Get-MclCleanupPlan -Result $result -Action $request.Action
                if (-not $Force) {
                    $lines = @($plan.Lines)
                    if ($request.Action -eq 'Cancel') { $lines += "Message: $($request.Comment)" }
                    $lines += 'A backup of the meetings is written first; the copies removed can be restored for the retention of deleted items (-Action Restore -FromReport <report folder>).'
                    Confirm-MclAction -Title $request.Action -Lines $lines
                }
                $runPath = New-MclRunFolder -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix -Action $request.Action
                $result = Invoke-MclCleanup -Settings $settings -Result $result -Action $request.Action -Comment $request.Comment -BackupPath (Join-Path $runPath "$($settings.ReportPrefix)-Backup.json")
                Write-MclMeetingTable -Meetings @($result.Meetings)
            }
            elseif ($request.Action -ne 'Report') { $result.Action = $request.Action }
        }

        # ---- report -------------------------------------------------------------------------------------
        $reportText = 'none (-NoReport)'
        if (-not $NoReport -or $runPath) {
            Write-MclNextStep 'Report' 'Report'
            $exportArgs = @{ Result = $result; OutputPath = $settings.OutputPath; Prefix = $settings.ReportPrefix; Formats = $settings.ReportFormats; Delimiter = $settings.CsvDelimiter; SummaryOnly = [bool]$NoReport }
            if ($runPath) { $exportArgs.Directory = $runPath }
            $report = Export-MclReport @exportArgs
            foreach ($f in $report.Files.Values) { Write-MclItem Ok $f -Icon File }
            $reportText = if ($report.Files.Contains('Html')) { $report.Files.Html } else { $report.Directory }
        }
        Write-MclRunSummary -Result $result -ReportText $reportText -LogPath $logPath
        $exitCode = switch ($result.Status) { 'Completed' { 0 } 'Failed' { 1 } default { 2 } }
    }
}
catch {
    if ($moduleLoaded) {
        Write-MclItem Fail $_.Exception.Message
        Write-MclLog 'ERROR' ($_.ScriptStackTrace -replace '\r?\n', ' | ')
        Write-Host ''
    }
    else {
        Write-Host "Meeting Cleanup: $($_.Exception.Message)" -ForegroundColor Red
    }
    $exitCode = 1
}
finally {
    if ($moduleLoaded) { Stop-MclLog }
    [Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
}
exit $exitCode
