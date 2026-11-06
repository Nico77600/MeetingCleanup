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

.PARAMETER SeriesScope
    A series: Whole (default, Search.SeriesScope) acts on the whole series, every occurrence; Occurrences acts only
    on its occurrences in the period - with a period of one day, one occurrence. The occurrences can also be chosen
    one by one in the window. With an action, the period must be given (-Start and -End).

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

.PARAMETER DevelopmentMode
    Allows a source checkout without PackageIntegrity.psd1. A present manifest is always verified.

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
    .\Invoke-MeetingCleanup.ps1 -OrganizerFile C:\Tools\MeetingCleanup-data\leavers.csv -SearchIn Organizer, Rooms, AllMailboxes
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
    Version : 1.3.1
    Exit codes : 0 = completed, 1 = failed, 2 = finished with warnings (a copy not removed or restored, a mailbox not read...).
    Documentation : docs\MeetingCleanup-UserGuide.html (user guide: prerequisites, everyday commands) and
                    docs\MeetingCleanup-Guide.html (developer guide); sources: docs\*.md

    Disclaimer:
    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

    As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.

    Removing, cancelling or transferring meetings changes real calendars: always start with a report, and test it in your environment before production use.
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
    [ValidateSet('Whole', 'Occurrences')]
    [string]$SeriesScope,
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
    [switch]$DevelopmentMode,
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
    # The package is verified before any of its code runs: the verifier must carry the owner signature, then it
    # checks the signed manifest. A tree without a manifest is accepted only with -DevelopmentMode.
    $integrityManifest = Join-Path $PSScriptRoot 'PackageIntegrity.psd1'
    $integrityHint = 'Download the complete release ZIP again from https://github.com/Nico77600/MeetingCleanup/releases/latest, extract it and unblock it. A development or CI build is started from its folder with Invoke-MeetingCleanup.ps1 <parameters> -DevelopmentMode.'
    if (Microsoft.PowerShell.Management\Test-Path -LiteralPath $integrityManifest) {
        $integrityVerifier = Join-Path $PSScriptRoot 'src\Test-PackageIntegrity.ps1'
        $integrityCursor = [IO.Path]::GetFullPath($integrityVerifier)
        while ($integrityCursor) {
            $integrityItem = Microsoft.PowerShell.Management\Get-Item -LiteralPath $integrityCursor -Force -ErrorAction Stop
            if ($integrityItem.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Package integrity: reparse point: $integrityCursor" }
            $integrityCursor = [IO.Path]::GetDirectoryName($integrityCursor)
        }
        $integritySignature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $integrityVerifier -ErrorAction Stop
        if ([string]$integritySignature.Status -cne 'Valid' -or $null -eq $integritySignature.SignerCertificate -or
            $integritySignature.SignerCertificate.Thumbprint -cne '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6' -or
            $null -eq $integritySignature.TimeStamperCertificate) {
            throw "Package integrity: verifier must have the trusted, timestamped owner signature. $integrityHint"
        }
        if ($DevelopmentMode) { & $integrityVerifier -Root $PSScriptRoot -DevelopmentMode | Out-Null } else { & $integrityVerifier -Root $PSScriptRoot -Distribution | Out-Null }
    }
    elseif (-not $DevelopmentMode) {
        throw "Package integrity: PackageIntegrity.psd1 is missing. $integrityHint"
    }
    else {
        Write-Warning 'DevelopmentMode: no package manifest; package byte integrity and publisher authentication are not established.'
        # The module does not repeat this warning in the same process.
        [AppDomain]::CurrentDomain.SetData('MeetingCleanup.UnverifiedTreeWarned', $true)
    }
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
        if ($SeriesScope) { $requestArgs.SeriesScope = $SeriesScope }
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

# SIG # Begin signature block
# MIImIgYJKoZIhvcNAQcCoIImEzCCJg8CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCATuKPaXZgC6YWT
# veWYVZE7sOkS4QMzEbPwAWGHhhHoP6CCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAK9szS1FeD
# RB9BmJhZ7k8JX+7DNSBhruJs5f1ELEPX2DANBgkqhkiG9w0BAQEFAASCAgB2dtLM
# gy1OL+MPH5TxHtcsnWc7tN//QOlzpdkhLP67gjrZbqVoc+gkIjG3YEWjYy4IC39m
# bSKVzTi8MQbe9MOcbQD5bQ/m0kpratGpiIF1Ee7eHMGpxoIuWHO/gSBV8KI6DXvz
# 1LDPwdHSzopEiju5GuoruMeRWYXgsfQSv8ka+kaYUsWbGMb3yvvynPVg07s7ehPw
# KGynjluJ0nS/eUuCVGt1VzyyW6E5/LFeIC3OXF31p3Deb8KBFmFqXTvqU1RsMWPt
# jOB772ndUL1AHlhPrNzZoKtg4Ul8UF0gRG9t8JxX2s0grBae8awWckOwgUxaVCSn
# EcpRqN8T4sTS9GiYF3Rrq1iCFa7SM7vnvXCtHfeNw0sfobBaup7grS/jx9oCcLg9
# GpgbPjH2OTUEHhWqYWf7uxnM7+nP+bx5JZwSbR9f0q41W89osulDsfvJmH1LTNBz
# pBa+o+uR4b/IDuJ5VbDtA+qTS1m6pjvBqV8VD3EK3e7n/zh8J+zSRtxBaT2ACt7w
# gyy1Y/hJFs0Oho8lLs27efcVVg9uw29G37N+fPnz9+klSACptt3IY15U3JsO953r
# lPa/gdj7ZiTNQrvNdXEDekUTfY0pczdT43zDqSU8tPZs0/WHOEBcX5TDga8GNLQg
# Yafg9A2X1bjONeZc7JaLl3LvsTczlo9z41jNXKGCFMYwghTCBgorBgEEAYI3AwMB
# MYIUsjCCFK4GCSqGSIb3DQEHAqCCFJ8wghSbAgEDMQ0wCwYJYIZIAWUDBAIBMIGB
# BgsqhkiG9w0BCRABBKByBHAwbgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIFlDGelunZggrEhCzeVliXzfp7NeFZiCFSGjKFAsrw6CAgg6bRJx1cFs
# yhgPMjAyNjExMDYwODQyMjFaMAMCAQECCBr+QiwAf/zAoIIRpzCCBPMwggLboAMC
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
# CSqGSIb3DQEJBTEPFw0yNjExMDYwODQyMjFaMCgGCSqGSIb3DQEJNDEbMBkwCwYJ
# YIZIAWUDBAIBoQoGCCqGSM49BAMCMC8GCSqGSIb3DQEJBDEiBCCHpNoYIQkbxTpQ
# sNt2PpX/LgWmBYoPFOtDpbJND/dbBzCByAYLKoZIhvcNAQkQAi8xgbgwgbUwgbIw
# ga8EIFdP6heq+l58aiubCMmYsXNeRQs4c8pE5qJMoJMOY/OwMIGKMHakdDByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7jMAoGCCqGSM49BAMCBEcw
# RQIgSPbHtwkEqrblxikWORGgBGoGjuvIcsnfhvzgr8AAtJACIQCRA5aLiaRbST9X
# VhduexOuju5AUxuJ/KNN7EUdHYW/rw==
# SIG # End signature block
