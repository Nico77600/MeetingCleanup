<#
.SYNOPSIS
    Meeting Cleanup - PowerShell module.

.DESCRIPTION
    Loads the parts of the tool, in the order of an execution:

        src\MeetingCleanup.Console.ps1   console output and log file (same rules as the other tools)
        src\MeetingCleanup.Config.ps1    configuration file, settings, dates
        src\MeetingCleanup.Graph.ps1     Microsoft Graph: token, requests, $batch scheduler
        src\MeetingCleanup.Search.ps1    organizer, mailboxes to search, meetings and their copies
        src\MeetingCleanup.Cleanup.ps1   backup, remove or cancel, then verify
        src\MeetingCleanup.Restore.ps1   restore of the copies removed (Recoverable Items, Exchange Online PowerShell)
        src\MeetingCleanup.Transfer.ps1  transfer of meetings to a new organizer (Exchange Online, or re-created)
        src\MeetingCleanup.Report.ps1    CSV, JSON and HTML report, replay of a report
        src\MeetingCleanup.Gui.ps1       WPF window (Fluent theme of Windows 11)

    The access token and the client secret stay in memory: they are never written to the console,
    the log or the report.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.2.0
    History : see CHANGELOG.md
#>
#Requires -Version 7.4
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Net.Http

$script:ToolVersion = '1.2.2'
$script:ToolRoot = $PSScriptRoot
$script:LogWriter = $null
$script:LogPath = $null
$script:Quiet = $false
# GUI hooks, set only while the window runs a search or an action: Queue (progress lines, read by the window),
# Sink (lines, inline runs), Cancel, Hold.
$script:Ui = $null
# Microsoft Graph connection of the current run (Connect-MclGraph).
$script:Graph = $null

# Compiled helpers (src\MeetingCleanup.Native.cs): once per PowerShell process.
$native = 'MeetingCleanupNative.Fast' -as [type]
if (-not $native) { Add-Type -Path (Join-Path $PSScriptRoot 'src\MeetingCleanup.Native.cs') }
elseif ($native::Version -ne $script:ToolVersion) { throw "Meeting Cleanup $($native::Version) is already loaded in this PowerShell session: open a new PowerShell window to use $($script:ToolVersion)." }

foreach ($part in 'Console', 'Config', 'Graph', 'Search', 'Cleanup', 'Restore', 'Transfer', 'Report', 'Gui') {
    . (Join-Path $PSScriptRoot "src\MeetingCleanup.$part.ps1")
}

Export-ModuleMember -Function @(
    'Import-MclConfiguration', 'Test-MclConfiguration', 'New-MclRequest', 'Test-MclRequest', 'Connect-MclGraph', 'Resolve-MclOrganizer', 'Get-MclSearchMailboxes'
    'Find-MclMeetings', 'Get-MclCleanupPlan', 'Invoke-MclCleanup', 'Import-MclReport', 'Export-MclReport', 'New-MclRunFolder', 'Show-MclGui', 'New-MclForm'
    'Import-MclRestoreSource', 'Get-MclRestorePlan', 'Connect-MclExchange', 'Disconnect-MclExchange', 'Invoke-MclRestore'
    'Resolve-MclNewOrganizer', 'Get-MclTransferPlan', 'Invoke-MclTransfer'
    'Start-MclLog', 'Stop-MclLog', 'Write-MclLog', 'Write-MclBanner', 'Write-MclStep', 'Write-MclItem', 'Write-MclSummary', 'Initialize-MclSteps', 'Write-MclNextStep'
    'Write-MclRunBanner', 'Write-MclMeetingTable', 'Write-MclRunSummary', 'Format-MclDuration'
)
