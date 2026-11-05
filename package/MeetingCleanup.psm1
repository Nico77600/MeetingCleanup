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
    Version : 1.3.1
    History : see CHANGELOG.md
#>
#Requires -Version 7.4
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Net.Http

$script:ToolVersion = '1.3.1'
$script:ToolRoot = $PSScriptRoot

# Package integrity. A package carries the owner-signed PackageIntegrity.psd1 manifest; when it is present, the signed
# verifier (src\Test-PackageIntegrity.ps1) must accept the tree before any part of the tool is compiled or loaded.
# A source tree without a manifest is an unverified development tree: the entry script refuses it unless
# -DevelopmentMode is given, and the module says so once per process.
function Assert-MclPackageIntegrity {
    [CmdletBinding()]
    param([string]$Root = $script:ToolRoot, [switch]$DevelopmentMode)
    $packageRoot = [IO.Path]::TrimEndingDirectorySeparator([IO.Path]::GetFullPath($Root))
    $manifest = Join-Path $packageRoot 'PackageIntegrity.psd1'
    if ((Test-Path -LiteralPath $manifest) -and -not (Test-Path -LiteralPath $manifest -PathType Leaf)) {
        throw 'Package integrity: PackageIntegrity.psd1 must be a file.'
    }
    if (-not (Test-Path -LiteralPath $manifest -PathType Leaf)) {
        if (-not [AppDomain]::CurrentDomain.GetData('MeetingCleanup.UnverifiedTreeWarned')) {
            [AppDomain]::CurrentDomain.SetData('MeetingCleanup.UnverifiedTreeWarned', $true)
            Write-Warning 'Package integrity: no PackageIntegrity.psd1; this is an unverified source tree (package byte integrity and publisher authentication are not established).'
        }
        return
    }
    $verifier = Join-Path $packageRoot 'src\Test-PackageIntegrity.ps1'
    $cursor = [IO.Path]::GetFullPath($verifier)
    while ($cursor) {
        if (-not (Test-Path -LiteralPath $cursor)) { throw 'Package integrity: the verifier is missing; nothing was loaded.' }
        $item = Microsoft.PowerShell.Management\Get-Item -LiteralPath $cursor -Force -ErrorAction Stop
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Package integrity: reparse point: $cursor" }
        $cursor = [IO.Path]::GetDirectoryName($cursor)
    }
    $signature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $verifier -ErrorAction Stop
    if ([string]$signature.Status -cne 'Valid' -or $null -eq $signature.SignerCertificate -or
        $signature.SignerCertificate.Thumbprint -cne '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6' -or
        $null -eq $signature.TimeStamperCertificate) {
        throw 'Package integrity: verifier must have the trusted, timestamped owner signature. Download the complete release ZIP again from https://github.com/Nico77600/MeetingCleanup/releases/latest, extract it and unblock it.'
    }
    $arguments = @{ Root = $packageRoot }
    if ($DevelopmentMode) { $arguments.DevelopmentMode = $true }
    $result = @(& $verifier @arguments)
    if ($result.Count -ne 1 -or -not $result[0].PSObject.Properties['Status'] -or $result[0].Status -notlike 'Valid*') {
        throw 'Package integrity: the verifier did not return an explicit Valid result; nothing was loaded.'
    }
}
Assert-MclPackageIntegrity
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
