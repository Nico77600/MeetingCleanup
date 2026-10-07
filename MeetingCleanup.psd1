#
#  Meeting Cleanup - module manifest
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : see ModuleVersion
#
#  Loaded by Invoke-MeetingCleanup.ps1 (Import-Module by path).
#
@{
    RootModule        = 'MeetingCleanup.psm1'
    ModuleVersion     = '1.2.2'
    GUID              = '3264d592-5f9b-4155-8bef-cb2bb4feff0e'
    Author            = 'Nicolas Fabert'
    Copyright         = '(c) 2026 Nicolas Fabert. MIT License.'
    Description       = 'Meeting Cleanup: finds the meetings of organizers (present or deleted mailbox) or of rooms in Exchange Online - one meeting, a series or a period - then removes them from the attendees and the rooms, cancels them, or transfers them to a new organizer, with a backup before any change, the restore of the copies removed (Recoverable Items), CSV, JSON and HTML reports and a window.'
    PowerShellVersion = '7.4'

    # Functions called by Invoke-MeetingCleanup.ps1, the tests and the documentation tools. The other functions stay internal.
    FunctionsToExport = @(
        'Import-MclConfiguration', 'Test-MclConfiguration', 'New-MclRequest', 'Test-MclRequest', 'Connect-MclGraph', 'Resolve-MclOrganizer', 'Get-MclSearchMailboxes'
        'Find-MclMeetings', 'Get-MclCleanupPlan', 'Invoke-MclCleanup', 'Import-MclReport', 'Export-MclReport', 'New-MclRunFolder', 'Show-MclGui', 'New-MclForm'
        'Import-MclRestoreSource', 'Get-MclRestorePlan', 'Connect-MclExchange', 'Disconnect-MclExchange', 'Invoke-MclRestore'
        'Resolve-MclNewOrganizer', 'Get-MclTransferPlan', 'Invoke-MclTransfer'
        'Start-MclLog', 'Stop-MclLog', 'Write-MclLog', 'Write-MclBanner', 'Write-MclStep', 'Write-MclItem', 'Write-MclSummary', 'Initialize-MclSteps', 'Write-MclNextStep'
        'Write-MclRunBanner', 'Write-MclMeetingTable', 'Write-MclRunSummary', 'Format-MclDuration'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags       = @('ExchangeOnline', 'Calendar', 'Meeting', 'Room', 'MicrosoftGraph', 'Cleanup')
            LicenseUri = 'https://opensource.org/licenses/MIT'
        }
    }
}
