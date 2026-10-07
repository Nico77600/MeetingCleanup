<#
.SYNOPSIS
    Meeting Cleanup - report files (dot-sourced by MeetingCleanup.psm1).

.DESCRIPTION
    One folder per execution, <FilePrefix>_<Action>_<yyyyMMdd-HHmmss>, with:
      <prefix>-Organizers.csv one row per organizer: address, state of the mailbox, meetings and copies, results
      <prefix>-Meetings.csv   one row per meeting: subject, organizer, start, series, organizer copy, copies, status
      <prefix>-Copies.csv     one row per mailbox: role, how it was found, action, result, Graph status, verified
      <prefix>-Summary.json   the whole result, for scripts and for -FromReport (replay, restore)
      <prefix>-Backup.json    Remove, Cancel and Transfer: the meetings and copies as they were, written before any change
      <prefix>.html           self-contained dashboard (templates\Report.template.html)
    CSV files: UTF-8 with BOM, configurable delimiter, text cells starting with = + - @ are prefixed with an
    apostrophe (no formula injection when opened in Excel). No token or secret is ever part of the result.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.2.2
#>

$script:ReportColumns = [ordered]@{
    Meetings = @('MeetingId', 'Subject', 'Organizer', 'OrganizerName', 'Kind', 'Scope', 'Occurrences', 'StartText', 'EndText', 'NextInPeriod', 'Recurrence', 'Location', 'OrganizerCopy', 'Copies', 'RoomCopies', 'AttendeeCopies', 'NotProcessed', 'Cancelled', 'Selected', 'Status', 'NewOrganizer', 'NewMeetingId', 'TransferMethod', 'Notes')
    Copies   = @('MeetingId', 'MeetingSubject', 'Organizer', 'Mailbox', 'Role', 'Via', 'Occurrence', 'Response', 'ShowAs', 'Cancelled', 'Action', 'Result', 'HttpStatus', 'Verified', 'ActionUtc', 'Detail', 'EventId')
    Organizers = @('Input', 'DisplayName', 'PrimaryAddress', 'State', 'Detail', 'Meetings', 'Series', 'Copies', 'Removed', 'Cancelled', 'Restored', 'Transferred', 'Failed')
}

function Format-MclCsvCell {
    <# A CSV cell: text starting with = + - @ (or tab, CR) prefixed with an apostrophe, quoted when needed (compiled). #>
    param([AllowNull()][object]$Value, [Parameter(Mandatory = $true)][string]$Delimiter)
    return [MeetingCleanupNative.Fast]::CsvCell($Value, $Delimiter)
}

function Write-MclCsv {
    <# A CSV file: UTF-8 with BOM, the columns given, one line per row (compiled). #>
    param(
        [AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory = $true)][string[]]$Columns,
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Delimiter = ';'
    )
    [MeetingCleanupNative.Fast]::WriteCsv($Rows, $Columns, $Path, $Delimiter)
}

function ConvertTo-MclEmbeddedJson {
    <# JSON safe inside a <script type="application/json"> block. #>
    param([AllowNull()][object]$Value)
    $json = ConvertTo-Json -InputObject $Value -Depth 10 -Compress
    if ([string]::IsNullOrEmpty($json)) { $json = 'null' }
    return $json.Replace('<', '\u003c').Replace('>', '\u003e').Replace('&', '\u0026')
}

function Get-MclMeetingRows {
    <# The meetings flattened for the CSV and the HTML: one copy per mailbox (an occurrence copy is counted once for its mailbox); the new organizer is not a copy. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result)
    [MeetingCleanupNative.Fast]::MeetingTable($Result.Meetings).ToObjects()
}

function Get-MclCopyRows {
    param([Parameter(Mandatory = $true)][pscustomobject]$Result)
    [MeetingCleanupNative.Fast]::CopyTable($Result.Meetings).ToObjects()
}

function Get-MclOrganizerRows {
    <# One row per organizer of the run: its mailbox and what was found and done for its meetings. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result)
    [MeetingCleanupNative.Fast]::OrganizerTable($Result.Organizers, $Result.Meetings).ToObjects()
}
function New-MclRunFolder {
    <# New folder of a run under OutputPath: <Prefix>_<Action>_<yyyyMMdd-HHmmss>. #>
    param([Parameter(Mandatory = $true)][string]$OutputPath, [string]$Prefix = 'MeetingCleanup', [string]$Action = 'Report')
    $base = Join-Path $OutputPath ('{0}_{1}_{2}' -f $Prefix, $Action, (Get-Date).ToString('yyyyMMdd-HHmmss'))
    $path = $base
    $n = 2
    while (Test-Path -LiteralPath $path) { $path = "$base-$n"; $n++ }
    [void][IO.Directory]::CreateDirectory($path)
    return $path
}

function Export-MclReport {
    <#
    .SYNOPSIS
        Writes the CSV, JSON and HTML files of one result in a new folder under OutputPath (or in Directory, the
        folder created before the action, where the backup already is).
    .PARAMETER SummaryOnly
        Only the Summary.json file (-NoReport of an action: the restore needs it).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Result,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [string]$Prefix = 'MeetingCleanup',
        [ValidateSet('Csv', 'Html')][string[]]$Formats = @('Csv', 'Html'),
        [ValidateSet(';', ',', "`t")][string]$Delimiter = ';',
        [string]$Directory,
        [switch]$SummaryOnly
    )

    $runPath = if ($Directory) { [void][IO.Directory]::CreateDirectory($Directory); $Directory } else { New-MclRunFolder -OutputPath $OutputPath -Prefix $Prefix -Action ([string]$Result.Action) }
    # The rows (thousands for a large search) as compiled tables, written to CSV and JSON without PowerShell objects.
    $meetings = [MeetingCleanupNative.Fast]::MeetingTable($Result.Meetings)
    $copies = [MeetingCleanupNative.Fast]::CopyTable($Result.Meetings)
    $organizers = [MeetingCleanupNative.Fast]::OrganizerTable($Result.Organizers, $Result.Meetings)
    $files = [ordered]@{}
    if ($Formats -contains 'Csv' -and -not $SummaryOnly) {
        $files.Meetings = Join-Path $runPath "$Prefix-Meetings.csv"
        [MeetingCleanupNative.Fast]::WriteTableCsv($meetings, $script:ReportColumns.Meetings, $files.Meetings, $Delimiter)
        $files.Copies = Join-Path $runPath "$Prefix-Copies.csv"
        [MeetingCleanupNative.Fast]::WriteTableCsv($copies, $script:ReportColumns.Copies, $files.Copies, $Delimiter)
        $files.Organizers = Join-Path $runPath "$Prefix-Organizers.csv"
        [MeetingCleanupNative.Fast]::WriteTableCsv($organizers, $script:ReportColumns.Organizers, $files.Organizers, $Delimiter)
    }
    $files.Summary = Join-Path $runPath "$Prefix-Summary.json"
    # RunRemoved (restore: the copies of the source run, also in Meetings) is not written again.
    $data = if ($Result.PSObject.Properties['RunRemoved']) { $Result | Select-Object -Property * -ExcludeProperty RunRemoved } else { $Result }
    [IO.File]::WriteAllText($files.Summary, (ConvertTo-Json -InputObject $data -Depth 10), [Text.UTF8Encoding]::new($false))
    $backup = Join-Path $runPath "$Prefix-Backup.json"
    if (Test-Path -LiteralPath $backup) { $files.Backup = $backup }

    if ($Formats -contains 'Html' -and -not $SummaryOnly) {
        $summary = [ordered]@{}
        foreach ($key in 'Tool', 'Version', 'Action', 'Status', 'Error', 'StartedUtc', 'CompletedUtc', 'DurationSeconds', 'Request', 'Tenant', 'Organization', 'AppId', 'AppName', 'Organizers', 'Searched', 'Warnings', 'Counts', 'CleanupComment', 'FromReport', 'BackupFile', 'SourceAction', 'NewOrganizer') {
            # Direct assignment: a list of one item stays a list (a function output would unroll it).
            $summary[$key] = $null
            $prop = $Result.PSObject.Properties[$key]
            if ($prop) { $summary[$key] = $prop.Value }
        }
        $summary.ActionText = Get-MclActionText ([string]$Result.Action)
        $summary.GeneratedText = Format-MclDate ([datetime]::UtcNow) ([string](Get-MclProperty $Result.Request 'TimeZone'))
        $html = [IO.File]::ReadAllText((Join-Path $script:ToolRoot 'templates\Report.template.html'))
        $names = @($Result.Organizers | ForEach-Object { if ($_.DisplayName) { $_.DisplayName } else { $_.Input } })
        $title = "Meeting Cleanup | $($Result.Action) | $(if ($names.Count -le 3) { $names -join ', ' } else { "$($names.Count) organizers" })"
        $html = $html.Replace('{{TITLE}}', [Net.WebUtility]::HtmlEncode($title))
        $html = $html.Replace('{{SUMMARY_JSON}}', (ConvertTo-MclEmbeddedJson $summary))
        # The rows (thousands for a large search): compiled JSON writer, HTML-safe as well.
        $html = $html.Replace('{{MEETINGS_JSON}}', [MeetingCleanupNative.Fast]::TableJson($meetings))
        $html = $html.Replace('{{COPIES_JSON}}', [MeetingCleanupNative.Fast]::TableJson($copies))
        $html = $html.Replace('{{ORGANIZERS_JSON}}', [MeetingCleanupNative.Fast]::TableJson($organizers))
        if ($html -match '\{\{[A-Z_]+\}\}') { throw "Report template marker not replaced: $($Matches[0])" }
        $files.Html = Join-Path $runPath "$Prefix.html"
        [IO.File]::WriteAllText($files.Html, $html, [Text.UTF8Encoding]::new($true))
    }
    [pscustomobject]@{ Directory = $runPath; Files = $files }
}