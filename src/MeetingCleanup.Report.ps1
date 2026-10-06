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
    Version : 1.2.0
#>

$script:ReportColumns = [ordered]@{
    Meetings = @('MeetingId', 'Subject', 'Organizer', 'OrganizerName', 'Kind', 'Scope', 'Occurrences', 'StartText', 'EndText', 'NextInPeriod', 'Recurrence', 'Location', 'OrganizerCopy', 'Copies', 'RoomCopies', 'AttendeeCopies', 'NotProcessed', 'Cancelled', 'Selected', 'Status', 'NewOrganizer', 'NewMeetingId', 'TransferMethod', 'Notes')
    Copies   = @('MeetingId', 'MeetingSubject', 'Organizer', 'Mailbox', 'Role', 'Via', 'Occurrence', 'Response', 'ShowAs', 'Cancelled', 'Action', 'Result', 'HttpStatus', 'Verified', 'ActionUtc', 'Detail', 'EventId')
    Organizers = @('Input', 'DisplayName', 'PrimaryAddress', 'State', 'Detail', 'Meetings', 'Series', 'Copies', 'Removed', 'Cancelled', 'Restored', 'Transferred', 'Failed')
}

function Format-MclCsvCell {
    param([AllowNull()][object]$Value, [Parameter(Mandatory = $true)][string]$Delimiter)

    if ($null -eq $Value) { return '' }
    if ($Value -is [bool]) { $text = if ($Value) { 'True' } else { 'False' } }
    elseif ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) { $text = (@($Value) -join ' | ') }
    elseif ($Value -is [string]) {
        $text = $Value
        # Formula injection: Excel evaluates a cell starting with = + - @ (or tab / CR).
        if ($text -match '^[=+\-@\t\r]') { $text = "'" + $text }
    }
    else { $text = [string]$Value }
    if ($text.Contains($Delimiter) -or $text.Contains('"') -or $text -match '[\r\n]') { $text = '"' + $text.Replace('"', '""') + '"' }
    return $text
}

function Write-MclCsv {
    param(
        [AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory = $true)][string[]]$Columns,
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Delimiter = ';'
    )

    $builder = [Text.StringBuilder]::new()
    [void]$builder.AppendLine((@($Columns | ForEach-Object { Format-MclCsvCell $_ $Delimiter }) -join $Delimiter))
    foreach ($row in @($Rows)) {
        [void]$builder.AppendLine((@($Columns | ForEach-Object { Format-MclCsvCell (Get-MclProperty $row $_) $Delimiter }) -join $Delimiter))
    }
    [IO.File]::WriteAllText($Path, $builder.ToString(), [Text.UTF8Encoding]::new($true))
}

function ConvertTo-MclEmbeddedJson {
    <# JSON safe inside a <script type="application/json"> block. #>
    param([AllowNull()][object]$Value)
    $json = ConvertTo-Json -InputObject $Value -Depth 10 -Compress
    if ([string]::IsNullOrEmpty($json)) { $json = 'null' }
    return $json.Replace('<', '\u003c').Replace('>', '\u003e').Replace('&', '\u0026')
}

function Get-MclMeetingRows {
    <# The meetings flattened for the CSV and the HTML. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result)
    foreach ($m in @($Result.Meetings)) {
        # One copy per mailbox (an occurrence copy is counted once for its mailbox); the new organizer is not a copy.
        $copies = @($m.Copies | Where-Object { $_.EventId -and $_.Role -in 'Organizer', 'Attendee', 'Room' } | Group-Object Mailbox | ForEach-Object { $_.Group[0] })
        [pscustomobject]@{
            MeetingId = $m.MeetingId; Subject = $m.Subject; Organizer = $m.Organizer; OrganizerName = $m.OrganizerName; Kind = $m.Kind
            Scope = [string](Get-MclProperty $m 'Scope'); Occurrences = [int](Get-MclProperty $m 'Occurrences')
            NewOrganizer = [string](Get-MclProperty $m 'NewOrganizer'); NewMeetingId = [string](Get-MclProperty $m 'NewMeetingId'); TransferMethod = [string](Get-MclProperty $m 'TransferMethod')
            StartText = $m.StartText; EndText = $m.EndText; NextInPeriod = $m.NextInPeriod; Recurrence = $m.Recurrence; Location = $m.Location
            OrganizerCopy = $m.OrganizerCopy; Copies = $copies.Count; RoomCopies = @($copies | Where-Object Role -eq 'Room').Count
            AttendeeCopies = @($copies | Where-Object Role -eq 'Attendee').Count; NotProcessed = @($m.Copies | Where-Object Result -eq 'Not processed').Count
            Cancelled = [bool]$m.Cancelled; Selected = [bool]$m.Selected; Status = $m.Status; Notes = @($m.Notes)
        }
    }
}

function Get-MclCopyRows {
    param([Parameter(Mandatory = $true)][pscustomobject]$Result)
    foreach ($m in @($Result.Meetings)) {
        foreach ($c in @($m.Copies)) {
            [pscustomobject]@{
                MeetingId = $m.MeetingId; MeetingSubject = $m.Subject; Organizer = $m.Organizer; Mailbox = $c.Mailbox; Role = $c.Role; Via = $c.Via; Occurrence = [string](Get-MclProperty $c 'Occurrence'); Response = $c.Response; ShowAs = (Get-MclProperty $c 'ShowAs')
                Cancelled = [bool]$c.Cancelled; Action = $c.Action; Result = $c.Result; HttpStatus = $(if ($c.HttpStatus) { $c.HttpStatus } else { '' })
                Verified = $c.Verified; ActionUtc = (Get-MclProperty $c 'ActionUtc'); Detail = $c.Detail; EventId = $c.EventId
            }
        }
    }
}

function Get-MclOrganizerRows {
    <# One row per organizer of the run: its mailbox and what was found and done for its meetings. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result)
    foreach ($o in @($Result.Organizers)) {
        $keys = @(@($o.Addresses) + [string]$o.PrimaryAddress | Where-Object { $_ } | ForEach-Object { ([string]$_).ToLowerInvariant() })
        $mine = @($Result.Meetings | Where-Object { $k = [string](Get-MclProperty $_ 'OrganizerKey'); if (-not $k) { $k = [string]$_.Organizer }; $keys -contains $k.ToLowerInvariant() })
        $copies = @($mine | ForEach-Object { @($_.Copies) })
        [pscustomobject]@{
            Input = $o.Input; DisplayName = $o.DisplayName; PrimaryAddress = $o.PrimaryAddress; State = $o.State; Detail = $o.Detail
            Meetings = $mine.Count; Series = @($mine | Where-Object Kind -eq 'Series').Count; Copies = @($copies | Where-Object EventId).Count
            Removed = @($copies | Where-Object Result -eq 'Removed').Count; Cancelled = @($copies | Where-Object Result -eq 'Cancelled').Count
            Restored = @($copies | Where-Object Result -eq 'Restored').Count; Transferred = @($mine | Where-Object Status -eq 'Transferred').Count; Failed = @($copies | Where-Object Result -eq 'Failed').Count
        }
    }
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
    $meetings = @(Get-MclMeetingRows $Result)
    $copies = @(Get-MclCopyRows $Result)
    $organizers = @(Get-MclOrganizerRows $Result)
    $files = [ordered]@{}
    if ($Formats -contains 'Csv' -and -not $SummaryOnly) {
        $files.Meetings = Join-Path $runPath "$Prefix-Meetings.csv"
        Write-MclCsv -Rows $meetings -Columns $script:ReportColumns.Meetings -Path $files.Meetings -Delimiter $Delimiter
        $files.Copies = Join-Path $runPath "$Prefix-Copies.csv"
        Write-MclCsv -Rows $copies -Columns $script:ReportColumns.Copies -Path $files.Copies -Delimiter $Delimiter
        $files.Organizers = Join-Path $runPath "$Prefix-Organizers.csv"
        Write-MclCsv -Rows $organizers -Columns $script:ReportColumns.Organizers -Path $files.Organizers -Delimiter $Delimiter
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
        $html = $html.Replace('{{MEETINGS_JSON}}', (ConvertTo-MclEmbeddedJson @($meetings)))
        $html = $html.Replace('{{COPIES_JSON}}', (ConvertTo-MclEmbeddedJson @($copies)))
        $html = $html.Replace('{{ORGANIZERS_JSON}}', (ConvertTo-MclEmbeddedJson @($organizers)))
        if ($html -match '\{\{[A-Z_]+\}\}') { throw "Report template marker not replaced: $($Matches[0])" }
        $files.Html = Join-Path $runPath "$Prefix.html"
        [IO.File]::WriteAllText($files.Html, $html, [Text.UTF8Encoding]::new($true))
    }
    [pscustomobject]@{ Directory = $runPath; Files = $files }
}