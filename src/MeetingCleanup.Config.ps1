<#
.SYNOPSIS
    Meeting Cleanup - configuration, request and dates (dot-sourced by MeetingCleanup.psm1).

.DESCRIPTION
    The configuration file has sections (Tenant, Authentication, Search, Cleanup, Restore, Transfer, Graph, Report, Logging),
    like the other tools. It is flattened into one settings hashtable used by the command line, the window
    and the tests; unknown sections or keys and invalid values are all reported at once.

    A request is what one run does: the organizers (or the rooms), the period, the subject or the meeting IDs,
    where to search and the action (with the new organizer of a transfer). The command line and the window both build one with New-MclRequest, so they
    are checked by the same rules (Test-MclRequest).

    Dates: the period is typed in the time zone of Report.TimeZone (empty = the one of Windows). A start
    date is the start of that day; an end date without a time is the end of that day (included).

.NOTES
    Author  : Nicolas Fabert
    Version : 1.2.0
#>

# Section.Key of the configuration file -> key of the settings hashtable.
$script:ConfigSchema = [ordered]@{
    Tenant         = [ordered]@{ TenantId = 'TenantId'; Organization = 'Organization' }
    Authentication = [ordered]@{ Mode = 'AuthMode'; AppId = 'AppId'; CertificateThumbprint = 'CertificateThumbprint'; ClientSecretVariable = 'ClientSecretVariable' }
    Search         = [ordered]@{ SearchIn = 'SearchIn'; PastDays = 'PastDays'; FutureDays = 'FutureDays'; Rooms = 'Rooms'; RoomFile = 'RoomFile'; MailboxFile = 'MailboxFile' }
    Cleanup        = [ordered]@{ CancelComment = 'CancelComment'; Verify = 'Verify' }
    Restore        = [ordered]@{ Connection = 'RestoreConnection'; UserPrincipalName = 'RestoreUser'; WindowMinutes = 'RestoreWindowMinutes'; ReAccept = 'RestoreReAccept' }
    Transfer       = [ordered]@{ Method = 'TransferMethod'; Comment = 'TransferComment' }
    Graph          = [ordered]@{ MaxConcurrency = 'MaxConcurrency'; PageSize = 'PageSize'; MaxRetries = 'MaxRetries'; TimeoutSeconds = 'TimeoutSeconds' }
    Report         = [ordered]@{ OutputPath = 'OutputPath'; FilePrefix = 'ReportPrefix'; Formats = 'ReportFormats'; CsvDelimiter = 'CsvDelimiter'; TimeZone = 'TimeZone' }
    Logging        = [ordered]@{ Path = 'LogPath'; RetentionDays = 'LogRetentionDays' }
}

$script:SearchScopes = @('Organizer', 'Rooms', 'Mailboxes', 'AllMailboxes')
$script:Actions = @('Report', 'Remove', 'Cancel', 'Restore', 'Transfer')
$script:TransferMethods = @('Auto', 'Native', 'Recreate')
$script:GuidPattern = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
$script:SmtpPattern = '^[^@\s<>"]+@[^@\s<>"]+\.[^@\s<>"]+$'
# legacyExchangeDN or X500 address of a mailbox (the address the copies may show once the mailbox is deleted).
$script:X500Pattern = '^(?i)(x500:)?/o=[^\r\n]+/cn=[^\r\n]+$'

function Get-MclDefaultConfiguration {
    @{
        TenantId              = ''
        Organization          = ''
        AuthMode              = 'Certificate'
        AppId                 = ''
        CertificateThumbprint = ''
        ClientSecretVariable  = 'MCL_CLIENT_SECRET'
        SearchIn              = @('Organizer', 'Rooms')
        PastDays              = 0
        FutureDays            = 365
        Rooms                 = @()
        RoomFile              = ''
        MailboxFile           = ''
        CancelComment         = 'This meeting has been cancelled by the IT department.'
        Verify                = $true
        RestoreConnection     = 'Application'
        RestoreUser           = ''
        RestoreWindowMinutes  = 10
        RestoreReAccept       = $true
        TransferMethod        = 'Auto'
        TransferComment       = 'This meeting is now organized by {0}.'
        MaxConcurrency        = 16
        PageSize              = 500
        MaxRetries            = 6
        TimeoutSeconds        = 120
        OutputPath            = '.\reports'
        ReportPrefix          = 'MeetingCleanup'
        ReportFormats         = @('Csv', 'Html')
        CsvDelimiter          = ';'
        TimeZone              = ''
        LogPath               = '.\logs'
        LogRetentionDays      = 30
    }
}

function Resolve-MclPath {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Root)
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    return [IO.Path]::GetFullPath($Path, $Root)
}

function Get-MclTimeZone {
    <# Time zone of the dates: Report.TimeZone (IANA or Windows ID), the one of Windows when empty. #>
    param([AllowEmptyString()][AllowNull()][string]$Id)
    if ([string]::IsNullOrWhiteSpace($Id)) { return [TimeZoneInfo]::Local }
    return [TimeZoneInfo]::FindSystemTimeZoneById($Id)
}

function Format-MclDate {
    <# A UTC date shown in the time zone of the report: yyyy-MM-dd HH:mm (or yyyy-MM-dd). -PeriodEnd: an end at 00:00 shows the day before (included). #>
    param([AllowNull()][object]$Utc, [AllowEmptyString()][AllowNull()][string]$TimeZone, [switch]$DateOnly, [switch]$PeriodEnd)
    if ($null -eq $Utc -or [string]$Utc -eq '') { return '' }
    $d = if ($Utc -is [datetime]) { $Utc } else { [datetime]::Parse([string]$Utc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal) }
    $d = [datetime]::SpecifyKind($d.ToUniversalTime(), [DateTimeKind]::Utc)
    $local = [TimeZoneInfo]::ConvertTimeFromUtc($d, (Get-MclTimeZone $TimeZone))
    if ($PeriodEnd -and $local.TimeOfDay -eq [TimeSpan]::Zero) { $local = $local.AddDays(-1); $DateOnly = $true }
    return $local.ToString($(if ($DateOnly) { 'yyyy-MM-dd' } else { 'yyyy-MM-dd HH:mm' }), [Globalization.CultureInfo]::InvariantCulture)
}

function ConvertTo-MclUtc {
    <#
        A date typed by the administrator (time zone of the report) to UTC. -EndOfDay: a date without a time
        is the end of that day (the next day at 00:00), so that the end date is included.
    #>
    param([Parameter(Mandatory = $true)][datetime]$Date, [AllowEmptyString()][AllowNull()][string]$TimeZone, [switch]$EndOfDay)
    if ($Date.Kind -eq [DateTimeKind]::Utc) { return $Date }
    $d = [datetime]::SpecifyKind($Date, [DateTimeKind]::Unspecified)
    if ($EndOfDay -and $d.TimeOfDay -eq [TimeSpan]::Zero) { $d = $d.AddDays(1) }
    return [TimeZoneInfo]::ConvertTimeToUtc($d, (Get-MclTimeZone $TimeZone))
}

function Get-MclDefaultPeriod {
    <# Start = today minus Search.PastDays, end = today plus Search.FutureDays (included), as UTC dates. #>
    param([Parameter(Mandatory = $true)][hashtable]$Settings)
    $zone = Get-MclTimeZone $Settings.TimeZone
    $today = [TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, $zone).Date
    [pscustomobject]@{
        Start = ConvertTo-MclUtc -Date $today.AddDays(-[int]$Settings.PastDays) -TimeZone $Settings.TimeZone
        End   = ConvertTo-MclUtc -Date $today.AddDays([int]$Settings.FutureDays) -TimeZone $Settings.TimeZone -EndOfDay
    }
}

function Test-MclConfiguration {
    <# Checks a settings hashtable (flattened configuration) and lists every problem. -ForConnection: tenant and application required. #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Configuration, [switch]$ForConnection)

    $problems = [Collections.Generic.List[string]]::new()
    $c = $Configuration
    $number = {
        param([string]$Key, [string]$Label, [int]$Min, [int]$Max)
        $n = 0
        if (-not [int]::TryParse([string]$c[$Key], [ref]$n) -or $n -lt $Min -or $n -gt $Max) { [void]$problems.Add("$Label must be a whole number between $Min and $Max.") }
    }

    if ([string]$c.TenantId -and [string]$c.TenantId -notmatch $script:GuidPattern -and [string]$c.TenantId -notmatch '^[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$') {
        [void]$problems.Add('Tenant.TenantId must be the tenant ID (GUID) or a domain of the tenant (contoso.onmicrosoft.com).')
    }
    if ([string]$c.AuthMode -notin 'Certificate', 'ClientSecret') { [void]$problems.Add("Authentication.Mode must be 'Certificate' or 'ClientSecret'.") }
    if ([string]$c.AppId -and [string]$c.AppId -notmatch $script:GuidPattern) { [void]$problems.Add('Authentication.AppId must be the application (client) ID, a GUID.') }
    if ([string]$c.AuthMode -eq 'Certificate' -and [string]$c.CertificateThumbprint -and [string]$c.CertificateThumbprint -notmatch '^[0-9a-fA-F]{40}$') {
        [void]$problems.Add('Authentication.CertificateThumbprint must be the 40 hexadecimal characters of the thumbprint.')
    }
    if ([string]$c.ClientSecretVariable -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') { [void]$problems.Add('Authentication.ClientSecretVariable must be the name of an environment variable.') }
    if ($ForConnection) {
        if (-not [string]$c.TenantId) { [void]$problems.Add('Tenant.TenantId is required.') }
        if (-not [string]$c.AppId) { [void]$problems.Add('Authentication.AppId is required.') }
        if ([string]$c.AuthMode -eq 'Certificate' -and -not [string]$c.CertificateThumbprint) { [void]$problems.Add('Authentication.CertificateThumbprint is required in Certificate mode.') }
    }
    $scopes = @($c.SearchIn)
    if ($scopes.Count -eq 0 -or @($scopes | Where-Object { $_ -notin $script:SearchScopes }).Count) { [void]$problems.Add("Search.SearchIn must contain one or more of: $($script:SearchScopes -join ', ').") }
    & $number 'PastDays' 'Search.PastDays' 0 3650
    & $number 'FutureDays' 'Search.FutureDays' 0 3650
    foreach ($room in @($c.Rooms)) { if ([string]$room -notmatch $script:SmtpPattern) { [void]$problems.Add("Search.Rooms: '$room' is not an SMTP address.") } }
    if ($c.Verify -isnot [bool]) { [void]$problems.Add('Cleanup.Verify must be $true or $false.') }
    if ([string]$c.RestoreConnection -notin 'Application', 'Interactive') { [void]$problems.Add("Restore.Connection must be 'Application' or 'Interactive'.") }
    if ([string]$c.RestoreUser -and [string]$c.RestoreUser -notmatch $script:SmtpPattern) { [void]$problems.Add('Restore.UserPrincipalName must be empty or a UPN (admin@contoso.com).') }
    & $number 'RestoreWindowMinutes' 'Restore.WindowMinutes' 1 240
    if ($c.RestoreReAccept -isnot [bool]) { [void]$problems.Add('Restore.ReAccept must be $true or $false.') }
    if ([string]$c.TransferMethod -notin $script:TransferMethods) { [void]$problems.Add("Transfer.Method must be one of: $($script:TransferMethods -join ', ').") }
    if ([string]$c.TransferComment -match '[<>]') { [void]$problems.Add('Transfer.Comment must be plain text (no < or >).') }
    if ([string]$c.CancelComment -match '[<>]') { [void]$problems.Add('Cleanup.CancelComment must be plain text (no < or >).') }
    & $number 'MaxConcurrency' 'Graph.MaxConcurrency' 1 32
    & $number 'PageSize' 'Graph.PageSize' 10 1000
    & $number 'MaxRetries' 'Graph.MaxRetries' 0 10
    & $number 'TimeoutSeconds' 'Graph.TimeoutSeconds' 10 600
    foreach ($pair in @(@('OutputPath', 'Report.OutputPath'), @('ReportPrefix', 'Report.FilePrefix'), @('LogPath', 'Logging.Path'))) {
        if ([string]::IsNullOrWhiteSpace([string]$c[$pair[0]])) { [void]$problems.Add("$($pair[1]) is required.") }
    }
    if ([string]$c.ReportPrefix -match '[\\/:*?"<>|\s]') { [void]$problems.Add('Report.FilePrefix must be a file name without spaces.') }
    $formats = @($c.ReportFormats)
    if ($formats.Count -eq 0 -or @($formats | Where-Object { $_ -notin 'Csv', 'Html' }).Count) { [void]$problems.Add("Report.Formats must contain 'Csv', 'Html' or both.") }
    if ([string]$c.CsvDelimiter -notin ';', ',', "`t") { [void]$problems.Add("Report.CsvDelimiter must be ';', ',' or a tab.") }
    if ([string]$c.TimeZone) { try { $null = Get-MclTimeZone $c.TimeZone } catch { [void]$problems.Add("Report.TimeZone '$($c.TimeZone)' is not a time zone of this computer (for example Europe/Paris, or empty).") } }
    & $number 'LogRetentionDays' 'Logging.RetentionDays' 1 365

    [pscustomobject]@{ IsValid = $problems.Count -eq 0; Problems = @($problems) }
}

function Import-MclConfiguration {
    <#
        Reads config\MeetingCleanup.config.psd1 (sections), applies the defaults, resolves the relative
        paths from the tool folder, checks everything and returns the settings hashtable.
    #>
    [CmdletBinding()]
    param(
        [string]$Path = (Join-Path $script:ToolRoot 'config\MeetingCleanup.config.psd1'),
        [string]$Root = $script:ToolRoot
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Configuration file not found: $Path" }
    $settings = Get-MclDefaultConfiguration
    $problems = [Collections.Generic.List[string]]::new()
    $data = Import-PowerShellDataFile -LiteralPath $Path
    foreach ($section in $data.Keys) {
        if (-not $script:ConfigSchema.Contains($section)) { [void]$problems.Add("Unknown section '$section'. Sections: $($script:ConfigSchema.Keys -join ', ')."); continue }
        if ($data[$section] -isnot [hashtable]) { [void]$problems.Add("Section '$section' must be a @{ } block."); continue }
        foreach ($key in $data[$section].Keys) {
            if (-not $script:ConfigSchema[$section].Contains($key)) {
                [void]$problems.Add("Unknown key '$section.$key'. Keys of $($section): $($script:ConfigSchema[$section].Keys -join ', ').")
                continue
            }
            $settings[$script:ConfigSchema[$section][$key]] = $data[$section][$key]
        }
    }
    foreach ($key in 'SearchIn', 'Rooms', 'ReportFormats') { $settings[$key] = @($settings[$key] | Where-Object { "$_" }) }
    foreach ($key in 'OutputPath', 'LogPath', 'RoomFile', 'MailboxFile') {
        if (-not [string]::IsNullOrWhiteSpace([string]$settings[$key])) { $settings[$key] = Resolve-MclPath -Path ([string]$settings[$key]) -Root $Root }
    }
    $settings.ConfigPath = [IO.Path]::GetFullPath($Path)
    foreach ($p in (Test-MclConfiguration -Configuration $settings).Problems) { [void]$problems.Add($p) }
    if ($problems.Count) { throw ("Invalid configuration ($Path):`n - " + ($problems -join "`n - ")) }
    return $settings
}

function Read-MclAddressFile {
    <#
        Addresses of a file: a text file with one address per line (# = comment), or a CSV file with a column
        PrimarySmtpAddress, EmailAddress, Mail, WindowsEmailAddress, UserPrincipalName, Address or Organizer
        (comma or semicolon). -AllowX500: X500 addresses (legacyExchangeDN) are kept too, also from a column
        LegacyExchangeDN (a list of organizers whose mailbox is deleted).
    #>
    param([Parameter(Mandatory = $true)][string]$Path, [switch]$AllowX500)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Address file not found: $Path" }
    $lines = @([IO.File]::ReadAllLines($Path) | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') })
    if (-not $lines.Count) { return @() }
    $columns = @('PrimarySmtpAddress', 'EmailAddress', 'Mail', 'WindowsEmailAddress', 'UserPrincipalName', 'Address', 'Organizer')
    if ($AllowX500) { $columns += 'LegacyExchangeDN' }
    $header = @($lines[0] -split '[;,]' | ForEach-Object { $_.Trim().Trim('"') })
    $found = @($columns | Where-Object { $header -contains $_ })
    if ($found.Count) {
        $delimiter = if ($lines[0].Contains(';')) { ';' } else { ',' }
        $rows = @($lines | ConvertFrom-Csv -Delimiter $delimiter)
        # The first address column of a row that holds a value (SMTP before X500).
        $values = foreach ($row in $rows) {
            $found | ForEach-Object { ([string]$row.$_).Trim() } | Where-Object { $_ } | Select-Object -First 1
        }
    }
    elseif ($AllowX500) {
        $values = foreach ($l in $lines) { $v = $l.Trim('"', ' '); if ($v -match $script:X500Pattern) { $v } else { ($v -split '[;,\s]')[0] } }
    }
    else { $values = @($lines | ForEach-Object { ($_ -split '[;,\s]')[0].Trim('"') }) }
    return @($values | Where-Object { $_ -match $script:SmtpPattern -or ($AllowX500 -and $_ -match $script:X500Pattern) } | Sort-Object -Unique)
}

function Split-MclAddressList {
    <# Addresses typed in one text: separated by ; , new lines or spaces (an X500 address keeps its spaces). #>
    param([AllowEmptyString()][AllowNull()][string[]]$Text)
    $all = foreach ($t in @($Text)) {
        foreach ($line in ([string]$t -split '[;\r\n]+')) {
            $v = $line.Trim()
            if (-not $v) { continue }
            if ($v -match $script:X500Pattern) { $v; continue }
            $v -split '[,\s]+' | Where-Object { $_ }
        }
    }
    return @($all | Select-Object -Unique)
}

function New-MclRequest {
    <#
        What one run does, from the command line or the window, with the defaults of the configuration.
        Mode: 'Organizers' (the meetings of -Organizer / -OrganizerFile) or 'Rooms' (every meeting of the rooms of
        -Room / -RoomFile, whatever its organizer; a series is then limited to its occurrences in the period).
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Settings,
        [string[]]$Organizer,
        [string]$OrganizerFile,
        [string[]]$Room,
        [string]$RoomFile,
        [Nullable[datetime]]$Start,
        [Nullable[datetime]]$End,
        [string]$Subject,
        [string[]]$MeetingId,
        [string[]]$SearchIn,
        [string[]]$Mailbox,
        [string]$MailboxFile,
        [ValidateSet('Report', 'Remove', 'Cancel', 'Restore', 'Transfer')][string]$Action = 'Report',
        [AllowNull()][string]$Comment,
        [string]$FromReport,
        [string]$NewOrganizer,
        [string]$TransferMethod,
        [Nullable[datetime]]$TransferFrom
    )

    $period = Get-MclDefaultPeriod -Settings $Settings
    $file = if ($OrganizerFile) { [IO.Path]::GetFullPath($OrganizerFile, (Get-Location).Path) } else { '' }
    $organizers = @(Split-MclAddressList $Organizer)
    if ($file -and (Test-Path -LiteralPath $file -PathType Leaf)) { $organizers = @($organizers + @(Read-MclAddressFile -Path $file -AllowX500) | Where-Object { $_ } | Select-Object -Unique) }
    $roomPath = if ($RoomFile) { [IO.Path]::GetFullPath($RoomFile, (Get-Location).Path) } else { '' }
    $rooms = @(Split-MclAddressList $Room | ForEach-Object { $_.ToLowerInvariant() })
    if ($roomPath -and (Test-Path -LiteralPath $roomPath -PathType Leaf)) { $rooms = @($rooms + @(Read-MclAddressFile -Path $roomPath | ForEach-Object { $_.ToLowerInvariant() }) | Where-Object { $_ } | Select-Object -Unique) }
    $mode = if (-not $organizers.Count -and -not $file -and ($rooms.Count -or $roomPath)) { 'Rooms' } else { 'Organizers' }
    $isComment = $null -ne $Comment -and $PSBoundParameters.ContainsKey('Comment')
    [pscustomobject]@{
        Mode           = $mode
        Organizer      = $organizers
        OrganizerFile  = $file
        Room           = $rooms
        RoomFile       = $roomPath
        Start          = if ($null -ne $Start) { ConvertTo-MclUtc -Date ([datetime]$Start) -TimeZone $Settings.TimeZone } else { $period.Start }
        End            = if ($null -ne $End) { ConvertTo-MclUtc -Date ([datetime]$End) -TimeZone $Settings.TimeZone -EndOfDay } else { $period.End }
        PeriodGiven    = $null -ne $Start -and $null -ne $End
        Subject        = [string]$Subject
        MeetingId      = @($MeetingId | ForEach-Object { ([string]$_ -split '[;,\s]+') } | Where-Object { $_ } | ForEach-Object { $_.ToUpperInvariant() } | Select-Object -Unique)
        SearchIn       = if ($mode -eq 'Rooms') { @('Rooms') } elseif ($SearchIn) { @($SearchIn | Select-Object -Unique) } else { @($Settings.SearchIn) }
        Mailboxes      = @(Split-MclAddressList $Mailbox)
        MailboxFile    = if ($MailboxFile) { [IO.Path]::GetFullPath($MailboxFile, (Get-Location).Path) } else { [string]$Settings.MailboxFile }
        Action         = $Action
        Comment        = if ($isComment) { $Comment } elseif ($Action -eq 'Transfer') { [string]$Settings.TransferComment } else { [string]$Settings.CancelComment }
        FromReport     = $FromReport
        NewOrganizer   = ([string]$NewOrganizer).Trim().ToLowerInvariant()
        TransferMethod = if ($TransferMethod) { $TransferMethod } else { [string]$Settings.TransferMethod }
        TransferFrom   = if ($null -ne $TransferFrom) { ConvertTo-MclUtc -Date ([datetime]$TransferFrom) -TimeZone $Settings.TimeZone } else { [datetime]::UtcNow }
    }
}

function Test-MclRequest {
    <# Checks a request (organizers or rooms, period, scopes, action, transfer) and lists every problem. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Request)

    $problems = [Collections.Generic.List[string]]::new()
    $r = $Request
    $mode = if ($r.PSObject.Properties['Mode']) { [string]$r.Mode } else { 'Organizers' }
    if ($r.FromReport) {
        if (-not (Test-Path -LiteralPath $r.FromReport)) { [void]$problems.Add("Report not found: $($r.FromReport)") }
        if ($r.Action -eq 'Report') { [void]$problems.Add('-FromReport replays a report: choose -Action Remove, Cancel, Transfer or Restore.') }
    }
    elseif ($r.Action -eq 'Restore') { [void]$problems.Add('-Action Restore needs -FromReport: the folder of the report of a Remove, Cancel or Transfer run.') }
    else {
        if ($mode -eq 'Rooms') {
            if ($r.RoomFile -and -not (Test-Path -LiteralPath $r.RoomFile -PathType Leaf)) { [void]$problems.Add("Room file not found: $($r.RoomFile)") }
            elseif (-not @($r.Room).Count) { [void]$problems.Add("No SMTP address in $($r.RoomFile): one room per line, or a CSV column PrimarySmtpAddress, Mail or UserPrincipalName.") }
            foreach ($a in @($r.Room)) { if ($a -notmatch $script:SmtpPattern) { [void]$problems.Add("Room '$a' is not an SMTP address.") } }
            if ($r.Action -in 'Remove', 'Cancel' -and -not $r.PeriodGiven) { [void]$problems.Add('Rooms: give the period of the action (-Start and -End): every meeting of the rooms in it is acted on.') }
        }
        else {
            if (@($r.Room).Count -or $r.RoomFile) { [void]$problems.Add('Give organizers (-Organizer, -OrganizerFile) or rooms (-Room, -RoomFile), not both.') }
            if ($r.OrganizerFile -and -not (Test-Path -LiteralPath $r.OrganizerFile -PathType Leaf)) { [void]$problems.Add("Organizer file not found: $($r.OrganizerFile)") }
            elseif ($r.OrganizerFile -and -not @($r.Organizer).Count) { [void]$problems.Add("No SMTP or X500 address in $($r.OrganizerFile): one per line, or a CSV column PrimarySmtpAddress, Mail, UserPrincipalName, Organizer or LegacyExchangeDN.") }
            elseif (-not @($r.Organizer).Count) { [void]$problems.Add('Give the organizer (-Organizer) or a list of organizers (-OrganizerFile): SMTP address, or the X500 address (legacyExchangeDN) of a deleted mailbox. Or the rooms (-Room, -RoomFile).') }
            foreach ($a in @($r.Organizer)) {
                if ($a -notmatch $script:SmtpPattern -and $a -notmatch $script:X500Pattern) { [void]$problems.Add("Organizer '$a' is neither an SMTP address nor an X500 address (/o=.../cn=...).") }
            }
            $scopes = @($r.SearchIn)
            if (-not $scopes.Count -or @($scopes | Where-Object { $_ -notin $script:SearchScopes }).Count) { [void]$problems.Add("Search in: one or more of $($script:SearchScopes -join ', ').") }
            if ($scopes -contains 'Mailboxes') {
                if (-not @($r.Mailboxes).Count -and -not $r.MailboxFile) { [void]$problems.Add('Mailboxes: give the list (-Mailbox) or a file (-MailboxFile, Search.MailboxFile).') }
                if ($r.MailboxFile -and -not (Test-Path -LiteralPath $r.MailboxFile -PathType Leaf)) { [void]$problems.Add("Mailbox file not found: $($r.MailboxFile)") }
                foreach ($a in @($r.Mailboxes)) { if ($a -notmatch $script:SmtpPattern) { [void]$problems.Add("Mailbox '$a' is not an SMTP address.") } }
            }
        }
        if ($r.End -le $r.Start) { [void]$problems.Add('The end of the period must be after its start.') }
        foreach ($id in @($r.MeetingId)) { if ($id -notmatch '^[0-9A-F]{40,}$') { [void]$problems.Add("Meeting ID '$id' is not an iCalUId (hexadecimal, column MeetingId of the report).") } }
    }
    if ($r.Action -eq 'Transfer') {
        if ($mode -eq 'Rooms' -and -not $r.FromReport) { [void]$problems.Add('Transfer moves the meetings of organizers: give -Organizer or -OrganizerFile, not rooms.') }
        if (-not $r.NewOrganizer) { [void]$problems.Add('Transfer: give the new organizer (-NewOrganizer), the SMTP address of a mailbox of the tenant.') }
        elseif ($r.NewOrganizer -notmatch $script:SmtpPattern) { [void]$problems.Add("New organizer '$($r.NewOrganizer)' is not an SMTP address.") }
        if ([string]$r.TransferMethod -notin $script:TransferMethods) { [void]$problems.Add("Transfer method must be one of: $($script:TransferMethods -join ', ').") }
        if ($r.TransferFrom -lt [datetime]::UtcNow.Date.AddDays(-1)) { [void]$problems.Add('Transfer: the date it starts from (-TransferFrom) cannot be in the past.') }
    }
    if ($r.Action -notin $script:Actions) { [void]$problems.Add("Action must be one of: $($script:Actions -join ', ').") }
    if ([string]$r.Comment -match '[<>]') { [void]$problems.Add('The message must be plain text (no < or >).') }
    [pscustomobject]@{ IsValid = $problems.Count -eq 0; Problems = @($problems) }
}