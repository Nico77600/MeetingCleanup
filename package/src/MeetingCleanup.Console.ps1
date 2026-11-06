<#
.SYNOPSIS
    Meeting Cleanup - console output and log file (dot-sourced by MeetingCleanup.psm1).

.DESCRIPTION
    Same rules as EAS OAuth Mailbox, Message Trace Report and Purview DLP Report:
      - ANSI colours are disabled when the output is redirected or NO_COLOR is set;
        MCL_FORCE_COLOR=1 forces them.
      - Icons: emoji in Windows Terminal / VS Code, symbols of the classic console fonts elsewhere.
        MCL_ICONS = Emoji | Symbols | Ascii forces a style.
      - Every line shown is also written to the daily log file, without colours or icons.
      - During a window run, the same lines are sent to the progress box of the window.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.3.1
#>

$script:C = @{ Reset = ''; Bold = ''; Dim = ''; Accent = ''; AccentBg = ''; Green = ''; Yellow = ''; Red = ''; Blue = ''; White = '' }
if ($env:MCL_FORCE_COLOR -eq '1' -or (-not [Console]::IsOutputRedirected -and -not $env:NO_COLOR)) {
    $e = [char]27
    $script:C = @{
        Reset = "$e[0m"; Bold = "$e[1m"; Dim = "$e[90m"; White = "$e[97m"
        Accent = "$e[38;2;214;62;115m"; AccentBg = "$e[48;2;177;31;75m$e[97m"
        Green = "$e[38;2;80;200;120m"; Yellow = "$e[38;2;240;200;90m"; Red = "$e[38;2;240;90;90m"; Blue = "$e[38;2;110;170;240m"
    }
}
$script:IconStyle = if ($env:MCL_ICONS -in 'Emoji', 'Symbols', 'Ascii') { $env:MCL_ICONS }
    elseif ([Console]::IsOutputRedirected) { 'Symbols' }
    elseif ($env:WT_SESSION -or $env:TERM_PROGRAM -eq 'vscode') { 'Emoji' }
    else { 'Symbols' }
$script:Dot = [char]0x00B7
$script:ProgressShown = $false
# The progress in course (Get-MclProgressEta): its label, when and where it began.
$script:ProgressEta = $null

function Get-MclIconSet {
    <# Icons of one console style. Symbols: only characters of the classic console fonts. #>
    param([Parameter(Mandatory = $true)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)

    $u = { param([int]$Code) [char]::ConvertFromUtf32($Code) }
    switch ($Style) {
        'Emoji' {
            return @{
                Logo = & $u 0x1F4C5; Ok = & $u 0x2705; Warn = (& $u 0x26A0) + [char]0xFE0F; Fail = & $u 0x274C; Info = & $u 0x1F539
                Skip = & $u 0x23E9; Key = & $u 0x1F511; Server = & $u 0x1F5A5; Shield = & $u 0x1F512; Room = & $u 0x1F3E2
                Mail = & $u 0x1F4E8; People = & $u 0x1F465; User = & $u 0x1F464; File = & $u 0x1F4C4; Log = & $u 0x1F4DD
                Report = & $u 0x1F4CA; Done = & $u 0x1F389; Target = & $u 0x1F3AF; Search = & $u 0x1F50E; Clock = & $u 0x23F3
                Calendar = & $u 0x1F4C6; Trash = (& $u 0x1F5D1) + [char]0xFE0F; Cancel = & $u 0x1F6AB; Cloud = (& $u 0x2601) + [char]0xFE0F; Refresh = & $u 0x1F504
            }
        }
        'Symbols' {
            return @{
                Logo = & $u 0x2666; Ok = & $u 0x221A; Warn = & $u 0x25B2; Fail = & $u 0x00D7; Info = & $u 0x2022
                Skip = & $u 0x00BB; Key = & $u 0x00A7; Server = & $u 0x2261; Shield = & $u 0x25CA; Room = & $u 0x2302
                Mail = '@'; People = & $u 0x2192; User = & $u 0x263A; File = & $u 0x25AC; Log = & $u 0x00B6
                Report = & $u 0x2261; Done = & $u 0x221A; Target = & $u 0x25D9; Search = & $u 0x25BA; Clock = & $u 0x25CB
                Calendar = & $u 0x25A1; Trash = & $u 0x00D7; Cancel = & $u 0x00F8; Cloud = & $u 0x2248; Refresh = & $u 0x00AB
            }
        }
        default {
            return @{
                Logo = '*'; Ok = '+'; Warn = '!'; Fail = 'x'; Info = '-'; Skip = '>'; Key = 'k'; Server = '='; Shield = 'o'; Room = '#'
                Mail = '@'; People = '&'; User = 'u'; File = '-'; Log = '='; Report = '='; Done = '*'; Target = 'o'; Search = '?'
                Clock = '~'; Calendar = '#'; Trash = 'x'; Cancel = '/'; Cloud = '~'; Refresh = 'r'
            }
        }
    }
}

function Get-MclFrameSet {
    <# Rounded corners in modern terminals (emoji style), square corners elsewhere (present in every console font). #>
    param([Parameter(Mandatory = $true)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)

    if ($Style -eq 'Ascii') {
        return @{ TopLeft = [char]'+'; TopRight = [char]'+'; BottomLeft = [char]'+'; BottomRight = [char]'+'; Horizontal = [char]'-'; Vertical = [char]'|' }
    }
    if ($Style -eq 'Symbols') {
        return @{ TopLeft = [char]0x250C; TopRight = [char]0x2510; BottomLeft = [char]0x2514; BottomRight = [char]0x2518; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
    }
    return @{ TopLeft = [char]0x256D; TopRight = [char]0x256E; BottomLeft = [char]0x2570; BottomRight = [char]0x256F; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
}

$script:Icons = Get-MclIconSet $script:IconStyle
$script:Frame = Get-MclFrameSet $script:IconStyle
$script:IconPad = if ($script:IconStyle -eq 'Emoji') { ' ' } else { '  ' }
$script:IconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { 1 }

function Get-MclIcon { param([Parameter(Mandatory = $true)][string]$Name) return $script:Icons[$Name] + $script:IconPad }

function Format-MclDuration {
    param([Parameter(Mandatory = $true)][double]$Seconds)

    $inv = [Globalization.CultureInfo]::InvariantCulture
    $t = [TimeSpan]::FromTicks([long]([Math]::Max(0.0, $Seconds) * 10000000))
    if ($t.TotalHours -ge 1) { return [string]::Format($inv, '{0} h {1:00} min', [int][Math]::Floor($t.TotalHours), $t.Minutes) }
    if ($t.TotalMinutes -ge 1) { return [string]::Format($inv, '{0} min {1:00} s', $t.Minutes, $t.Seconds) }
    return [string]::Format($inv, '{0:0.0} s', $t.TotalSeconds)
}

function Format-MclText {
    <# Text cut to a width with an ellipsis, padded to the width. #>
    param([AllowEmptyString()][AllowNull()][string]$Text, [int]$Width)
    $t = [string]$Text -replace '[\r\n\t]+', ' '
    if ($Width -le 0) { return $t }
    if ($t.Length -gt $Width) { return $t.Substring(0, [Math]::Max(0, $Width - 1)) + [char]0x2026 }
    return $t.PadRight($Width)
}

function Send-MclUi {
    <#
        Forwards a console line to the window while a window run is in progress: into the queue the window reads
        (background run, Ui.Queue) or to its sink (Ui.Sink).
    #>
    param([string]$Status, [string]$Text)
    $u = $script:Ui
    if (-not $u) { return }
    if ($u.Queue) { $u.Queue.Enqueue([string[]]@($Status, $Text)) }
    elseif ($u.Sink) { & $u.Sink $Status $Text }
}

function Start-MclLog {
    <# Opens (or continues) today's log file and deletes the log files older than the retention. #>
    param([Parameter(Mandatory = $true)][string]$Directory, [int]$RetentionDays = 30)

    Stop-MclLog
    [void][IO.Directory]::CreateDirectory($Directory)
    $script:LogPath = Join-Path $Directory ('MeetingCleanup_{0:yyyyMMdd}.log' -f (Get-Date))
    $stream = [IO.FileStream]::new($script:LogPath, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
    $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
    $writer.AutoFlush = $true
    # Synchronized: the window and its background run write to the same log.
    $script:LogWriter = [IO.TextWriter]::Synchronized($writer)
    $limit = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -LiteralPath $Directory -Filter 'MeetingCleanup_*.log' -File -ErrorAction SilentlyContinue |
        Where-Object LastWriteTime -lt $limit | Remove-Item -Force -ErrorAction SilentlyContinue
    return $script:LogPath
}

function Stop-MclLog {
    if ($script:LogWriter) { $script:LogWriter.Dispose(); $script:LogWriter = $null }
}

function Write-MclLog {
    <# One line in the log file only. The log never contains colours, icons, tokens or secrets. #>
    param(
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP')][string]$Level = 'INFO',
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Message
    )
    if ($script:LogWriter) { $script:LogWriter.WriteLine(('{0:yyyy-MM-ddTHH:mm:ss.fffzzz} [{1,-5}] {2}' -f (Get-Date), $Level, $Message)) }
}

function Get-MclConsoleWidth {
    try { $w = [Console]::WindowWidth; if ($w -ge 40) { return $w } } catch { }
    return 120
}

function Clear-MclProgress {
    <# Ends the live progress line (if any) so that the next line starts on a new row. #>
    if ($script:ProgressShown) {
        [Console]::Write("`r" + (' ' * [Math]::Max(10, (Get-MclConsoleWidth) - 1)) + "`r")
        $script:ProgressShown = $false
    }
}

function Write-MclBanner {
    <# Title card at the start of an execution, followed by the context rows (label -> @(Icon, Text)). #>
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Subtitle,
        [System.Collections.Specialized.OrderedDictionary]$Details
    )

    Write-MclLog 'STEP' "=== $Title v$($script:ToolVersion) ==="
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $v = $Details[$key]
            Write-MclLog 'INFO' ('{0}: {1}' -f $key, $(if ($v -is [array]) { $v[1] } else { $v }))
        }
    }
    if ($script:Quiet) { return }
    $C = $script:C; $F = $script:Frame; $width = 78
    $right = "v$($script:ToolVersion) $($script:Dot) Nicolas Fabert"
    $left = "  $($script:Icons.Logo)  $Title"
    $gap = [Math]::Max(1, $width - ($left.Length - $script:Icons.Logo.Length + $script:IconWidth) - $right.Length - 2)
    Write-Host ''
    Write-Host ('  {0}{1}{2}{3}{4}' -f $C.Accent, $F.TopLeft, [string]::new($F.Horizontal, $width), $F.TopRight, $C.Reset)
    Write-Host ('  {0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}' -f $C.Accent, $F.Vertical, $C.Reset, $C.Bold, $left, $C.Reset, [string]::new(' ', $gap), $C.Dim, $right, '  ', ($C.Accent + $F.Vertical), $C.Reset)
    if ($Subtitle) {
        Write-Host ('  {0}{1}{2}{3}{4}{5}{0}{6}{2}' -f $C.Accent, $F.Vertical, $C.Reset, $C.Dim, (Format-MclText "     $Subtitle" $width), $C.Reset, $F.Vertical)
    }
    Write-Host ('  {0}{1}{2}{3}{4}' -f $C.Accent, $F.BottomLeft, [string]::new($F.Horizontal, $width), $F.BottomRight, $C.Reset)
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $value = $Details[$key]
            $icon, $text = if ($value -is [array]) { (Get-MclIcon $value[0]), $value[1] } else { '   ', $value }
            Write-Host ('     {0}{1}{2,-11}{3} {4}' -f $icon, $C.Dim, $key, $C.Reset, $text)
        }
    }
}

function Write-MclStep {
    <# Step header with a coloured number pill and an icon:  ─ 3/6 ─ 🔎  Search #>
    param(
        [Parameter(Mandatory = $true)][int]$Number,
        [Parameter(Mandatory = $true)][int]$Total,
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Icon = 'Info'
    )

    Write-MclLog 'STEP' "[$Number/$Total] $Title"
    $script:ProgressEta = $null
    Send-MclUi 'Step' "[$Number/$Total] $Title"
    if ($script:Quiet) { return }
    Clear-MclProgress
    $C = $script:C
    Write-Host ''
    Write-Host ('  {0} {1}/{2} {3} {4}{5}{6}{3}' -f $C.AccentBg, $Number, $Total, $C.Reset, (Get-MclIcon $Icon), $C.Bold, $Title)
}

function Write-MclItem {
    <# One indented result line with a status icon, also written to the log and to the window. #>
    param(
        [ValidateSet('Ok', 'Warn', 'Fail', 'Info', 'Skip')][string]$Status = 'Info',
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text,
        [string]$Icon
    )

    $level = @{ Ok = 'OK'; Warn = 'WARN'; Fail = 'ERROR'; Info = 'INFO'; Skip = 'INFO' }[$Status]
    Write-MclLog $level $Text
    Send-MclUi $Status $Text
    if ($script:Quiet) { return }
    Clear-MclProgress
    $color = @{ Ok = $script:C.Green; Warn = $script:C.Yellow; Fail = $script:C.Red; Info = ''; Skip = $script:C.Dim }[$Status]
    $symbol = Get-MclIcon $(if ($Icon) { $Icon } else { $Status })
    $textColor = if ($Status -in 'Warn', 'Fail', 'Skip') { $color } else { '' }
    Write-Host ('      {0}{1}{2}{3}{4}{2}' -f $color, $symbol, $script:C.Reset, $textColor, $Text)
}

function Format-MclTimeLeft {
    <#
        The time left of a progress, rounded as a person would say it: a few seconds, about 25 s (5 s steps under a
        minute), about 1 min 30 s (10 s steps under 5 minutes), about 12 min, about 1 h 05 min.
    #>
    param([Parameter(Mandatory = $true)][double]$Seconds)

    $inv = [Globalization.CultureInfo]::InvariantCulture
    if ($Seconds -lt 10) { return 'a few seconds left' }
    $away = [MidpointRounding]::AwayFromZero
    $r = [int]$(if ($Seconds -lt 60) { [Math]::Ceiling($Seconds / 5) * 5 } elseif ($Seconds -lt 300) { [Math]::Round($Seconds / 10, $away) * 10 } else { [Math]::Round($Seconds / 60, $away) * 60 })
    if ($r -lt 60) { return [string]::Format($inv, 'about {0} s left', $r) }
    $h = [int][Math]::Floor($r / 3600); $m = [int][Math]::Floor(($r % 3600) / 60); $s = $r % 60
    if ($h) { return [string]::Format($inv, 'about {0} h {1:00} min left', $h, $m) }
    if ($s) { return [string]::Format($inv, 'about {0} min {1:00} s left', $m, $s) }
    return [string]::Format($inv, 'about {0} min left', $m)
}

function Get-MclProgressEta {
    <#
        Time left of the progress in course, from its speed since it began; empty until it can be told (2 s and
        2 % of progress since its first value). Another label (the counts aside), or a value going back, is a
        new progress. A step starts with none (Write-MclStep). -Now: tests.
    #>
    param([Parameter(Mandatory = $true)][double]$Fraction, [AllowEmptyString()][string]$Text, [datetime]$Now = [datetime]::UtcNow)

    $key = $Text -replace '[\d\s,.\u00A0\u202F/]+', ''
    $s = $script:ProgressEta
    if (-not $s -or $s.Key -ne $key -or $Fraction -lt $s.Last) {
        $script:ProgressEta = @{ Key = $key; Start = $Now; From = $Fraction; Last = $Fraction; Left = -1.0; At = $Now }
        return ''
    }
    $s.Last = $Fraction
    $done = $Fraction - $s.From
    $elapsed = ($Now - $s.Start).TotalSeconds
    if ($Fraction -ge 1 -or $done -lt 0.02 -or $elapsed -lt 2) { return '' }
    $left = $elapsed / $done * (1 - $Fraction)
    # Graph answers come in bursts (16 calls in flight): half the new figure, half the last one brought forward.
    if ($s.Left -ge 0) { $left = 0.5 * $left + 0.5 * [Math]::Max(0.0, $s.Left - ($Now - $s.At).TotalSeconds) }
    $s.Left = $left; $s.At = $Now
    return Format-MclTimeLeft $left
}

function Write-MclProgress {
    <#
        Live progress line, rewritten in place (interactive console); sent to the window during a window run
        (fraction|text|time left).
              ⏳  ███████░░░░░  58%  1,077/1,858 mailboxes searched · about 40 s left
    #>
    param([Parameter(Mandatory = $true)][double]$Fraction, [Parameter(Mandatory = $true)][string]$Text)

    $left = Get-MclProgressEta -Fraction $Fraction -Text $Text
    Send-MclUi 'Progress' ('{0}|{1}|{2}' -f $Fraction.ToString('0.000', [Globalization.CultureInfo]::InvariantCulture), $Text, $left)
    if ($script:Quiet -or [Console]::IsOutputRedirected) { return }
    $C = $script:C
    $percent = [int][Math]::Floor(100 * [Math]::Min(1.0, [Math]::Max(0.0, $Fraction)))
    $filled = [int][Math]::Round(12 * $percent / 100.0)
    $bar = $C.Accent + [string]::new([char]0x2588, $filled) + $C.Dim + [string]::new([char]0x2591, 12 - $filled) + $C.Reset
    $line = if ($left) { "$Text $($script:Dot) $left" } else { $Text }
    $plain = Format-MclText $line ([Math]::Max(10, (Get-MclConsoleWidth) - 30))
    [Console]::Write(("`r      {0}{1} {2,3}%  {3}{4}{5}" -f (Get-MclIcon 'Clock'), $bar, $percent, $C.Dim, $plain.TrimEnd(), $C.Reset))
    $script:ProgressShown = $true
}

function Write-MclTable {
    <#
        Aligned table, one row per object, with a status icon in front of each row.
        Columns: @{ Name = 'Header'; Property = 'PropertyName'; Width = 20; Align = 'Right' } - Width 0 = the rest of the console.
        StatusProperty: Ok | Warn | Fail | Info | Skip (colour and icon of the row).
    #>
    param(
        [Parameter(Mandatory = $true)][object[]]$Columns,
        [AllowEmptyCollection()][AllowNull()][object[]]$Rows,
        [string]$StatusProperty = 'Status',
        [int]$Indent = 6,
        [int]$MaxWidth = 170
    )

    if (-not $Rows -or -not $Rows.Count) { return }
    foreach ($row in $Rows) { Write-MclLog 'INFO' (($Columns | ForEach-Object { "$($_.Name)=$([string]$row.($_.Property))" }) -join ' | ') }
    if ($script:Quiet) { return }
    Clear-MclProgress
    $C = $script:C
    $consoleWidth = [Math]::Min($MaxWidth, (Get-MclConsoleWidth) - 1)
    if ($consoleWidth -lt 80) { $consoleWidth = 120 }
    $fixed = [int](($Columns | ForEach-Object { [int]$_['Width'] } | Measure-Object -Sum).Sum) + 2 * $Columns.Count
    $last = [Math]::Max(20, $consoleWidth - $Indent - 3 - $fixed)
    $pad = ' ' * $Indent
    $cell = {
        param($col, $text)
        $w = if ([int]$col['Width']) { [int]$col['Width'] } else { $last }
        if ($col['Align'] -eq 'Right') { (Format-MclText $text $w).Trim().PadLeft($w) } else { Format-MclText $text $w }
    }
    $header = ($Columns | ForEach-Object { & $cell $_ $_['Name'] }) -join '  '
    Write-Host ('{0}{1}{2}{3}{4}' -f $pad, $C.Dim, (' ' * ($script:IconWidth + $script:IconPad.Length)), $header.TrimEnd(), $C.Reset)
    foreach ($row in $Rows) {
        $status = [string]$row.$StatusProperty
        if ($status -notin 'Ok', 'Warn', 'Fail', 'Info', 'Skip') { $status = 'Info' }
        $color = @{ Ok = $C.Green; Warn = $C.Yellow; Fail = $C.Red; Info = $C.Blue; Skip = $C.Dim }[$status]
        $cells = foreach ($col in $Columns) { & $cell $col ([string]$row.($col.Property)) }
        $textColor = if ($status -eq 'Skip') { $C.Dim } else { '' }
        Write-Host ('{0}{1}{2}{3}{4}{5}{3}' -f $pad, $color, (Get-MclIcon $status), $C.Reset, $textColor, (($cells -join '  ').TrimEnd()))
    }
}

function Write-MclSummary {
    <# Final summary card (label -> @(Icon, Text)). #>
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][System.Collections.Specialized.OrderedDictionary]$Values,
        [ValidateSet('Ok', 'Warn', 'Fail')][string]$Status = 'Ok'
    )

    foreach ($key in $Values.Keys) {
        $v = $Values[$key]
        Write-MclLog 'INFO' ('Summary - {0}: {1}' -f $key, $(if ($v -is [array]) { $v[1] } else { $v }))
    }
    if ($script:Quiet) { return }
    Clear-MclProgress
    $C = $script:C; $F = $script:Frame; $width = 78
    $color = @{ Ok = $C.Green; Warn = $C.Yellow; Fail = $C.Red }[$Status]
    $icon = $script:Icons[@{ Ok = 'Done'; Warn = 'Warn'; Fail = 'Fail' }[$Status]]
    $head = " $icon  $Title "
    $rest = [Math]::Max(2, $width - 1 - ($head.Length - $icon.Length + $script:IconWidth))
    Write-Host ''
    Write-Host ('  {0}{1}{2}{3}{4}{0}{5}{6}{7}' -f $color, $F.TopLeft, $F.Horizontal, $C.Bold, $head, ($C.Reset + $color), ([string]::new($F.Horizontal, $rest) + $F.TopRight), $C.Reset)
    foreach ($key in $Values.Keys) {
        $value = $Values[$key]
        $rowIcon, $text = if ($value -is [array]) { (Get-MclIcon $value[0]), $value[1] } else { '   ', $value }
        Write-Host ('    {0}{1}{2,-10}{3} {4}' -f $rowIcon, $C.Dim, $key, $C.Reset, $text)
    }
    Write-Host ('  {0}{1}{2}{3}{4}' -f $color, $F.BottomLeft, [string]::new($F.Horizontal, $width), $F.BottomRight, $C.Reset)
    Write-Host ''
}

function Get-MclActionText {
    <# The action in words, for the banner, the window and the report. #>
    param([Parameter(Mandatory = $true)][string]$Action)
    switch ($Action) {
        'Remove' { 'Remove silently: the copies of the attendees and the rooms, no message' }
        'Cancel' { 'Cancel and clean: the organizer cancels (message to the attendees), then the copies left are removed' }
        'Restore' { 'Restore: the copies removed by a run come back from Recoverable Items, no message' }
        'Transfer' { 'Transfer: the meetings move to a new organizer (moved by Exchange Online, or re-created by him)' }
        default { 'Report only: nothing is changed' }
    }
}

function Get-MclScopeText {
    param([Parameter(Mandatory = $true)][string]$Scope)
    switch ($Scope) {
        'Organizer' { "organizer's calendar" }
        'Rooms' { 'room mailboxes' }
        'Mailboxes' { 'mailboxes of the list' }
        'AllMailboxes' { 'every mailbox' }
        default { $Scope }
    }
}

function Format-MclOrganizerList {
    <# The organizers in one line: all of them up to 3, else the first ones and the count (and the file). #>
    param([string[]]$Organizer, [string]$File)
    $list = @($Organizer)
    $text = if ($list.Count -le 3) { $list -join ', ' } else { '{0} organizers ({1}, ...)' -f $list.Count, (($list | Select-Object -First 2) -join ', ') }
    if ($File) { $text += " $($script:Dot) file $([IO.Path]::GetFileName($File))" }
    return $text
}

function Write-MclRunBanner {
    <# Title card of a command-line run: organizer, meetings, where to search, action, application, report and log. #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Settings,
        [Parameter(Mandatory = $true)][pscustomobject]$Request,
        [string]$LogPath,
        [switch]$NoReport
    )

    $dot = $script:Dot
    $banner = [ordered]@{}
    if ($Request.FromReport) {
        $banner['Plan'] = @('File', $(if ($Request.Action -eq 'Restore') { "copies removed by the run of $($Request.FromReport)" } else { "meetings reviewed in $($Request.FromReport)" }))
    }
    else {
        if ((Get-MclProperty $Request 'Mode') -eq 'Rooms') {
            $list = @($Request.Room)
            $text = if ($list.Count -le 3) { $list -join ', ' } else { '{0} rooms ({1}, ...)' -f $list.Count, (($list | Select-Object -First 2) -join ', ') }
            if ($Request.RoomFile) { $text += " $dot file $([IO.Path]::GetFileName($Request.RoomFile))" }
            $banner['Rooms'] = @('Room', "$text $dot every organizer")
        }
        else { $banner['Organizer'] = @('User', (Format-MclOrganizerList -Organizer $Request.Organizer -File $Request.OrganizerFile)) }
        $what = "from $(Format-MclDate $Request.Start $Settings.TimeZone -DateOnly) to $(Format-MclDate $Request.End $Settings.TimeZone -DateOnly -PeriodEnd)"
        if ($Request.Subject) { $what += " $dot subject contains '$($Request.Subject)'" }
        if (@($Request.MeetingId).Count) { $what += " $dot $(@($Request.MeetingId).Count) meeting ID(s)" }
        $banner['Meetings'] = @('Calendar', $what)
        if ((Get-MclProperty $Request 'Mode') -eq 'Rooms') { $banner['Search in'] = @('Search', 'these rooms only (a series: its occurrences in the period)') }
        else { $banner['Search in'] = @('Search', ((@($Request.SearchIn) | ForEach-Object { Get-MclScopeText $_ }) -join " $dot ")) }
    }
    $banner['Action'] = @($(switch ($Request.Action) { 'Remove' { 'Trash' } 'Cancel' { 'Cancel' } 'Restore' { 'Refresh' } 'Transfer' { 'People' } default { 'Report' } }), (Get-MclActionText $Request.Action))
    if ($Request.Action -eq 'Transfer') {
        $banner['New organizer'] = @('User', ('{0} {1} method {2}{3}' -f $Request.NewOrganizer, $dot, $Request.TransferMethod, $(if ($Request.TransferFrom -gt [datetime]::UtcNow.AddMinutes(5)) { " $dot from $(Format-MclDate $Request.TransferFrom $Settings.TimeZone)" } else { '' })))
    }
    $banner['Tenant'] = @('Cloud', $(if ($Settings.Organization) { "$($Settings.Organization) $dot $($Settings.TenantId)" } else { $Settings.TenantId }))
    $banner['App'] = @('Key', "$($Settings.AppId) $dot $(if ($Settings.AuthMode -eq 'Certificate') { "certificate $($Settings.CertificateThumbprint)" } else { "client secret (`$env:$($Settings.ClientSecretVariable) or prompt)" })")
    if ($Request.Action -eq 'Restore') { $banner['Exchange'] = @('Server', $(if ($Settings.RestoreConnection -eq 'Interactive') { "Exchange Online PowerShell as an administrator $($Settings.RestoreUser)" } else { 'Exchange Online PowerShell as the application (role Mailbox Import Export)' })) }
    if ($Request.Action -eq 'Transfer' -and $Request.TransferMethod -ne 'Recreate') { $banner['Exchange'] = @('Server', 'Exchange Online PowerShell for the meetings whose organizer has a mailbox (Invoke-ChangeMeetingOrganizer)') }
    $banner['Report'] = @('Report', $(if ($NoReport) { 'backup and Summary.json only (-NoReport)' } else { $Settings.OutputPath }))
    if ($LogPath) { $banner['Log'] = @('Log', $LogPath) }
    Write-MclBanner -Title 'Meeting Cleanup' -Subtitle "Exchange Online $dot meetings of organizers or rooms, in every calendar" -Details $banner
}

function Write-MclMeetingTable {
    <# The meetings of a result, one line each: start, kind, organizer copy, copies, result, (organizer), subject. #>
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Meetings)

    $many = @($Meetings | ForEach-Object Organizer | Select-Object -Unique).Count -gt 1
    $rows = foreach ($m in $Meetings) {
        # One line per mailbox: an occurrence copy is counted once for its mailbox.
        $copies = @($m.Copies | Where-Object { $_.EventId -and $_.Role -in 'Organizer', 'Attendee', 'Room' } | Group-Object Mailbox | ForEach-Object { $_.Group[0] })
        $rooms = @($copies | Where-Object Role -eq 'Room').Count
        [pscustomobject]@{
            Status    = switch ($m.Status) { { $_ -in 'Removed', 'Cancelled', 'Restored', 'Transferred' } { 'Ok' } { $_ -in 'Partial', 'Not restorable' } { 'Warn' } 'Failed' { 'Fail' } { $_ -in 'Skipped', 'Nothing to do' } { 'Skip' } default { 'Info' } }
            Start     = $m.StartText
            Kind      = if ((Get-MclProperty $m 'Scope') -eq 'Occurrences') { '{0} occ.' -f $m.Occurrences } else { $m.Kind }
            Subject   = $m.Subject
            Organizer = $m.OrganizerCopy
            Who       = if ($m.OrganizerName) { $m.OrganizerName } else { $m.Organizer }
            Copies    = '{0} ({1} room{2})' -f $copies.Count, $rooms, $(if ($rooms -eq 1) { '' } else { 's' })
            Result    = $m.Status
        }
    }
    $columns = @(
        @{ Name = 'Start'; Property = 'Start'; Width = 16 }
        @{ Name = 'Kind'; Property = 'Kind'; Width = 7 }
        @{ Name = 'Organizer copy'; Property = 'Organizer'; Width = 15 }
        @{ Name = 'Copies'; Property = 'Copies'; Width = 13 }
        @{ Name = 'Result'; Property = 'Result'; Width = 14 }
    )
    if ($many) { $columns += @{ Name = 'Organizer'; Property = 'Who'; Width = 22 } }
    $columns += @{ Name = 'Subject'; Property = 'Subject'; Width = 0 }
    Write-MclTable -Rows @($rows) -Columns $columns
}

function Write-MclRunSummary {
    <# Final card of a command-line run: status, meetings, copies, report, log and what to do next. #>
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Result,
        [string]$ReportText = 'none (-NoReport)',
        [string]$LogPath
    )

    $dot = $script:Dot
    $n = $Result.Counts
    $values = [ordered]@{}
    $values['Status'] = @($(switch ($Result.Status) { 'Completed' { 'Ok' } 'Failed' { 'Fail' } default { 'Warn' } }), "$($Result.Status) $dot $(Get-MclActionText $Result.Action)")
    if ($n.Organizers -gt 1) { $values['Organizers'] = @('People', ('{0} {1} {2} with meetings' -f $n.Organizers, $dot, @($Result.Meetings | ForEach-Object OrganizerKey | Select-Object -Unique).Count)) }
    $values['Meetings'] = @('Calendar', ('{0} found {1} {2} series {1} {3} selected' -f $n.Meetings, $dot, $n.Series, $n.Selected))
    if ($Result.Action -eq 'Restore') {
        $values['Restored'] = @('Refresh', ('{0} restored {1} {2} already present {1} {3} not found {1} {4} not restorable {1} {5} failed' -f $n.Restored, $dot, $n.AlreadyPresent, $n.NotFound, $n.NotRestorable, $n.Failed))
    }
    else {
        $values['Copies'] = @('People', ('{0} in {1} mailbox(es) {2} {3} in rooms {2} {4} attendee(s) not processed' -f $n.Copies, $n.Mailboxes, $dot, $n.RoomCopies, $n.NotProcessed))
        if ($Result.Action -eq 'Transfer') {
            $selected = @($Result.Meetings | Where-Object Selected)
            $values['Transferred'] = @('People', ('{0} transferred to {1} {2} {3} by Exchange Online {2} {4} re-created {2} {5} not transferred {2} {6} failed' -f $n.Transferred, (Get-MclProperty $Result 'NewOrganizer'), $dot, @($selected | Where-Object { $_.Status -eq 'Transferred' -and $_.TransferMethod -eq 'Native' }).Count, @($selected | Where-Object { $_.Status -in 'Transferred', 'Partial' -and $_.TransferMethod -eq 'Recreate' }).Count, @($selected | Where-Object Status -eq 'Skipped').Count, @($selected | Where-Object Status -eq 'Failed').Count))
            $values['Done'] = @('Trash', ('old copies: {0} removed {1} {2} cancelled by their old organizer {1} {3} failed' -f $n.Removed, $dot, $n.Cancelled, $n.Failed))
        }
        elseif ($Result.Action -ne 'Report') {
            $values['Done'] = @('Trash', ('{0} removed {1} {2} cancelled {1} {3} already gone {1} {4} kept {1} {5} failed' -f $n.Removed, $dot, $n.Cancelled, $n.AlreadyGone, $n.Kept, $n.Failed))
        }
        if ($n.OccurrenceCopies) { $values['Occurrences'] = @('Calendar', ('{0} series limited to the period: {1} occurrence copies (not restorable once removed)' -f @($Result.Meetings | Where-Object { (Get-MclProperty $_ 'Scope') -eq 'Occurrences' }).Count, $n.OccurrenceCopies)) }
    }
    if ($Result.PSObject.Properties['BackupFile'] -and $Result.BackupFile) { $values['Backup'] = @('Shield', $Result.BackupFile) }
    if ($Result.Error) { $values['First issue'] = @('Fail', $Result.Error) }
    $values['Duration'] = @('Clock', (Format-MclDuration $Result.DurationSeconds))
    $values['Report'] = @('Report', $ReportText)
    if ($LogPath) { $values['Log'] = @('Log', $LogPath) }
    # The folder of this run, for the command to give next (-FromReport).
    $folder = if ($Result.PSObject.Properties['BackupFile'] -and $Result.BackupFile) { Split-Path $Result.BackupFile -Parent } elseif ($ReportText -match '[\\/]') { Split-Path $ReportText -Parent } else { '<report folder>' }
    $values['Next'] = @('Info', $(switch ($Result.Status) {
                'Completed' {
                    if ($Result.Action -eq 'Report') {
                        if ($n.Meetings) { "Review the report, then: -FromReport '$folder' -Action Remove (or Cancel, or Transfer -NewOrganizer <address>), with -MeetingId <id> to act on some of them only." }
                        else { 'Nothing found: widen the period or search more mailboxes (-SearchIn Rooms, Mailboxes, AllMailboxes).' }
                    }
                    elseif ($Result.Action -eq 'Restore') { 'Nothing to do. The copies are back, answered again; no message was sent.' }
                    elseif ($Result.Action -eq 'Transfer') { "Nothing to do. The attendees answer the invitation of $(Get-MclProperty $Result 'NewOrganizer') again (re-created meetings); a Teams link may need to be renewed by the new organizer." }
                    elseif ($Result.Action -eq 'Cancel') {
                        if (@($Result.Meetings | Where-Object Status -eq 'Removed').Count) { "Nothing to do. A cancellation cannot be undone (the attendees received it); the meetings removed silently (no organizer copy) can: -Action Restore -FromReport '$folder'." }
                        else { 'Nothing to do. A cancellation cannot be undone: the attendees received it.' }
                    }
                    else { "Nothing to do. To undo it, within the retention of deleted items (14 days by default): -Action Restore -FromReport '$folder'." }
                }
                'Failed' { 'Read the first issue above and the log; nothing more was changed after it.' }
                default {
                    if ($Result.Action -eq 'Restore' -and $n.NotRestorable -and -not ($n.Failed + $n.NotFound)) { 'Nothing more to do: a cancelled or transferred meeting, or an occurrence, cannot be restored.' }
                    elseif ($Result.Action -eq 'Transfer' -and @($Result.Meetings | Where-Object { $_.Status -eq 'Failed' -and $_.TransferMethod -eq 'Native' }).Count) { "Open the report. A meeting Exchange Online could not move can be re-created: -FromReport '$folder' -Action Transfer -NewOrganizer <address> -TransferMethod Recreate -MeetingId <id>." }
                    else { 'Open the report: each copy gives its result. Running the same command again retries what is left.' }
                }
            }))
    $title = switch ($Result.Status) { 'Completed' { switch ($Result.Action) { 'Report' { 'Search finished' } 'Restore' { 'Restore finished' } 'Transfer' { 'Transfer finished' } default { 'Cleanup finished' } } } 'Failed' { 'Run failed' } default { 'Finished with warnings' } }
    $card = switch ($Result.Status) { 'Completed' { 'Ok' } 'Failed' { 'Fail' } default { 'Warn' } }
    Write-MclSummary -Title $title -Values $values -Status $card
}
# SIG # Begin signature block
# MIImIwYJKoZIhvcNAQcCoIImFDCCJhACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAM/dmQQ4stlhr3
# MG9a8pksoK6zMoPjWM/ppL9rJQi5GqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCADZRKtunaE
# TRtQBaBRVVdu3lfaRuqyjDpna58hUhBhNzANBgkqhkiG9w0BAQEFAASCAgBIs+vr
# 2OpJeKTrZ7d2aQfiCtrI/3FTGr0ABxzVznePWCGhP0jMbawI0f3OLFoAMsVyo2eS
# Jc/FiAG1aUmP1LFOnxCA8cmdgoP/4kkXt19aJFiJTKAtvxr5ELxC7pE1pYzPvsSv
# 4OEqijSFfAoGgu4otHcv7McOHws59y+4D+Z0X7IU3+DtFQr4dIP5hVJOfmLSOvwA
# CK7hA5aCoGnlQ6+kTkoPBeEjJ0CZo7+5tuaLNjjCxXSvAR2cRXAyjv8ffHmQNNkM
# H83Vhde7c1yKI4O4HPRWr5pXshXLWKiuNrjZi3wwpTBCItG6pocFqnQdsnZtDQSe
# 6F9HjzsLeFOaMLJDgtc+FA1E9xBHqLQjb+Gc1obTPTSy9/UIdBmOb0/qBiccWoRz
# IenVFfc+jmL8/OuJaV2Cx1LfYqA/+REOyeXt8c9rSU1/OBDzsgM9v2Xf9H1oNjWq
# 2DQONfyRxsIKFRHXWzY+lSGmyAgVmHld8Njoyq9APDLZuMaPYV2XyN5TZhvasDly
# bIHSsTLHmyw92SmY1I+nCk7MSDsRrzknYnhWA0d/WKVJCjT9ZWUgrY0TE4JP4gYf
# e+Fq9KpJckKsVeGh5pJu/fZUZPCkBCyHVeC+4erdpWztNHw3armlAIpHoviRIDw5
# zSOZuZ/VXdG2jL6Bm0X0n0WUaIfj3HY64L8zEqGCFMcwghTDBgorBgEEAYI3AwMB
# MYIUszCCFK8GCSqGSIb3DQEHAqCCFKAwghScAgEDMQ0wCwYJYIZIAWUDBAIBMIGB
# BgsqhkiG9w0BCRABBKByBHAwbgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIFz5Ls327GFvXVWr+dKPZsIKN/J4ZAEYtWvDakwAaHSAAghdmEkRxPNl
# eBgPMjAyNjExMDYwODQyMjNaMAMCAQECCDA+ZvTQ7h2PoIIRpzCCBPMwggLboAMC
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
# CSqGSIb3DQEJBTEPFw0yNjExMDYwODQyMjNaMCgGCSqGSIb3DQEJNDEbMBkwCwYJ
# YIZIAWUDBAIBoQoGCCqGSM49BAMCMC8GCSqGSIb3DQEJBDEiBCB00kViMj50lslr
# iX6xi6vzdWXQgK12PzXO2MdoMiurqzCByAYLKoZIhvcNAQkQAi8xgbgwgbUwgbIw
# ga8EIFdP6heq+l58aiubCMmYsXNeRQs4c8pE5qJMoJMOY/OwMIGKMHakdDByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7jMAoGCCqGSM49BAMCBEgw
# RgIhAPcJnAmnWD82RYdsJ4cQLXNyGljWeIG9sZ9ksGePYOuXAiEAgNg8oHLKMvNW
# bjnZX6F/tqNUb04S8xlD4LvLqWOHM3A=
# SIG # End signature block
