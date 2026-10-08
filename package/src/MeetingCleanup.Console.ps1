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
    Version : 1.3.0
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