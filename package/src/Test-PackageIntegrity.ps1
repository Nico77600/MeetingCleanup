#Requires -Version 7.4
<#
.SYNOPSIS
    Verifies the authenticated Meeting Cleanup package before any of its code runs.

.DESCRIPTION
    Invoked by Invoke-MeetingCleanup.ps1, by the module (import) and by the package builder. It requires a valid,
    timestamped owner signature on this script, on PackageIntegrity.psd1 and on every PowerShell file of the
    inventory, then compares the SHA256 of every inventory file (C# source, report template, HTML guides, README,
    notices) with the signed manifest. Missing, changed, unknown or reparse-point files fail closed.

    Not protected, by design: config\MeetingCleanup.config.psd1 (edited by the administrator) and the generated
    folders reports\ and logs\ (never executed).

.PARAMETER Root
    Package folder to verify, not its src subfolder. This script must belong to that package's src folder.

.PARAMETER Distribution
    Verifies a release package: the files copied by the package builder (CHANGELOG.md) are required and hashed.
    Refused together with -DevelopmentMode.

.PARAMETER DevelopmentMode
    For an unsigned continuous-integration build or a developer package. The files and the PowerShell signatures
    are still enforced when a manifest is present.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.3.1
#>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$Root, [switch]$Distribution, [switch]$DevelopmentMode)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($Distribution -and $DevelopmentMode) { throw 'Package integrity: -Distribution and -DevelopmentMode cannot be combined.' }
$signer = '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6'
# What to do, at the end of the refusals an unsigned or re-saved package meets first.
$releaseHint = 'Download the complete release ZIP again from https://github.com/Nico77600/MeetingCleanup/releases/latest, extract it and unblock it. A modified, re-saved or unsigned file cannot be started. A development or CI build is started from its folder with Invoke-MeetingCleanup.ps1 <parameters> -DevelopmentMode.'
$rootPath = [IO.Path]::GetFullPath($Root).TrimEnd('\')
$signed = @(
    'Invoke-MeetingCleanup.ps1'
    'MeetingCleanup.psm1'
    'MeetingCleanup.psd1'
    'src\Test-PackageIntegrity.ps1'
    'src\MeetingCleanup.Cleanup.ps1'
    'src\MeetingCleanup.Config.ps1'
    'src\MeetingCleanup.Console.ps1'
    'src\MeetingCleanup.Graph.ps1'
    'src\MeetingCleanup.Gui.ps1'
    'src\MeetingCleanup.Report.ps1'
    'src\MeetingCleanup.Restore.ps1'
    'src\MeetingCleanup.Search.ps1'
    'src\MeetingCleanup.Transfer.ps1'
)
$required = @($signed) + @(
    'src\MeetingCleanup.Native.cs'
    'templates\Report.template.html'
    'docs\MeetingCleanup-UserGuide.html'
    'docs\MeetingCleanup-Guide.html'
    'README.md'
    'THIRD-PARTY-NOTICES.md'
    'LICENSE'
)
# Copied from the repository root by the package build: hashed, required only in a distribution.
$distributionOnly = @('CHANGELOG.md')
# Repository-only documents are not loaded or copied into the distribution.
# These exact names (not an extension or directory wildcard) are the only unhashed source extras.
$sourceDocuments = @(
    'docs\MeetingCleanup-UserGuide.md'
    'docs\MeetingCleanup-Guide.md'
    'docs\images\gui-done-light.png'
    'docs\images\gui-occurrences-light.png'
    'docs\images\gui-progress-light.png'
    'docs\images\gui-restore-light.png'
    'docs\images\gui-rooms-light.png'
    'docs\images\gui-search-dark.png'
    'docs\images\gui-search-light.png'
    'docs\images\gui-transfer-light.png'
    'docs\images\readme-banner-dark.png'
    'docs\images\readme-banner-light.png'
    'docs\images\readme-how-it-works-dark.png'
    'docs\images\readme-how-it-works-light.png'
    'docs\images\readme-principles-dark.png'
    'docs\images\readme-principles-light.png'
    'docs\images\readme-transfer-dark.png'
    'docs\images\readme-transfer-light.png'
    'docs\images\report-dark.png'
    'docs\images\report-overview.png'
    'docs\images\report-transfers.png'
)
# Written or edited by the administrator; data, never executed.
$editableFiles = @('config\MeetingCleanup.config.psd1')
$inventory = @($required) + @($distributionOnly)
function Assert-PlainPath([string]$Path, [bool]$Directory = $false) {
    $item = Microsoft.PowerShell.Management\Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint -or $item.PSIsContainer -ne $Directory) {
        throw "Package integrity: unsafe path or unexpected type: $Path"
    }
    $parent = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($Path))
    while ($parent) {
        $ancestor = Microsoft.PowerShell.Management\Get-Item -LiteralPath $parent -Force -ErrorAction Stop
        if ($ancestor.Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw "Package integrity: reparse-point ancestor: $parent"
        }
        $parent = [IO.Path]::GetDirectoryName($parent)
    }
}
function Assert-OwnerSignature([string]$Path) {
    $signature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $Path -ErrorAction Stop
    if ([string]$signature.Status -cne 'Valid' -or $null -eq $signature.SignerCertificate -or
        $signature.SignerCertificate.Thumbprint -cne $signer -or $null -eq $signature.TimeStamperCertificate) {
        throw "Package integrity: expected trusted, timestamped owner signature: $Path. $releaseHint"
    }
}
function Test-Entry($Entry) {
    $Entry -is [hashtable] -and $Entry.Count -eq 2 -and $Entry.ContainsKey('Path') -and $Entry.ContainsKey('Sha256') -and
        $Entry.Path -is [string] -and $Entry.Sha256 -is [string] -and $Entry.Sha256 -cmatch '^[0-9A-F]{64}$'
}
Assert-PlainPath $rootPath $true
if (-not [string]::Equals([IO.Path]::GetFullPath($PSCommandPath),
    [IO.Path]::Combine($rootPath, 'src', 'Test-PackageIntegrity.ps1'), [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Package integrity: verifier must belong to this package root.'
}
Assert-PlainPath $PSCommandPath
Assert-OwnerSignature $PSCommandPath
$manifestPath = [IO.Path]::Combine($rootPath, 'PackageIntegrity.psd1')
if (-not [IO.File]::Exists($manifestPath)) { throw "Package integrity: PackageIntegrity.psd1 is missing. $releaseHint" }
Assert-PlainPath $manifestPath
Assert-OwnerSignature $manifestPath
# Authenticate the raw PSD1 before parsing its restricted data language.
$manifest = Microsoft.PowerShell.Utility\Import-PowerShellDataFile -LiteralPath $manifestPath -ErrorAction Stop
if ($manifest.Count -ne 5 -or -not $manifest.ContainsKey('SchemaVersion') -or
    -not $manifest.ContainsKey('Algorithm') -or -not $manifest.ContainsKey('SignerThumbprint') -or
    -not $manifest.ContainsKey('Version') -or -not $manifest.ContainsKey('Files') -or
    $manifest.SchemaVersion -isnot [int] -or $manifest.SchemaVersion -ne 1 -or $manifest.Algorithm -cne 'SHA256' -or
    $manifest.SignerThumbprint -cne $signer -or $manifest.Version -isnot [string] -or $manifest.Files -isnot [array]) {
    throw 'Package integrity: unsupported or incomplete manifest schema.'
}
if ($manifest.Files.Count -ne $inventory.Count) {
    throw "Package integrity: the manifest lists $($manifest.Files.Count) files, $($inventory.Count) expected."
}
$seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($entry in $manifest.Files) {
    if (-not (Test-Entry $entry) -or $entry.Path -cnotin $inventory -or -not $seen.Add($entry.Path)) {
        throw 'Package integrity: invalid, duplicate or unapproved inventory path/hash.'
    }
    $path = [IO.Path]::Combine($rootPath, $entry.Path)
    if ($entry.Path -cin $distributionOnly -and -not $Distribution -and
        -not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $path)) { continue }
    Assert-PlainPath $path
    $hash = (Microsoft.PowerShell.Utility\Get-FileHash -LiteralPath $path -Algorithm SHA256 -ErrorAction Stop).Hash
    if ($hash -cne $entry.Sha256) { throw "Package integrity: SHA256 mismatch: $($entry.Path)" }
    if ($entry.Path -cmatch '\.ps(m1|d1|1)$') { Assert-OwnerSignature $path }
}
foreach ($path in $inventory) {
    if (-not $seen.Contains($path)) { throw "Package integrity: missing inventory path: $path" }
}
$status = if ($Distribution) { 'Valid' } else { 'Valid (files only)' }
$queue = [Collections.Generic.Queue[string]]::new()
$queue.Enqueue($rootPath)
while ($queue.Count) {
    $directory = $queue.Dequeue()
    foreach ($item in Microsoft.PowerShell.Management\Get-ChildItem -LiteralPath $directory -Force -ErrorAction Stop) {
        $relative = [IO.Path]::GetRelativePath($rootPath, $item.FullName)
        if ($directory -ceq $rootPath -and $item.PSIsContainer -and $item.Name -cin @('reports', 'logs')) { continue }
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw "Package integrity: protected reparse point: $relative"
        }
        if ($item.PSIsContainer) { $queue.Enqueue($item.FullName); continue }
        # The owner-signed manifest itself (authenticated above, before it was parsed).
        if ($relative -ceq 'PackageIntegrity.psd1') { continue }
        if ($relative -cin $editableFiles -or $relative -cin $sourceDocuments) { continue }
        if ($relative -cin $inventory) { continue }
        throw "Package integrity: unexpected stable file: $relative. Move your own files (lists, certificates, copies of the configuration) out of the tool folder, for example to C:\Tools\MeetingCleanup-data\. A file of the package that is changed or missing: download the ZIP again from https://github.com/Nico77600/MeetingCleanup/releases/latest."
    }
}
[pscustomobject]@{
    Name    = 'MeetingCleanup'
    Version = $manifest.Version
    Files   = $manifest.Files.Count
    Status  = $status
    Signer  = $signer
}
