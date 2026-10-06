<#
.SYNOPSIS
    Copies the files needed to run Meeting Cleanup into a separate folder, ready to be zipped.

.DESCRIPTION
    The package contains only what Invoke-MeetingCleanup.ps1 needs at run time, plus the HTML guide:
        Invoke-MeetingCleanup.ps1, MeetingCleanup.psd1, MeetingCleanup.psm1, src\, config\, templates\,
        docs\MeetingCleanup-Guide.html, README.md, CHANGELOG.md, LICENSE, THIRD-PARTY-NOTICES.md
    The HTML guide is rebuilt first from docs\MeetingCleanup-Guide.md (tools\Build-Documentation.ps1): it is
    self-contained (images inline), so the Markdown source and the images are not copied.
    It never copies reports\, logs\, artifacts\, tests\ (with the simulated tenant) or tools\.

    The configuration is copied as delivered (empty tenant and application). The script checks the content
    of the package and that the module loads from it.

.PARAMETER Destination
    Package folder. Default: package\MeetingCleanup-<version>, next to the tool folder.

.PARAMETER Force
    Replace the destination folder if it already contains a package. A folder that contains reports\ or
    logs\ (a package that has been run) is never replaced.

.EXAMPLE
    .\tools\New-MeetingCleanupPackage.ps1
    Creates ..\package\MeetingCleanup-<version>.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [string]$Destination,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$version = (Import-PowerShellDataFile -LiteralPath (Join-Path $root 'MeetingCleanup.psd1')).ModuleVersion
if (-not $Destination) { $Destination = Join-Path (Split-Path $root -Parent) "package\MeetingCleanup-$version" }
$Destination = [IO.Path]::GetFullPath($Destination, (Get-Location).Path).TrimEnd('\')
$rootPrefix = [IO.Path]::GetFullPath($root).TrimEnd('\') + '\'
if (($Destination + '\').StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase) -or $rootPrefix.StartsWith($Destination + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "The destination must be outside the tool folder: $Destination"
}
if (Test-Path -LiteralPath $Destination) {
    if (-not $Force) { throw "The destination already exists: $Destination. Use -Force to replace it." }
    if (-not (Test-Path -LiteralPath (Join-Path $Destination 'Invoke-MeetingCleanup.ps1'))) { throw "The destination is not a Meeting Cleanup package, it is not replaced: $Destination" }
    foreach ($used in 'reports', 'logs') {
        if (Test-Path -LiteralPath (Join-Path $Destination $used)) { throw "The destination contains a $used folder (a package that has been run), it is not replaced: $Destination" }
    }
    Remove-Item -LiteralPath $Destination -Recurse -Force
}

# ---- HTML guide, rebuilt from the Markdown source -------------------------------------------------
& (Join-Path $PSScriptRoot 'Build-Documentation.ps1') | Out-Null

# ---- Files needed at run time -------------------------------------------------------------------
$files = [Collections.Generic.List[string]]::new()
foreach ($f in 'Invoke-MeetingCleanup.ps1', 'MeetingCleanup.psd1', 'MeetingCleanup.psm1', 'README.md', 'CHANGELOG.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md',
    'config\MeetingCleanup.config.psd1', 'templates\Report.template.html', 'docs\MeetingCleanup-Guide.html') { $files.Add($f) }
Get-ChildItem -LiteralPath (Join-Path $root 'src') -Filter '*.ps1' -File | ForEach-Object { $files.Add("src\$($_.Name)") }
foreach ($f in $files) {
    $source = Join-Path $root $f
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Missing file in the tool folder: $f" }
    $target = Join-Path $Destination $f
    [void][IO.Directory]::CreateDirectory((Split-Path $target -Parent))
    Copy-Item -LiteralPath $source -Destination $target
}

# ---- Checks ---------------------------------------------------------------------------------------
$problems = [Collections.Generic.List[string]]::new()
foreach ($name in 'reports', 'logs', 'tests', 'artifacts', 'tools', 'docs\images') {
    if (Test-Path -LiteralPath (Join-Path $Destination $name)) { $problems.Add("Folder $name\ must not be in the package.") }
}
Get-ChildItem -LiteralPath $Destination -Recurse -File -Include '*.log', '*.csv', '*.json', '*.png', '*.Tests.ps1', '*FakeGraph*' |
    ForEach-Object { $problems.Add("Not a run-time file: $($_.Name)") }
foreach ($part in 'Console', 'Config', 'Graph', 'Search', 'Cleanup', 'Restore', 'Report', 'Gui') {
    if (-not (Test-Path -LiteralPath (Join-Path $Destination "src\MeetingCleanup.$part.ps1"))) { $problems.Add("Missing in the package: src\MeetingCleanup.$part.ps1") }
}
$config = Import-PowerShellDataFile -LiteralPath (Join-Path $Destination 'config\MeetingCleanup.config.psd1')
if ($config.Tenant.TenantId -or $config.Authentication.AppId -or $config.Authentication.CertificateThumbprint -or $config.Restore.UserPrincipalName) { $problems.Add('The configuration of the package must not hold a tenant, an application, a certificate or an administrator.') }
if ($problems.Count) { throw ("Package not valid ($Destination):`n - " + ($problems -join "`n - ")) }

# The module loads from the package, and its reports go under the package folder.
$loaded = & pwsh -NoProfile -Command "Import-Module '$Destination\MeetingCleanup.psd1'; (Get-Command -Module MeetingCleanup).Count; (Import-MclConfiguration).OutputPath"
if ($LASTEXITCODE -ne 0 -or [int]$loaded[0] -lt 20 -or -not ([string]$loaded[1]).StartsWith($Destination)) { throw "The module does not load correctly from the package: $loaded" }

$all = Get-ChildItem -LiteralPath $Destination -Recurse -File
Write-Host ''
Write-Host "  Meeting Cleanup $version - package ready" -ForegroundColor Green
Write-Host "  Folder   : $Destination"
Write-Host ("  Content  : {0} files, {1:N1} MB" -f $all.Count, (($all | Measure-Object Length -Sum).Sum / 1MB))
Write-Host "  Check    : module loads ($($loaded[0]) commands), reports written under the package folder"
Write-Host "  Config   : empty tenant and application - fill them in before the first run (guide, chapters 5 and 7)"
Write-Host ''
$all | Sort-Object FullName | ForEach-Object { '    {0,10:N0}  {1}' -f $_.Length, $_.FullName.Substring($Destination.Length + 1) }
