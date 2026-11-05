<#
.SYNOPSIS
    Builds the Meeting Cleanup release package in three phases, so that the owner signs between them.

.DESCRIPTION
    The package contains only what Invoke-MeetingCleanup.ps1 needs at run time, plus the HTML guides:
        Invoke-MeetingCleanup.ps1, MeetingCleanup.psd1, MeetingCleanup.psm1, src\, config\, templates\,
        docs\MeetingCleanup-UserGuide.html, docs\MeetingCleanup-Guide.html, README.md, CHANGELOG.md, LICENSE,
        THIRD-PARTY-NOTICES.md, PackageIntegrity.psd1
    The HTML guides are built from their Markdown source with tools\Build-Documentation.ps1 and committed in
    package\docs; they are self-contained (images inline), so the Markdown sources and the images are not copied.
    The README of the package is package\README.md, copied as is.
    It never copies reports\, logs\, artifacts\, tests\ or tools\.

    Build (default): copies the files into a new folder and refuses to continue unless the PowerShell files of
    package\ carry the timestamped owner signature. No manifest is written.
    -Unsigned (continuous integration): the same build without the signature requirement. The committed
    package\PackageIntegrity.psd1 is copied when it exists. An '-unsigned-ci' ZIP is written; it is never a release.
    -Manifest <package folder>: hashes the built folder and writes an unsigned PackageIntegrity.psd1. The owner signs it.
    -Seal <package folder>: verifies the signed manifest on the built folder and on package\ (the bytes of the
    protected files must be the ones in the repository), then writes the release ZIP, with the single root folder
    MeetingCleanup-<version>, and prints its SHA-256.
    Signing key material is never read by this script.

.PARAMETER Destination
    Package folder to build. Default: package\MeetingCleanup-<version>, next to the repository folder.

.PARAMETER Force
    Replace the destination folder if it already contains a package. A folder that contains reports\ or
    logs\ (a package that has been run) is never replaced.

.PARAMETER Unsigned
    Continuous integration: builds without requiring the owner signature.

.PARAMETER Manifest
    Built package folder: writes its unsigned PackageIntegrity.psd1.

.PARAMETER Seal
    Built package folder: verifies it against the signed manifest and writes the release ZIP.

.EXAMPLE
    .\tools\New-MeetingCleanupPackage.ps1 -Destination C:\work\package-out\MeetingCleanup -Unsigned
    Builds an unsigned package for the tests and writes its ZIP.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.3.1
#>
#Requires -Version 7.4
[CmdletBinding(DefaultParameterSetName = 'Build')]
param(
    [Parameter(ParameterSetName = 'Build')][string]$Destination,
    [Parameter(ParameterSetName = 'Build')][switch]$Force,
    [Parameter(ParameterSetName = 'Build')][switch]$Unsigned,
    [Parameter(ParameterSetName = 'Manifest', Mandatory)][string]$Manifest,
    [Parameter(ParameterSetName = 'Seal', Mandatory)][string]$Seal
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot -Parent
$root = Join-Path $repoRoot 'package'
$version = (Import-PowerShellDataFile -LiteralPath (Join-Path $root 'MeetingCleanup.psd1')).ModuleVersion
$signer = '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6'
$signedScripts = @('Invoke-MeetingCleanup.ps1', 'MeetingCleanup.psm1', 'MeetingCleanup.psd1', 'src\Test-PackageIntegrity.ps1') +
    @('Cleanup', 'Config', 'Console', 'Graph', 'Gui', 'Report', 'Restore', 'Search', 'Transfer' | ForEach-Object { "src\MeetingCleanup.$_.ps1" })

function Get-RelativeFile([string]$Base) {
    Get-ChildItem -LiteralPath $Base -Recurse -File -Force | ForEach-Object { $_.FullName.Substring($Base.Length + 1) } | Sort-Object { $_ } -CaseSensitive
}
function Assert-OwnerSigned([string]$Path) {
    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    if ([string]$signature.Status -cne 'Valid' -or $null -eq $signature.SignerCertificate -or
        $signature.SignerCertificate.Thumbprint -cne $signer -or $null -eq $signature.TimeStamperCertificate) {
        throw "Not signed with the trusted, timestamped owner certificate ($($signature.Status)): $Path"
    }
}
function Write-Manifest([string]$Package) {
    $entries = foreach ($relative in Get-RelativeFile $Package) {
        if ($relative -ceq 'PackageIntegrity.psd1' -or $relative -ceq 'config\MeetingCleanup.config.psd1') { continue }
        [pscustomobject]@{ Path = $relative; Sha256 = (Get-FileHash -LiteralPath (Join-Path $Package $relative) -Algorithm SHA256).Hash }
    }
    $packageVersion = (Import-PowerShellDataFile -LiteralPath (Join-Path $Package 'MeetingCleanup.psd1')).ModuleVersion
    $text = [Text.StringBuilder]::new()
    [void]$text.AppendLine('@{')
    [void]$text.AppendLine('    SchemaVersion = 1')
    [void]$text.AppendLine("    Algorithm = 'SHA256'")
    [void]$text.AppendLine("    SignerThumbprint = '$signer'")
    [void]$text.AppendLine("    Version = '$packageVersion'")
    [void]$text.AppendLine('    Files = @(')
    foreach ($entry in $entries) {
        if ($entry.Path.Contains("'")) { throw "A packaged path contains a quote: $($entry.Path)" }
        [void]$text.AppendLine("        @{ Path = '$($entry.Path)'; Sha256 = '$($entry.Sha256)' }")
    }
    [void]$text.AppendLine('    )')
    [void]$text.AppendLine('}')
    $target = Join-Path $Package 'PackageIntegrity.psd1'
    [IO.File]::WriteAllText($target, ($text.ToString() -replace "`r?`n", "`r`n"), [Text.UTF8Encoding]::new($false))
    $target
}
function Write-PackageZip([string]$Package, [string]$ZipPath) {
    # One root folder, MeetingCleanup-<version>, whatever the name of the built folder.
    if (Test-Path -LiteralPath $ZipPath) { throw "The ZIP already exists: $ZipPath" }
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::Open($ZipPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($relative in Get-RelativeFile $Package) {
            [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, (Join-Path $Package $relative),
                ("MeetingCleanup-$version/" + $relative.Replace('\', '/')), [IO.Compression.CompressionLevel]::Optimal)
        }
    }
    finally { $archive.Dispose() }
}

if ($PSCmdlet.ParameterSetName -eq 'Manifest') {
    $Manifest = [IO.Path]::GetFullPath($Manifest, (Get-Location).Path).TrimEnd('\')
    foreach ($relative in $signedScripts) { Assert-OwnerSigned (Join-Path $Manifest $relative) }
    $path = Write-Manifest $Manifest
    Write-Output "Unsigned manifest written: $path"
    Write-Output "Sign it, then copy the signed file to $(Join-Path $root 'PackageIntegrity.psd1') and run: .\tools\New-MeetingCleanupPackage.ps1 -Seal `"$Manifest`""
    return
}

if ($PSCmdlet.ParameterSetName -eq 'Seal') {
    $Seal = [IO.Path]::GetFullPath($Seal, (Get-Location).Path).TrimEnd('\')
    foreach ($used in 'data', 'reports', 'logs', 'bin') {
        if (Test-Path -LiteralPath (Join-Path $Seal $used)) { throw "The built folder contains a $used folder (a package that has been run), it is not sealed: $Seal" }
    }
    $signedManifest = Join-Path $root 'PackageIntegrity.psd1'
    if (-not (Test-Path -LiteralPath $signedManifest)) { throw "Copy the signed manifest to $signedManifest first." }
    if ((Get-FileHash -LiteralPath $signedManifest -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath (Join-Path $Seal 'PackageIntegrity.psd1') -Algorithm SHA256).Hash) {
        Copy-Item -LiteralPath $signedManifest -Destination (Join-Path $Seal 'PackageIntegrity.psd1') -Force
    }
    $result = & (Join-Path $Seal 'src\Test-PackageIntegrity.ps1') -Root $Seal -Distribution
    if ($result.Status -cne 'Valid') { throw 'The built package is not Valid.' }
    $sourceResult = & (Join-Path $root 'src\Test-PackageIntegrity.ps1') -Root $root
    if ($sourceResult.Status -notlike 'Valid*') { throw 'The repository package folder is not Valid.' }
    foreach ($relative in Get-RelativeFile $Seal) {
        $sourceFile = Join-Path $root $relative
        if ($relative -ceq 'CHANGELOG.md') { $sourceFile = Join-Path $repoRoot $relative }
        if ($relative -ceq 'config\MeetingCleanup.config.psd1') { continue }
        if (-not (Test-Path -LiteralPath $sourceFile)) { throw "The packaged file has no source in the repository: $relative" }
        if ((Get-FileHash -LiteralPath $sourceFile -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath (Join-Path $Seal $relative) -Algorithm SHA256).Hash) {
            throw "The packaged file differs from the repository file: $relative"
        }
    }
    $zip = "$Seal.zip"
    Write-PackageZip $Seal $zip
    [pscustomobject]@{
        Package = $Seal; Zip = $zip; ZipBytes = (Get-Item -LiteralPath $zip).Length
        ZipSha256 = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash
        Files = $result.Files + 1; Signed = $true; Published = $false
    } | ConvertTo-Json
    return
}

# ---- Build -----------------------------------------------------------------------------------------
if (-not $Destination) { $Destination = Join-Path (Split-Path $repoRoot -Parent) "package\MeetingCleanup-$version" }
$Destination = [IO.Path]::GetFullPath($Destination, (Get-Location).Path).TrimEnd('\')
$rootPrefix = [IO.Path]::GetFullPath($repoRoot).TrimEnd('\') + '\'
if (($Destination + '\').StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase) -or $rootPrefix.StartsWith($Destination + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "The destination must be outside the repository folder: $Destination"
}
if (Test-Path -LiteralPath $Destination) {
    if (-not $Force) { throw "The destination already exists: $Destination. Use -Force to replace it." }
    if (-not (Test-Path -LiteralPath (Join-Path $Destination 'Invoke-MeetingCleanup.ps1'))) { throw "The destination is not a Meeting Cleanup package, it is not replaced: $Destination" }
    foreach ($used in 'reports', 'logs') {
        if (Test-Path -LiteralPath (Join-Path $Destination $used)) { throw "The destination contains a $used folder (a package that has been run), it is not replaced: $Destination" }
    }
    Remove-Item -LiteralPath $Destination -Recurse -Force
}
$zip = if ($Unsigned) { "$Destination-unsigned-ci.zip" } else { $null }
if ($zip -and (Test-Path -LiteralPath $zip)) { throw "The ZIP already exists: $zip" }
if (-not $Unsigned) { foreach ($relative in $signedScripts) { Assert-OwnerSigned (Join-Path $root $relative) } }

# ---- Files needed at run time (the HTML guides are built beforehand: tools\Build-Documentation.ps1) ----
$files = [Collections.Generic.List[string]]::new()
foreach ($f in 'Invoke-MeetingCleanup.ps1', 'MeetingCleanup.psd1', 'MeetingCleanup.psm1', 'README.md', 'CHANGELOG.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md',
    'config\MeetingCleanup.config.psd1', 'templates\Report.template.html', 'docs\MeetingCleanup-UserGuide.html', 'docs\MeetingCleanup-Guide.html') { $files.Add($f) }
Get-ChildItem -LiteralPath (Join-Path $root 'src') -File | Where-Object Extension -in '.ps1', '.cs' | ForEach-Object { $files.Add("src\$($_.Name)") }
foreach ($f in $files) {
    $sourceRoot = if ($f -eq 'CHANGELOG.md') { $repoRoot } else { $root }
    $source = Join-Path $sourceRoot $f
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Missing file in the tool folder: $f" }
    $target = Join-Path $Destination $f
    [void][IO.Directory]::CreateDirectory((Split-Path $target -Parent))
    Copy-Item -LiteralPath $source -Destination $target
}

# ---- Checks ----------------------------------------------------------------------------------------
$problems = [Collections.Generic.List[string]]::new()
foreach ($name in 'reports', 'logs', 'tests', 'artifacts', 'tools', 'docs\images') {
    if (Test-Path -LiteralPath (Join-Path $Destination $name)) { $problems.Add("Folder $name\ must not be in the package.") }
}
Get-ChildItem -LiteralPath $Destination -Recurse -File -Include '*.log', '*.csv', '*.json', '*.png', '*.Tests.ps1', '*FakeGraph*', '*-UserGuide.md', '*-Guide.md' |
    ForEach-Object { $problems.Add("Not a run-time file: $($_.Name)") }
foreach ($part in 'Console.ps1', 'Config.ps1', 'Graph.ps1', 'Search.ps1', 'Cleanup.ps1', 'Restore.ps1', 'Transfer.ps1', 'Report.ps1', 'Gui.ps1', 'Native.cs') {
    if (-not (Test-Path -LiteralPath (Join-Path $Destination "src\MeetingCleanup.$part"))) { $problems.Add("Missing in the package: src\MeetingCleanup.$part") }
}
if (-not (Test-Path -LiteralPath (Join-Path $Destination 'src\Test-PackageIntegrity.ps1'))) { $problems.Add('Missing in the package: src\Test-PackageIntegrity.ps1') }
$config = Import-PowerShellDataFile -LiteralPath (Join-Path $Destination 'config\MeetingCleanup.config.psd1')
if ($config.Tenant.TenantId -or $config.Authentication.AppId -or $config.Authentication.CertificateThumbprint -or $config.Restore.UserPrincipalName) { $problems.Add('The configuration of the package must not hold a tenant, an application, a certificate or an administrator.') }

# Every local link of the packaged README resolves inside the package.
$readme = [IO.File]::ReadAllText((Join-Path $Destination 'README.md'))
foreach ($match in [regex]::Matches($readme, '\]\((?!https?:|#|mailto:)([^)#\s]+)')) {
    $linked = $match.Groups[1].Value
    if (-not (Test-Path -LiteralPath (Join-Path $Destination $linked))) { $problems.Add("README.md links a file that is not in the package: $linked") }
}
if ($problems.Count) { throw ("Package not valid ($Destination):`n - " + ($problems -join "`n - ")) }

if ($Unsigned) {
    $committed = Join-Path $root 'PackageIntegrity.psd1'
    if (Test-Path -LiteralPath $committed) {
        Copy-Item -LiteralPath $committed -Destination (Join-Path $Destination 'PackageIntegrity.psd1')
        & (Join-Path $Destination 'src\Test-PackageIntegrity.ps1') -Root $Destination -DevelopmentMode | Out-Null
    }
    else { Write-Warning 'package\PackageIntegrity.psd1 is not committed: this unsigned build carries no manifest.' }
}

# The module loads from the package, and its reports go under the package folder.
$loaded = & pwsh -NoProfile -Command "Import-Module '$Destination\MeetingCleanup.psd1' 3>`$null; (Get-Command -Module MeetingCleanup).Count; (Import-MclConfiguration).OutputPath"
if ($LASTEXITCODE -ne 0 -or [int]$loaded[0] -lt 20 -or -not ([string]$loaded[1]).StartsWith($Destination)) { throw "The module does not load correctly from the package: $loaded" }

$all = Get-ChildItem -LiteralPath $Destination -Recurse -File
Write-Host ''
Write-Host "  Meeting Cleanup $version - package ready$(if ($Unsigned) { ' (unsigned: continuous integration only)' } else { ' (no manifest yet)' })" -ForegroundColor Green
Write-Host "  Folder   : $Destination"
Write-Host ("  Content  : {0} files, {1:N1} MB" -f $all.Count, (($all | Measure-Object Length -Sum).Sum / 1MB))
Write-Host "  Check    : module loads ($($loaded[0]) commands), reports written under the package folder, README links resolve"
Write-Host ''
$all | Sort-Object FullName | ForEach-Object { '    {0,10:N0}  {1}' -f $_.Length, $_.FullName.Substring($Destination.Length + 1) }
if ($Unsigned) {
    Write-PackageZip $Destination $zip
    Write-Host ''
    Write-Host "  ZIP      : $zip" -ForegroundColor Green
}
else {
    Write-Host ''
    Write-Host "  Next     : .\tools\New-MeetingCleanupPackage.ps1 -Manifest `"$Destination`"" -ForegroundColor Yellow
}
