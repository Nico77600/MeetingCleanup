#Requires -Version 7.4
# Signature-dependent cases carry the RealSignature tag: they need a sealed release package (MCL_RELEASE_PACKAGE) and skip without it.
# The other cases run the real verifier on a fixture copy whose Authenticode check is replaced by a no-op, so that the
# manifest, the inventory, the hashes and the folder rules are exercised before the owner signs.
BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Source = Join-Path $script:RepoRoot 'package'
    $script:Pwsh = Join-Path $PSHOME 'pwsh.exe'
    # TestDrive is disabled by Run-Tests.ps1: the fixtures live in a private folder of the ignored artifacts folder.
    $script:Work = Join-Path $script:RepoRoot ('artifacts\package-tests-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($script:Work)
    $script:Signer = '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6'
    $script:FixtureNumber = 0
    $script:Fixture = $null

    # The manifest writer of the package builder, used as is.
    $builderPath = Join-Path $script:RepoRoot 'tools\New-MeetingCleanupPackage.ps1'
    $ast = [Management.Automation.Language.Parser]::ParseFile($builderPath, [ref]$null, [ref]$null)
    foreach ($name in 'Get-RelativeFile', 'Write-Manifest') {
        $function = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq $name }, $true)
        if ($null -eq $function) { throw "The builder function $name is missing." }
        . ([scriptblock]::Create($function.Extent.Text))
    }
    $signer = $script:Signer

    function New-FixtureTemplate {
        $template = Join-Path $script:Work 'template'
        Copy-Item -LiteralPath $script:Source -Destination $template -Recurse
        foreach ($folder in 'reports', 'logs') { Remove-Item -LiteralPath (Join-Path $template $folder) -Recurse -Force -ErrorAction SilentlyContinue }
        Remove-Item -LiteralPath (Join-Path $template 'PackageIntegrity.psd1') -ErrorAction SilentlyContinue
        # A release package holds the HTML guides only.
        Remove-Item -LiteralPath (Join-Path $template 'docs\images') -Recurse -Force
        Get-ChildItem -LiteralPath (Join-Path $template 'docs') -Filter '*.md' | Remove-Item
        Copy-Item -LiteralPath (Join-Path $script:RepoRoot 'CHANGELOG.md') -Destination $template
        # The test copy of the verifier does not check Authenticode: everything else is the delivered code.
        $verifier = Join-Path $template 'src\Test-PackageIntegrity.ps1'
        $text = [IO.File]::ReadAllText($verifier)
        $stubbed = [regex]::Replace($text, '(?s)function Assert-OwnerSignature\(\[string\]\$Path\) \{.*?\r?\n\}\r?\n', "function Assert-OwnerSignature([string]`$Path) { }`r`n")
        if ($stubbed -ceq $text) { throw 'The owner-signature check was not found in the verifier.' }
        [IO.File]::WriteAllText($verifier, $stubbed, [Text.UTF8Encoding]::new($false))
        [void](Write-Manifest $template)
        $template
    }
    function Reset-Fixture([switch]$Source) {
        Remove-Module MeetingCleanup -Force -ErrorAction SilentlyContinue
        $script:FixtureNumber++
        $script:Fixture = Join-Path $script:Work ('package-{0:D4}' -f $script:FixtureNumber)
        Copy-Item -LiteralPath $script:Template -Destination $script:Fixture -Recurse
        # A source checkout has no file copied at build time.
        if ($Source) { Remove-Item -LiteralPath (Join-Path $script:Fixture 'CHANGELOG.md') }
    }
    function Invoke-Guard([string[]]$Switch = @()) {
        $verifier = Join-Path $script:Fixture 'src\Test-PackageIntegrity.ps1'
        $arguments = @{ Root = $script:Fixture }
        foreach ($name in $Switch) { $arguments[$name] = $true }
        & $verifier @arguments 3> $null
    }
    function Invoke-Child([string]$Command) {
        $output = & $script:Pwsh -NoProfile -NonInteractive -Command $Command 2>&1 | ForEach-Object { "$_" }
        [pscustomobject]@{ ExitCode = $LASTEXITCODE; Text = ($output -join "`n") }
    }
    function Edit-Protected([string]$Path) {
        if ([IO.Path]::GetExtension($Path) -in '.ps1', '.psm1', '.psd1') {
            [IO.File]::WriteAllText($Path, "# Evidence-only tamper`r`n" + [IO.File]::ReadAllText($Path), [Text.UTF8Encoding]::new($true))
        } else { [IO.File]::AppendAllBytes($Path, [Text.Encoding]::ASCII.GetBytes("`r`n# Evidence-only tamper")) }
    }
    $script:Template = New-FixtureTemplate
    $manifest = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Template 'PackageIntegrity.psd1')
    $script:Entries = @($manifest.Files)
}

AfterAll {
    Remove-Item -LiteralPath $script:Work -Recurse -Force -ErrorAction SilentlyContinue
}

Describe 'Signed manifest and verifier' {
    BeforeEach { Reset-Fixture }
    It 'accepts the distribution, the same files as a source checkout, and each in its own mode' {
        (Invoke-Guard @('Distribution')).Status | Should -Be 'Valid'
        (Invoke-Guard @('DevelopmentMode')).Status | Should -BeLike 'Valid*'
        Reset-Fixture -Source
        (Invoke-Guard).Status | Should -Be 'Valid (files only)'
    }
    It 'refuses -Distribution together with -DevelopmentMode' {
        { Invoke-Guard @('Distribution', 'DevelopmentMode') } | Should -Throw -ExpectedMessage '*cannot be combined*'
    }
    It 'inventories the protected files, and not the editable configuration' {
        $paths = @($script:Entries.Path)
        $paths.Count | Should -Be 21
        $paths | Should -Not -Contain 'config\MeetingCleanup.config.psd1'
        $paths | Should -Not -Contain 'PackageIntegrity.psd1'
        foreach ($required in 'Invoke-MeetingCleanup.ps1', 'MeetingCleanup.psm1', 'MeetingCleanup.psd1', 'src\Test-PackageIntegrity.ps1', 'src\MeetingCleanup.Native.cs',
            'templates\Report.template.html', 'docs\MeetingCleanup-UserGuide.html', 'docs\MeetingCleanup-Guide.html', 'README.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md', 'CHANGELOG.md') {
            $paths | Should -Contain $required
        }
        @($paths | Where-Object { $_ -match '\.ps(m1|d1|1)$' }).Count | Should -Be 13
    }
    It 'keeps the verifier only in src' {
        Copy-Item -LiteralPath (Join-Path $script:Fixture 'src\Test-PackageIntegrity.ps1') -Destination (Join-Path $script:Fixture 'Test-PackageIntegrity.ps1')
        { & (Join-Path $script:Fixture 'Test-PackageIntegrity.ps1') -Root $script:Fixture } | Should -Throw -ExpectedMessage '*verifier must belong to this package root*'
    }
    It 'accepts the editable configuration and the generated folders' {
        Add-Content (Join-Path $script:Fixture 'config\MeetingCleanup.config.psd1') '# Fictitious customization'
        foreach ($folder in 'reports', 'logs') {
            [void][IO.Directory]::CreateDirectory((Join-Path $script:Fixture $folder))
            Set-Content (Join-Path $script:Fixture "$folder\generated.txt") 'generated'
        }
        (Invoke-Guard @('Distribution')).Status | Should -Be 'Valid'
    }
    It 'rejects a changed byte in every protected entry' {
        foreach ($entry in $script:Entries) {
            Reset-Fixture
            Edit-Protected (Join-Path $script:Fixture $entry.Path)
            { Invoke-Guard @('Distribution') } | Should -Throw -ExpectedMessage '*SHA256 mismatch*' -Because $entry.Path
        }
    }
    It 'rejects missing mandatory files' {
        foreach ($relative in 'src\MeetingCleanup.Native.cs', 'docs\MeetingCleanup-Guide.html', 'README.md', 'src\MeetingCleanup.Graph.ps1') {
            Reset-Fixture
            Remove-Item -LiteralPath (Join-Path $script:Fixture $relative)
            { Invoke-Guard } | Should -Throw -Because $relative
        }
        Reset-Fixture
        Remove-Item -LiteralPath (Join-Path $script:Fixture 'CHANGELOG.md')
        { Invoke-Guard @('Distribution') } | Should -Throw
    }
    It 'rejects an unexpected file anywhere in the package' {
        foreach ($relative in 'extra.ps1', 'src\extra.cs', 'docs\extra.html', 'templates\extra.html', 'config\extra.psd1', 'docs\images\extra.png') {
            Reset-Fixture
            $path = Join-Path $script:Fixture $relative
            [void][IO.Directory]::CreateDirectory((Split-Path $path -Parent))
            Set-Content -LiteralPath $path -Value 'unexpected'
            { Invoke-Guard } | Should -Throw -ExpectedMessage '*unexpected stable file*' -Because $relative
        }
    }
    It 'rejects a junction in a protected subtree' {
        $outside = Join-Path $script:Work 'junction-target'
        [void][IO.Directory]::CreateDirectory($outside)
        foreach ($relative in 'src', 'templates', 'docs', 'config') {
            Reset-Fixture
            $link = Join-Path $script:Fixture $relative
            Remove-Item -LiteralPath $link -Recurse -Force
            New-Item -ItemType Junction -Path $link -Target $outside | Out-Null
            try { { Invoke-Guard } | Should -Throw -Because $relative }
            finally { (Get-Item -LiteralPath $link -Force).Delete() }
        }
    }
    It 'rejects a junction in place of a generated folder' {
        $outside = Join-Path $script:Work 'junction-generated'
        [void][IO.Directory]::CreateDirectory($outside)
        foreach ($name in 'reports', 'logs') {
            Reset-Fixture
            $link = Join-Path $script:Fixture $name
            New-Item -ItemType Junction -Path $link -Target $outside | Out-Null
            try { { Invoke-Guard } | Should -Throw -ExpectedMessage '*reparse point*' -Because $name }
            finally { (Get-Item -LiteralPath $link -Force).Delete() }
        }
    }
    It 'rejects an absent manifest and a manifest of the wrong shape' {
        Remove-Item -LiteralPath (Join-Path $script:Fixture 'PackageIntegrity.psd1')
        { Invoke-Guard } | Should -Throw -ExpectedMessage '*PackageIntegrity.psd1 is missing*releases/latest*'
        Reset-Fixture
        Set-Content -LiteralPath (Join-Path $script:Fixture 'PackageIntegrity.psd1') -Value '@{ SchemaVersion = 1 }'
        { Invoke-Guard } | Should -Throw -ExpectedMessage '*unsupported or incomplete manifest*'
    }
    It 'rejects a manifest that omits, duplicates, adds or retargets an entry' {
        $path = Join-Path $script:Template 'PackageIntegrity.psd1'
        $lines = [IO.File]::ReadAllLines($path)
        $first = $lines | Where-Object { $_ -match "Path = 'README.md'" } | Select-Object -First 1
        $mutations = @(
            ($lines | Where-Object { $_ -ne $first })
            (@($lines[0..($lines.IndexOf($first))]) + @($first) + @($lines[($lines.IndexOf($first) + 1)..($lines.Count - 1)]))
            ($lines | ForEach-Object { if ($_ -eq $first) { $_; $_.Replace('README.md', 'Other.md') } else { $_ } })
            ($lines | ForEach-Object { $_.Replace("Path = 'README.md'", "Path = 'LICENSE'") })
        )
        foreach ($mutation in $mutations) {
            Reset-Fixture
            [IO.File]::WriteAllText((Join-Path $script:Fixture 'PackageIntegrity.psd1'), ($mutation -join "`r`n") + "`r`n", [Text.UTF8Encoding]::new($false))
            { Invoke-Guard } | Should -Throw
        }
    }
    It 'refuses a manifest without the owner signature when the real verifier runs' {
        # The delivered verifier, with its Authenticode check, on an unsigned fixture.
        $real = Join-Path $script:Fixture 'src\Test-PackageIntegrity.ps1'
        Copy-Item -LiteralPath (Join-Path $script:Source 'src\Test-PackageIntegrity.ps1') -Destination $real -Force
        { & $real -Root $script:Fixture } | Should -Throw -ExpectedMessage '*expected trusted, timestamped owner signature*'
    }
}

Describe 'Integrity at the entry points' {
    BeforeEach { Reset-Fixture }
    It 'refuses to start from a package whose verifier is not owner-signed, before the module is loaded' {
        foreach ($relative in 'src\MeetingCleanup.Native.cs', 'src\Test-PackageIntegrity.ps1', 'MeetingCleanup.psm1') {
            Reset-Fixture
            Edit-Protected (Join-Path $script:Fixture $relative)
            $run = Invoke-Child ("& '$(Join-Path $script:Fixture 'Invoke-MeetingCleanup.ps1')' -Organizer a@contoso.test; 'EXIT=' + `$LASTEXITCODE")
            $run.Text | Should -Match 'Package integrity: verifier must have the trusted, timestamped owner signature' -Because $relative
            $run.Text | Should -Not -Match 'EXIT=0'
            Test-Path -LiteralPath (Join-Path $script:Fixture 'reports') | Should -BeFalse
            Test-Path -LiteralPath (Join-Path $script:Fixture 'logs') | Should -BeFalse
        }
    }
    It 'refuses a tampered verifier before the module is imported' {
        Edit-Protected (Join-Path $script:Fixture 'src\Test-PackageIntegrity.ps1')
        { Import-Module (Join-Path $script:Fixture 'MeetingCleanup.psd1') -Force -ErrorAction Stop } | Should -Throw -ExpectedMessage '*Package integrity:*'
    }
    It 'does not run a verifier that lacks the owner signature, even from an otherwise correct package' {
        # The fixture verifier is unsigned: the module and the entry script refuse it instead of executing it.
        { Import-Module (Join-Path $script:Fixture 'MeetingCleanup.psd1') -Force -ErrorAction Stop } | Should -Throw -ExpectedMessage '*owner signature*'
        $run = Invoke-Child ("& '$(Join-Path $script:Fixture 'Invoke-MeetingCleanup.ps1')' -DevelopmentMode -Organizer a@contoso.test; 'EXIT=' + `$LASTEXITCODE")
        $run.Text | Should -Match 'owner signature'
    }
    It 'accepts an unverified source checkout only with -DevelopmentMode' {
        Reset-Fixture -Source
        Remove-Item -LiteralPath (Join-Path $script:Fixture 'PackageIntegrity.psd1')
        $plain = Invoke-Child ("& '$(Join-Path $script:Fixture 'Invoke-MeetingCleanup.ps1')' -Organizer a@contoso.test; 'EXIT=' + `$LASTEXITCODE")
        $plain.Text | Should -Match 'PackageIntegrity.psd1 is missing\. Download the complete release ZIP .*releases/latest.*-DevelopmentMode'
        $development = Invoke-Child ("& '$(Join-Path $script:Fixture 'Invoke-MeetingCleanup.ps1')' -Organizer a@contoso.test -ConfigPath '$(Join-Path $script:Work 'no-such-config.psd1')' -DevelopmentMode 3>&1 2>&1 | Out-String")
        $development.Text | Should -Match 'DevelopmentMode: no package manifest'
        # The run reaches the loading of the configuration: a configuration error, not a start failure.
        $development.Text | Should -Match 'no-such-config'
        $development.Text | Should -Not -Match 'positional|cannot find a positional'
        $development.Text | Should -Not -Match 'Package integrity'
        # One warning per process: the module does not repeat it at import.
        ([regex]::Matches($development.Text, 'no package manifest|no PackageIntegrity\.psd1')).Count | Should -Be 1
    }
}

Describe 'Owner signatures' -Tag 'RealSignature' {
    It 'requires valid timestamped owner signatures on the PowerShell files and the manifest' {
        $package = $env:MCL_RELEASE_PACKAGE
        if (-not $package -or -not (Test-Path -LiteralPath (Join-Path $package 'PackageIntegrity.psd1'))) {
            Set-ItResult -Skipped -Because 'MCL_RELEASE_PACKAGE names no sealed release package folder (New-MeetingCleanupPackage.ps1 -Seal)'
            return
        }
        $scripts = @(Get-ChildItem -LiteralPath $package -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.psm1' -or ($_.Extension -eq '.psd1' -and $_.Name -ne 'MeetingCleanup.config.psd1') })
        $scripts.Count | Should -Be 14
        foreach ($file in $scripts) {
            $signature = Get-AuthenticodeSignature -LiteralPath $file.FullName
            [string]$signature.Status | Should -Be 'Valid' -Because $file.Name
            $signature.SignerCertificate.Thumbprint | Should -BeExactly $script:Signer
            $signature.TimeStamperCertificate | Should -Not -BeNullOrEmpty
        }
    }
    It 'accepts the sealed package with the real verifier and refuses a copy with one changed byte' {
        $package = $env:MCL_RELEASE_PACKAGE
        if (-not $package -or -not (Test-Path -LiteralPath (Join-Path $package 'PackageIntegrity.psd1'))) {
            Set-ItResult -Skipped -Because 'MCL_RELEASE_PACKAGE names no sealed release package folder (New-MeetingCleanupPackage.ps1 -Seal)'
            return
        }
        (& (Join-Path $package 'src\Test-PackageIntegrity.ps1') -Root $package -Distribution).Status | Should -Be 'Valid'
        $copy = Join-Path $script:Work 'sealed-copy'
        Copy-Item -LiteralPath $package -Destination $copy -Recurse
        [IO.File]::AppendAllBytes((Join-Path $copy 'templates\Report.template.html'), [byte[]]@(32))
        { & (Join-Path $copy 'src\Test-PackageIntegrity.ps1') -Root $copy -Distribution } | Should -Throw -ExpectedMessage '*SHA256 mismatch*'
    }
    It 'accepts the committed package with the real verifier' {
        if (-not (Test-Path -LiteralPath (Join-Path $script:Source 'PackageIntegrity.psd1'))) {
            Set-ItResult -Skipped -Because 'package\PackageIntegrity.psd1 is not committed yet (the owner signs it)'
            return
        }
        (& (Join-Path $script:Source 'src\Test-PackageIntegrity.ps1') -Root $script:Source).Status | Should -Be 'Valid (files only)'
    }
    It 'rejects an altered manifest before parsing it' {
        $package = $env:MCL_RELEASE_PACKAGE
        if (-not $package -or -not (Test-Path -LiteralPath (Join-Path $package 'PackageIntegrity.psd1'))) {
            Set-ItResult -Skipped -Because 'MCL_RELEASE_PACKAGE names no sealed release package folder (New-MeetingCleanupPackage.ps1 -Seal)'
            return
        }
        $copy = Join-Path $script:Work 'altered-manifest'
        Copy-Item -LiteralPath $package -Destination $copy -Recurse
        $path = Join-Path $copy 'PackageIntegrity.psd1'
        [IO.File]::WriteAllText($path, "# Changed`r`n" + [IO.File]::ReadAllText($path), [Text.UTF8Encoding]::new($true))
        { & (Join-Path $copy 'src\Test-PackageIntegrity.ps1') -Root $copy -Distribution } | Should -Throw -ExpectedMessage '*expected trusted, timestamped owner signature*'
    }
    It 'starts from the sealed package up to the configuration, and refuses a copy with one changed byte' {
        $package = $env:MCL_RELEASE_PACKAGE
        if (-not $package -or -not (Test-Path -LiteralPath (Join-Path $package 'PackageIntegrity.psd1'))) {
            Set-ItResult -Skipped -Because 'MCL_RELEASE_PACKAGE names no sealed release package folder (New-MeetingCleanupPackage.ps1 -Seal)'
            return
        }
        $missing = Join-Path $script:Work 'no-such-config.psd1'
        $run = Invoke-Child ("& '$(Join-Path $package 'Invoke-MeetingCleanup.ps1')' -Organizer a@contoso.test -ConfigPath '$missing' 2>&1 | Out-String")
        $run.Text | Should -Match 'no-such-config'
        $run.Text | Should -Not -Match 'Package integrity|positional'
        $copy = Join-Path $script:Work 'sealed-start'
        Copy-Item -LiteralPath $package -Destination $copy -Recurse
        [IO.File]::AppendAllBytes((Join-Path $copy 'templates\Report.template.html'), [byte[]]@(32))
        $changed = Invoke-Child ("& '$(Join-Path $copy 'Invoke-MeetingCleanup.ps1')' -Organizer a@contoso.test -ConfigPath '$missing' 2>&1 | Out-String")
        $changed.Text | Should -Match 'SHA256 mismatch'
    }
    It 'rejects a cryptographically valid manifest with its timestamp removed' {
        $package = $env:MCL_RELEASE_PACKAGE
        if (-not $package -or -not (Test-Path -LiteralPath (Join-Path $package 'PackageIntegrity.psd1'))) {
            Set-ItResult -Skipped -Because 'MCL_RELEASE_PACKAGE names no sealed release package folder (New-MeetingCleanupPackage.ps1 -Seal)'
            return
        }
        Add-Type -AssemblyName System.Security.Cryptography.Pkcs
        $fixture = Join-Path $script:Work 'real-timestamp'
        Copy-Item -LiteralPath $package -Destination $fixture -Recurse
        $path = Join-Path $fixture 'PackageIntegrity.psd1'
        $encoding = [Text.UTF8Encoding]::new($false, $true)
        $text = $encoding.GetString([IO.File]::ReadAllBytes($path))
        $match = [regex]::Match($text, '(?ms)^# SIG # Begin signature block\r?\n(?<Data>.*?)^# SIG # End signature block\r?\n?\z')
        $match.Success | Should -BeTrue
        $cms = [Security.Cryptography.Pkcs.SignedCms]::new()
        $cms.Decode([Convert]::FromBase64String((($match.Groups['Data'].Value -replace '(?m)^#\s?', '') -replace '\s', '')))
        foreach ($attribute in @($cms.SignerInfos[0].UnsignedAttributes)) {
            foreach ($value in $attribute.Values) {
                $cms.SignerInfos[0].RemoveUnsignedAttribute([Security.Cryptography.AsnEncodedData]::new($attribute.Oid, $value.RawData))
            }
        }
        $base64 = [Convert]::ToBase64String($cms.Encode())
        $lines = for ($i = 0; $i -lt $base64.Length; $i += 64) { '# ' + $base64.Substring($i, [Math]::Min(64, $base64.Length - $i)) }
        [IO.File]::WriteAllBytes($path, $encoding.GetBytes($text.Substring(0, $match.Index) +
            "# SIG # Begin signature block`r`n" + ($lines -join "`r`n") + "`r`n# SIG # End signature block`r`n"))
        [string](Get-AuthenticodeSignature -LiteralPath $path).Status | Should -Be 'Valid'
        (Get-AuthenticodeSignature -LiteralPath $path).TimeStamperCertificate | Should -BeNullOrEmpty
        { & (Join-Path $fixture 'src\Test-PackageIntegrity.ps1') -Root $fixture -Distribution } | Should -Throw -ExpectedMessage '*expected trusted, timestamped owner signature*'
    }
}

Describe 'Package builder and source archive' {
    It 'writes a manifest that the verifier parses, in a stable order, with the five keys' {
        $text = [IO.File]::ReadAllText((Join-Path $script:Template 'PackageIntegrity.psd1'))
        $data = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Template 'PackageIntegrity.psd1')
        $data.Keys.Count | Should -Be 5
        $data.SchemaVersion | Should -Be 1
        $data.Algorithm | Should -Be 'SHA256'
        $data.SignerThumbprint | Should -BeExactly $script:Signer
        $data.Version | Should -Be (Import-PowerShellDataFile -LiteralPath (Join-Path $script:Source 'MeetingCleanup.psd1')).ModuleVersion
        $paths = @($data.Files.Path)
        $paths | Should -Be @($paths | Sort-Object { $_ } -CaseSensitive)
        $text | Should -Match "`r`n"
    }
    It 'offers the build, -Manifest and -Seal phases and builds unsigned only on request' {
        $builder = [IO.File]::ReadAllText((Join-Path $script:RepoRoot 'tools\New-MeetingCleanupPackage.ps1'))
        foreach ($name in 'Unsigned', 'Manifest', 'Seal') { $builder | Should -Match "\[Parameter\(ParameterSetName = '\w+'(, Mandatory)?\)\]\[\w+\]\`$$name" }
        $builder | Should -Match "if \(-not \`$Unsigned\) \{ foreach \(\`$relative in \`$signedScripts\) \{ Assert-OwnerSigned"
        $builder | Should -Match 'MeetingCleanup-\$version/'
    }
    It 'builds an unsigned package, and -Seal refuses a built folder that has been run, without writing a ZIP' {
        # The builder refuses a destination inside the repository: use the temporary folder.
        $built = Join-Path ([IO.Path]::GetTempPath()) ('mcl-seal-' + [guid]::NewGuid().ToString('N'))
        $builder = Join-Path $script:RepoRoot 'tools\New-MeetingCleanupPackage.ps1'
        $build = Invoke-Child ("& '$builder' -Destination '$built' -Unsigned *>&1 | Out-String; 'EXIT=' + `$LASTEXITCODE")
        $build.Text | Should -Match 'package ready'
        Test-Path -LiteralPath "$built-unsigned-ci.zip" | Should -BeTrue
        [void][IO.Directory]::CreateDirectory((Join-Path $built 'reports'))
        $seal = Invoke-Child ("& '$builder' -Seal '$built' *>&1 | Out-String")
        $seal.Text | Should -Match 'contains a reports folder'
        Test-Path -LiteralPath "$built.zip" | Should -BeFalse
        Remove-Item -LiteralPath $built, "$built-unsigned-ci.zip" -Recurse -Force -ErrorAction SilentlyContinue
    }
    It 'keeps the packaged README to files the package holds' {
        $readme = [IO.File]::ReadAllText((Join-Path $script:Source 'README.md'))
        $links = @([regex]::Matches($readme, '\]\((?!https?:|#|mailto:)([^)#\s]+)') | ForEach-Object { $_.Groups[1].Value })
        $links.Count | Should -BeGreaterThan 0
        foreach ($link in $links) {
            Test-Path -LiteralPath (Join-Path $script:Template $link) | Should -BeTrue -Because $link
            $link | Should -Not -Match 'docs/.*\.md|images|tests|tools|Run-Tests' -Because 'the packaged README links shipped files only'
        }
        $readme | Should -Not -Match 'Chemin'
    }
    It 'regenerates the HTML guides from their Markdown sources without a difference' {
        foreach ($guide in 'MeetingCleanup-UserGuide', 'MeetingCleanup-Guide') {
            $target = Join-Path $script:Work "$guide.html"
            & (Join-Path $script:RepoRoot 'tools\Build-Documentation.ps1') -Source (Join-Path $script:Source "docs\$guide.md") -Destination $target | Out-Null
            (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath (Join-Path $script:Source "docs\$guide.html") -Algorithm SHA256).Hash -Because "$guide.html is the committed build of its Markdown source"
        }
    }
    It 'keeps every protected file in one line ending (CRLF), without a lone CR' {
        foreach ($entry in $script:Entries) {
            if ($entry.Path -match '\.(png|dll)$') { continue }
            $bytes = [IO.File]::ReadAllBytes((Join-Path $script:Template $entry.Path))
            $text = [Text.Encoding]::UTF8.GetString($bytes)
            ([regex]::Matches($text, "(?<!\r)\n")).Count | Should -Be 0 -Because "$($entry.Path) has a lone LF"
            ([regex]::Matches($text, "\r(?!\n)")).Count | Should -Be 0 -Because "$($entry.Path) has a lone CR"
        }
    }
    It 'marks the signed files byte for byte and leaves the development paths out of the source archive' {
        $attributes = [IO.File]::ReadAllText((Join-Path $script:RepoRoot '.gitattributes'))
        foreach ($rule in 'package/**/*.ps1', 'package/**/*.psm1', 'package/MeetingCleanup.psd1', 'package/PackageIntegrity.psd1') {
            $attributes | Should -Match ([regex]::Escape($rule) + ' -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol')
        }
        foreach ($ignored in '/tests/', '/tools/', '/.github/', '/Run-Tests.ps1', '/AGENTS.md', '/package/docs/*.md', '/package/docs/images/') {
            $attributes | Should -Match ([regex]::Escape($ignored) + ' export-ignore')
        }
        foreach ($kept in 'package/', 'README.md', 'CHANGELOG.md', 'LICENSE', 'SECURITY.md', 'PROVENANCE.md') { $attributes | Should -Not -Match ([regex]::Escape("/$kept") + ' export-ignore') }
    }
}

Describe 'No cached assembly' {
    It 'compiles the verified C# source in memory on each run and loads no cached or prebuilt assembly' {
        $sources = Get-ChildItem -LiteralPath $script:Source -Recurse -File -Include '*.ps1', '*.psm1' | Where-Object { $_.Name -ne 'Test-PackageIntegrity.ps1' }
        foreach ($file in $sources) {
            $text = [IO.File]::ReadAllText($file.FullName)
            $text | Should -Not -Match 'OutputAssembly|\bLoadFrom\b|Assembly\]::(Load|LoadFile)|LOCALAPPDATA|SpecialFolder\]::LocalApplicationData|Add-Type\s+-Path\s+\S*\.dll' -Because $file.Name
        }
        @(Get-ChildItem -LiteralPath $script:Source -Recurse -File -Filter '*.dll').Count | Should -Be 0
        $script = "`$sw = [Diagnostics.Stopwatch]::StartNew(); Import-Module '$(Join-Path $script:Source 'MeetingCleanup.psd1')' 3>`$null; `$sw.Stop(); `$t = [MeetingCleanupNative.Fast] -as [type]; 'LOCATION=[' + `$t.Assembly.Location + ']'; 'DYNAMIC=' + `$t.Assembly.IsDynamic; 'MS=' + [int]`$sw.ElapsedMilliseconds"
        $run = Invoke-Child $script
        $run.Text | Should -Match 'LOCATION=\[\]'
        $run.Text | Should -Match 'MS=\d+'
        Write-Host ("Module import with in-memory compilation: " + ([regex]::Match($run.Text, 'MS=(\d+)').Groups[1].Value) + ' ms')
    }
}
