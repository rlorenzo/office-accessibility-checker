<#
.SYNOPSIS
    Pester suite for the Office accessibility checkers.

.DESCRIPTION
    Drives `check-office-accessibility.ps1` against every fixture listed in
    fixtures/manifest.psd1 and asserts the exit code + rule output match the
    manifest. Adding a new rule is "drop a fixture in the folder + add a
    manifest entry" -- no test code changes required.
#>

BeforeDiscovery {
    $manifestPath = Join-Path $PSScriptRoot 'fixtures' 'manifest.psd1'
    $script:Fixtures = (Import-PowerShellDataFile -LiteralPath $manifestPath).Fixtures
}

BeforeAll {
    $script:CheckerPath = Join-Path $PSScriptRoot '..' 'check-office-accessibility.ps1'
    $script:FixtureDir  = Join-Path $PSScriptRoot 'fixtures'

    # Make sure the SDK is present locally so the checker can load it. CI runs
    # setup-accessibility-checker.ps1 in a separate step, but a developer who
    # invokes the suite directly may not have done so yet.
    $sdkPath = Join-Path $PSScriptRoot '..' 'lib' 'DocumentFormat.OpenXml.dll'
    if (-not (Test-Path -LiteralPath $sdkPath)) {
        & (Join-Path $PSScriptRoot '..' 'setup-accessibility-checker.ps1')
    }

    function Invoke-Checker {
        param(
            [Parameter(Mandatory)] [string] $Path
        )
        # Capture stdout. The checker emits one tab-separated line per issue
        # in detailed mode, then a final PASS/FAIL line. We don't care about
        # stderr here -- it's used for tool errors, which are exit-code asserted.
        $stdout  = & $script:CheckerPath -FilePath $Path -Format detailed 2>$null
        $exit    = $LASTEXITCODE
        $lines   = @($stdout | ForEach-Object { [string]$_ })
        # Issue lines look like "SEVERITY`tRuleName`tDescription".
        $rules   = $lines |
            Where-Object { $_ -match "`t" } |
            ForEach-Object { ($_ -split "`t")[1] }
        [pscustomobject]@{
            ExitCode  = $exit
            Lines     = $lines
            RuleNames = @($rules)
        }
    }

    function Invoke-Fix {
        param(
            [Parameter(Mandatory)] [string] $Path,
            [string] $Format = 'detailed'
        )
        $stdout = & $script:CheckerPath -FilePath $Path -Format $Format -Fix 2>$null
        $exit   = $LASTEXITCODE
        $lines  = @($stdout | ForEach-Object { [string]$_ })
        # Resolve the FIXED line — its target is the absolute path the checker
        # wrote. Tests assert against that file rather than a derived guess.
        $fixedLine = $lines | Where-Object { $_ -like 'FIXED *' } | Select-Object -First 1
        $fixedPath = $null
        if ($fixedLine) {
            # "FIXED <abs path>: rule (n), rule (m)"
            if ($fixedLine -match '^FIXED\s+(.+?):\s+') { $fixedPath = $Matches[1] }
        }
        [pscustomobject]@{
            ExitCode  = $exit
            Lines     = $lines
            FixedLine = $fixedLine
            FixedPath = $fixedPath
        }
    }

    function Invoke-Bulk {
        param(
            [Parameter(Mandatory)] [string] $Path,
            [switch] $Recurse,
            [string] $Format = 'text'
        )
        $invokeArgs = @{ Path = $Path; Format = $Format }
        if ($Recurse) { $invokeArgs.Recurse = $true }
        $stdout = & $script:CheckerPath @invokeArgs 2>$null
        $exit   = $LASTEXITCODE
        $lines  = @($stdout | ForEach-Object { [string]$_ })
        [pscustomobject]@{
            ExitCode = $exit
            Lines    = $lines
            Pass     = @($lines | Where-Object { $_ -like 'PASS *' }).Count
            Fail     = @($lines | Where-Object { $_ -like 'FAIL *' }).Count
            Error    = @($lines | Where-Object { $_ -like 'ERROR *' }).Count
        }
    }
}

Describe 'Office accessibility checker' {
    Context 'Pre-flight error handling' {
        It 'exits 2 when the file does not exist' {
            $missing = Join-Path $TestDrive 'does-not-exist.docx'
            $result  = Invoke-Checker -Path $missing
            $result.ExitCode | Should -Be 2
        }

        It 'exits 2 for unsupported extensions' {
            $txt = Join-Path $TestDrive 'sample.txt'
            'not an office file' | Set-Content -LiteralPath $txt
            $result = Invoke-Checker -Path $txt
            $result.ExitCode | Should -Be 2
        }

        It 'exits 2 for legacy .doc files' {
            $doc = Join-Path $TestDrive 'legacy.doc'
            'placeholder' | Set-Content -LiteralPath $doc
            $result = Invoke-Checker -Path $doc
            $result.ExitCode | Should -Be 2
        }
    }

    Context 'Fixture: <File>' -ForEach $script:Fixtures {
        It 'matches manifest expectations' {
            $fixturePath = Join-Path $script:FixtureDir $File
            $fixturePath | Should -Exist -Because 'the manifest references it; rebuild via Build-Fixtures.ps1 if missing'

            $result = Invoke-Checker -Path $fixturePath

            $result.ExitCode | Should -Be $ExpectedExit -Because (
                "checker output was:`n" + ($result.Lines -join "`n")
            )

            foreach ($expected in $MustContain) {
                $result.RuleNames | Should -Contain $expected -Because (
                    "fixture is meant to trigger $expected. Got rules: " +
                    ($result.RuleNames -join ', ')
                )
            }

            foreach ($forbidden in $MustNotContain) {
                $result.RuleNames | Should -Not -Contain $forbidden -Because (
                    "fixture should not trigger $forbidden. Got rules: " +
                    ($result.RuleNames -join ', ')
                )
            }
        }
    }

    Context 'Bulk scan' {
        BeforeEach {
            $script:BulkDir = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $script:BulkDir | Out-Null
        }

        It 'all accessible files: exit 0 with one PASS per file' {
            'word-accessible-baseline.docx',
            'excel-accessible-baseline.xlsx' | ForEach-Object {
                Copy-Item -LiteralPath (Join-Path $script:FixtureDir $_) -Destination $script:BulkDir
            }

            $result = Invoke-Bulk -Path $script:BulkDir
            $result.ExitCode | Should -Be 0
            $result.Pass     | Should -Be 2
            $result.Fail     | Should -Be 0
            $result.Error    | Should -Be 0
        }

        It 'mixed pass/fail: exits 1 and reports both' {
            'word-accessible-baseline.docx',
            'excel-accessible-baseline.xlsx',
            'word-missing-alt-text.docx' | ForEach-Object {
                Copy-Item -LiteralPath (Join-Path $script:FixtureDir $_) -Destination $script:BulkDir
            }

            $result = Invoke-Bulk -Path $script:BulkDir
            $result.ExitCode | Should -Be 1
            $result.Pass     | Should -Be 2
            $result.Fail     | Should -Be 1
            $result.Error    | Should -Be 0
        }

        It 'does not descend into subdirectories without -Recurse' {
            $sub = Join-Path $script:BulkDir 'nested'
            New-Item -ItemType Directory -Path $sub | Out-Null
            Copy-Item -LiteralPath (Join-Path $script:FixtureDir 'word-accessible-baseline.docx') -Destination $sub

            $result = Invoke-Bulk -Path $script:BulkDir
            $result.ExitCode | Should -Be 0
            $result.Pass     | Should -Be 0
            $result.Fail     | Should -Be 0
            $result.Error    | Should -Be 0
        }

        It '-Recurse picks up fixtures in subdirectories' {
            $sub = Join-Path $script:BulkDir 'nested'
            New-Item -ItemType Directory -Path $sub | Out-Null
            Copy-Item -LiteralPath (Join-Path $script:FixtureDir 'word-accessible-baseline.docx') -Destination $sub

            $result = Invoke-Bulk -Path $script:BulkDir -Recurse
            $result.ExitCode | Should -Be 0
            $result.Pass     | Should -Be 1
        }

        It 'silently skips unsupported extensions' {
            Copy-Item -LiteralPath (Join-Path $script:FixtureDir 'word-accessible-baseline.docx') -Destination $script:BulkDir
            'irrelevant content' | Set-Content -LiteralPath (Join-Path $script:BulkDir 'notes.txt')
            'irrelevant content' | Set-Content -LiteralPath (Join-Path $script:BulkDir 'report.pdf')

            $result = Invoke-Bulk -Path $script:BulkDir
            $result.ExitCode | Should -Be 0
            $result.Pass     | Should -Be 1
            $result.Error    | Should -Be 0
        }

        It 'collect-and-continue: a corrupt file does not halt the scan' {
            Copy-Item -LiteralPath (Join-Path $script:FixtureDir 'word-accessible-baseline.docx') -Destination $script:BulkDir
            # Zero-byte file with valid extension. The SDK throws FileFormatException
            # on Open(), which the per-format checker sniffs and surfaces as
            # DocumentProtected (exit 1). Exit 2 (true tool errors) is reserved
            # for IOException / SDK-missing paths -- not reliably reproducible
            # from a unit test, so this case validates only the exit-1 path.
            New-Item -ItemType File -Path (Join-Path $script:BulkDir 'corrupt.docx') | Out-Null

            $result = Invoke-Bulk -Path $script:BulkDir
            $result.ExitCode | Should -Be 1
            $result.Pass     | Should -Be 1 -Because 'the good fixture must still be scanned after the corrupt one fails'
            $result.Fail     | Should -Be 1
        }

        It 'empty directory: exit 0' {
            $result = Invoke-Bulk -Path $script:BulkDir
            $result.ExitCode | Should -Be 0
            $result.Pass     | Should -Be 0
            $result.Fail     | Should -Be 0
            $result.Error    | Should -Be 0
        }

        It 'ignores Office lock files (~$*)' {
            Copy-Item -LiteralPath (Join-Path $script:FixtureDir 'word-accessible-baseline.docx') -Destination $script:BulkDir
            # ~$ files are zero-byte by convention; they would otherwise fail noisily.
            New-Item -ItemType File -Path (Join-Path $script:BulkDir '~$lock.docx') | Out-Null

            $result = Invoke-Bulk -Path $script:BulkDir
            $result.ExitCode | Should -Be 0
            $result.Pass     | Should -Be 1
            $result.Error    | Should -Be 0
        }
    }

    Context 'Autofix' {
        BeforeEach {
            # Each test copies the relevant fixture into a per-test scratch dir
            # so Invoke-Fix can write *.fixed.<ext> sibling without polluting
            # the committed fixtures folder.
            $script:FixDir = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $script:FixDir | Out-Null
        }

        # Each row: source fixture, the rule we expect the fixer to clear from
        # the post-fix run, and whether the fix should result in an exit-0
        # (true) or exit-1 (false — there are remaining errors after the fix).
        $cases = @(
            @{ File = 'word-missing-table-headers.docx';       Rule = 'MissingTableHeaders';       Passes = $true  }
            @{ File = 'word-repeated-blanks.docx';             Rule = 'RepeatedBlanks';            Passes = $true  }
            @{ File = 'word-low-contrast.docx';                Rule = 'LowContrast';               Passes = $true  }
            @{ File = 'excel-missing-table-headers.xlsx';      Rule = 'MissingTableHeaders';       Passes = $true  }
            @{ File = 'excel-red-only-negative-formatting.xlsx'; Rule = 'RedOnlyNegativeFormatting'; Passes = $true }
            @{ File = 'excel-low-contrast.xlsx';               Rule = 'LowContrast';               Passes = $true  }
            @{ File = 'powerpoint-missing-table-headers.pptx'; Rule = 'MissingTableHeaders';       Passes = $true  }
            @{ File = 'powerpoint-low-contrast.pptx';          Rule = 'LowContrast';               Passes = $true  }
        )

        It '<File>: -Fix produces a .fixed copy that no longer reports <Rule>' -ForEach $cases {
            $src  = Join-Path $script:FixtureDir $File
            $copy = Join-Path $script:FixDir $File
            Copy-Item -LiteralPath $src -Destination $copy

            $result = Invoke-Fix -Path $copy
            $result.FixedLine | Should -Not -BeNullOrEmpty -Because (
                "expected a FIXED line. Output was:`n" + ($result.Lines -join "`n")
            )
            $result.FixedLine | Should -Match $Rule -Because (
                "FIXED line should mention the rule that was repaired"
            )
            $result.FixedPath | Should -Not -BeNullOrEmpty
            Test-Path -LiteralPath $result.FixedPath | Should -BeTrue

            # Re-check the fixed file directly: the rule must no longer fire.
            $verify = Invoke-Checker -Path $result.FixedPath
            $verify.RuleNames | Should -Not -Contain $Rule -Because (
                "fix was supposed to eliminate $Rule. Post-fix output was:`n" +
                ($verify.Lines -join "`n")
            )

            if ($Passes) {
                $verify.ExitCode | Should -Be 0 -Because (
                    "fixture only fails on $Rule, so the fixed file should pass"
                )
            }
        }

        It 'is a no-op on already-accessible files' {
            $src  = Join-Path $script:FixtureDir 'word-accessible-baseline.docx'
            $copy = Join-Path $script:FixDir 'word-accessible-baseline.docx'
            Copy-Item -LiteralPath $src -Destination $copy

            $result = Invoke-Fix -Path $copy
            $result.ExitCode | Should -Be 0
            $result.FixedLine | Should -BeNullOrEmpty
            # No .fixed copy should be written when nothing was fixed.
            (Test-Path -LiteralPath (Join-Path $script:FixDir 'word-accessible-baseline.fixed.docx')) | Should -BeFalse
        }

        It 'leaves the original file unmodified' {
            $src  = Join-Path $script:FixtureDir 'word-missing-table-headers.docx'
            $copy = Join-Path $script:FixDir 'word-missing-table-headers.docx'
            Copy-Item -LiteralPath $src -Destination $copy
            $beforeHash = (Get-FileHash -LiteralPath $copy).Hash

            $result = Invoke-Fix -Path $copy
            $result.FixedLine | Should -Not -BeNullOrEmpty

            (Get-FileHash -LiteralPath $copy).Hash | Should -Be $beforeHash -Because (
                "the input file must never be mutated by -Fix"
            )
        }

        It 'bulk-scan with -Fix skips previously produced .fixed files' {
            # First run: creates word-missing-table-headers.fixed.docx alongside the input.
            Copy-Item -LiteralPath (Join-Path $script:FixtureDir 'word-missing-table-headers.docx') `
                      -Destination $script:FixDir
            Copy-Item -LiteralPath (Join-Path $script:FixtureDir 'word-accessible-baseline.docx') `
                      -Destination $script:FixDir

            & $script:CheckerPath -Path $script:FixDir -Fix 2>$null | Out-Null
            $LASTEXITCODE | Should -BeIn @(0, 1)

            # Second run on the same dir: should not pick up *.fixed.docx as an
            # input (else we'd re-emit FIXED on it and create *.fixed.fixed.docx).
            & $script:CheckerPath -Path $script:FixDir -Fix 2>$null | Out-Null
            $fixedFixed = Join-Path $script:FixDir 'word-missing-table-headers.fixed.fixed.docx'
            (Test-Path -LiteralPath $fixedFixed) | Should -BeFalse
        }
    }
}
