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
}
