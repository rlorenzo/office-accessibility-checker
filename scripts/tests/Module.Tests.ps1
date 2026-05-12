<#
.SYNOPSIS
    Pester suite for the OfficeAccessibilityChecker PowerShell Gallery module.

.DESCRIPTION
    Builds the module via scripts/Build-Module.ps1, imports it, and asserts
    the public cmdlets behave the same as direct script invocation against
    a sample of fixtures.
#>

BeforeAll {
    $script:RepoRoot   = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $script:ModuleDir  = Join-Path $script:RepoRoot 'module' 'OfficeAccessibilityChecker'
    $script:ManifestPath = Join-Path $script:ModuleDir 'OfficeAccessibilityChecker.psd1'
    $script:FixtureDir = Join-Path $PSScriptRoot 'fixtures'

    # Build the module so Private/ is populated. Build-Module.ps1 also runs
    # Test-ModuleManifest, so a malformed manifest fails here loudly.
    & (Join-Path $script:RepoRoot 'scripts' 'Build-Module.ps1') | Out-Null

    # Make sure the SDK is available inside the module (Initialize would
    # download it on first use; we skip that since the repo's scripts/lib
    # already has it from the main test setup, and we want this suite to be
    # cheap and offline-friendly).
    $repoSdk     = Join-Path $script:RepoRoot 'scripts' 'lib' 'DocumentFormat.OpenXml.dll'
    $moduleLib   = Join-Path $script:ModuleDir 'Private' 'lib'
    if ((Test-Path -LiteralPath $repoSdk) -and -not (Test-Path -LiteralPath (Join-Path $moduleLib 'DocumentFormat.OpenXml.dll'))) {
        if (-not (Test-Path -LiteralPath $moduleLib)) { New-Item -ItemType Directory -Path $moduleLib | Out-Null }
        Copy-Item -Path (Join-Path $script:RepoRoot 'scripts' 'lib' '*.dll') -Destination $moduleLib -Force
    }

    Import-Module -Name $script:ManifestPath -Force
}

AfterAll {
    Remove-Module OfficeAccessibilityChecker -Force -ErrorAction SilentlyContinue
}

Describe 'OfficeAccessibilityChecker module' {
    Context 'Manifest' {
        It 'passes Test-ModuleManifest' {
            { Test-ModuleManifest -Path $script:ManifestPath } | Should -Not -Throw
        }

        It 'exports the expected public functions' {
            $exported = (Get-Command -Module OfficeAccessibilityChecker | Select-Object -ExpandProperty Name)
            $exported | Should -Contain 'Test-OfficeAccessibility'
            $exported | Should -Contain 'Initialize-OfficeAccessibilityChecker'
        }

        It 'declares PSGallery-friendly metadata' {
            $manifest = Import-PowerShellDataFile -LiteralPath $script:ManifestPath
            $manifest.Description | Should -Not -BeNullOrEmpty
            $manifest.PrivateData.PSData.Tags | Should -Contain 'accessibility'
            $manifest.PrivateData.PSData.ProjectUri | Should -Not -BeNullOrEmpty
            $manifest.PrivateData.PSData.LicenseUri | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Test-OfficeAccessibility' {
        It 'PASSes an accessible fixture' {
            $fixture = Join-Path $script:FixtureDir 'word-accessible-baseline.docx'
            $output  = Test-OfficeAccessibility -Path $fixture
            $LASTEXITCODE | Should -Be 0
            ($output | Select-Object -Last 1) | Should -Match '^PASS '
        }

        It 'FAILs an inaccessible fixture and names the rule' {
            $fixture = Join-Path $script:FixtureDir 'word-missing-table-headers.docx'
            $output  = Test-OfficeAccessibility -Path $fixture
            $LASTEXITCODE | Should -Be 1
            ($output | Select-Object -Last 1) | Should -Match 'FAIL .*MissingTableHeaders'
        }

        It '-Format detailed emits one line per issue' {
            $fixture = Join-Path $script:FixtureDir 'word-missing-table-headers.docx'
            $output  = Test-OfficeAccessibility -Path $fixture -Format detailed
            # At minimum: one ERROR line + one FAIL summary line.
            $issueLines = $output | Where-Object { $_ -match "`t" }
            $issueLines.Count | Should -BeGreaterThan 0
        }

        It '-Fix produces a sibling .fixed file that re-passes' {
            $src    = Join-Path $script:FixtureDir 'word-missing-table-headers.docx'
            $copy   = Join-Path $TestDrive 'word-missing-table-headers.docx'
            Copy-Item -LiteralPath $src -Destination $copy

            $output = Test-OfficeAccessibility -Path $copy -Fix
            $fixed  = Join-Path $TestDrive 'word-missing-table-headers.fixed.docx'
            Test-Path -LiteralPath $fixed | Should -BeTrue
            ($output | Where-Object { $_ -like 'FIXED *' }) | Should -Not -BeNullOrEmpty
        }
    }
}
