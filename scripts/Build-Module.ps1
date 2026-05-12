<#
.SYNOPSIS
    Assemble the OfficeAccessibilityChecker module folder for Publish-Module.

.DESCRIPTION
    Copies the canonical scripts under scripts/ into
    module/OfficeAccessibilityChecker/Private/ so the module folder becomes a
    self-contained, publish-ready directory. The repo's scripts/ remain the
    single source of truth — this script only mirrors them into the module.

    After running, the module folder can be:
      - Imported directly:  Import-Module ./module/OfficeAccessibilityChecker
      - Validated:          Test-ModuleManifest ./module/OfficeAccessibilityChecker/OfficeAccessibilityChecker.psd1
      - Published:          Publish-Module -Path ./module/OfficeAccessibilityChecker -NuGetApiKey <key>

.PARAMETER Clean
    Remove the Private/ folder before copying. Use when removing scripts
    between releases.
#>

[CmdletBinding()]
param(
    [switch] $Clean
)

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

$repoRoot   = Split-Path -Parent $PSScriptRoot
$scriptsDir = Join-Path $repoRoot 'scripts'
$moduleDir  = Join-Path (Join-Path $repoRoot 'module') 'OfficeAccessibilityChecker'
$privateDir = Join-Path $moduleDir 'Private'

if (-not (Test-Path -LiteralPath $moduleDir)) {
    throw "Expected module directory at '$moduleDir'. Has the module been removed?"
}

if ($Clean -and (Test-Path -LiteralPath $privateDir)) {
    Remove-Item -LiteralPath $privateDir -Recurse -Force
}
if (-not (Test-Path -LiteralPath $privateDir)) {
    New-Item -ItemType Directory -Path $privateDir | Out-Null
}

# Copy the entry-point dispatcher, the per-format checkers, and the SDK
# downloader. Anything under scripts/tests/ is excluded — the module only
# needs runtime code.
$bundledScripts = @(
    'check-office-accessibility.ps1'
    'check-docx-accessibility.ps1'
    'check-xlsx-accessibility.ps1'
    'check-pptx-accessibility.ps1'
    'setup-accessibility-checker.ps1'
)
foreach ($name in $bundledScripts) {
    $src = Join-Path $scriptsDir $name
    if (-not (Test-Path -LiteralPath $src)) {
        throw "Source script '$src' not found. The repo layout has changed; update Build-Module.ps1."
    }
    Copy-Item -LiteralPath $src -Destination (Join-Path $privateDir $name) -Force
}

# Validate the manifest now so a broken release isn't discovered at
# Publish-Module time.
$manifestPath = Join-Path $moduleDir 'OfficeAccessibilityChecker.psd1'
$null = Test-ModuleManifest -Path $manifestPath

Write-Information "Module assembled at: $moduleDir"
Write-Information "Bundled scripts:"
Get-ChildItem -LiteralPath $privateDir -Filter '*.ps1' | ForEach-Object { Write-Information "  $($_.Name)" }
