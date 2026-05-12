function Initialize-OfficeAccessibilityChecker {
<#
.SYNOPSIS
    Download the Open XML SDK into the module's Private/lib/ folder.

.DESCRIPTION
    Thin wrapper over the bundled setup-accessibility-checker.ps1. Run this
    once after Install-Module before invoking Test-OfficeAccessibility.
    Idempotent — re-running is a no-op unless -Force is supplied.

.PARAMETER Force
    Re-download even when the SDK DLL is already present.

.EXAMPLE
    Initialize-OfficeAccessibilityChecker

.EXAMPLE
    Initialize-OfficeAccessibilityChecker -Force
#>
    [CmdletBinding()]
    param(
        [switch] $Force
    )

    $setup = Join-Path $PSScriptRoot '..' 'Private' 'setup-accessibility-checker.ps1'
    if (-not (Test-Path -LiteralPath $setup)) {
        throw "OfficeAccessibilityChecker is missing its bundled setup script at '$setup'. Reinstall the module."
    }

    $invokeArgs = @{}
    if ($Force) { $invokeArgs.Force = $true }

    & $setup @invokeArgs
}
