function Test-OfficeAccessibility {
<#
.SYNOPSIS
    Run accessibility checks against a Word, Excel, or PowerPoint file (or a
    directory of them).

.DESCRIPTION
    Thin wrapper over the bundled check-office-accessibility.ps1 dispatcher.
    Returns the same PASS/FAIL output the script emits on stdout. The exit
    code from the underlying script is exposed via $LASTEXITCODE.

    On first use, run Initialize-OfficeAccessibilityChecker once to download
    the Open XML SDK into the module's Private/lib/ folder.

.PARAMETER Path
    A file (.docx/.docm/.xlsx/.xlsm/.pptx/.pptm) or a directory containing
    such files.

.PARAMETER Format
    'text' (default) emits one PASS/FAIL line per file. 'detailed' emits
    every issue, one tab-separated line each, then the PASS/FAIL summary.

.PARAMETER Recurse
    When -Path is a directory, also descend into subdirectories.

.PARAMETER Fix
    Write a sibling <basename>.fixed.<ext> with deterministic structural
    remediations applied. The original file is never modified. After fixing,
    the checker re-runs against the fixed file and prints its PASS/FAIL.

.EXAMPLE
    Test-OfficeAccessibility report.docx

.EXAMPLE
    Test-OfficeAccessibility ./docs -Recurse -Format detailed

.EXAMPLE
    Test-OfficeAccessibility report.docx -Fix
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string] $Path,

        [ValidateSet('text', 'detailed')]
        [string] $Format = 'text',

        [switch] $Recurse,

        [switch] $Fix
    )

    $dispatcher = Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'Private') 'check-office-accessibility.ps1'
    if (-not (Test-Path -LiteralPath $dispatcher)) {
        throw "OfficeAccessibilityChecker is missing its bundled scripts at '$dispatcher'. Reinstall the module."
    }

    $invokeArgs = @{ Path = $Path; Format = $Format }
    if ($Recurse) { $invokeArgs.Recurse = $true }
    if ($Fix)     { $invokeArgs.Fix     = $true }

    & $dispatcher @invokeArgs
}
