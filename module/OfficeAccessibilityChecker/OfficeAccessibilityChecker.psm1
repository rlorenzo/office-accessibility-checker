<#
.SYNOPSIS
    Module loader for OfficeAccessibilityChecker.

.DESCRIPTION
    Dot-sources every public function under Public/ so the manifest's
    FunctionsToExport list resolves at import time. The actual rule logic
    lives in Private/, which is populated from the canonical scripts/ in
    the repo by Build-Module.ps1 before publishing.
#>

$ErrorActionPreference = 'Stop'

$publicFunctions = Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public') -Filter '*.ps1' -ErrorAction SilentlyContinue
foreach ($file in $publicFunctions) {
    . $file.FullName
}
