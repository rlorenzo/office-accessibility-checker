<#
.SYNOPSIS
    Runs PSScriptAnalyzer against scripts\ using the project's pinned settings.

.DESCRIPTION
    Lints every .ps1 / .psd1 / .psm1 file under scripts\ with the rule
    exclusions defined in PSScriptAnalyzerSettings.psd1 at the project root.
    Exits non-zero if any finding is emitted, so this is safe to wire into
    pre-commit hooks or CI.

.NOTES
    Install PSScriptAnalyzer if missing:
        Install-Module PSScriptAnalyzer -Scope CurrentUser
#>

[CmdletBinding()]
param(
    [ValidateSet('Error','Warning','Information')]
    [string] $MinSeverity = 'Warning'
)

$ErrorActionPreference  = 'Stop'
$InformationPreference  = 'Continue'

if (-not (Get-Module -ListAvailable -Name PSScriptAnalyzer)) {
    [Console]::Error.WriteLine("PSScriptAnalyzer is not installed. Run: Install-Module PSScriptAnalyzer -Scope CurrentUser")
    exit 2
}

$settingsPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'PSScriptAnalyzerSettings.psd1'
$scriptsPath  = $PSScriptRoot

$results = Invoke-ScriptAnalyzer `
    -Path $scriptsPath `
    -Recurse `
    -Settings $settingsPath `
    -Severity $MinSeverity

if ($results) {
    $results | Format-Table -AutoSize Severity, RuleName, Line, ScriptName, Message | Out-String | Write-Information
    Write-Information ("FAIL  {0} finding(s) at severity >= {1}" -f $results.Count, $MinSeverity)
    exit 1
}

Write-Information "PASS  no PSScriptAnalyzer findings at severity >= $MinSeverity"
exit 0
