<#
.SYNOPSIS
    Convenience wrapper: ensures the Open XML SDK + Pester 5 are present,
    then runs the Pester suite under scripts/tests/.

.DESCRIPTION
    Idempotent setup. Re-runs are no-ops once the SDK and Pester 5+ are
    cached on the machine. Pester is installed into the CurrentUser scope
    only -- never machine-wide.
#>

[CmdletBinding()]
param(
    [ValidateSet('None','Normal','Detailed','Diagnostic')]
    [string] $Verbosity = 'Detailed'
)

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

if (-not (Get-Module -ListAvailable -Name Pester | Where-Object { $_.Version -ge [version]'5.0.0' })) {
    Write-Information 'Pester 5+ not found; installing into CurrentUser scope from PSGallery...'
    Install-Module Pester -MinimumVersion 5.5.0 -Force -SkipPublisherCheck -Scope CurrentUser
}

& (Join-Path (Split-Path -Parent $PSScriptRoot) 'setup-accessibility-checker.ps1')

$config = New-PesterConfiguration
$config.Run.Path        = $PSScriptRoot
$config.Output.Verbosity = $Verbosity

Invoke-Pester -Configuration $config
