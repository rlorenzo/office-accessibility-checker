<#
.SYNOPSIS
    Dispatcher that runs the right OOXML accessibility checker for a Word or
    Excel file, mirroring the veraPDF-style CLI surface.

.PARAMETER FilePath
    Path to a .docx, .docm, .xlsx, or .xlsm file.

.PARAMETER Format
    'text' (default) emits a single PASS/FAIL line. 'detailed' emits every
    issue found, one per line.

.OUTPUTS
    Exit codes:
      0  no errors found
      1  one or more accessibility errors found
      2  tool error (file not found, unsupported format, SDK missing, etc.)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $FilePath,

    [ValidateSet('text','detailed')]
    [string] $Format = 'text'
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $FilePath)) {
    [Console]::Error.WriteLine("File not found: $FilePath")
    exit 2
}

$ext = [IO.Path]::GetExtension($FilePath).ToLowerInvariant()

switch ($ext) {
    { $_ -in '.docx', '.docm' } {
        & (Join-Path $PSScriptRoot 'check-docx-accessibility.ps1') -FilePath $FilePath -Format $Format
        exit $LASTEXITCODE
    }
    { $_ -in '.xlsx', '.xlsm' } {
        & (Join-Path $PSScriptRoot 'check-xlsx-accessibility.ps1') -FilePath $FilePath -Format $Format
        exit $LASTEXITCODE
    }
    { $_ -in '.doc', '.xls' } {
        [Console]::Error.WriteLine("Unsupported format: legacy binary Office files ($ext) are not supported. Re-save as .docx/.xlsx.")
        exit 2
    }
    default {
        [Console]::Error.WriteLine("Unsupported file type: $ext")
        exit 2
    }
}
