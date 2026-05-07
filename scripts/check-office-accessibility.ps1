<#
.SYNOPSIS
    Dispatcher that runs the right OOXML accessibility checker for a Word or
    Excel file, mirroring the veraPDF-style CLI surface. Also supports bulk
    scanning when given a directory.

.PARAMETER Path
    Path to a .docx, .docm, .xlsx, or .xlsm file, or a directory containing
    such files. The legacy parameter name -FilePath is kept as an alias.

.PARAMETER Recurse
    When -Path is a directory, also descend into subdirectories. Ignored when
    -Path is a file.

.PARAMETER Format
    'text' (default) emits a single PASS/FAIL line per file. 'detailed' emits
    every issue found, one per line, followed by a PASS/FAIL summary per file.

.OUTPUTS
    Exit codes (single-file mode, and worst-of in bulk mode):
      0  no errors found
      1  one or more accessibility errors found
      2  tool error (file not found, unsupported format, SDK missing, etc.)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, Position=0)]
    [Alias('FilePath')]
    [string] $Path,

    [switch] $Recurse,

    [ValidateSet('text','detailed')]
    [string] $Format = 'text'
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $Path)) {
    [Console]::Error.WriteLine("Path not found: $Path")
    exit 2
}

$item = Get-Item -LiteralPath $Path

# --- Single-file dispatch (leaf) ----------------------------------------------
if (-not $item.PSIsContainer) {
    $ext = [IO.Path]::GetExtension($item.FullName).ToLowerInvariant()

    switch ($ext) {
        { $_ -in '.docx', '.docm' } {
            & (Join-Path $PSScriptRoot 'check-docx-accessibility.ps1') -FilePath $item.FullName -Format $Format
            exit $LASTEXITCODE
        }
        { $_ -in '.xlsx', '.xlsm' } {
            & (Join-Path $PSScriptRoot 'check-xlsx-accessibility.ps1') -FilePath $item.FullName -Format $Format
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
}

# --- Bulk scan (container) ----------------------------------------------------

$gciParams = @{
    LiteralPath = $item.FullName
    File        = $true
    Force       = $true
    ErrorAction = 'SilentlyContinue'
}
if ($Recurse) { $gciParams.Recurse = $true }

# Inaccessible subdirectories (permissions, reparse points) are skipped via
# SilentlyContinue rather than aborting the scan.
$allFiles = @(Get-ChildItem @gciParams)

# Office creates ~$<name> lock files in the same folder while a document is
# open. They share the .docx/.xlsx extension but are not real documents and
# would error noisily; filter them out.
$supported = @(
    $allFiles | Where-Object {
        $_.Extension -match '^\.(docx|docm|xlsx|xlsm)$' -and $_.Name -notlike '~$*'
    }
)
$skipped = $allFiles.Count - $supported.Count
$total   = $supported.Count

$pass = 0
$fail = 0
$err  = 0
$worstExit = 0

for ($i = 0; $i -lt $total; $i++) {
    $f = $supported[$i]

    if ($total -gt 1) {
        Write-Progress -Activity 'Scanning Office files' `
            -Status $f.FullName `
            -PercentComplete (($i / $total) * 100)
    }

    & $PSCommandPath -Path $f.FullName -Format $Format
    $exit = $LASTEXITCODE

    switch ($exit) {
        0 { $pass++ }
        1 { $fail++ }
        default {
            $err++
            # Child writes its diagnostic to stderr and exits without a stdout
            # PASS/FAIL line. Emit a stable stdout marker so downstream greps
            # see the failure.
            Write-Output "ERROR $($f.FullName)"
        }
    }
    if ($exit -gt $worstExit) { $worstExit = $exit }
}

if ($total -gt 1) {
    Write-Progress -Activity 'Scanning Office files' -Completed
}

[Console]::Error.WriteLine(
    "Scanned $total files: $pass passed, $fail failed, $err errors, $skipped skipped"
)
exit $worstExit
