<#
.SYNOPSIS
    One-time setup: fetches the Open XML SDK DLLs into scripts/lib/.

.DESCRIPTION
    Downloads DocumentFormat.OpenXml (and required transitive deps) directly
    from nuget.org as .nupkg archives, extracts the runtime DLLs into
    scripts/lib/, and cleans up the temp folder. Idempotent — re-running
    is a no-op when the pinned DLL already exists.

.NOTES
    Pinned versions are the single source of truth for SDK versions used by
    all check-*-accessibility.ps1 scripts. Bump $Packages to upgrade.
#>

[CmdletBinding()]
param(
    [switch] $Force
)

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

# Pinned packages. DocumentFormat.OpenXml v3 split the assembly — the
# Framework package is a transitive dependency we must also fetch.
#
# Sha512 is the SHA-512 of the .nupkg as published on nuget.org. TLS + version
# pinning alone do not prove the bytes match the expected artifact (compromised
# feed, proxy cache, or local trust-store issue could swap content), so each
# download is hashed and rejected if it does not match.
$Packages = @(
    @{
        Name    = 'DocumentFormat.OpenXml'
        Version = '3.0.2'
        Sha512  = 'd521e5e470ae0ca9a00742f43a005ca21ca415190a511d478e3c40e2110b77335422dbffeff8283f70169ddbbd0b91ad26e385be56ed2eb96bffc7316230a501'
    }
    @{
        Name    = 'DocumentFormat.OpenXml.Framework'
        Version = '3.0.2'
        Sha512  = 'e124628d7e7e1c84774bd3f0244c1c96b71c5e19f608c218402d440f9f2aad102296a0a38f4e6f9fd16de48cef0b0d5f6c66dd33a17a4dfafe08efb16c3b8aa4'
    }
)

$LibDir  = Join-Path $PSScriptRoot 'lib'
$MainDll = Join-Path $LibDir 'DocumentFormat.OpenXml.dll'

if ((Test-Path $MainDll) -and -not $Force) {
    Write-Information "Open XML SDK already present at $LibDir (use -Force to refresh)."
    exit 0
}

if (-not (Test-Path $LibDir)) {
    New-Item -ItemType Directory -Path $LibDir | Out-Null
}

$tempDir = Join-Path ([IO.Path]::GetTempPath()) ("openxml-sdk-fetch-" + [Guid]::NewGuid())
New-Item -ItemType Directory -Path $tempDir | Out-Null

# Pick a target framework folder compatible with the host runtime.
#
# Windows PowerShell 5.1 runs on .NET Framework 4.x and cannot load net5.0+
# assemblies, so a naive "highest net*" preference produces a setup that looks
# successful but breaks every Add-Type call at checker invocation time. We pin
# to netstandard2.0/2.1 (loadable by both PS 5.1 and modern .NET) unless the
# package only ships net*/lib targets. When we do fall back to a net* folder,
# we filter to majors compatible with the running .NET runtime.
function Select-BestFramework {
    param([string[]] $FrameworkNames)

    if ($FrameworkNames -contains 'netstandard2.0') { return 'netstandard2.0' }
    if ($FrameworkNames -contains 'netstandard2.1') { return 'netstandard2.1' }

    # netstandard1.x is loadable on .NET Framework 4.6.1+ and netcoreapp 2.0+.
    $netstandard = $FrameworkNames | Where-Object { $_ -match '^netstandard\d' } | Sort-Object -Descending
    if ($netstandard) { return $netstandard | Select-Object -First 1 }

    # Only net*/lib targets remain. Pick the highest major <= the host runtime.
    # PSEdition 'Desktop' = Windows PowerShell on .NET Framework 4.x (no net5+).
    $hostMajor = if ($PSVersionTable.PSEdition -eq 'Desktop') { 4 }
                 else { [Environment]::Version.Major }

    $compatible = foreach ($n in $FrameworkNames) {
        if ($n -match '^net(\d+)(?:\.(\d+))?') {
            $major = [int]$Matches[1]
            # Folder names like "net48" / "net472" are .NET Framework 4.x.
            if ($major -ge 5 -and $major -le $hostMajor) {
                [pscustomobject]@{ Name = $n; Major = $major }
            }
            elseif ($major -eq 4 -and $hostMajor -eq 4) {
                [pscustomobject]@{ Name = $n; Major = $major }
            }
        }
    }
    if ($compatible) {
        return ($compatible | Sort-Object -Property Major -Descending | Select-Object -First 1).Name
    }

    # Nothing compatible — let the caller fail loudly rather than copy a
    # DLL that will only error at Add-Type time.
    return $null
}

try {
    $copied = @{}

    foreach ($pkg in $Packages) {
        $name    = $pkg.Name
        $version = $pkg.Version
        $url     = "https://www.nuget.org/api/v2/package/$name/$version"
        $nupkg   = Join-Path $tempDir "$name.$version.nupkg"
        $extract = Join-Path $tempDir "$name.$version"

        Write-Information "Downloading $name $version..."
        Invoke-WebRequest -Uri $url -OutFile $nupkg -UseBasicParsing

        $expected = $pkg.Sha512
        if ([string]::IsNullOrWhiteSpace($expected)) {
            throw "No pinned SHA-512 for $name $version. Refusing to install an unverified package."
        }
        # Get-FileHash returns uppercase; pinned values are lowercase.
        $actual = (Get-FileHash -LiteralPath $nupkg -Algorithm SHA512).Hash
        if ($actual -ine $expected) {
            throw "SHA-512 mismatch for $name $version.`nExpected: $expected`nActual:   $actual"
        }

        Write-Information "Extracting $name..."
        Expand-Archive -Path $nupkg -DestinationPath $extract -Force

        $libRoot = Join-Path $extract 'lib'
        if (-not (Test-Path $libRoot)) {
            throw "Package $name does not contain a lib/ directory."
        }

        $frameworks = Get-ChildItem -Path $libRoot -Directory | ForEach-Object Name
        $best       = Select-BestFramework -FrameworkNames $frameworks
        if (-not $best) {
            throw "Package $name has no usable framework folder under lib/."
        }
        $bestPath = Join-Path $libRoot $best
        Write-Information "  Using $best"

        Get-ChildItem -Path $bestPath -Filter '*.dll' | ForEach-Object {
            $dest = Join-Path $LibDir $_.Name
            Copy-Item -Path $_.FullName -Destination $dest -Force
            $copied[$_.Name] = $true
        }
    }

    if (-not (Test-Path $MainDll)) {
        throw "DocumentFormat.OpenXml.dll not found after extraction. Inspect $tempDir."
    }

    Write-Information ""
    Write-Information "Open XML SDK installed to $LibDir"
    Write-Information "Files:"
    $copied.Keys | Sort-Object | ForEach-Object { Write-Information "  $_" }
}
finally {
    if (Test-Path $tempDir) {
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
