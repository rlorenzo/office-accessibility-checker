<#
.SYNOPSIS
    Runs Microsoft Word-style accessibility rules against a .docx/.docm file
    using the Open XML SDK.

.DESCRIPTION
    Inspects the document body, every header part, and every footer part for
    common accessibility issues (missing alt text, missing table headers,
    document protection, missing content control titles, merged cells,
    heading-order skips, floating objects, repeated blank characters, lack of
    headings, low contrast). Emits a single PASS/FAIL line in 'text' mode, or
    a sorted list of issues followed by the PASS/FAIL line in 'detailed' mode.

.PARAMETER FilePath
    Path to a .docx or .docm file.

.PARAMETER Format
    'text' (default) emits only PASS/FAIL. 'detailed' emits one tab-separated
    issue per line plus the PASS/FAIL summary.

.PARAMETER Fix
    When set, after running the checks the script writes a sibling file named
    <basename>.fixed.docx with deterministic structural remediations applied
    for the rules MissingTableHeaders, RepeatedBlanks, and LowContrast. The
    original file is never modified. After fixing, the checker is re-run on
    the fixed file and its PASS/FAIL is appended.

.OUTPUTS
    Exit codes:
      0  no errors found (warnings/tips do not fail)
      1  one or more accessibility errors found (includes IRM/password-protected files)
      2  tool error (file not found, unsupported format, SDK load failure, file locked)

      With -Fix, the exit code reflects the FIXED file when fixes were applied,
      otherwise the original file.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $FilePath,

    [ValidateSet('text','detailed')]
    [string] $Format = 'text',

    [switch] $Fix
)

$ErrorActionPreference = 'Stop'

# --- Pinned OOXML namespaces -------------------------------------------------
$NS = @{
    w   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    wp  = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    pic = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
    wps = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
    wpg = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
    a   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
}

# --- Issue record helper ------------------------------------------------------
function New-Issue {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'New-Issue is a pure record factory — it constructs a [pscustomobject] and has no side effects. The verb New triggers this rule, but the function does not change any system state.'
    )]
    param(
        [ValidateSet('ERROR','WARNING','TIP')]
        [string] $Severity,
        [string] $RuleName,
        [string] $Description
    )
    [pscustomobject]@{
        Severity    = $Severity
        RuleName    = $RuleName
        Description = $Description
    }
}

# --- Pre-flight: file existence and extension --------------------------------
if (-not (Test-Path -LiteralPath $FilePath)) {
    [Console]::Error.WriteLine("File not found: $FilePath")
    exit 2
}

$ext = [IO.Path]::GetExtension($FilePath).ToLowerInvariant()
if ($ext -ne '.docx' -and $ext -ne '.docm') {
    [Console]::Error.WriteLine("Unsupported format: expected .docx or .docm, got $ext")
    exit 2
}

# --- Locate and load the Open XML SDK ----------------------------------------
$sdkPath = $null
if ($env:OPENXML_SDK_PATH -and (Test-Path -LiteralPath $env:OPENXML_SDK_PATH)) {
    $sdkPath = $env:OPENXML_SDK_PATH
} else {
    $candidate = Join-Path $PSScriptRoot 'lib\DocumentFormat.OpenXml.dll'
    if (Test-Path -LiteralPath $candidate) {
        $sdkPath = $candidate
    }
}

if (-not $sdkPath) {
    [Console]::Error.WriteLine("Open XML SDK not found. Run scripts\setup-accessibility-checker.ps1 first.")
    exit 2
}

try {
    Add-Type -Path $sdkPath -ErrorAction Stop
} catch {
    [Console]::Error.WriteLine("Failed to load Open XML SDK from '$sdkPath': $($_.Exception.Message)")
    exit 2
}

# --- XML helpers --------------------------------------------------------------
function Get-PartXDocument {
    param($Part)
    if (-not $Part) { return $null }
    $stream = $Part.GetStream([IO.FileMode]::Open, [IO.FileAccess]::Read)
    try {
        $reader = [Xml.XmlReader]::Create($stream)
        try {
            return [Xml.Linq.XDocument]::Load($reader)
        } finally {
            $reader.Dispose()
        }
    } finally {
        $stream.Dispose()
    }
}

function Get-XAttr {
    param([Xml.Linq.XElement] $Element, [string] $Namespace, [string] $LocalName)
    if (-not $Element) { return $null }
    $attr = $Element.Attribute([Xml.Linq.XName]::Get($LocalName, $Namespace))
    if ($attr) { return $attr.Value }
    # Try no-namespace fallback (some attributes are unqualified).
    $attr = $Element.Attribute([Xml.Linq.XName]::Get($LocalName))
    if ($attr) { return $attr.Value }
    return $null
}

function Get-Descendant {
    param([Xml.Linq.XContainer] $Root, [string] $Namespace, [string] $LocalName)
    if (-not $Root) { return @() }
    return $Root.Descendants([Xml.Linq.XName]::Get($LocalName, $Namespace))
}

function Get-ChildElement {
    param([Xml.Linq.XElement] $Parent, [string] $Namespace, [string] $LocalName)
    if (-not $Parent) { return $null }
    return $Parent.Element([Xml.Linq.XName]::Get($LocalName, $Namespace))
}

# --- Rule: MissingAltText -----------------------------------------------------
# Walk every wp:anchor and wp:inline drawing wrapper. Pass if:
#   - wp:docPr/@descr non-empty, OR
#   - wp:docPr/@title non-empty, OR
#   - wp:docPr/@decorative="1"
# Then recurse into wpg:wgp groups: every child shape's per-shape cNvPr
# (pic:cNvPr / wps:cNvPr / a:graphicFrame//a:cNvPr / nested wpg) gets the
# same check.
function Test-AltText {
    param([Xml.Linq.XContainer] $Root, [ref] $Issues)

    if (-not $Root) { return }

    $wpName = [Xml.Linq.XName]::Get('anchor', $NS.wp)
    $inlineName = [Xml.Linq.XName]::Get('inline', $NS.wp)

    $wrappers = @()
    $wrappers += $Root.Descendants($wpName)
    $wrappers += $Root.Descendants($inlineName)

    foreach ($wrapper in $wrappers) {
        $docPr = Get-ChildElement -Parent $wrapper -Namespace $NS.wp -LocalName 'docPr'
        $name = if ($docPr) { Get-XAttr -Element $docPr -Namespace $null -LocalName 'name' } else { $null }
        if ([string]::IsNullOrEmpty($name)) { $name = '(unnamed)' }

        if (-not (Test-AltTextAttribute $docPr)) {
            $Issues.Value += New-Issue -Severity ERROR -RuleName 'MissingAltText' `
                -Description "Image/object `"$name`" has no alt text"
        }

        # Recurse into group shapes.
        $groups = Get-Descendant -Root $wrapper -Namespace $NS.wpg -LocalName 'wgp'
        foreach ($grp in $groups) {
            Test-GroupAltText -Group $grp -Issues $Issues
        }
    }
}

# Pass if any of: @descr non-empty, @title non-empty, @decorative="1".
# Applies uniformly to wp:docPr (drawing wrapper) and per-shape *:cNvPr.
function Test-AltTextAttribute {
    param([Xml.Linq.XElement] $Element)
    if (-not $Element) { return $false }
    if (-not [string]::IsNullOrEmpty((Get-XAttr -Element $Element -Namespace $null -LocalName 'descr'))) { return $true }
    if (-not [string]::IsNullOrEmpty((Get-XAttr -Element $Element -Namespace $null -LocalName 'title'))) { return $true }
    if ((Get-XAttr -Element $Element -Namespace $null -LocalName 'decorative') -eq '1') { return $true }
    return $false
}

function Test-GroupAltText {
    param([Xml.Linq.XElement] $Group, [ref] $Issues)
    if (-not $Group) { return }

    # Direct child shapes of the group.
    foreach ($child in $Group.Elements()) {
        $ln = $child.Name.LocalName
        $nsName = $child.Name.NamespaceName
        switch ($ln) {
            'pic' {
                if ($nsName -eq $NS.pic) {
                    $nvPicPr = Get-ChildElement -Parent $child -Namespace $NS.pic -LocalName 'nvPicPr'
                    $cNvPr = Get-ChildElement -Parent $nvPicPr -Namespace $NS.pic -LocalName 'cNvPr'
                    $name = if ($cNvPr) { Get-XAttr -Element $cNvPr -Namespace $null -LocalName 'name' } else { '(unnamed)' }
                    if ([string]::IsNullOrEmpty($name)) { $name = '(unnamed)' }
                    if (-not (Test-AltTextAttribute $cNvPr)) {
                        $Issues.Value += New-Issue -Severity ERROR -RuleName 'MissingAltText' `
                            -Description "Image/object `"$name`" has no alt text"
                    }
                }
            }
            'wsp' {
                if ($nsName -eq $NS.wps) {
                    $nvSpPr = Get-ChildElement -Parent $child -Namespace $NS.wps -LocalName 'nvSpPr'
                    $cNvPr = Get-ChildElement -Parent $nvSpPr -Namespace $NS.wps -LocalName 'cNvPr'
                    $name = if ($cNvPr) { Get-XAttr -Element $cNvPr -Namespace $null -LocalName 'name' } else { '(unnamed)' }
                    if ([string]::IsNullOrEmpty($name)) { $name = '(unnamed)' }
                    if (-not (Test-AltTextAttribute $cNvPr)) {
                        $Issues.Value += New-Issue -Severity ERROR -RuleName 'MissingAltText' `
                            -Description "Image/object `"$name`" has no alt text"
                    }
                }
            }
            'graphicFrame' {
                if ($nsName -eq $NS.a -or $nsName -eq $NS.wpg) {
                    # Look for a cNvPr descendant in any namespace.
                    $cNvPr = $child.Descendants() | Where-Object { $_.Name.LocalName -eq 'cNvPr' } | Select-Object -First 1
                    $name = if ($cNvPr) { Get-XAttr -Element $cNvPr -Namespace $null -LocalName 'name' } else { '(unnamed)' }
                    if ([string]::IsNullOrEmpty($name)) { $name = '(unnamed)' }
                    if (-not (Test-AltTextAttribute $cNvPr)) {
                        $Issues.Value += New-Issue -Severity ERROR -RuleName 'MissingAltText' `
                            -Description "Image/object `"$name`" has no alt text"
                    }
                }
            }
            'wgp' {
                if ($nsName -eq $NS.wpg) {
                    Test-GroupAltText -Group $child -Issues $Issues
                }
            }
        }
    }
}

# --- Rule: MissingTableHeaders -----------------------------------------------
# Pass only if the first w:tr carries w:trPr/w:tblHeader (presence; default
# value="1"). w:tblLook is table-style metadata that drives visual first-row
# formatting; it does not expose header semantics to assistive technology, so
# it must NOT satisfy this rule.
#
# Layout tables — tables used purely for visual arrangement (e.g. photo grids)
# rather than tabular data — are exempt. A table with w:tblPr/w:tblDescription
# is treated as declaring its purpose to assistive tech; forcing tblHeader on
# the first row would mislabel arrangement cells as data headers, so we use
# the presence of tblDescription as the OOXML signal of layout intent.
function Test-TableHeader {
    param([Xml.Linq.XContainer] $Root, [int] $StartIndex, [ref] $Issues)
    if (-not $Root) { return $StartIndex }

    $tables = Get-Descendant -Root $Root -Namespace $NS.w -LocalName 'tbl'
    $idx = $StartIndex
    foreach ($tbl in $tables) {
        $idx++
        $rows = $tbl.Elements([Xml.Linq.XName]::Get('tr', $NS.w))
        if (-not $rows -or $rows.Count -eq 0) { continue }

        $tblPr = Get-ChildElement -Parent $tbl -Namespace $NS.w -LocalName 'tblPr'
        if ($tblPr) {
            $tblDescription = Get-ChildElement -Parent $tblPr -Namespace $NS.w -LocalName 'tblDescription'
            if ($tblDescription) {
                # Layout table — skip the header check.
                continue
            }
        }

        $firstRow = $rows | Select-Object -First 1
        $hasHeader = $false

        $trPr = Get-ChildElement -Parent $firstRow -Namespace $NS.w -LocalName 'trPr'
        if ($trPr) {
            $tblHeader = Get-ChildElement -Parent $trPr -Namespace $NS.w -LocalName 'tblHeader'
            if ($tblHeader) {
                $val = Get-XAttr -Element $tblHeader -Namespace $NS.w -LocalName 'val'
                # Default is "1" (true) if attr absent.
                if ([string]::IsNullOrEmpty($val) -or $val -eq '1' -or $val -eq 'true' -or $val -eq 'on') {
                    $hasHeader = $true
                }
            }
        }

        if (-not $hasHeader) {
            $Issues.Value += New-Issue -Severity ERROR -RuleName 'MissingTableHeaders' `
                -Description "Table $idx does not have header information"
        }
    }
    return $idx
}

# --- Rule: MissingContentControlTitle ----------------------------------------
function Test-ContentControlTitle {
    param([Xml.Linq.XContainer] $Root, [int] $StartIndex, [ref] $Issues)
    if (-not $Root) { return $StartIndex }

    $sdts = Get-Descendant -Root $Root -Namespace $NS.w -LocalName 'sdt'
    $idx = $StartIndex
    foreach ($sdt in $sdts) {
        $idx++
        $sdtPr = Get-ChildElement -Parent $sdt -Namespace $NS.w -LocalName 'sdtPr'
        $alias = Get-ChildElement -Parent $sdtPr -Namespace $NS.w -LocalName 'alias'
        $val = if ($alias) { Get-XAttr -Element $alias -Namespace $NS.w -LocalName 'val' } else { $null }
        if ([string]::IsNullOrEmpty($val)) {
            $Issues.Value += New-Issue -Severity ERROR -RuleName 'MissingContentControlTitle' `
                -Description "Content control at position $idx has no title"
        }
    }
    return $idx
}

# --- Rule: MergedTableCells ---------------------------------------------------
function Test-MergedCell {
    param([Xml.Linq.XContainer] $Root, [int] $StartIndex, [ref] $Issues)
    if (-not $Root) { return $StartIndex }

    $tables = Get-Descendant -Root $Root -Namespace $NS.w -LocalName 'tbl'
    $idx = $StartIndex
    foreach ($tbl in $tables) {
        $idx++
        $cells = Get-Descendant -Root $tbl -Namespace $NS.w -LocalName 'tc'
        if (-not $cells -or $cells.Count -eq 0) { continue }

        $hasMerge = $false

        foreach ($tc in $cells) {
            $tcPr = Get-ChildElement -Parent $tc -Namespace $NS.w -LocalName 'tcPr'
            if (-not $tcPr) { continue }

            $vMerge = Get-ChildElement -Parent $tcPr -Namespace $NS.w -LocalName 'vMerge'
            if ($vMerge) {
                # Continuation of vertical merge if no val or val != "restart".
                $vmVal = Get-XAttr -Element $vMerge -Namespace $NS.w -LocalName 'val'
                if ([string]::IsNullOrEmpty($vmVal) -or $vmVal -ne 'restart') {
                    $hasMerge = $true
                    break
                }
            }

            $gridSpan = Get-ChildElement -Parent $tcPr -Namespace $NS.w -LocalName 'gridSpan'
            if ($gridSpan) {
                $gsVal = Get-XAttr -Element $gridSpan -Namespace $NS.w -LocalName 'val'
                $gsInt = 0
                if ([int]::TryParse($gsVal, [ref] $gsInt)) {
                    if ($gsInt -gt 1) {
                        $hasMerge = $true
                        break
                    }
                }
            }
        }

        if (-not $hasMerge) {
            # Detect nested table: any descendant w:tbl whose parent chain includes this $tbl through a w:tc.
            $nested = Get-Descendant -Root $tbl -Namespace $NS.w -LocalName 'tbl'
            foreach ($nt in $nested) {
                if ($nt -ne $tbl) {
                    $hasMerge = $true
                    break
                }
            }
        }

        if ($hasMerge) {
            $Issues.Value += New-Issue -Severity WARNING -RuleName 'MergedTableCells' `
                -Description "Table $idx contains merged or nested cells"
        }
    }
    return $idx
}

# --- Rule: HeadingOrderSkip ---------------------------------------------------
# Walks paragraphs in document order, tracks heading levels from
# w:pPr/w:pStyle/@w:val matching ^Heading([1-9])$. Warn if level > prev + 1.
function Test-HeadingOrder {
    param([Xml.Linq.XContainer] $Root, [ref] $Issues)
    if (-not $Root) { return }

    $paras = Get-Descendant -Root $Root -Namespace $NS.w -LocalName 'p'
    $prevLevel = 0
    foreach ($p in $paras) {
        $pPr = Get-ChildElement -Parent $p -Namespace $NS.w -LocalName 'pPr'
        $pStyle = Get-ChildElement -Parent $pPr -Namespace $NS.w -LocalName 'pStyle'
        $val = if ($pStyle) { Get-XAttr -Element $pStyle -Namespace $NS.w -LocalName 'val' } else { $null }
        if ([string]::IsNullOrEmpty($val)) { continue }
        if ($val -match '^Heading([1-9])$') {
            $level = [int] $Matches[1]
            if ($prevLevel -gt 0 -and $level -gt ($prevLevel + 1)) {
                $Issues.Value += New-Issue -Severity WARNING -RuleName 'HeadingOrderSkip' `
                    -Description "Heading level skipped from $prevLevel to $level"
            }
            $prevLevel = $level
        }
    }
}

# --- Rule: FloatingObject -----------------------------------------------------
function Test-FloatingObject {
    param([Xml.Linq.XContainer] $Root, [ref] $Issues)
    if (-not $Root) { return }

    $anchors = Get-Descendant -Root $Root -Namespace $NS.wp -LocalName 'anchor'
    foreach ($a in $anchors) {
        $docPr = Get-ChildElement -Parent $a -Namespace $NS.wp -LocalName 'docPr'
        $name = if ($docPr) { Get-XAttr -Element $docPr -Namespace $null -LocalName 'name' } else { $null }
        if ([string]::IsNullOrEmpty($name)) { $name = '(unnamed)' }
        $Issues.Value += New-Issue -Severity WARNING -RuleName 'FloatingObject' `
            -Description "Object `"$name`" is floating (not inline with text)"
    }
}

# --- Rule: RepeatedBlanks -----------------------------------------------------
# Concat w:t per paragraph; flag runs of 3+ space (U+0020) or NBSP (U+00A0).
# Tabs are NOT flagged (per spec).
function Test-RepeatedBlank {
    param([Xml.Linq.XContainer] $Root, [ref] $Issues)
    if (-not $Root) { return }

    $paras = Get-Descendant -Root $Root -Namespace $NS.w -LocalName 'p'
    foreach ($p in $paras) {
        $texts = Get-Descendant -Root $p -Namespace $NS.w -LocalName 't'
        if (-not $texts -or $texts.Count -eq 0) { continue }
        $sb = New-Object System.Text.StringBuilder
        foreach ($t in $texts) {
            [void] $sb.Append($t.Value)
        }
        $combined = $sb.ToString()
        if ($combined -match '[  ]{3,}') {
            $Issues.Value += New-Issue -Severity WARNING -RuleName 'RepeatedBlanks' `
                -Description 'Paragraph contains repeated blank characters'
        }
    }
}

# --- Rule: NoHeadingStyles (TIP) ---------------------------------------------
function Test-NoHeadingStyle {
    param([Xml.Linq.XContainer[]] $Roots, [ref] $Issues)
    $found = $false
    foreach ($root in $Roots) {
        if (-not $root) { continue }
        $pStyles = Get-Descendant -Root $root -Namespace $NS.w -LocalName 'pStyle'
        foreach ($ps in $pStyles) {
            $val = Get-XAttr -Element $ps -Namespace $NS.w -LocalName 'val'
            if (-not [string]::IsNullOrEmpty($val) -and $val -match '^Heading[1-9]$') {
                $found = $true
                break
            }
        }
        if ($found) { break }
    }
    if (-not $found) {
        $Issues.Value += New-Issue -Severity TIP -RuleName 'NoHeadingStyles' `
            -Description 'Document contains no heading-style paragraphs'
    }
}

# --- Contrast helpers --------------------------------------------------------
# WCAG 2.x relative luminance and contrast ratio.
# https://www.w3.org/TR/WCAG21/#dfn-relative-luminance
function Get-RelativeLuminance {
    param([Parameter(Mandatory)] [string] $Hex)
    $r = [Convert]::ToInt32($Hex.Substring(0, 2), 16) / 255.0
    $g = [Convert]::ToInt32($Hex.Substring(2, 2), 16) / 255.0
    $b = [Convert]::ToInt32($Hex.Substring(4, 2), 16) / 255.0
    $rL = if ($r -le 0.03928) { $r / 12.92 } else { [Math]::Pow(($r + 0.055) / 1.055, 2.4) }
    $gL = if ($g -le 0.03928) { $g / 12.92 } else { [Math]::Pow(($g + 0.055) / 1.055, 2.4) }
    $bL = if ($b -le 0.03928) { $b / 12.92 } else { [Math]::Pow(($b + 0.055) / 1.055, 2.4) }
    return 0.2126 * $rL + 0.7152 * $gL + 0.0722 * $bL
}

function Get-ContrastRatio {
    param(
        [Parameter(Mandatory)] [string] $ForegroundHex,
        [Parameter(Mandatory)] [string] $BackgroundHex
    )
    $fgL = Get-RelativeLuminance -Hex $ForegroundHex
    $bgL = Get-RelativeLuminance -Hex $BackgroundHex
    $lighter = [Math]::Max($fgL, $bgL)
    $darker  = [Math]::Min($fgL, $bgL)
    return ($lighter + 0.05) / ($darker + 0.05)
}

# Returns the explicit 6-digit hex value of a w:shd element's fill, or $null
# if the shading uses a theme reference, "auto", or no fill.
function Get-ExplicitShadingFill {
    param([Xml.Linq.XElement] $ShadingParent)
    if (-not $ShadingParent) { return $null }
    $shd = Get-ChildElement -Parent $ShadingParent -Namespace $NS.w -LocalName 'shd'
    if (-not $shd) { return $null }
    if (-not [string]::IsNullOrEmpty((Get-XAttr -Element $shd -Namespace $NS.w -LocalName 'themeFill'))) { return $null }
    $fill = Get-XAttr -Element $shd -Namespace $NS.w -LocalName 'fill'
    if ([string]::IsNullOrEmpty($fill) -or $fill -eq 'auto') { return $null }
    if ($fill -match '^[0-9A-Fa-f]{6}$') { return $fill }
    return $null
}

# --- Rule: LowContrast (WARNING) ---------------------------------------------
# Best-effort contrast check. Only flags runs where BOTH the foreground (run
# color) and background (run shading or paragraph shading) are explicit 6-digit
# hex values. Skips: "auto" colors, themeColor / themeFill references, runs
# with no explicit color, paragraphs with no explicit shading. False positives
# from style/theme inheritance are avoided by refusing to guess.
#
# Threshold follows WCAG 2.x:
#   3.0:1 for large text (18pt+, or 14pt+ bold)
#   4.5:1 otherwise
# Font size in OOXML is in half-points: 22 -> 11pt (default), 36 -> 18pt,
# 28 -> 14pt.
function Test-LowContrast {
    param([Xml.Linq.XContainer] $Root, [ref] $Issues)
    if (-not $Root) { return }

    $runs = Get-Descendant -Root $Root -Namespace $NS.w -LocalName 'r'
    foreach ($run in $runs) {
        $rPr = Get-ChildElement -Parent $run -Namespace $NS.w -LocalName 'rPr'
        if (-not $rPr) { continue }

        $colorEl = Get-ChildElement -Parent $rPr -Namespace $NS.w -LocalName 'color'
        if (-not $colorEl) { continue }
        if (-not [string]::IsNullOrEmpty((Get-XAttr -Element $colorEl -Namespace $NS.w -LocalName 'themeColor'))) { continue }
        $fgVal = Get-XAttr -Element $colorEl -Namespace $NS.w -LocalName 'val'
        if ([string]::IsNullOrEmpty($fgVal) -or $fgVal -eq 'auto') { continue }
        if ($fgVal -notmatch '^[0-9A-Fa-f]{6}$') { continue }

        $bgVal = Get-ExplicitShadingFill -ShadingParent $rPr
        if (-not $bgVal) {
            $p = $run.Parent
            while ($p -and $p.Name.LocalName -ne 'p') { $p = $p.Parent }
            if ($p) {
                $pPr = Get-ChildElement -Parent $p -Namespace $NS.w -LocalName 'pPr'
                if ($pPr) { $bgVal = Get-ExplicitShadingFill -ShadingParent $pPr }
            }
        }
        if (-not $bgVal) { continue }

        $halfPts = 22
        $sz = Get-ChildElement -Parent $rPr -Namespace $NS.w -LocalName 'sz'
        if ($sz) {
            $szVal = Get-XAttr -Element $sz -Namespace $NS.w -LocalName 'val'
            $parsed = 0
            if ([int]::TryParse($szVal, [ref] $parsed)) { $halfPts = $parsed }
        }
        $isBold = ($null -ne (Get-ChildElement -Parent $rPr -Namespace $NS.w -LocalName 'b'))
        $isLarge = ($halfPts -ge 36) -or ($isBold -and $halfPts -ge 28)
        $threshold = if ($isLarge) { 3.0 } else { 4.5 }

        $ratio = Get-ContrastRatio -ForegroundHex $fgVal -BackgroundHex $bgVal
        if ($ratio -lt $threshold) {
            $sb = New-Object System.Text.StringBuilder
            foreach ($t in (Get-Descendant -Root $run -Namespace $NS.w -LocalName 't')) { [void] $sb.Append($t.Value) }
            $preview = $sb.ToString()
            if ([string]::IsNullOrEmpty($preview)) { $preview = '(empty run)' }
            elseif ($preview.Length -gt 30) { $preview = $preview.Substring(0, 30) + '...' }
            $Issues.Value += New-Issue -Severity WARNING -RuleName 'LowContrast' `
                -Description ("Text `"{0}`" has contrast {1:N2}:1 (#{2} on #{3}); WCAG requires {4}:1" -f $preview, $ratio, $fgVal.ToUpper(), $bgVal.ToUpper(), $threshold)
        }
    }
}

# --- Autofix helpers ---------------------------------------------------------
# Scope: only the rules whose remediation is a deterministic structural OOXML
# edit. Anything that requires authoring (alt text, content control title,
# heading promotion, descriptive link text), is destructive (unmerge cells,
# remove protection), or changes layout (anchor -> inline) is intentionally
# out of scope.
#
# All fix functions mutate the supplied XDocument in place and return the
# count of items they fixed. Saving the part is the caller's job.

# HSL <-> sRGB conversion. Used by Get-NearestPassingForeground to find the
# perceptually-closest foreground color (preserving hue and saturation) that
# meets a given WCAG contrast ratio against a fixed background. CIELCh would
# be more uniform but HSL is good enough for "nudge until it passes" and keeps
# the math (and the runtime cost per low-contrast run) small.
function ConvertTo-Hsl {
    param([Parameter(Mandatory)] [string] $Hex)
    $r = [Convert]::ToInt32($Hex.Substring(0, 2), 16) / 255.0
    $g = [Convert]::ToInt32($Hex.Substring(2, 2), 16) / 255.0
    $b = [Convert]::ToInt32($Hex.Substring(4, 2), 16) / 255.0
    $max = [Math]::Max([Math]::Max($r, $g), $b)
    $min = [Math]::Min([Math]::Min($r, $g), $b)
    $L = ($max + $min) / 2.0
    if ($max -eq $min) { return [pscustomobject]@{ H = 0.0; S = 0.0; L = $L } }
    $d = $max - $min
    $S = if ($L -gt 0.5) { $d / (2.0 - $max - $min) } else { $d / ($max + $min) }
    $H = 0.0
    if     ($max -eq $r) { $H = (($g - $b) / $d) + ($(if ($g -lt $b) { 6.0 } else { 0.0 })) }
    elseif ($max -eq $g) { $H = (($b - $r) / $d) + 2.0 }
    else                 { $H = (($r - $g) / $d) + 4.0 }
    return [pscustomobject]@{ H = ($H / 6.0); S = $S; L = $L }
}

function Get-HueChannel {
    param([double] $p, [double] $q, [double] $t)
    if ($t -lt 0) { $t += 1 }
    if ($t -gt 1) { $t -= 1 }
    if ($t -lt (1.0 / 6.0)) { return $p + ($q - $p) * 6.0 * $t }
    if ($t -lt 0.5) { return $q }
    if ($t -lt (2.0 / 3.0)) { return $p + ($q - $p) * ((2.0 / 3.0) - $t) * 6.0 }
    return $p
}

function ConvertFrom-Hsl {
    param([Parameter(Mandatory)] [double] $H,
          [Parameter(Mandatory)] [double] $S,
          [Parameter(Mandatory)] [double] $L)
    if ($S -eq 0) { $r = $L; $g = $L; $b = $L }
    else {
        $q = if ($L -lt 0.5) { $L * (1.0 + $S) } else { $L + $S - ($L * $S) }
        $p = (2.0 * $L) - $q
        $r = Get-HueChannel $p $q ($H + (1.0 / 3.0))
        $g = Get-HueChannel $p $q $H
        $b = Get-HueChannel $p $q ($H - (1.0 / 3.0))
    }
    $rByte = [int][Math]::Round($r * 255)
    $gByte = [int][Math]::Round($g * 255)
    $bByte = [int][Math]::Round($b * 255)
    return ('{0:X2}{1:X2}{2:X2}' -f $rByte, $gByte, $bByte)
}

# Find the smallest |L - L_orig| at which the foreground passes contrast.
# Search both directions (darkening, lightening) and pick the closer side.
# Step 0.005 = 200 candidates per direction worst-case; far smaller in practice.
function Get-NearestPassingForeground {
    param(
        [Parameter(Mandatory)] [string] $ForegroundHex,
        [Parameter(Mandatory)] [string] $BackgroundHex,
        [double] $Threshold = 4.5
    )
    if ((Get-ContrastRatio -ForegroundHex $ForegroundHex -BackgroundHex $BackgroundHex) -ge $Threshold) {
        return $ForegroundHex
    }
    $hsl = ConvertTo-Hsl -Hex $ForegroundHex
    $step = 0.005

    $bestDark = $null; $deltaDark = [double]::PositiveInfinity
    for ($L = $hsl.L - $step; $L -ge 0; $L -= $step) {
        $cand = ConvertFrom-Hsl -H $hsl.H -S $hsl.S -L $L
        if ((Get-ContrastRatio -ForegroundHex $cand -BackgroundHex $BackgroundHex) -ge $Threshold) {
            $bestDark = $cand; $deltaDark = $hsl.L - $L; break
        }
    }
    $bestLight = $null; $deltaLight = [double]::PositiveInfinity
    for ($L = $hsl.L + $step; $L -le 1; $L += $step) {
        $cand = ConvertFrom-Hsl -H $hsl.H -S $hsl.S -L $L
        if ((Get-ContrastRatio -ForegroundHex $cand -BackgroundHex $BackgroundHex) -ge $Threshold) {
            $bestLight = $cand; $deltaLight = $L - $hsl.L; break
        }
    }
    if ($bestDark -and $deltaDark -le $deltaLight) { return $bestDark }
    if ($bestLight) { return $bestLight }

    # Mid-tone bg with no passing color along the lightness axis: fall back to
    # whichever pole has more contrast headroom.
    $blackR = Get-ContrastRatio -ForegroundHex '000000' -BackgroundHex $BackgroundHex
    $whiteR = Get-ContrastRatio -ForegroundHex 'FFFFFF' -BackgroundHex $BackgroundHex
    if ($blackR -ge $whiteR) { return '000000' } else { return 'FFFFFF' }
}

function Repair-WordTableHeader {
    param([Xml.Linq.XContainer] $Root)
    if (-not $Root) { return 0 }
    $fixed = 0
    $tables = Get-Descendant -Root $Root -Namespace $NS.w -LocalName 'tbl'
    foreach ($tbl in $tables) {
        $rows = $tbl.Elements([Xml.Linq.XName]::Get('tr', $NS.w))
        if (-not $rows -or $rows.Count -eq 0) { continue }

        $tblPr = Get-ChildElement -Parent $tbl -Namespace $NS.w -LocalName 'tblPr'
        if ($tblPr) {
            # Layout-table opt-out: skip tables that declare themselves as
            # layout via tblDescription, matching the rule.
            if (Get-ChildElement -Parent $tblPr -Namespace $NS.w -LocalName 'tblDescription') { continue }
        }

        $firstRow = $rows | Select-Object -First 1
        $trPr = Get-ChildElement -Parent $firstRow -Namespace $NS.w -LocalName 'trPr'
        if ($trPr -and (Get-ChildElement -Parent $trPr -Namespace $NS.w -LocalName 'tblHeader')) { continue }

        if (-not $trPr) {
            $trPr = New-Object Xml.Linq.XElement -ArgumentList ([Xml.Linq.XName]::Get('trPr', $NS.w))
            # w:trPr must be the first child of w:tr per the schema.
            $firstRow.AddFirst($trPr)
        }
        $tblHeader = New-Object Xml.Linq.XElement -ArgumentList ([Xml.Linq.XName]::Get('tblHeader', $NS.w))
        $trPr.Add($tblHeader)
        $fixed++
    }
    return $fixed
}

# Replace runs of 3+ space (U+0020) or NBSP (U+00A0) with a single tab character
# inside each w:t element. Cross-w:t blank runs are not handled (the typical
# case is "user typed N spaces in one run"); they would require splitting and
# re-stitching runs, which we deliberately avoid. xml:space="preserve" must be
# set on any modified w:t so the tab survives serialization.
function Repair-WordRepeatedBlank {
    param([Xml.Linq.XContainer] $Root)
    if (-not $Root) { return 0 }
    $fixed = 0
    $texts = Get-Descendant -Root $Root -Namespace $NS.w -LocalName 't'
    foreach ($t in $texts) {
        $orig = $t.Value
        if ($orig -notmatch '[  ]{3,}') { continue }
        $new = [regex]::Replace($orig, "[  ]{3,}", "`t")
        if ($new -eq $orig) { continue }
        $t.Value = $new
        $xmlNs = 'http://www.w3.org/XML/1998/namespace'
        $existing = $t.Attribute([Xml.Linq.XName]::Get('space', $xmlNs))
        if (-not $existing) {
            $t.SetAttributeValue([Xml.Linq.XName]::Get('space', $xmlNs), 'preserve')
        }
        $fixed++
    }
    return $fixed
}

function Repair-WordLowContrast {
    param([Xml.Linq.XContainer] $Root)
    if (-not $Root) { return 0 }
    $fixed = 0
    $runs = Get-Descendant -Root $Root -Namespace $NS.w -LocalName 'r'
    foreach ($run in $runs) {
        $rPr = Get-ChildElement -Parent $run -Namespace $NS.w -LocalName 'rPr'
        if (-not $rPr) { continue }
        $colorEl = Get-ChildElement -Parent $rPr -Namespace $NS.w -LocalName 'color'
        if (-not $colorEl) { continue }
        if (-not [string]::IsNullOrEmpty((Get-XAttr -Element $colorEl -Namespace $NS.w -LocalName 'themeColor'))) { continue }
        $fgVal = Get-XAttr -Element $colorEl -Namespace $NS.w -LocalName 'val'
        if ([string]::IsNullOrEmpty($fgVal) -or $fgVal -eq 'auto') { continue }
        if ($fgVal -notmatch '^[0-9A-Fa-f]{6}$') { continue }

        $bgVal = Get-ExplicitShadingFill -ShadingParent $rPr
        if (-not $bgVal) {
            $p = $run.Parent
            while ($p -and $p.Name.LocalName -ne 'p') { $p = $p.Parent }
            if ($p) {
                $pPr = Get-ChildElement -Parent $p -Namespace $NS.w -LocalName 'pPr'
                if ($pPr) { $bgVal = Get-ExplicitShadingFill -ShadingParent $pPr }
            }
        }
        if (-not $bgVal) { continue }

        $halfPts = 22
        $sz = Get-ChildElement -Parent $rPr -Namespace $NS.w -LocalName 'sz'
        if ($sz) {
            $szVal = Get-XAttr -Element $sz -Namespace $NS.w -LocalName 'val'
            $parsed = 0
            if ([int]::TryParse($szVal, [ref] $parsed)) { $halfPts = $parsed }
        }
        $isBold = ($null -ne (Get-ChildElement -Parent $rPr -Namespace $NS.w -LocalName 'b'))
        $isLarge = ($halfPts -ge 36) -or ($isBold -and $halfPts -ge 28)
        $threshold = if ($isLarge) { 3.0 } else { 4.5 }

        if ((Get-ContrastRatio -ForegroundHex $fgVal -BackgroundHex $bgVal) -ge $threshold) { continue }

        $newFg = Get-NearestPassingForeground -ForegroundHex $fgVal -BackgroundHex $bgVal -Threshold $threshold
        if ($newFg -and $newFg -ne $fgVal.ToUpper()) {
            $colorEl.SetAttributeValue([Xml.Linq.XName]::Get('val', $NS.w), $newFg)
            $fixed++
        }
    }
    return $fixed
}

function Save-PartXDocument {
    param($Part, [Xml.Linq.XDocument] $Document)
    $stream = $Part.GetStream([IO.FileMode]::Create, [IO.FileAccess]::Write)
    try {
        $Document.Save($stream)
    } finally {
        $stream.Dispose()
    }
}

# --- Open document and run rules ---------------------------------------------
$issues = @()
$doc = $null

try {
    try {
        $doc = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Open($FilePath, $false)
    } catch [System.IO.IOException] {
        [Console]::Error.WriteLine("File is locked or unreadable: $($_.Exception.Message)")
        exit 2
    } catch {
        # IRM/password failures surface as FileFormatException, InvalidDataException,
        # or a generic OpenXmlPackageException wrapping one of those. Sniff the chain.
        $protected = $false
        $cur = $_.Exception
        while ($cur) {
            if ($cur -is [System.IO.FileFormatException] -or $cur -is [System.IO.InvalidDataException]) {
                $protected = $true
                break
            }
            $cur = $cur.InnerException
        }
        if (-not $protected) {
            [Console]::Error.WriteLine("Failed to open document: $($_.Exception.Message)")
            exit 2
        }
        $issues += New-Issue -Severity ERROR -RuleName 'DocumentProtected' `
            -Description 'Document is IRM- or password-protected'
        $doc = $null
    }

    if ($doc) {
        $main = $doc.MainDocumentPart
        if (-not $main) {
            [Console]::Error.WriteLine("Document has no main part.")
            exit 2
        }

        $bodyDoc = Get-PartXDocument $main
        $bodyRoot = if ($bodyDoc) { $bodyDoc.Root } else { $null }

        $headerRoots = @()
        foreach ($hp in $main.HeaderParts) {
            $hd = Get-PartXDocument $hp
            if ($hd -and $hd.Root) { $headerRoots += $hd.Root }
        }

        $footerRoots = @()
        foreach ($fp in $main.FooterParts) {
            $fd = Get-PartXDocument $fp
            if ($fd -and $fd.Root) { $footerRoots += $fd.Root }
        }

        $allRoots = @()
        if ($bodyRoot) { $allRoots += $bodyRoot }
        $allRoots += $headerRoots
        $allRoots += $footerRoots

        $issuesRef = [ref] $issues

        # Alt text + floating + repeated blanks + heading order + content control title
        # apply across body, headers, footers.
        foreach ($root in $allRoots) {
            Test-AltText           -Root $root -Issues $issuesRef
            Test-FloatingObject    -Root $root -Issues $issuesRef
            Test-RepeatedBlank    -Root $root -Issues $issuesRef
            Test-HeadingOrder      -Root $root -Issues $issuesRef
        }

        # Tables and content controls: count indices continuously across roots.
        $tableIdx = 0
        $mergeIdx = 0
        $sdtIdx = 0
        foreach ($root in $allRoots) {
            $tableIdx = Test-TableHeader         -Root $root -StartIndex $tableIdx -Issues $issuesRef
            $mergeIdx = Test-MergedCell          -Root $root -StartIndex $mergeIdx -Issues $issuesRef
            $sdtIdx   = Test-ContentControlTitle  -Root $root -StartIndex $sdtIdx   -Issues $issuesRef
        }

        # TIP rule (always-on; only emitted in detailed mode)
        Test-NoHeadingStyle -Roots $allRoots -Issues $issuesRef

        # WARNING: best-effort contrast check (skips theme/auto colors)
        foreach ($root in $allRoots) {
            Test-LowContrast -Root $root -Issues $issuesRef
        }

        $issues = $issuesRef.Value
    }
}
finally {
    if ($doc) {
        try { $doc.Dispose() } catch { Write-Verbose "Dispose failed: $($_.Exception.Message)" }
    }
}

# --- Decide pass/fail ---------------------------------------------------------
$hasError = $false
foreach ($i in $issues) {
    if ($i.Severity -eq 'ERROR') { $hasError = $true; break }
}

# --- Emit output --------------------------------------------------------------
if ($Format -eq 'detailed') {
    $severityRank = @{ 'ERROR' = 0; 'WARNING' = 1; 'TIP' = 2 }
    $emit = $issues | Sort-Object `
        @{ Expression = { $severityRank[$_.Severity] } }, `
        @{ Expression = { $_.RuleName } }
    foreach ($i in $emit) {
        # In detailed mode every issue prints; TIPs only appear in detailed mode
        # (which is what we're in), and ERROR/WARNING always print.
        Write-Output ("{0}`t{1}`t{2}" -f $i.Severity, $i.RuleName, $i.Description)
    }
}

if ($hasError) {
    $errorRules = ($issues | Where-Object { $_.Severity -eq 'ERROR' } | Select-Object -ExpandProperty RuleName -Unique) -join ', '
    Write-Output "FAIL $FilePath`: $errorRules"
    $originalExit = 1
} else {
    Write-Output "PASS $FilePath"
    $originalExit = 0
}

# --- Autofix ------------------------------------------------------------------
# Applies only the deterministic structural remediations; everything else
# requires authoring or layout judgment and is intentionally skipped. The
# original is never touched: we copy to <basename>.fixed.docx and edit the copy.
if (-not $Fix) { exit $originalExit }

$fixableRuleNames = @('MissingTableHeaders', 'RepeatedBlanks', 'LowContrast')
$hasFixable = $false
foreach ($i in $issues) {
    if ($fixableRuleNames -contains $i.RuleName) { $hasFixable = $true; break }
}
if (-not $hasFixable) { exit $originalExit }

$fixedPath = [IO.Path]::ChangeExtension($FilePath, $null).TrimEnd('.') + '.fixed' + $ext
try {
    Copy-Item -LiteralPath $FilePath -Destination $fixedPath -Force
} catch {
    [Console]::Error.WriteLine("Failed to create fixed copy at '$fixedPath': $($_.Exception.Message)")
    exit 2
}

$fixCounts = @{ MissingTableHeaders = 0; RepeatedBlanks = 0; LowContrast = 0 }
$fixDoc = $null
try {
    try {
        $fixDoc = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Open($fixedPath, $true)
    } catch {
        [Console]::Error.WriteLine("Failed to open fixed copy '$fixedPath' for editing: $($_.Exception.Message)")
        exit 2
    }

    $main = $fixDoc.MainDocumentPart
    $partsToFix = @()
    if ($main) { $partsToFix += $main }
    foreach ($hp in $main.HeaderParts) { $partsToFix += $hp }
    foreach ($fp in $main.FooterParts) { $partsToFix += $fp }

    foreach ($part in $partsToFix) {
        $xdoc = Get-PartXDocument $part
        if (-not $xdoc -or -not $xdoc.Root) { continue }

        $fixCounts.MissingTableHeaders += (Repair-WordTableHeader  -Root $xdoc.Root)
        $fixCounts.RepeatedBlanks      += (Repair-WordRepeatedBlank -Root $xdoc.Root)
        $fixCounts.LowContrast         += (Repair-WordLowContrast   -Root $xdoc.Root)

        Save-PartXDocument -Part $part -Document $xdoc
    }
} finally {
    if ($fixDoc) {
        try { $fixDoc.Dispose() } catch { Write-Verbose "Dispose failed: $($_.Exception.Message)" }
    }
}

$fixSummaryParts = @()
foreach ($k in $fixableRuleNames) {
    if ($fixCounts[$k] -gt 0) { $fixSummaryParts += ('{0} ({1})' -f $k, $fixCounts[$k]) }
}
if ($fixSummaryParts.Count -eq 0) {
    # Issues were reported but the autofix paths matched none of the actual
    # OOXML state (e.g., theme-driven contrast, cross-w:t blank runs). Drop
    # the empty .fixed file rather than leaving a misleading artifact.
    Remove-Item -LiteralPath $fixedPath -Force -ErrorAction SilentlyContinue
    exit $originalExit
}

Write-Output ("FIXED {0}: {1}" -f $fixedPath, ($fixSummaryParts -join ', '))

# Re-run on the fixed file so the user sees its PASS/FAIL (and any residual
# issues in detailed mode). Recursion is bounded: -Fix is intentionally not
# forwarded, so this terminates.
& $PSCommandPath -FilePath $fixedPath -Format $Format
exit $LASTEXITCODE
