<#
.SYNOPSIS
    Runs Microsoft Office Accessibility Checker rules against a .xlsx/.xlsm
    file using the Open XML SDK.

.DESCRIPTION
    Implements the Excel rule subset from Microsoft's published Accessibility
    Checker rules by traversing the underlying OOXML directly with the
    Open XML SDK. No Office installation is required.

    Rule coverage:
      ERROR:   MissingAltText, MissingTableHeaders, DocumentProtected,
               RedOnlyNegativeFormatting
      WARNING: MergedCells, DefaultSheetTabName, LowContrast
      TIP:     DefaultTableName

    Notes:
      - RedOnlyNegativeFormatting: best-effort detection of "[Red]" numFmt
        codes that lack a complementary minus/parens/secondary color.
      - LowContrast: best-effort. Only flags cells whose font color and
        cell fill are both explicit RGB values; theme references, indexed
        palette colors, "auto", non-solid fills, and conditional formatting
        are skipped.

.PARAMETER FilePath
    Path to a .xlsx or .xlsm file.

.PARAMETER Format
    'text' (default) emits a single PASS/FAIL line.
    'detailed' emits every issue (one tab-separated line) followed by the
    PASS/FAIL summary line.

.PARAMETER Fix
    When set, after running the checks the script writes a sibling file named
    <basename>.fixed.xlsx (or .xlsm) with deterministic structural remediations
    applied for the rules MissingTableHeaders, RedOnlyNegativeFormatting, and
    LowContrast. The original file is never modified.

.OUTPUTS
    Exit codes:
      0  no errors found (warnings/tips do not fail)
      1  one or more accessibility errors found (includes IRM/password-protected)
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

# Localized default sheet-tab name patterns. Adding a locale is a one-line edit.
$DefaultSheetNamePatterns = @(
    '^(Sheet|Chart)\d+$',     # English
    '^Tabelle\d+$',            # German
    '^(Feuil|Feuille)\d+$',    # French
    '^Hoja\d+$',               # Spanish
    '^Foglio\d+$',             # Italian
    '^Planilha\d+$',           # Portuguese
    '^シート\d+$'              # Japanese
)

# OOXML namespaces
$NS = @{
    s   = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    xdr = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    r   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

#--------------------------------------------------------------------
# Pre-flight: file existence + extension
#--------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $FilePath)) {
    [Console]::Error.WriteLine("File not found: $FilePath")
    exit 2
}

$ext = [IO.Path]::GetExtension($FilePath).ToLowerInvariant()
if ($ext -ne '.xlsx' -and $ext -ne '.xlsm') {
    [Console]::Error.WriteLine("Unsupported file type for Excel checker: $ext (expected .xlsx or .xlsm)")
    exit 2
}

# Resolve to absolute path so SDK doesn't rely on cwd.
$FilePath = (Resolve-Path -LiteralPath $FilePath).ProviderPath

#--------------------------------------------------------------------
# Locate and load Open XML SDK
#--------------------------------------------------------------------

$sdkCandidates = @()
if ($env:OPENXML_SDK_PATH) {
    $sdkCandidates += $env:OPENXML_SDK_PATH
}
$sdkCandidates += (Join-Path $PSScriptRoot 'lib\DocumentFormat.OpenXml.dll')

$sdkPath = $null
foreach ($candidate in $sdkCandidates) {
    if ($candidate -and (Test-Path -LiteralPath $candidate)) {
        $sdkPath = $candidate
        break
    }
}

if (-not $sdkPath) {
    [Console]::Error.WriteLine("Open XML SDK not found. Run scripts/setup-accessibility-checker.ps1 first.")
    exit 2
}

try {
    Add-Type -Path $sdkPath
}
catch {
    [Console]::Error.WriteLine("Failed to load Open XML SDK from '$sdkPath': $($_.Exception.Message)")
    exit 2
}

#--------------------------------------------------------------------
# Helpers
#--------------------------------------------------------------------

function New-Issue {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'New-Issue is a pure record factory — it constructs a [pscustomobject] and has no side effects. The verb New triggers this rule, but the function does not change any system state.'
    )]
    param(
        [Parameter(Mandatory)][string] $Severity,
        [Parameter(Mandatory)][string] $RuleName,
        [Parameter(Mandatory)][string] $Description
    )
    [pscustomobject]@{
        Severity    = $Severity
        RuleName    = $RuleName
        Description = $Description
    }
}

function Get-PartXDocument {
    param($Part)
    if ($null -eq $Part) { return $null }
    try {
        $stream = $Part.GetStream([IO.FileMode]::Open, [IO.FileAccess]::Read)
        try {
            return [System.Xml.Linq.XDocument]::Load($stream)
        }
        finally {
            $stream.Dispose()
        }
    }
    catch {
        return $null
    }
}

function Get-XAttr {
    param([System.Xml.Linq.XElement] $Element, [string] $Namespace, [string] $LocalName)
    if ($null -eq $Element) { return $null }
    $attr = $Element.Attribute([System.Xml.Linq.XName]::Get($LocalName, $Namespace))
    if ($attr) { return $attr.Value }
    # Fall back to unqualified attribute (most xlsx attributes have no namespace).
    $attr = $Element.Attribute([System.Xml.Linq.XName]::Get($LocalName))
    if ($attr) { return $attr.Value }
    return $null
}

function Get-ChildElement {
    param(
        [System.Xml.Linq.XElement] $Parent,
        [string] $Namespace,
        [string] $LocalName,
        [switch] $All
    )
    if ($All) {
        if ($null -eq $Parent) { return @() }
        return ,@($Parent.Elements([System.Xml.Linq.XName]::Get($LocalName, $Namespace)))
    }
    if ($null -eq $Parent) { return $null }
    return $Parent.Element([System.Xml.Linq.XName]::Get($LocalName, $Namespace))
}

#--------------------------------------------------------------------
# Sheet-name lookup: relate WorksheetPart -> sheet display name
#--------------------------------------------------------------------

function Get-SheetNameMap {
    param($WorkbookPart)
    $map = @{}
    $wbDoc = Get-PartXDocument -Part $WorkbookPart
    if ($null -eq $wbDoc) { return $map }
    $sheetsEl = Get-ChildElement -Parent $wbDoc.Root -Namespace $NS.s -LocalName 'sheets'
    if ($null -eq $sheetsEl) { return $map }
    foreach ($sheet in (Get-ChildElement -Parent $sheetsEl -Namespace $NS.s -LocalName 'sheet' -All)) {
        $name = Get-XAttr -Element $sheet -Namespace $null -LocalName 'name'
        $rid  = Get-XAttr -Element $sheet -Namespace $NS.r -LocalName 'id'
        if ($rid) { $map[$rid] = $name }
    }
    return $map
}

function Get-WorksheetDisplayName {
    param($WorkbookPart, $WorksheetPart, $SheetMap)
    try {
        $rid = $WorkbookPart.GetIdOfPart($WorksheetPart)
        if ($SheetMap.ContainsKey($rid)) { return $SheetMap[$rid] }
    } catch {
        Write-Verbose "GetIdOfPart failed: $($_.Exception.Message)"
    }
    return '(unknown)'
}

#--------------------------------------------------------------------
# Rule: MissingAltText
#--------------------------------------------------------------------

# Map shape local-name -> non-visual-properties container local-name
$XdrNvContainerByShape = @{
    'pic'          = 'nvPicPr'
    'graphicFrame' = 'nvGraphicFramePr'
    'sp'           = 'nvSpPr'
    'grpSp'        = 'nvGrpSpPr'
}

# Pass if any of: @descr non-empty, @title non-empty, @decorative in {1,true}.
function Test-XdrAltTextAttribute {
    param([System.Xml.Linq.XElement] $CNvPr)
    if ($null -eq $CNvPr) { return $false }
    if (-not [string]::IsNullOrWhiteSpace((Get-XAttr -Element $CNvPr -Namespace $null -LocalName 'descr'))) { return $true }
    if (-not [string]::IsNullOrWhiteSpace((Get-XAttr -Element $CNvPr -Namespace $null -LocalName 'title'))) { return $true }
    $decorative = Get-XAttr -Element $CNvPr -Namespace $null -LocalName 'decorative'
    return ($decorative -eq '1' -or $decorative -eq 'true')
}

function Test-XdrShapeContainer {
    param([System.Xml.Linq.XElement] $Container, [string] $SheetName)
    $issues = @()
    foreach ($shapeLocalName in $XdrNvContainerByShape.Keys) {
        $nvContainerName = $XdrNvContainerByShape[$shapeLocalName]
        foreach ($shape in (Get-ChildElement -Parent $Container -Namespace $NS.xdr -LocalName $shapeLocalName -All)) {
            $nv = Get-ChildElement -Parent $shape -Namespace $NS.xdr -LocalName $nvContainerName
            $cNvPr = Get-ChildElement -Parent $nv -Namespace $NS.xdr -LocalName 'cNvPr'
            if ($cNvPr -and -not (Test-XdrAltTextAttribute $cNvPr)) {
                $name = Get-XAttr -Element $cNvPr -Namespace $null -LocalName 'name'
                if ([string]::IsNullOrWhiteSpace($name)) { $name = '(unnamed)' }
                $issues += New-Issue -Severity 'ERROR' -RuleName 'MissingAltText' `
                    -Description "Image/object `"$name`" on sheet `"$SheetName`" has no alt text"
            }
            if ($shapeLocalName -eq 'grpSp') {
                $issues += Test-XdrShapeContainer -Container $shape -SheetName $SheetName
            }
        }
    }
    return $issues
}

function Test-AltText {
    param($DrawingsPart, [string] $SheetName)
    $issues = @()
    $doc = Get-PartXDocument -Part $DrawingsPart
    if ($null -eq $doc -or $null -eq $doc.Root) { return $issues }
    foreach ($anchorType in @('twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor')) {
        foreach ($anchor in (Get-ChildElement -Parent $doc.Root -Namespace $NS.xdr -LocalName $anchorType -All)) {
            $issues += Test-XdrShapeContainer -Container $anchor -SheetName $SheetName
        }
    }
    return $issues
}

#--------------------------------------------------------------------
# Rule: MissingTableHeaders + DefaultTableName
#--------------------------------------------------------------------

function Test-TableDefinition {
    param($TableDefinitionPart)

    $issues = @()
    $doc = Get-PartXDocument -Part $TableDefinitionPart
    if ($null -eq $doc -or $null -eq $doc.Root) { return $issues }

    $tableEl = $doc.Root
    $name        = Get-XAttr -Element $tableEl -Namespace $null -LocalName 'name'
    $displayName = Get-XAttr -Element $tableEl -Namespace $null -LocalName 'displayName'
    $effectiveName = if ([string]::IsNullOrWhiteSpace($name)) { $displayName } else { $name }
    if ([string]::IsNullOrWhiteSpace($effectiveName)) { $effectiveName = '(unnamed)' }

    # MissingTableHeaders: headerRowCount="0" => fail
    $hrc = Get-XAttr -Element $tableEl -Namespace $null -LocalName 'headerRowCount'
    if ($null -ne $hrc -and $hrc -eq '0') {
        $issues += (New-Issue -Severity 'ERROR' -RuleName 'MissingTableHeaders' `
            -Description "Table `"$effectiveName`" has no header row")
    }

    # DefaultTableName tip
    if ($effectiveName -match '^Table\d+$') {
        $issues += (New-Issue -Severity 'TIP' -RuleName 'DefaultTableName' `
            -Description "Table `"$effectiveName`" uses a default placeholder name")
    }

    return $issues
}

#--------------------------------------------------------------------
# Rule: MergedCells
#--------------------------------------------------------------------

function Test-MergedCell {
    param($WorksheetPart, [string] $SheetName)

    $issues = @()
    $doc = Get-PartXDocument -Part $WorksheetPart
    if ($null -eq $doc -or $null -eq $doc.Root) { return $issues }

    $mergeCellsEl = Get-ChildElement -Parent $doc.Root -Namespace $NS.s -LocalName 'mergeCells'
    if ($null -eq $mergeCellsEl) { return $issues }

    $mergeCells = Get-ChildElement -Parent $mergeCellsEl -Namespace $NS.s -LocalName 'mergeCell' -All
    if ($mergeCells.Count -gt 0) {
        $issues += (New-Issue -Severity 'WARNING' -RuleName 'MergedCells' `
            -Description "Sheet `"$SheetName`" contains merged cells")
    }
    return $issues
}

#--------------------------------------------------------------------
# Rule: DefaultSheetTabName
#--------------------------------------------------------------------

function Test-SheetTabName {
    param($WorkbookPart)

    $issues = @()
    $doc = Get-PartXDocument -Part $WorkbookPart
    if ($null -eq $doc -or $null -eq $doc.Root) { return $issues }

    $sheetsEl = Get-ChildElement -Parent $doc.Root -Namespace $NS.s -LocalName 'sheets'
    if ($null -eq $sheetsEl) { return $issues }

    foreach ($sheet in (Get-ChildElement -Parent $sheetsEl -Namespace $NS.s -LocalName 'sheet' -All)) {
        $name = Get-XAttr -Element $sheet -Namespace $null -LocalName 'name'

        if ([string]::IsNullOrWhiteSpace($name)) {
            $issues += (New-Issue -Severity 'WARNING' -RuleName 'DefaultSheetTabName' `
                -Description "Sheet tab `"$name`" uses a default placeholder name")
            continue
        }

        $matched = $false
        foreach ($pattern in $DefaultSheetNamePatterns) {
            if ($name -match $pattern) { $matched = $true; break }
        }
        if ($matched) {
            $issues += (New-Issue -Severity 'WARNING' -RuleName 'DefaultSheetTabName' `
                -Description "Sheet tab `"$name`" uses a default placeholder name")
        }
    }

    return $issues
}

#--------------------------------------------------------------------
# Rule: RedOnlyNegativeFormatting
#--------------------------------------------------------------------

# Collect the set of numFmtId values that are actually referenced by a cell on
# any worksheet. Reporting unused custom <numFmt> entries as ERROR makes the
# CLI fail accessible workbooks because of stale style metadata that no cell
# uses. We resolve cell `s` -> cellXfs index -> numFmtId and return the set.
function Get-UsedNumFmtId {
    param($WorkbookPart)

    $used = New-Object 'System.Collections.Generic.HashSet[string]'

    if ($null -eq $WorkbookPart -or $null -eq $WorkbookPart.WorkbookStylesPart) {
        return $used
    }

    $stylesDoc = Get-PartXDocument -Part $WorkbookPart.WorkbookStylesPart
    if ($null -eq $stylesDoc -or $null -eq $stylesDoc.Root) { return $used }

    # Build cellXf index -> numFmtId map.
    $cellXfsEl = Get-ChildElement -Parent $stylesDoc.Root -Namespace $NS.s -LocalName 'cellXfs'
    if ($null -eq $cellXfsEl) { return $used }

    $cellXfNumFmtIds = New-Object 'System.Collections.Generic.List[string]'
    foreach ($xf in (Get-ChildElement -Parent $cellXfsEl -Namespace $NS.s -LocalName 'xf' -All)) {
        $id = Get-XAttr -Element $xf -Namespace $null -LocalName 'numFmtId'
        if ([string]::IsNullOrEmpty($id)) { $id = '0' }
        $cellXfNumFmtIds.Add($id) | Out-Null
    }

    # Walk every worksheet's cells and record the numFmtId of each used style.
    $cName = [System.Xml.Linq.XName]::Get('c', $NS.s)
    foreach ($wsPart in $WorkbookPart.WorksheetParts) {
        $wsDoc = Get-PartXDocument -Part $wsPart
        if ($null -eq $wsDoc -or $null -eq $wsDoc.Root) { continue }

        foreach ($cell in $wsDoc.Root.Descendants($cName)) {
            $sAttr = $cell.Attribute([System.Xml.Linq.XName]::Get('s'))
            $sIdx = if ($sAttr) { $sAttr.Value } else { '0' }
            $parsed = 0
            if ([int]::TryParse($sIdx, [ref] $parsed) -and $parsed -ge 0 -and $parsed -lt $cellXfNumFmtIds.Count) {
                [void] $used.Add($cellXfNumFmtIds[$parsed])
            }
        }
    }

    return $used
}

function Test-RedOnlyNumberFormat {
    param(
        $WorkbookStylesPart,
        [System.Collections.Generic.HashSet[string]] $UsedNumFmtIds
    )

    $issues = @()
    if ($null -eq $WorkbookStylesPart) { return $issues }

    $doc = Get-PartXDocument -Part $WorkbookStylesPart
    if ($null -eq $doc -or $null -eq $doc.Root) { return $issues }

    $numFmtsEl = Get-ChildElement -Parent $doc.Root -Namespace $NS.s -LocalName 'numFmts'
    if ($null -eq $numFmtsEl) { return $issues }

    foreach ($numFmt in (Get-ChildElement -Parent $numFmtsEl -Namespace $NS.s -LocalName 'numFmt' -All)) {
        $numFmtId = Get-XAttr -Element $numFmt -Namespace $null -LocalName 'numFmtId'
        if ($UsedNumFmtIds -and -not [string]::IsNullOrEmpty($numFmtId) -and -not $UsedNumFmtIds.Contains($numFmtId)) {
            continue
        }
        $code = Get-XAttr -Element $numFmt -Namespace $null -LocalName 'formatCode'
        if ([string]::IsNullOrEmpty($code)) { continue }

        # Excel format sections: positive ; negative ; zero ; text
        # Section delimiter is ';' (not escaped in standard usage).
        $sections = $code -split ';'
        if ($sections.Count -lt 2) { continue }

        $negative = $sections[1]
        if ([string]::IsNullOrEmpty($negative)) { continue }

        # Check if [Red] appears in the negative section
        if ($negative -notmatch '(?i)\[Red\]') { continue }

        # Strip [Red] then check for any other color marker, parentheses, or leading minus
        $stripped = $negative -replace '(?i)\[Red\]', ''

        # Other color markers: any [Black], [Blue], [Cyan], [Green], [Magenta], [White], [Yellow], [Color N]
        $hasOtherColor = $stripped -match '(?i)\[(Black|Blue|Cyan|Green|Magenta|White|Yellow|Color\s*\d+)\]'

        # Parentheses around the negative format
        $hasParens = $stripped.Contains('(') -or $stripped.Contains(')')

        # Leading minus sign (allow leading whitespace) — strip remaining bracketed conditionals
        # (e.g. "[<0]") so we look at the actual format text.
        $strippedNoBrackets = $stripped -replace '\[[^\]]*\]', ''
        $hasMinus = $strippedNoBrackets.TrimStart() -match '^\s*-'

        if (-not $hasOtherColor -and -not $hasParens -and -not $hasMinus) {
            $issues += (New-Issue -Severity 'ERROR' -RuleName 'RedOnlyNegativeFormatting' `
                -Description "Number format `"$code`" uses red-only differentiation for negative values")
        }
    }

    return $issues
}

#--------------------------------------------------------------------
# Rule: LowContrast (WARNING) -- best-effort
#--------------------------------------------------------------------
#
# Walks every text-bearing cell on every worksheet, resolves its effective
# font color and fill color via the cellXf -> font/fill chain in styles.xml,
# and flags pairs whose WCAG 2.x contrast ratio falls below 4.5:1.
#
# Skipped (cannot be evaluated without rendering): "auto" colors, theme
# references, indexed-palette references, gradient/non-solid fills, and any
# cell whose font or fill resolves to a non-explicit value. Conditional
# formatting rules are also out of scope (they depend on cell values).
#
# Note: font-size-driven large-text relaxation (3.0:1 vs 4.5:1) is omitted
# because Excel cells inherit font size in ways the static check cannot
# resolve reliably; we use the stricter 4.5:1 universally.

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

# Excel <color> elements may carry rgb=, theme=, indexed=, or auto=. Only
# explicit rgb values can be evaluated; everything else returns $null.
# rgb is stored as 8-hex ARGB; we strip the alpha channel.
function Get-ExcelExplicitColorHex {
    param([System.Xml.Linq.XElement] $ColorEl)
    if ($null -eq $ColorEl) { return $null }
    if (-not [string]::IsNullOrEmpty((Get-XAttr -Element $ColorEl -Namespace $null -LocalName 'theme'))) { return $null }
    if (-not [string]::IsNullOrEmpty((Get-XAttr -Element $ColorEl -Namespace $null -LocalName 'indexed'))) { return $null }
    $auto = Get-XAttr -Element $ColorEl -Namespace $null -LocalName 'auto'
    if ($auto -eq '1' -or $auto -eq 'true') { return $null }
    $rgb = Get-XAttr -Element $ColorEl -Namespace $null -LocalName 'rgb'
    if ([string]::IsNullOrEmpty($rgb)) { return $null }
    if ($rgb -match '^[0-9A-Fa-f]{8}$') { return $rgb.Substring(2) }
    if ($rgb -match '^[0-9A-Fa-f]{6}$') { return $rgb }
    return $null
}

function Test-LowContrast {
    param($WorkbookPart, $SheetMap)
    $issues = @()
    if ($null -eq $WorkbookPart -or $null -eq $WorkbookPart.WorkbookStylesPart) { return $issues }

    $stylesDoc = Get-PartXDocument -Part $WorkbookPart.WorkbookStylesPart
    if ($null -eq $stylesDoc -or $null -eq $stylesDoc.Root) { return $issues }

    # fonts[i] -> hex (or $null if not explicit)
    $fontColors = New-Object 'System.Collections.Generic.List[object]'
    $fontsEl = Get-ChildElement -Parent $stylesDoc.Root -Namespace $NS.s -LocalName 'fonts'
    if ($fontsEl) {
        foreach ($font in (Get-ChildElement -Parent $fontsEl -Namespace $NS.s -LocalName 'font' -All)) {
            $colorEl = Get-ChildElement -Parent $font -Namespace $NS.s -LocalName 'color'
            $fontColors.Add((Get-ExcelExplicitColorHex $colorEl)) | Out-Null
        }
    }

    # fills[i] -> hex (only solid patternFill with explicit fgColor); else $null
    $fillColors = New-Object 'System.Collections.Generic.List[object]'
    $fillsEl = Get-ChildElement -Parent $stylesDoc.Root -Namespace $NS.s -LocalName 'fills'
    if ($fillsEl) {
        foreach ($fill in (Get-ChildElement -Parent $fillsEl -Namespace $NS.s -LocalName 'fill' -All)) {
            $hex = $null
            $pf = Get-ChildElement -Parent $fill -Namespace $NS.s -LocalName 'patternFill'
            if ($pf) {
                $type = Get-XAttr -Element $pf -Namespace $null -LocalName 'patternType'
                if ($type -eq 'solid') {
                    $fgColorEl = Get-ChildElement -Parent $pf -Namespace $NS.s -LocalName 'fgColor'
                    $hex = Get-ExcelExplicitColorHex $fgColorEl
                }
            }
            $fillColors.Add($hex) | Out-Null
        }
    }

    # cellXfs[i] -> (FontId, FillId)
    $cellXfs = New-Object 'System.Collections.Generic.List[object]'
    $cellXfsEl = Get-ChildElement -Parent $stylesDoc.Root -Namespace $NS.s -LocalName 'cellXfs'
    if ($cellXfsEl) {
        foreach ($xf in (Get-ChildElement -Parent $cellXfsEl -Namespace $NS.s -LocalName 'xf' -All)) {
            $fontId = 0; [int]::TryParse((Get-XAttr -Element $xf -Namespace $null -LocalName 'fontId'), [ref] $fontId) | Out-Null
            $fillId = 0; [int]::TryParse((Get-XAttr -Element $xf -Namespace $null -LocalName 'fillId'), [ref] $fillId) | Out-Null
            $cellXfs.Add([pscustomobject]@{ FontId = $fontId; FillId = $fillId }) | Out-Null
        }
    }

    if ($cellXfs.Count -eq 0) { return $issues }

    $cName = [System.Xml.Linq.XName]::Get('c', $NS.s)
    foreach ($wsPart in $WorkbookPart.WorksheetParts) {
        $wsDoc = Get-PartXDocument -Part $wsPart
        if ($null -eq $wsDoc -or $null -eq $wsDoc.Root) { continue }
        $sheetName = Get-WorksheetDisplayName -WorkbookPart $WorkbookPart -WorksheetPart $wsPart -SheetMap $SheetMap

        foreach ($cell in $wsDoc.Root.Descendants($cName)) {
            $sAttr = $cell.Attribute([System.Xml.Linq.XName]::Get('s'))
            if ($null -eq $sAttr) { continue }

            # Skip cells with no displayable text. Empty cells render nothing,
            # so contrast is meaningless and would produce noise on large sheets
            # with widespread fill formatting.
            $hasContent = $false
            foreach ($child in $cell.Elements()) {
                $ln = $child.Name.LocalName
                if ($ln -eq 'v' -or $ln -eq 'f' -or $ln -eq 'is') { $hasContent = $true; break }
            }
            if (-not $hasContent) { continue }

            $sIdx = -1
            if (-not [int]::TryParse($sAttr.Value, [ref] $sIdx)) { continue }
            if ($sIdx -lt 0 -or $sIdx -ge $cellXfs.Count) { continue }
            $xf = $cellXfs[$sIdx]
            if ($xf.FontId -lt 0 -or $xf.FontId -ge $fontColors.Count) { continue }
            if ($xf.FillId -lt 0 -or $xf.FillId -ge $fillColors.Count) { continue }
            $fg = $fontColors[$xf.FontId]
            $bg = $fillColors[$xf.FillId]
            if ([string]::IsNullOrEmpty($fg) -or [string]::IsNullOrEmpty($bg)) { continue }

            $cellRef = Get-XAttr -Element $cell -Namespace $null -LocalName 'r'
            if ([string]::IsNullOrEmpty($cellRef)) { $cellRef = '(unknown)' }
            $ratio = Get-ContrastRatio -ForegroundHex $fg -BackgroundHex $bg
            if ($ratio -lt 4.5) {
                $issues += New-Issue -Severity 'WARNING' -RuleName 'LowContrast' `
                    -Description ("Cell {0} on sheet `"{1}`" has contrast {2:N2}:1 (#{3} on #{4}); WCAG requires 4.5:1" -f $cellRef, $sheetName, $ratio, $fg.ToUpper(), $bg.ToUpper())
            }
        }
    }

    return $issues
}

#--------------------------------------------------------------------
# Autofix helpers
#--------------------------------------------------------------------
# See check-docx-accessibility.ps1 for the full HSL/contrast notes — the
# helpers are duplicated rather than shared because each checker is meant
# to be runnable on its own.

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
    $blackR = Get-ContrastRatio -ForegroundHex '000000' -BackgroundHex $BackgroundHex
    $whiteR = Get-ContrastRatio -ForegroundHex 'FFFFFF' -BackgroundHex $BackgroundHex
    if ($blackR -ge $whiteR) { return '000000' } else { return 'FFFFFF' }
}

# Removes headerRowCount="0" from a single TableDefinitionPart's table element.
# Excel defaults headerRowCount to 1 when the attribute is absent, so deleting
# it is the correct way to re-enable a header row. Returns 1 if it changed
# anything, else 0.
function Repair-XlsxTableHeader {
    param($TableDefinitionPart)
    $doc = Get-PartXDocument -Part $TableDefinitionPart
    if ($null -eq $doc -or $null -eq $doc.Root) { return 0 }
    $tableEl = $doc.Root
    $hrcAttr = $tableEl.Attribute([System.Xml.Linq.XName]::Get('headerRowCount'))
    if ($null -eq $hrcAttr -or $hrcAttr.Value -ne '0') { return 0 }
    $hrcAttr.Remove()
    Save-PartXDocument -Part $TableDefinitionPart -Document $doc
    return 1
}

# Rewrite [Red]-only negative format codes to add a leading minus to the
# negative section. We deliberately do not strip [Red] (the user picked that
# styling); we just augment the format with a sign so the value is also
# distinguishable for color-blind users. Returns the count of numFmt entries
# rewritten in this stylesheet.
function Repair-XlsxRedOnlyNumberFormat {
    param($WorkbookStylesPart,
          [System.Collections.Generic.HashSet[string]] $UsedNumFmtIds)
    if ($null -eq $WorkbookStylesPart) { return 0 }
    $doc = Get-PartXDocument -Part $WorkbookStylesPart
    if ($null -eq $doc -or $null -eq $doc.Root) { return 0 }
    $numFmtsEl = Get-ChildElement -Parent $doc.Root -Namespace $NS.s -LocalName 'numFmts'
    if ($null -eq $numFmtsEl) { return 0 }

    $fixed = 0
    foreach ($numFmt in (Get-ChildElement -Parent $numFmtsEl -Namespace $NS.s -LocalName 'numFmt' -All)) {
        $numFmtId = Get-XAttr -Element $numFmt -Namespace $null -LocalName 'numFmtId'
        if ($UsedNumFmtIds -and -not [string]::IsNullOrEmpty($numFmtId) -and -not $UsedNumFmtIds.Contains($numFmtId)) { continue }
        $code = Get-XAttr -Element $numFmt -Namespace $null -LocalName 'formatCode'
        if ([string]::IsNullOrEmpty($code)) { continue }
        $sections = $code -split ';'
        if ($sections.Count -lt 2) { continue }
        $negative = $sections[1]
        if ([string]::IsNullOrEmpty($negative)) { continue }
        if ($negative -notmatch '(?i)\[Red\]') { continue }
        $stripped = $negative -replace '(?i)\[Red\]', ''
        if ($stripped -match '(?i)\[(Black|Blue|Cyan|Green|Magenta|White|Yellow|Color\s*\d+)\]') { continue }
        if ($stripped.Contains('(') -or $stripped.Contains(')')) { continue }
        $strippedNoBrackets = $stripped -replace '\[[^\]]*\]', ''
        if ($strippedNoBrackets.TrimStart() -match '^\s*-') { continue }

        # Insert a literal '-' immediately after the [Red] marker (or any
        # other leading bracket-conditional). regex captures [Red] at start
        # of the section, and any preceding [<...>] conditionals.
        $newNegative = $negative -replace '(?i)^(\s*(?:\[[^\]]*\]\s*)*)\[Red\]', '${1}[Red]-'
        if ($newNegative -eq $negative) { continue }
        $sections[1] = $newNegative
        $newCode = $sections -join ';'
        $numFmt.SetAttributeValue([System.Xml.Linq.XName]::Get('formatCode'), $newCode)
        $fixed++
    }

    if ($fixed -gt 0) { Save-PartXDocument -Part $WorkbookStylesPart -Document $doc }
    return $fixed
}

# Walk every text cell, resolve the cellXf -> font/fill chain, and for each
# cell whose contrast falls below 4.5:1 create (or reuse) a font with the
# nearest passing color and rebind the cellXf to it. Cells whose color comes
# from theme/indexed/auto, or whose fill is non-solid, are skipped — same
# bar the rule uses.
function Repair-XlsxLowContrast {
    param($WorkbookPart)
    if ($null -eq $WorkbookPart -or $null -eq $WorkbookPart.WorkbookStylesPart) { return 0 }
    $stylesPart = $WorkbookPart.WorkbookStylesPart
    $stylesDoc  = Get-PartXDocument -Part $stylesPart
    if ($null -eq $stylesDoc -or $null -eq $stylesDoc.Root) { return 0 }

    $fontsEl   = Get-ChildElement -Parent $stylesDoc.Root -Namespace $NS.s -LocalName 'fonts'
    $fillsEl   = Get-ChildElement -Parent $stylesDoc.Root -Namespace $NS.s -LocalName 'fills'
    $cellXfsEl = Get-ChildElement -Parent $stylesDoc.Root -Namespace $NS.s -LocalName 'cellXfs'
    if ($null -eq $fontsEl -or $null -eq $cellXfsEl) { return 0 }

    # Get-ChildElement -All returns its array via the unary-comma trick so that
    # a single-result call still iterates correctly. Wrapping with @() here would
    # double-wrap (count collapses to 1), so we assign the bare result.
    $fontEls = Get-ChildElement -Parent $fontsEl -Namespace $NS.s -LocalName 'font' -All
    $fillEls = if ($fillsEl) { Get-ChildElement -Parent $fillsEl -Namespace $NS.s -LocalName 'fill' -All } else { @() }
    $xfEls   = Get-ChildElement -Parent $cellXfsEl -Namespace $NS.s -LocalName 'xf' -All

    # Resolve fill colors once.
    $fillHex = New-Object 'System.Collections.Generic.List[object]'
    foreach ($fill in $fillEls) {
        $hex = $null
        $pf = Get-ChildElement -Parent $fill -Namespace $NS.s -LocalName 'patternFill'
        if ($pf -and (Get-XAttr -Element $pf -Namespace $null -LocalName 'patternType') -eq 'solid') {
            $fgColorEl = Get-ChildElement -Parent $pf -Namespace $NS.s -LocalName 'fgColor'
            $hex = Get-ExcelExplicitColorHex $fgColorEl
        }
        $fillHex.Add($hex) | Out-Null
    }

    # Track desired color overrides: original FontId -> (hex(BG) -> new FontId).
    # Each unique (font, bg) pair gets one synthesised font; cellXfs that share
    # the pair share the new font.
    $newFontByPair = @{}
    $fixed = 0

    $cName = [System.Xml.Linq.XName]::Get('c', $NS.s)
    foreach ($wsPart in $WorkbookPart.WorksheetParts) {
        $wsDoc = Get-PartXDocument -Part $wsPart
        if ($null -eq $wsDoc -or $null -eq $wsDoc.Root) { continue }

        $wsChanged = $false
        foreach ($cell in $wsDoc.Root.Descendants($cName)) {
            $sAttr = $cell.Attribute([System.Xml.Linq.XName]::Get('s'))
            if ($null -eq $sAttr) { continue }
            $hasContent = $false
            foreach ($child in $cell.Elements()) {
                $ln = $child.Name.LocalName
                if ($ln -eq 'v' -or $ln -eq 'f' -or $ln -eq 'is') { $hasContent = $true; break }
            }
            if (-not $hasContent) { continue }
            $sIdx = -1
            if (-not [int]::TryParse($sAttr.Value, [ref] $sIdx)) { continue }
            if ($sIdx -lt 0 -or $sIdx -ge $xfEls.Count) { continue }

            $xf = $xfEls[$sIdx]
            $fontId = 0; [int]::TryParse((Get-XAttr -Element $xf -Namespace $null -LocalName 'fontId'), [ref] $fontId) | Out-Null
            $fillId = 0; [int]::TryParse((Get-XAttr -Element $xf -Namespace $null -LocalName 'fillId'), [ref] $fillId) | Out-Null
            if ($fontId -lt 0 -or $fontId -ge $fontEls.Count) { continue }
            if ($fillId -lt 0 -or $fillId -ge $fillHex.Count) { continue }
            $fontEl = $fontEls[$fontId]
            $fontColorEl = Get-ChildElement -Parent $fontEl -Namespace $NS.s -LocalName 'color'
            $fg = Get-ExcelExplicitColorHex $fontColorEl
            $bg = $fillHex[$fillId]
            if ([string]::IsNullOrEmpty($fg) -or [string]::IsNullOrEmpty($bg)) { continue }
            if ((Get-ContrastRatio -ForegroundHex $fg -BackgroundHex $bg) -ge 4.5) { continue }

            $newFg = Get-NearestPassingForeground -ForegroundHex $fg -BackgroundHex $bg -Threshold 4.5
            if (-not $newFg -or $newFg -eq $fg.ToUpper()) { continue }

            $pairKey = "{0}|{1}" -f $fontId, $bg.ToUpper()
            if (-not $newFontByPair.ContainsKey($pairKey)) {
                # Clone the original font, set the new color, append, and
                # record its new index.
                $cloned = New-Object System.Xml.Linq.XElement -ArgumentList $fontEl
                $clonedColor = Get-ChildElement -Parent $cloned -Namespace $NS.s -LocalName 'color'
                if ($clonedColor) {
                    foreach ($attrName in @('theme','indexed','auto','rgb','tint')) {
                        $a = $clonedColor.Attribute([System.Xml.Linq.XName]::Get($attrName))
                        if ($a) { $a.Remove() }
                    }
                    $clonedColor.SetAttributeValue([System.Xml.Linq.XName]::Get('rgb'), 'FF' + $newFg)
                } else {
                    $newColor = New-Object System.Xml.Linq.XElement -ArgumentList ([System.Xml.Linq.XName]::Get('color', $NS.s))
                    $newColor.SetAttributeValue([System.Xml.Linq.XName]::Get('rgb'), 'FF' + $newFg)
                    $cloned.Add($newColor)
                }
                $fontsEl.Add($cloned)
                $newFontByPair[$pairKey] = $fontEls.Count
                $fontEls += ,$cloned
            }

            # Clone the cellXf, point at the new font, append, and rebind the cell.
            $newFontId = $newFontByPair[$pairKey]
            $clonedXfKey = "$sIdx|$newFontId"
            $clonedXfIdx = $null
            if ($newFontByPair.ContainsKey("xf:$clonedXfKey")) {
                $clonedXfIdx = $newFontByPair["xf:$clonedXfKey"]
            } else {
                $clonedXf = New-Object System.Xml.Linq.XElement -ArgumentList $xf
                $clonedXf.SetAttributeValue([System.Xml.Linq.XName]::Get('fontId'), $newFontId)
                $clonedXf.SetAttributeValue([System.Xml.Linq.XName]::Get('applyFont'), '1')
                $cellXfsEl.Add($clonedXf)
                $clonedXfIdx = $xfEls.Count
                $xfEls += ,$clonedXf
                $newFontByPair["xf:$clonedXfKey"] = $clonedXfIdx
            }
            $sAttr.Value = "$clonedXfIdx"
            $fixed++
            $wsChanged = $true
        }
        if ($wsChanged) { Save-PartXDocument -Part $wsPart -Document $wsDoc }
    }

    if ($fixed -gt 0) {
        # Update count attributes on <fonts> / <cellXfs> if they exist.
        $fontsCount = $fontsEl.Attribute([System.Xml.Linq.XName]::Get('count'))
        if ($fontsCount) { $fontsCount.Value = "$($fontEls.Count)" }
        $xfCount = $cellXfsEl.Attribute([System.Xml.Linq.XName]::Get('count'))
        if ($xfCount) { $xfCount.Value = "$($xfEls.Count)" }
        Save-PartXDocument -Part $stylesPart -Document $stylesDoc
    }
    return $fixed
}

function Save-PartXDocument {
    param($Part, [System.Xml.Linq.XDocument] $Document)
    $stream = $Part.GetStream([IO.FileMode]::Create, [IO.FileAccess]::Write)
    try {
        $Document.Save($stream)
    } finally {
        $stream.Dispose()
    }
}

#--------------------------------------------------------------------
# Main: open document and run all rules
#--------------------------------------------------------------------

$issues = New-Object System.Collections.Generic.List[object]
$doc = $null

try {
    try {
        $doc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($FilePath, $false)
    }
    catch [System.IO.IOException] {
        [Console]::Error.WriteLine("Could not open file (is it locked by another process?): $($_.Exception.Message)")
        exit 2
    }
    catch {
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
            [Console]::Error.WriteLine("Failed to open workbook: $($_.Exception.Message)")
            exit 2
        }
        $issues.Add((New-Issue -Severity 'ERROR' -RuleName 'DocumentProtected' `
            -Description 'Workbook is IRM- or password-protected')) | Out-Null
    }

    if ($null -ne $doc) {
        $wbPart = $doc.WorkbookPart
        if ($null -eq $wbPart) {
            [Console]::Error.WriteLine("Workbook part missing — file may be corrupt.")
            exit 2
        }

        $sheetMap = Get-SheetNameMap -WorkbookPart $wbPart

        # Workbook-level rules
        foreach ($i in (Test-SheetTabName -WorkbookPart $wbPart)) { $issues.Add($i) | Out-Null }
        $usedNumFmtIds = Get-UsedNumFmtId -WorkbookPart $wbPart
        foreach ($i in (Test-RedOnlyNumberFormat -WorkbookStylesPart $wbPart.WorkbookStylesPart -UsedNumFmtIds $usedNumFmtIds)) {
            $issues.Add($i) | Out-Null
        }
        foreach ($i in (Test-LowContrast -WorkbookPart $wbPart -SheetMap $sheetMap)) {
            $issues.Add($i) | Out-Null
        }

        # Per-worksheet rules
        foreach ($wsPart in $wbPart.WorksheetParts) {
            $sheetName = Get-WorksheetDisplayName -WorkbookPart $wbPart -WorksheetPart $wsPart -SheetMap $sheetMap

            foreach ($i in (Test-MergedCell -WorksheetPart $wsPart -SheetName $sheetName)) {
                $issues.Add($i) | Out-Null
            }

            # Drawings on this worksheet
            $drawingsPart = $wsPart.DrawingsPart
            if ($null -ne $drawingsPart) {
                foreach ($i in (Test-AltText -DrawingsPart $drawingsPart -SheetName $sheetName)) {
                    $issues.Add($i) | Out-Null
                }
            }

            # Tables on this worksheet
            foreach ($tdp in $wsPart.TableDefinitionParts) {
                foreach ($i in (Test-TableDefinition -TableDefinitionPart $tdp)) {
                    $issues.Add($i) | Out-Null
                }
            }
        }
    }
}
finally {
    if ($null -ne $doc) {
        try { $doc.Dispose() } catch { Write-Verbose "Dispose failed: $($_.Exception.Message)" }
    }
}

#--------------------------------------------------------------------
# Output + exit
#--------------------------------------------------------------------

$errorIssues = @($issues | Where-Object { $_.Severity -eq 'ERROR' })
$errorCount = $errorIssues.Count
$exitCode = if ($errorCount -gt 0) { 1 } else { 0 }
if ($errorCount -gt 0) {
    $errorRules = ($errorIssues | Select-Object -ExpandProperty RuleName -Unique) -join ', '
    $summary = "FAIL $FilePath`: $errorRules"
} else {
    $summary = "PASS $FilePath"
}

if ($Format -eq 'detailed') {
    $severityRank = @{ 'ERROR' = 0; 'WARNING' = 1; 'TIP' = 2 }
    $sorted = $issues | Sort-Object `
        @{ Expression = { $severityRank[$_.Severity] } }, `
        @{ Expression = { $_.RuleName } }

    foreach ($issue in $sorted) {
        # TIPs only emit in detailed mode (text mode skips them entirely;
        # but in detailed mode all issues — including TIPs — appear).
        Write-Output ("{0}`t{1}`t{2}" -f $issue.Severity, $issue.RuleName, $issue.Description)
    }
    Write-Output $summary
}
else {
    Write-Output $summary
}

#--------------------------------------------------------------------
# Autofix
#--------------------------------------------------------------------

if (-not $Fix) { exit $exitCode }

$fixableRuleNames = @('MissingTableHeaders', 'RedOnlyNegativeFormatting', 'LowContrast')
$hasFixable = $false
foreach ($i in $issues) {
    if ($fixableRuleNames -contains $i.RuleName) { $hasFixable = $true; break }
}
if (-not $hasFixable) { exit $exitCode }

$fixedPath = [IO.Path]::ChangeExtension($FilePath, $null).TrimEnd('.') + '.fixed' + $ext
try {
    Copy-Item -LiteralPath $FilePath -Destination $fixedPath -Force
}
catch {
    [Console]::Error.WriteLine("Failed to create fixed copy at '$fixedPath': $($_.Exception.Message)")
    exit 2
}

$fixCounts = @{ MissingTableHeaders = 0; RedOnlyNegativeFormatting = 0; LowContrast = 0 }
$fixDoc = $null
try {
    try {
        $fixDoc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($fixedPath, $true)
    } catch {
        [Console]::Error.WriteLine("Failed to open fixed copy '$fixedPath' for editing: $($_.Exception.Message)")
        exit 2
    }
    $wbPartFix = $fixDoc.WorkbookPart
    if ($null -ne $wbPartFix) {
        foreach ($wsPart in $wbPartFix.WorksheetParts) {
            foreach ($tdp in $wsPart.TableDefinitionParts) {
                $fixCounts.MissingTableHeaders += (Repair-XlsxTableHeader -TableDefinitionPart $tdp)
            }
        }
        $fixCounts.RedOnlyNegativeFormatting = Repair-XlsxRedOnlyNumberFormat `
            -WorkbookStylesPart $wbPartFix.WorkbookStylesPart `
            -UsedNumFmtIds (Get-UsedNumFmtId -WorkbookPart $wbPartFix)
        $fixCounts.LowContrast = Repair-XlsxLowContrast -WorkbookPart $wbPartFix
    }
}
finally {
    if ($null -ne $fixDoc) {
        try { $fixDoc.Dispose() } catch { Write-Verbose "Dispose failed: $($_.Exception.Message)" }
    }
}

$fixSummaryParts = @()
foreach ($k in $fixableRuleNames) {
    if ($fixCounts[$k] -gt 0) { $fixSummaryParts += ('{0} ({1})' -f $k, $fixCounts[$k]) }
}
if ($fixSummaryParts.Count -eq 0) {
    Remove-Item -LiteralPath $fixedPath -Force -ErrorAction SilentlyContinue
    exit $exitCode
}

Write-Output ("FIXED {0}: {1}" -f $fixedPath, ($fixSummaryParts -join ', '))

& $PSCommandPath -FilePath $fixedPath -Format $Format
exit $LASTEXITCODE
