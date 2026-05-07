<#
.SYNOPSIS
    Runs Microsoft Office Accessibility Checker rules against a .pptx/.pptm
    file using the Open XML SDK.

.DESCRIPTION
    Mirrors the rule subset the built-in PowerPoint Accessibility Checker
    surfaces, walking presentation XML directly via the Open XML SDK. No Office
    install required.

    Rule coverage:
      ERROR:   MissingAltText, MissingSlideTitle, MissingTableHeaders,
               DocumentProtected
      WARNING: DuplicateSlideTitle, MergedTableCells, LowContrast,
               NonDescriptiveLinkText

    Notes:
      - LowContrast: best-effort. Only flags runs whose foreground text color
        and the owning-shape (or slide-background) fill are both explicit RGB.
        Theme/scheme references, gradient/non-solid fills, and any inherited
        colors are skipped.
      - NonDescriptiveLinkText: run-level hyperlinks only (a:hlinkClick on a
        run). Fires when the visible text is empty, equals the URL, or matches
        a known generic blacklist.

.PARAMETER FilePath
    Path to a .pptx or .pptm file.

.PARAMETER Format
    'text' (default) emits a single PASS/FAIL line.
    'detailed' emits every issue (one tab-separated line) followed by the
    PASS/FAIL summary line.

.OUTPUTS
    Exit codes:
      0  no errors found (warnings do not fail)
      1  one or more accessibility errors found (includes IRM/password-protected)
      2  tool error (file not found, unsupported format, SDK load failure, file locked)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $FilePath,

    [ValidateSet('text','detailed')]
    [string] $Format = 'text'
)

$ErrorActionPreference = 'Stop'

# --- Pinned OOXML namespaces -------------------------------------------------
$NS = @{
    p    = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    a    = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    r    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    adec = 'http://schemas.microsoft.com/office/drawing/2017/decorative'
}

# Generic phrases that fail NonDescriptiveLinkText. Compared case-insensitively
# against the run's visible text after trimming whitespace and trailing
# punctuation.
$NonDescriptiveLinkPhrases = @(
    'click here', 'click', 'here', 'more', 'read more', 'link', 'this link', 'this'
)

#--------------------------------------------------------------------
# Pre-flight: file existence + extension
#--------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $FilePath)) {
    [Console]::Error.WriteLine("File not found: $FilePath")
    exit 2
}

$ext = [IO.Path]::GetExtension($FilePath).ToLowerInvariant()
if ($ext -ne '.pptx' -and $ext -ne '.pptm') {
    [Console]::Error.WriteLine("Unsupported file type for PowerPoint checker: $ext (expected .pptx or .pptm)")
    exit 2
}

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
    [Console]::Error.WriteLine("Open XML SDK not found. Run scripts\setup-accessibility-checker.ps1 first.")
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
        Justification = 'New-Issue is a pure record factory.'
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

function Get-Descendant {
    param([System.Xml.Linq.XContainer] $Root, [string] $Namespace, [string] $LocalName)
    if ($null -eq $Root) { return @() }
    return $Root.Descendants([System.Xml.Linq.XName]::Get($LocalName, $Namespace))
}

#--------------------------------------------------------------------
# Slide ordering: walk presentation.SlideIdList, resolve each entry to a
# SlidePart (skip notes/handouts that share the relationship space).
#--------------------------------------------------------------------

function Get-OrderedSlideList {
    param($PresentationPart)

    $result = New-Object 'System.Collections.Generic.List[object]'
    $presDoc = Get-PartXDocument -Part $PresentationPart
    if ($null -eq $presDoc -or $null -eq $presDoc.Root) { return $result }

    $sldIdLst = Get-ChildElement -Parent $presDoc.Root -Namespace $NS.p -LocalName 'sldIdLst'
    if ($null -eq $sldIdLst) { return $result }

    $i = 0
    foreach ($sldId in (Get-ChildElement -Parent $sldIdLst -Namespace $NS.p -LocalName 'sldId' -All)) {
        $i++
        $rid = Get-XAttr -Element $sldId -Namespace $NS.r -LocalName 'id'
        if ([string]::IsNullOrEmpty($rid)) { continue }
        $part = $null
        try { $part = $PresentationPart.GetPartById($rid) } catch { Write-Verbose "GetPartById failed for $rid" }
        if ($part -is [DocumentFormat.OpenXml.Packaging.SlidePart]) {
            $result.Add([pscustomobject]@{
                Number    = $i
                SlidePart = $part
            }) | Out-Null
        }
    }
    return $result
}

#--------------------------------------------------------------------
# Common shape-walking helpers
#--------------------------------------------------------------------

# Concatenate every a:t descendant of an element. Used for placeholder text,
# table cell text, and run text alike.
function Get-XmlText {
    param([System.Xml.Linq.XElement] $Element)
    if ($null -eq $Element) { return '' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in (Get-Descendant -Root $Element -Namespace $NS.a -LocalName 't')) {
        [void] $sb.Append($t.Value)
    }
    return $sb.ToString()
}

# Walk a spTree, yielding every "leaf" shape we want to evaluate. Group shapes
# (p:grpSp) are recursed into; their containers themselves are not yielded.
# Returns an array of XElement.
function Get-LeafShape {
    param([System.Xml.Linq.XElement] $SpTree)
    if ($null -eq $SpTree) { return @() }
    $shapes = New-Object 'System.Collections.Generic.List[object]'
    foreach ($child in $SpTree.Elements()) {
        if ($child.Name.NamespaceName -ne $NS.p) { continue }
        switch ($child.Name.LocalName) {
            'sp'           { $shapes.Add($child) | Out-Null }
            'pic'          { $shapes.Add($child) | Out-Null }
            'graphicFrame' { $shapes.Add($child) | Out-Null }
            'cxnSp'        { $shapes.Add($child) | Out-Null }
            'grpSp' {
                foreach ($nested in (Get-LeafShape -SpTree $child)) {
                    $shapes.Add($nested) | Out-Null
                }
            }
        }
    }
    return $shapes
}

# Locate the cNvPr inside a shape's non-visual properties container. Returns
# $null if the container is missing.
function Get-ShapeCNvPr {
    param([System.Xml.Linq.XElement] $Shape)
    if ($null -eq $Shape) { return $null }
    $local = $Shape.Name.LocalName
    $nvName = switch ($local) {
        'sp'           { 'nvSpPr' }
        'pic'          { 'nvPicPr' }
        'graphicFrame' { 'nvGraphicFramePr' }
        'cxnSp'        { 'nvCxnSpPr' }
        default        { $null }
    }
    if (-not $nvName) { return $null }
    $nv = Get-ChildElement -Parent $Shape -Namespace $NS.p -LocalName $nvName
    if ($null -eq $nv) { return $null }
    return Get-ChildElement -Parent $nv -Namespace $NS.p -LocalName 'cNvPr'
}

#--------------------------------------------------------------------
# Rule: MissingAltText
#--------------------------------------------------------------------

# Pass if any of:
#   - @descr non-empty
#   - @title non-empty
#   - @decorative in {1,true}    (older direct attribute)
#   - <a:extLst><a:ext><adec:decorative val="1"/>   (Office 2019/365 modern marker)
function Test-PptxAltTextAttribute {
    param([System.Xml.Linq.XElement] $CNvPr)
    if ($null -eq $CNvPr) { return $false }
    if (-not [string]::IsNullOrWhiteSpace((Get-XAttr -Element $CNvPr -Namespace $null -LocalName 'descr'))) { return $true }
    if (-not [string]::IsNullOrWhiteSpace((Get-XAttr -Element $CNvPr -Namespace $null -LocalName 'title'))) { return $true }
    $decorative = Get-XAttr -Element $CNvPr -Namespace $null -LocalName 'decorative'
    if ($decorative -eq '1' -or $decorative -eq 'true') { return $true }
    return (Test-DecorativeExtension -CNvPr $CNvPr)
}

# Office 2019+ stores the "Mark as decorative" flag in an extension list:
#   <a:extLst>
#     <a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">
#       <adec:decorative xmlns:adec="..." val="1"/>
#     </a:ext>
#   </a:extLst>
# Walk it and treat val in {1,true} as decorative.
function Test-DecorativeExtension {
    param([System.Xml.Linq.XElement] $CNvPr)
    if ($null -eq $CNvPr) { return $false }
    foreach ($dec in (Get-Descendant -Root $CNvPr -Namespace $NS.adec -LocalName 'decorative')) {
        $val = Get-XAttr -Element $dec -Namespace $null -LocalName 'val'
        if ($val -eq '1' -or $val -eq 'true') { return $true }
    }
    return $false
}

# Decide whether a shape requires alt text. PowerPoint's checker exempts:
#   - placeholder shapes (title/body etc.) whose visual content is inherited
#     from the layout/master and whose role is to hold text;
#   - graphicFrames containing a table — tables are text content, screen
#     readers consume the cells directly, so a redundant description is not
#     required.
# Pictures, charts/SmartArt graphicFrames, group children (already recursed),
# connectors, and non-placeholder p:sp shapes all require alt text.
function Test-ShapeNeedsAltText {
    param([System.Xml.Linq.XElement] $Shape)
    if ($null -eq $Shape) { return $false }
    $local = $Shape.Name.LocalName

    if ($local -eq 'sp') {
        $nvSpPr = Get-ChildElement -Parent $Shape -Namespace $NS.p -LocalName 'nvSpPr'
        $nvPr = Get-ChildElement -Parent $nvSpPr -Namespace $NS.p -LocalName 'nvPr'
        $ph = Get-ChildElement -Parent $nvPr -Namespace $NS.p -LocalName 'ph'
        if ($ph) { return $false }
        return $true
    }

    if ($local -eq 'graphicFrame') {
        # Tables are accessible content; charts and SmartArt are not.
        $tbl = Get-Descendant -Root $Shape -Namespace $NS.a -LocalName 'tbl' | Select-Object -First 1
        if ($tbl) { return $false }
        return $true
    }

    return $true   # pic, cxnSp
}

function Test-PptxAltText {
    param(
        [System.Xml.Linq.XElement] $SpTree,
        [int] $SlideNumber,
        [System.Collections.Generic.List[object]] $Issues
    )
    if ($null -eq $SpTree) { return }

    foreach ($shape in (Get-LeafShape -SpTree $SpTree)) {
        if (-not (Test-ShapeNeedsAltText $shape)) { continue }

        $cNvPr = Get-ShapeCNvPr -Shape $shape
        if (-not (Test-PptxAltTextAttribute $cNvPr)) {
            $name = if ($cNvPr) { Get-XAttr -Element $cNvPr -Namespace $null -LocalName 'name' } else { $null }
            if ([string]::IsNullOrWhiteSpace($name)) { $name = '(unnamed)' }
            [void] $Issues.Add((New-Issue -Severity 'ERROR' -RuleName 'MissingAltText' `
                -Description "Image/object `"$name`" on slide $SlideNumber has no alt text"))
        }
    }
}

#--------------------------------------------------------------------
# Rule: MissingSlideTitle
#--------------------------------------------------------------------

# Returns the slide's displayed title text, or $null if neither the slide nor
# its layout defines a title placeholder.
#
# Resolution order:
#   1. Walk the slide spTree for a p:sp with p:ph type ∈ {title, ctrTitle}.
#      If found, return its txBody text (may be the empty string if the slide
#      author left the placeholder empty).
#   2. If the slide has no title-typed placeholder, consult the slide's layout
#      (slidePart.SlideLayoutPart). When the layout defines a title placeholder
#      the slide *inherits* it; its rendered text is whatever the slide's
#      matching <p:ph> shape contains, or empty if the slide doesn't override
#      the placeholder. Either way the slide should have a title — we return
#      the empty string to surface MissingSlideTitle as "empty title".
#   3. If neither slide nor layout defines a title-typed placeholder, return
#      $null and the caller reports "no title placeholder".
function Get-SlideTitleText {
    param([System.Xml.Linq.XElement] $SpTree, $SlidePart)
    if ($null -eq $SpTree) { return $null }

    foreach ($shape in (Get-LeafShape -SpTree $SpTree)) {
        if ($shape.Name.LocalName -ne 'sp') { continue }
        $nvSpPr = Get-ChildElement -Parent $shape -Namespace $NS.p -LocalName 'nvSpPr'
        $nvPr = Get-ChildElement -Parent $nvSpPr -Namespace $NS.p -LocalName 'nvPr'
        $ph = Get-ChildElement -Parent $nvPr -Namespace $NS.p -LocalName 'ph'
        if ($null -eq $ph) { continue }
        $type = Get-XAttr -Element $ph -Namespace $null -LocalName 'type'
        if ($type -ne 'title' -and $type -ne 'ctrTitle') { continue }
        $txBody = Get-ChildElement -Parent $shape -Namespace $NS.p -LocalName 'txBody'
        return (Get-XmlText -Element $txBody)
    }

    # Layout fallback: check whether the slide's layout declares a title-typed
    # placeholder. If so, the slide inherits a title shape but didn't override
    # it — surface as an empty title rather than "no title placeholder".
    if ($null -ne $SlidePart -and $null -ne $SlidePart.SlideLayoutPart) {
        $layoutDoc = Get-PartXDocument -Part $SlidePart.SlideLayoutPart
        if ($null -ne $layoutDoc -and $null -ne $layoutDoc.Root) {
            $layoutCSld = Get-ChildElement -Parent $layoutDoc.Root -Namespace $NS.p -LocalName 'cSld'
            $layoutSpTree = if ($layoutCSld) { Get-ChildElement -Parent $layoutCSld -Namespace $NS.p -LocalName 'spTree' } else { $null }
            foreach ($lShape in (Get-LeafShape -SpTree $layoutSpTree)) {
                if ($lShape.Name.LocalName -ne 'sp') { continue }
                $lNvSp = Get-ChildElement -Parent $lShape -Namespace $NS.p -LocalName 'nvSpPr'
                $lNvPr = Get-ChildElement -Parent $lNvSp -Namespace $NS.p -LocalName 'nvPr'
                $lPh   = Get-ChildElement -Parent $lNvPr -Namespace $NS.p -LocalName 'ph'
                if ($null -eq $lPh) { continue }
                $lType = Get-XAttr -Element $lPh -Namespace $null -LocalName 'type'
                if ($lType -eq 'title' -or $lType -eq 'ctrTitle') {
                    return ''
                }
            }
        }
    }

    return $null
}

function Test-SlideTitle {
    param(
        [System.Xml.Linq.XElement] $SpTree,
        $SlidePart,
        [int] $SlideNumber,
        [System.Collections.Generic.List[object]] $Issues
    )
    $title = Get-SlideTitleText -SpTree $SpTree -SlidePart $SlidePart
    if ($null -eq $title) {
        [void] $Issues.Add((New-Issue -Severity 'ERROR' -RuleName 'MissingSlideTitle' `
            -Description "Slide $SlideNumber has no title placeholder"))
        return
    }
    if ([string]::IsNullOrWhiteSpace($title)) {
        [void] $Issues.Add((New-Issue -Severity 'ERROR' -RuleName 'MissingSlideTitle' `
            -Description "Slide $SlideNumber has an empty title"))
    }
}

#--------------------------------------------------------------------
# Rule: MissingTableHeaders + MergedTableCells
#--------------------------------------------------------------------

# Locate every a:tbl on the slide. Tables sit inside p:graphicFrame shapes
# (a:graphic/a:graphicData/a:tbl), so descendants is the simplest walk.
function Get-SlideTable {
    param([System.Xml.Linq.XElement] $SpTree)
    if ($null -eq $SpTree) { return @() }
    return Get-Descendant -Root $SpTree -Namespace $NS.a -LocalName 'tbl'
}

function Test-TableHeader {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '',
        Justification = 'Function evaluates the MissingTableHeaders rule.')]
    param(
        [System.Xml.Linq.XElement] $SpTree,
        [int] $SlideNumber,
        [System.Collections.Generic.List[object]] $Issues
    )

    $tables = @(Get-SlideTable -SpTree $SpTree)
    if ($tables.Count -eq 0) { return }

    $idx = 0
    foreach ($tbl in $tables) {
        $idx++
        $tblPr = Get-ChildElement -Parent $tbl -Namespace $NS.a -LocalName 'tblPr'
        $firstRow = if ($tblPr) { Get-XAttr -Element $tblPr -Namespace $null -LocalName 'firstRow' } else { $null }
        if ($firstRow -ne '1' -and $firstRow -ne 'true') {
            [void] $Issues.Add((New-Issue -Severity 'ERROR' -RuleName 'MissingTableHeaders' `
                -Description "Table $idx on slide $SlideNumber has no header row"))
        }
    }
}

function Test-TableMergedCell {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '',
        Justification = 'Function evaluates the MergedTableCells rule.')]
    param(
        [System.Xml.Linq.XElement] $SpTree,
        [int] $SlideNumber,
        [System.Collections.Generic.List[object]] $Issues
    )

    $tables = @(Get-SlideTable -SpTree $SpTree)
    if ($tables.Count -eq 0) { return }

    $idx = 0
    foreach ($tbl in $tables) {
        $idx++
        $hasMerge = $false
        foreach ($tc in (Get-Descendant -Root $tbl -Namespace $NS.a -LocalName 'tc')) {
            $gridSpan = Get-XAttr -Element $tc -Namespace $null -LocalName 'gridSpan'
            $rowSpan  = Get-XAttr -Element $tc -Namespace $null -LocalName 'rowSpan'
            $hMerge   = Get-XAttr -Element $tc -Namespace $null -LocalName 'hMerge'
            $vMerge   = Get-XAttr -Element $tc -Namespace $null -LocalName 'vMerge'

            $gsInt = 0
            $rsInt = 0
            if ([int]::TryParse($gridSpan, [ref] $gsInt) -and $gsInt -gt 1) { $hasMerge = $true; break }
            if ([int]::TryParse($rowSpan, [ref] $rsInt) -and $rsInt -gt 1)   { $hasMerge = $true; break }
            if ($hMerge -eq '1' -or $hMerge -eq 'true') { $hasMerge = $true; break }
            if ($vMerge -eq '1' -or $vMerge -eq 'true') { $hasMerge = $true; break }
        }
        if ($hasMerge) {
            [void] $Issues.Add((New-Issue -Severity 'WARNING' -RuleName 'MergedTableCells' `
                -Description "Table $idx on slide $SlideNumber contains merged cells"))
        }
    }
}

#--------------------------------------------------------------------
# Rule: NonDescriptiveLinkText
#--------------------------------------------------------------------

# Strip leading/trailing whitespace and trailing punctuation so "Click here."
# matches the same blacklist as "click here".
function Get-NormalizedLinkText {
    param([string] $Text)
    if ($null -eq $Text) { return '' }
    $stripped = $Text.Trim()
    $stripped = $stripped -replace '[.!?,:;\s]+$', ''
    return $stripped.ToLowerInvariant()
}

function Test-LinkText {
    param(
        $SlidePart,
        [System.Xml.Linq.XElement] $SpTree,
        [int] $SlideNumber,
        [System.Collections.Generic.List[object]] $Issues
    )
    if ($null -eq $SpTree) { return }

    # Build relationship-id -> URL map up front; iterating the runs more than
    # once and looking up by Id each time avoids round-tripping the SDK
    # collection per-run.
    $relMap = @{}
    if ($SlidePart -and $SlidePart.HyperlinkRelationships) {
        foreach ($rel in $SlidePart.HyperlinkRelationships) {
            $relMap[$rel.Id] = $rel.Uri.ToString()
        }
    }

    foreach ($run in (Get-Descendant -Root $SpTree -Namespace $NS.a -LocalName 'r')) {
        $rPr = Get-ChildElement -Parent $run -Namespace $NS.a -LocalName 'rPr'
        if ($null -eq $rPr) { continue }
        $hlink = Get-ChildElement -Parent $rPr -Namespace $NS.a -LocalName 'hlinkClick'
        if ($null -eq $hlink) { continue }

        $rid = Get-XAttr -Element $hlink -Namespace $NS.r -LocalName 'id'
        $url = $null
        if ($rid -and $relMap.ContainsKey($rid)) { $url = $relMap[$rid] }

        $tEl = Get-ChildElement -Parent $run -Namespace $NS.a -LocalName 't'
        $rawText = if ($tEl) { $tEl.Value } else { '' }
        $normalized = Get-NormalizedLinkText -Text $rawText

        $bad = $false
        if ([string]::IsNullOrWhiteSpace($rawText)) {
            $bad = $true
        }
        elseif ($url -and ($rawText.Trim().ToLowerInvariant() -eq $url.ToLowerInvariant())) {
            $bad = $true
        }
        else {
            foreach ($phrase in $NonDescriptiveLinkPhrases) {
                if ($normalized -eq $phrase) { $bad = $true; break }
            }
        }

        if ($bad) {
            $preview = if ([string]::IsNullOrEmpty($rawText)) { '(empty)' }
                       elseif ($rawText.Length -gt 30) { $rawText.Substring(0, 30) + '...' }
                       else { $rawText }
            [void] $Issues.Add((New-Issue -Severity 'WARNING' -RuleName 'NonDescriptiveLinkText' `
                -Description "Hyperlink text `"$preview`" on slide $SlideNumber is not descriptive"))
        }
    }
}

#--------------------------------------------------------------------
# Rule: LowContrast
#--------------------------------------------------------------------

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

# Resolve a fillable element's explicit srgb value, or $null. Skips
# scheme/system colors and non-solid fill types.
function Get-ExplicitSolidFillHex {
    param([System.Xml.Linq.XElement] $Container)
    if ($null -eq $Container) { return $null }
    $solidFill = Get-ChildElement -Parent $Container -Namespace $NS.a -LocalName 'solidFill'
    if ($null -eq $solidFill) { return $null }
    $srgb = Get-ChildElement -Parent $solidFill -Namespace $NS.a -LocalName 'srgbClr'
    if ($null -eq $srgb) { return $null }
    $val = Get-XAttr -Element $srgb -Namespace $null -LocalName 'val'
    if ($val -and $val -match '^[0-9A-Fa-f]{6}$') { return $val }
    return $null
}

# Slide-background fill: p:cSld/p:bg/p:bgPr/a:solidFill/a:srgbClr.
function Get-SlideBackgroundHex {
    param([System.Xml.Linq.XElement] $CSld)
    if ($null -eq $CSld) { return $null }
    $bg = Get-ChildElement -Parent $CSld -Namespace $NS.p -LocalName 'bg'
    if ($null -eq $bg) { return $null }
    $bgPr = Get-ChildElement -Parent $bg -Namespace $NS.p -LocalName 'bgPr'
    return Get-ExplicitSolidFillHex -Container $bgPr
}

function Test-LowContrast {
    param(
        [System.Xml.Linq.XElement] $CSld,
        [System.Xml.Linq.XElement] $SpTree,
        [int] $SlideNumber,
        [System.Collections.Generic.List[object]] $Issues
    )
    if ($null -eq $SpTree) { return }

    $slideBg = Get-SlideBackgroundHex -CSld $CSld

    foreach ($shape in (Get-LeafShape -SpTree $SpTree)) {
        if ($shape.Name.LocalName -ne 'sp') { continue }
        $spPr = Get-ChildElement -Parent $shape -Namespace $NS.p -LocalName 'spPr'
        $shapeFill = Get-ExplicitSolidFillHex -Container $spPr
        $bg = if ($shapeFill) { $shapeFill } else { $slideBg }
        if (-not $bg) { continue }

        $txBody = Get-ChildElement -Parent $shape -Namespace $NS.p -LocalName 'txBody'
        if ($null -eq $txBody) { continue }

        foreach ($run in (Get-Descendant -Root $txBody -Namespace $NS.a -LocalName 'r')) {
            $rPr = Get-ChildElement -Parent $run -Namespace $NS.a -LocalName 'rPr'
            $fg = Get-ExplicitSolidFillHex -Container $rPr
            if (-not $fg) { continue }

            $ratio = Get-ContrastRatio -ForegroundHex $fg -BackgroundHex $bg
            if ($ratio -lt 4.5) {
                $tEl = Get-ChildElement -Parent $run -Namespace $NS.a -LocalName 't'
                $preview = if ($tEl) { $tEl.Value } else { '(empty run)' }
                if ([string]::IsNullOrEmpty($preview)) { $preview = '(empty run)' }
                elseif ($preview.Length -gt 30) { $preview = $preview.Substring(0, 30) + '...' }
                [void] $Issues.Add((New-Issue -Severity 'WARNING' -RuleName 'LowContrast' `
                    -Description ("Text `"{0}`" on slide {1} has contrast {2:N2}:1 (#{3} on #{4}); WCAG requires 4.5:1" -f $preview, $SlideNumber, $ratio, $fg.ToUpper(), $bg.ToUpper())))
            }
        }
    }
}

#--------------------------------------------------------------------
# Main: open document and run all rules
#--------------------------------------------------------------------

$issues = New-Object 'System.Collections.Generic.List[object]'
$doc = $null

try {
    try {
        $doc = [DocumentFormat.OpenXml.Packaging.PresentationDocument]::Open($FilePath, $false)
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
            [Console]::Error.WriteLine("Failed to open presentation: $($_.Exception.Message)")
            exit 2
        }
        [void] $issues.Add((New-Issue -Severity 'ERROR' -RuleName 'DocumentProtected' `
            -Description 'Presentation is IRM- or password-protected'))
        $doc = $null
    }

    if ($null -ne $doc) {
        $presPart = $doc.PresentationPart
        if ($null -eq $presPart) {
            [Console]::Error.WriteLine("Presentation part missing — file may be corrupt.")
            exit 2
        }

        $orderedSlides = Get-OrderedSlideList -PresentationPart $presPart
        $titleSeen = @{}

        foreach ($entry in $orderedSlides) {
            $slidePart = $entry.SlidePart
            $slideNum  = $entry.Number

            $slideDoc = Get-PartXDocument -Part $slidePart
            if ($null -eq $slideDoc -or $null -eq $slideDoc.Root) { continue }
            $cSld = Get-ChildElement -Parent $slideDoc.Root -Namespace $NS.p -LocalName 'cSld'
            $spTree = if ($cSld) { Get-ChildElement -Parent $cSld -Namespace $NS.p -LocalName 'spTree' } else { $null }

            Test-PptxAltText      -SpTree $spTree -SlideNumber $slideNum -Issues $issues
            Test-SlideTitle       -SpTree $spTree -SlidePart $slidePart -SlideNumber $slideNum -Issues $issues
            Test-TableHeader      -SpTree $spTree -SlideNumber $slideNum -Issues $issues
            Test-TableMergedCell  -SpTree $spTree -SlideNumber $slideNum -Issues $issues
            Test-LinkText         -SlidePart $slidePart -SpTree $spTree -SlideNumber $slideNum -Issues $issues
            Test-LowContrast      -CSld $cSld -SpTree $spTree -SlideNumber $slideNum -Issues $issues

            # Track titles for DuplicateSlideTitle (skip empty/missing titles —
            # those already fired MissingSlideTitle).
            $titleText = Get-SlideTitleText -SpTree $spTree -SlidePart $slidePart
            if (-not [string]::IsNullOrWhiteSpace($titleText)) {
                $key = $titleText.Trim().ToLowerInvariant()
                if (-not $titleSeen.ContainsKey($key)) { $titleSeen[$key] = @() }
                $titleSeen[$key] += $slideNum
            }
        }

        foreach ($key in $titleSeen.Keys) {
            $slidesWithTitle = $titleSeen[$key]
            if ($slidesWithTitle.Count -gt 1) {
                $list = ($slidesWithTitle | Sort-Object) -join ', '
                [void] $issues.Add((New-Issue -Severity 'WARNING' -RuleName 'DuplicateSlideTitle' `
                    -Description "Slides $list share the title `"$key`""))
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
        Write-Output ("{0}`t{1}`t{2}" -f $issue.Severity, $issue.RuleName, $issue.Description)
    }
    Write-Output $summary
}
else {
    Write-Output $summary
}

exit $exitCode
