<#
.SYNOPSIS
    Runs Microsoft Word-style accessibility rules against a .docx/.docm file
    using the Open XML SDK.

.DESCRIPTION
    Inspects the document body, every header part, and every footer part for
    common accessibility issues (missing alt text, missing table headers,
    document protection, missing content control titles, merged cells,
    heading-order skips, floating objects, repeated blank characters, lack of
    headings). Emits a single PASS/FAIL line in 'text' mode, or a sorted list
    of issues followed by the PASS/FAIL line in 'detailed' mode.

.PARAMETER FilePath
    Path to a .docx or .docm file.

.PARAMETER Format
    'text' (default) emits only PASS/FAIL. 'detailed' emits one tab-separated
    issue per line plus the PASS/FAIL summary.

.OUTPUTS
    Exit codes:
      0  no errors found (warnings/tips do not fail)
      1  one or more accessibility errors found (includes IRM/password-protected files)
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
function Test-TableHeader {
    param([Xml.Linq.XContainer] $Root, [int] $StartIndex, [ref] $Issues)
    if (-not $Root) { return $StartIndex }

    $tables = Get-Descendant -Root $Root -Namespace $NS.w -LocalName 'tbl'
    $idx = $StartIndex
    foreach ($tbl in $tables) {
        $idx++
        $rows = $tbl.Elements([Xml.Linq.XName]::Get('tr', $NS.w))
        if (-not $rows -or $rows.Count -eq 0) { continue }

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

# --- Rule: ContrastSkipped (TIP) ---------------------------------------------
function Test-ContrastSkipped {
    param([ref] $Issues)
    $Issues.Value += New-Issue -Severity TIP -RuleName 'ContrastSkipped' `
        -Description 'Contrast check skipped (requires rendering)'
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

        # TIP rules only matter for detailed mode but cheap to compute always.
        Test-NoHeadingStyle -Roots $allRoots -Issues $issuesRef
        Test-ContrastSkipped -Issues $issuesRef

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
    Write-Output "FAIL $FilePath"
    exit 1
} else {
    Write-Output "PASS $FilePath"
    exit 0
}
