<#
.SYNOPSIS
    One-shot regenerator for the .docx/.xlsx fixtures the Pester suite reads.

.DESCRIPTION
    Each fixture is a minimal OOXML package built directly via the Open XML
    SDK + raw XML, intentionally containing (or omitting) the precise marker
    each accessibility rule looks for. The output is deterministic: re-running
    overwrites whatever is already on disk.

    The fixtures themselves are committed to the repo -- run this script only
    when you change a rule, change a builder, or want to verify the committed
    binaries match the source. After running, commit the resulting binaries
    alongside the rule/builder change so reviewers see both moves in one PR.

    The DocumentProtected rule has no fixture: it fires when the package is
    wrapped in an encrypted Compound File envelope, which the Open XML SDK
    does not emit.

.NOTES
    Run scripts/setup-accessibility-checker.ps1 first to populate
    scripts/lib/. This script reuses the same SDK DLL the checkers load.
#>

[CmdletBinding()]
param(
    [string] $OutputDir = (Join-Path $PSScriptRoot 'fixtures')
)

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

# --- Locate and load the Open XML SDK (same logic as the checkers) ----------
$sdkPath = $null
if ($env:OPENXML_SDK_PATH -and (Test-Path -LiteralPath $env:OPENXML_SDK_PATH)) {
    $sdkPath = $env:OPENXML_SDK_PATH
} else {
    $candidate = Join-Path (Split-Path -Parent $PSScriptRoot) 'lib\DocumentFormat.OpenXml.dll'
    if (Test-Path -LiteralPath $candidate) { $sdkPath = $candidate }
}
if (-not $sdkPath) {
    [Console]::Error.WriteLine("Open XML SDK not found. Run scripts/setup-accessibility-checker.ps1 first.")
    exit 2
}
Add-Type -Path $sdkPath

if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

# --- Shared helpers ---------------------------------------------------------

# Write a UTF-8 XML string into the named part of an Open XML package.
function Set-PartXml {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Set-PartXml writes a byte stream into a temporary OOXML part inside the test fixture being built; -WhatIf adds nothing here because the entire output directory is rebuilt deterministically every run.'
    )]
    param(
        [Parameter(Mandatory)] $Part,
        [Parameter(Mandatory)] [string] $Xml
    )
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Xml)
    $stream = New-Object System.IO.MemoryStream(,$bytes)
    try {
        $Part.FeedData($stream)
    } finally {
        $stream.Dispose()
    }
}

# AddNewPart<T>(string id) is a generic method on OpenXmlPartContainer that
# PowerShell can't dispatch by type inference. Wrap it in reflection so the
# call sites stay readable.
$AddNewPartCache = @{}
function Add-OpenXmlPart {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Helper that adds a typed part to an Open XML container; no environmental side effects.'
    )]
    param(
        [Parameter(Mandatory)] $Container,
        [Parameter(Mandatory)] [type] $PartType,
        [Parameter(Mandatory)] [string] $RelId
    )
    $key = "$($Container.GetType().FullName)|$($PartType.FullName)"
    if (-not $AddNewPartCache.ContainsKey($key)) {
        $method = $Container.GetType().GetMethods() |
            Where-Object {
                $_.Name -eq 'AddNewPart' -and
                $_.IsGenericMethodDefinition -and
                $_.GetParameters().Count -eq 1 -and
                $_.GetParameters()[0].ParameterType -eq [string]
            } |
            Select-Object -First 1
        if (-not $method) {
            throw "Could not locate AddNewPart<T>(string) on $($Container.GetType().FullName)"
        }
        $AddNewPartCache[$key] = $method.MakeGenericMethod($PartType)
    }
    return $AddNewPartCache[$key].Invoke($Container, @($RelId))
}

# --- Word XML fragments -----------------------------------------------------
#
# Every fixture's document.xml is built by concatenating these fragments. The
# checker only walks specific OOXML markers, so the bodies stay minimal --
# no styles part, no themes, no real image binary.

$WordNamespaces = @'
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
'@

function Get-WordParagraph {
    param(
        [string] $Text = '',
        [string] $Style = $null,
        [string] $InnerXml = $null
    )
    $pPr = if ($Style) { "<w:pPr><w:pStyle w:val=`"$Style`"/></w:pPr>" } else { '' }
    $run = if ($InnerXml) { "<w:r>$InnerXml</w:r>" }
           elseif ($Text) { "<w:r><w:t xml:space=`"preserve`">$Text</w:t></w:r>" }
           else { '' }
    "<w:p>$pPr$run</w:p>"
}

# Inline drawing (passes FloatingObject; alt-text status controlled by params).
function Get-WordInlineDrawing {
    param(
        [int] $Id = 1,
        [string] $Name = "Picture $Id",
        [string] $Descr = '',
        [string] $Title = '',
        [bool] $Decorative = $false
    )
    $descrAttr = if ($Descr)     { " descr=`"$Descr`"" } else { '' }
    $titleAttr = if ($Title)     { " title=`"$Title`"" } else { '' }
    $decoAttr  = if ($Decorative){ ' decorative="1"' }   else { '' }
    @"
<w:drawing>
  <wp:inline distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="914400" cy="914400"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:docPr id="$Id" name="$Name"$descrAttr$titleAttr$decoAttr/>
    <wp:cNvGraphicFramePr/>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/>
    </a:graphic>
  </wp:inline>
</w:drawing>
"@
}

# Anchored (floating) drawing -- triggers FloatingObject.
function Get-WordAnchorDrawing {
    param(
        [int] $Id = 2,
        [string] $Name = "Picture $Id",
        [string] $Descr = ''
    )
    $descrAttr = if ($Descr) { " descr=`"$Descr`"" } else { '' }
    @"
<w:drawing>
  <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
    <wp:extent cx="914400" cy="914400"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:wrapSquare wrapText="bothSides"/>
    <wp:docPr id="$Id" name="$Name"$descrAttr/>
    <wp:cNvGraphicFramePr/>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/>
    </a:graphic>
  </wp:anchor>
</w:drawing>
"@
}

function Get-WordTable {
    param(
        [bool] $HasHeader = $true,
        [bool] $HasMerge = $false
    )
    $hdrPr = if ($HasHeader) { '<w:trPr><w:tblHeader/></w:trPr>' } else { '' }
    $cell1Pr = if ($HasMerge) { '<w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>' }
                else         { '<w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>' }
    $cell2 = if ($HasMerge) { '' }
             else           { '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>' }
    @"
<w:tbl>
  <w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr>
  <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
  <w:tr>
    $hdrPr
    <w:tc>$cell1Pr<w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
    $cell2
  </w:tr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>
"@
}

function Get-WordContentControl {
    param([string] $Title = '')
    $alias = if ($Title) { "<w:alias w:val=`"$Title`"/>" } else { '' }
    @"
<w:sdt>
  <w:sdtPr>$alias<w:text/></w:sdtPr>
  <w:sdtContent>
    <w:p><w:r><w:t>field</w:t></w:r></w:p>
  </w:sdtContent>
</w:sdt>
"@
}

# Wrap a body fragment in the document/body envelope.
function Get-WordDocumentXml {
    param([string] $BodyXml)
    @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document $WordNamespaces>
  <w:body>
$BodyXml
    <w:sectPr/>
  </w:body>
</w:document>
"@
}

# Create a Word package, write document.xml, save.
function Build-WordFixture {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $DocumentXml
    )
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $type = [DocumentFormat.OpenXml.WordprocessingDocumentType]::Document
    $doc = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Create($Path, $type)
    try {
        $main = $doc.AddMainDocumentPart()
        Set-PartXml -Part $main -Xml $DocumentXml
    } finally {
        $doc.Dispose()
    }
}

# --- Word fixture builders --------------------------------------------------

function Build-WordAccessibleBaseline {
    param([string] $Path)
    $body = @(
        Get-WordParagraph -Text 'Document Title' -Style 'Heading1'
        Get-WordParagraph -Text 'Body text.'
        Get-WordParagraph -InnerXml (Get-WordInlineDrawing -Id 1 -Descr 'Company logo')
        Get-WordTable -HasHeader $true
        Get-WordContentControl -Title 'My Field'
    ) -join "`n"
    Build-WordFixture -Path $Path -DocumentXml (Get-WordDocumentXml $body)
}

function Build-WordMissingAltText {
    param([string] $Path)
    $body = @(
        Get-WordParagraph -Text 'Document Title' -Style 'Heading1'
        Get-WordParagraph -Text 'Body text.'
        Get-WordParagraph -InnerXml (Get-WordInlineDrawing -Id 1 -Descr '')
        Get-WordTable -HasHeader $true
        Get-WordContentControl -Title 'My Field'
    ) -join "`n"
    Build-WordFixture -Path $Path -DocumentXml (Get-WordDocumentXml $body)
}

function Build-WordMissingTableHeaders {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '',
        Justification = 'Function name mirrors the MissingTableHeaders rule which is intentionally plural.')]
    param([string] $Path)
    $body = @(
        Get-WordParagraph -Text 'Document Title' -Style 'Heading1'
        Get-WordParagraph -Text 'Body text.'
        Get-WordParagraph -InnerXml (Get-WordInlineDrawing -Id 1 -Descr 'Company logo')
        Get-WordTable -HasHeader $false
        Get-WordContentControl -Title 'My Field'
    ) -join "`n"
    Build-WordFixture -Path $Path -DocumentXml (Get-WordDocumentXml $body)
}

function Build-WordMissingContentControlTitle {
    param([string] $Path)
    $body = @(
        Get-WordParagraph -Text 'Document Title' -Style 'Heading1'
        Get-WordParagraph -Text 'Body text.'
        Get-WordParagraph -InnerXml (Get-WordInlineDrawing -Id 1 -Descr 'Company logo')
        Get-WordTable -HasHeader $true
        Get-WordContentControl -Title ''
    ) -join "`n"
    Build-WordFixture -Path $Path -DocumentXml (Get-WordDocumentXml $body)
}

function Build-WordMergedTableCells {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '',
        Justification = 'Function name mirrors the MergedTableCells rule which is intentionally plural.')]
    param([string] $Path)
    $body = @(
        Get-WordParagraph -Text 'Document Title' -Style 'Heading1'
        Get-WordParagraph -Text 'Body text.'
        Get-WordParagraph -InnerXml (Get-WordInlineDrawing -Id 1 -Descr 'Company logo')
        Get-WordTable -HasHeader $true -HasMerge $true
        Get-WordContentControl -Title 'My Field'
    ) -join "`n"
    Build-WordFixture -Path $Path -DocumentXml (Get-WordDocumentXml $body)
}

function Build-WordHeadingOrderSkip {
    param([string] $Path)
    $body = @(
        Get-WordParagraph -Text 'Section A' -Style 'Heading1'
        Get-WordParagraph -Text 'Subsection A.1' -Style 'Heading3'
        Get-WordParagraph -Text 'Body text.'
        Get-WordParagraph -InnerXml (Get-WordInlineDrawing -Id 1 -Descr 'Company logo')
        Get-WordTable -HasHeader $true
        Get-WordContentControl -Title 'My Field'
    ) -join "`n"
    Build-WordFixture -Path $Path -DocumentXml (Get-WordDocumentXml $body)
}

function Build-WordFloatingObject {
    param([string] $Path)
    # Anchored drawing carries alt text so MissingAltText doesn't co-fire.
    $body = @(
        Get-WordParagraph -Text 'Document Title' -Style 'Heading1'
        Get-WordParagraph -Text 'Body text.'
        Get-WordParagraph -InnerXml (Get-WordAnchorDrawing -Id 2 -Descr 'Decorative banner')
        Get-WordTable -HasHeader $true
        Get-WordContentControl -Title 'My Field'
    ) -join "`n"
    Build-WordFixture -Path $Path -DocumentXml (Get-WordDocumentXml $body)
}

function Build-WordRepeatedBlanks {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '',
        Justification = 'Function name mirrors the RepeatedBlanks rule which is intentionally plural.')]
    param([string] $Path)
    $body = @(
        Get-WordParagraph -Text 'Document Title' -Style 'Heading1'
        Get-WordParagraph -Text 'Hello   world.'  # three spaces
        Get-WordParagraph -InnerXml (Get-WordInlineDrawing -Id 1 -Descr 'Company logo')
        Get-WordTable -HasHeader $true
        Get-WordContentControl -Title 'My Field'
    ) -join "`n"
    Build-WordFixture -Path $Path -DocumentXml (Get-WordDocumentXml $body)
}

function Build-WordNoHeadingStyles {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '',
        Justification = 'Function name mirrors the NoHeadingStyles rule which is intentionally plural.')]
    param([string] $Path)
    # No paragraph carries a Heading style anywhere. Everything else accessible.
    $body = @(
        Get-WordParagraph -Text 'Plain body paragraph one.'
        Get-WordParagraph -Text 'Plain body paragraph two.'
        Get-WordParagraph -InnerXml (Get-WordInlineDrawing -Id 1 -Descr 'Company logo')
        Get-WordTable -HasHeader $true
        Get-WordContentControl -Title 'My Field'
    ) -join "`n"
    Build-WordFixture -Path $Path -DocumentXml (Get-WordDocumentXml $body)
}

function Build-WordLayoutTable {
    param([string] $Path)
    # Table flagged as a layout table via w:tblDescription. First row has no
    # tblHeader, but the description signals "this is for visual arrangement,
    # not tabular data" -- so the MissingTableHeaders rule must skip it.
    $layoutTable = @'
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="4000" w:type="dxa"/>
    <w:tblDescription w:val="Photo grid -- layout only"/>
  </w:tblPr>
  <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>
'@
    $body = @(
        Get-WordParagraph -Text 'Document Title' -Style 'Heading1'
        Get-WordParagraph -Text 'Body text.'
        Get-WordParagraph -InnerXml (Get-WordInlineDrawing -Id 1 -Descr 'Company logo')
        $layoutTable
        Get-WordContentControl -Title 'My Field'
    ) -join "`n"
    Build-WordFixture -Path $Path -DocumentXml (Get-WordDocumentXml $body)
}

function Build-WordLowContrast {
    param([string] $Path)
    # One run with explicit light-grey color (CCCCCC) on explicit white run
    # shading (FFFFFF). Computed contrast ratio is ~1.61:1 -- well below the
    # WCAG 4.5:1 normal-text threshold, so the LowContrast rule fires.
    $lowContrastPara = @'
<w:p>
  <w:r>
    <w:rPr>
      <w:color w:val="CCCCCC"/>
      <w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/>
    </w:rPr>
    <w:t xml:space="preserve">Hard to read</w:t>
  </w:r>
</w:p>
'@
    $body = @(
        Get-WordParagraph -Text 'Document Title' -Style 'Heading1'
        $lowContrastPara
        Get-WordParagraph -InnerXml (Get-WordInlineDrawing -Id 1 -Descr 'Company logo')
        Get-WordTable -HasHeader $true
        Get-WordContentControl -Title 'My Field'
    ) -join "`n"
    Build-WordFixture -Path $Path -DocumentXml (Get-WordDocumentXml $body)
}

# --- Excel: build a workbook by composing parts -----------------------------

# Common namespaces used in Excel parts
$XlsxNs = @{
    s   = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    r   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    xdr = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    a   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
}

# Build workbook.xml referencing one sheet by relationship ID.
function Get-XlsxWorkbookXml {
    param([string] $SheetName = 'Inventory', [string] $SheetRelId = 'rId1')
    @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="$($XlsxNs.s)" xmlns:r="$($XlsxNs.r)">
  <sheets>
    <sheet name="$SheetName" sheetId="1" r:id="$SheetRelId"/>
  </sheets>
</workbook>
"@
}

# Build a worksheet.xml. Optional mergeCells, optional drawing reference, optional
# table reference, optional cell with style index.
function Get-XlsxWorksheetXml {
    param(
        [bool]   $WithMergeCells   = $false,
        [string] $DrawingRelId     = $null,
        [string] $TableRelId       = $null,
        [int]    $StyleIndexForA1  = 0
    )
    $mergePart = if ($WithMergeCells) { '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>' } else { '' }
    $drawingPart = if ($DrawingRelId) { "<drawing r:id=`"$DrawingRelId`"/>" } else { '' }
    $tablePart = if ($TableRelId) { "<tableParts count=`"1`"><tablePart r:id=`"$TableRelId`"/></tableParts>" } else { '' }
    $cellStyle = if ($StyleIndexForA1 -gt 0) { " s=`"$StyleIndexForA1`"" } else { '' }
    @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="$($XlsxNs.s)" xmlns:r="$($XlsxNs.r)">
  <sheetData>
    <row r="1"><c r="A1"$cellStyle><v>-100</v></c><c r="B1" t="inlineStr"><is><t>Header B</t></is></c></row>
    <row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row>
    <row r="3"><c r="A3"><v>3</v></c><c r="B3"><v>4</v></c></row>
  </sheetData>
  $mergePart
  $drawingPart
  $tablePart
</worksheet>
"@
}

function Get-XlsxTableXml {
    param(
        [string] $Name = 'Inventory',
        [string] $DisplayName = 'Inventory',
        [int]    $HeaderRowCount = 1,
        [string] $Ref = 'A1:B3'
    )
    @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="$($XlsxNs.s)" id="1" name="$Name" displayName="$DisplayName" ref="$Ref" totalsRowShown="0" headerRowCount="$HeaderRowCount">
  <autoFilter ref="$Ref"/>
  <tableColumns count="2">
    <tableColumn id="1" name="ColA"/>
    <tableColumn id="2" name="ColB"/>
  </tableColumns>
</table>
"@
}

function Get-XlsxDrawingXml {
    param(
        [string] $Descr = 'Logo',
        [bool]   $Decorative = $false
    )
    $descrAttr = if ($Descr) { " descr=`"$Descr`"" } else { '' }
    $decoAttr  = if ($Decorative) { ' decorative="1"' } else { '' }
    @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="$($XlsxNs.xdr)" xmlns:a="$($XlsxNs.a)" xmlns:r="$($XlsxNs.r)">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="1" name="Picture 1"$descrAttr$decoAttr/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr/>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
"@
}

# Build styles.xml. If $RedOnly is true, includes a custom numFmt "[Red]0.00"
# with id 164 and a cellXf at index 1 referencing it (so a cell with s="1"
# pulls the format into the "used" set the checker requires).
function Get-XlsxStylesXml {
    param([bool] $RedOnly = $false)
    if ($RedOnly) {
        @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="$($XlsxNs.s)">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="[Red]0.00;[Red]0.00"/></numFmts>
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>
"@
    } else {
        @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="$($XlsxNs.s)">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>
"@
    }
}

# Compose an xlsx fixture. The flags drive which optional parts get written.
function Build-XlsxFixture {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [string] $SheetName        = 'Inventory',
        [bool]   $WithMergeCells   = $false,
        [bool]   $WithDrawing      = $true,
        [string] $DrawingDescr     = 'Company logo',
        [bool]   $WithTable        = $true,
        [string] $TableName        = 'Inventory',
        [int]    $TableHeaderRows  = 1,
        [bool]   $RedOnlyNumberFmt = $false
    )
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $type = [DocumentFormat.OpenXml.SpreadsheetDocumentType]::Workbook
    $doc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Create($Path, $type)
    try {
        $wbPart = $doc.AddWorkbookPart()
        $wsPart = Add-OpenXmlPart -Container $wbPart -PartType ([DocumentFormat.OpenXml.Packaging.WorksheetPart]) -RelId 'rId1'
        $stylesPart = Add-OpenXmlPart -Container $wbPart -PartType ([DocumentFormat.OpenXml.Packaging.WorkbookStylesPart]) -RelId 'rId2'

        $tableRelId = $null
        if ($WithTable) {
            $tablePart = Add-OpenXmlPart -Container $wsPart -PartType ([DocumentFormat.OpenXml.Packaging.TableDefinitionPart]) -RelId 'rIdTable'
            $tableRelId = 'rIdTable'
            Set-PartXml -Part $tablePart -Xml (Get-XlsxTableXml -Name $TableName -DisplayName $TableName -HeaderRowCount $TableHeaderRows)
        }

        $drawingRelId = $null
        if ($WithDrawing) {
            $drawingPart = Add-OpenXmlPart -Container $wsPart -PartType ([DocumentFormat.OpenXml.Packaging.DrawingsPart]) -RelId 'rIdDrawing'
            $drawingRelId = 'rIdDrawing'
            Set-PartXml -Part $drawingPart -Xml (Get-XlsxDrawingXml -Descr $DrawingDescr)
        }

        $styleIdx = if ($RedOnlyNumberFmt) { 1 } else { 0 }
        Set-PartXml -Part $wsPart -Xml (Get-XlsxWorksheetXml -WithMergeCells $WithMergeCells -DrawingRelId $drawingRelId -TableRelId $tableRelId -StyleIndexForA1 $styleIdx)
        Set-PartXml -Part $stylesPart -Xml (Get-XlsxStylesXml -RedOnly $RedOnlyNumberFmt)
        Set-PartXml -Part $wbPart -Xml (Get-XlsxWorkbookXml -SheetName $SheetName -SheetRelId 'rId1')
    } finally {
        $doc.Dispose()
    }
}

# --- Excel fixture builders -------------------------------------------------

function Build-XlsxAccessibleBaseline {
    param([string] $Path)
    Build-XlsxFixture -Path $Path
}

function Build-XlsxMissingAltText {
    param([string] $Path)
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $type = [DocumentFormat.OpenXml.SpreadsheetDocumentType]::Workbook
    $doc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Create($Path, $type)
    try {
        $wbPart = $doc.AddWorkbookPart()
        $wsPart = Add-OpenXmlPart -Container $wbPart -PartType ([DocumentFormat.OpenXml.Packaging.WorksheetPart]) -RelId 'rId1'
        $stylesPart = Add-OpenXmlPart -Container $wbPart -PartType ([DocumentFormat.OpenXml.Packaging.WorkbookStylesPart]) -RelId 'rId2'
        $tablePart = Add-OpenXmlPart -Container $wsPart -PartType ([DocumentFormat.OpenXml.Packaging.TableDefinitionPart]) -RelId 'rIdTable'
        $drawingPart = Add-OpenXmlPart -Container $wsPart -PartType ([DocumentFormat.OpenXml.Packaging.DrawingsPart]) -RelId 'rIdDrawing'
        Set-PartXml -Part $tablePart -Xml (Get-XlsxTableXml -Name 'Inventory' -DisplayName 'Inventory' -HeaderRowCount 1)
        Set-PartXml -Part $drawingPart -Xml (Get-XlsxDrawingXml -Descr '')
        Set-PartXml -Part $wsPart -Xml (Get-XlsxWorksheetXml -DrawingRelId 'rIdDrawing' -TableRelId 'rIdTable')
        Set-PartXml -Part $stylesPart -Xml (Get-XlsxStylesXml)
        Set-PartXml -Part $wbPart -Xml (Get-XlsxWorkbookXml -SheetName 'Inventory' -SheetRelId 'rId1')
    } finally {
        $doc.Dispose()
    }
}

function Build-XlsxMissingTableHeaders {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '',
        Justification = 'Function name mirrors the MissingTableHeaders rule which is intentionally plural.')]
    param([string] $Path)
    Build-XlsxFixture -Path $Path -TableHeaderRows 0
}

function Build-XlsxRedOnlyNegativeFormatting {
    param([string] $Path)
    Build-XlsxFixture -Path $Path -RedOnlyNumberFmt $true
}

function Build-XlsxMergedCells {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '',
        Justification = 'Function name mirrors the MergedCells rule which is intentionally plural.')]
    param([string] $Path)
    Build-XlsxFixture -Path $Path -WithMergeCells $true
}

function Build-XlsxDefaultSheetTabName {
    param([string] $Path)
    Build-XlsxFixture -Path $Path -SheetName 'Sheet1'
}

function Build-XlsxDefaultTableName {
    param([string] $Path)
    Build-XlsxFixture -Path $Path -TableName 'Table1'
}

# Styles for the LowContrast fixture: font 1 has explicit grey CCCCCC, fill 2
# is solid white, cellXf 1 references both. Computed contrast ~1.61:1.
function Get-XlsxStylesXmlLowContrast {
    @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="$($XlsxNs.s)">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><sz val="11"/><name val="Calibri"/><color rgb="FFCCCCCC"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
</styleSheet>
"@
}

function Build-XlsxLowContrast {
    param([string] $Path)
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $type = [DocumentFormat.OpenXml.SpreadsheetDocumentType]::Workbook
    $doc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Create($Path, $type)
    try {
        $wbPart = $doc.AddWorkbookPart()
        $wsPart = Add-OpenXmlPart -Container $wbPart -PartType ([DocumentFormat.OpenXml.Packaging.WorksheetPart]) -RelId 'rId1'
        $stylesPart = Add-OpenXmlPart -Container $wbPart -PartType ([DocumentFormat.OpenXml.Packaging.WorkbookStylesPart]) -RelId 'rId2'
        $tablePart = Add-OpenXmlPart -Container $wsPart -PartType ([DocumentFormat.OpenXml.Packaging.TableDefinitionPart]) -RelId 'rIdTable'
        $drawingPart = Add-OpenXmlPart -Container $wsPart -PartType ([DocumentFormat.OpenXml.Packaging.DrawingsPart]) -RelId 'rIdDrawing'

        Set-PartXml -Part $tablePart -Xml (Get-XlsxTableXml -Name 'Inventory' -DisplayName 'Inventory' -HeaderRowCount 1)
        Set-PartXml -Part $drawingPart -Xml (Get-XlsxDrawingXml -Descr 'Logo')
        # Cell A1 uses cellXf 1 -> font 1 (grey) on fill 2 (white).
        Set-PartXml -Part $wsPart -Xml (Get-XlsxWorksheetXml -DrawingRelId 'rIdDrawing' -TableRelId 'rIdTable' -StyleIndexForA1 1)
        Set-PartXml -Part $stylesPart -Xml (Get-XlsxStylesXmlLowContrast)
        Set-PartXml -Part $wbPart -Xml (Get-XlsxWorkbookXml -SheetName 'Inventory' -SheetRelId 'rId1')
    } finally {
        $doc.Dispose()
    }
}

# --- PowerPoint XML fragments ----------------------------------------------
#
# PowerPoint fixtures use a deliberately minimal package: presentation.xml
# referencing one or more slides, plus per-slide XML. The checker walks slides
# directly from the presentation's sldIdLst and never visits the layout/master
# chain, so the fixtures omit those parts.

$PptxNs = @{
    p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

$PptxSlideAttrs = @"
xmlns:p="$($PptxNs.p)" xmlns:a="$($PptxNs.a)" xmlns:r="$($PptxNs.r)"
"@

# Build a presentation.xml referencing the supplied slide relationship IDs in
# order. Slides display in the order their <p:sldId> entries appear here.
function Get-PptxPresentationXml {
    param([string[]] $SlideRelIds)
    $sldId = 256
    $entries = foreach ($rid in $SlideRelIds) {
        "    <p:sldId id=`"$sldId`" r:id=`"$rid`"/>"
        $sldId++
    }
    $sldList = ($entries -join "`n")
    @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation $PptxSlideAttrs>
  <p:sldIdLst>
$sldList
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
"@
}

# Wrap a spTree body in the slide envelope.
function Get-PptxSlideXml {
    param([string] $SpTreeChildren)
    @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld $PptxSlideAttrs>
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
$SpTreeChildren
    </p:spTree>
  </p:cSld>
</p:sld>
"@
}

# Slide envelope with an explicit white background — used by the LowContrast
# fixture so the contrast lookup has a known background to read.
function Get-PptxSlideXmlWhiteBackground {
    param([string] $SpTreeChildren)
    @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld $PptxSlideAttrs>
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
$SpTreeChildren
    </p:spTree>
  </p:cSld>
</p:sld>
"@
}

# Title placeholder shape. $Text = '' means PowerPoint would render the
# layout's "Click to add title" prompt; the checker treats that as missing.
function Get-PptxTitleShape {
    param([int] $Id = 2, [string] $Text = '')
    $body = if ([string]::IsNullOrEmpty($Text)) {
        '<a:p><a:endParaRPr lang="en-US"/></a:p>'
    } else {
        "<a:p><a:r><a:rPr lang=`"en-US`"/><a:t>$Text</a:t></a:r></a:p>"
    }
    @"
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="$Id" name="Title $Id"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          $body
        </p:txBody>
      </p:sp>
"@
}

# Picture shape. $Descr controls the alt text on cNvPr; $ExtLstDecorative
# instead embeds the modern Office <adec:decorative val="1"/> extLst marker
# (omitting descr) — used to test the alt-text rule's recognition of the
# Office 2019/365 "Mark as Decorative" form.
function Get-PptxPictureShape {
    param(
        [int] $Id = 3,
        [string] $Name = "Picture $Id",
        [string] $Descr = '',
        [bool] $ExtLstDecorative = $false
    )
    $descrAttr = if ($Descr) { " descr=`"$Descr`"" } else { '' }
    $extLst = if ($ExtLstDecorative) {
        '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>'
    } else { '' }
    @"
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="$Id" name="$Name"$descrAttr>$extLst</p:cNvPr>
          <p:cNvPicPr/>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill><a:blip/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
        <p:spPr/>
      </p:pic>
"@
}

# Two-column, two-row table inside a graphicFrame. $FirstRow controls the
# tblPr/@firstRow attribute. $Merge produces a horizontally-merged first row.
function Get-PptxTableShape {
    param(
        [int] $Id = 4,
        [bool] $FirstRow = $true,
        [bool] $Merge = $false
    )
    $firstAttr = if ($FirstRow) { ' firstRow="1"' } else { '' }
    $mergeAttrs = if ($Merge) { ' gridSpan="2"' } else { '' }
    $cellB = if ($Merge) { '' }
             else        { '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>B</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>' }
    @"
      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="$Id" name="Table $Id"/>
          <p:cNvGraphicFramePr/>
          <p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="2000000"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tblPr$firstAttr/>
              <a:tblGrid><a:gridCol w="3000000"/><a:gridCol w="3000000"/></a:tblGrid>
              <a:tr h="500000">
                <a:tc$mergeAttrs><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>A</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                $cellB
              </a:tr>
              <a:tr h="500000">
                <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
                <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>2</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
              </a:tr>
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
"@
}

# Body text shape carrying a single hyperlink run. $RelId references a
# hyperlink relationship the caller is responsible for adding to the slide
# part. $Text is the visible text of the linked run.
function Get-PptxLinkShape {
    param(
        [int] $Id = 5,
        [string] $RelId = 'rIdLink1',
        [string] $Text = 'Annual report'
    )
    @"
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="$Id" name="Body $Id"/>
          <p:cNvSpPr/>
          <p:nvPr><p:ph idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US"><a:hlinkClick r:id="$RelId"/></a:rPr>
              <a:t>$Text</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
"@
}

# Shape with explicit white fill and a single text run colored grey (CCCCCC).
# Computed contrast with white is ~1.61:1, well below the 4.5:1 threshold.
# Marked decorative so the alt-text rule does not co-fire.
function Get-PptxLowContrastShape {
    param([int] $Id = 6, [string] $Text = 'Hard to read')
    @"
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="$Id" name="Tile $Id" decorative="1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="800000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US">
                <a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill>
              </a:rPr>
              <a:t>$Text</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
"@
}

# Compose a PowerPoint fixture. $SlideEnvelopes is a list of full <p:sld> XML
# strings (one per slide). $LinkRels (optional) attaches hyperlink
# relationships per slide, indexed by 0-based slide position; each entry is a
# hashtable of @{ RelId = Url }.
#
# IMPORTANT — fixture realism scope:
#
#   These fixtures are deliberately minimal SDK-only packages. They include a
#   PresentationPart and one or more SlidePart(s) but DO NOT include the
#   SlideMaster / SlideLayout / Theme parts that a real PowerPoint file
#   requires. The dispatcher and checker open them via
#   `PresentationDocument.Open` and walk only what the rules need (slide
#   spTree contents and hyperlink relationships), so the missing parts have
#   no behavioural impact on tests.
#
#   Microsoft PowerPoint itself, however, will refuse to open these fixtures
#   ("PowerPoint found a problem with content"). If you want a fixture you
#   can manually inspect in PowerPoint to verify a checker change against
#   the native Accessibility Checker, build it in PowerPoint and copy it
#   into scripts/tests/fixtures/ rather than threading a full
#   master/layout/theme skeleton through this builder.
function Build-PptxFixture {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '',
        Justification = 'Fixture builder mirrors slide-collection input.')]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string[]] $SlideEnvelopes,
        [object[]] $LinkRels = $null
    )
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $type = [DocumentFormat.OpenXml.PresentationDocumentType]::Presentation
    $doc = [DocumentFormat.OpenXml.Packaging.PresentationDocument]::Create($Path, $type)
    try {
        $presPart = $doc.AddPresentationPart()

        $relIds = @()
        for ($i = 0; $i -lt $SlideEnvelopes.Count; $i++) {
            $rid = "rIdSlide{0}" -f ($i + 1)
            $relIds += $rid
            $slidePart = Add-OpenXmlPart -Container $presPart -PartType ([DocumentFormat.OpenXml.Packaging.SlidePart]) -RelId $rid

            Set-PartXml -Part $slidePart -Xml $SlideEnvelopes[$i]

            if ($LinkRels -and $i -lt $LinkRels.Count -and $LinkRels[$i]) {
                foreach ($entry in $LinkRels[$i].GetEnumerator()) {
                    [void] $slidePart.AddHyperlinkRelationship([uri] $entry.Value, $true, $entry.Key)
                }
            }
        }

        Set-PartXml -Part $presPart -Xml (Get-PptxPresentationXml -SlideRelIds $relIds)
    } finally {
        $doc.Dispose()
    }
}

# --- PowerPoint fixture builders -------------------------------------------

function Build-PptxAccessibleBaseline {
    param([string] $Path)
    $sp = @(
        Get-PptxTitleShape   -Id 2 -Text 'Welcome'
        Get-PptxPictureShape -Id 3 -Descr 'Company logo'
        Get-PptxTableShape   -Id 4 -FirstRow $true
        Get-PptxLinkShape    -Id 5 -RelId 'rIdLink1' -Text 'Annual report'
    ) -join "`n"
    Build-PptxFixture -Path $Path `
        -SlideEnvelopes @((Get-PptxSlideXml -SpTreeChildren $sp)) `
        -LinkRels @(@{ 'rIdLink1' = 'https://example.com/annual-report' })
}

function Build-PptxMissingSlideTitle {
    param([string] $Path)
    # Title placeholder present but empty text — checker reports
    # "empty title" rather than "no title placeholder".
    $sp = (Get-PptxTitleShape -Id 2 -Text '')
    Build-PptxFixture -Path $Path -SlideEnvelopes @((Get-PptxSlideXml -SpTreeChildren $sp))
}

function Build-PptxMissingAltText {
    param([string] $Path)
    $sp = @(
        Get-PptxTitleShape   -Id 2 -Text 'Welcome'
        Get-PptxPictureShape -Id 3 -Descr ''
    ) -join "`n"
    Build-PptxFixture -Path $Path -SlideEnvelopes @((Get-PptxSlideXml -SpTreeChildren $sp))
}

function Build-PptxMissingTableHeaders {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '',
        Justification = 'Fixture function name mirrors the MissingTableHeaders rule which is intentionally plural.')]
    param([string] $Path)
    $sp = @(
        Get-PptxTitleShape -Id 2 -Text 'Inventory'
        Get-PptxTableShape -Id 4 -FirstRow $false
    ) -join "`n"
    Build-PptxFixture -Path $Path -SlideEnvelopes @((Get-PptxSlideXml -SpTreeChildren $sp))
}

function Build-PptxDuplicateSlideTitle {
    param([string] $Path)
    $envelopes = @(
        Get-PptxSlideXml -SpTreeChildren (Get-PptxTitleShape -Id 2 -Text 'Overview')
        Get-PptxSlideXml -SpTreeChildren (Get-PptxTitleShape -Id 2 -Text 'Details')
        Get-PptxSlideXml -SpTreeChildren (Get-PptxTitleShape -Id 2 -Text 'Overview')
    )
    Build-PptxFixture -Path $Path -SlideEnvelopes $envelopes
}

function Build-PptxMergedTableCells {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '',
        Justification = 'Fixture function name mirrors the MergedTableCells rule which is intentionally plural.')]
    param([string] $Path)
    $sp = @(
        Get-PptxTitleShape -Id 2 -Text 'Roster'
        Get-PptxTableShape -Id 4 -FirstRow $true -Merge $true
    ) -join "`n"
    Build-PptxFixture -Path $Path -SlideEnvelopes @((Get-PptxSlideXml -SpTreeChildren $sp))
}

function Build-PptxNonDescriptiveLink {
    param([string] $Path)
    $sp = @(
        Get-PptxTitleShape -Id 2 -Text 'Resources'
        Get-PptxLinkShape  -Id 5 -RelId 'rIdLink1' -Text 'click here'
    ) -join "`n"
    Build-PptxFixture -Path $Path `
        -SlideEnvelopes @((Get-PptxSlideXml -SpTreeChildren $sp)) `
        -LinkRels @(@{ 'rIdLink1' = 'https://example.com/resources' })
}

function Build-PptxExtLstDecorative {
    param([string] $Path)
    # Picture has no descr/title/@decorative attribute; the alt-text rule must
    # recognize the modern extLst <adec:decorative> marker as satisfying.
    $sp = @(
        Get-PptxTitleShape   -Id 2 -Text 'Welcome'
        Get-PptxPictureShape -Id 3 -Descr '' -ExtLstDecorative $true
    ) -join "`n"
    Build-PptxFixture -Path $Path -SlideEnvelopes @((Get-PptxSlideXml -SpTreeChildren $sp))
}

function Build-PptxLowContrast {
    param([string] $Path)
    $sp = @(
        Get-PptxTitleShape       -Id 2 -Text 'Status'
        Get-PptxLowContrastShape -Id 6 -Text 'Hard to read'
    ) -join "`n"
    Build-PptxFixture -Path $Path -SlideEnvelopes @((Get-PptxSlideXmlWhiteBackground -SpTreeChildren $sp))
}

# --- Drive every builder ----------------------------------------------------

$builders = @(
    @{ Name = 'word-accessible-baseline.docx';            Build = { param($p) Build-WordAccessibleBaseline           -Path $p } }
    @{ Name = 'word-missing-alt-text.docx';               Build = { param($p) Build-WordMissingAltText                -Path $p } }
    @{ Name = 'word-missing-table-headers.docx';          Build = { param($p) Build-WordMissingTableHeaders           -Path $p } }
    @{ Name = 'word-missing-content-control-title.docx';  Build = { param($p) Build-WordMissingContentControlTitle    -Path $p } }
    @{ Name = 'word-merged-table-cells.docx';             Build = { param($p) Build-WordMergedTableCells              -Path $p } }
    @{ Name = 'word-heading-order-skip.docx';             Build = { param($p) Build-WordHeadingOrderSkip              -Path $p } }
    @{ Name = 'word-floating-object.docx';                Build = { param($p) Build-WordFloatingObject                -Path $p } }
    @{ Name = 'word-repeated-blanks.docx';                Build = { param($p) Build-WordRepeatedBlanks                -Path $p } }
    @{ Name = 'word-no-heading-styles.docx';              Build = { param($p) Build-WordNoHeadingStyles               -Path $p } }
    @{ Name = 'word-low-contrast.docx';                   Build = { param($p) Build-WordLowContrast                   -Path $p } }
    @{ Name = 'word-layout-table.docx';                   Build = { param($p) Build-WordLayoutTable                   -Path $p } }
    @{ Name = 'excel-accessible-baseline.xlsx';           Build = { param($p) Build-XlsxAccessibleBaseline            -Path $p } }
    @{ Name = 'excel-missing-alt-text.xlsx';              Build = { param($p) Build-XlsxMissingAltText                -Path $p } }
    @{ Name = 'excel-missing-table-headers.xlsx';         Build = { param($p) Build-XlsxMissingTableHeaders           -Path $p } }
    @{ Name = 'excel-red-only-negative-formatting.xlsx';  Build = { param($p) Build-XlsxRedOnlyNegativeFormatting     -Path $p } }
    @{ Name = 'excel-merged-cells.xlsx';                  Build = { param($p) Build-XlsxMergedCells                   -Path $p } }
    @{ Name = 'excel-default-sheet-tab-name.xlsx';        Build = { param($p) Build-XlsxDefaultSheetTabName           -Path $p } }
    @{ Name = 'excel-default-table-name.xlsx';            Build = { param($p) Build-XlsxDefaultTableName              -Path $p } }
    @{ Name = 'excel-low-contrast.xlsx';                  Build = { param($p) Build-XlsxLowContrast                   -Path $p } }
    @{ Name = 'powerpoint-accessible-baseline.pptx';      Build = { param($p) Build-PptxAccessibleBaseline            -Path $p } }
    @{ Name = 'powerpoint-missing-slide-title.pptx';      Build = { param($p) Build-PptxMissingSlideTitle             -Path $p } }
    @{ Name = 'powerpoint-missing-alt-text.pptx';         Build = { param($p) Build-PptxMissingAltText                -Path $p } }
    @{ Name = 'powerpoint-missing-table-headers.pptx';    Build = { param($p) Build-PptxMissingTableHeaders           -Path $p } }
    @{ Name = 'powerpoint-duplicate-slide-title.pptx';    Build = { param($p) Build-PptxDuplicateSlideTitle           -Path $p } }
    @{ Name = 'powerpoint-merged-table-cells.pptx';       Build = { param($p) Build-PptxMergedTableCells              -Path $p } }
    @{ Name = 'powerpoint-non-descriptive-link.pptx';     Build = { param($p) Build-PptxNonDescriptiveLink            -Path $p } }
    @{ Name = 'powerpoint-low-contrast.pptx';             Build = { param($p) Build-PptxLowContrast                   -Path $p } }
    @{ Name = 'powerpoint-extlst-decorative.pptx';        Build = { param($p) Build-PptxExtLstDecorative              -Path $p } }
)

foreach ($b in $builders) {
    $target = Join-Path $OutputDir $b.Name
    & $b.Build $target
    Write-Information "Built $($b.Name)"
}

Write-Information ("Done. {0} fixture(s) written to {1}" -f $builders.Count, $OutputDir)
