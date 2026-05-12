# Rule reference

Detailed, OOXML-level documentation for every rule the checker implements. The README has the user-facing summary; this file is for contributors, integrators, or anyone trying to reason about a specific finding.

For Microsoft's authoritative rule list, see the [Accessibility Checker reference][rules]. We mirror that rule set, with documented exceptions noted below.

[rules]: https://support.microsoft.com/en-us/office/rules-for-the-accessibility-checker-651e08f2-0fc3-4e10-aaca-74b4a67101c1

## Word (`check-docx-accessibility.ps1`)

| Severity | Rule | What it checks |
|---|---|---|
| ERROR | `MissingAltText` | Every drawing in the body, headers, footers, and groups has alt text, a title, or `decorative="1"`. |
| ERROR | `MissingTableHeaders` | First row carries `w:trPr/w:tblHeader` (the semantic table-header marker). Visual first-row styling via `tblLook` is intentionally **not** accepted, since accepting it would let tables that look like they have headers pass while exposing nothing to screen readers. Tables marked as layout-only via `w:tblPr/w:tblDescription` are exempt. |
| ERROR | `MissingContentControlTitle` | Every `w:sdt` has a non-empty `w:alias`. |
| ERROR | `DocumentProtected` | File is IRM- or password-protected (Restrict Editing is **not** flagged). |
| WARNING | `MergedTableCells` | Tables with merged or nested cells. |
| WARNING | `HeadingOrderSkip` | Heading levels skip (e.g. Heading1 to Heading3). |
| WARNING | `FloatingObject` | Drawings using `wp:anchor` (not inline). |
| WARNING | `RepeatedBlanks` | 3+ consecutive spaces or non-breaking spaces. Tabs are intentionally **not** flagged (Microsoft's own remediation advice is to use tabs). |
| WARNING | `LowContrast` | Best-effort. Flags runs whose foreground (`w:rPr/w:color`) and background (run or paragraph `w:shd/@w:fill`) are both explicit hex values and whose WCAG contrast ratio falls below 4.5:1 (3:1 for 18pt+ text or 14pt+ bold). Theme references, "auto" colors, and style/theme inheritance are skipped. |
| TIP | `NoHeadingStyles` | Document has no Heading-style paragraphs anywhere. |

## Excel (`check-xlsx-accessibility.ps1`)

| Severity | Rule | What it checks |
|---|---|---|
| ERROR | `MissingAltText` | Every drawing (pictures, charts, shapes, group shapes) across all anchor types has alt text, a title, or `decorative="1"`. |
| ERROR | `MissingTableHeaders` | `<table>` does not have `headerRowCount="0"`. |
| ERROR | `RedOnlyNegativeFormatting` | Best-effort detection of `[Red]` numFmt patterns lacking complementary minus/parens. |
| ERROR | `DocumentProtected` | Workbook is IRM- or password-protected (workbook/sheet protection is **not** flagged). |
| WARNING | `MergedCells` | Worksheet has any `<mergeCells>` entries (sheet-wide check, not table-scoped). |
| WARNING | `DefaultSheetTabName` | Sheet tab matches a default placeholder (`Sheet1`, `Tabelle1`, `Feuil1`, `Hoja1`, `Foglio1`, `Planilha1`, `シート1`, etc.). Add locales by editing the regex array at the top of the script. |
| WARNING | `LowContrast` | Best-effort. Flags cells whose font color and cell fill are both explicit RGB values and whose WCAG contrast ratio falls below 4.5:1. Theme references, indexed-palette colors, "auto", non-solid fills, and conditional formatting are skipped. |
| TIP | `DefaultTableName` | Table name matches `^Table\d+$`. |

## PowerPoint (`check-pptx-accessibility.ps1`)

| Severity | Rule | What it checks |
|---|---|---|
| ERROR | `MissingAltText` | Every non-placeholder shape, picture, chart/SmartArt graphic frame, connector, and group child has alt text, a title, or `decorative="1"`. Tables (graphic frames containing `a:tbl`) are exempt because their cells already expose accessible text. |
| ERROR | `MissingSlideTitle` | Every slide has a title placeholder (`p:ph/@type` ∈ {`title`, `ctrTitle`}) with non-empty text. Triggers both for slides with no title placeholder and for slides whose placeholder text is blank. |
| ERROR | `MissingTableHeaders` | Every `a:tbl` has `a:tblPr/@firstRow="1"` (the table-style header-row marker). |
| ERROR | `DocumentProtected` | File is IRM- or password-protected. |
| WARNING | `DuplicateSlideTitle` | Two or more slides share the same title text (case-insensitive, trimmed). |
| WARNING | `MergedTableCells` | Tables with cells using `gridSpan>1`, `rowSpan>1`, `hMerge="1"`, or `vMerge="1"`. |
| WARNING | `LowContrast` | Best-effort. Flags runs whose foreground (`a:rPr/a:solidFill/a:srgbClr`) and background (owning shape's `p:spPr` solid fill, falling back to the slide's `p:cSld/p:bg` solid fill) are both explicit RGB and whose WCAG contrast ratio falls below 4.5:1. Theme/scheme colors, gradients, and inherited fills are skipped. |
| WARNING | `NonDescriptiveLinkText` | Run-level hyperlinks (`a:rPr/a:hlinkClick`) whose visible text is empty, equals the URL, or matches a generic phrase such as `click here`, `here`, `more`, `read more`, `link`, `this link`. |

## Known limitations

- **Contrast.** Best-effort only. The check requires both foreground and background to be explicit RGB values; runs/cells whose colors come from theme references, "auto", indexed palettes, style inheritance, or conditional formatting are skipped because resolving them faithfully requires rendering. Expect false negatives for documents that rely on themed colors.
- **Red-only number format.** Detection is best-effort against common `[Red]` patterns. Custom numFmt edge cases may slip through.
- **Localized default sheet names.** The locale list lives in `check-xlsx-accessibility.ps1` as a single named constant; add a regex to extend.
- **Macro-enabled formats** (`.docm`, `.xlsm`, `.pptm`) are treated identically to their non-macro siblings, since macros are irrelevant to OOXML structural checks.
- **`DocumentProtected` has no fixture** because creating an encrypted Compound File envelope from scratch isn't something the SDK does; coverage is intentionally deferred.

## Autofix

`-Fix` writes a sibling `<basename>.fixed.<ext>` with a small set of deterministic structural repairs applied. The original file is never modified. After applying fixes the checker recursively re-runs against the fixed file (without `-Fix`, so no infinite recursion) and prints its PASS/FAIL line. The exit code reflects the fixed file when fixes were applied, otherwise the original.

In bulk-scan mode `-Fix` is forwarded per file; previously-produced `*.fixed.<ext>` outputs are excluded from the file list so re-scans don't generate `*.fixed.fixed.<ext>`.

### Scope

| Format | Rule | OOXML edit |
|---|---|---|
| Word | `MissingTableHeaders` | Inserts `<w:tblHeader/>` into the first row's `<w:trPr>`. Layout tables (carrying `w:tblPr/w:tblDescription`) are skipped. |
| Word | `RepeatedBlanks` | Replaces runs of 3+ U+0020 / U+00A0 inside a single `<w:t>` with one `\t`, ensuring `xml:space="preserve"` is set on the modified element. Cross-`<w:t>` runs are intentionally not handled (would require splitting and re-stitching runs). |
| Word, Excel, PowerPoint | `LowContrast` | Rewrites the foreground color to the perceptually-closest hex (algorithm below). Only runs/cells whose foreground and background are both already explicit RGB are touched — same bar the rule uses. |
| Excel | `MissingTableHeaders` | Removes `headerRowCount="0"` from the table element. Excel's default of 1 then re-treats the first row as a header. |
| Excel | `RedOnlyNegativeFormatting` | Inserts a literal `-` immediately after the `[Red]` marker in the negative section of the format code. Other bracket conditionals before `[Red]` are preserved. |
| PowerPoint | `MissingTableHeaders` | Sets `a:tblPr/@firstRow="1"` (creating `a:tblPr` if absent). |

### What is intentionally not autofixed

- **Authoring**: alt text, slide titles, content control titles, descriptive link text, default sheet/table names, headings (`NoHeadingStyles`).
- **Destructive**: unmerging cells (Word/Excel/PowerPoint), removing protection, anchor → inline (changes layout).
- **Subjective design**: heading-level remapping, contrast color shifts on theme/indexed colors.
- **Anything that would require lying to assistive tech**: blanket `decorative="1"`, fabricated alt text, placeholder content control titles.

The bar is "deterministic structural OOXML edit whose visible effect is either invisible or expected as a consequence of the recommended remediation." A surprising number of rules don't qualify.

### Low-contrast color algorithm

For each run/cell that fails contrast, find the perceptually-closest replacement foreground that meets the WCAG 2.x threshold (3:1 for large text in Word; 4.5:1 elsewhere):

1. Convert the foreground hex to HSL.
2. Hold hue and saturation fixed; binary-walk lightness in both directions (down toward 0, up toward 1) at step 0.005, computing contrast against the unchanged background at each step.
3. Stop on the first L that meets the threshold in each direction.
4. Pick the candidate with the smaller `|L_new − L_orig|`. If neither direction reaches the threshold within `[0, 1]` (a mid-tone background with no passing color along the lightness axis), fall back to whichever pole — black or white — has more contrast headroom.

LCh would be more perceptually uniform than HSL, but HSL is good enough for "nudge the color until it passes" and keeps the per-run cost trivial.

## Project layout

```
scripts/
  setup-accessibility-checker.ps1   # one-time SDK download
  check-office-accessibility.ps1    # dispatcher (.docx / .xlsx / .pptx -> right checker)
  check-docx-accessibility.ps1      # Word rules + fixes
  check-xlsx-accessibility.ps1      # Excel rules + fixes
  check-pptx-accessibility.ps1      # PowerPoint rules + fixes
  Build-Module.ps1                  # assembles module/OfficeAccessibilityChecker for publish
  lint.ps1                          # PSScriptAnalyzer wrapper
  lib/                              # gitignored; populated by setup script
  tests/
    AccessibilityChecker.Tests.ps1  # manifest-driven Pester suite
    Module.Tests.ps1                # PowerShell Gallery module smoke tests
    Build-Fixtures.ps1              # one-shot regenerator (run when rules change)
    Invoke-Tests.ps1                # local convenience wrapper
    fixtures/
      manifest.psd1                 # filename -> expected exit + rule names
      *.docx, *.xlsx, *.pptx        # committed; rebuild via Build-Fixtures.ps1
module/
  OfficeAccessibilityChecker/       # PowerShell Gallery package
    OfficeAccessibilityChecker.psd1 # manifest
    OfficeAccessibilityChecker.psm1 # loader
    Public/                         # exported cmdlets
    Private/                        # build artifact: gitignored, mirrors scripts/
docs/
  RULES.md                          # this document
.github/workflows/ci.yml            # lint (Windows) + Pester matrix on Windows/macOS/Linux
```
