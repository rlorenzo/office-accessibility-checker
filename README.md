# office-accessibility-checker

CLI accessibility checker for Office Open XML documents (`.docx`/`.docm`,
`.xlsx`/`.xlsm`). Mirrors veraPDF's invocation shape: feed it a file, get a
PASS/FAIL line and a meaningful exit code.

Implements Microsoft's published [Accessibility Checker rules][rules]
directly against OOXML using the [Open XML SDK][sdk]. No Office install
required — works headless on any machine with PowerShell.

[rules]: https://support.microsoft.com/en-us/office/rules-for-the-accessibility-checker-651e08f2-0fc3-4e10-aaca-74b4a67101c1
[sdk]: https://www.nuget.org/packages/documentformat.openxml

## Setup (one-time)

```powershell
.\scripts\setup-accessibility-checker.ps1
```

Downloads the Open XML SDK DLLs into `scripts\lib\` (gitignored). Idempotent
— re-runs are no-ops unless `-Force` is passed.

## Usage

```powershell
# Text mode (default) — single PASS/FAIL line
.\scripts\check-office-accessibility.ps1 path\to\file.docx
# PASS path\to\file.docx

# Detailed mode — every issue, one per line, then PASS/FAIL summary
.\scripts\check-office-accessibility.ps1 path\to\file.xlsx -Format detailed
# WARNING  MergedCells          Sheet "Sheet1" contains merged cells
# TIP      ContrastSkipped      Contrast check skipped (requires rendering)
# PASS path\to\file.xlsx
```

The dispatcher picks the right checker based on file extension. You can
also call `check-docx-accessibility.ps1` or `check-xlsx-accessibility.ps1`
directly with the same parameters.

## Exit codes

| Code | Meaning |
|------|---------|
| 0    | No accessibility errors. Warnings/tips do not fail. |
| 1    | One or more accessibility errors found. Includes IRM/password-protected files. |
| 2    | Tool error (file not found, unsupported format, SDK missing, file locked). |

## Supported formats

| Extension | Handled by |
|-----------|------------|
| `.docx`, `.docm` | `check-docx-accessibility.ps1` |
| `.xlsx`, `.xlsm` | `check-xlsx-accessibility.ps1` |
| `.doc`, `.xls`   | **Not supported** — legacy binary formats. Re-save as `.docx`/`.xlsx`. Exits 2. |

## Rules implemented

### Word

| Severity | Rule | What it checks |
|----------|------|----------------|
| ERROR    | `MissingAltText` | Every drawing in body, headers, footers, and groups has alt text, a title, or `decorative="1"`. |
| ERROR    | `MissingTableHeaders` | First row carries `w:trPr/w:tblHeader` (the semantic table-header marker). Visual first-row styling via `tblLook` is intentionally **not** accepted — it is style metadata, not header semantics, so accepting it would let tables that look like they have headers pass while exposing nothing to screen readers. |
| ERROR    | `MissingContentControlTitle` | Every `w:sdt` has a non-empty `w:alias`. |
| ERROR    | `DocumentProtected` | File is IRM- or password-protected (Restrict Editing is **not** flagged). |
| WARNING  | `MergedTableCells` | Tables with merged or nested cells. |
| WARNING  | `HeadingOrderSkip` | Heading levels skip (e.g., Heading1 → Heading3). |
| WARNING  | `FloatingObject` | Drawings using `wp:anchor` (not inline). |
| WARNING  | `RepeatedBlanks` | 3+ consecutive spaces or non-breaking spaces. Tabs are intentionally **not** flagged (Microsoft's own remediation advice is to use tabs). |
| TIP      | `NoHeadingStyles` | Document has no Heading-style paragraphs anywhere. |
| TIP      | `ContrastSkipped` | Always emitted; contrast can't be checked without rendering. |

### Excel

| Severity | Rule | What it checks |
|----------|------|----------------|
| ERROR    | `MissingAltText` | Every drawing — pictures, charts, shapes, group shapes — across all anchor types has alt text, a title, or `decorative="1"`. |
| ERROR    | `MissingTableHeaders` | `<table>` does not have `headerRowCount="0"`. |
| ERROR    | `RedOnlyNegativeFormatting` | Best-effort detection of `[Red]` numFmt patterns lacking complementary minus/parens. |
| ERROR    | `DocumentProtected` | Workbook is IRM- or password-protected (workbook/sheet protection is **not** flagged). |
| WARNING  | `MergedCells` | Worksheet has any `<mergeCells>` entries — sheet-wide check, not table-scoped. |
| WARNING  | `DefaultSheetTabName` | Sheet tab matches a default placeholder (`Sheet1`, `Tabelle1`, `Feuil1`, `Hoja1`, `Foglio1`, `Planilha1`, `シート1`, etc.). Add locales by editing the regex array at the top of the script. |
| TIP      | `DefaultTableName` | Table name matches `^Table\d+$`. |
| TIP      | `ContrastSkipped` | Always emitted. |

## Known limitations

- **Contrast.** Cannot be checked without rendering — a single TIP records this.
- **Red-only number format.** Detection is best-effort against common `[Red]` patterns. Custom numFmt edge cases may slip through.
- **Localized default sheet names.** The locale list lives in `check-xlsx-accessibility.ps1` as a single named constant; add a regex to extend.
- **Macro-enabled formats** (`.docm`, `.xlsm`) are treated identically to their non-macro siblings — macros are irrelevant to OOXML structural checks.
- **Tests / fixtures.** Pester suite and hand-crafted fixtures (planned in §10 of the design doc) are out-of-scope follow-up work.

## Override the SDK location

```powershell
$env:OPENXML_SDK_PATH = 'C:\path\to\DocumentFormat.OpenXml.dll'
```

If set and the file exists, this takes precedence over the bundled `scripts\lib\` copy.

## Files

```
scripts/
  setup-accessibility-checker.ps1   # one-time SDK download
  check-office-accessibility.ps1    # dispatcher (.docx / .xlsx → right checker)
  check-docx-accessibility.ps1      # Word rules
  check-xlsx-accessibility.ps1      # Excel rules
  lib/                              # gitignored; populated by setup script
```
