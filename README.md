# office-accessibility-checker

[![CI](https://github.com/rlorenzo/office-accessibility-checker/actions/workflows/ci.yml/badge.svg)](https://github.com/rlorenzo/office-accessibility-checker/actions/workflows/ci.yml)

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

## Tests

A manifest-driven Pester suite lives under [scripts/tests/](scripts/tests/). It runs each checker against committed accessible/inaccessible Word and Excel fixtures and asserts the exit code + reported rules match [manifest.psd1](scripts/tests/fixtures/manifest.psd1).

```powershell
# Idempotent: fetches the Open XML SDK + Pester 5 if missing, then runs the suite.
.\scripts\tests\Invoke-Tests.ps1

# Or, if you've already installed Pester 5+ yourself:
Invoke-Pester scripts\tests
```

Fixtures are committed binaries under [scripts/tests/fixtures/](scripts/tests/fixtures/). They're produced by [Build-Fixtures.ps1](scripts/tests/Build-Fixtures.ps1), which writes minimal OOXML packages directly via the Open XML SDK; treat the script as a one-shot regenerator -- run it after editing a rule or fixture builder, then commit the updated binaries in the same PR.

The `DocumentProtected` rule has no fixture: it fires when the package is wrapped in an encrypted Compound File envelope, which the SDK doesn't write. Coverage is intentionally deferred -- if you need it, hand-craft a password-encrypted file and add a manifest entry.

CI runs the lint + test jobs on every push and PR via [.github/workflows/ci.yml](.github/workflows/ci.yml).

## Linting

```powershell
.\scripts\lint.ps1
```

Runs PSScriptAnalyzer against `scripts\` using the rules in [PSScriptAnalyzerSettings.psd1](PSScriptAnalyzerSettings.psd1). Exits non-zero on findings at severity ≥ Warning, so it's safe to wire into pre-commit hooks. CI runs the same script.

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
  lint.ps1                          # PSScriptAnalyzer wrapper
  lib/                              # gitignored; populated by setup script
  tests/
    AccessibilityChecker.Tests.ps1  # manifest-driven Pester suite
    Build-Fixtures.ps1              # one-shot regenerator (run when rules change)
    Invoke-Tests.ps1                # local convenience wrapper
    fixtures/
      manifest.psd1                 # filename → expected exit + rule names
      *.docx, *.xlsx                # committed; rebuild via Build-Fixtures.ps1
.github/workflows/ci.yml            # lint + test on windows-latest
```
