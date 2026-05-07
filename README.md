# office-accessibility-checker

[![CI](https://github.com/rlorenzo/office-accessibility-checker/actions/workflows/ci.yml/badge.svg)](https://github.com/rlorenzo/office-accessibility-checker/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![GitHub Sponsors](https://img.shields.io/github/sponsors/rlorenzo?label=Sponsor&logo=GitHub)](https://github.com/sponsors/rlorenzo)
[![PowerShell 7+](https://img.shields.io/badge/PowerShell-7%2B-blue?logo=powershell)](https://learn.microsoft.com/en-us/powershell/)
![Platform: Windows](https://img.shields.io/badge/Platform-Windows-blue?logo=windows)

CLI accessibility checker for Office Open XML documents (`.docx`/`.docm`, `.xlsx`/`.xlsm`). Checks Microsoft's published [Accessibility Checker rules][rules] directly against OOXML using the [Open XML SDK][sdk]. Headless. No Office install required.

[rules]: https://support.microsoft.com/en-us/office/rules-for-the-accessibility-checker-651e08f2-0fc3-4e10-aaca-74b4a67101c1
[sdk]: https://www.nuget.org/packages/documentformat.openxml

## Quick start

```powershell
.\scripts\setup-accessibility-checker.ps1
.\scripts\check-office-accessibility.ps1 path\to\file.docx
# PASS path\to\file.docx
```

The setup script downloads the Open XML SDK DLLs into `scripts\lib\` (gitignored). It is idempotent; re-runs are no-ops unless `-Force` is passed.

## Usage

Text mode (default) emits a single PASS/FAIL line. Detailed mode emits every issue found, one per line, followed by a PASS/FAIL summary.

```powershell
.\scripts\check-office-accessibility.ps1 path\to\file.xlsx -Format detailed
# WARNING  MergedCells          Sheet "Sheet1" contains merged cells
# TIP      ContrastSkipped      Contrast check skipped (requires rendering)
# PASS path\to\file.xlsx
```

The dispatcher picks the right checker based on file extension. You can also call `check-docx-accessibility.ps1` or `check-xlsx-accessibility.ps1` directly with the same parameters.

### Bulk scan

Pass a directory instead of a file to scan many files in one invocation. Supported extensions (`.docx`, `.docm`, `.xlsx`, `.xlsm`) are picked up; everything else is silently skipped. Office lock files (`~$*`) are ignored.

```powershell
.\scripts\check-office-accessibility.ps1 path\to\folder
# PASS path\to\folder\report.docx
# FAIL path\to\folder\budget.xlsx
# PASS path\to\folder\memo.docx
```

Subdirectories are not descended into by default. Add `-Recurse` to walk the tree:

```powershell
.\scripts\check-office-accessibility.ps1 path\to\folder -Recurse -Format detailed
```

A summary footer is written to **stderr** (`Scanned N files: X passed, Y failed, Z errors, W skipped`), so stdout stays clean for piping. A single corrupt or unreadable file does not halt the scan; it produces an `ERROR <path>` line on stdout and the scan continues. The aggregate exit code is the worst per-file exit (0 if all passed, 1 if any accessibility errors, 2 if any tool errors).

## Exit codes

| Code | Meaning |
|------|---------|
| 0    | No accessibility errors. Warnings and tips do not fail. |
| 1    | One or more accessibility errors found. Includes IRM/password-protected files. |
| 2    | Tool error (file not found, unsupported format, SDK missing, file locked). |

## Supported formats

| Extension | Handled by |
|-----------|------------|
| `.docx`, `.docm` | `check-docx-accessibility.ps1` |
| `.xlsx`, `.xlsm` | `check-xlsx-accessibility.ps1` |
| `.doc`, `.xls`   | **Not supported.** Legacy binary formats; re-save as `.docx`/`.xlsx`. Exits 2. |

## Rules implemented

### Word

| Severity | Rule | What it checks |
|----------|------|----------------|
| ERROR    | `MissingAltText` | Every drawing in body, headers, footers, and groups has alt text, a title, or `decorative="1"`. |
| ERROR    | `MissingTableHeaders` | First row carries `w:trPr/w:tblHeader` (the semantic table-header marker). Visual first-row styling via `tblLook` is intentionally **not** accepted, since accepting it would let tables that look like they have headers pass while exposing nothing to screen readers. |
| ERROR    | `MissingContentControlTitle` | Every `w:sdt` has a non-empty `w:alias`. |
| ERROR    | `DocumentProtected` | File is IRM- or password-protected (Restrict Editing is **not** flagged). |
| WARNING  | `MergedTableCells` | Tables with merged or nested cells. |
| WARNING  | `HeadingOrderSkip` | Heading levels skip (e.g., Heading1 to Heading3). |
| WARNING  | `FloatingObject` | Drawings using `wp:anchor` (not inline). |
| WARNING  | `RepeatedBlanks` | 3+ consecutive spaces or non-breaking spaces. Tabs are intentionally **not** flagged (Microsoft's own remediation advice is to use tabs). |
| TIP      | `NoHeadingStyles` | Document has no Heading-style paragraphs anywhere. |
| TIP      | `ContrastSkipped` | Always emitted; contrast cannot be checked without rendering. |

### Excel

| Severity | Rule | What it checks |
|----------|------|----------------|
| ERROR    | `MissingAltText` | Every drawing (pictures, charts, shapes, group shapes) across all anchor types has alt text, a title, or `decorative="1"`. |
| ERROR    | `MissingTableHeaders` | `<table>` does not have `headerRowCount="0"`. |
| ERROR    | `RedOnlyNegativeFormatting` | Best-effort detection of `[Red]` numFmt patterns lacking complementary minus/parens. |
| ERROR    | `DocumentProtected` | Workbook is IRM- or password-protected (workbook/sheet protection is **not** flagged). |
| WARNING  | `MergedCells` | Worksheet has any `<mergeCells>` entries (sheet-wide check, not table-scoped). |
| WARNING  | `DefaultSheetTabName` | Sheet tab matches a default placeholder (`Sheet1`, `Tabelle1`, `Feuil1`, `Hoja1`, `Foglio1`, `Planilha1`, `シート1`, etc.). Add locales by editing the regex array at the top of the script. |
| TIP      | `DefaultTableName` | Table name matches `^Table\d+$`. |
| TIP      | `ContrastSkipped` | Always emitted. |

## Known limitations

- **Contrast.** Cannot be checked without rendering. A single TIP records this.
- **Red-only number format.** Detection is best-effort against common `[Red]` patterns. Custom numFmt edge cases may slip through.
- **Localized default sheet names.** The locale list lives in `check-xlsx-accessibility.ps1` as a single named constant; add a regex to extend.
- **Macro-enabled formats** (`.docm`, `.xlsm`) are treated identically to their non-macro siblings, since macros are irrelevant to OOXML structural checks.

## Development

### Running tests

A manifest-driven Pester suite lives under [scripts/tests/](scripts/tests/). It runs each checker against committed accessible/inaccessible Word and Excel fixtures and asserts that the exit code and reported rules match [manifest.psd1](scripts/tests/fixtures/manifest.psd1).

```powershell
# Idempotent: fetches the Open XML SDK + Pester 5 if missing, then runs the suite.
.\scripts\tests\Invoke-Tests.ps1

# Or, if you have already installed Pester 5+ yourself:
Invoke-Pester scripts\tests
```

Fixtures are committed binaries under [scripts/tests/fixtures/](scripts/tests/fixtures/), produced by [Build-Fixtures.ps1](scripts/tests/Build-Fixtures.ps1). See [CONTRIBUTING.md](CONTRIBUTING.md) for the workflow when changing rules or regenerating fixtures.

The `DocumentProtected` rule has no fixture: it fires when the package is wrapped in an encrypted Compound File envelope, which the SDK does not write. Coverage is intentionally deferred. To exercise it, hand-craft a password-encrypted file and add a manifest entry.

CI runs the lint and test jobs on every push and pull request via [.github/workflows/ci.yml](.github/workflows/ci.yml).

### Linting

```powershell
.\scripts\lint.ps1
```

Runs PSScriptAnalyzer against `scripts\` using the rules in [PSScriptAnalyzerSettings.psd1](PSScriptAnalyzerSettings.psd1). Exits non-zero on findings at severity >= Warning, so it is safe to wire into pre-commit hooks. CI runs the same script.

### Overriding the SDK location

```powershell
$env:OPENXML_SDK_PATH = 'C:\path\to\DocumentFormat.OpenXml.dll'
```

If set and the file exists, this takes precedence over the bundled `scripts\lib\` copy.

### Project layout

```
scripts/
  setup-accessibility-checker.ps1   # one-time SDK download
  check-office-accessibility.ps1    # dispatcher (.docx / .xlsx -> right checker)
  check-docx-accessibility.ps1      # Word rules
  check-xlsx-accessibility.ps1      # Excel rules
  lint.ps1                          # PSScriptAnalyzer wrapper
  lib/                              # gitignored; populated by setup script
  tests/
    AccessibilityChecker.Tests.ps1  # manifest-driven Pester suite
    Build-Fixtures.ps1              # one-shot regenerator (run when rules change)
    Invoke-Tests.ps1                # local convenience wrapper
    fixtures/
      manifest.psd1                 # filename -> expected exit + rule names
      *.docx, *.xlsx                # committed; rebuild via Build-Fixtures.ps1
.github/workflows/ci.yml            # lint + test on windows-latest
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). For security issues, see [SECURITY.md](SECURITY.md).

## License

MIT. See [LICENSE](LICENSE).
