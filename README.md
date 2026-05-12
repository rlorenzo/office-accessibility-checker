# office-accessibility-checker

[![CI](https://github.com/rlorenzo/office-accessibility-checker/actions/workflows/ci.yml/badge.svg)](https://github.com/rlorenzo/office-accessibility-checker/actions/workflows/ci.yml)
[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/OfficeAccessibilityChecker?logo=powershell&label=PSGallery)](https://www.powershellgallery.com/packages/OfficeAccessibilityChecker)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![GitHub Sponsors](https://img.shields.io/github/sponsors/rlorenzo?label=Sponsor&logo=GitHub)](https://github.com/sponsors/rlorenzo)
[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)](https://learn.microsoft.com/en-us/powershell/)
![Cross-platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-blue)

A command-line tool that checks Word, Excel, and PowerPoint files for accessibility problems — the same kinds of issues Microsoft's built-in Accessibility Checker reports. No copy of Office required. Runs on Windows, macOS, and Linux.

## Requirements

Windows PowerShell 5.1 or PowerShell 7+. Windows 10/11 ships 5.1 preinstalled, so nothing to install there. On macOS and Linux, install PowerShell 7 — Microsoft maintains [install instructions](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell).

## Install

**Via PowerShell Gallery** (recommended):

```powershell
Install-Module OfficeAccessibilityChecker
Initialize-OfficeAccessibilityChecker        # one-time SDK download
Test-OfficeAccessibility path/to/file.docx
# PASS path/to/file.docx
```

**Or, from a clone:**

```powershell
git clone https://github.com/rlorenzo/office-accessibility-checker.git
cd office-accessibility-checker
./scripts/setup-accessibility-checker.ps1
./scripts/check-office-accessibility.ps1 path/to/file.docx
```

The rest of this README uses the script form (`./scripts/check-office-accessibility.ps1`). Module users can substitute `Test-OfficeAccessibility` — the parameters are identical.

## Checking a file

By default you get a single line per file:

```powershell
./scripts/check-office-accessibility.ps1 report.docx
# FAIL report.docx: MissingAltText, MissingTableHeaders
```

For the full list of issues, add `-Format detailed`:

```powershell
./scripts/check-office-accessibility.ps1 report.docx -Format detailed
# ERROR    MissingAltText        Image "chart" has no alt text
# WARNING  MergedTableCells      Table 1 contains merged cells
# FAIL report.docx: MissingAltText
```

Issues come in three categories: **ERROR** (likely blocks people with disabilities — causes `FAIL`), **WARNING** (hard to use), and **TIP** (room to improve). Only ERRORs cause the script to exit non-zero, so you can wire it into CI without warnings tripping a build.

## Checking many files at once

Pass a folder instead of a file. Subfolders are not searched unless you add `-Recurse`.

```powershell
./scripts/check-office-accessibility.ps1 path/to/folder -Recurse
# PASS folder/report.docx
# FAIL folder/budget.xlsx: MissingAltText
# PASS folder/memo.docx
```

A summary footer (`Scanned N files: X passed, …`) is written to standard error. A single corrupt file produces an `ERROR` line and the scan keeps going.

## Fixing what it can

Add `-Fix` and a sibling file named `<original>.fixed.<ext>` is written with safe, automatic repairs applied. **Your original file is never modified.**

```powershell
./scripts/check-office-accessibility.ps1 report.docx -Fix
# FAIL report.docx: MissingTableHeaders, LowContrast
# FIXED report.fixed.docx: MissingTableHeaders (2), LowContrast (5)
# PASS report.fixed.docx
```

Only fixes that don't drastically change how the document looks are applied:

| Format | Issue | What gets fixed |
|---|---|---|
| Word | Missing table headers | Marks the first row of each table as a header. |
| Word | Repeated blanks | Replaces 3+ spaces with a tab (Microsoft's recommended remediation). |
| Word, Excel, PowerPoint | Low contrast text | Nudges the text color to the closest shade that meets WCAG contrast — same hue, just a bit darker or lighter. |
| Excel | Missing table headers | Re-enables the header row Excel disabled. |
| Excel | Red-only negative numbers | Adds a leading `-` so negatives are also distinguishable for color-blind readers. |
| PowerPoint | Missing table headers | Turns on the table's header-row option. |

Things that need human judgement — writing alt text, naming a slide, picking a meaningful link label, unmerging cells — are not autofixed.

## What it checks

### Word
| Severity | Issue | What it looks for |
|---|---|---|
| ERROR | Missing alt text | Images, charts, and shapes without alt text or a "decorative" mark. |
| ERROR | Missing table headers | Tables whose first row isn't marked as a header. (Layout-only tables are exempt.) |
| ERROR | Missing content control title | Form fields with no title set. |
| ERROR | Document protected | File is password- or rights-protected. |
| WARNING | Merged table cells | Tables with merged or nested cells. |
| WARNING | Heading order skip | Heading levels jump (e.g. Heading 1 to Heading 3). |
| WARNING | Floating object | Images that aren't inline with text. |
| WARNING | Repeated blanks | Three or more spaces in a row (use tabs instead). |
| WARNING | Low contrast | Text whose color contrast falls below WCAG thresholds. |
| TIP | No heading styles | Document has no headings at all. |

### Excel
| Severity | Issue | What it looks for |
|---|---|---|
| ERROR | Missing alt text | Pictures, charts, shapes without alt text. |
| ERROR | Missing table headers | Excel table with the header row disabled. |
| ERROR | Red-only negative formatting | Number formats that show negatives in red but with no `-` or parentheses. |
| ERROR | Document protected | Workbook is password- or rights-protected. |
| WARNING | Merged cells | Sheets containing merged cells. |
| WARNING | Default sheet tab name | Sheet still named "Sheet1", "Tabelle1", "Hoja1", etc. |
| WARNING | Low contrast | Cells whose font/fill colors fall below 4.5:1 contrast. |
| TIP | Default table name | Table still named "Table1", "Table2", … |

### PowerPoint
| Severity | Issue | What it looks for |
|---|---|---|
| ERROR | Missing alt text | Images, charts, shapes without alt text. (Tables are exempt.) |
| ERROR | Missing slide title | Slide has no title or a blank title. |
| ERROR | Missing table headers | Table without a header row enabled. |
| ERROR | Document protected | File is password- or rights-protected. |
| WARNING | Duplicate slide title | Two or more slides share the same title. |
| WARNING | Merged table cells | Tables with merged cells. |
| WARNING | Low contrast | Text whose color contrast falls below 4.5:1. |
| WARNING | Non-descriptive link text | Hyperlinks that read "click here", "more", the URL itself, etc. |

For exact behavior — including OOXML-level details, intentional exemptions, and best-effort caveats — see [`docs/RULES.md`](docs/RULES.md). Microsoft's full rule reference lives [here][rules].

## A few things to know

- **Contrast checking is best-effort.** It only catches problems where both colors are explicitly set. Theme colors and inherited styles are skipped.
- **Macro-enabled files** (`.docm`/`.xlsm`/`.pptm`) work the same way as their non-macro versions.
- **Legacy `.doc`/`.xls`/`.ppt` files are not supported** — re-save them as the modern format first.

## Exit codes

| Code | Meaning |
|---|---|
| 0 | No errors found. |
| 1 | One or more accessibility errors. |
| 2 | Tool error (file missing, unsupported format, etc.). |

## Development

```powershell
./scripts/tests/Invoke-Tests.ps1   # run the test suite
./scripts/lint.ps1                  # run the linter
```

Tests and fixtures live under [`scripts/tests/`](scripts/tests/). The technical rule catalog and autofix design notes are in [`docs/RULES.md`](docs/RULES.md). See [CONTRIBUTING.md](CONTRIBUTING.md) for the developer workflow.

To use a different copy of the Open XML SDK, set `$env:OPENXML_SDK_PATH`.

## License

MIT. See [LICENSE](LICENSE). Security disclosures: [SECURITY.md](SECURITY.md).

[rules]: https://support.microsoft.com/en-us/office/rules-for-the-accessibility-checker-651e08f2-0fc3-4e10-aaca-74b4a67101c1
[sdk]: https://www.nuget.org/packages/documentformat.openxml
