# Contributing

Thanks for your interest in improving `office-accessibility-checker`. This guide covers reporting issues, setting up locally, and the conventions for changing rules or fixtures.

## Reporting issues

[Open a GitHub issue](https://github.com/rlorenzo/office-accessibility-checker/issues) and include:

- OS and PowerShell version (`$PSVersionTable.PSVersion`).
- The command you ran and its full output.
- A sample file that reproduces the problem, ideally minimized. If the file is sensitive, attach a sanitized excerpt or describe the structure (drawings, tables, headers, etc.) instead.
- What you expected to happen and what actually happened.

For suspected security issues, see [SECURITY.md](SECURITY.md) instead.

## Local setup

Requires PowerShell 7+ on Windows.

```powershell
# Download the Open XML SDK into scripts\lib\ (idempotent)
.\scripts\setup-accessibility-checker.ps1

# Run the test suite
.\scripts\tests\Invoke-Tests.ps1
```

`Invoke-Tests.ps1` installs Pester 5 if missing, then runs the manifest-driven Pester suite at [scripts/tests/](scripts/tests/).

## Adding or changing a rule

1. Edit the relevant checker: [check-docx-accessibility.ps1](scripts/check-docx-accessibility.ps1) or [check-xlsx-accessibility.ps1](scripts/check-xlsx-accessibility.ps1).
2. Add a fixture-builder branch in [Build-Fixtures.ps1](scripts/tests/Build-Fixtures.ps1) that produces both an accessible (passing) and inaccessible (failing) sample for the rule.
3. Regenerate fixtures by running `Build-Fixtures.ps1`. Treat it as a one-shot regenerator.
4. Update [scripts/tests/fixtures/manifest.psd1](scripts/tests/fixtures/manifest.psd1) so the new fixtures map to expected exit codes and rule names.
5. Update the rules tables in [README.md](README.md). The README is the source of truth for what's checked; keep it in sync.
6. Commit the regenerated binaries in the same PR as the rule change.

## Running checks before pushing

```powershell
.\scripts\lint.ps1                  # PSScriptAnalyzer
.\scripts\tests\Invoke-Tests.ps1    # Pester suite
```

Both run in CI on every push and pull request via [.github/workflows/ci.yml](.github/workflows/ci.yml).

## Pull request conventions

- One logical change per PR.
- Commit messages in the imperative mood ("Add MissingAltText fixture", not "Added").
- Reference issue numbers in the PR description when relevant.
- Keep PRs focused: rule changes that ship without updated fixtures and manifest entries will not pass CI.
