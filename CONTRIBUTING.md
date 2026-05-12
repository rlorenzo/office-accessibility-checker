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

Requires PowerShell 7+. Microsoft maintains [install instructions for Windows, macOS, and Linux](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell).

```powershell
# Download the Open XML SDK into scripts/lib/ (idempotent)
./scripts/setup-accessibility-checker.ps1

# Run the test suite
./scripts/tests/Invoke-Tests.ps1
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
./scripts/lint.ps1                  # PSScriptAnalyzer
./scripts/tests/Invoke-Tests.ps1    # Pester suite
```

Both run in CI on every push and pull request via [.github/workflows/ci.yml](.github/workflows/ci.yml).

## Pull request conventions

- One logical change per PR.
- Commit messages in the imperative mood ("Add MissingAltText fixture", not "Added").
- Reference issue numbers in the PR description when relevant.
- Keep PRs focused: rule changes that ship without updated fixtures and manifest entries will not pass CI.

## Releasing to PowerShell Gallery

The published module lives under [module/OfficeAccessibilityChecker/](module/OfficeAccessibilityChecker/). The hand-written surface (`.psd1` manifest, `.psm1` loader, `Public/`) is committed; the `Private/` folder is a build artifact mirrored from `scripts/` and is gitignored.

Publishing is automated by [.github/workflows/publish.yml](.github/workflows/publish.yml): pushing a `v*` tag triggers a job that re-assembles the module and uploads it to the Gallery. The standard release flow:

1. **Bump `ModuleVersion`** in [OfficeAccessibilityChecker.psd1](module/OfficeAccessibilityChecker/OfficeAccessibilityChecker.psd1). Open a PR, merge after CI is green.
2. **Tag the merge commit on `main`** with the matching version:
   ```powershell
   git checkout main && git pull
   git tag v1.x.x && git push origin v1.x.x
   ```
3. The `Publish to PowerShell Gallery` workflow runs, verifies that the tag matches the manifest, and runs `Publish-Module`. Watch the run under the **Actions** tab.

The workflow fails fast if the tag and `ModuleVersion` disagree — that's the guard against "tagged v1.1.0 but forgot to bump the manifest."

### Manual fallback

If the workflow is broken or you need to push without a tag, the same flow runs locally:

```powershell
./scripts/Build-Module.ps1
Import-Module ./module/OfficeAccessibilityChecker -Force
Test-OfficeAccessibility ./scripts/tests/fixtures/word-accessible-baseline.docx
Publish-Module -Path ./module/OfficeAccessibilityChecker -NuGetApiKey $env:PSGALLERY_API_KEY
```

### Required GitHub configuration

A one-time setup, already done for this repo but documented for forks:

- Create a GitHub environment named `powershell-gallery` (**Settings → Environments**).
- Add a **Deployment branches and tags** rule restricted to tags matching `v*`. This means only a `v*` tag push can access the secret.
- Add a repository-scoped or environment-scoped `PSGALLERY_API_KEY` secret with a Gallery API key that has push rights to `OfficeAccessibilityChecker`.
