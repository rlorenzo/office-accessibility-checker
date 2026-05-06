# Security policy

## Reporting a vulnerability

If you find a security issue, please report it privately rather than opening a public issue.

- Preferred: [GitHub private vulnerability reporting](https://github.com/rlorenzo/office-accessibility-checker/security/advisories/new).
- Email: rexlorenzo@gmail.com.

Please include reproduction steps and, if possible, a sample file that triggers the behavior. You will get an initial response within a reasonable timeframe; a fix will be coordinated before public disclosure.

## Threat model

`office-accessibility-checker` parses Office Open XML packages (`.docx`, `.docm`, `.xlsx`, `.xlsm`) on the user's machine using the [Open XML SDK](https://www.nuget.org/packages/documentformat.openxml). Inputs may be untrusted.

The tool:

- Reads OOXML packages in read-only mode.
- Does not execute macros, formulas, or scripts contained in the file.
- Does not render content (no Word, Excel, or browser engine is invoked).
- Does not write to or modify the file being checked.
- Does not make network calls during a check. The one-time setup script downloads the SDK from nuget.org.

The most likely class of vulnerability is a malformed or maliciously crafted OOXML package that causes the SDK to crash, hang, or consume excessive memory while parsing. Reports of any such input are welcome.

## Supported versions

Only the latest commit on `main` is supported. There are no tagged releases yet.
