@{
    RootModule        = 'OfficeAccessibilityChecker.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = '68562f32-bc7c-4d75-b931-177916720dbe'
    Author            = 'Rex Lorenzo'
    Copyright         = '(c) Rex Lorenzo. MIT License.'
    Description       = 'Cross-platform CLI accessibility checker for Word, Excel, and PowerPoint files. Implements Microsoft''s Accessibility Checker rules directly against OOXML using the Open XML SDK. No Office install required. Supports bulk scans and -Fix to write a sibling .fixed.<ext> with safe, structural remediations applied.'

    PowerShellVersion = '7.0'

    FunctionsToExport = @(
        'Test-OfficeAccessibility'
        'Initialize-OfficeAccessibilityChecker'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    PrivateData = @{
        PSData = @{
            Tags         = @(
                'accessibility', 'a11y', 'wcag', 'office', 'ooxml',
                'word', 'excel', 'powerpoint', 'docx', 'xlsx', 'pptx',
                'compliance', 'cross-platform'
            )
            LicenseUri   = 'https://github.com/rlorenzo/office-accessibility-checker/blob/main/LICENSE'
            ProjectUri   = 'https://github.com/rlorenzo/office-accessibility-checker'
            ReleaseNotes = 'See https://github.com/rlorenzo/office-accessibility-checker/releases for the full changelog.'
        }
    }
}
