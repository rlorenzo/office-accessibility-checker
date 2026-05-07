@{
    # Each entry maps a fixture filename to its expected checker output.
    #
    # Fields:
    #   File           -- filename relative to scripts/tests/fixtures/
    #   ExpectedExit   -- required exit code (0 = no errors, 1 = errors found, 2 = tool error)
    #   MustContain    -- rule names that MUST appear in `-Format detailed` output
    #   MustNotContain -- rule names that MUST NOT appear (catches accidental over-triggering)
    #   Notes          -- short reminder of what the fixture targets
    #
    # Fixtures are committed binaries. Build-Fixtures.ps1 is the one-shot
    # regenerator -- run it when you change a rule, then commit the updated
    # binaries.
    Fixtures = @(
        # --- Word: accessible baseline -------------------------------------
        @{
            File           = 'word-accessible-baseline.docx'
            ExpectedExit   = 0
            MustContain    = @()
            MustNotContain = @(
                'MissingAltText', 'MissingTableHeaders', 'MissingContentControlTitle',
                'MergedTableCells', 'HeadingOrderSkip', 'FloatingObject',
                'RepeatedBlanks', 'NoHeadingStyles', 'LowContrast'
            )
            Notes          = 'Heading 1, alt-tagged inline image, table with repeating header row, content control with title.'
        },

        # --- Word: per-rule inaccessible fixtures --------------------------
        @{
            File           = 'word-missing-alt-text.docx'
            ExpectedExit   = 1
            MustContain    = @('MissingAltText')
            MustNotContain = @()
            Notes          = 'One inline image with alt text deleted (Picture Format > Alt Text > clear).'
        },
        @{
            File           = 'word-missing-table-headers.docx'
            ExpectedExit   = 1
            MustContain    = @('MissingTableHeaders')
            MustNotContain = @()
            Notes          = 'Table without "Repeat as header row at the top of each page" set on the first row.'
        },
        @{
            File           = 'word-missing-content-control-title.docx'
            ExpectedExit   = 1
            MustContain    = @('MissingContentControlTitle')
            MustNotContain = @()
            Notes          = 'Plain Text content control inserted via Developer tab, with empty Title in Properties.'
        },
        @{
            File           = 'word-merged-table-cells.docx'
            ExpectedExit   = 0
            MustContain    = @('MergedTableCells')
            MustNotContain = @()
            Notes          = 'Table with two cells merged horizontally (Layout > Merge Cells).'
        },
        @{
            File           = 'word-heading-order-skip.docx'
            ExpectedExit   = 0
            MustContain    = @('HeadingOrderSkip')
            MustNotContain = @()
            Notes          = 'Heading 1 immediately followed by Heading 3 (skips Heading 2).'
        },
        @{
            File           = 'word-floating-object.docx'
            ExpectedExit   = 0
            MustContain    = @('FloatingObject')
            MustNotContain = @()
            Notes          = 'Image with wrap set to anything other than "In Line with Text" (e.g. Square).'
        },
        @{
            File           = 'word-repeated-blanks.docx'
            ExpectedExit   = 0
            MustContain    = @('RepeatedBlanks')
            MustNotContain = @()
            Notes          = 'A paragraph containing three or more consecutive spaces.'
        },
        @{
            File           = 'word-no-heading-styles.docx'
            ExpectedExit   = 0
            MustContain    = @('NoHeadingStyles')
            MustNotContain = @()
            Notes          = 'Document with body text only -- no Heading 1/2/3 styles applied anywhere.'
        },
        @{
            File           = 'word-low-contrast.docx'
            ExpectedExit   = 0
            MustContain    = @('LowContrast')
            MustNotContain = @()
            Notes          = 'One run with explicit light-grey color on white run shading (~1.6:1 contrast).'
        },
        @{
            File           = 'word-layout-table.docx'
            ExpectedExit   = 0
            MustContain    = @()
            MustNotContain = @('MissingTableHeaders')
            Notes          = 'Table marked as layout via w:tblDescription with no header row -- MissingTableHeaders must skip it.'
        },

        # --- Excel: accessible baseline ------------------------------------
        @{
            File           = 'excel-accessible-baseline.xlsx'
            ExpectedExit   = 0
            MustContain    = @()
            MustNotContain = @(
                'MissingAltText', 'MissingTableHeaders', 'RedOnlyNegativeFormatting',
                'MergedCells', 'DefaultSheetTabName', 'DefaultTableName', 'LowContrast'
            )
            Notes          = 'Sheet renamed away from "Sheet1", named table "Inventory" with header row, alt-tagged image.'
        },

        # --- Excel: per-rule inaccessible fixtures -------------------------
        @{
            File           = 'excel-missing-alt-text.xlsx'
            ExpectedExit   = 1
            MustContain    = @('MissingAltText')
            MustNotContain = @()
            Notes          = 'Picture inserted with alt text deleted (Picture Format > Alt Text > clear).'
        },
        @{
            File           = 'excel-missing-table-headers.xlsx'
            ExpectedExit   = 1
            MustContain    = @('MissingTableHeaders')
            MustNotContain = @()
            Notes          = 'Insert > Table with "My table has headers" UNCHECKED, so headerRowCount="0".'
        },
        @{
            File           = 'excel-red-only-negative-formatting.xlsx'
            ExpectedExit   = 1
            MustContain    = @('RedOnlyNegativeFormatting')
            MustNotContain = @()
            Notes          = 'Custom number format like [Red]0.00 with no minus or parens for the negative branch.'
        },
        @{
            File           = 'excel-merged-cells.xlsx'
            ExpectedExit   = 0
            MustContain    = @('MergedCells')
            MustNotContain = @()
            Notes          = 'Two adjacent cells merged via Home > Merge & Center.'
        },
        @{
            File           = 'excel-default-sheet-tab-name.xlsx'
            ExpectedExit   = 0
            MustContain    = @('DefaultSheetTabName')
            MustNotContain = @()
            Notes          = 'Sheet tab left at the default "Sheet1".'
        },
        @{
            File           = 'excel-default-table-name.xlsx'
            ExpectedExit   = 0
            MustContain    = @('DefaultTableName')
            MustNotContain = @()
            Notes          = 'Insert > Table accepting the auto-assigned "Table1" name.'
        },
        @{
            File           = 'excel-low-contrast.xlsx'
            ExpectedExit   = 0
            MustContain    = @('LowContrast')
            MustNotContain = @()
            Notes          = 'Cell A1 styled with explicit light-grey font on solid white fill (~1.6:1 contrast).'
        },

        # --- PowerPoint: accessible baseline -------------------------------
        @{
            File           = 'powerpoint-accessible-baseline.pptx'
            ExpectedExit   = 0
            MustContain    = @()
            MustNotContain = @(
                'MissingAltText', 'MissingSlideTitle', 'MissingTableHeaders',
                'DuplicateSlideTitle', 'MergedTableCells', 'NonDescriptiveLinkText',
                'LowContrast'
            )
            Notes          = 'Title shape with text, alt-tagged picture, table with header row, descriptively-labeled hyperlink.'
        },

        # --- PowerPoint: per-rule inaccessible fixtures --------------------
        @{
            File           = 'powerpoint-missing-slide-title.pptx'
            ExpectedExit   = 1
            MustContain    = @('MissingSlideTitle')
            MustNotContain = @()
            Notes          = 'Slide carries a title placeholder shape but the placeholder text is empty.'
        },
        @{
            File           = 'powerpoint-missing-alt-text.pptx'
            ExpectedExit   = 1
            MustContain    = @('MissingAltText')
            MustNotContain = @()
            Notes          = 'Picture inserted with alt text deleted (Picture Format > Alt Text > clear).'
        },
        @{
            File           = 'powerpoint-missing-table-headers.pptx'
            ExpectedExit   = 1
            MustContain    = @('MissingTableHeaders')
            MustNotContain = @()
            Notes          = 'Table inserted with the "Header Row" table-style option turned off.'
        },
        @{
            File           = 'powerpoint-duplicate-slide-title.pptx'
            ExpectedExit   = 0
            MustContain    = @('DuplicateSlideTitle')
            MustNotContain = @()
            Notes          = 'Three-slide deck where the first and third slide share the same title text.'
        },
        @{
            File           = 'powerpoint-merged-table-cells.pptx'
            ExpectedExit   = 0
            MustContain    = @('MergedTableCells')
            MustNotContain = @()
            Notes          = 'Table with two cells merged horizontally in the first row (gridSpan="2").'
        },
        @{
            File           = 'powerpoint-non-descriptive-link.pptx'
            ExpectedExit   = 0
            MustContain    = @('NonDescriptiveLinkText')
            MustNotContain = @()
            Notes          = 'Hyperlink whose visible text is "click here" rather than something descriptive.'
        },
        @{
            File           = 'powerpoint-low-contrast.pptx'
            ExpectedExit   = 0
            MustContain    = @('LowContrast')
            MustNotContain = @()
            Notes          = 'Shape with explicit white fill containing a single light-grey-coloured text run (~1.6:1).'
        },
        @{
            File           = 'powerpoint-extlst-decorative.pptx'
            ExpectedExit   = 0
            MustContain    = @()
            MustNotContain = @('MissingAltText')
            Notes          = 'Picture marked decorative via the modern Office extLst marker, not the legacy @decorative attribute.'
        }
    )
}
