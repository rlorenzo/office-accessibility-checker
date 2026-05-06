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
                'RepeatedBlanks', 'NoHeadingStyles'
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

        # --- Excel: accessible baseline ------------------------------------
        @{
            File           = 'excel-accessible-baseline.xlsx'
            ExpectedExit   = 0
            MustContain    = @()
            MustNotContain = @(
                'MissingAltText', 'MissingTableHeaders', 'RedOnlyNegativeFormatting',
                'MergedCells', 'DefaultSheetTabName', 'DefaultTableName'
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
        }
    )
}
