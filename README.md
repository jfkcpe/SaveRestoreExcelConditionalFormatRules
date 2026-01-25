# SaveRestoreExcelConditionalFormatRules
VBA Macros to Save most current Conditional Format Rules as plain, human-readable text, then Restore them from that string.

Updated Jan 24 2026 - Clean up & condense code, normalize string variables, fix bug in font color.
Updated Jan 23 2026 - Added a lot more capability, moved utility functions to a separate module.

The Save subroutine will attempt to save the existing Conditional Formatting rules for the current worksheet.
The rules are saved as a compact plain-text string in a cell somewhere in the workbook (e.g., on some "ReadMe" tab or something).
That cell must be Named <Tabname>_CF_RULES (e.g., Data_CF_RULES)
Utility functions are in a separate Module

The Restore subroutine below will read that string and attempt to recreate all the rules for the worksheet from scratch.
Currently, there are probably some conditions and formats that are not saved (or restored).
Formats Handled: Interior Color (including Color Scale), Font Color/Bold/Italic and some Borders.
Formats Not Yet Handled: Fill Patterns, .
So far, limited testing indicates successfully handling of condition types:
  xlExpression
  xlCellValue (Operators: xlBetween, xlNotBetween, xlEqual,)
  xlUniqueValues (including DupeUnique or "Duplicates")
  xlTextString (TextOperators: xlContains, xlDoesNotContain, xlBeginsWith, xlEndsWith)
  xlColorScale (2 & 3 colors), xlBlanksCondition, xlNoBlanksCondition, xlTimePeriod (all DateOperators probably)
  xlTop10 (so far only top N %)
Condition Types Not Yet Tested (thus, perhaps not handled):
  xlIconSet, xlErrorsCondition, xlNoErrorsCondition

Note about programming style: this is not efficient or elegant code.  It enumerates each and every parameter
  that I discovered through experimentation and debugging (and is no doubt missing some options).
I suppose I could have written (or obtained) some code to programmatically and exhaustively render (and restore) each ConditionalFormattingRule
  as some kind ofhierarchical, plain-text dictionary object with all the potentially unnecessary parameters.
  But, I wanted the saved string to be uncluttered and easy to read and understand.
  I also enjoyed the rather exhausting project of unpacking and trying to understand the small portion of the CF Rules that I managed to handle.
  I'm a little inconsistent with naming of String vs. Long (and Boolean) data types.  I'll clean that up some day.
