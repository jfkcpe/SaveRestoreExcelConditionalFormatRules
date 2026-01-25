# SaveRestoreExcelConditionalFormatRules
VBA Macros to Save most current Conditional Format Rules as plain, human-readable text, then Restore them from that string.

## Updates:

-   Updated Jan 24 2026 - Clean up & condense code, normalize string variables, fix bug in font color, deeper dive into borders.
-   Updated Jan 23 2026 - Added a lot more capability, moved utility functions to a separate module.

## About:

My Problem: My *Conditional Formatting Rules* often get all munged up with dozens of duplicate rules and fragmented *Applies To* ranges as I grow my worksheets adding and moving columns and generally playing around. This might be some flaw in my approach to excel, but I’ve seen others complain about it also.

I tried the simple solution of recording myself setting up all my formatting rules from scratch and saving that as a very specific (and brittle) macro. But it became difficult whenever I made changes and had to either re-record all the rules or mess around with the recorded commands in the macro.

I looked for some way to save and restore *Conditional Formatting Rules*, but my Word 365 didn’t appear to have that capability. I know there are probably plug-ins to handle CF Rules, but I decided to take on the challenge as a learning exercise and share it if I got anywhere.

The Save subroutine will attempt to save the existing *Conditional Formatting Rules* for the current worksheet.

The rules are saved as a compact plain-text string in a cell somewhere in the workbook (e.g., on some "ReadMe" tab or something).

That cell must be Named \<Tabname\>_CF_RULES (e.g., Data_CF_RULES)

Utility functions are in a separate Module

The Restore subroutine will read that string and attempt to recreate all the rules for the worksheet from scratch.

## Note about programming style:

I am not a developer and this is not efficient or elegant code. It enumerates each and every parameter that I discovered through experimentation and debugging (and is no doubt missing some options).

I suppose I could have written (or obtained) some code to programmatically and exhaustively render each ConditionalFormattingRule as some kind of hierarchical, plain-text listing with all the parameters (including many potentially unnecessary ones) and then rebuild the Object during Restore.

But, I wanted the saved string to be uncluttered and easy to read and understand. I also enjoyed the rather exhausting project of unpacking and trying to understand the small portion of the CF Rules that I managed to handle.

## Limitations:

Currently, there are probably some conditions and formats that are not saved (or restored).

Formats Handled: Interior Color (including Color Scale), Font Color/Bold/Italic and some Borders.

Formats Not Yet Handled: Fill Patterns,

So far, limited testing indicates successfully handling of condition types:

-   xlExpression
-   xlCellValue (Operators: xlBetween, xlNotBetween, xlEqual,)
-   xlUniqueValues (including DupeUnique or "Duplicates")
-   xlTextString (TextOperators: xlContains, xlDoesNotContain, xlBeginsWith, xlEndsWith)
-   xlColorScale (2 & 3 colors), xlBlanksCondition, xlNoBlanksCondition, xlTimePeriod (all DateOperators probably)
-   xlTop10 (so far only top N %)

Condition Types Not Yet Tested (thus, perhaps not handled):

-   xlIconSet
-   xlErrorsCondition
-   xlNoErrorsCondition

## Operational Notes:

1.  To Install, go to Developer -\> Visual Basic. Right click on Left Nav “Modules” and select “Import File”
2.  Do that for two files: “SaveRestore.bas” and “SaveRestoreUtil.bas”
3.  Assuming you want to save and restore Conditional Formatting Rules for your tab “MyData”… Somewhere on your workbook (possibly on a “Readme” or “Misc” tab) select a cell and Name it (try clicking in the pull-down just under the left side of the menu bar and type the name).  
    The name should be in the form \<TabName\>_CF_RULES. An example would be: MyData_CF_RULES

## Test Case:

In my repository there is a file called TestCase.txt which is a recent copy of the saved rules (from my WatchedHistory tv/movie bingeing spreadsheet). The rules are mostly arbitrary – I just wanted to have a test case with a lot of different rules.
