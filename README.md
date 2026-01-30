# SaveRestoreExcelConditionalFormatRules
VBA Macros to Save most current Conditional Format Rules as plain, human-readable text, then Recreate them from that string.

## Updates:

-   Updated Jan 30 2026 – ver 0.4
    -   Merged the “Util” module of utility/helper functions back into the main SaveRestore.bas file, so you won’t see any reference to the SaveRestoreUtil.bas file anymore.
    -   Add “Column(s) Affected:” string showing column header for each Rule in saved string
    -   Updated Borders logic, made code more concise with extended use of arrays.
    -   Added xlIconSet support (whew!)
    -   Reversed Excel’s defaults for some programmatically created rules on StopIfTrue – our default will be False, just like Excel’s default for MANUALLY created rules.
-   Updated Jan 24 2026 - Clean up & condense code, normalize string variables, fix bug in font color, deeper dive into borders.
-   Updated Jan 23 2026 - Added a lot more capability, moved utility functions to a separate module.

## About:

My Problem: My *Conditional Formatting Rules* often get all munged up with dozens of duplicate rules and fragmented *Applies To* ranges as I grow my worksheets by copy/pasting cells (esPECIALLY this!), inserting or moving columns, and generally playing around. This might be some flaw in my approach to excel, but I’ve seen others complain about it also.

I tried the simple solution of recording myself setting up all my formatting rules from scratch and saving that as a very specific (and brittle) macro. But it became difficult whenever I made changes and had to either re-record all the rules or mess around with the recorded commands in the macro.

I looked for some way to save and later restore *Conditional Formatting Rules*, but my Word 365 didn’t appear to have that capability. I know there are probably plug-ins to handle CF Rules, but I decided to take on the challenge as a learning exercise and share it if I got anywhere.

The Save subroutine will attempt to save the existing *Conditional Formatting Rules* for the current worksheet.

The rules are saved as a compact plain-text string in a cell somewhere in the workbook (e.g., on some "ReadMe" tab or something – See “Operational Notes” below.)

The Recreate subroutine will read that string and attempt to recreate all the rules for the worksheet from scratch. (And there are several utility or helper functions at the bottom.)

## Note about programming style:

I am not a developer and this is not efficient or elegant code. It enumerates each and every parameter that I deduced would be useful through trial-and-error combined with debugging and inspection of the Excel Object Model (and we are, no doubt, missing some options).

I suppose I could have written (or obtained) some code to programmatically and exhaustively render each ConditionalFormattingRule as some kind of hierarchical, plain-text listing (perhaps JSON?) with all the parameters (including many potentially unnecessary ones) and then rebuild the Object during Recreate. But, I wanted the saved string to be uncluttered and easy to read and understand (and edit, if the spirit moves you). I also enjoyed the rather exhausting project of unpacking and trying to understand the small portion of the CF Rules that I managed to handle.

## Limitations:

Currently, there are probably some conditions and formats that are not saved (or restored).

Formats Handled: Interior Color (including Color Scale), Font Color/Bold/Italic, Icon Sets, and most Borders.

Formats Not Yet Handled: Fill Patterns,

So far, limited testing indicates successful handling of condition types:

-   xlExpression
-   xlCellValue (Operators: xlBetween, xlNotBetween, xlEqual,)
-   xlUniqueValues (including DupeUnique or "Duplicates")
-   xlTextString (TextOperators: xlContains, xlDoesNotContain, xlBeginsWith, xlEndsWith)
-   xlColorScale (2 & 3 colors), xlTimePeriod (all DateOperators – like xlLastWeek - probably)
-   xlBlanksCondition, xlNoBlanksCondition, xlErrorsCondition, xlNoErrorsCondition
-   xlTop10 (so far only top N %)
-   xlIconSet

## Operational Notes:

1.  To Install, go to Developer -\> Visual Basic. Right click on Left Nav “Modules” and select “Import File”. Do that for file: “SaveRestore.bas”
2.  Let’s say you want to save and later recreate Conditional Formatting Rules for your tab “MyData”, do this: Somewhere on your workbook (possibly on a “Readme” or “Misc” tab) select a cell and give it a range name (try clicking in the pull-down just under the left side of the menu bar and type the name).  
    The name should be in the form \<TabName\>_CF_RULES. An example would be: MyData_CF_RULES
3.  If you also want your saved Rules to show which columns (by the name you give them in the header row) are affected by each Rule, select the header row in your spreadsheet and give a range name in the form \<TabName\>_CF_hdrRow, for example MyDate_CF_hdrRow.
4.  Go to your data worksheet. On menu tab Developer click “Macros” and select SaveConditionalFormattingToString to test out the save function.

## Test Case:

In my repository there is a file called TestCase.txt which is a recent copy of the saved rules (from my WatchedHistory tv/movie bingeing spreadsheet). The rules are mostly arbitrary and some are weird – I just wanted to have a test case with a lot of different rules.
