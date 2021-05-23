title: Stop using INDIRECT

# Introduction

`INDIRECT` is the top function that should be avoided in Excel. It has the following drawbacks:
1. `INDIRECT` makes Excel calculation slow
2. Having `INDIRECT` function present is an indicator of bad spreadsheet design
3. Spreadsheet is prone to errors upon modification
4. Formula auditing is impossible

This article will explore each of the drawback, and explain the correct spreadsheet design approach which will eliminate `INDIRECT`.

# Brief Description about INDIRECT

You can refer to the [official manual of `INDIRECT`](doc-indirect).

The `INDIRECT` function takes one required argument that represents a cell address and returns the value in that address. As an example: `=INDIRECT("A1")` gives you the value in the cell `A1`.

# There are 4 key problems associated with using INDIRECT

## 1. INDIRECT makes calculation slow

When an Excel workbook contains an `INDIRECT`, with automatic calculation turned on, you will notice that the spreadsheet calculates upon every edit in any cells. This is the first sign that there is an `INDIRECT` somewhere in the opened Excel workbooks (the `INDIRECT` can be in another opened workbook).

This recalculated-every-time behavior is triggered by the fact that `INDIRECT` is a [volatile function](doc-volatile-function). Volatile functions should be avoided. More examples are `NOW`, `TODAY`, `OFFSET`.

TODO: (How much slower?)

## 2. `INDIRECT` is an indicator of bad spreadsheet design

The most common use case of `INDIRECT` is when the same kind of data is put in different worksheets. A common example is the same policy data report of last 12 months are put in 12 different worksheets of names 'month-1' to 'month-12'. Each of the worksheet has a total count at the cell `A1`, therefore the "Summary" worksheet uses `=INDIRECT("month-"&C5&"!A1")` (`C5` is a number between 1 to 12) to obtain the total count for each month.

This design to put the same kind of data in different worksheets creates the need to use `INDIRECT`, making the spreadsheet slow. It also limits the capability for further analysis of data across months.

The next section of this article illustrates the proper design that will eliminate the needs for `INDIRECT` and enables more complex analysis of the by-month data.
When the spreadsheet is designed well, `INDIRECT` will not be necessary.

## 3. Spreadsheet is prone to error upon modifications

When rows and columns are inserted or deleted, the formula in the affected cells are automatically adjusted by Excel. This ensures the correctness of the spreadsheet is maintained and is not affected by the layout changes.

See my previous post on [Excel's logic-layout dependency](blog-logic-layout-dependency) which talks about dependency between logic and layout.

Any `INDIRECT` functions cannot be automatically adjusted, because the address in an `INDIRECT` function is a variable. When a workbook can contain `INDIRECT` (where the user/maintainer may not be aware of), it can be hard to know whether it is safe to rename a worksheet, delete a worksheet, or insert rows and columns. Any modifications can potentially break a worksheet.

## 4. Spreadsheet formula auditing is impossible

To study and review an Excel workbook, it is common to trace the dependents and precedents of a cell. However any cell referenced by `INDIRECT` cannot be traced.
This is similar to the problem above for being prone to errors.

# The spreadsheet design that will eliminate INDIRECT

The golden rule is to put the same kind of data in one single worksheet. The attribute that is used to distinguish the data across worksheets, i.e. the month, is present as a separate column.

With such design, the data across different months can be summarized easily, through a pivot table or simple `SUMIF` formula. Investigation of data throughout the months is also easier with AutoFilter. More complicated analysis can also be done.

If the data volume is too large to be contained in one single worksheet, alternatives such as database software should be considered.

# Does Manual Calculation mode help?

No. While using manual calculation mode seems to eliminate the slowness upon editing, all the problems mentioned above remain there because `INDIRECT` is not removed.

1. The calculation is still slow when the workbook is calculated.
2. The spreadsheet is still prone to errors due to layout changes.
3. The spreadsheet design remains bad, which hinders future maintenance and limits the analysis capabilities.
4. Formula auditing is still impossible.

Even worse, manual calculation mode introducts further problem:

## Problem 1: Wrong results when forgotten to press F9

Excel gives wrong results when calculation is not triggered, and Excel does not provide a clear-enough indicator for this.

## Problem 2: Workbooks saved in Manual calculation mode can alter other users' calculation mode

The actual calculation mode of Excel is not a simple workbook-only or machine-only settings. It depends on the workbooks opened.

I strongly encourage you to read the [official documentation](doc-calculation-mode) about how your Excel's calculation mode is determined.

After reading the official documentation, you will understand how saving one workbook in manual calculation mode can affect other users of the same workbook, and in turn their other opened workbooks, even if they never use manual calculation in their other workbooks.

It poses a serious threat because even a user who always uses automatic calculation can unintentionally have his/her Excel switched to manual calculation mode simply by opening one workbook. They may not be aware of the fact that their workbooks are no longer automatically calculated and thus may use the wrong un-refreshed results from their workbooks.

# Keyboard Shortcuts

When Excel is set as automatic calculation mode, these shortcuts will no longer be necessary. Calculation-related shortcuts are listed below nevertheless, in case there is a workbook that calculates so slowly that it forces manual calculation mode temporarily.

You can also read the detailed article about [Excel Recalculation](doc-excel-recalculation) to understand how and when Excel performs calculation.

| Description | Shortcut |
| --- | --- |
| Calculate worksheet | `Shift+F9` |
| Calculate all opened workbooks | `F9` |
| Calculate workbook (with tree rebuild) | `Ctrl+Alt~Shift+F9` |
| Calculate all workbooks (with tree rebuild) | `Ctrl+Alt+F9` |
| Calculate selected cell range (Replacing `=` by `=` in all cells) | `Ctrl+H`, `=`, `Tab`, `=`, `Alt+A` |

# Conclusion

`INDIRECT` should not be used in all circumstances. If you find that you need to use the `INDIRECT` function, it is time to review the spreaadsheet design. The same kind of data should be put in one single worksheet instead of multiple worksheets, i.e. no more "one-tab-per-month" design.

Manual calculation is not a solution for eliminating the slowness and unresponsiveness brought by the `INDIRECT` function. And it will create even more problems affecting other users as well.

[doc-indirect]: https://support.microsoft.com/en-us/office/indirect-function-474b3a3a-8a26-4f44-b491-92b6306fa261
[doc-volatile-function]: https://docs.microsoft.com/en-us/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions
[doc-calculation-mode]: https://docs.microsoft.com/en-us/office/troubleshoot/excel/current-mode-of-calculation
[doc-excel-recalculation]: https://docs.microsoft.com/en-us/office/client-developer/excel/excel-recalculation
[blog-logic-layout-dependency]: https://www.actuaries.tech/5-examples-of-logic-layout-dependency-in-excel/
