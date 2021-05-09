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

From my experience seeing the usage `INDIRECT` function is usually used in a scenario like this, where `INDIRECT` is used to obtain a value from a variable worksheet name:

# INDIRECT makes calculation slow

When an Excel workbook contains an `INDIRECT`, with automatic calculation turned on, you will notice that the spreadsheet calculates upon every edit in any cells. This is the first sign that there is an `INDIRECT` somewhere in the opened Excel workbooks (the `INDIRECT` can be in another opened workbook).

This recalculated-every-time behavior is triggered by the fact that `INDIRECT` is a [volatile function](doc-volatile-function). Volatile functions should be avoided. More examples are `NOW`, `TODAY`, `OFFSET`.

TODO: (How much slower?)

# `INDIRECT` is an indicator of bad spreadsheet design

The most common use case of `INDIRECT` is when the same kind of data is put in different worksheets. A common example is the same policy data report of last 12 months are put in 12 different worksheets of names '1' to '12'. Each of the worksheet has a total count at the cell `A1`, therefore the "Summary" worksheet uses `=INDIRECT(C5&"!A1")` (`C5` is a number between 1 to 12) to obtain the total count for each month.

This design to put the same kind of data in different worksheets creates the need to use `INDIRECT`, making the spreadsheet slow. It also limits the capability for further analysis of data across months.

The final section of this article illustrates the proper design that will eliminate the needs for `INDIRECT` and enables more complex analysis of the by-month data.
When the spreadsheet is designed well, `INDIRECT` will not be necessary.
`INDIRECT` can be avoided. The best way is by adopting a good spreadsheet design.

# Spreadsheet is prone to error upon modifications

When rows and columns are inserted or deleted, the formula in the affected cells are automatically adjusted by Excel. This ensures the correctness of the spreadsheet is maintained, and is not affected by the layout changes.

See my previous post on [Excel's logic-layout dependency](blog-logic-layout-dependency).

Any `INDIRECT` functions cannot be automatically adjusted, because the address in an `INDIRECT` function is a variable. When a workbook can contain `INDIRECT` (where the user/maintainer may not be aware of), it can be hard to know whether it is safe to rename a worksheet, delete a worksheet, or insert rows and columns. Any modifications can potentially break a worksheet.

# Spreadsheet formula auditing is impossible

To study and review an Excel workbook, it is common to trace the dependents and precedents of a cell. However any cell referenced by `INDIRECT` cannot be traced.
This is similar to the problem above for being prone to errors.

# The spreadsheet design that will eliminate INDIRECT

The golden rule is to put the same kind of data in one single worksheet. The attribute that is used to distinguish the data across worksheets, i.e. the month, is present as a separate column.

With such design, the data across different months can be summarized easily, through a pivot table or simple `SUMIF` formula. Investigation of data throughout the months is also easier with AutoFilter. More complicated analysis can also be done.

If the data volume is too large to be contained in one single worksheet, you should explore other alternatives such as a database for data storage.

[doc-indirect]: https://support.microsoft.com/en-us/office/indirect-function-474b3a3a-8a26-4f44-b491-92b6306fa261
[doc-volatile-function]: https://docs.microsoft.com/en-us/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions
