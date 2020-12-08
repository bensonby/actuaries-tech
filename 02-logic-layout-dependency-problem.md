title: 5 Examples of Logic-Layout Dependency in Excel

# Introduction

[Model-view-controller (MVC)][mvc-wiki], first described by Trygve Reenskaug in 1979, is an architectural pattern widely used in software engineering.

One of the key principles of MVC is the separation of the business logic (controller) and the presentation layer (view), which can reduce the complexity in architectural design and increase the code flexibility and maintability.

In simple words, either the layout or the business logic can be changed without affecting the other.

# Excel's logic-layout dependency

Excel can be considered as a programming interface:

- the cell formula is the logic (controller)
- the layout and format of the cells are the presentation (view)

Cell addresses in the formula, being the basic design of Excel, makes the separation of logic and layout difficult.

A well-designed workbook should work fine even when any cell is moved or any column is inserted at any place. Excel already does a great job in adjusting the formula automatically when cells are moved, but there are still many cases where automatic adjustments do not work.

Below are 5 examples of Excel features that involves strong logic-layout dependency that are hard to separate.

## Example 1 – VLOOKUP

```=VLOOKUP(C1, Data!$A$3:$Z1000, 6, FALSE)```

The number `6` in the above [`VLOOKUP`][vlookup-doc] is a simple example.

When the columns of the table array are rearranged, this function will produce the wrong result because the column number does not change accordingly.

Some people try to avoid this problem by using a number specifying the column index above the headers. Yet this is not the cleanest solution because it still suffers from the logic layout dependency.

The proper solution is to use [`INDEX` and `MATCH`][blog-vlookup].

[![`VLOOKUP` formula failure][image-vlookup]][image-vlookup]

## Example 2 – INDIRECT

The [`INDIRECT`][indirect-doc] function fails when any kind of layout changes is made, because the `INDIRECT` function accepts a string, not a cell reference. The string will never be able to adapt to layout changes automatically.

A "bonus" flaw is that `INDIRECT` makes your workbooks significantly slower!

[![`INDIRECT` formula failure][image-indirect]][image-indirect]

## Example 3 – Pivot Tables

When a cell refers to the values in a pivot table by their cell addresses (`A1`, `C5`, etc.), these cell references do not change automatically when the layout of the pivot table is changed, e.g. a new field is added to the pivot table.

Excel also provides a `GETPIVOTDATA` function yet it also has some caveats. (To be covered in later article on pivot tables)

[![Pivot table formula failure][image-pivot-table]][image-pivot-table]

## Example 4 – Data Tables

In order for [data tables][data-table-doc] to work, the cells and the formula need to be laid out in the designated way. In this way the logic and layout are closely tied together and cannot be separated.

[![Data Table formula failure][image-data-table]][image-data-table]

## Example 5 – VBA

A lot of VBA programs "macros" suffer from the logic-layout dependency problems. Many VBA programs are not bulletproof in the way that these operations can easy break the macro:

- inserting/deleting rows and columns, or
- simply moving cells and renaming worksheets.

The below code snippet is an illustration – it contains the references `B6`,  `output` and `A10`. Whenever the cells `B6` and `A10` are moved, or the worksheet `output` is renamed or deleted, the snippet will produce errors.

(And you will see yet another vague error message from VBA! Yay!)

```vb
Public Sub testing()
  inputPlanCode = Range("B6").Value
  Worksheets("output").Range("A10").Value = inputPlanCode
End Sub
```

There are some best practices for writing better VBA. However the best way is to avoid VBA whenever possible, given the amount of care needed to write readable and maintenable VBA codes.

It turns out that most of the VBA macros are not necessary at the first place when:

- spreadsheets are well-designed, and
- we use the right tool for the right task


# Summary

The grid layout and the labeling of rows and columns (e.g. a cell is called `E9`) are great obstacles for good readability and maintainability of the spreadsheets. They make the separation of logic and layout difficult to achieve.

Possible solutions to the above issues will be discussed in the upcoming articles. The key concept is to reduce the logic-layout dependency, which will make our lives with Excel much easier. For programmers, such concept is even more crucial when writing codes.

[mvc-wiki]: http://en.wikipedia.org/wiki/Model%E2%80%93view%E2%80%93controller
[vlookup-doc]: https://support.office.com/en-us/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1
[indirect-doc]: https://support.office.com/en-us/article/INDIRECT-function-474b3a3a-8a26-4f44-b491-92b6306fa261
[data-table-doc]: https://support.microsoft.com/en-us/office/calculate-multiple-results-by-using-a-data-table-e95e2487-6ca6-4413-ad12-77542a5ea50b
[image-data-table]: /content/images/2017/12/logic-layout-data-table.png
[image-pivot-table]: /content/images/2020/12/logic-layout-pivot-table.png
[image-indirect]: /content/images/2020/12/logic-layout-indirect.png
[image-vlookup]: /content/images/2020/12/logic-layout-vlookup.png
[blog-vlookup]: /stop-using-vlookup/
