title: 5 Examples of Logic-Layout Dependency in Excel

TODO: reference to each of example article

# Introduction

[Model-view-controller (MVC)][mvc-wiki], first described by Trygve Reenskaug in 1979, is an architectural pattern widely used in software engineering. One of MVC’s virtue is the separation of the business logic and the presentation layer, which can reduce the complexity in architectural design and to increase the code flexibility and maintability . Either the layout or the business logic can be changed without affecting the other.

# Excel's logic and layout dependency

Excel can be considered as a programming interface, as it allows logic to be built by using formula, and output to be shown in cells. In this way we can regard the formula as the business logic, and the placement of cells in the workbook as the presentation layer.

The basic design of Excel makes the business logic and layout largely dependent of each other, making the separation difficult. A well-designed workbook should work fine even when any cell is moved or any column is inserted at any place. Excel already does a great job in the automatic adjustment of formula when cells are moved, but there are still many cases where automatic adjustment does not work.

Below are 5 examples of Excel features or formula that involves strong logic and layout dependency.

## Example 1 – VLOOKUP

The column number, or the third attribute, of the [`VLOOKUP`][vlookup-doc] function is a simple example. When a column of the table array is moved, this function will fail because the column number does not change accordingly. Some people try to avoid this problem by using a number above the column. But this is not the best solution as it still suffers from the logic layout dependency. The proper solution is explained in [Stop using VLOOKUP][article-vlookup].

[![`VLOOKUP` formula failure][image-vlookup]][image-vlookup]

## Example 2 – INDIRECT

The [`INDIRECT`][indirect-doc] function fails when any kind of layout changes is made, because the `INDIRECT` function accepts a string, not a cell reference. The string will never adapt to the layout changes automatically.

[![`INDIRECT` formula failure][image-indirect]][image-indirect]

## Example 3 – Pivot Tables

When a cell refers to the values in a pivot table by their cell addresses (`A1`, `C5`, etc.), these cell references do not change automatically when the layout of the pivot table is changed, e.g. a new field is added to the pivot table.

[![Pivot table formula failure][image-pivot-table]][image-pivot-table]

## Example 4 – Data Tables

In order for data tables to work, the cells and the formula need to be laid out in the designated way. In this way the logic and layout are closely tied together and cannot be separated.

[![Data Table formula failure][image-data-table]][image-data-table]

## Example 5 – VBA

A lot of VBA programs suffer from the logic-layout dependency problems. Many VBA programs are so poorly written that a) inserting/deleting rows and columns, or b) simply moving cells and renaming worksheets, make them fail. The below code snippet is an illustration – it contains the references `B6`,  `output` and `A10`. Whenever the cells `B6` and `A10` are moved, or the worksheet `output` is renamed or deleted, the snippet will fail.

```vb
Public Sub testing()
  inputPlanCode = Range("B6").Value
    Worksheets("output").Range("A10").Value = inputPlanCode
    End Sub
    ```

# Summary

The grid layout and the labeling of rows and columns (e.g. a cell is called `E9`) are great obstacles for good stability, readability and maintainability of the spreadsheets. They make the separation of logic and layout difficult to achieve. Possible solutions to the above issues will be discussed in the upcoming articles.

The key concept is to reduce the interdependency between the logic and layout in the spreadsheets, which will make our lives with Excel much easier. For programmers, such concept is even more crucial when writing codes.

[mvc-wiki]: http://en.wikipedia.org/wiki/Model%E2%80%93view%E2%80%93controller
[vlookup-doc]: https://support.office.com/en-us/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1
[indirect-doc]: https://support.office.com/en-us/article/INDIRECT-function-474b3a3a-8a26-4f44-b491-92b6306fa261
[image-data-table]: /content/images/2017/12/logic-layout-data-table.png
[image-pivot-table]: /content/images/2017/12/logic-layout-pivot-table.png
[image-indirect]: /content/images/2017/12/logic-layout-indirect.png
[image-vlookup]: /content/images/2017/12/logic-layout-vlookup.png
[article-vlookup]: /stop-using-vlookup
