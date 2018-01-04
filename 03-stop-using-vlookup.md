title: Avoid using VLOOKUP

# Introduction

[`VLOOKUP`][doc-vlookup] is a function very well-known because everyone uses it. However `VLOOKUP` creates problems most of the time. A better approach will be using the combination of `INDEX` and `MATCH` functions. It may sound complex (using two functions makes the formula longer) at first, but it definitely improves in terms of readability and speed.

# Problems with VLOOKUP

`VLOOKUP` is very popular. Nearly all of us know how to use it. Yet it is important to understand the weaknesses of this function.

Throughout this article, the formula `=VLOOKUP(C9, Data!$A$2:$Z$6000, 4, FALSE)` will be used for the illustration purpose.

## 1. Dependency of layout and logic - the order of columns

The third argument `4`, the column number of the table array, is in this case the 4th column, column D.

This "magic number" does not adjust itself when the columns are changed in the `Data` worksheet.

To eliminate the magic number, many people add a number reference above the columns, as shown in the screenshot below.

[![VLOOKUP restrictions on column orders][image-column-order]][image-column-order]

However manual adjustment is still necessary whenever columns are added (Filling in formula in new columns or adjusting references to deleted columns). It still creates extra work; and the worksheet is not clean enough since it is "polluted" by the magic numbers.

In addition, there is a restriction in using `VLOOKUP`: the value (`C9`) must be matched with the leftmost column of the table array, and the target column hence must be placed on the right side of it. This is an example of [dependency of layout and logic][article-02-logic-layout-dependency]. It is nonsense that we must order the columns this way just for the sake of the `VLOOKUP` function.

## 2. Referencing unused ranges and values

The table array in `VLOOKUP` (the second argument) has to include both the matching column and the target column. Usually this range consists of far more columns than the two that are necessary. It creates trouble for formula auditing because it will be difficult to trace which cells are dependents of a particular cell. As shown in the below exhibit, because of the data array covering the entire table, Excel thinks that the Sex column is necessary for the computation of the "plan code" (as well as all other fields using the whole data range), but it is obviously not the case.

[![VLOOKUP causes cell dependency problem][image-unused-cells]][image-unused-cells]

# Using INDEX with MATCH as an alternative for VLOOKUP

Imagine you are about to enter the formula:

```
=VLOOKUP(A3, $A$8:$D$16, 4, FALSE)
```

You will now replace it with:

```
=INDEX($D$8:$D$16, MATCH($A3, $A$8:$A$16, 0))
```

_The last argument `0` is equivalent to the last argument in `VLOOKUP` (`false`)._

Official documentations of the `INDEX` and `MATCH` functions:

- [`INDEX`][doc-index]
- [`MATCH`][doc-match]

[![Use INDEX and MATCH instead of VLOOKUP][image-index-match]][image-index-match]

It gives you a few advantages:

1. All arguments involved are real cell references, and are exactly the cells to be used. No more reference to column B and C!

2. It enables quick navigation to the target column – Press `Ctrl+[` from `B3` (plan code) and it will bring you to `$D$8:$D$16` (plan code column) directly.

3. When you need multiple columns as results, you can do the `MATCH` only once, and use `INDEX` multiple times. The speed improvement will be huge in a complicated workbook!

[![Reuse MATCH result for multiple INDEX result cells][image-multple-match]][image-multiple-match]

# When should I use VLOOKUP?

The only scenario where using `VLOOKUP` can still be acceptable (but not fully justified) is when the target column is a dynamic number.

However when you get used to the `INDEX`-`MATCH` functions, you can easily accomplish the same instead of using `VLOOKUP`.

For example, in order to find a premium rate by issue age and duration in a select rate table, you may use something like:

```
=VLOOKUP(Sex & IssueAge, PREM_RATE_RANGE, Duration+1, FALSE)
```

However this example of usage still suffers from the problem of layout-logic dependency. Imagine when a column is inserted between the Rate Key and the Premium Rates. The correct way to solve the problem is left for you as an exercise.

In all other scenarios `VLOOKUP` should never be used.

# Summary

Avoid the `VLOOKUP` function. It damages the layout structure of Excel. Please use `INDEX+MATCH` as the replacement. You will a) find your workbook less prone to error due to layout changes, and b) find the formula easier to read because you see the exact cell ranges.

[doc-vlookup]: https://support.office.com/en-us/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1
[doc-index]: https://support.office.com/en-us/article/INDEX-function-a5dcf0dd-996d-40a4-a822-b56b061328bd
[doc-match]: https://support.office.com/en-us/article/MATCH-function-e8dffd45-c762-47d6-bf89-533f4a37673a
[image-column-order]: /content/images/2017/12/vlookup-numbering-columns.png
[image-unused-cells]: /content/images/2017/12/vlookup-unused-cells.png
[image-index-match]: /content/images/2017/12/vlookup-index-match.png
[article-02-logic-layout-dependency]: /5-examples-of-logic-layout-dependency-in-excel
