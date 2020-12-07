title: Avoid using VLOOKUP

# Introduction

[`VLOOKUP`][doc-vlookup] is a function very well-known because everyone uses it. However `VLOOKUP` creates problems most of the time. A better approach will be using the combination of `INDEX` and `MATCH` functions. It may sound complex (using two functions makes the formula longer) at first, but I will explain the rationale in this article.

# Problems with VLOOKUP

`VLOOKUP` is very popular. Nearly all of us know how to use it. Yet it is important to understand the weaknesses of this function.

Throughout this article, the formula `=VLOOKUP(A3, $A$8:$D$16, 4, FALSE)` will be used for the illustration purpose.

## 1. Dependency of layout and logic - the order of columns

The third argument: `4` in `VLOOKUP(A3, $A$8:$D$16, 4, FALSE)` refers to the 4th column. This [magic number][wiki-magic-number] does not adjust itself when the content of the columns `$A$8:$D$16` are changed.

To prevent the problem of the magic number, some people add a number reference above the columns, as shown in the screenshot below.

However manual adjustment is still necessary whenever columns are added (Filling in formula in new columns or adjusting references to deleted columns). It creates extra work; and the worksheet is not clean enough since it is "polluted" by the magic numbers.

In addition, there is a restriction in using `VLOOKUP`: the value (`A3`) must be matched with the leftmost column of the table array, and the target column hence must be placed on the right side of it. This is another dependency of layout and logic. It is non-sense that we must order the columns this way just for the sake of the `VLOOKUP` function.

[![VLOOKUP restrictions on column orders][image-column-order]][image-column-order]

## 2. Referencing unused ranges and values

The table array in `VLOOKUP` (the second argument) has to include both the matching column and the target column. Usually this range consists of far more columns than the two that are necessary. It creates trouble for formula auditing because it will be difficult to trace which cells are dependents of a particular cell. As shown in the below exhibit, because of the data array covering the entire table, Excel thinks that the Sex column is necessary for the computation of the "plan code" (as well as all other fields using the whole data range), but it is obviously not the case.

[![VLOOKUP causes cell dependency problem][image-unused-cells]][image-unused-cells]

# What should I use as an alternative for VLOOKUP?

Imagine you are about to enter the formula:

```
=VLOOKUP(A3, $A$8:$D$16, 4, FALSE)
```

The followings are two possible alternatives which can avoid the problems mentioned above:

## 1. INDEX with MATCH

```
=INDEX($D$8:$D$16, MATCH($A3, $A$8:$A$16, 0))
```

Official documentations of the `INDEX` and `MATCH` functions:

- [`INDEX`][doc-index]
- [`MATCH`][doc-match]

This formula perfectly solves the problems because all arguments involved are real cell references, and they are exactly the cells to be used. There are no unused cells nor column order problems.

The last argument (zero) is equivalent to the last argument in `VLOOKUP` (`false`), while the `MATCH` function provides more than that.

An extra benefit is the quick navigation to the target column – Press `Ctrl+[` in the cell and it will bring you to `$D$8:$D$16` directly.

Using `INDEX` with `MATCH` has another important, and often unnoticed advantage over the other alternatives: calculation speed optimization on large amounts of data. This will be discussed in another article.

[![Use INDEX and MATCH instead of VLOOKUP][image-index-match]][image-index-match]

## 2. LOOKUP

```
=LOOKUP($A3, $A$8:$A$16, $D$8:$D$16)
```

[`LOOKUP`][doc-lookup] is not as flexible as the above function due to its limited nature, but is definitely shorter. The range to be matched must be sorted in ascending order; also it can only return the match for the largest value that is less than or equal to the value.

[![LOOKUP syntax with exact reference data columns][image-lookup]][image-lookup]

## 3. XLOOKUP

How about `XLOOKUP`?

[`XLOOKUP`][doc-xlookup] is a new function only available through Microsoft 365 subscribers. Since `XLOOKUP` is not available in earlier versions, the compatibility issue is the major issue here, it should only be used when you are certain that there will not be other users opening your file with an older version of Excel.

While `XLOOKUP` addresses some of the pitfalls of `VLOOKUP`, it still suffers from two problems:

1. Only exact match mode is supported. This is a downgrade from the similar `MATCH` and `VLOOKUP` functions.
2. Searching efforts cannot be reduced When there are multiple instances of `XLOOKUP` with the same key. Such searching time can only be reduced with `MATCH`.

# When should I use VLOOKUP?

The only scenario where using `VLOOKUP` can still be acceptable (but not fully justified because the same can be easily achieved with `INDEX` and `MATCH` as well) is when the target column is a dynamic number. For example, in order to find a premium rate by issue age and duration in a select rate table, you may use something like:

```
=VLOOKUP(Sex & IssueAge, PREM_RATE_RANGE, Duration+1, FALSE)
```

However this example of usage still suffers from the problem of layout-logic dependency. Imagine when a column is inserted between the first column (the key) and the Premium Rates... The correct way to solve the problem is left for you as an exercise.

In all other scenarios `VLOOKUP` should never be used.

# Summary

Avoid the `VLOOKUP` function. It damages the layout structure of Excel.

I recommend using `INDEX+MATCH` because of the possible options and the convenience in navigation that it gives. This will be the only one function you need (or two functions, if you prefer).

You will:
a) find your workbook less prone to error due to layout changes, and
b) find the formula easier to read because you see the exact cell ranges.

[doc-vlookup]: https://support.office.com/en-us/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1
[doc-index]: https://support.office.com/en-us/article/INDEX-function-a5dcf0dd-996d-40a4-a822-b56b061328bd
[doc-match]: https://support.office.com/en-us/article/MATCH-function-e8dffd45-c762-47d6-bf89-533f4a37673a
[doc-lookup]: https://support.office.com/en-us/article/LOOKUP-function-446d94af-663b-451d-8251-369d5e3864cb
[doc-xlookup]: https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929
[wiki-magic-number]: https://en.wikipedia.org/wiki/Magic_number_(programming)
[image-column-order]: /content/images/2017/12/vlookup-numbering-columns.png
[image-unused-cells]: /content/images/2017/12/vlookup-unused-cells.png
[image-index-match]: /content/images/2017/12/vlookup-index-match.png
[image-lookup]: /content/images/2017/12/vlookup-lookup.png
