title: Stop using VLOOKUP

post-url: stop-using-vlookup

# Introduction

[`VLOOKUP`](https://support.office.com/en-us/article/VLOOKUP-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1) is a function very well-known because of... sorry, `VLOOKUP` is not really that useful - it is not necessary at all. Please avoid using `VLOOKUP`.

# Problem with VLOOKUP

1. Dependency of layout and logic - the order of columns

 The value to be found is expressed as the column number. If the column order is altered, the function fails to adjust itself. In light of this, some people added a number reference above the columns, as shown in the screenshot below. However manual adjustment is still needed whenever columns are added (Filling in formula). Excel is not supposed to be that manual.

 In addition, there is a restriction where the value to be matched must be in the leftmost column of the range, and the target column hence must be placed on the right of it. This is another dependency of layout and logic. Please think: Why must I place my columns in this way just for the sake of the stupid `VLOOKUP` function?

 ![`VLOOKUP` restrictions on column orders](assets/posts/vlookup-numbering-columns.png)

2. Referencing unused ranges and values

 The range in `VLOOKUP` (the second argument) has to include both the matching column and the target column. Usually the range consists of far more columns other than the two in interest. It creates trouble for formula auditing. It is hard to trace which cells are dependent on a particular cell. As shown in the below exhibit, Excel thinks that the Sex is necessary for the computation of the "plan code" (as well as all other fields using the whole data range), but it is obviously not the case.

 ![`VLOOKUP` causes cell dependency problem](assets/posts/vlookup-unused-cells.png)

# What to use then, if not VLOOKUP?

Imagine you are about to type in the formula:

```
=VLOOKUP(A3, $A$8:$D$16, 4, FALSE)
```

The followings are two possible alternatives which can avoid the problems mentioned above:

## 1. INDEX with MATCH

 ```
 =INDEX($D$8:$D$16, MATCH($A3, $A$8:$A$16, 0))
 ```

 Please refer to the documentations of [`INDEX`](https://support.office.com/en-us/article/INDEX-function-a5dcf0dd-996d-40a4-a822-b56b061328bd) and [`MATCH`](https://support.office.com/en-us/article/MATCH-function-e8dffd45-c762-47d6-bf89-533f4a37673a) if necessary. This formula perfectly solves the problems because all arguments involved are real cell references, and they are exactly the cells in use. No unused cells. No column order problems. The last argument (zero) is equivalent to the last argument in `VLOOKUP` (`false`), while the `MATCH` function provides more than that. An extra benefit is the quick navigation to the target column – Press `Ctrl+[` in the cell and it will bring you to `$D$8:$D$16` directly.

 ![Use `INDEX` and `MATCH` instead of `VLOOKUP`](assets/posts/vlookup-index-match.png)

## 2. LOOKUP

 ```
 =LOOKUP($A3, $A$8:$A$16, $D$8:$D$16)
 ```

 Please refer to the documentation of [`LOOKUP`](https://support.office.com/en-us/article/LOOKUP-function-446d94af-663b-451d-8251-369d5e3864cb) if necessary. This function is not as flexible as the above function due to its limited nature, but is definitely shorter. The range to be matched must be sorted in ascending order; also it can only return the match for the largest value that is less than or equal to the value.

 ![`VLOOKUP` restrictions on column orders](assets/posts/vlookup-lookup.png)

# When should I use VLOOKUP?

The only scenario where the use of `VLOOKUP` can still be justified (acceptable but not fully justified) is when the target column is a dynamic number. For example, in order to find a premium rate by issue age and duration in a select rate table, you may use something like:

```
=VLOOKUP(Sex & IssueAge, PREM_RATE_RANGE, Duration+1, FALSE)
```

Please note that this usage still suffers from the problem of layout-logic dependency. Imagine when a column is inserted between the Rate Key and the Premium Rates... The correct way to solve the problem is left for you as an exercise.

In all other scenarios `VLOOKUP` should never be used.

# Summary

From now on, forget about the `VLOOKUP` function. It damages the layout structure of Excel. Please use `INDEX+MATCH` or `LOOKUP` to replace it. Personally I will recommend `INDEX+MATCH` because of the possible options and the convenience in navigation it brings. You will a) find your workbook less prone to error due to layout changes, and b) find the formula easier to read because you see the exact cell ranges.
