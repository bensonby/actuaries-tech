title: About Dates (Part 1: Basics)

# Introduction

Dates are special in Excel, and are worth discussing. This article will explain the basics of dates in Excel and illustrate some common ways to manipulate dates.
# Dates are numbers

Dates are just numbers in Excel.

The date `31-Dec-2011` is equivalent to the number 40908, i.e. the display “31-Dec-2011″ is a format setting for the number 40908, just like you apply cell format to change a normal number 40908 to percentage `4090800%` (`Ctrl+Shift+5`), or to standard number format `40,908.00` (`Ctrl+Shift+1`). The value 40908 stays the same regardless of which cell format you apply.

This numeric representation of a date, called the "serial number", is the number of days since 1st Jan 1900.

# Formula for dates
[`DATE`][doc-date] can be used to create a date by providing the year, month and day as arguments, e.g. `DATE(2011, 10, 31)`.

Yet these also work:
- `DATE(2010, 14, 1)` is 1st Feb 2011
- `DATE(2011, 3, 0)` is 28th Feb 2011
- `DATE(2011, 3, -1)` is 27th Feb 2011
- `DATE(2012, 0, 0)` is ... left as an exercise for you

[`EOMONTH`][doc-eomonth], "end of month", is a useful function since actuaries often deal with month end dates.
- `EOMONTH(DATE(2020, 9, 30), 1)` is 31st Oct 2020
- `EOMONTH(DATE(2020, 9, 30), 2)` is 30th Nov 2020
- `EOMONTH(DATE(2020, 9, 30), -1)` is 31st Aug 2020
- `EOMONTH(DATE(2020, 9, 15), 0)` is 30th Sep 2020

Other (less useful) date creation functions include [`DATEVALUE`][doc-datevalue] and [`TODAY`][doc-today].

The year, month and day of a date can be retrieved by the functions [`YEAR`][doc-year], [`MONTH`][doc-month] and [`DAY`][doc-day] respectively.
# Arithmetics on dates

Since dates are equivalent to numbers representing the number of days since 1st Jan 1900, basic arithmetics apply:

- `DATE(2020, 12, 31)-DATE(2020, 5, 2)`: number of days between 2nd May and 31st Dec in 2020;
- `DATE(2020, 11, 28)+7`: the same day next week from 28 Nov 2020.

You often need more complicated calculations, don’t you?
[`DATEDIF`][doc-datedif] is the must-know function.
# The useful function: DATEDIF
[`DATEDIF`][doc-datedif] has long been available in Excel – probably since the earliest version of Excel you have worked on, but it has not been well documented and thus not many actuaries know about it.

`DATEDIF` calculates the differences in number of years/months/days given two dates.

Examples:
- `DATEDIF(DATE(2019, 3, 16), DATE(2020, 12, 1), "m")` is 20, i.e. 20 complete months in between
- `DATEDIF(DATE(1994, 3, 16), DATE(2020, 12, 31), "y")` is 26, i.e. the age last birthday of a person born in 16th Mar 1994 as of 31st Dec 2020.
- `DATEDIF(DATE(2019, 11, 30)+1, DATE(2020, 8, 31)+1, "m")` is 9, i.e. number of months between the two month ends. Notice how the `+1` helps to align the two dates to the first day of a month for `DATEDIF` to work.

`DATEDIF` also supports calculating the differences neglecting the year/month. Using `ym` as the last argument will calculate the difference in months ignoring the year. The same applies for `yd` and `md`. However according to the [documentation][doc-datedif], using `md` may produce inaccurate results. Nevertheless `ym`, `yd` and `md` are seldom used.

# Date Format
The display of the date can be adjusted by “Format Cells” (Ctrl+1) -> “Date”. There are several preset formats available for you to choose.

If none of them suit your needs, you can select “Custom” and enter your own “formatting code” to control how the date is displayed.

The list of available formatting codes can be found in the [official documentation][doc-date-format].
# Using dates in text
The second most important function after `DATEDIF` will be [`TEXT`][doc-text]. This function is not specifically created for dates but as an actuary it is often used with presenting dates.

When we generate financial reports in Excel, the file often takes the current reporting date as an input, e.g. 30-Nov-2020. And there will be a heading like "Balance Sheet as of Nov 2020".

This can be created by `="Balance Sheet as of "&TEXT(DATE(2020, 11, 30), "mmm yyyy")`.

I have seen people creating a mapping from the number 1-12 to the text Jan-Dec but this is not necessary.

The format `mmm yyyy` used in the example above is the same [formatting code][doc-date-format] mentioned above. What `TEXT` does is to return a string by formatting your value (`DATE(2020, 11, 30)`) using the provided format (`mmm yyyy`).
# Keyboard Shortcuts
1. `Ctrl+;` to insert current date to the cell
2. `Ctrl+Shift+3` to format a cell as date
3. `Ctrl+Shift+~` to format a cell as general number format, i.e. showing 40908 for the date 31 Dec 2011.
4. `Ctrl+1 right Alt+C d` to customize the [date format][doc-date-format]
# Conclusion (TL;DR)
1. Dates in Excel are numbers representing the number of days since 1/1/1900.
2. Make good use of the arithmetic functions and the useful functions `DATEDIF`, `EOMONTH` and `TEXT`. Do not create long and stupid date formula in Excel ever again.

# Reference
There are quite a few functions introduced in this article. For easier reference I re-list them below:

- [`DATE`][doc-date]
- [`DATEVALUE`][doc-datevalue]
- [`TODAY`][doc-today]
- [`YEAR`][doc-year]
- [`MONTH`][doc-month]
- [`DAY`][doc-day]
- [`DATEDIF`][doc-datedif] (Must-know!)
- [`EOMONTH`][doc-eomonth] (Must-know!)
- [`TEXT`][doc-text] (Must-know!)
- [List of date formatting codes][doc-date-format]

[doc-date]: https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349
[doc-datevalue]: https://support.microsoft.com/en-us/office/datevalue-function-df8b07d4-7761-4a93-bc33-b7471bbff252
[doc-today]: https://support.microsoft.com/en-us/office/today-function-5eb3078d-a82c-4736-8930-2f51a028fdd9
[doc-year]: https://support.microsoft.com/en-us/office/year-function-c64f017a-1354-490d-981f-578e8ec8d3b9
[doc-month]: https://support.microsoft.com/en-us/office/month-function-579a2881-199b-48b2-ab90-ddba0eba86e8
[doc-day]: https://support.microsoft.com/en-us/office/day-function-8a7d1cbb-6c7d-4ba1-8aea-25c134d03101
[doc-datedif]: https://support.microsoft.com/en-us/office/datedif-function-25dba1a4-2812-480b-84dd-8b32a451b35c
[doc-eomonth]: https://support.microsoft.com/en-us/office/eomonth-function-7314ffa1-2bc9-4005-9d66-f49db127d628
[doc-text]: https://support.microsoft.com/en-us/office/text-function-20d5ac4d-7b94-49fd-bb38-93d29371225c
[doc-date-format]: https://support.microsoft.com/en-us/office/format-a-date-the-way-you-want-8e10019e-d5d8-47a1-ba95-db95123d273e#ID0EAACAAA=Windows
