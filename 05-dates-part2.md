title: About Dates (Part 2: Date formats)

In case you haven’t read, here is [About Dates (Part 1: basics)][blog-about-dates-part-1].

# Default Date Format in Excel

Excel has a default date format defined according to the Regional Settings in the installed machine, e.g. m/d/y on a United States settings. If we enter a date in a cell using the correct format `4/28/2020`, it is recognized as a date automatically and is right-aligned. On the contrary, if we enter `28/4/2020`, Excel cannot recognize it as a date but as a text which is left-aligned.

TODO: add regional setting winows + excel screenshot of dates

This system-specific default format generally does not cause issues because we know how our computer is set up and do not change this settings often, thus entering dates in Excel is not a problem.

However, when CSV files containing dates are opened in Excel, the dates may not be interpreted correctly if the dates are not saved properly.

# Potential Problems with dates in CSV files

To understand the possible problem with dates in CSV files, we need to look at the process of:
- saving an Excel file into CSV format
- reading a CSV file from Excel

## Saving an Excel file into CSV format

When an excel worksheet is saved in CSV format, in the CSV the dates are stored in the same format as displayed in Excel.
TODO: add illustration, excel shows vs csv shows

## Opening CSV files in Exel

Let's illustrate with 3 examples.

1. When a csv file contains a cell "08/06/2020":

    | Machine default date format | Excel interpretation of 08/06/2020 |
    | --- | --- |
    | mm/dd/yyyy | 6th August 2020 |
    | dd/mm/yyyy | 8th June 2020

2. When a csv file contains a cell "27/06/2020":

    | Machine default date format | Excel interpretation of 06/27/2020 |
    | --- | --- |
    | mm/dd/yyyy | 27th June 2020 |
    | dd/mm/yyyy | 06/27/2020 as a text, not a date |

3. When a csv file contains a cell "08/30/35":

    | Machine default date format | Excel interpretation of 08/30/35 |
    | --- | --- |
    | mm/dd/yyyy with century cut-off year at 2029 | 30th August 1935 |
    | mm/dd/yyyy with century cut-off year at 2040 | 30th August 2035 |
    | dd/mm/yyyy | 08/30/35 as a text, not a date |

For two-digit years like "35", and other date format inputs like "09/13", there is a [detailed explanation from the official documentation][doc-two-digit-years] about how Excel interprets these date inputs.

From the above examples, it can be concluded that the interpretation of a date in the csv file depends entirely on the settings in the computer that opens the csv file.

This is a serious problem because the same data can be interpreted inconsistently by people on different machines.

# Avoid the date ambiguity

The problems with date ambiguity in CSV files can be avoided by setting all dates to an appropriate format like `YYYY-MM-DD` or `DD-MMM-YYYY` before saving. An alternative is to save the date as an 8-digit number `yyyymmdd`.

- `MDY` or `DMY` ordering leads to unrecognized dates
- Using two-digit years leads to ambiguity for which century being used. This ambiguity is common in date of birth (past dates) or policy maturity dates (future dates).

Not limited to CSV files, using unambiguous date formats is also considered good practice in Excel files because the dates can be interpreted by human more easily without causing confusions.

# How to fix the mis-interpreted dates

When we find that the dates in the CSV files are not equivalent to Excel's default date format, there are two ways to fix. The first method (formula-driven) is the recommended because there are no manual adjustments involved.

TODO: screenshot of wrong dates

1. Formula Approach

    Since dates are equivalent to serial numbers, we can use the `ISNUMBER` function to check if a date in CSV is interpreted as a date in Excel, or as a text due to incompatible date formats.

    - Simple case when dates are in MM/DD/YYYY format:
       ```
        =IF(ISNUMBER(A1), DATE(YEAR(A1), DAY(A1), MONTH(A1)), DATE(RIGHT(A1, 4), MID(A1, 4, 2), LEFT(A1, 2)))
       ```
    - More complicated case when dates are in M/D/YYYY format:
      ```
        =IF(ISNUMBER(A1), DATE(YEAR(A1), DAY(A1), MONTH(A1)), DATE(RIGHT(A1, 4), LEFT(RIGHT(A1, LEN(A1)-FIND("/", A1)), FIND("/", A1, FIND("/", A1)+1)-FIND("/", A1)-1), LEFT(A1, FIND("/", A1)-1)))
      ```

2. Text delimiting approach

    (Assume the default date format is M/D/YYYY and it opens a CSV file with date format D/M/YYYY)

    The method is to use the Text to Columns (Alt+D+E) feature in Excel to edit the dates.

    TODO: illustrations

# Dates as key for lookup

Dates are also used to "look up" information. When dates are used as a component of the key, for easier reference and better readability, use `TEXT` to create a readable key instead of leaving it as the serial number of the date.

TODO: insert example

# Date format on file and folder management

Although not directly related to Excel, in terms of file and folder management, the most common mistake is using a wrong date format to name a file or a folder.

Examples of correct formats are:
- `2020-11-30`
- `2020-11`
- `2020Q4`
- `20201120` (not perfect but not likely to cause ambiguity)

Examples of wrong formats are:
- `0812` (2008 Dec / 8th Dec / Aug 2012?)
- `201208` (2012 Aug / 20th Dec 2008 / 8th Dec 2020?)
- `19Q4` (99Q4 will be sorted after 19Q4, or what if there will be year 2119?)
- `4Q2020` (Confusing sorting order, all 1Q being sorted first)
- `Apr2019` (The ugliest sorting order ever, Apr, Aug, Dec, Feb, Jan, etc.)

Correct formats are those that enforce date hierarchy from year to month to day, making file listing order correct, and introduce no ambiguity in the numbers.

There will be separate articles talking about file and folder management.

# Hotkeys

Hotkeys are always essential for fast manipulations in Excel.

`Ctrl+Shift+3` to make a cell in date format (dd-mmm-yy)
`Ctrl+Shift+~` to make a cell in general format (serial number for dates)
`Ctrl+1 right Alt+C d` to customize the [date format][doc-date-format]
`Alt+D+E` for opening Text to Columns wizard

# Conclusion

While the concept of serial number brings convenience in date arithmetics, it creates issues due to the different date formats, especially on csv file.

The interpretation of dates (m/d/y v.s. d/m/y) in Excel depends on system settings. Thus the same CSV file can be interpreted differently on machines.

Two suggested good practices when dealing with output of dates to CSV are:
  - use the number `yyyymmdd` to represent a date (An 8-digit number)
  - use the format `yyyy-mm-dd` or dd-mmm-yyyy to save dates into CSV files
    
[doc-two-digit-years]: http://support.microsoft.com/kb/214391
[blog-about-dates-part-1]: /about-dates-part-1
