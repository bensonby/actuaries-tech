Preface - How is Tech important for actuaries?

# Tech skills is (was) the missing piece in the actuarial education

Actuaries work with computers every day. Besides the basic features of the computer (Microsoft Windows), actuaries spend the majority of time with other kinds of software, probably with large amounts of data. This includes:
1. spreadsheet software (Microsoft Excel)
1. database software (Microsoft Access, Oracle, FoxPro, etc.)
1. actuarial software (Prophet, AXIS, etc.)
1. Coding skill (VBA, SQL, command line)
1. Email software (Microsoft Outlook, Lotus Notes)


While they are essential elements in the everyday work, actuaries are not taught on how to use them properly; and there are no exams on the technological skills.

The consequences are:

1. Actuaries are self-taught on technological. Skills are acquired mostly through the work done by "previous" colleagues
1. It is easy to develop bad habits which will reduce the working efficiency.
1. In particular, Microsoft Excel spreadsheets are not well designed, meaning that they are not easily understood and cannot be easily maintained and extended. They also suffer from slow calculation speeds or even slow VBA run time.

When the technological skills are matured (?), you will be able to, as examples:
1. Use the right tool for the right tasks. (When you only have hammer blah blah)
1. [MS Excel] Create workbooks that are easily understood and self-documented (without extra efforts)
1. [MS Excel] Revamp existing poorly-designed worksheets, e.g. improving update time from 8 hours to 30 minutes from my experience, with 2-day effort
1. [MS Excel] Eliminate the manual calculation mode in Microsoft Excel because spreadsheets calculate fast.
1. [MS Word] Create easily maintainable Microsoft Word documents

I am starting this blog (hopefully it will turn into a e-book) to share some best practice regarding the different technical aspects of using the softwares.
Most of the times adopting the best practice does not involve any sophisticated technical knowledge. The essence is to create easily-understood output that others (or yourself) can follow, understand and maintain comfortably. Using the top-niche skill that others do not know does not help but only hinder the future maintenance.

-------------------------------------------

How is Tech important for actuaries?

# Technology skills is (was) the missing piece in the actuarial education   

We learn actuarial science and we take actuarial exams; we work with computers every day, yet we do not have proper lessons nor exams on IT skills.

Microsoft Excel skill is a particularly important one because most actuaries work with Microsoft Excel every day. Therefore developing good habits on spreadsheet design can boost the efficiency and avoid the followings:

<pic on large excel file size>
<pic on manual calculation>
<pic on #REF!>
<pic on 1004 error VBA>

With a good habit on spreadsheet designs, one can work much faster (and most people underestimate the potential improvement)!

# Good spreadsheet design does not require sophisticated use of Excel

Knowledge on VBA and all the uncommon Excel functions (`ADDRESS`) is useful (good-to-have), but these kinds of knowledge do not automatically make you an expert in spreadsheet design.

The truth is actually the opposite: the use of sophisticated features is an indicator of a badly designed spreadsheet.

For example,

1. a good spreadsheet design uses short and neat formula, which are easy to understand and follow.
<pic on repeated vlookup for if case>
2. A bad spreadsheet design may require VBA to achieve the desired outcome; a good spreadsheet design solves it not by VBA, but by proper data management and layout design.
<bad: each file represents a month; good: single worksheet for all data>
3. A good spreadsheet design shows the purpose of the file and each worksheet clearly by itself, without any supplementary long text description.
<bad: step by step guide, good: inline with good worksheet names and highlight of inputs>

# A good spreadsheet design improves the working efficiency significantly

A good spreadsheet design means the following:
1. Clear illustration of process flow
   It is not meant to be done with long text descriptions. This should be accomplished mainly using good worksheet flows, descriptive names, logical layout and clean formula.
1. Fast spreadsheet calculation with automatic calculation mode
   The use of manual calculation setting indicates spreadsheet design problem (leading to slow calculations) and creates risks of wrong results due to non-refreshed spreadsheets.
1. Clear and short formula
   Formula with multiple lines imply a spreadsheet design problem. Formula should be easy to understand and follow.
1. Clear descriptions of what purposes the workbook and each of the worksheet are for.
   A step-by-step update guide can hardly explain the concept.
1. Clear input data source
   Always listing the source file location or scripts allow the data to be traced.
1. Clean data with proper data format layout in the worksheets
   Data quality is the key. Bad data formats require numerous tweaks. It is important to clean the data first before building things on them.
1. No obscure? layout, e.g. hidden worksheets, groups of rows and columns, blocks of data at the right or bottom off the screen
1. Concise and clear file names and version control
1. Flexible for spreadsheet enhancement
   Always starting from the lowest level data to the final presentational summary (bottom-up) makes the entire spreadsheet easy to tweak.
1. Standalone spreadsheets
   Workbooks should not have external links because there are always risks that the external workbook is modified.

# Simple rule of thumb: 0.5-second limit on calculation speed

Many are used to turn on the Manual Calculation Mode in Excel, because the spreadsheet calculates slowly. In my experience, I have worked through a lot of these Excels, and have revamped many of these.

In most of the cases, reducing the calculation time from 10 seconds to 0.2 seconds can be done within several hours of spreadsheet revamp. If I am to give a single piece of the most important advice on this, it will have to be:

## Never use `INDIRECT` in formula.

There are other techniques which help keeps a minimal calculation speed. It is necessary to stress again that they are related to spreadsheet design habits, and no sophisticated knowledge is necessary.

Important Note

Upgrading your computer with a faster CPU is not a solution - it will not improve the spreadsheet calculationi speed significantly considering the current CPU speed.

# There is a list of best practice on each Excel feature

Since Excel is a heavily used software with tons of features, it will be necessary to visit the different features one by one. Later blog posts will explore the following topics.

- Good documentation
- Use of VLOOKUP
- Use of SUMIF
- Use of colors
- Use of VBA
- Dealing with Dates
- Data management in Excel
- File management and version control
- Avoiding volatile functions such as INDIRECT
- And more!

# Good tech skills are also necessary on the other MS products as well as Windows

While Excel is the most commonly used software, actuaries also work with MS Outlook, MS Word, MS Powerpoint as well as database management softwares frequently.

A good command of corresponding skills ensure an efficient work flow process. Using the right tool for the job is the fundamental to boosting the work efficiency.

Actuaries work with computers every day. It will be sad to have the application of the actuarial skillset hindered by the limitations on one's technological ability.
