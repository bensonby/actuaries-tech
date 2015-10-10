title: Why are Excel Workbooks difficult to follow and check?

post-url: why-are-excel-workbooks-difficult-to-follow-and-check

# Introduction

Excel aims to be an easy-to-use spreadsheet software. This is pretty straightforward even for novice users. I am sure they find it easy enough to create spreadsheets with Excel.

However people (actuaries) start to build insanely complicated workbooks with complex formulae. This is very common In the actuarial industry: how long does it take to understand a workbook?

In this article I list several reasons that make an Excel workbook difficult to understand.

# The Problem

It is difficult for anyone to understand a new workbook. Please do not treat this as normal. This problem lies in the root objective of Excel: “being easy to use”. The complex and hard-to-read workbooks are simply the trade-offs. Excel workbooks act like computer programs, but, are hard to be as good as them.

1. Sequence of logic is not visible

 Presented as a grid with cell addresses, Excel allows us to enter formula in any cell. We need not specify the order of calculations: Excel does it automatically, magically, and invisibly for you. It becomes the responsibility of the users to interpret the overall logic flow. One must study the workbook thoroughly to understand it.

 The invisible logical flow deviates from a normal programming languages which have clear execution order. With plenty of worksheets and cells, they take hours for first-time users to read. Although we can understanding the formula of all the cells, we will not be able to recall their calculation sequences and cell dependencies quickly.

1. No control on variable scopes

 If you have learned basic programming, you must have read the advice “do not use global variables”. It is related to variable scope. Each variable should only “live” in its own world so readers know where they are used.

 In Excel, however, there are no such concepts being emphasized. Any worksheet or even external workbooks can be used as an input to a formula. Tracing the usage of variables is so difficult that the "tracing precedents/dependents" feature can barely help.
 
1. Excel formula are poorly readable and not comprehensible

 Formula often look like dealing with memory addresses. You know what is going on only after referring to each of the cell.
`(I7+D8-E8)*((1+$B$3)^(1/12)-1)`

 Why not like this?
`(ending_balance[t-1]+premium[t]-commission[t])*((1+interest_rate)^(1/12)-1)`

 Function names are not meaningful. We are building cells with the elementary functions. Let’s try to guess what this formula means: (The cell `B1` contains a string)

 `LEN(B1)-LEN(SUBSTITUTE(SUBSTITUTE(B1,"B",""),"b",""))`

 This formula counts the number of “B”s (case-insensitive) in the cell `B1`. Why not use a function instead? `count_char(B1, “b”)`

 It is much more intuitive and easier to understand.

 _Note: the proposed improvements only intend to show the bad comprehensibility of the formula. The proposed alternatives do not solve the problem perfectly because of other fundamental problems in Excel._

1. Irregular and Restricted Layout

 Blocks are put anywhere, in any worksheet. There can be unnoticeable checking cells. There can also be blocks that are far away from the visible screen, e.g. some people try to avoid pivot table collision by putting a second pivot table far on the right in a worksheet.

 A programming principle says layout should be separated from logic. In the world of Excel, they are totally stuck together. The most serious examples are pivot tables and data tables, where both of them have to be in specific formats defined by Microsoft Excel.

1. Poorly written VBA programs

 When some tasks are too difficult to be done with simple Excel formula, we often choose VBA as the solution. Yet we, as actuaries, are not technically trained to write programs: Getting programs to work is not difficult, but making the program easily readable, maintainable and extensible requires SKILLS. Most actuaries do not possess the required level of skills.

 The lack of programming skill is “kindly” supplemented by Excel VBA’s unfriendly error descriptions: how many times have you encountered `1004 Application-defined or object-defined error` and `Type Mismatch`?

 We end up with numerous horrible VBA programs – the more complicated the program is, the worse the VBA code quality is. Less than 1% of actuarial programmers can write good VBA programs.

# Conclusion
Excel is easy to learn but hard to master and control. Its strengths lead to its weaknesses. Certainly there are ways to ease the negative impacts brought by the above. But please keep in mind that part of them are unresolvable – they are the “nice features” that make Excel “user-friendly”.

In my other articles, I will explain some design principles that can ease the drawbacks.

