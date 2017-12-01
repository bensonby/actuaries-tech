title: 5 Reasons Why Excel Workbooks Are Difficult to Read

# Introduction

Excel aims to be an easy-to-use spreadsheet software. The learning curve is not steep even for novice users. I am sure they find it easy enough to create spreadsheets with Excel.

Yet it is another story when people (actuaries) start to build insanely complicated workbooks with complex formulae. In the actuarial industry, we often ask ourselves: how long does it take to understand this bloody workbook?

In this article I list several reasons that make an Excel workbook difficult to understand.

# The Problem

In most of the cases, it is difficult for anyone to understand a new workbook. Please do not treat this as a normal phenomenon. The problem lies in the root objective of Excel: "being easy to use". The complex and hard-to-read workbooks are thus the trade-offs. Excel workbooks act like computer programs, but, are hard to be as good as them.

## 1. Sequence of logic is not visible

Presented as a grid with cell addresses (`A1`, `C100`, `DE1234`, etc.), Excel allows us to enter a formula in any cell. We need not specify the order of calculations: Excel does it automatically, magically, and invisibly for you. It becomes the responsibility of the users to interpret the overall logic flow. One must study the workbook thoroughly to understand it.

The invisible logical flow deviates from a normal programming languages which have clear execution order. It can take hours for a first-time user to understand a workbook with plenty of worksheets and cells. Although we can understand the formula of all the cells, we will not be able to recall their calculation sequences and cell dependencies quickly.

## 2. No control on variable scopes

If you have learned basic programming, you must have read the advice “do not use global variables”. It is related to variable scopes. Each variable should only "live" in its own world to limit its effect on other parts of the world, i.e. readers can easily know the extent the variable is used.

In Excel, however, such concepts are not emphasized. A cell can refer to any worksheet or even external workbooks. Tracing the usage of variables is so difficult that sometimes even the "tracing precedents/dependents" feature can barely help.
 
## 3. Excel formulae are poorly readable and not comprehensible

Formula often look like dealing with computer memory addresses. You know what is going on only by reading each of the dependent cell.
`(I7+D8-E8)*((1+$B$3)^(1/12)-1)`

Why not like this?
`(ending_balance[t-1]+premium[t]-commission[t])*((1+interest_rate)^(1/12)-1)`

Function names are also not meaningful. We are building cells with the elementary functions. Let's try to guess what this formula means: (The cell `B1` contains a string)

`LEN(B1)-LEN(SUBSTITUTE(SUBSTITUTE(B1,"B",""),"b",""))`

Answer: This formula counts the number of “B”s (case-insensitive) in the cell `B1`.

Why not use a function instead? `count_char(B1, “b”)`

That would be much more intuitive and easier to understand.

_Note: the proposed improvement only intends to show the bad comprehensibility of the formula. The proposed alternative does not solve the problem perfectly because of other fundamental problems in Excel._

## 4. Irregular and Restricted Layout

Data blocks can be put everywhere, in any worksheet. There can be unnoticeable checking cells. There can be hidden cells too. There can also be blocks that are far away from the visible screen, e.g. some people try to avoid pivot table collision by putting a second pivot table far on the right in a worksheet.

[This programming principle][soc] states that layout should be separated from logic. In the world of Excel, they are totally stuck together. The most obvious examples are pivot tables and data tables, where both of them have to be in the specific formats defined by Microsoft Excel.

## 5. Poorly written VBA programs

When a task is too difficult to be done with simple Excel formula, we often choose VBA as the solution. Yet we, as actuaries, are not technically trained to write programs: It is not difficult to create a program that works; but creating a program that is easily readable, maintainable and extensible requires a lot of skills. Most actuaries do not possess such level of skills.

The lack of programming skill is "kindly" supplemented by Excel VBA’s unfriendly error descriptions. Do you remember the number of times you have seen `1004 Application-defined or object-defined error` or `Type Mismatch`?

We end up with numerous horrible VBA programs – the more complicated the program is, the worse the VBA code quality is. I believe that less than 1% of actuarial programmers can write good VBA programs. And they become nightmare for people (in most cases even the writer itself) to maintain and debug these programs.

# Conclusion

Excel is easy to learn but hard to master and control. Its strengths lead to its weaknesses. Certainly there are ways to ease the negative impacts brought by the above. But please keep in mind that some of the weaknesses cannot be perfectly solved – because they are the "nice features" that make Excel so "user-friendly".

In my other articles, I will explain some design principles that can ease the drawbacks.

[soc]: https://en.wikipedia.org/wiki/Separation_of_concerns
