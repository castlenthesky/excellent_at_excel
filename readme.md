# Excellent at Excel

---
So, you want to "*get good*" at Excel. Congratulations. This is a list of some shortcuts and formulas that are going to set you on the path. Example files are attached here in this repo. 

Based on the fact you are using github, these templates are presupposing you have a working knowledge of Excel. If that's not the case, check out LinkedIn's [Excel 2016 Essential Training](https://www.linkedin.com/learning/excel-2016-essential-training/welcome) course and come back after.

Let's jump right into the good stuff:

---

## `SUMPRODUCT`
Good for weighted averages. Also good for counting distinct values in a column and conditional ranking.

### Counting Distinct Values
Google Sheets has a `COUNTDISTINCT` function like SQL. Excel doesn't. Let's make one.

### Conditional Ranking
SUM   --> SUMIF   --> SUMIFS
COUNT --> COUNTIF --> COUNTIFS
RANK  --> ***?!?!?!***

Let's add some conditioning in our ranking formula so we can rank within specific categories.

---
## Lookups
Ok. So there are somewhere in the ballpark of 10 metric-tons' worth of ways to do lookups. Let's break the most important ones down:
### Beginner: `VLOOKUP` OR `HLOOKUP`
Want to really impress people that don't know anything about Excel? Tell them about `VLOOKUP`. Want to doubly-impress them? Tell them about `HLOOKUP`. Want to make yourself look foolish to anybody that knows anything about Excel? Brag to them about knowing the `VLOOKUP` and `HLOOKUP` functions. Let's move on.

### Intermediate: `INDEX(MATCH())`
You are an **INTEREMEDIATE** user and have more discerning tastes in how you build your Excel models. Therefore, you have outgrown `VLOOKUP`. But what if you want to find a cell, and then return the value from the cell to the **LEFT**? This is the kind of situation that makes `VLOOKUP` look like a bigger disappointment than I was to my parents. `VLOOKUP` can only return values to the **RIGHT** of a matched value. Pairing the `INDEX` and `MATCH` functions together allows you to return cells to the **LEFT** of your match.

### Advanced: `OFFSET` & `MATCH`
The `OFFSET` function is probably my favorite function in Excel.
![alt text](https://i.imgflip.com/3cqujr.jpg "Car Salesman Meme")


### *MORE* ADVANCEDER: `INDIRECT`
So... now you have some P&L statements which are most likely stored in some really bad format. Somebody probably designed the file where each year is on its own tab, or worse yet: each month has its own tab.
