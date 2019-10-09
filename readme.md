# Excellent at Excel

So, you want to "*get good*" at Excel. Congratulations. This is a list of some shortcuts and formulas that are going to set you on the path of being most excellent and Excel modeling. Example files are attached within this repo. If you have any unique uses of formulas or advance tips, drop me a line over in the _______ section of this repo. 

Based on the fact you are using github, I am assuming you have a working knowledge of Excel. If that's not the case, check out LinkedIn's [Excel 2016 Essential Training](https://www.linkedin.com/learning/excel-2016-essential-training/welcome) course and come back after.

Let's jump right into the good stuff:

## Data Connections
OK. You know what really rustles my jimmies? Having to access some system, download a file, open that file, copy the data, and THEN past it into my analyses. WE LIVE IN THE 21st CENTURY! JUST MAKE A DATA CONNECTION!

Got a monthly report *\*\*\*COUGH\*\*\** ***P&L*** *\*\*\*COUGH\*\*\** that you need to create? Monthly variances to breakdown? Wire that Excel up to your database. Push `Refresh` on the Data ribbon and be done with it!

### Dynamic SQL Connections
Even better than wiring up a query to pull your data: Wiring up a dynamic query with variables. Microsoft hasn't made this easy or clean, but man it can be useful for some awesome models.

## Useful Function Uses

### Lookups
Ok. So there are somewhere in the ballpark of 10 metric-tons' worth of ways to do lookups. Let's break the most important ones down:

#### Beginner: `VLOOKUP` OR `HLOOKUP`
Want to really impress people that don't know anything about Excel? Tell them about `VLOOKUP`. Want to doubly-impress them? Tell them about `HLOOKUP`. Want to make yourself look foolish to anybody that knows anything about Excel? Brag to them about knowing the `VLOOKUP` and `HLOOKUP` functions. Let's move on.

#### Intermediate: `INDEX(MATCH())`
You are an **INTEREMEDIATE** user and have more discerning tastes in how you build your Excel models. Therefore, you have outgrown `VLOOKUP`. But what if you want to find a cell, and then return the value from the cell to the **LEFT**? This is the kind of situation that makes `VLOOKUP` look like a bigger disappointment than I was to my parents. `VLOOKUP` can only return values to the **RIGHT** of a matched value. Pairing the `INDEX` and `MATCH` functions together allows you to return cells to the **LEFT** of your match.

#### Advanced: `OFFSET` & `MATCH`
The `OFFSET` function is probably my favorite function in Excel. I might actually get a little excited when I get to use it to solve an actual problem.
![alt text](https://i.imgflip.com/3cqujr.jpg "Car Salesman Meme")


#### *MORE* ADVANCEDER: `INDIRECT`
So... now you have some P&L statements which are most likely stored in some really bad format. Somebody probably designed the file where each year is on its own tab, or worse yet: each month has its own tab. This is where `INDIRECT` really shines!

### `SUMPRODUCT`
Good for weighted averages. Also good for counting distinct values in a column and conditional ranking.

#### Counting Distinct Values
Google Sheets has a `COUNTDISTINCT` function that's pretty handy. Excel doesn't. Let's make one.
`=SUMPRODUCT((1/COUNTIF([RANGE_OF_VALUES],[RANGE_OF_VALUES]&"")))`


#### Conditional Ranking
SUM   --> SUMIF   --> SUMIFS
COUNT --> COUNTIF --> COUNTIFS
RANK  --> ***?!?!?!***

Sometimes it is nice to know the rank of an item among a subset of other similar items. For exampe: I have the monthly sales of a given set of products spread across various areas of the business or within different geographic regions. I don't only want to know which product is the largest overall, but the largest in a given region and product category. Believe it or not, we can do this with the `SUMPRODUCT` function alone.

