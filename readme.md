# Excellent at Excel

So, you want to "*get good*" at Excel. Congratulations. This is a list of some shortcuts and formulas that are going to set you on the path of being most excellent and Excel modeling. Example files are attached within this repo. If you have any unique uses of formulas or advance tips, drop me a line over in the _______ section of this repo. 

Based on the fact you are using github, I am assuming you have a working knowledge of Excel. If that's not the case, check out LinkedIn's [Excel 2016 Essential Training](https://www.linkedin.com/learning/excel-2016-essential-training/welcome) course and come back after.

Let's jump right into the good stuff:

## 01. Beginner Lookups: `VLOOKUP` and `HLOOKUP`
Want to really impress people that don't know anything about Excel? Tell them about `VLOOKUP`. Want to doubly-impress them? Tell them about `HLOOKUP`. Want to make yourself look foolish to anybody that knows anything about Excel? Brag to them about knowing the `VLOOKUP` and `HLOOKUP` functions. Let's move on. [Check out the file](https://github.com/castlenthesky/excellent_at_excel/blob/master/01.%20Begginer%20Lookups.xlsx) if you need to.

## 02. Intermediate Lookups: `INDEX`, `OFFSET` and `MATCH`
What if you want to find a cell, and then return the value from the cell to the **LEFT**? This is the kind of situation that makes `VLOOKUP` look like a bigger disappointment than I was to my parents. This is because `VLOOKUP` can only return values to the **RIGHT** of a matched value. `INDEX`, `MATCH`, and `OFFSET` are all mixed and matched together in the [Intermediate Lookups file](https://github.com/castlenthesky/excellent_at_excel/blob/master/02.%20Intermediate%20Lookups.xlsx) to demonstrate how to overcome `VLOOKUP`'s shortcomings.

### `INDEX(MATCH())`
Pairing the `INDEX` and `MATCH` functions together allows you to return cells to the **LEFT** of your match. It's a pretty handy combination and the example file holds a simple demonstration for your reference.

### `OFFSET(MATCH())`

The `OFFSET` function is probably my favorite function in Excel. I get a little excited when I get to use it to solve an actual problem.

![alt text](https://i.imgflip.com/3cqujr.jpg "Car Salesman Meme")

But seriously - `OFFSET` is SOOOOO useful is SOOOO many sitiuations. It not only allows you to lookup a value in a table, but also allows you the ability to dynamically return a RANGE of cells. "Who cares," you ask? Wouldn't it be nice to easily pull the average movement of a product in the 3 months prior to your promotion so you can evaluate your promotional lift and subsequent cannibalization? Yes, it would be nice, and `OFFSET` makes this a piece of cake. Check out the example file to see how it's done.

## 03. Advanced Lookups: `INDIRECT`
So... now you have some P&L statements which are most likely stored in some really bad format. Somebody probably designed the file where each year is on its own tab, or worse yet: each month has its own tab. This is where `INDIRECT` really shines!

The [example file](https://github.com/castlenthesky/excellent_at_excel/blob/master/03.%20Advanced%20Lookups.xlsx) shows how to use `INDIRECT` to collect annual sales and COGS from different tabs in a summary table. It's a lot easier than clicking back and forth between tabs 90 times. It also allows your models to be dynamic enough to allow for future expansion with little more than a copy and paste.

## 04. Useful Uses of Useful Functions
There are a few functions in Excel that are really under-utilized, or which can be used in some non-traditional ways. [Check out the example file](https://github.com/castlenthesky/excellent_at_excel/blob/master/04.%20Useful%20Uses%20of%20Useful%20Functions.xlsx) to see how these functions can upgrade your models.
### `SUMPRODUCT`
Good for weighted averages. Also good for counting distinct values in a column and conditional ranking.

##### Counting Distinct Values
Google Sheets has a `COUNTDISTINCT` function that's pretty handy. Excel doesn't. Let's make one.

`=SUMPRODUCT((1/COUNTIF([RANGE_OF_VALUES],[RANGE_OF_VALUES]&"")))`

##### Conditional Ranking
SUM   --> SUMIF   --> SUMIFS

COUNT --> COUNTIF --> COUNTIFS

RANK  --> ***?!?!?!***

Sometimes it is nice to know the rank of an item among a subset of other similar items. For exampe: I have the monthly sales of a given set of products spread across various areas of the business or within different geographic regions. I don't only want to know which product is the largest overall, but the largest in a given region and product category. Thankfully, we can do this with the `SUMPRODUCT` function alone. Check out the example file.

## 05. Data Connections
OK. You know what really rustles my jimmies? Having to access some system, download a file, open that file, copy the data, and THEN past it into my analyses. WE LIVE IN THE 21st CENTURY! [JUST MAKE A DATA CONNECTION!](https://github.com/castlenthesky/excellent_at_excel/blob/master/05.%20Data%20Connections.xlsx)

### Basic SQL Connections
Have a monthly report *\*\*\*COUGH\*\*\** ***P&L*** *\*\*\*COUGH\*\*\** that you need to create with data you source from your corporate database? Monthly variances to breakdown? Wire Excel up to your database. Push `Refresh` on the Data ribbon and be done with it!

### Dynamic SQL Connections
Even better than wiring up a query to pull your data: Wiring up a dynamic query with variables. Microsoft hasn't made this easy or clean, but man it can be useful for some awesome models.




