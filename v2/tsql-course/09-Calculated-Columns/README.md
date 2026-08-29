# Video 09 — Calculated Columns in T-SQL

**YouTube:** Calculated Columns in T-SQL — Build New Data from Existing Data | Video 09 of 36

**Module 2: Selecting and Shaping Data**

In this video, we look at calculated columns: expressions written directly in the `SELECT` list that create new values from existing table data at query time.

You will learn how to perform arithmetic inside a query, assign readable aliases to calculated columns, combine text using `CONCAT()`, and take a first look at date calculations using `GETDATE()` and `DATEDIFF()`.

We also finish with an important gotcha: integer division, where a calculation can return a perfectly valid-looking number that is actually wrong.

---

## 1. What Is a Calculated Column?

A calculated column is an expression written inside the `SELECT` list.

SQL Server evaluates that expression for every row and returns the calculated value alongside the existing table columns.

The underlying table is not changed.

For example:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       ListPrice
FROM Production.Product;
```

This query returns existing columns directly from the table.

Now we can add a calculated column.

---

# Calculating a VAT-Inclusive Price

Suppose we want to show each product's list price with 20 percent tax added.

We can multiply `ListPrice` by `1.20`.

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       ListPrice,
       ListPrice * 1.20 AS PriceWithVAT
FROM Production.Product;
```

The expression:

```sql
ListPrice * 1.20
```

is evaluated separately for every row.

The alias:

```sql
AS PriceWithVAT
```

gives the calculated column a readable name in the result set.

Without an alias, SQL Server would still perform the calculation, but the output column would not have a useful business-friendly heading.

---

# Filtering Before the Calculation

We can also filter the rows used by the query.

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       ListPrice,
       ListPrice * 1.20 AS PriceWithVAT
FROM Production.Product
WHERE ListPrice > 0;
```

This removes products where the list price is zero.

It can also help prevent unwanted results where NULL values are involved.

As covered earlier in the course, arithmetic involving NULL normally produces NULL.

For example:

```text
NULL * 1.20
```

returns:

```text
NULL
```

If you wanted to keep rows with missing values but substitute zero for the calculation, you could use:

```sql
ISNULL(ListPrice, 0) * 1.20
```

The correct choice depends on the business requirement.

---

# Adding More Than One Calculated Column

A query can contain multiple calculated columns.

For example, we can calculate both a VAT-inclusive price and a simple margin estimate.

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       ListPrice,
       StandardCost,
       ListPrice * 1.20 AS PriceWithVAT,
       ListPrice - StandardCost AS MarginEstimate
FROM Production.Product
WHERE ListPrice > 0;
```

Here:

```sql
ListPrice * 1.20 AS PriceWithVAT
```

creates the VAT-inclusive price.

And:

```sql
ListPrice - StandardCost AS MarginEstimate
```

calculates the difference between list price and standard cost.

A calculated value can be positive, zero, or negative depending on the source data.

---

# Arithmetic Operators in T-SQL

T-SQL supports the standard arithmetic operators.

```text
+   Addition
-   Subtraction
*   Multiplication
/   Division
```

These work much like arithmetic formulas in Excel.

For example:

```sql
ListPrice + 10
```

```sql
ListPrice - StandardCost
```

```sql
ListPrice * 1.20
```

```sql
ListPrice / 2
```

Parentheses can be used to control the order of operations.

For example:

```sql
(ListPrice - StandardCost) * 1.20
```

The calculation inside the parentheses is performed first.

---

# Calculated Columns Are Not Stored

A calculated column created in a `SELECT` query does not change the table.

For example:

```sql
ListPrice * 1.20 AS PriceWithVAT
```

does not create a permanent `PriceWithVAT` column in `Production.Product`.

The calculation exists only in the result returned by the query.

Each time the query runs, SQL Server recalculates the value using the current table data.

For reporting and analysis, this is often exactly what you need.

---

# A First Look at CONCAT()

Calculated columns are not limited to numbers.

You can also create new text values.

T-SQL provides the `CONCAT()` function for joining values together.

For example, `Person.Person` contains separate first-name and last-name columns.

```sql
USE AdventureWorks2025;

SELECT BusinessEntityID,
       FirstName,
       LastName,
       CONCAT(FirstName, ' ', LastName) AS FullName
FROM Person.Person;
```

The expression:

```sql
CONCAT(FirstName, ' ', LastName)
```

combines:

```text
FirstName
+
a space
+
LastName
```

into one value.

For example:

```text
John Smith
```

The single space is included as a string literal:

```sql
' '
```

---

# CONCAT() and NULL Values

`CONCAT()` has useful behaviour when NULL values are involved.

It treats NULL as an empty string rather than allowing one NULL argument to make the entire result NULL.

It can also automatically convert non-text values to text when necessary.

This makes it useful for building display values from multiple columns.

String functions are covered in more detail later in the course.

---

# A First Look at GETDATE()

Calculated columns can also use functions.

`GETDATE()` returns the current date and time from SQL Server.

```sql
SELECT GETDATE() AS Today;
```

It takes no arguments.

Each time the query runs, SQL Server returns the current server date and time.

---

# A First Look at DATEDIFF()

`DATEDIFF()` calculates the difference between two dates.

Its basic structure is:

```text
DATEDIFF(unit, start_date, end_date)
```

For example, we can estimate an employee's years of service.

```sql
USE AdventureWorks2025;

SELECT BusinessEntityID,
       HireDate,
       GETDATE() AS Today,
       DATEDIFF(year, HireDate, GETDATE()) AS YearsOfService
FROM HumanResources.Employee
ORDER BY HireDate ASC;
```

The expression:

```sql
DATEDIFF(year, HireDate, GETDATE())
```

contains three arguments:

```text
year       → Unit to measure
HireDate   → Start date
GETDATE()  → End date
```

The result is returned as an integer.

---

# Changing the DATEDIFF Unit

The same function can measure different time units.

For example:

```sql
DATEDIFF(year, HireDate, GETDATE())
```

measures year boundaries.

You could change the unit to:

```sql
DATEDIFF(month, HireDate, GETDATE())
```

or:

```sql
DATEDIFF(day, HireDate, GETDATE())
```

The start and end dates stay the same.

Only the unit changes.

Date functions are covered in more detail later in the course.

---

# GOTCHA — Integer Division

One of the easiest mistakes to miss in a calculated column is integer division.

Consider:

```sql
USE AdventureWorks2025;

SELECT 7 / 2 AS WrongResult;
```

You might expect:

```text
3.5
```

But SQL Server returns:

```text
3
```

Both `7` and `2` are integers.

When SQL Server divides one integer by another integer, it performs integer division.

Any remainder is discarded.

There is:

* No error
* No warning
* No indication that precision has been lost

That is what makes this mistake easy to overlook.

---

# Why Integer Division Is Dangerous

The result:

```text
3
```

looks like a completely valid number.

In a large result set containing ratios, percentages, averages, or margins, the error may not be obvious.

For example, calculations such as:

```text
Sales / Quantity
Margin / Revenue
Completed / Total
```

can produce misleading results if both source columns are integer types.

---

# Fixing Integer Division

To force decimal arithmetic, at least one side of the calculation needs to be treated as a decimal value.

A simple technique is to multiply one side by `1.0`.

```sql
USE AdventureWorks2025;

SELECT 7 * 1.0 / 2 AS CorrectResult;
```

Now SQL Server returns:

```text
3.5
```

The expression:

```sql
7 * 1.0
```

introduces a decimal value before the division happens.

SQL Server therefore performs decimal arithmetic instead of integer arithmetic.

---

# Integer Division with Table Columns

The same rule applies when calculations use table columns.

Suppose both columns are integer types.

This:

```sql
ColumnA / ColumnB
```

may perform integer division.

A common defensive pattern is:

```sql
ColumnA * 1.0 / ColumnB
```

This forces decimal arithmetic.

For example:

```sql
SalesAmount * 1.0 / Quantity
```

If a calculated percentage, ratio, or average looks suspiciously rounded, check the data types of the columns involved.

---

# Complete Reference Script

Below is a consolidated script containing the main examples from this video.

```sql
USE AdventureWorks2025;


/* =========================================================
   EXISTING COLUMNS
   ========================================================= */

SELECT ProductID,
       Name,
       ListPrice
FROM Production.Product;


/* =========================================================
   CALCULATED COLUMN
   VAT-INCLUSIVE PRICE
   ========================================================= */

SELECT ProductID,
       Name,
       ListPrice,
       ListPrice * 1.20 AS PriceWithVAT
FROM Production.Product;


/* =========================================================
   CALCULATED COLUMN WITH FILTER
   ========================================================= */

SELECT ProductID,
       Name,
       ListPrice,
       ListPrice * 1.20 AS PriceWithVAT
FROM Production.Product
WHERE ListPrice > 0;


/* =========================================================
   MULTIPLE CALCULATED COLUMNS
   ========================================================= */

SELECT ProductID,
       Name,
       ListPrice,
       StandardCost,
       ListPrice * 1.20 AS PriceWithVAT,
       ListPrice - StandardCost AS MarginEstimate
FROM Production.Product
WHERE ListPrice > 0;


/* =========================================================
   OPTIONAL NULL HANDLING

   Use when the business requirement is to treat a missing
   ListPrice as zero.
   ========================================================= */

SELECT ProductID,
       Name,
       ListPrice,
       ISNULL(ListPrice, 0) * 1.20 AS PriceWithVAT
FROM Production.Product;


/* =========================================================
   CONCAT()
   BUILD A FULL NAME
   ========================================================= */

SELECT BusinessEntityID,
       FirstName,
       LastName,
       CONCAT(FirstName, ' ', LastName) AS FullName
FROM Person.Person;


/* =========================================================
   GETDATE() AND DATEDIFF()
   YEARS OF SERVICE
   ========================================================= */

SELECT BusinessEntityID,
       HireDate,
       GETDATE() AS Today,
       DATEDIFF(year, HireDate, GETDATE()) AS YearsOfService
FROM HumanResources.Employee
ORDER BY HireDate ASC;


/* =========================================================
   DATEDIFF() USING MONTHS
   ========================================================= */

SELECT BusinessEntityID,
       HireDate,
       DATEDIFF(month, HireDate, GETDATE()) AS MonthsOfService
FROM HumanResources.Employee;


/* =========================================================
   DATEDIFF() USING DAYS
   ========================================================= */

SELECT BusinessEntityID,
       HireDate,
       DATEDIFF(day, HireDate, GETDATE()) AS DaysOfService
FROM HumanResources.Employee;


/* =========================================================
   INTEGER DIVISION GOTCHA

   Both values are integers.
   SQL Server returns 3 rather than 3.5.
   ========================================================= */

SELECT 7 / 2 AS WrongResult;


/* =========================================================
   FORCE DECIMAL ARITHMETIC

   Multiplying one side by 1.0 introduces a decimal value.
   ========================================================= */

SELECT 7 * 1.0 / 2 AS CorrectResult;
```

---

# Key Takeaways

By the end of this video, you should be comfortable with:

* Understanding what a calculated column is
* Writing expressions directly in the `SELECT` list
* Creating new values from existing table columns
* Using aliases with calculated columns
* Performing addition, subtraction, multiplication, and division
* Using parentheses to control arithmetic order
* Understanding that calculated columns do not modify the underlying table
* Recognising that calculated values are recomputed when the query runs
* Understanding how NULL can affect calculations
* Using `ISNULL()` where the business requirement calls for a replacement value
* Combining text using `CONCAT()`
* Understanding how `CONCAT()` handles NULL values
* Using `GETDATE()` to retrieve the current server date and time
* Using `DATEDIFF()` to measure the difference between dates
* Changing the unit used by `DATEDIFF()`
* Recognising the integer division problem
* Understanding why integer division can silently produce incorrect results
* Using `1.0` to force decimal arithmetic when necessary

---

## Next Video

**Video 10 — GROUP BY: Summarising Your Data**

In the next video, we move from returning individual rows to summarising groups of rows.

We will start using aggregate functions with `GROUP BY` to answer business questions such as totals, counts, averages, minimums, and maximums across different categories of data.
