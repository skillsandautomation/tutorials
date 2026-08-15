# Video 04 — Learn the SELECT Statement in T-SQL

**YouTube:** Learn the SELECT Statement in T-SQL — Write, Run & Understand Your First Queries | Video 04 of 36

**Module 2: Selecting and Shaping Data**

In this video, we start writing T-SQL queries using the `SELECT` statement.

You will learn how `SELECT` and `FROM` work together, how to reference tables correctly, how to explore unfamiliar tables with `SELECT *`, how to read common errors in SQL Server Management Studio, and how to make your queries easier to read with formatting, comments, and column aliases.

We also look at a practical workflow for writing queries in SSMS, copying results into Excel, and using AI tools to help explain, troubleshoot, and clean up SQL.

---

## 1. Your First SELECT Query

A basic SELECT query contains two important clauses:

* `SELECT` tells SQL Server which columns you want returned.
* `FROM` tells SQL Server which table those columns should come from.

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM Production.Product;
```

### What each part does

```sql
USE AdventureWorks2025;
```

Sets `AdventureWorks2025` as the current database.

```sql
SELECT ProductID,
       Name,
       ListPrice,
       Color
```

Specifies the columns that should appear in the result set.

```sql
FROM Production.Product;
```

Specifies the table containing the data.

`Production` is the schema and `Product` is the table.

---

## 2. Always Reference the Schema

When referring to a table, it is good practice to include the schema name.

Correct:

```sql
SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM Production.Product;
```

If you leave out the schema:

```sql
SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM Product;
```

SQL Server may return an error such as:

```text
Invalid object name 'Product'.
```

SQL Server normally attempts to resolve an unqualified table name using your default schema.

Using the schema explicitly removes ambiguity.

---

## 3. Selecting the Correct Database

If you open a new query window without selecting `AdventureWorks2025`, your query may be running against another database such as `master`.

One solution is to explicitly set the database:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM Production.Product;
```

You can also select `AdventureWorks2025` from the database dropdown in SSMS.

Another useful workflow is:

1. Expand **Databases** in Object Explorer.
2. Find `AdventureWorks2025`.
3. Right-click the database.
4. Select **New Query**.

The query window will automatically use that database.

---

## 4. Let SSMS Generate a Starting Query

When working with an unfamiliar table, SSMS can generate a query for you.

In Object Explorer:

1. Expand `AdventureWorks2025`.
2. Expand **Tables**.
3. Find `Production.Product`.
4. Right-click the table.
5. Select **Select Top 1000 Rows**.

SSMS generates a query containing the table and its columns.

This is useful when you do not know the exact table or column names yet.

---

# Query vs Statement

These terms are closely related, but they are not exactly the same thing.

## Statement

A **statement** is a complete instruction sent to SQL Server.

Examples include:

### Data Manipulation

```sql
SELECT
INSERT
UPDATE
DELETE
```

### Database Structure

```sql
CREATE TABLE
ALTER DATABASE
DROP INDEX
```

### Access Control

```sql
GRANT
REVOKE
```

### Flow Control

```sql
IF...ELSE
WHILE
BEGIN...END
```

## Query

A query is a type of statement whose purpose is to retrieve information from the database.

In T-SQL, queries are primarily built around the `SELECT` statement and clauses such as:

```sql
WHERE
JOIN
GROUP BY
ORDER BY
```

Most of this course focuses on retrieving and analysing data.

Later in the course we will also cover statements that modify data, including:

```sql
INSERT
UPDATE
DELETE
```

---

# SELECT *

The asterisk means:

> Return every column from the table.

```sql
USE AdventureWorks2025;

SELECT *
FROM Production.Product;
```

This is extremely useful when exploring a table for the first time.

You can quickly see:

* What columns exist
* What the data looks like
* Which columns contain values
* Which columns contain `NULL`
* The general structure of the table

---

## Why SELECT * Can Be Problematic

`SELECT *` is useful for exploration, but it is normally better to explicitly name the columns in queries that will be saved, shared, or used by downstream systems.

Instead of:

```sql
SELECT *
FROM Production.Product;
```

Prefer:

```sql
SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM Production.Product;
```

This makes the query more predictable and clearly documents which fields are actually required.

---

# A Practical Workflow for Writing SELECT Queries

When exploring a table, a useful starting point is:

```sql
SELECT *
FROM Production.Product;
```

Run the query first and inspect the table.

Then replace the `*` with the columns you actually need:

```sql
SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM Production.Product;
```

Once SSMS knows which table you are using, IntelliSense can help suggest the available column names.

This reduces typing and helps avoid mistakes.

---

# Reading Errors in SSMS

Mistakes are normal when writing SQL.

SSMS provides several ways to identify them.

---

## Example: Incorrect Column Name

Suppose we accidentally type:

```sql
USE AdventureWorks2025;

SELECT ProductD,
       Name,
       ListPrice,
       Color
FROM Production.Product;
```

`ProductD` does not exist.

SSMS may show a red squiggly underline beneath the invalid column name.

If the query is executed, SQL Server will return an error in the **Messages** tab.

---

## Example: Extra Comma

Consider this query:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       ListPrice,
       Color,
FROM Production.Product;
```

The problem is the comma after `Color`.

SQL Server may report the error near:

```sql
FROM
```

even though the actual mistake is on the previous line.

This is an important debugging lesson:

> The line SQL Server reports is where the parser realised something was wrong. It is not always the exact location where the mistake was made.

When troubleshooting syntax errors, inspect the surrounding lines as well.

---

# Using AI to Troubleshoot SQL Errors

If an error is not obvious, AI tools can help explain it.

A useful prompt is:

```text
I am trying to run the following T-SQL query:

[Paste your query]

I am getting the following error:

[Paste the SQL Server error message]

What is the issue?
```

Do not blindly copy the response.

Use the explanation to understand the problem, then correct and test the query yourself.

---

# Writing Clean T-SQL

SQL Server does not require beautifully formatted code.

This query works:

```sql
select productid, name, listprice, color from production.product;
```

But it is much easier to read like this:

```sql
SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM Production.Product;
```

Both queries return the same result.

The formatting is for humans, not SQL Server.

---

## Formatting Conventions Used in This Course

Throughout the course we will generally use:

* SQL keywords in uppercase
* Table and column names using their database casing
* One selected column per line
* Consistent indentation
* Schema-qualified table names
* Semicolons at the end of statements

Example:

```sql
SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM Production.Product;
```

---

# Comments in T-SQL

Comments allow you to leave notes inside your SQL scripts.

SQL Server ignores commented text.

---

## Single-Line Comments

Single-line comments start with:

```sql
--
```

Example:

```sql
USE AdventureWorks2025;

-- Returns product ID, name, price, and colour from the product catalogue
SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM Production.Product;
```

Everything after `--` on that line is ignored by SQL Server.

---

## Multi-Line Comments

Multi-line comments start with:

```sql
/*
```

and end with:

```sql
*/
```

Example:

```sql
USE AdventureWorks2025;

/*
    Product catalogue report
    Author: [your name]
    Last updated: May 2025
*/

SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM Production.Product;
```

---

# Temporarily Commenting Out Code

Comments can also be used to temporarily disable part of a query.

For example:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       ListPrice
       -- Color
FROM Production.Product;
```

The commented line remains in your script but is ignored when the query runs.

This can be useful while testing and troubleshooting.

> When commenting out items from a comma-separated SELECT list, make sure the remaining commas still produce valid SQL.

---

# Column Aliases

Database column names are not always ideal for the people reading your output.

Aliases allow you to change the column heading in the result set without changing the underlying database.

Use the `AS` keyword.

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name AS ProductName,
       ListPrice AS Price,
       Color
FROM Production.Product;
```

The result now displays headings such as:

```text
ProductID
ProductName
Price
Color
```

The underlying database columns remain unchanged.

---

## Aliases Containing Spaces

If an alias contains spaces, wrap it in square brackets.

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name AS ProductName,
       ListPrice AS [Unit Price],
       Color AS ProductColour
FROM Production.Product;
```

The result heading will display:

```text
Unit Price
```

---

## Avoiding Spaces in Aliases

Another option is to use underscores:

```sql
SELECT ListPrice AS Unit_Price
FROM Production.Product;
```

This avoids the need for square brackets.

Throughout your own SQL, choose a naming convention and use it consistently.

---

# Running One Query from a Larger Script

When working in SSMS, it is common to have several exploratory queries in the same query window.

For example:

```sql
select * from Production.Product

select productid, name, listprice from Production.Product

select name, color from Production.Product
```

To execute only one query:

1. Highlight the query you want to run.
2. Press `F5` or click **Execute**.

Only the highlighted text will execute.

---

## Important Gotcha

If nothing is highlighted when you press `F5`, SSMS executes everything in the query window.

For SELECT queries this may simply return several result sets.

Later, when working with statements such as:

```sql
UPDATE
DELETE
INSERT
```

executing more code than intended can have much more serious consequences.

Build the habit of checking what is highlighted before pressing Execute.

---

# Copying SQL Server Results into Excel

SQL query results are often the starting point for additional analysis in Excel.

After running a query, select the results grid.

To copy the data:

```text
Ctrl + C
```

To copy the data **including column headers**:

```text
Ctrl + Shift + C
```

You can also right-click the result grid and choose:

```text
Copy with Headers
```

Then paste the results into Excel.

---

# Using AI with T-SQL

AI tools such as ChatGPT, Claude, and Copilot can be useful when working with SQL.

They can help:

* Format SQL
* Explain unfamiliar queries
* Diagnose errors
* Suggest improvements
* Generate starting queries

But AI-generated SQL should always be reviewed and tested.

A query can be syntactically valid while still returning the wrong business result.

---

# AI Tip 1 — Clean Up a Messy Query

Suppose you have:

```sql
select productid, name, listprice, color from production.product
```

You could use the following AI prompt:

```text
Clean up this T-SQL query.

Capitalise all SQL keywords, put each column on its own line in the SELECT list with consistent indentation, use PascalCase for table and column names, and add a brief comment above the SELECT explaining what the query returns.

select productid, name, listprice, color from production.product
```

A cleaned-up version could look like:

```sql
-- Returns product ID, name, list price, and colour for all products
SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM Production.Product;
```

---

# AI Tip 2 — Explain a Query You Did Not Write

You will often encounter SQL written by somebody else.

A useful AI prompt is:

```text
Explain what this T-SQL query does in plain English, step by step.

Describe what each clause is doing and what the result set will contain.
```

Example query:

```sql
SELECT soh.SalesOrderID,
       soh.OrderDate,
       soh.TotalDue,
       p.FirstName,
       p.LastName
FROM Sales.SalesOrderHeader soh
JOIN Sales.Customer c
    ON soh.CustomerID = c.CustomerID
JOIN Person.Person p
    ON c.PersonID = p.BusinessEntityID
WHERE soh.TotalDue > 1000
ORDER BY soh.OrderDate DESC;
```

Even if you have not learned every part of the query yet, AI can help you understand its overall purpose before you analyse each clause in detail.

---

# Complete Reference Script

Below is a consolidated script containing the main examples from this video.

```sql
USE AdventureWorks2025;


/* =========================================================
   BASIC SELECT
   ========================================================= */

SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM Production.Product;


/* =========================================================
   SELECT ALL COLUMNS
   ========================================================= */

SELECT *
FROM Production.Product;


/* =========================================================
   COLUMN ALIASES
   ========================================================= */

SELECT ProductID,
       Name AS ProductName,
       ListPrice AS Price,
       Color
FROM Production.Product;


/* =========================================================
   ALIASES WITH SPACES
   ========================================================= */

SELECT ProductID,
       Name AS ProductName,
       ListPrice AS [Unit Price],
       Color AS ProductColour
FROM Production.Product;


/* =========================================================
   SINGLE-LINE COMMENT
   ========================================================= */

-- Returns product ID, name, price, and colour from the product catalogue
SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM Production.Product;


/* =========================================================
   MULTI-LINE COMMENT
   ========================================================= */

/*
    Product catalogue report
    Author: [your name]
    Last updated: May 2025
*/

SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM Production.Product;


/* =========================================================
   EXPLORATORY QUERIES
   Highlight only the query you want to execute.
   ========================================================= */

SELECT *
FROM Production.Product;

SELECT ProductID,
       Name,
       ListPrice
FROM Production.Product;

SELECT Name,
       Color
FROM Production.Product;


/* =========================================================
   QUERY USED FOR AI EXPLANATION EXAMPLE
   ========================================================= */

SELECT soh.SalesOrderID,
       soh.OrderDate,
       soh.TotalDue,
       p.FirstName,
       p.LastName
FROM Sales.SalesOrderHeader soh
JOIN Sales.Customer c
    ON soh.CustomerID = c.CustomerID
JOIN Person.Person p
    ON c.PersonID = p.BusinessEntityID
WHERE soh.TotalDue > 1000
ORDER BY soh.OrderDate DESC;
```

---

# Key Takeaways

By the end of this video, you should be comfortable with:

* Writing a basic `SELECT ... FROM` query
* Selecting individual columns
* Using `SELECT *` to explore unfamiliar tables
* Referencing tables using `Schema.Table`
* Selecting the correct database in SSMS
* Understanding the difference between a query and a statement
* Reading basic SQL Server error messages
* Formatting T-SQL consistently
* Writing single-line and multi-line comments
* Using column aliases with `AS`
* Running only highlighted SQL in SSMS
* Copying query results into Excel with headers
* Using AI to format, explain, and troubleshoot SQL

---

## Next Video

**Video 05 — The WHERE Clause**

In the next video, we add `WHERE`, which is where querying starts to become genuinely useful.
