# Video 08 — Sorting and Limiting Results with ORDER BY and TOP

**YouTube:** Sorting and Limiting Results in T-SQL — ORDER BY, TOP & WITH TIES | Video 08 of 36

**Module 2: Selecting and Shaping Data**

In this video, we look at how to control the order of query results using `ORDER BY` and how to limit the number of rows returned using `TOP`.

You will learn how to sort data in ascending and descending order, sort by multiple columns, and understand why SQL Server does not guarantee any particular row order unless you explicitly request one.

We also look at `TOP`, `TOP 1`, and `TOP WITH TIES`, before finishing with an important SQL processing-order concept: why a column alias can be used in `ORDER BY` but not in `WHERE`.

---

## 1. Sorting Results with ORDER BY

Without an `ORDER BY` clause, SQL Server does not guarantee the order in which rows will be returned.

If the order of your results matters, you need to specify it explicitly.

`ORDER BY` goes at the end of the query.

Let's start by sorting products from lowest to highest list price.

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       ListPrice
FROM Production.Product
ORDER BY ListPrice ASC;
```

`ASC` means **ascending**.

Depending on the type of data, this means:

* Lowest to highest
* A to Z
* Earliest to latest

`ASC` is the default direction, but explicitly including it can make the intended sort easier to understand.

---

## 2. Sorting in Descending Order

Use `DESC` to reverse the sort.

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       ListPrice
FROM Production.Product
ORDER BY ListPrice DESC;
```

Now the products with the highest list prices appear first.

`DESC` means **descending**:

* Highest to lowest
* Z to A
* Latest to earliest

The rows returned by the query have not changed. Only their order has changed.

---

## 3. Sorting Text

`ORDER BY` also works with text columns.

For example, we can sort the product catalogue alphabetically by product name.

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       ListPrice
FROM Production.Product
ORDER BY Name ASC;
```

The products are now returned from A to Z based on `Name`.

The same `ORDER BY` syntax works across different data types.

---

# Sorting by Multiple Columns

You can sort by more than one column.

SQL Server applies the sort conditions in the order they appear in the `ORDER BY` clause.

For example:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       Color,
       ListPrice
FROM Production.Product
ORDER BY Color ASC,
         ListPrice DESC;
```

SQL Server first sorts by:

```text
Color ASC
```

Then, within each colour group, it sorts by:

```text
ListPrice DESC
```

The result might therefore be thought of as:

```text
Black
    Highest price
    ↓
    Lowest price

Blue
    Highest price
    ↓
    Lowest price

Red
    Highest price
    ↓
    Lowest price
```

Each column in the `ORDER BY` list can have its own sort direction.

For example:

```sql
ORDER BY Color ASC,
         ListPrice DESC;
```

mixes ascending and descending sorting in the same query.

---

# NULL Values in ORDER BY

NULL values also participate in sorting.

In SQL Server, NULL values appear before non-NULL values when sorting in ascending order.

For example:

```sql
ORDER BY Color ASC;
```

will place products where `Color` is NULL before the populated colours.

With descending order:

```sql
ORDER BY Color DESC;
```

NULL values appear after the non-NULL values.

This is another behaviour to be aware of when nullable columns are used in reports.

---

# SQL Has No Default Row Order

One of the most important concepts to understand about SQL is:

> A table has no guaranteed default row order.

Consider:

```sql
USE AdventureWorks2025;

SELECT TOP 20
       ProductID,
       Name,
       ListPrice
FROM Production.Product;
```

You may run this query several times and see exactly the same twenty rows.

That does **not** mean those rows have a guaranteed position in the table.

SQL Server is free to return the rows in whatever order is appropriate for that particular execution.

The apparent order may be influenced by factors such as:

* Physical data storage
* Indexes
* Query execution plans
* Changes to the table
* Changes to the data

A small, stable table may appear to return rows consistently for a long time.

You should still not rely on that order.

The rule is simple:

> If row order matters, use `ORDER BY`.

---

# Why TOP Without ORDER BY Is Dangerous

The previous query contains:

```sql
SELECT TOP 20
```

but no `ORDER BY`.

That means SQL Server returns twenty rows, but there is no guarantee **which** twenty rows will be returned.

For example:

```sql
SELECT TOP 10
       ProductID,
       Name,
       ListPrice
FROM Production.Product;
```

does not mean:

```text
10 most expensive products
```

or:

```text
10 cheapest products
```

It simply requests ten rows.

To give `TOP` a meaningful business definition, it normally needs to be combined with `ORDER BY`.

---

# TOP — Limiting the Number of Rows

`TOP` limits the number of rows returned by a query.

It appears immediately after `SELECT`.

Suppose a procurement manager wants the ten products with the highest list price.

```sql
USE AdventureWorks2025;

SELECT TOP 10
       ProductID,
       Name,
       ListPrice
FROM Production.Product
ORDER BY ListPrice DESC;
```

The two clauses work together:

```text
ORDER BY ListPrice DESC
        ↓
Highest prices first
        ↓
TOP 10
        ↓
Return the first 10
```

The result is the ten highest-priced products.

---

# TOP 1

`TOP 1` is useful when you want a single row representing an extreme value.

For example, to find the most expensive product:

```sql
USE AdventureWorks2025;

SELECT TOP 1
       ProductID,
       Name,
       ListPrice
FROM Production.Product
ORDER BY ListPrice DESC;
```

Because the rows are sorted from highest to lowest price, the first row is the product with the highest `ListPrice`.

The same pattern can be used for questions such as:

```text
Most expensive product
Earliest order
Latest transaction
Longest-serving employee
Highest-value customer
```

The important part is choosing the appropriate `ORDER BY`.

---

# TOP WITH TIES

There is an additional consideration when using `TOP`.

Suppose you request:

```sql
TOP 10
```

and the tenth and eleventh rows have exactly the same value in the column being used for sorting.

Plain `TOP 10` still returns ten rows.

One tied row may therefore be included while another is excluded.

`WITH TIES` allows you to include every row that shares the boundary sort value.

Compare these two queries.

### Plain TOP 10

```sql
USE AdventureWorks2025;

SELECT TOP 10
       ProductID,
       Name,
       ListPrice
FROM Production.Product
ORDER BY ListPrice DESC;
```

### TOP 10 WITH TIES

```sql
USE AdventureWorks2025;

SELECT TOP 10 WITH TIES
       ProductID,
       Name,
       ListPrice
FROM Production.Product
ORDER BY ListPrice DESC;
```

If several products share the tenth-highest price, `WITH TIES` includes all of them.

The result can therefore contain more than ten rows.

---

## When Should You Use WITH TIES?

It depends on the business requirement.

For something such as:

```text
Top-performing products
Highest salespeople
Highest-value customers
```

including ties may provide a fairer result.

But suppose another system requires exactly ten records.

In that case:

```sql
TOP 10
```

may be the appropriate choice.

`WITH TIES` is useful when equality at the boundary matters more than maintaining a strict row count.

---

# Column Aliases in ORDER BY

Earlier in the course, we introduced column aliases using `AS`.

For example:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       ListPrice AS Price
FROM Production.Product
ORDER BY Price DESC;
```

This works.

`Price` is not the real database column name.

It is an alias created here:

```sql
ListPrice AS Price
```

Yet `ORDER BY` is able to reference it.

---

# Why Can ORDER BY See an Alias?

The reason relates to SQL's **logical processing order**.

Conceptually, `SELECT` is processed before `ORDER BY`.

By the time SQL Server logically reaches:

```sql
ORDER BY Price DESC;
```

the alias:

```sql
ListPrice AS Price
```

already exists.

That is why this works:

```sql
SELECT ProductID,
       Name,
       ListPrice AS Price
FROM Production.Product
ORDER BY Price DESC;
```

---

# Why the Same Alias Does Not Work in WHERE

Now consider:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       ListPrice AS Price
FROM Production.Product
WHERE Price > 500;
```

This fails.

SQL Server returns an error similar to:

```text
Invalid column name 'Price'.
```

The reason is that `WHERE` is logically processed before `SELECT`.

At the point where SQL Server evaluates:

```sql
WHERE Price > 500
```

the alias `Price` has not yet been created.

The correct version uses the original column name:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       ListPrice AS Price
FROM Production.Product
WHERE ListPrice > 500
ORDER BY Price DESC;
```

Here:

```sql
WHERE ListPrice > 500
```

uses the real column name.

And:

```sql
ORDER BY Price DESC
```

can use the alias.

---

# ORDER BY Is the Exception

A useful rule to remember is that aliases created in the `SELECT` list cannot normally be referenced by clauses that are logically processed before `SELECT`.

For example:

```text
WHERE
GROUP BY
HAVING
```

cannot generally use a `SELECT` alias in this way.

`ORDER BY` can because it is logically processed after `SELECT`.

We will revisit SQL's logical processing order as more clauses are introduced throughout the course.

---

# Complete Reference Script

Below is a consolidated script containing the main examples from this video.

```sql
USE AdventureWorks2025;


/* =========================================================
   ORDER BY — ASCENDING
   Lowest price first
   ========================================================= */

SELECT ProductID,
       Name,
       ListPrice
FROM Production.Product
ORDER BY ListPrice ASC;


/* =========================================================
   ORDER BY — DESCENDING
   Highest price first
   ========================================================= */

SELECT ProductID,
       Name,
       ListPrice
FROM Production.Product
ORDER BY ListPrice DESC;


/* =========================================================
   SORT TEXT ALPHABETICALLY
   ========================================================= */

SELECT ProductID,
       Name,
       ListPrice
FROM Production.Product
ORDER BY Name ASC;


/* =========================================================
   SORT BY MULTIPLE COLUMNS

   Color:
       Ascending

   ListPrice:
       Descending within each Color
   ========================================================= */

SELECT ProductID,
       Name,
       Color,
       ListPrice
FROM Production.Product
ORDER BY Color ASC,
         ListPrice DESC;


/* =========================================================
   TOP WITHOUT ORDER BY

   WARNING:
   Returns 20 rows, but there is no guaranteed definition
   of which 20 rows will be returned.
   ========================================================= */

SELECT TOP 20
       ProductID,
       Name,
       ListPrice
FROM Production.Product;


/* =========================================================
   TOP WITH ORDER BY
   Return the 10 highest-priced products
   ========================================================= */

SELECT TOP 10
       ProductID,
       Name,
       ListPrice
FROM Production.Product
ORDER BY ListPrice DESC;


/* =========================================================
   TOP 1
   Return the single highest-priced product
   ========================================================= */

SELECT TOP 1
       ProductID,
       Name,
       ListPrice
FROM Production.Product
ORDER BY ListPrice DESC;


/* =========================================================
   PLAIN TOP 10
   May exclude rows tied at the boundary
   ========================================================= */

SELECT TOP 10
       ProductID,
       Name,
       ListPrice
FROM Production.Product
ORDER BY ListPrice DESC;


/* =========================================================
   TOP 10 WITH TIES
   Include all rows tied at the boundary
   ========================================================= */

SELECT TOP 10 WITH TIES
       ProductID,
       Name,
       ListPrice
FROM Production.Product
ORDER BY ListPrice DESC;


/* =========================================================
   COLUMN ALIAS IN ORDER BY

   This works because ORDER BY is logically processed
   after SELECT.
   ========================================================= */

SELECT ProductID,
       Name,
       ListPrice AS Price
FROM Production.Product
ORDER BY Price DESC;


/* =========================================================
   COLUMN ALIAS IN WHERE

   WARNING:
   This query fails because WHERE is logically processed
   before SELECT creates the Price alias.
   ========================================================= */

SELECT ProductID,
       Name,
       ListPrice AS Price
FROM Production.Product
WHERE Price > 500;


/* =========================================================
   CORRECT VERSION

   Use the original column name in WHERE.
   The alias can still be used in ORDER BY.
   ========================================================= */

SELECT ProductID,
       Name,
       ListPrice AS Price
FROM Production.Product
WHERE ListPrice > 500
ORDER BY Price DESC;
```

---

# Key Takeaways

By the end of this video, you should be comfortable with:

* Sorting query results using `ORDER BY`
* Using `ASC` for ascending order
* Using `DESC` for descending order
* Sorting numeric and text columns
* Sorting by multiple columns
* Mixing ascending and descending directions
* Understanding how NULL values behave when sorted
* Recognising that SQL tables have no guaranteed default row order
* Knowing why you should use `ORDER BY` whenever row order matters
* Limiting query results using `TOP`
* Understanding why `TOP` without `ORDER BY` can return arbitrary rows
* Using `TOP 1` to retrieve a single highest or lowest result
* Using `TOP WITH TIES` when equal boundary values should all be included
* Choosing between a strict row limit and including ties based on the business requirement
* Using column aliases in `ORDER BY`
* Understanding why a `SELECT` alias cannot be used in `WHERE`
* Recognising how SQL's logical processing order explains alias visibility

---

## Next Video

**Video 09 — Calculated Columns: Build New Data from Existing Data**

In the next video, we start creating new values directly inside a query using calculated columns.

We will use arithmetic expressions to calculate values such as prices and margins, combine text using `CONCAT()`, take a first look at date calculations with `GETDATE()` and `DATEDIFF()`, and cover the integer division gotcha that can quietly produce incorrect results.
