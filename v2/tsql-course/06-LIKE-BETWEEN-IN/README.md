# Video 06 — LIKE, BETWEEN and IN

**T-SQL for Business Professionals**
**Module 2: Selecting and Shaping Data**

In the previous lesson, we introduced the `WHERE` clause and used comparison operators to filter rows.

In this lesson, we extend that filtering toolkit with three operators that come up constantly in practical SQL work:

* `LIKE` for text pattern matching
* `BETWEEN` for ranges
* `IN` for matching against a list of values

We also look at several important behaviours that can produce unexpected results, including datetime boundaries, `NULL` values inside `NOT IN`, and case sensitivity when using `LIKE`.

---

## What You Will Learn

By the end of this lesson, you will be able to:

* Search text using `LIKE`
* Use `%` to match zero or more characters
* Use `_` to match exactly one character
* Use `NOT LIKE`
* Understand the performance implication of leading wildcards
* Filter numeric and date ranges using `BETWEEN`
* Understand that `BETWEEN` includes both boundaries
* Safely filter datetime values
* Replace long `OR` conditions with `IN`
* Use `NOT IN`
* Understand why `NULL` can cause problems with `NOT IN`
* Understand how database collation affects `LIKE`

---

# 1. LIKE — Filtering by Pattern

In the previous video, we filtered text using exact comparisons such as:

```sql
WHERE Color = 'Black'
```

But sometimes we do not know the complete value.

We might want:

* names beginning with a particular word
* names ending with particular text
* names containing a word anywhere

For this, T-SQL provides `LIKE`.

`LIKE` compares a text column against a **pattern**.

---

# 2. The Percent Wildcard

The `%` wildcard represents **zero or more characters**.

Suppose we want every product whose name starts with `Mountain`.

```sql
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE Name LIKE 'Mountain%';
```

The pattern:

```sql
'Mountain%'
```

means:

> Start with Mountain, followed by zero or more characters.

This can match values such as:

```text
Mountain-100
Mountain-200
Mountain Bike Frame
```

---

# 3. Finding Values That End With Text

Move the wildcard to the beginning:

```sql
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE Name LIKE '%Bike';
```

The pattern:

```sql
'%Bike'
```

means:

> Anything followed by Bike.

So the value must end with `Bike`.

---

# 4. Finding Text Anywhere

Place `%` at both ends:

```sql
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE Name LIKE '%Mountain%';
```

Now `Mountain` can appear anywhere in the value.

The pattern means:

```text
anything + Mountain + anything
```

This is useful for searches where you know part of the text but not its position.

---

# 5. The Underscore Wildcard

The second wildcard used with `LIKE` is `_`.

An underscore represents **exactly one character**.

Compare the two:

| Wildcard | Meaning                 |
| -------- | ----------------------- |
| `%`      | Zero or more characters |
| `_`      | Exactly one character   |

AdventureWorks contains product names such as:

```text
Mountain-100
Mountain-200
Mountain-300
```

We can use:

```sql
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE Name LIKE 'Mountain-_00%';
```

The `_` allows exactly one character between the hyphen and the two zeros.

The final `%` allows additional characters after that pattern.

In practice, `%` is used much more frequently, while `_` is useful when matching a specific text structure.

---

# 6. NOT LIKE

`NOT LIKE` reverses the condition.

For example:

```sql
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE Name NOT LIKE 'Mountain%';
```

This returns products whose names do **not** begin with `Mountain`.

---

# 7. LIKE and Performance

There is an important performance difference between these patterns:

```sql
LIKE 'Mountain%'
```

and:

```sql
LIKE '%Mountain%'
```

With:

```sql
LIKE 'Mountain%'
```

the known text is at the beginning of the value.

SQL Server may still be able to use an appropriate database index efficiently.

With:

```sql
LIKE '%Mountain%'
```

the pattern begins with a wildcard.

SQL Server cannot know where `Mountain` will appear, so it may need to scan the rows to evaluate the condition.

On a small table such as `Production.Product`, the difference may not be noticeable.

On tables containing millions of rows, leading wildcards in frequently executed queries can become a performance consideration.

---

# 8. BETWEEN — Filtering by Range

`BETWEEN` tests whether a value falls between a lower and upper boundary.

Suppose we want products priced between 500 and 1000.

```sql
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE ListPrice BETWEEN 500 AND 1000;
```

An important property of `BETWEEN` is that **both boundaries are included**.

Conceptually:

```sql
WHERE ListPrice BETWEEN 500 AND 1000
```

is equivalent to:

```sql
WHERE ListPrice >= 500
  AND ListPrice <= 1000
```

`BETWEEN` simply provides a cleaner way of expressing the range.

---

# 9. BETWEEN with Dates

`BETWEEN` can also be used with dates.

Suppose we want orders within a date range:

```sql
USE AdventureWorks2025;

SELECT
    SalesOrderID,
    OrderDate,
    TotalDue
FROM Sales.SalesOrderHeader
WHERE OrderDate BETWEEN '2025-01-01' AND '2025-03-31';
```

The lower and upper values are both included by `BETWEEN`.

However, there is an important issue when the column contains a **time component**. We cover that shortly.

---

# 10. Put the Lower Value First

The lower boundary comes first in a `BETWEEN` expression.

This works logically:

```sql
WHERE OrderDate BETWEEN '2025-01-01' AND '2025-03-31'
```

Reversing the dates:

```sql
WHERE OrderDate BETWEEN '2025-03-31' AND '2025-01-01'
```

returns no qualifying rows.

SQL Server is effectively looking for a value that is simultaneously:

```text
on or after 31 March
```

and:

```text
on or before 1 January
```

Nothing can satisfy both conditions.

### Rule

Put the **smaller/lower value first** and the **larger/upper value second**.

---

# 11. NOT BETWEEN

`NOT BETWEEN` returns values outside the specified range.

```sql
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE ListPrice NOT BETWEEN 500 AND 1000;
```

This returns products priced:

* below 500
* or above 1000

---

# 12. GOTCHA — BETWEEN and Datetime Values

There is an important issue when using `BETWEEN` against columns containing both a date and a time.

`OrderDate` in `Sales.SalesOrderHeader` uses a datetime data type.

Consider:

```sql
WHERE OrderDate BETWEEN '2025-01-01' AND '2025-03-31'
```

The upper value:

```text
2025-03-31
```

is interpreted as the beginning of that day:

```text
2025-03-31 00:00:00
```

That means records later on the final day may fall outside the range.

---

# 13. A Safer Datetime Pattern

For datetime columns, a safer approach is to make the upper boundary **exclusive** and use the following day.

The general pattern is:

```sql
WHERE DateColumn >= 'StartDate'
  AND DateColumn < 'DayAfterEndDate'
```

For example:

```sql
USE AdventureWorks2025;

SELECT
    SalesOrderID,
    OrderDate,
    TotalDue
FROM Sales.SalesOrderHeader
WHERE OrderDate >= '2024-01-01'
  AND OrderDate < '2025-01-01';
```

This captures every datetime from the beginning of 2024 up to, but not including, the beginning of 2025.

Because the next period's starting point is excluded, the time component on the final day does not cause records to be missed.

We will return to dates and date functions later in the course.

---

# 14. IN — Filtering Against a List

`IN` tests whether a value matches **any value in a specified list**.

Suppose we want products that are:

* Red
* Blue
* Yellow

We could write:

```sql
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color = 'Red'
   OR Color = 'Blue'
   OR Color = 'Yellow';
```

This works.

But `IN` provides a cleaner alternative.

---

# 15. Replacing OR with IN

The same filter can be written as:

```sql
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color IN ('Red', 'Blue', 'Yellow');
```

Both versions produce the same result.

The `IN` version is:

* shorter
* easier to read
* easier to maintain

If another colour needs to be added later, simply add it to the list.

```sql
WHERE Color IN ('Red', 'Blue', 'Yellow', 'Silver');
```

---

# 16. IN with Numbers

`IN` also works with numeric values.

Suppose we want orders from three specific sales territories:

```sql
USE AdventureWorks2025;

SELECT
    SalesOrderID,
    OrderDate,
    TotalDue,
    TerritoryID
FROM Sales.SalesOrderHeader
WHERE TerritoryID IN (1, 4, 6);
```

Numbers do not require quotes.

---

# 17. Combining IN with Other Conditions

`IN` is simply another condition inside `WHERE`.

That means it can be combined with `AND`, `OR`, and other predicates.

For example:

```sql
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE Color IN ('Red', 'Blue')
  AND ListPrice > 500;
```

This returns products that are:

```text
Red OR Blue
```

and also:

```text
priced above 500
```

---

# 18. NOT IN

`NOT IN` reverses an `IN` condition.

For example:

```sql
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color NOT IN ('Red', 'Blue', 'Yellow');
```

This returns values that do not match any item in the list.

However, `NOT IN` has an important relationship with `NULL`.

---

# 19. NULL Values and NOT IN

First, consider the normal version:

```sql
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color NOT IN ('Red', 'Blue', 'Yellow');
```

Rows where `Color` itself is `NULL` are not included.

A comparison involving `NULL` evaluates to `UNKNOWN`, not `TRUE`, so those rows do not pass the filter.

This becomes even more important when `NULL` appears **inside the NOT IN list itself**.

---

# 20. GOTCHA — NULL Inside a NOT IN List

Consider this query:

```sql
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color NOT IN ('Red', 'Blue', 'Yellow', NULL);
```

This can cause the result to collapse to **zero rows**.

Why?

For `NOT IN` to return a row, SQL Server must be able to establish that the value does not match the values in the list.

But once `NULL` appears in that list, SQL Server cannot establish that the value is definitively not equal to the unknown value represented by `NULL`.

The result becomes `UNKNOWN`, so the row does not pass the filter.

This becomes especially important when the list comes from a **subquery**, because a `NULL` can enter the results without being obvious from the query itself.

We explore `NULL` behaviour properly in the next video.

---

# 21. GOTCHA — Is LIKE Case-Sensitive?

Consider these two conditions:

```sql
WHERE Name LIKE 'mountain%'
```

and:

```sql
WHERE Name LIKE 'Mountain%'
```

Do they return the same result?

The answer depends on the **collation** of the database.

Collation controls how SQL Server compares and sorts text.

Many SQL Server databases use a case-insensitive collation, but you should not assume that every database does.

---

# 22. Check the Database Collation

We can inspect the collation of AdventureWorks using:

```sql
USE AdventureWorks2025;

SELECT
    name,
    collation_name
FROM sys.databases
WHERE name = 'AdventureWorks2025';
```

You may see a collation similar to:

```text
SQL_Latin1_General_CP1_CI_AS
```

For our purposes, two parts are particularly useful:

```text
CI = Case Insensitive
AS = Accent Sensitive
```

With a case-insensitive collation:

```text
mountain
```

and:

```text
Mountain
```

are treated as equivalent for this comparison.

If the database uses `CS` instead:

```text
CS = Case Sensitive
```

the casing matters.

---

# 23. Compare the LIKE Results

Try both versions:

```sql
USE AdventureWorks2025;

-- Lowercase
SELECT
    ProductID,
    Name
FROM Production.Product
WHERE Name LIKE 'mountain%';


-- Matching case
SELECT
    ProductID,
    Name
FROM Production.Product
WHERE Name LIKE 'Mountain%';
```

On a case-insensitive database, both can return the same rows.

On a case-sensitive database, they can produce different results.

### Practical Rule

Do not automatically assume that text comparisons are case-insensitive.

When working with an unfamiliar database:

1. Look at the actual data.
2. Match the casing used in the database.
3. Check the database collation if case sensitivity matters.

String functions such as `UPPER()` and `LOWER()` can also be used to normalise text. We cover those later in the course.

---

# LIKE Wildcard Reference

| Pattern                | Meaning                                 |
| ---------------------- | --------------------------------------- |
| `'Mountain%'`          | Starts with Mountain                    |
| `'%Bike'`              | Ends with Bike                          |
| `'%Mountain%'`         | Contains Mountain                       |
| `'Mountain-_00%'`      | Exactly one character where `_` appears |
| `NOT LIKE 'Mountain%'` | Does not start with Mountain            |

---

# BETWEEN Reference

### Numeric range

```sql
WHERE ListPrice BETWEEN 500 AND 1000
```

Equivalent logic:

```sql
WHERE ListPrice >= 500
  AND ListPrice <= 1000
```

### Outside a range

```sql
WHERE ListPrice NOT BETWEEN 500 AND 1000
```

### Safer datetime pattern

```sql
WHERE OrderDate >= '2024-01-01'
  AND OrderDate < '2025-01-01'
```

---

# IN Reference

### Text values

```sql
WHERE Color IN ('Red', 'Blue', 'Yellow')
```

### Numeric values

```sql
WHERE TerritoryID IN (1, 4, 6)
```

### Combine with another condition

```sql
WHERE Color IN ('Red', 'Blue')
  AND ListPrice > 500
```

### NOT IN

```sql
WHERE Color NOT IN ('Red', 'Blue', 'Yellow')
```

### Dangerous NULL example

```sql
WHERE Color NOT IN ('Red', 'Blue', 'Yellow', NULL)
```

---

# Complete Practice Script

The following script brings together the main examples from this lesson.

```sql
USE AdventureWorks2025;

--------------------------------------------------
-- LIKE: Starts with Mountain
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE Name LIKE 'Mountain%';


--------------------------------------------------
-- LIKE: Ends with Bike
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE Name LIKE '%Bike';


--------------------------------------------------
-- LIKE: Contains Mountain
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE Name LIKE '%Mountain%';


--------------------------------------------------
-- LIKE: Underscore wildcard
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE Name LIKE 'Mountain-_00%';


--------------------------------------------------
-- NOT LIKE
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE Name NOT LIKE 'Mountain%';


--------------------------------------------------
-- BETWEEN: Numeric range
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE ListPrice BETWEEN 500 AND 1000;


--------------------------------------------------
-- BETWEEN: Date range
--------------------------------------------------

SELECT
    SalesOrderID,
    OrderDate,
    TotalDue
FROM Sales.SalesOrderHeader
WHERE OrderDate BETWEEN '2025-01-01' AND '2025-03-31';


--------------------------------------------------
-- BETWEEN: Reversed boundaries
-- Returns zero rows
--------------------------------------------------

SELECT
    SalesOrderID,
    OrderDate,
    TotalDue
FROM Sales.SalesOrderHeader
WHERE OrderDate BETWEEN '2025-03-31' AND '2025-01-01';


--------------------------------------------------
-- NOT BETWEEN
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE ListPrice NOT BETWEEN 500 AND 1000;


--------------------------------------------------
-- Safer datetime filtering pattern
--------------------------------------------------

SELECT
    SalesOrderID,
    OrderDate,
    TotalDue
FROM Sales.SalesOrderHeader
WHERE OrderDate >= '2024-01-01'
  AND OrderDate < '2025-01-01';


--------------------------------------------------
-- Multiple OR conditions
--------------------------------------------------

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color = 'Red'
   OR Color = 'Blue'
   OR Color = 'Yellow';


--------------------------------------------------
-- Replace the OR conditions with IN
--------------------------------------------------

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color IN ('Red', 'Blue', 'Yellow');


--------------------------------------------------
-- Numeric IN
--------------------------------------------------

SELECT
    SalesOrderID,
    OrderDate,
    TotalDue,
    TerritoryID
FROM Sales.SalesOrderHeader
WHERE TerritoryID IN (1, 4, 6);


--------------------------------------------------
-- IN combined with AND
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE Color IN ('Red', 'Blue')
  AND ListPrice > 500;


--------------------------------------------------
-- Standard NOT IN
--------------------------------------------------

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color NOT IN ('Red', 'Blue', 'Yellow');


--------------------------------------------------
-- NOT IN with NULL
-- Demonstrates the NULL problem
--------------------------------------------------

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color NOT IN ('Red', 'Blue', 'Yellow', NULL);


--------------------------------------------------
-- Check database collation
--------------------------------------------------

SELECT
    name,
    collation_name
FROM sys.databases
WHERE name = 'AdventureWorks2025';


--------------------------------------------------
-- LIKE and case sensitivity
--------------------------------------------------

SELECT
    ProductID,
    Name
FROM Production.Product
WHERE Name LIKE 'mountain%';


SELECT
    ProductID,
    Name
FROM Production.Product
WHERE Name LIKE 'Mountain%';
```

---

# Key Takeaways

1. `LIKE` searches for text patterns rather than exact values.
2. `%` represents zero or more characters.
3. `_` represents exactly one character.
4. A leading `%` can have performance implications on large tables.
5. `BETWEEN` includes both its lower and upper boundaries.
6. Put the lower boundary before the upper boundary.
7. Be careful using inclusive end dates with datetime columns.
8. An exclusive upper boundary using the following date is often safer for datetime filtering.
9. `IN` is a cleaner alternative to long chains of `OR` conditions.
10. `IN` works with both text and numeric values.
11. Be especially careful with `NULL` and `NOT IN`.
12. `LIKE` case sensitivity depends on the database collation.
13. Do not assume every SQL Server database is case-insensitive.

---

# Next Video

## Video 07 — Understanding NULL Values

In the next lesson, we take a deeper look at `NULL`.

`NULL` has already appeared several times in our filtering examples, and it behaves differently from normal values.

We will look at:

* what `NULL` actually represents
* why normal comparisons with `NULL` do not work
* three-valued logic
* `IS NULL` and `IS NOT NULL`
* replacing `NULL` in query output
* how `NULL` affects arithmetic and aggregation
* why `NOT IN` can behave unexpectedly when `NULL` is involved

Understanding `NULL` properly can prevent an entire category of silent errors in business reports.

---

# Course

This repository accompanies the **T-SQL for Business Professionals** YouTube course.

The course is designed for analysts, finance professionals, operations teams, Excel users, and other business professionals who want to learn how to work directly with data using SQL Server.

**Database used:** AdventureWorks2025
**Database platform:** Microsoft SQL Server
**Query tool:** SQL Server Management Studio (SSMS)
