# Video 07 — NULL Values in T-SQL

**YouTube:** NULL Values in T-SQL — IS NULL, ISNULL, COALESCE & the NOT IN Trap | Video 07 of 36

**Module 2: Selecting and Shaping Data**

In this video, we look at one of the most important and commonly misunderstood concepts in SQL: `NULL`.

You will learn what `NULL` actually means, why normal comparison operators do not work with it, how to correctly test for missing values using `IS NULL` and `IS NOT NULL`, and how to replace NULL values in query output using `ISNULL()` and `COALESCE()`.

We also look at how NULL affects arithmetic and aggregation, and finish with an important `NOT IN` problem that can cause a query to unexpectedly return zero rows.

---

## 1. What NULL Actually Means

`NULL` represents the absence of a value.

It does **not** mean:

* Zero
* An empty string
* A blank space
* The word "NULL"

It means that the value is unknown, missing, was never recorded, or does not apply.

For example, the `Production.Product` table contains several nullable columns.

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       Color,
       Size,
       Weight
FROM Production.Product;
```

When you run this query, you will see blank-looking cells in columns such as `Color`, `Size`, and `Weight`.

Those cells contain `NULL`.

---

# Three-Valued Logic

Most programming comparisons return one of two outcomes:

```text
TRUE
FALSE
```

SQL introduces a third:

```text
UNKNOWN
```

A comparison involving `NULL` normally produces `UNKNOWN`.

For example:

```sql
NULL = NULL
```

does not evaluate to `TRUE`.

It evaluates to:

```text
UNKNOWN
```

When SQL Server evaluates a `WHERE` clause, rows where the condition is `FALSE` or `UNKNOWN` are excluded.

This is why NULL-related queries can produce unexpected results without generating an error.

---

# Why = NULL Does Not Work

Suppose you want to find products that have no colour recorded.

A beginner might try:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       Color
FROM Production.Product
WHERE Color = NULL;
```

This returns zero rows.

The problem is this expression:

```sql
Color = NULL
```

SQL Server cannot determine whether a value is equal to an unknown value.

The comparison therefore returns:

```text
UNKNOWN
```

The rows are excluded.

---

## The Correct Test: IS NULL

To test for the absence of a value, use:

```sql
IS NULL
```

Example:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       Color
FROM Production.Product
WHERE Color IS NULL;
```

This correctly returns products where no colour has been recorded.

`IS NULL` is specifically designed to test whether a value is absent.

---

## IS NOT NULL

The inverse test is:

```sql
IS NOT NULL
```

Example:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       Color
FROM Production.Product
WHERE Color IS NOT NULL;
```

This returns products where a colour value exists.

The important rule is:

```text
IS NULL
IS NOT NULL
```

Use these when testing whether a value is present or absent.

Do not use:

```sql
= NULL
```

---

# Replacing NULL Values with ISNULL()

Sometimes you do not want to filter NULL rows out.

Instead, you want to keep the row but replace the NULL with something more useful in the query output.

For example, rather than displaying a blank colour, we could display:

```text
Not Specified
```

Use the `ISNULL()` function.

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       ISNULL(Color, 'Not Specified') AS Color
FROM Production.Product;
```

`ISNULL()` takes two arguments:

```text
ISNULL(value_to_check, replacement_value)
```

In this example:

```sql
ISNULL(Color, 'Not Specified')
```

means:

> Return `Color` if it has a value. Otherwise return `Not Specified`.

The original data in the table is not changed.

Only the query result is changed.

---

# COALESCE()

`COALESCE()` provides another way to work with NULL values.

Unlike `ISNULL()`, which checks one value and provides one replacement, `COALESCE()` can evaluate several possible values.

It returns the first value that is not NULL.

Example:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       COALESCE(Color, Size, 'No Details') AS Description
FROM Production.Product;
```

SQL Server evaluates the arguments from left to right.

Conceptually:

```text
Color
  ↓
If NULL, try Size
  ↓
If NULL, return 'No Details'
```

So:

```sql
COALESCE(Color, Size, 'No Details')
```

returns:

1. `Color`, if available
2. Otherwise `Size`
3. Otherwise `No Details`

This is useful when several optional fields could provide a suitable value.

---

# NULL in Arithmetic

NULL also affects calculations.

If one part of an arithmetic expression is NULL, the result normally becomes NULL.

For example:

```text
100 * NULL
```

returns:

```text
NULL
```

This behaviour is sometimes described as **NULL propagation**.

---

## Checking Column Nullability

Before using the example calculation, we can inspect whether the columns allow NULL values.

```sql
USE AdventureWorks2025;

SELECT COLUMNPROPERTY(
           OBJECT_ID('Production.Product'),
           'StandardCost',
           'AllowsNull'
       ) AS StandardCost_Nullable,
       COLUMNPROPERTY(
           OBJECT_ID('Production.Product'),
           'Weight',
           'AllowsNull'
       ) AS Weight_Nullable;
```

For these columns:

```text
StandardCost_Nullable = 0
Weight_Nullable       = 1
```

`0` means NULL is not allowed.

`1` means NULL is allowed.

So `StandardCost` always contains a value, while `Weight` can contain NULL.

---

# NULL Propagation in a Calculation

Consider this calculated column:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       StandardCost,
       Weight,
       StandardCost * Weight AS TotalWeightCost
FROM Production.Product;
```

Where `Weight` contains a value, SQL Server performs the multiplication normally.

Where `Weight` is NULL:

```text
StandardCost * NULL
```

produces:

```text
NULL
```

Even though `StandardCost` itself contains a valid number.

---

# Protecting a Calculation with ISNULL()

If the business requirement is to treat a missing weight as zero, wrap `Weight` in `ISNULL()`.

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       StandardCost,
       Weight,
       StandardCost * ISNULL(Weight, 0) AS TotalWeightCost
FROM Production.Product;
```

Now:

```text
NULL Weight
```

is replaced with:

```text
0
```

before the multiplication occurs.

So:

```text
StandardCost * 0
```

returns:

```text
0
```

instead of NULL.

---

## Defensive NULL Handling

If you have not checked whether either column allows NULL, you may sometimes see calculations written like this:

```sql
ISNULL(StandardCost, 0) * ISNULL(Weight, 0) AS TotalWeightCost
```

For this particular table, wrapping `StandardCost` is unnecessary because the schema guarantees that it cannot be NULL.

But on columns where nullability has not been verified, this can be a useful defensive pattern.

The important point is that the replacement value should make sense for the business scenario.

Do not automatically replace every NULL with zero.

Sometimes NULL genuinely means:

```text
Unknown
```

and replacing it with zero would change the meaning of the data.

---

# COUNT(*) vs COUNT(column)

NULL also affects aggregate functions.

A useful distinction is the difference between:

```sql
COUNT(*)
```

and:

```sql
COUNT(ColumnName)
```

Consider:

```sql
USE AdventureWorks2025;

SELECT COUNT(*) AS TotalRows,
       COUNT(Color) AS RowsWithColor,
       COUNT(Size) AS RowsWithSize
FROM Production.Product;
```

`COUNT(*)` counts every row.

`COUNT(Color)` counts only rows where `Color` is not NULL.

`COUNT(Size)` counts only rows where `Size` is not NULL.

The rule is:

```text
COUNT(*)      → counts rows
COUNT(column) → counts non-NULL values
```

We will return to aggregate functions in more detail when we cover `GROUP BY`.

---

# The NOT IN NULL Trap

One of the more surprising NULL behaviours appears with `NOT IN`.

Suppose you want products whose colour is not Red or Blue.

You might write:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       Color
FROM Production.Product
WHERE Color NOT IN ('Red', 'Blue');
```

This works as expected.

But consider this:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       Color
FROM Production.Product
WHERE Color NOT IN ('Red', 'Blue', NULL);
```

This returns zero rows.

---

## Why Does NOT IN Return Zero Rows?

Conceptually:

```sql
Color NOT IN ('Red', 'Blue', NULL)
```

requires SQL Server to confirm that the colour is different from every value in the list.

That includes:

```sql
Color <> NULL
```

But a comparison with NULL produces:

```text
UNKNOWN
```

SQL Server can therefore no longer prove that the row satisfies the full `NOT IN` condition.

The result becomes UNKNOWN and the row is excluded.

---

## Removing NULL from the List

Remove the NULL:

```sql
USE AdventureWorks2025;

SELECT ProductID,
       Name,
       Color
FROM Production.Product
WHERE Color NOT IN ('Red', 'Blue');
```

Rows are returned again.

The only change is that NULL is no longer present in the list.

---

# Why This Matters in Real Queries

You probably will not intentionally write:

```sql
NOT IN ('Red', 'Blue', NULL)
```

The more realistic problem appears when `NOT IN` uses the result of another query.

For example, later you may encounter patterns such as:

```sql
WHERE SomeColumn NOT IN
(
    SELECT SomeOtherColumn
    FROM SomeTable
)
```

If that subquery unexpectedly returns a NULL, the outer query may produce results you were not expecting.

A useful troubleshooting rule is:

> If `NOT IN` unexpectedly returns zero rows, check whether NULL values are entering the comparison.

Later in the course we will look at `NOT EXISTS`, which is often a safer alternative for this type of subquery.

---

# Complete Reference Script

Below is a consolidated script containing the main examples from this video.

```sql
USE AdventureWorks2025;


/* =========================================================
   EXPLORE NULL VALUES
   ========================================================= */

SELECT ProductID,
       Name,
       Color,
       Size,
       Weight
FROM Production.Product;


/* =========================================================
   INCORRECT NULL COMPARISON
   = NULL DOES NOT WORK
   ========================================================= */

SELECT ProductID,
       Name,
       Color
FROM Production.Product
WHERE Color = NULL;


/* =========================================================
   CORRECT NULL TEST
   ========================================================= */

SELECT ProductID,
       Name,
       Color
FROM Production.Product
WHERE Color IS NULL;


/* =========================================================
   IS NOT NULL
   ========================================================= */

SELECT ProductID,
       Name,
       Color
FROM Production.Product
WHERE Color IS NOT NULL;


/* =========================================================
   REPLACE NULL USING ISNULL()
   ========================================================= */

SELECT ProductID,
       Name,
       ISNULL(Color, 'Not Specified') AS Color
FROM Production.Product;


/* =========================================================
   COALESCE()
   Return the first non-NULL value
   ========================================================= */

SELECT ProductID,
       Name,
       COALESCE(Color, Size, 'No Details') AS Description
FROM Production.Product;


/* =========================================================
   CHECK COLUMN NULLABILITY

   AllowsNull:
   0 = NULL not allowed
   1 = NULL allowed
   ========================================================= */

SELECT COLUMNPROPERTY(
           OBJECT_ID('Production.Product'),
           'StandardCost',
           'AllowsNull'
       ) AS StandardCost_Nullable,
       COLUMNPROPERTY(
           OBJECT_ID('Production.Product'),
           'Weight',
           'AllowsNull'
       ) AS Weight_Nullable;


/* =========================================================
   NULL PROPAGATION THROUGH ARITHMETIC
   ========================================================= */

SELECT ProductID,
       Name,
       StandardCost,
       Weight,
       StandardCost * Weight AS TotalWeightCost
FROM Production.Product;


/* =========================================================
   PROTECTING THE CALCULATION WITH ISNULL()

   StandardCost is NOT NULL by schema constraint.
   Weight allows NULL.
   ========================================================= */

SELECT ProductID,
       Name,
       StandardCost,
       Weight,
       StandardCost * ISNULL(Weight, 0) AS TotalWeightCost
FROM Production.Product;


/* =========================================================
   DEFENSIVE NULL HANDLING

   Useful when column nullability has not been verified.
   ========================================================= */

SELECT ProductID,
       Name,
       StandardCost,
       Weight,
       ISNULL(StandardCost, 0) * ISNULL(Weight, 0)
           AS TotalWeightCost
FROM Production.Product;


/* =========================================================
   COUNT(*) VS COUNT(column)
   ========================================================= */

SELECT COUNT(*) AS TotalRows,
       COUNT(Color) AS RowsWithColor,
       COUNT(Size) AS RowsWithSize
FROM Production.Product;


/* =========================================================
   NOT IN WITH NULL
   WARNING: RETURNS ZERO ROWS
   ========================================================= */

SELECT ProductID,
       Name,
       Color
FROM Production.Product
WHERE Color NOT IN ('Red', 'Blue', NULL);


/* =========================================================
   NOT IN WITHOUT NULL
   ========================================================= */

SELECT ProductID,
       Name,
       Color
FROM Production.Product
WHERE Color NOT IN ('Red', 'Blue');
```

---

# Key Takeaways

By the end of this video, you should be comfortable with:

* Understanding what `NULL` represents in SQL
* Recognising that NULL is not zero or an empty string
* Understanding SQL's three-valued logic
* Knowing why comparisons involving NULL produce `UNKNOWN`
* Using `IS NULL` to find missing values
* Using `IS NOT NULL` to find rows where values exist
* Understanding why `= NULL` does not work
* Replacing NULL values in query output with `ISNULL()`
* Using `COALESCE()` to return the first available non-NULL value
* Understanding NULL propagation through arithmetic
* Protecting calculations from NULL where appropriate
* Recognising that replacement values such as zero depend on business context
* Understanding the difference between `COUNT(*)` and `COUNT(column)`
* Recognising the `NOT IN` NULL trap
* Knowing to check for NULLs when `NOT IN` unexpectedly returns zero rows

---

## Next Video

**Video 08 — Sorting and Limiting Results with ORDER BY and TOP**

In the next video, we look at how to control the order of query results using `ORDER BY`, sort using `ASC` and `DESC`, limit results using `TOP`, and handle ties at the boundary with `TOP WITH TIES`.
