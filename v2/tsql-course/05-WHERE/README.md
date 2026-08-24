Video 05 — Filtering Rows with WHERE

T-SQL for Business Professionals
Module 2: Selecting and Shaping Data

In this lesson, we introduce the WHERE clause and use it to filter query results down to exactly the rows we need.

We start with simple conditions on text, numbers, and dates, then build up to multiple conditions using AND, OR, and NOT.

We also look at two areas that regularly cause problems in real queries:

mixing AND and OR without parentheses
trying to compare values directly with NULL
What You Will Learn

By the end of this lesson, you will be able to:

Filter rows using WHERE
Write conditions against text, numeric, and date columns
Use the main T-SQL comparison operators
Combine conditions using AND, OR, and NOT
Understand why column aliases cannot normally be referenced in WHERE
Control operator precedence using parentheses
Correctly test for missing values with IS NULL and IS NOT NULL
1. Adding WHERE to a Query

Without a WHERE clause, a query normally returns every row that satisfies the rest of the query.

WHERE adds a condition, or predicate, that SQL Server evaluates for each row.

Rows where the condition evaluates to true are returned.

Example — Filter by Text
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE Color = 'Black';

The predicate is:

WHERE Color = 'Black'

SQL Server returns only rows where the Color column contains Black.

String Literals

Text values written directly into a T-SQL query are called string literals.

String literals are wrapped in single quotes:

'Black'

Use single quotes for text values.

2. Filtering Numeric Values

Numbers do not need quotes.

USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE ListPrice > 500;

Here SQL Server compares the value in ListPrice with the number 500.

3. Filtering Dates

The same approach works with dates.

USE AdventureWorks2025;

SELECT
    SalesOrderID,
    OrderDate,
    TotalDue
FROM Sales.SalesOrderHeader
WHERE OrderDate >= '2025-01-01';

For date literals, this course uses:

YYYY-MM-DD

For example:

2025-01-01

This format avoids ambiguity caused by regional date formats.

4. Values Depend on the Data Type

The way a value is written in a predicate depends on the column's data type.

Text
WHERE Color = 'Black'
Number
WHERE ListPrice > 500
Date
WHERE OrderDate >= '2025-01-01'
5. Column Aliases Cannot Be Used in WHERE

Suppose we give ListPrice a friendlier name:

USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice AS Price,
    Color
FROM Production.Product
WHERE Price > 500;

This fails because Price is an alias created in the SELECT list.

SQL Server evaluates FROM and WHERE before it evaluates the SELECT list.

At the point where WHERE is evaluated, the alias does not yet exist.

Use the original column name instead:

USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice AS Price,
    Color
FROM Production.Product
WHERE ListPrice > 500;
Rule

Use the real table column name inside WHERE, not an alias created in SELECT.

6. Comparison Operators

The main comparison operators used in T-SQL are:

Operator	Meaning
=	Equal to
<>	Not equal to
>	Greater than
<	Less than
>=	Greater than or equal to
<=	Less than or equal to
Equal To
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color = 'Red';
Not Equal To
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color <> 'Red';

One important detail: rows where Color is NULL are not returned by this condition.

We come back to this later in the lesson.

Less Than or Equal To
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE ListPrice <= 100;
Comparing Dates
USE AdventureWorks2025;

SELECT
    SalesOrderID,
    OrderDate,
    TotalDue
FROM Sales.SalesOrderHeader
WHERE OrderDate < '2025-06-01';
7. The Basic Structure of a Predicate

A typical predicate follows this pattern:

Column
Operator
Value

For example:

ListPrice > 500

or:

Color = 'Black'
8. Combining Conditions with AND

AND requires all connected conditions to be true.

For example, find products that are:

Black
Priced above 500
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE Color = 'Black'
  AND ListPrice > 500;

A product must satisfy both conditions to appear.

Adding a Third Condition

We can continue adding conditions.

In Production.Product, a NULL SellEndDate indicates that the product is still active.

USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice,
    Color,
    SellEndDate
FROM Production.Product
WHERE Color = 'Black'
  AND ListPrice > 500
  AND SellEndDate IS NULL;

Now all three conditions must be true.

9. Combining Conditions with OR

OR requires at least one of the conditions to be true.

Find products that are either black or red:

USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE Color = 'Black'
   OR Color = 'Red';
AND vs OR

AND generally narrows a result set.

OR generally broadens a result set.

OR Across Different Columns

Conditions do not need to test the same column.

For example:

USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE Color = 'Black'
   OR ListPrice > 1000;

This returns:

black products at any price
products above 1000 regardless of colour
10. Using NOT

NOT reverses a condition.

USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE NOT Color = 'Black';

For this example, the following is usually clearer:

WHERE Color <> 'Black'

Both versions exclude rows where Color is NULL.

11. Parentheses and Operator Precedence

This is one of the most important parts of the lesson.

When AND and OR appear in the same condition, SQL Server does not simply evaluate them from left to right.

AND has higher precedence than OR.

Consider this query:

USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE Color = 'Black'
   OR Color = 'Red'
  AND ListPrice > 500;

A human might read this as:

Black or Red
AND
Price above 500

But SQL Server effectively interprets it as:

WHERE Color = 'Black'
   OR (Color = 'Red' AND ListPrice > 500);

That means:

every black product qualifies, regardless of price
red products qualify only when their price is above 500

This can produce incorrect results without producing a syntax error.

12. Fixing the Logic with Parentheses

If the requirement is:

Black or Red products, with both colours required to have a price above 500

write:

USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE (Color = 'Black' OR Color = 'Red')
  AND ListPrice > 500;

The parentheses force SQL Server to evaluate the colour condition first.

Practical Rule

Whenever you mix AND and OR, use parentheses to make the intended logic explicit.

Even when the parentheses are technically unnecessary, they can make the query much easier to understand.

13. Multiple Logical Groups

Parentheses also allow you to build separate groups of conditions.

For example:

Black products above 500
OR
Silver products below 100
USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE (Color = 'Black' AND ListPrice > 500)
   OR (Color = 'Silver' AND ListPrice < 100);

Each pair of conditions forms its own logical group.

14. GOTCHA — Do Not Use = NULL

A common beginner mistake is:

WHERE Color = NULL

For example:

USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color = NULL;

This returns no matching rows.

NULL represents an unknown or missing value.

You cannot test an unknown value using a normal equality comparison.

15. Use IS NULL

To find rows where the value is missing, use:

IS NULL

Example:

USE AdventureWorks2025;

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color IS NULL;
16. IS NOT NULL

To return rows where a value exists:

WHERE Color IS NOT NULL

Do not use:

Color != NULL

or:

Color <> NULL
Rule

Use:

IS NULL

or:

IS NOT NULL

when testing for NULL.

17. NULL and Not-Equal Comparisons

Consider:

WHERE Color <> 'Red'

You might expect this to return every row that is not red.

However, rows where Color is NULL are not included.

If you want both:

colours other than red
rows where no colour has been assigned

write:

WHERE Color <> 'Red'
   OR Color IS NULL;

We explore NULL behaviour in much more detail later in the course.

Complete Practice Script

The following queries bring together the main examples from this lesson.

USE AdventureWorks2025;

--------------------------------------------------
-- Filter text values
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE Color = 'Black';


--------------------------------------------------
-- Filter numeric values
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE ListPrice > 500;


--------------------------------------------------
-- Filter dates
--------------------------------------------------

SELECT
    SalesOrderID,
    OrderDate,
    TotalDue
FROM Sales.SalesOrderHeader
WHERE OrderDate >= '2025-01-01';


--------------------------------------------------
-- Column aliases cannot normally be used in WHERE
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice AS Price,
    Color
FROM Production.Product
WHERE ListPrice > 500;


--------------------------------------------------
-- Equal to
--------------------------------------------------

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color = 'Red';


--------------------------------------------------
-- Not equal to
--------------------------------------------------

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color <> 'Red';


--------------------------------------------------
-- Less than or equal to
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice
FROM Production.Product
WHERE ListPrice <= 100;


--------------------------------------------------
-- Date comparison
--------------------------------------------------

SELECT
    SalesOrderID,
    OrderDate,
    TotalDue
FROM Sales.SalesOrderHeader
WHERE OrderDate < '2025-06-01';


--------------------------------------------------
-- AND
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE Color = 'Black'
  AND ListPrice > 500;


--------------------------------------------------
-- Multiple AND conditions
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice,
    Color,
    SellEndDate
FROM Production.Product
WHERE Color = 'Black'
  AND ListPrice > 500
  AND SellEndDate IS NULL;


--------------------------------------------------
-- OR
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE Color = 'Black'
   OR Color = 'Red';


--------------------------------------------------
-- OR across different columns
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE Color = 'Black'
   OR ListPrice > 1000;


--------------------------------------------------
-- NOT
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE NOT Color = 'Black';


--------------------------------------------------
-- Incorrect AND / OR logic
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE Color = 'Black'
   OR Color = 'Red'
  AND ListPrice > 500;


--------------------------------------------------
-- Correct AND / OR logic
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE (Color = 'Black' OR Color = 'Red')
  AND ListPrice > 500;


--------------------------------------------------
-- Multiple logical groups
--------------------------------------------------

SELECT
    ProductID,
    Name,
    ListPrice,
    Color
FROM Production.Product
WHERE (Color = 'Black' AND ListPrice > 500)
   OR (Color = 'Silver' AND ListPrice < 100);


--------------------------------------------------
-- Incorrect NULL comparison
--------------------------------------------------

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color = NULL;


--------------------------------------------------
-- Correct NULL comparison
--------------------------------------------------

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color IS NULL;


--------------------------------------------------
-- Return non-NULL values
--------------------------------------------------

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color IS NOT NULL;


--------------------------------------------------
-- Not Red, including NULL values
--------------------------------------------------

SELECT
    ProductID,
    Name,
    Color
FROM Production.Product
WHERE Color <> 'Red'
   OR Color IS NULL;
Key Takeaways
WHERE controls which rows are returned by a query.
Text values use single quotes.
Numbers are normally written without quotes.
Dates in this course use the YYYY-MM-DD format.
Use the actual column name inside WHERE, not an alias defined in SELECT.
Use AND when every condition must be true.
Use OR when at least one condition must be true.
AND has higher precedence than OR.
Use parentheses whenever you mix AND and OR.
Never test for missing values using = NULL or <> NULL.
Use IS NULL and IS NOT NULL instead.
Next Video
Video 06 — LIKE, BETWEEN and IN

In the next lesson, we extend our filtering toolkit with three operators that come up constantly in practical SQL work:

BETWEEN for ranges
IN for matching against a list
LIKE for pattern matching in text
Course

This repository accompanies the T-SQL for Business Professionals YouTube course.

The course is designed for analysts, finance professionals, operations teams, Excel users, and other business professionals who want to learn how to work directly with data using SQL Server.

Database used: AdventureWorks2025
Database platform: Microsoft SQL Server
Query tool: SQL Server Management Studio (SSMS)
