-- ============================================================
-- Learn SQL in 30 Minutes — Quick Start for Business Professionals
-- Database: AdventureWorks2025
-- ============================================================

USE AdventureWorks2025;

-- ------------------------------------------------------------
-- 1. Your First Query — SELECT and FROM
-- ------------------------------------------------------------
SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM   Production.Product;


-- ------------------------------------------------------------
-- 2. Filtering Rows — WHERE
-- ------------------------------------------------------------

-- Filter by a numeric column (no quotes needed)
SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM   Production.Product
WHERE  ListPrice > 1000;

-- Filter by a text column (quotes required)
SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM   Production.Product
WHERE  Color = 'Red';

-- Combine conditions with AND
SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM   Production.Product
WHERE  Color = 'Red'
AND    ListPrice > 500;

-- Pattern matching with LIKE and wildcard %
SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM   Production.Product
WHERE  Name LIKE 'Mountain%';


-- ------------------------------------------------------------
-- 3. Sorting and Limiting — ORDER BY and TOP
-- ------------------------------------------------------------

SELECT ProductID,
       Name,
       ListPrice,
       Color
FROM   Production.Product
ORDER BY ListPrice DESC;

-- TOP + ORDER BY together = a meaningful "top N"
SELECT TOP 5
       ProductID,
       Name,
       ListPrice,
       Color
FROM   Production.Product
ORDER BY ListPrice DESC;


-- ------------------------------------------------------------
-- 4. Calculated Columns
-- ------------------------------------------------------------

-- Simple calculation with an alias (AS)
SELECT TOP 5
       ProductID,
       Name,
       ListPrice,
       ListPrice * 0.9 AS DiscountedPrice
FROM   Production.Product
ORDER BY ListPrice DESC;

-- CONCAT to build a readable label from multiple columns
SELECT TOP 5
       ProductID,
       CONCAT(Name, ' - ', Color) AS ProductLabel,
       ListPrice
FROM   Production.Product
ORDER BY ListPrice DESC;


-- ------------------------------------------------------------
-- 5. Summarising Data — GROUP BY and HAVING
-- ------------------------------------------------------------

-- Aggregate with no grouping = single overall number
SELECT SUM(TotalDue) AS TotalRevenue
FROM   Sales.SalesOrderHeader;

-- GROUP BY breaks the aggregate down by year
SELECT YEAR(OrderDate) AS OrderYear,
       SUM(TotalDue)   AS TotalRevenue,
       COUNT(*)        AS OrderCount
FROM   Sales.SalesOrderHeader
GROUP BY YEAR(OrderDate)
ORDER BY OrderYear;

-- HAVING filters groups AFTER aggregation (WHERE can't do this)
SELECT YEAR(OrderDate) AS OrderYear,
       SUM(TotalDue)   AS TotalRevenue,
       COUNT(*)        AS OrderCount
FROM   Sales.SalesOrderHeader
GROUP BY YEAR(OrderDate)
HAVING SUM(TotalDue) > 1000000
ORDER BY OrderYear;


-- ------------------------------------------------------------
-- 6. Bringing Tables Together — JOIN
-- ------------------------------------------------------------

-- INNER JOIN: only rows with a match on both sides
SELECT t.Name            AS Territory,
       SUM(soh.TotalDue) AS TotalRevenue,
       COUNT(*)          AS OrderCount
FROM   Sales.SalesOrderHeader AS soh
JOIN   Sales.SalesTerritory   AS t ON soh.TerritoryID = t.TerritoryID
GROUP BY t.Name
ORDER BY TotalRevenue DESC;

-- LEFT JOIN: keeps every row from the left table, even with no match
-- Here: find customers who have NEVER placed an order
SELECT c.CustomerID,
       soh.SalesOrderID
FROM      Sales.Customer         AS c
LEFT JOIN Sales.SalesOrderHeader AS soh ON c.CustomerID = soh.CustomerID
WHERE  soh.SalesOrderID IS NULL;
