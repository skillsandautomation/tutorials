# 00 - Learn SQL in 30 Minutes (Quick Start for Business Professionals)

📺 Watch: (YouTube link here)
🗄️ Database: AdventureWorks2025
⏱️ Duration: ~30 minutes

## Summary

A standalone quick-start video for business professionals with no coding or database background. In 30 minutes, it covers everything needed to go from zero to writing real T-SQL queries against a live database:

- Installing the 3 required tools (SQL Server Express, SSMS, AdventureWorks2025 sample database)
- **SELECT / FROM** — retrieving data from a table
- **WHERE** — filtering rows (comparisons, AND/OR, LIKE wildcards)
- **ORDER BY / TOP** — sorting and limiting results
- Calculated columns (AS, CONCAT)
- **GROUP BY / HAVING** — summarizing and filtering aggregated data
- **JOIN** — combining tables (INNER JOIN and LEFT JOIN)

This video moves fast on purpose. It's meant to get business professionals productive immediately. The full 35-video T-SQL course (see [`../`](../)) covers each of these topics in depth, with the reasoning behind every choice and common mistakes to avoid.

## Setup

1. Install **SQL Server Express** (free) — search "SQL Server Express download"
2. Install **SSMS** (SQL Server Management Studio) — search "download SSMS"
3. Download & restore **AdventureWorks2025** — search "AdventureWorks sample databases" on Microsoft Docs, then in SSMS: right-click *Databases* → *Restore Database*

## Code

See [`30-min-tsql-quickstart.sql`](./30-min-tsql-quickstart.sql) for every query used in the video, in order, with comments.

## Key Concepts

`SELECT`, `FROM`, schema-qualified table names, `WHERE`, comparison operators, comments (`--`), `AND`/`OR`, `LIKE` & wildcards (`%`), `ORDER BY`, `ASC`/`DESC`, `TOP`, calculated columns, `AS`, `CONCAT`, aggregate functions (`SUM`, `COUNT`), `GROUP BY`, `HAVING`, `JOIN`, `INNER JOIN`, table aliases, `ON`, `LEFT JOIN`, `NULL`
