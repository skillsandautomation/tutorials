How Data Is Organised in SQL Server

This folder accompanies Video 3 — How Data Is Organised: Databases, Schemas, Tables, and Constraints from the T-SQL for Business Professionals course.

Before we start writing T-SQL properly, it's important to understand where the data actually lives and how SQL Server organises it.

🎯 What You'll Learn

In this video, we'll explore:

The SQL Server hierarchy
Databases and schemas
Tables, rows, and columns
Common SQL Server data types
Primary and foreign keys
Constraints and data integrity
Indexes
How to explore an unfamiliar database using system queries
🗂️ The SQL Server Hierarchy

A useful mental model is:

Server
└── Database
    └── Schema
        └── Object
            └── Table
                ├── Columns
                ├── Keys
                ├── Constraints
                └── Indexes

For example:

SQL Server
└── AdventureWorks2025
    └── Production
        └── Product

This gives us the schema-qualified table name:

Production.Product

Throughout the course, we'll generally use these two-part object names when referencing tables.

📊 Tables, Rows, and Columns

A table normally represents one type of business entity, such as:

A product
A customer
An employee
A sales order

Each row represents one instance of that entity, while each column represents an attribute.

For example:

Production.Product

ProductID | Name | ProductNumber | Color | ListPrice

SQL Server also assigns a data type to every column.

Common examples include:

int
bigint
decimal
money
varchar
nvarchar
date
datetime2
bit

These data types define what kind of information SQL Server allows the column to contain.

🔑 Primary and Foreign Keys
Primary Key

A primary key uniquely identifies each row in a table.

For example:

Production.Product
        │
        └── ProductID  ← Primary Key

No two products can have the same ProductID.

Foreign Key

A foreign key creates a relationship between tables.

For example:

Sales.SalesOrderHeader
        │
        └── CustomerID
               │
               ▼
        Sales.Customer
               │
               └── CustomerID

This is the structural relationship that allows information stored across different tables to be connected.

We'll make extensive use of these relationships later when we learn JOINs.

🛡️ Constraints

SQL Server can enforce rules directly at the database level.

Some important constraints include:

Constraint	Purpose
PRIMARY KEY	Uniquely identifies each row
FOREIGN KEY	Creates and protects relationships between tables
NOT NULL	Requires a value
CHECK	Requires data to satisfy a condition
DEFAULT	Supplies a value automatically when one isn't provided
UNIQUE	Prevents duplicate values

These rules are one of the reasons data stored in a properly designed relational database can be considerably more controlled than data stored in a spreadsheet.

⚡ Indexes

Indexes are structures maintained by SQL Server to help queries find data efficiently.

You'll see clustered and non-clustered indexes inside Object Explorer.

We aren't designing indexes yet, but it's useful to understand that they exist because they play an important role in query performance.

We'll return to indexes later in the course.

🔍 Exploring an Unfamiliar Database

One of the most useful skills you can develop is being able to connect to an unfamiliar database and quickly understand its structure.

Here are three useful metadata queries from this video.

1. Find All Tables in the Database
USE AdventureWorks2025;

SELECT TABLE_SCHEMA,
       TABLE_NAME
FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_TYPE = 'BASE TABLE'
ORDER BY TABLE_SCHEMA,
         TABLE_NAME;

This gives you a quick inventory of all the physical tables available in the database, organised by schema.

2. Inspect All Columns in a Specific Table
SELECT COLUMN_NAME,
       DATA_TYPE,
       CHARACTER_MAXIMUM_LENGTH,
       IS_NULLABLE,
       COLUMN_DEFAULT
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = 'Production'
  AND TABLE_NAME = 'Product'
ORDER BY ORDINAL_POSITION;

This lets you inspect:

Column names
Data types
Maximum lengths
Whether NULL values are allowed
Default values

It's the query-based equivalent of opening the Columns folder in Object Explorer.

3. Inspect All Foreign Key Relationships in the Database
SELECT 
    fk.name AS ForeignKeyName,
    OBJECT_SCHEMA_NAME(fkc.parent_object_id) + '.' + tp.name AS ParentTable,
    cp.name AS ParentColumn,
    OBJECT_SCHEMA_NAME(fkc.referenced_object_id) + '.' + tr.name AS ReferencedTable,
    cr.name AS ReferencedColumn
FROM sys.foreign_keys fk
JOIN sys.foreign_key_columns fkc 
    ON fk.object_id = fkc.constraint_object_id
JOIN sys.tables tp 
    ON fkc.parent_object_id = tp.object_id
JOIN sys.columns cp 
    ON fkc.parent_object_id = cp.object_id 
    AND fkc.parent_column_id = cp.column_id
JOIN sys.tables tr 
    ON fkc.referenced_object_id = tr.object_id
JOIN sys.columns cr 
    ON fkc.referenced_object_id = cr.object_id 
    AND fkc.referenced_column_id = cr.column_id
ORDER BY ParentTable, fkc.constraint_column_id;

This query maps the foreign key relationships across the database and shows:

The foreign key name
The parent table
The foreign key column
The referenced table
The referenced column

This is particularly useful when you start working with an unfamiliar database. Before writing any JOINs, you can use this query to understand how the tables are connected.

Together, these three queries give you a useful starting toolkit for investigating almost any SQL Server database:

What tables exist? → What columns do they contain? → How are those tables related?

You can also investigate much of this without writing any T-SQL.

In SSMS:

AdventureWorks2025
└── Tables
    └── Production.Product
        ├── Columns
        ├── Keys
        ├── Constraints
        ├── Indexes
        ├── Triggers
        └── Statistics

Another useful option is:

Right-click Table
    → Script Table as
        → CREATE To
            → New Query Editor Window

This generates the T-SQL definition of the table and lets you inspect how it was actually created.

✅ Key Takeaway

The core hierarchy to remember is:

Server → Database → Schema → Object

Within our tables:

Rows contain records → Columns contain attributes → Primary keys identify rows → Foreign keys connect tables → Constraints protect the data.

Once you understand this structure, an unfamiliar corporate database becomes much easier to navigate.

In the next video, we'll move from understanding the database to querying it properly with the SELECT statement.
