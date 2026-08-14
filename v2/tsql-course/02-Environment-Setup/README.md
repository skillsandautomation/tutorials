# Setting Up Your T-SQL Environment

This folder accompanies **Video 2 — Setting Up Your Environment and Running Your First Query** from the **T-SQL for Business Professionals** course.

In this video, we'll build the complete local environment used throughout the course and finish by running our first query against the AdventureWorks database.

## 🎯 What We'll Set Up

By the end of this video, you'll have:

* **SQL Server Express** installed as your database engine
* **SQL Server Management Studio (SSMS)** installed as your interface
* SSMS connected to your local SQL Server instance
* **AdventureWorks2025** restored as your practice database
* Your first query successfully executed

## 🧩 The Three Components

Think of the environment as three separate layers:

| Component              | Purpose                                                        |
| ---------------------- | -------------------------------------------------------------- |
| **SQL Server Express** | The database engine that stores and processes the data         |
| **SSMS**               | The interface where we write and execute T-SQL                 |
| **AdventureWorks**     | The sample business database we'll query throughout the course |

If you're coming from Excel, Excel effectively combines all three roles into one application. With SQL Server, the **engine, interface, and data are separate**.

## 1️⃣ Install SQL Server Express

Download and install the free **SQL Server Express** edition.

For this course, choose the **Basic** installation option.

Once installed, SQL Server Express should create the named instance:

```text
SQLEXPRESS
```

This is the SQL Server instance we'll connect to from SSMS.

## 2️⃣ Install SQL Server Management Studio

Install **SQL Server Management Studio (SSMS)**.

SSMS is where we'll spend most of our time during this course. We'll use it to:

* Connect to SQL Server
* Explore databases and tables
* Write T-SQL
* Execute queries
* View query results

## 3️⃣ Connect SSMS to SQL Server Express

Open SSMS and connect using:

```text
Server name: .\SQLEXPRESS
Authentication: Windows Authentication
```

You can also use:

```text
localhost\SQLEXPRESS
```

If your version of SSMS uses mandatory encryption for the connection, select:

```text
Trust Server Certificate
```

### Connection Problems?

If SSMS cannot connect, check these first:

1. Confirm the server name is `.\SQLEXPRESS`
2. Check that **SQL Server (SQLEXPRESS)** is running in Windows Services
3. Restart your computer after installing SQL Server Express

## 4️⃣ Load AdventureWorks

We'll use **AdventureWorks2025** throughout the course.

Download the standard **OLTP** backup:

```text
AdventureWorks2025.bak
```

In SSMS:

```text
Databases
   → Right-click
   → Restore Database
   → Device
   → Select AdventureWorks2025.bak
   → Restore
```

Once complete, you should see:

```text
Databases
└── AdventureWorks2025
```

This will be our working database for the rest of the course.

## 5️⃣ Run Your First Query

In Object Explorer, navigate to:

```text
AdventureWorks2025
└── Tables
    └── Production.Product
```

Right-click **Production.Product** and select:

```text
Select Top 1000 Rows
```

SSMS will automatically generate a T-SQL query.

Press:

```text
F5
```

to execute it.

Your first live SQL result set should appear in the **Results** window.

## ✅ Where You Should Be Now

At the end of this video, you should have:

**SQL Server Express** → Running
**SSMS** → Installed and connected
**AdventureWorks2025** → Restored
**Production.Product** → Successfully queried

Your T-SQL environment is now ready.

From the next video onward, we can stop setting things up and start learning what we can actually **do with the data**.
