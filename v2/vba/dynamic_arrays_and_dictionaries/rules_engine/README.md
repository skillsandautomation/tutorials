# Excel VBA Rules Engine

Build a reusable **Rules Engine in Excel VBA** where the transformation logic lives in an Excel worksheet instead of being hard-coded inside the VBA macro.

The core idea is simple:

> **Change the logic without changing the code.**

Instead of writing separate VBA logic for every business rule, the workbook uses a **Mapping sheet** to define what the output should contain and how each output field should be calculated.

The VBA engine simply reads those rules and executes them.

---

## What This Project Does

The workbook contains three main worksheets:

* **Input** — contains the raw source data
* **Mapping** — contains the transformation rules
* **Output** — contains the transformed result

The Mapping sheet has two columns:

| Field Name | Expression                       |
| ---------- | -------------------------------- |
| OrderID    | `[OrderID]`                      |
| FullName   | `[FirstName] & " " & [LastName]` |
| LineTotal  | `[Qty] * [UnitPrice]`            |

Anything inside square brackets represents a field from the Input data.

Everything else is written using normal Excel formula syntax.

For example:

```text
[Qty] * [UnitPrice]
```

For an input row where:

```text
Qty = 5
UnitPrice = 12.5
```

the engine converts the expression into:

```text
5 * 12.5
```

and passes it to Excel's formula engine using:

```vba
Application.Evaluate
```

The result is then written into the Output array.

---

## Why Build It This Way?

A traditional VBA macro might contain logic such as:

```text
Column 2 + Column 3 = Full Name
Column 4 × Column 5 = Line Total
```

That works until the business changes its requirements.

Then somebody has to open the VBA editor and modify the code.

With this Rules Engine, the transformation logic is stored in Excel instead.

For example, this:

```text
[Qty] * [UnitPrice]
```

can be changed to:

```text
ROUND([Qty] * [UnitPrice], 2)
```

without modifying the VBA code.

You can also add an entirely new output column:

```text
Category
```

with:

```text
IF([Qty] > 100, "Bulk", "Standard")
```

Run the engine again and the new column is automatically created.

---

## The Mapping Sheet Controls the Output

Each row in the Mapping sheet represents one output column.

That means the Mapping sheet controls both:

1. **What fields appear in the Output**
2. **How each field is calculated**

A simple pass-through field can use:

```text
[OrderID]
```

A calculated field can use:

```text
[Qty] * [UnitPrice]
```

A text transformation can use:

```text
UPPER([LastName]) & ", " & [FirstName]
```

A conditional rule can use:

```text
IF([Qty] > 100, "Bulk", "Standard")
```

Because the engine uses `Application.Evaluate`, it can take advantage of Excel's existing formula language.

---

## Key VBA Concepts Used

This project brings together several useful intermediate VBA concepts.

### Dynamic Arrays

The Input, Mapping and Output data are loaded into arrays so that most of the processing happens in memory rather than repeatedly reading and writing worksheet cells.

This is significantly faster for larger datasets.

### Scripting.Dictionary

For each Input row, a Dictionary is created where:

```text
Key   = Field Name
Value = Current Row Value
```

For example:

```text
"Qty"       → 5
"UnitPrice" → 12.5
"FirstName" → John
```

This allows the engine to locate data by **field name instead of column position**.

If somebody rearranges the Input columns, the engine can still work because the Mapping rules reference names such as:

```text
[Qty]
```

rather than:

```text
Column 4
```

### Token Substitution

The engine searches each Mapping expression for bracketed field names.

For example:

```text
[Qty] * [UnitPrice]
```

becomes:

```text
5 * 12.5
```

The same process works for text:

```text
[FirstName] & " " & [LastName]
```

becomes:

```text
"John" & " " & "Smith"
```

Text values must be wrapped in quotation marks before being passed to `Application.Evaluate`.

### Application.Evaluate

Once the field references have been replaced with actual values, VBA passes the completed expression to Excel's formula engine.

For example:

```vba
Application.Evaluate("5 * 12.5")
```

returns:

```text
62.5
```

This means the VBA engine does not need separate code for every possible formula.

---

## Reuse the Same Engine With Different Data

The important part of this project is that the VBA code does not know anything about orders, quantities, prices or payroll.

For example, you could replace the original Input data with:

| EmpID | FirstName | LastName | HoursWorked | HourlyRate |
| ----- | --------- | -------- | ----------: | ---------: |

Then replace the Mapping rules with:

| Field Name | Expression                               |
| ---------- | ---------------------------------------- |
| EmpID      | `[EmpID]`                                |
| Employee   | `UPPER([LastName]) & ", " & [FirstName]` |
| GrossPay   | `ROUND([HoursWorked] * [HourlyRate], 2)` |

The same VBA engine can then produce a payroll output without changing the code.

That is why this project is better thought of as a **template** rather than a single-purpose macro.

---

# Complete VBA Code

```vba
Sub Rules_Engine()

Dim wsIn As Worksheet, wsMap As Worksheet, wsOut As Worksheet

Set wsIn = ThisWorkbook.Worksheets("Input")
Set wsMap = ThisWorkbook.Worksheets("Mapping")
Set wsOut = ThisWorkbook.Worksheets("Output")

wsOut.Cells.Clear

If wsIn.Range("A2").Value = "" Then
    MsgBox "no data"
    Exit Sub
End If

Dim arrIn As Variant, arrOut As Variant, arrMap As Variant
Dim arrInHead As Variant, arrOutHead As Variant

Dim lrow As Long, lcol As Long

lrow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row
lcol = wsIn.Cells(1, wsIn.Columns.Count).End(xlToLeft).Column

' Read everything in one shot - one touch of the worksheet
arrInHead = wsIn.Range(wsIn.Cells(1, 1), wsIn.Cells(1, lcol)).Value
arrIn = wsIn.Range(wsIn.Cells(2, 1), wsIn.Cells(lrow, lcol)).Value
arrMap = wsMap.Range("A1").CurrentRegion.Value

'1. Build the Output Header from the Mapping array
ReDim arrOutHead(1 To 1, 1 To UBound(arrMap, 1) - 1)

Dim m As Long

For m = 2 To UBound(arrMap, 1)
    arrOutHead(1, m - 1) = arrMap(m, 1)
Next m

'2. Paste Output Header onto Output Worksheet
wsOut.Range(wsOut.Cells(1, 1), _
            wsOut.Cells(1, UBound(arrOutHead, 2))).Value = arrOutHead

'3. Size the Output array. One row per Input row. One column per Mapping row
ReDim arrOut(1 To UBound(arrIn, 1), _
             1 To UBound(arrMap, 1) - 1)

'4. Create the Dictionary. Late Binding - no references needed, runs anywhere
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

Dim i As Long, j As Long
Dim sExpr As String
Dim vKey As Variant, vValue As Variant

'5. Transfer the data. For every Input row, run every Mapping expression
For i = LBound(arrIn, 1) To UBound(arrIn, 1)

    dict.RemoveAll ' each row starts fresh

    ' Load the row into the Dictionary. Key = Field Name, Value = cell
    For j = LBound(arrIn, 2) To UBound(arrIn, 2)
        dict(CStr(arrInHead(1, j))) = arrIn(i, j)
    Next j

    For m = 2 To UBound(arrMap, 1)

        sExpr = arrMap(m, 2)

        ' Substitute [FieldName] tokens with this row's values
        ' Text values get quote-wrapped so Evaluate treats them as literals
        For Each vKey In dict.Keys

            vValue = dict(vKey)

            If IsNumeric(vValue) Then
                sExpr = Replace(sExpr, "[" & vKey & "]", vValue)
            Else
                sExpr = Replace(sExpr, "[" & vKey & "]", """" & vValue & """")
            End If

        Next vKey

        ' Hand the finished string to Excel's formula engine
        arrOut(i, m - 1) = Application.Evaluate(sExpr)

    Next m

Next i

'6. Dump final array onto Output sheet - one touch on the way out
wsOut.Range(wsOut.Cells(2, 1), _
            wsOut.Cells(UBound(arrOut, 1) + 1, _
            UBound(arrOut, 2))).Value = arrOut

End Sub
```

---

## How to Use the Template

### 1. Add Your Input Data

Place your source data on the **Input** worksheet.

The first row must contain field names.

Example:

```text
OrderID | FirstName | LastName | Qty | UnitPrice
```

The data begins on Row 2.

---

### 2. Create Your Mapping Rules

On the **Mapping** worksheet, create two columns:

```text
Field Name | Expression
```

For example:

```text
OrderID   | [OrderID]
FullName  | [FirstName] & " " & [LastName]
LineTotal | [Qty] * [UnitPrice]
```

---

### 3. Run the Macro

Run:

```vba
Rules_Engine
```

The engine will:

1. Read the Input data
2. Read the Mapping rules
3. Build the required Output columns
4. Substitute the current row's values into each expression
5. Evaluate the expression
6. Populate the Output worksheet

---

## Example: Adding a Rule Without Changing VBA

Suppose the business asks for Line Total to be rounded to two decimal places.

Change:

```text
[Qty] * [UnitPrice]
```

to:

```text
ROUND([Qty] * [UnitPrice], 2)
```

Run the macro again.

No VBA changes are required.

---

## Example: Adding a New Output Field

Add another Mapping row:

```text
Category | IF([Qty] > 100, "Bulk", "Standard")
```

Run the macro.

The Output automatically gains a new `Category` column.

Again, no changes to the VBA engine are required.

---

## Important Limitations

The version demonstrated in the video is deliberately kept relatively simple.

There are several improvements you may want to make for production use.

### 1. Sheet Names Are Still Hard-Coded

The following worksheet names are currently inside the VBA:

```vba
Set wsIn = ThisWorkbook.Worksheets("Input")
Set wsMap = ThisWorkbook.Worksheets("Mapping")
Set wsOut = ThisWorkbook.Worksheets("Output")
```

These could also be moved into a configuration area.

---

### 2. Per-Row Error Handling

A bad Mapping expression can currently interrupt execution.

A production version could trap errors and write information such as:

```text
Input Row
Mapping Rule
Expression
Error Description
```

to a separate error log.

---

### 3. Similar Field Names

If your Input contains names such as:

```text
[Price]
```

and:

```text
[PriceUSD]
```

the Dictionary keys should be processed longest-first.

Otherwise `[Price]` could accidentally match part of `[PriceUSD]`.

---

### 4. Dates

This version explicitly handles numbers and text.

Dates require additional treatment before being passed into `Application.Evaluate`.

One approach is to convert them into an expression that Excel can reliably evaluate, such as using `DATEVALUE`.

---

### 5. Application.Evaluate Expression Length

`Application.Evaluate` has practical limitations on the length of the expression it can evaluate.

For very large or complex expressions, a different evaluation strategy may be required.

---

## The Architecture

The important lesson from this project is not the individual VBA syntax.

It is the architecture:

```text
INPUT DATA
     ↓
FIELD NAMES + VALUES
     ↓
DICTIONARY
     ↓
MAPPING RULE
     ↓
TOKEN SUBSTITUTION
     ↓
EXCEL EXPRESSION
     ↓
Application.Evaluate
     ↓
OUTPUT ARRAY
     ↓
OUTPUT SHEET
```

The VBA code becomes the **engine**.

The worksheet becomes the **configuration**.

And the business rules become **data instead of code**.

---

## YouTube Tutorial

🎥 **Build a Dynamic Rules Engine in Excel VBA — Change the Logic Without Changing the Code**

[ADD YOUTUBE VIDEO LINK]

In the full tutorial, I build the engine from an empty VBA module and explain each part step by step.

---

## Related Tutorials

* **VBA Dynamic Arrays** — [ADD LINK]
* **VBA Dictionaries** — [ADD LINK]
* **Config-Driven Data Validation Engine** — [ADD LINK]

---

## Download

You can download the completed Excel workbook from this repository and replace the sample Input data and Mapping rules with your own.

The idea is to use the workbook as a reusable starting point for different transformation requirements.

---

## Skills & Automation

This project is part of my **Skills & Automation** Excel VBA tutorial series.

The focus is practical automation for business users, analysts and developers who want to build reusable solutions in Excel.

If you found the project useful, check out the full video tutorial and the other VBA projects in this repository.
