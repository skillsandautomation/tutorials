
Dynamic Validation Engine in Excel VBA

Build a configurable data validation engine in Excel VBA where the validation rules live on a worksheet instead of being hard-coded into the VBA.

The engine reads each rule, substitutes field names with values from the current input row, evaluates the resulting expression with Application.Evaluate, and writes failed validations to an Exceptions sheet.

What This Project Demonstrates

Config-driven VBA

Application.Evaluate

Dynamic Arrays

Scripting.Dictionary

Late Binding

Rule-based validation

Exception reporting

Dynamic TRUE/FALSE evaluation

Handling blanks and dates correctly

Preserving the evaluated expression as an audit trail

Workbook Structure

Input

Contains the data to validate.

Example fields used in the video:

OrderID

Customer

Email

Region

Qty

UnitPrice

OrderDate

Rules

Contains the validation configuration.

Field

Rule

Severity

Message

Qty

[Qty] > 0

ERROR

Quantity must be positive

OrderID

LEN([OrderID]) = 8

ERROR

Order ID must be 8 characters

Email

[Email] <> ""

ERROR

Email is missing

UnitPrice

[UnitPrice] * [Qty] < 100000

WARN

Order value above review threshold

OrderDate

[OrderDate] <= TODAY()

ERROR

Order date cannot be in the future

Exceptions

The macro writes failed validations to this sheet.

Output columns:

Input Row

Field

Severity

Message

Rule Evaluated

How It Works

For every row in the Input sheet:

The row is loaded into a Scripting.Dictionary.

Each rule is copied into a working string.

Tokens such as [Qty] and [OrderDate] are replaced with the current row's values.

Dates are converted to their Excel serial number.

Numeric values are inserted directly.

Text and blank values are wrapped in quotes.

Application.Evaluate evaluates the completed expression.

If the result is False, an exception row is recorded.

The rules can therefore be changed without modifying the VBA itself.

VBA Code

The complete macro is in:

Validation_Engine.bas

You can also copy the code directly from below.

Sub Validation_Engine()

Dim wsIn As Worksheet, wsRules As Worksheet, wsOut As Worksheet
Set wsIn = ThisWorkbook.Worksheets("Input")
Set wsRules = ThisWorkbook.Worksheets("Rules")
Set wsOut = ThisWorkbook.Worksheets("Exceptions")

wsOut.Cells.Clear

If wsIn.Range("A2").Value = "" Then
    MsgBox "no data"
    Exit Sub
End If

Dim arrIn As Variant, arrRules As Variant, arrOut As Variant
Dim arrInHead As Variant

Dim lrow As Long, lcol As Long
lrow = wsIn.Cells(wsIn.Rows.Count, 1).End(xlUp).Row
lcol = wsIn.Cells(1, wsIn.Columns.Count).End(xlToLeft).Column

' Read everything in one shot - one touch of the worksheet
arrInHead = wsIn.Range(wsIn.Cells(1, 1), wsIn.Cells(1, lcol)).Value
arrIn = wsIn.Range(wsIn.Cells(2, 1), wsIn.Cells(lrow, lcol)).Value
arrRules = wsRules.Range("A1").CurrentRegion.Value

'1. Size the Exceptions array for the worst case: every row fails every rule
ReDim arrOut(1 To UBound(arrIn, 1) * (UBound(arrRules, 1) - 1) + 1, 1 To 5)

'2. Write the Exceptions header into Row 1 of the array
arrOut(1, 1) = "Input Row"
arrOut(1, 2) = "Field"
arrOut(1, 3) = "Severity"
arrOut(1, 4) = "Message"
arrOut(1, 5) = "Rule Evaluated"

'3. Create the Dictionary. Late Binding - no references needed, runs anywhere
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

Dim i As Long, j As Long, m As Long
Dim sExpr As String
Dim vKey As Variant, vValue As Variant, vResult As Variant
Dim lrowOut As Long
lrowOut = 1

'4. Validate the data. For every Input row, run every Rule
For i = LBound(arrIn, 1) To UBound(arrIn, 1)

    dict.RemoveAll ' each row starts fresh

    ' Load the row into the Dictionary. Key = Field Name, Value = cell
    For j = LBound(arrIn, 2) To UBound(arrIn, 2)
        dict(CStr(arrInHead(1, j))) = arrIn(i, j)
    Next j

    For m = 2 To UBound(arrRules, 1)

        sExpr = arrRules(m, 2)

        ' Substitute [FieldName] tokens with this row's values
        ' Text values get quote-wrapped so Evaluate treats them as literals
        ' Len check: a blank cell fools IsNumeric, so force blanks down the text branch
        ' VarType check: a Date fails IsNumeric, so substitute its serial number instead
        For Each vKey In dict.Keys

            vValue = dict(vKey)

            If VarType(vValue) = vbDate Then
                sExpr = Replace(sExpr, "[" & vKey & "]", CDbl(vValue))

            ElseIf IsNumeric(vValue) And Len(vValue) > 0 Then
                sExpr = Replace(sExpr, "[" & vKey & "]", vValue)

            Else
                sExpr = Replace(sExpr, "[" & vKey & "]",  & vValue & )
            End If

        Next vKey

        ' Hand the finished string to Excel's formula engine
        vResult = Application.Evaluate(sExpr)

        ' FALSE = rule broken. Log the exception
        If vResult = False Then

            lrowOut = lrowOut + 1

            arrOut(lrowOut, 1) = i + 1                  ' array row back to sheet row
            arrOut(lrowOut, 2) = arrRules(m, 1)         ' Field
            arrOut(lrowOut, 3) = arrRules(m, 3)         ' Severity
            arrOut(lrowOut, 4) = arrRules(m, 4)         ' Message
            arrOut(lrowOut, 5) = sExpr                  ' the evidence

        End If

    Next m

Next i

'5. Dump final array onto Exceptions sheet - only the rows actually filled
wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(lrowOut, 5)).Value = arrOut

MsgBox lrowOut - 1 & " exceptions found."

End Sub

Why the Blank Check Matters

A blank cell can behave unexpectedly with IsNumeric. The additional Len(vValue) > 0 check prevents blank values from being inserted into the expression as if they were numbers.

That allows a blank Email, for example, to become an empty quoted string rather than leaving a hole in the expression.

Why Dates Need Special Handling

A genuine Excel date is handled before the numeric/text branches:

If VarType(vValue) = vbDate Then
    sExpr = Replace(sExpr, "[" & vKey & "]", CDbl(vValue))

CDbl converts the date to the serial number Excel stores underneath. That allows a rule such as:

[OrderDate] <= TODAY()

to compare a number with a number.

Why the Output Array Is Sized for the Worst Case

Unlike a transformation engine, a validation engine does not know how many output rows it will produce.

A clean file may produce zero exceptions. A problematic file could produce one failed rule for every rule on every input row.

The macro therefore allocates enough space for the worst case once, before the loops start, and only writes the portion that was actually filled.

Requirements

Microsoft Excel

VBA enabled workbook

Three worksheets named:

Input

Rules

Exceptions

No reference needs to be added for Scripting.Dictionary because the code uses Late Binding.

Running the Project

Create the Input, Rules, and Exceptions worksheets.

Add your input data with headers in Row 1.

Add your validation rules to the Rules sheet.

Import Validation_Engine.bas into the VBA Editor, or paste the code into a standard module.

Run Validation_Engine.

Review the Exceptions sheet.

Extending the Engine

Possible enhancements include:

Rule IDs

Multiple severity levels

Rule categories

Exception filtering

Per-rule enable/disable flags

Better formatting for evaluated dates

Error handling for malformed rules

Logging rules that return an Excel error instead of TRUE/FALSE

Settings for worksheet names

A reusable workbook template

Related Video

This project is the companion to the Rules Engine project.

The Rules Engine uses worksheet expressions to produce values.

The Validation Engine uses worksheet expressions to test conditions and record exceptions.

Both use the same Config-Driven VBA idea:

The sheet contains the logic. VBA provides the engine.

YouTube

Watch the full step-by-step build on the Skills & Automation YouTube channel.

If you found the project useful, consider starring the repository and subscribing for more Excel VBA, automation, SQL, and programming tutorials for business professionals.
