# Windows Batch Files for Beginners

### Your First Automation Script | Skills and Automation

Learn how to automate everyday Windows tasks using batch files (`.bat`).

This repository contains the code demonstrated in my **Windows Batch Files for Beginners** YouTube tutorial.

We start by creating two sample CSV reports, build a batch file that copies and logs them, test command-line arguments and errors, and finally launch an existing Excel VBA automation using a second batch file.

**Requirements:** Windows 10/11, Command Prompt (CMD), Notepad, and optionally Microsoft Excel for the VBA example.

---

## 1. Initial Folder and File Setup

We'll work inside a folder called `CmdDemo` under the current Windows user's profile.

Open **Command Prompt (CMD)** and execute the following commands one at a time.

```bat
cd /d "%USERPROFILE%"

if not exist "CmdDemo" mkdir "CmdDemo"
cd CmdDemo

if not exist "Incoming" mkdir "Incoming"

echo Department,Amount>Incoming\Sales.csv
echo Sales,1200>>Incoming\Sales.csv

echo Department,Amount>Incoming\Expenses.csv
echo Operations,650>>Incoming\Expenses.csv

dir /b Incoming\*.csv
```

These commands create two sample CSV reports:

**Sales.csv**

```csv
Department,Amount
Sales,1200
```

**Expenses.csv**

```csv
Department,Amount
Operations,650
```

The initial folder structure should look like this:

```text
CmdDemo/
└── Incoming/
    ├── Sales.csv
    └── Expenses.csv
```

**Note:** These are dummy reports. Running the setup commands again will overwrite the sample CSV files.

---

## 2. Create ProcessReports.bat

Open Notepad and save a file named `ProcessReports.bat` inside `CmdDemo`.

Choose **Save as type: All files** so Windows does not append `.txt` to the filename.

Paste the following code into the file.

```bat
@echo off
rem Copy daily CSV reports into the Archive folder
setlocal
cd /d "%~dp0"

set "SourceFolder=Incoming"
if not "%~1"=="" set "SourceFolder=%~1"

echo Starting report processing...

if not exist "%SourceFolder%\*.csv" (
    echo ERROR: No CSV files found in "%SourceFolder%".
    >>"ProcessLog.txt" echo FAILED: No CSV files in "%SourceFolder%" at %DATE% %TIME%
    exit /b 1
)

if not exist "Archive" mkdir "Archive"
>>"ProcessLog.txt" echo Started by %USERNAME% at %DATE% %TIME%

copy /Y "%SourceFolder%\*.csv" "Archive\" >>"ProcessLog.txt" 2>&1

if errorlevel 1 (
    >>"ProcessLog.txt" echo FAILED: Copy error at %DATE% %TIME%
    echo ERROR: Copy failed. Check ProcessLog.txt.
    exit /b 1
)

>>"ProcessLog.txt" echo Completed at %DATE% %TIME%
echo Processing complete. Check Archive and ProcessLog.txt.
exit /b 0
```

### What the script does

1. Uses its own folder as the working directory.
2. Sets `Incoming` as the default source folder.
3. Allows a different source folder to be passed as an argument.
4. Checks for CSV files before processing.
5. Creates an `Archive` folder if required.
6. Copies CSV files into `Archive`.
7. Records execution details in `ProcessLog.txt`.
8. Returns an exit code indicating success or failure.

**Important:** `COPY /Y` overwrites matching files in `Archive` without asking for confirmation. This is a demonstration of repeatable processing, not versioned backup.

---

## 3. Run and Test the Batch File

Open CMD and navigate to the demonstration folder.

```bat
cd /d "%USERPROFILE%\CmdDemo"
```

### Test 1 — Run the script normally

```bat
ProcessReports.bat
```

Immediately inspect the exit code:

```bat
echo %ERRORLEVEL%
```

Expected result: `0` (success).

Check the files copied into `Archive`:

```bat
dir /b Archive
```

Expected files:

```text
Expenses.csv
Sales.csv
```

View the processing log:

```bat
type ProcessLog.txt
```

The log should contain entries showing when the processing started, the copying result, and when it completed.

### Test 2 — Pass a different source folder

First, create another folder with a sample CSV.

```bat
if not exist "OtherReports" mkdir "OtherReports"

echo Department,Amount>OtherReports\Extra.csv
echo Sales,2500>>OtherReports\Extra.csv
```

Now execute:

```bat
ProcessReports.bat OtherReports
```

The script will use `OtherReports` instead of `Incoming`.

Check the destination again:

```bat
dir /b Archive
```

You should now find `Extra.csv` alongside the original two reports.

### Test 3 — Test error handling

Run the script with a folder that does not exist.

```bat
ProcessReports.bat MissingFolder
```

Immediately check the exit code:

```bat
echo %ERRORLEVEL%
```

Expected result: `1` (failure).

Check the log:

```bat
type ProcessLog.txt
```

You should see a `FAILED` entry indicating that no CSV files were found in `MissingFolder`.

**Note:** Always inspect `%ERRORLEVEL%` immediately after running the batch file, before executing another command.

---

## 4. Bonus Example — Launch Excel VBA Using a Batch File

A batch file can also launch an existing application or automation.

In the video, we use a second batch file to open a macro-enabled Excel workbook. Its `Workbook_Open` event runs a VBA procedure that reads our two sample CSV reports and generates a summary.

### Create RunDailyReport.bat

Save another file named `RunDailyReport.bat` in the `CmdDemo` folder.

```bat
@echo off
rem Open the pre-built Excel VBA demo from the same folder as this batch file.
setlocal
cd /d "%~dp0"

if not exist "DailyReport.xlsm" (
    echo ERROR: DailyReport.xlsm was not found next to this batch file.
    exit /b 1
)

start "" "%~dp0DailyReport.xlsm"
echo Workbook launch requested. Check VBAOutput\DailyReportLog.txt and VBAOutput\DailySummary.csv.
rem IMPORTANT: A successful launch does not prove the VBA macro completed.
exit /b 0
```

This batch file checks that `DailyReport.xlsm` exists and then requests Windows to open it.

**Important:** The `START` command does not wait for the Excel macro to finish. A successful launch request is not proof that the VBA automation completed successfully.

---

## 5. Create the Excel VBA Workbook

This is an optional bonus example.

Open Microsoft Excel and create a blank workbook.

Save it as:

```text
%USERPROFILE%\CmdDemo\DailyReport.xlsm
```

Select **Excel Macro-Enabled Workbook (*.xlsm)** as the file type.

Press `Alt + F11` to open the VBA editor.

### Step A — Create the standard VBA module

Choose **Insert → Module**, and rename the module to `modDailyReport`.

Paste the code below into that module.

```vb
Option Explicit

' Demo only: each input CSV has a header and ONE data row
' containing an integer amount.

Public Sub GenerateDailyReport()

    Dim demoFolder As String
    Dim outputFolder As String
    Dim outputPath As String
    Dim sales As Long
    Dim expenses As Long
    Dim outputFile As Integer
    Dim outputOpen As Boolean
    Dim failureDetails As String

    On Error GoTo Failed

    demoFolder = ThisWorkbook.Path

    If Len(demoFolder) = 0 Then
        Err.Raise vbObjectError + 100, , _
                  "Save DailyReport.xlsm before running."
    End If

    outputFolder = demoFolder & "\VBAOutput"

    If Len(Dir$(outputFolder, vbDirectory)) = 0 Then
        MkDir outputFolder
    End If

    sales = ReadDemoAmount(demoFolder & "\Incoming\Sales.csv")
    expenses = ReadDemoAmount(demoFolder & "\Incoming\Expenses.csv")

    outputPath = outputFolder & "\DailySummary.csv"

    outputFile = FreeFile
    Open outputPath For Output As #outputFile
    outputOpen = True

    Print #outputFile, "Metric,Amount"
    Print #outputFile, "Sales," & sales
    Print #outputFile, "Expenses," & expenses
    Print #outputFile, "Net," & (sales - expenses)

    Close #outputFile
    outputOpen = False

    AppendStatus outputFolder & "\DailyReportLog.txt", _
                 "SUCCESS | " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
                 " | Created DailySummary.csv"

    Exit Sub

Failed:

    failureDetails = Err.Description

    On Error Resume Next

    If outputOpen Then Close #outputFile

    If Len(outputFolder) > 0 Then

        If Len(Dir$(outputFolder, vbDirectory)) = 0 Then
            MkDir outputFolder
        End If

        AppendStatus outputFolder & "\DailyReportLog.txt", _
                     "FAILED | " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
                     " | " & failureDetails
    End If

End Sub

Private Function ReadDemoAmount(ByVal csvPath As String) As Long

    Dim inputFile As Integer
    Dim inputOpen As Boolean
    Dim headerRow As String
    Dim dataRow As String
    Dim parts() As String
    Dim failureDetails As String

    On Error GoTo Failed

    inputFile = FreeFile

    Open csvPath For Input As #inputFile
    inputOpen = True

    If EOF(inputFile) Then
        Err.Raise vbObjectError + 101, , "CSV file is empty."
    End If

    Line Input #inputFile, headerRow

    If EOF(inputFile) Then
        Err.Raise vbObjectError + 102, , "CSV has no data row."
    End If

    Line Input #inputFile, dataRow

    parts = Split(dataRow, ",")

    If UBound(parts) < 1 Then
        Err.Raise vbObjectError + 103, , "CSV has no amount column."
    End If

    If Not IsNumeric(Trim$(parts(1))) Then
        Err.Raise vbObjectError + 104, , "Amount is not numeric."
    End If

    ReadDemoAmount = CLng(Trim$(parts(1)))

    Close #inputFile
    Exit Function

Failed:

    failureDetails = Err.Description

    On Error Resume Next

    If inputOpen Then Close #inputFile

    On Error GoTo 0

    Err.Raise vbObjectError + 105, _
              "ReadDemoAmount", csvPath & " - " & failureDetails

End Function

Private Sub AppendStatus(ByVal logPath As String, _
                         ByVal message As String)

    Dim logFile As Integer

    logFile = FreeFile

    Open logPath For Append As #logFile
    Print #logFile, message
    Close #logFile

End Sub
```

This procedure reads the amounts from the two sample CSV reports, calculates the net amount, writes a summary CSV and records a status entry.

The CSV reader is intentionally designed for the simple one-row integer examples used in this lesson. It is not a general-purpose CSV parser.

### Step B — Add the Workbook_Open event

In the VBA editor, find **ThisWorkbook** under Microsoft Excel Objects.

Double-click it and paste:

```vb
Private Sub Workbook_Open()

    On Error GoTo Finish

    modDailyReport.GenerateDailyReport

Finish:

    ' If this was the only workbook open,
    ' end that Excel session.
    ' Otherwise close only this demo workbook.

    On Error Resume Next

    If Application.Workbooks.Count = 1 Then
        Application.Quit
    Else
        ThisWorkbook.Close SaveChanges:=False
    End If

End Sub
```

The `Workbook_Open` event automatically calls `GenerateDailyReport` when the workbook opens.

Save and close the workbook.

Use Excel's normal macro trust and security controls. Do not bypass corporate macro policies or globally disable macro protections.

---

## 6. Run the Excel VBA Automation

Make sure that `DailyReport.xlsm` is closed and the batch file is saved in the same folder.

From CMD:

```bat
cd /d "%USERPROFILE%\CmdDemo"

RunDailyReport.bat
```

Or double-click `RunDailyReport.bat` in File Explorer.

If Excel is configured to allow the trusted macro, it will:

1. Open `DailyReport.xlsm`.
2. Trigger `Workbook_Open`.
3. Execute `GenerateDailyReport`.
4. Read the two CSV reports.
5. Generate a summary and status log.
6. Close the demonstration workbook.

### Expected output

The VBA procedure creates a `VBAOutput` folder containing:

```text
VBAOutput/
├── DailySummary.csv
└── DailyReportLog.txt
```

**DailySummary.csv**

```csv
Metric,Amount
Sales,1200
Expenses,650
Net,550
```

The status log should include a timestamped `SUCCESS` entry.

Check the newest log entry and the output file's modification time rather than relying on the batch file's launch message.

---

## 7. Final Folder Structure

Once both demonstrations have been completed successfully, your folder will look similar to this:

```text
CmdDemo/
│
├── Incoming/
│   ├── Sales.csv
│   └── Expenses.csv
│
├── OtherReports/
│   └── Extra.csv
│
├── Archive/
│   ├── Sales.csv
│   ├── Expenses.csv
│   └── Extra.csv
│
├── VBAOutput/
│   ├── DailySummary.csv
│   └── DailyReportLog.txt
│
├── ProcessReports.bat
├── ProcessLog.txt
├── RunDailyReport.bat
└── DailyReport.xlsm
```

Folders such as `Archive` and `VBAOutput`, and the log files, are created when their corresponding automations run.

---

## 8. A Few Important Notes

- All examples use disposable sample data.
- `ProcessReports.bat` overwrites matching CSV files in `Archive`.
- The batch-file log is appended to rather than overwritten.
- An exit code of `0` indicates reported success; a non-zero code indicates failure.
- The Excel launcher does not wait for VBA completion.
- The Excel example is intended for interactive desktop use, not as a production-ready background service.
- There is deliberately no `PAUSE` at the end of the batch files, so they do not wait for keyboard input.

For more advanced data processing, API calls and structured error handling, consider PowerShell rather than continuing to expand a batch script.

---

## What's Next?

In the next video, we'll use **Windows Task Scheduler** to run batch files at a chosen time, without manually double-clicking them.

We'll also discuss practical considerations such as execution accounts, permissions, working folders, and computers that are asleep or switched off.

**Skills and Automation**  
Practical coding and automation tutorials for business professionals.
