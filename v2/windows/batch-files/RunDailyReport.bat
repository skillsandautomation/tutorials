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
