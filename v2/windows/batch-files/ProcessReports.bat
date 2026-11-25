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
