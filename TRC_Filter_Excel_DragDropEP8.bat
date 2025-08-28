@echo off
REM TRC Filter Excel - Drag & Drop Launcher
REM This batch file enables drag & drop functionality for TRC_Filter_Excel_3.py

echo TRC Filter Excel - Drag ^& Drop Launcher
echo ======================================
echo.

if "%~1"=="" (
    echo No files provided. Please drag and drop TRC log files onto this batch file.
    echo.
    echo Usage: Drag TRC log files onto this .bat file
    echo.
    pause
    exit /b 1
)

echo Processing files:
set file_count=0
for %%f in (%*) do (
    echo   - %%~nxf
    set /a file_count+=1
)

echo.
echo Total files: %file_count%
echo.

REM Run the Python script with all dropped files
python "%~dp0TRC_Filter_Excel_3_EP8.py" %*

if errorlevel 1 (
    echo.
    echo ❌ Error occurred during processing.
    echo.
    pause
    exit /b 1
) else (
    echo.
    echo ✅ Processing completed successfully!
    echo.
    pause
)
