@echo off
echo ========================================
echo OTRS Ticket Analysis Setup Script
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in PATH
    echo Please install Python 3.6 or higher from:
    echo https://www.python.org/downloads/
    echo.
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

echo Python is installed
python --version
echo.

REM Check if pip is available
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo pip is not available
    echo Please ensure Python installation includes pip
    pause
    exit /b 1
)

echo pip is available
pip --version
echo.

REM Install required packages
echo Installing required packages...
echo.

pip install pandas openpyxl numpy

if %errorlevel% neq 0 (
    echo Failed to install required packages
    pause
    exit /b 1
)

echo.
echo ========================================
echo Installation completed successfully!
echo ========================================
echo.
echo You can now run the analysis script with:
echo python analyze_tickets.py [excel_file_path]
echo.
pause
