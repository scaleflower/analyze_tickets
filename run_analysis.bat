@echo off
echo ========================================
echo OTRS Ticket Analysis Runner
echo ========================================
echo.

REM Check if requirements are installed
python -c "import pandas, openpyxl, numpy" >nul 2>&1
if %errorlevel% neq 0 (
    echo Required packages not found. Running installation...
    call install_requirements.bat
    if %errorlevel% neq 0 (
        echo Installation failed. Please check the requirements.
        pause
        exit /b 1
    )
)

echo.
echo ========================================
echo Starting OTRS Ticket Analysis...
echo ========================================
echo.

REM Check if file path is provided as argument
if "%~1"=="" (
    echo No Excel file specified, using default file...
    python analyze_tickets.py
) else (
    echo Using specified file: %~1
    python analyze_tickets.py "%~1"
)

echo.
echo ========================================
echo Analysis completed!
echo ========================================
echo.
pause
