@echo off
echo Southern Starz Inventory Checker
echo ================================
echo.

REM Check if python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in PATH
    echo Please install Python 3.8 or higher from https://www.python.org/downloads/
    echo.
    pause
    exit /b
)

REM Check and install dependencies
echo Checking dependencies...
python -c "import pandas, pdfplumber, selenium, openpyxl, customtkinter" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing required packages...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo Failed to install dependencies.
        echo Please run "pip install -r requirements.txt" manually.
        pause
        exit /b
    )
)

echo.
echo Starting application...
python inventory_gui.py

if %errorlevel% neq 0 (
    echo.
    echo Application exited with an error.
    echo.
)

pause 