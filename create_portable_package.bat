@echo off
echo Creating portable package of Southern Starz Inventory Checker...
echo ===========================================================
echo.

REM Check if the dist directory exists
if not exist dist (
    echo Error: dist directory not found!
    echo Please run build_installer.bat first to create the executable.
    pause
    exit /b 1
)

REM Create a zip file using PowerShell
echo Creating zip file...
powershell -command "Compress-Archive -Path 'dist\*' -DestinationPath 'Southern_Starz_Checker_Portable.zip' -Force"

echo.
echo =====================================================
echo Package created: Southern_Starz_Checker_Portable.zip
echo.
echo To use on your work computer:
echo 1. Copy the zip file to your work computer
echo 2. Extract the contents
echo 3. Double-click Southern_Starz_Checker.exe to run
echo =====================================================
echo.

pause 