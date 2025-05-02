@echo off
echo Starting clean build process...

REM Close any processes that might be locking files
taskkill /F /IM Southern_Starz_Checker.exe 2>nul

REM Create a temporary build directory outside of OneDrive
set BUILD_DIR=%TEMP%\starz_build_%RANDOM%
echo Using temporary build directory: %BUILD_DIR%
mkdir "%BUILD_DIR%"

REM Clean up any previous builds
echo Cleaning previous build files...
if exist build (
    rd /s /q build
)
if exist dist (
    rd /s /q dist
)
if exist *.spec (
    del *.spec
)

echo Installing required packages...
pip install pyinstaller
pip install -r requirements.txt

REM Copy necessary files to temp directory
echo Copying files to temporary directory...
xcopy inventory_checker.py "%BUILD_DIR%\" /Y
xcopy inventory_gui.py "%BUILD_DIR%\" /Y
xcopy requirements.txt "%BUILD_DIR%\" /Y
if exist README.md (
    xcopy README.md "%BUILD_DIR%\" /Y
)

REM Build in the temp directory
echo Building executable in temp directory...
cd /d "%BUILD_DIR%"
pyinstaller --clean --noconfirm --onefile --noconsole --name "Southern_Starz_Checker" inventory_gui.py

REM Check if build was successful
if not exist "dist\Southern_Starz_Checker.exe" (
    echo ERROR: Build failed! Executable not created.
    cd /d "%~dp0"
    goto error
)

REM Copy back the dist folder
echo Copying executable back to original directory...
mkdir "%~dp0\dist"
xcopy "dist\*.*" "%~dp0\dist\" /E /H /Y

REM Create batch file
echo Creating launcher batch file...
echo @echo off > "%~dp0\dist\Run_Inventory_Checker.bat"
echo echo Starting Southern Starz Inventory Checker... >> "%~dp0\dist\Run_Inventory_Checker.bat"
echo start Southern_Starz_Checker.exe >> "%~dp0\dist\Run_Inventory_Checker.bat"

REM Return to original directory
cd /d "%~dp0"

REM Clean up temp directory
echo Cleaning temporary build directory...
rd /s /q "%BUILD_DIR%"

echo.
echo Build completed successfully!
echo The executable is in the "dist" folder.
echo.
echo To distribute:
echo 1. Copy the entire "dist" folder to another PC
echo 2. Run the executable or the batch file inside
goto end

:error
echo.
echo Build process encountered errors.
echo.
rd /s /q "%BUILD_DIR%" 2>nul

:end
pause 