@echo off
echo Building Southern Starz Inventory Checker Executable...
echo =====================================================
echo.

REM Check if Python is installed
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Install PyInstaller and required packages
echo Installing required packages...
pip install -U pyinstaller
pip install -r requirements.txt

REM Create directories for build
if not exist build mkdir build
if not exist dist mkdir dist

REM Create Spec File
echo Creating build specification...
echo from PyInstaller.building.api import PYZ, EXE, COLLECT
echo from PyInstaller.building.build_main import Analysis
echo block_cipher = None
echo.
echo a = Analysis(['inventory_gui.py'],
echo              pathex=['.'],
echo              binaries=[],
echo              datas=[],
echo              hiddenimports=['PIL._tkinter_finder', 'openpyxl', 'pandas', 'pdfplumber', 'selenium', 'customtkinter', 
echo                             'tkinter', 'webdriver_manager.chrome', 'selenium.webdriver.chrome.service'],
echo              hookspath=[],
echo              runtime_hooks=[],
echo              excludes=[],
echo              win_no_prefer_redirects=False,
echo              win_private_assemblies=False,
echo              cipher=block_cipher,
echo              noarchive=False)
echo.
echo pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
echo.
echo exe = EXE(pyz,
echo           a.scripts,
echo           a.binaries,
echo           a.zipfiles,
echo           a.datas,
echo           [],
echo           name='Southern_Starz_Checker',
echo           debug=False,
echo           bootloader_ignore_signals=False,
echo           strip=False,
echo           upx=True,
echo           upx_exclude=[],
echo           runtime_tmpdir=None,
echo           console=False,
echo           icon='NONE')
echo. > build\Southern_Starz_Checker.spec

REM Run PyInstaller
echo.
echo Building executable (this may take a few minutes)...
pyinstaller --clean --noconfirm --onefile --windowed --name "Southern_Starz_Checker" inventory_gui.py

REM Copy additional files
echo.
echo Creating distribution package...
copy requirements.txt dist\
copy README.md dist\
echo @echo off > dist\Run_Inventory_Checker.bat
echo echo Starting Southern Starz Inventory Checker... >> dist\Run_Inventory_Checker.bat
echo start Southern_Starz_Checker.exe >> dist\Run_Inventory_Checker.bat
echo echo. >> dist\Run_Inventory_Checker.bat

REM Create a simple README if it doesn't exist
if not exist README.md (
    echo # Southern Starz Inventory Checker > dist\README.md
    echo. >> dist\README.md
    echo ## How to Use >> dist\README.md
    echo. >> dist\README.md
    echo 1. Double-click on 'Southern_Starz_Checker.exe' to run the application >> dist\README.md
    echo 2. Click 'Browse' to select a PDF inventory file >> dist\README.md
    echo 3. Click 'Process Inventory' to analyze the inventory >> dist\README.md
    echo 4. When complete, click 'Download Excel' to open the generated report >> dist\README.md
)

echo.
echo =================================================
echo Build complete! Your executable is in the 'dist' folder.
echo.
echo To use on another computer:
echo 1. Copy the entire 'dist' folder to your work computer
echo 2. Double-click Southern_Starz_Checker.exe to run
echo =================================================
echo.

pause 