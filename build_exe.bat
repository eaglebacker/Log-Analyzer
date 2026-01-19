@echo off
echo ============================================
echo    Log Analyzer - Build Executable
echo ============================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Installing required packages...
echo.

REM Install required packages
python -m pip install --upgrade pip
python -m pip install openpyxl pyinstaller Pillow tkinterdnd2

echo.
echo Creating Sasquatch icon...
echo.

REM Create the icon
python create_icon.py

echo.
echo Building executable with Sasquatch icon...
echo.

REM Build the executable with custom icon (output to "Log Analyzer" folder)
pyinstaller --onefile --windowed --name "LogAnalyzer" --icon=sasquatch.ico --distpath "Log Analyzer" log_analyzer_gui.py

echo.
echo ============================================
if exist "Log Analyzer\LogAnalyzer.exe" (
    echo SUCCESS! Executable created at:
    echo    Log Analyzer\LogAnalyzer.exe
    echo.
    echo You can move this file anywhere and run it!
) else (
    echo BUILD FAILED. Check the error messages above.
)
echo ============================================
echo.
pause
