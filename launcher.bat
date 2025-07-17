@echo off
setlocal

:: --- Configuration ---
set PYTHON_SCRIPT=adVideo_UI.py

title Ad Video Tool Launcher

:: --- 1. Check for Python installation ---
echo Checking for Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not found in your system's PATH.
    pause
    exit /b
)

:: --- 2. Install required packages directly ---
echo Installing required packages (if missing)...
pip install pandas openpyxl requests --upgrade --quiet --disable-pip-version-check

:: --- 3. Launch the UI and Exit ---
echo Launching the Ad Video Tool...

:: Use "pythonw.exe" to run the script without a console window.
:: The "start" command ensures this batch script doesn't wait for the UI to close.
:: The empty "" is a required placeholder for the start command's window title.
start "" pythonw "%PYTHON_SCRIPT%"

endlocal
exit /b
