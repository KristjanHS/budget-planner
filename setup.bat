@echo off
echo === Budget Planner - Setup ===
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python not found. Install Python 3.10+ from https://www.python.org/downloads/
    echo Make sure to check "Add to PATH" during installation.
    pause
    exit /b 1
)

echo Creating virtual environment...
python -m venv .venv
if errorlevel 1 (
    echo ERROR: Failed to create virtual environment.
    pause
    exit /b 1
)

echo Installing dependencies...
call .venv\Scripts\activate.bat
pip install -e . >nul 2>nul
if errorlevel 1 (
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
)

echo.
echo Setup complete. You can now run:
echo   init.bat     - Create an empty workbook
echo   demo.bat     - Create a workbook with demo data
echo   rebuild.bat  - Rebuild from an existing workbook
echo.
pause
