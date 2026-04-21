@echo off
git pull
if errorlevel 1 (
    echo WARNING: git pull failed. Continuing with local version.
    pause
)
echo === Budget Planner - Setup ===
echo.

py -3.14 --version >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python 3.14 not found via the py launcher.
    echo Install Python 3.14 from https://www.python.org/downloads/
    echo During install, check "Add to PATH" and "Install launcher".
    pause
    exit /b 1
)

if exist .venv (
    echo Removing stale .venv...
    rmdir /s /q .venv
    if errorlevel 1 (
        echo ERROR: Failed to remove existing .venv directory. Close any process using it and retry.
        pause
        exit /b 1
    )
)

echo Creating virtual environment with Python 3.14...
py -3.14 -m venv .venv
if errorlevel 1 (
    echo ERROR: Failed to create virtual environment.
    pause
    exit /b 1
)

echo Installing dependencies...
call .venv\Scripts\activate.bat
pip install -e .
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
