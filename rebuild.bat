@echo off
setlocal

set "XLSX=output\budget_plan.xlsx"

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "STAMP=%%i"
set "ARCHIVE=output\budget_plan.%STAMP%.xlsx"

if exist "%XLSX%" (
    echo Archiving current workbook to %ARCHIVE%
    ren "%XLSX%" "budget_plan.%STAMP%.xlsx"
    if errorlevel 1 (
        echo ERROR: Could not rename current workbook.
        pause
        exit /b 1
    )
) else (
    echo No existing workbook to archive.
    set "ARCHIVE="
)

git pull
if errorlevel 1 (
    echo WARNING: git pull failed. Continuing with local version.
    pause
)

if defined ARCHIVE (
    echo Restoring local workbook from %ARCHIVE%
    copy /Y "%ARCHIVE%" "%XLSX%" >nul
    if errorlevel 1 (
        echo ERROR: Could not restore workbook from archive.
        pause
        exit /b 1
    )
)

if not exist .venv\Scripts\activate.bat (
    echo ERROR: Run setup.bat first.
    pause
    exit /b 1
)
call .venv\Scripts\activate.bat
python -m budget_planner rebuild
echo.
echo Workbook rebuilt in %XLSX%
echo Open it in Excel to recalculate formulas, then save.
