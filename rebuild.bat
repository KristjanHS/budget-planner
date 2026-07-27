@echo off
setlocal

set "XLSX=output\budget_plan.xlsx"

if not exist .venv\Scripts\activate.bat (
    echo ERROR: Run setup.bat first.
    pause
    exit /b 1
)
call .venv\Scripts\activate.bat

REM No archive step here: `rebuild` backs up %XLSX% to
REM <stem>.<YYYYmmdd_HHMMSS>.xlsx itself (workbook._backup_existing) right
REM before overwriting. Calling `python -m budget_planner archive` first produced a
REM second, identical timestamped copy every run. Timestamps stay in Python —
REM `for /f` over `powershell Get-Date` leaks PS stderr into captured stdout.
git pull
if errorlevel 1 (
    echo WARNING: git pull failed. Continuing with local version.
    pause
)

python -m budget_planner rebuild
if errorlevel 1 (
    echo.
    echo ERROR: rebuild failed - workbook NOT regenerated. See the error above.
    echo        Your data is intact in %XLSX% - rebuild backs up and overwrites
    echo        only as its final step, so a failure leaves the file untouched.
    pause
    exit /b 1
)
echo.
echo Workbook rebuilt in %XLSX%
echo Open it in Excel to recalculate formulas, then save.
