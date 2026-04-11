@echo off
git pull
if errorlevel 1 (
    echo WARNING: git pull failed. Continuing with local version.
    pause
)
if not exist .venv\Scripts\activate.bat (
    echo ERROR: Run setup.bat first.
    pause
    exit /b 1
)
call .venv\Scripts\activate.bat
python -m budget_planner init
echo.
echo Workbook created in output\budget_plan.xlsx
echo Open it in Excel to recalculate formulas, then save.
pause
