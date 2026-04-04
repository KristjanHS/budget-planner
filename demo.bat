@echo off
pushd "%~dp0"
if not exist .venv\Scripts\activate.bat (
    echo ERROR: Run setup.bat first.
    pause
    exit /b 1
)
call .venv\Scripts\activate.bat
python -m budget_planner init --demo
echo.
echo Demo workbook created in output\budget_plan.xlsx
echo Open it in Excel to recalculate formulas, then save.
pause
