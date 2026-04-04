@echo off
if not exist .venv\Scripts\activate.bat (
    echo ERROR: Run setup.bat first.
    pause
    exit /b 1
)
call .venv\Scripts\activate.bat
python -m budget_planner rebuild
echo.
echo Workbook rebuilt in output\budget_plan.xlsx
echo Open it in Excel to recalculate formulas, then save.
pause
