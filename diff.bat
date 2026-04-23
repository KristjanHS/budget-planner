@echo off
setlocal enableextensions

set "DIFFDIR=diff"
set "OUT=%DIFFDIR%\diff_result.json"

if not exist "%DIFFDIR%\" (
    echo ERROR: "%DIFFDIR%" folder not found.
    pause
    exit /b 1
)

if not exist .venv\Scripts\activate.bat (
    echo ERROR: Run setup.bat first.
    pause
    exit /b 1
)

REM File enumeration, oldest/newest selection, and archival of any existing
REM diff_result.json are all handled inside the Python script. diff.bat is
REM pure cmd - no PowerShell dependency.
REM
REM Default output: --mode=logic + compact JSON (context-friendly for AI
REM consumers; drops pure-style cell noise). For human-readable full output,
REM invoke the script directly with --mode=full --pretty.
call .venv\Scripts\activate.bat
python scripts\xlsx_structural_diff.py --dir "%DIFFDIR%" --out "%OUT%"
if errorlevel 1 (
    echo.
    echo xlsx_structural_diff failed.
    pause
    exit /b 1
)

echo.
echo Diff written to %OUT%
pause
endlocal
