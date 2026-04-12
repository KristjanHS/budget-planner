@echo off
setlocal enableextensions enabledelayedexpansion

REM ============================================================
REM  make_reproducer.bat  (Windows)
REM
REM  Copy this file into the folder that holds your confidential
REM  .xlsx workbook. Double-click (or run from cmd) and it will:
REM    1. Find the single .xlsx in its own folder.
REM    2. Run scripts.make_reproducer against it.
REM    3. Write fp.json and synthetic.xlsx into <REPO>\tmp\.
REM
REM  BEFORE FIRST USE: edit the REPO= line below to point at
REM  your local checkout of the budget-planner repo.
REM ============================================================

set "REPO=C:\Users\%USERNAME%\projects\budget-planner"
set "PYTHON=python"

REM --- Resolve the folder this .bat lives in ---
set "HERE=%~dp0"
REM %~dp0 always ends with a backslash.

REM --- Find the single .xlsx in %HERE% (skip temp ~$ lock files) ---
set "XLSX="
set "XLSX_COUNT=0"
for %%F in ("%HERE%*.xlsx") do (
    set "NAME=%%~nxF"
    if not "!NAME:~0,2!"=="~$" (
        set /a XLSX_COUNT+=1
        set "XLSX=%%~fF"
    )
)

if "%XLSX_COUNT%"=="0" (
    echo ERROR: no .xlsx file found in "%HERE%".
    pause
    exit /b 1
)
if not "%XLSX_COUNT%"=="1" (
    echo ERROR: expected exactly one .xlsx in "%HERE%", found %XLSX_COUNT%.
    echo        Remove or move the extras and try again.
    pause
    exit /b 1
)

echo Real workbook : %XLSX%
echo Repo          : %REPO%
echo Output        : %REPO%\tmp\fp.json  +  synthetic.xlsx
echo.

if not exist "%REPO%\scripts\make_reproducer.py" (
    echo ERROR: REPO path does not look right — "%REPO%\scripts\make_reproducer.py" is missing.
    echo        Edit the REPO= line at the top of this .bat file.
    pause
    exit /b 1
)

pushd "%REPO%"
%PYTHON% -m scripts.make_reproducer "%XLSX%"
set "RC=%ERRORLEVEL%"
popd

if not "%RC%"=="0" (
    echo.
    echo make_reproducer failed with exit code %RC%.
    pause
    exit /b %RC%
)

echo.
echo Done. Artifacts in %REPO%\tmp\ are safe to share.
pause
endlocal
