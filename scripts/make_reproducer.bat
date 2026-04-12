@echo off
setlocal enableextensions enabledelayedexpansion

REM ============================================================
REM  make_reproducer.bat  (Windows, runs via WSL)
REM
REM  Copy this file into the folder that holds your confidential
REM  .xlsx workbook. Double-click (or run from cmd) and it will:
REM    1. Find the single .xlsx in its own folder.
REM    2. Invoke the WSL python to run scripts.make_reproducer.
REM    3. Write fp.json + synthetic.xlsx into <REPO>/tmp/ in WSL.
REM
REM  No editing required. Assumes the repo lives at the WSL path
REM  below (adjust once if you ever move the checkout).
REM ============================================================

set "DISTRO=Ubuntu-22.04"
set "REPO_WSL=~/projects/budget-planner"

REM --- Resolve the folder this .bat lives in ---
set "HERE=%~dp0"

REM --- Find the single .xlsx in %HERE% (skip Excel ~$ lock files) ---
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

REM --- Convert the Windows path to a WSL path via wslpath ---
for /f "usebackq delims=" %%P in (`wsl -d %DISTRO% wslpath -a "%XLSX%"`) do set "XLSX_WSL=%%P"

if "%XLSX_WSL%"=="" (
    echo ERROR: wslpath failed. Is WSL distro "%DISTRO%" installed?
    pause
    exit /b 1
)

echo Real workbook : %XLSX%
echo WSL path      : %XLSX_WSL%
echo Repo (WSL)    : %REPO_WSL%
echo Output        : %REPO_WSL%/tmp/fp.json  +  synthetic.xlsx
echo.

wsl -d %DISTRO% bash -lc "cd '%REPO_WSL%' && python3 -m scripts.make_reproducer '%XLSX_WSL%'"
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo.
    echo make_reproducer failed with exit code %RC%.
    pause
    exit /b %RC%
)

echo.
echo Done. Artifacts in %REPO_WSL%/tmp/ are safe to share.
pause
endlocal
