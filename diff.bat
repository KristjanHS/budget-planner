@echo off
setlocal enableextensions enabledelayedexpansion

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

REM --- Find oldest and newest .xlsx in diff\ (skip Excel ~$ lock files) ---
set "OLDEST="
set "NEWEST="
for /f "usebackq delims=" %%P in (`powershell -NoProfile -Command "Get-ChildItem -LiteralPath '%DIFFDIR%' -Filter '*.xlsx' -File ^| Where-Object { -not $_.Name.StartsWith('~$') } ^| Sort-Object LastWriteTime ^| Select-Object -First 1 -ExpandProperty FullName"`) do set "OLDEST=%%P"
for /f "usebackq delims=" %%P in (`powershell -NoProfile -Command "Get-ChildItem -LiteralPath '%DIFFDIR%' -Filter '*.xlsx' -File ^| Where-Object { -not $_.Name.StartsWith('~$') } ^| Sort-Object LastWriteTime ^| Select-Object -Last 1 -ExpandProperty FullName"`) do set "NEWEST=%%P"

if not defined OLDEST (
    echo ERROR: No .xlsx files found in "%DIFFDIR%".
    pause
    exit /b 1
)

if /I "%OLDEST%"=="%NEWEST%" (
    echo ERROR: Only one .xlsx file in "%DIFFDIR%"; need at least two to diff.
    pause
    exit /b 1
)

echo Oldest : %OLDEST%
echo Newest : %NEWEST%
echo Output : %OUT%
echo.

REM --- Archive any existing diff result ---
if exist "%OUT%" (
    for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "STAMP=%%i"
    set "ARCHIVE=%DIFFDIR%\diff_result.!STAMP!.json"
    echo Archiving previous result to !ARCHIVE!
    ren "%OUT%" "diff_result.!STAMP!.json"
    if errorlevel 1 (
        echo ERROR: Could not archive previous diff result.
        pause
        exit /b 1
    )
)

call .venv\Scripts\activate.bat
python scripts\xlsx_structural_diff.py "%OLDEST%" "%NEWEST%" --out "%OUT%"
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
