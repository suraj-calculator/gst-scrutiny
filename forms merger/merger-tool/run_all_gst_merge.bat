@echo off
setlocal enabledelayedexpansion

REM ============================================================
REM   GST Merger Tool - runs all 4 merge scripts one by one
REM   Path yahan set karo - agar folder ka naam/location alag
REM   hai to sirf yeh line edit karo:
REM ============================================================
set "BASE_DIR=C:\Users\admin\Documents\merger\forms merger\merger-tool"

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python nahi mila is PC par.
    echo Pehle Python install karo aur "Add to PATH" tick karo, phir dubara try karo.
    pause
    exit /b
)

echo ============================================
echo   GST Merger Tool - Running all 4 scripts
echo ============================================

call :run_script "gstr1"     "merge_gstr1.py"
call :run_script "gstr2a"    "merge_r2a.py"
call :run_script "gstr2b"    "merge_gstr2b.py"
call :run_script "gstr3b"    "merge_gstr3b.py"
call :run_script "e invoice" "merge_einv.py"

echo.
echo ============================================
echo   Saare scripts complete ho gaye.
echo ============================================
pause
exit /b

:run_script
set "FOLDER=%~1"
set "SCRIPT=%~2"
echo.
echo --- %SCRIPT% chal raha hai "%FOLDER%" folder mein ---
pushd "%BASE_DIR%\%FOLDER%"
if errorlevel 1 (
    echo [FAIL] Folder nahi mila: %BASE_DIR%\%FOLDER%
    exit /b
)
python "%SCRIPT%"
if errorlevel 1 (
    echo [FAIL] %SCRIPT% mein error aayi. Upar dekho.
) else (
    echo [OK] %SCRIPT% successfully complete.
)
popd
exit /b
