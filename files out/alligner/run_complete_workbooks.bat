@echo off
REM =====================================================
REM  Complete Workbooks Tool - One Click Runner
REM  Ye .bat file usi folder ki saari .xlsx files ko
REM  process karegi jahan ye file rakhi hai.
REM =====================================================

REM ==== YAHAN APNA FOLDER PATH SET KARO ====
set TARGET_FOLDER=C:\Users\admin\Documents\merger\files out\alligner
REM ==========================================

cd /d "%~dp0"

echo.
echo ============================================
echo   GST Workbook Sheet-Completer Tool
echo ============================================
echo.
echo Target Folder: %TARGET_FOLDER%
echo.

if not exist "%TARGET_FOLDER%" (
    echo [ERROR] Ye folder nahi mila: %TARGET_FOLDER%
    echo Path check karo ya .bat file mein TARGET_FOLDER line update karo.
    echo.
    pause
    exit /b
)

REM Check karo Python installed hai ya nahi
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Python nahi mila. Pehle Python install karo:
    echo https://www.python.org/downloads/
    echo Install karte waqt "Add Python to PATH" tick karna mat bhoolna.
    echo.
    pause
    exit /b
)

REM Check karo openpyxl installed hai ya nahi, nahi to install kar do
python -c "import openpyxl" >nul 2>nul
if %errorlevel% neq 0 (
    echo [INFO] openpyxl install ho rahi hai, thoda ruko...
    pip install openpyxl
    echo.
)

REM Check karo complete_workbooks.py isi folder mein hai (jahan .bat file rakhi hai)
if not exist "complete_workbooks.py" (
    echo [ERROR] complete_workbooks.py is folder mein nahi mili.
    echo Is .bat file ko complete_workbooks.py ke saath ek hi folder mein rakho.
    echo.
    pause
    exit /b
)

echo Processing saari .xlsx files TARGET_FOLDER ki...
echo.

python complete_workbooks.py --folder "%TARGET_FOLDER%"

echo.
echo ============================================
echo   DONE! "_completed.xlsx" files banayi ja chuki hain.
echo ============================================
echo.
pause
