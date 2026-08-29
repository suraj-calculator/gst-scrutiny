@echo off
setlocal enabledelayedexpansion

set "TOOL_DIR=C:\Users\admin\Documents\merger\merger\pdf to excel
cd /d "%TOOL_DIR%"

echo ============================================
echo  GST PDF to Excel Converter
echo ============================================
echo  Folder: %TOOL_DIR%
echo.

set "PY_CMD="
where python >nul 2>nul
if not errorlevel 1 set "PY_CMD=python"

if not defined PY_CMD (
    where py >nul 2>nul
    if not errorlevel 1 set "PY_CMD=py"
)

if not defined PY_CMD (
    echo ERROR: Python was not found on this computer.
    echo Install it from https://www.python.org/downloads/
    echo IMPORTANT: during setup, tick the box "Add python.exe to PATH".
    echo.
    pause
    exit /b 1
)

echo Using: %PY_CMD%
echo.
echo Checking/installing required packages (pdfplumber, openpyxl)...
%PY_CMD% -m pip install --quiet --disable-pip-version-check pdfplumber openpyxl

echo.
echo Converting all PDFs found in this folder...
echo.

%PY_CMD% gst_pdf_to_excel.py "%TOOL_DIR%"

echo.
echo ============================================
echo  Done. Look for the new .xlsx files next to
echo  each PDF in this same folder.
echo ============================================
pause
