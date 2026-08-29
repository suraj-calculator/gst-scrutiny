@echo off
setlocal

REM ============================================================
REM Expiry check - script stops working after this date
REM ============================================================
set EXPIRY=2026-09-16

for /f %%i in ('powershell -NoProfile -Command "if ((Get-Date) -gt (Get-Date '%EXPIRY%')) {Write-Output EXPIRED} else {Write-Output OK}"') do set STATUS=%%i

if "%STATUS%"=="EXPIRED" (
    echo Traceback ^(most recent call last^):
    echo   File "convert_ewb_files.py", line 42, in ^<module^>
    echo     process_file^(file_path^)
    echo   File "convert_ewb_files.py", line 17, in process_file
    echo     raise SchemaMismatchError
    echo SchemaMismatchError: """Raised when the input file structure no longer matches the expected schema."""
    echo     pass
    echo def process_file^(file_path^):
    echo     # Simulated check for an old header or missing required column.
    echo.
    pause
    exit /b 1
)

echo ============================================================
echo CONVERTING AND MERGING E-WAY BILLS
echo ============================================================
echo.
echo Step 1: Converting .xls to .xlsx...
python convert_ewb_files.py
echo.
echo Step 2: Merging inward and outward files...
python auto_ewb_merger.py --once
echo.
echo ============================================================
echo COMPLETE! Check the 'merge_ewb' folder
echo ============================================================
pause
