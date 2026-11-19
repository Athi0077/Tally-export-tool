@echo off
echo ====================================================================
echo Building Standalone Single .exe for X2T - Excel to Tally Converter
echo ====================================================================
echo.

python -m PyInstaller --noconfirm --onefile --windowed ^
    --add-data "x2t_desktop\templates;templates" ^
    --add-data "sample_data;sample_data" ^
    --collect-all customtkinter ^
    --name "X2T_Converter" ^
    run.py

echo.
if %ERRORLEVEL% EQU 0 (
    echo ====================================================================
    echo BUILD SUCCESSFUL!
    echo Standalone executable is located at: dist\X2T_Converter.exe
    echo ====================================================================
) else (
    echo ====================================================================
    echo BUILD FAILED. Check errors above.
    echo ====================================================================
)
pause
