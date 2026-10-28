@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_exe.ps1"
if errorlevel 1 (
    echo.
    echo Build failed. Check the error messages above.
) else (
    echo.
    echo Build finished: dist\OracleBatchQuery.exe
)
pause
