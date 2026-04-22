@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -File "%~dp0Convert-to-CloudOnly-GUI.ps1"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Script exited with an error. Press any key to close.
    pause > nul
)
