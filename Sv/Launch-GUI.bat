@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -File "%~dp0Convert-to-CloudOnly-GUI.ps1"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Script avslutades med fel. Tryck en tangent for att stanga.
    pause > nul
)
