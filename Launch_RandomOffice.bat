@echo off
title RandomOffice Desktop
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -Path '%~dp0dist\standalone\RandomOffice.exe' -ErrorAction SilentlyContinue"
if exist "dist\standalone\RandomOffice.exe" (
    start "" "dist\standalone\RandomOffice.exe"
    exit
)
start "" .venv\Scripts\pythonw.exe -m randomoffice desktop
exit

