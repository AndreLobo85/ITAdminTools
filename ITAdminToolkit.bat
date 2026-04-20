@echo off
REM Lanca o IT Admin Toolkit.
set "SCRIPT=%~dp0ITAdminToolkit.ps1"
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File "%SCRIPT%"
