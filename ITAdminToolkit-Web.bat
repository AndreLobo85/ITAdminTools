@echo off
REM Launcher do IT Admin Toolkit (versao WebView2 / Claude Design).
set "SCRIPT=%~dp0ITAdminToolkit-Web.ps1"
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File "%SCRIPT%"
