@echo off
REM Launcher do IT Admin Toolkit (versao WebView2 / Claude Design).
REM Forca PowerShell 5.1 x64 via SysNative se estivermos num host 32-bit.
set "SCRIPT=%~dp0ITAdminToolkit-Web.ps1"

REM SysNative so existe quando o processo pai e 32-bit. Se o processo for ja
REM 64-bit, usamos directamente System32 (que e 64-bit em SO 64-bit).
set "PS_SYSNATIVE=%SystemRoot%\SysNative\WindowsPowerShell\v1.0\powershell.exe"
set "PS_SYSTEM32=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if exist "%PS_SYSNATIVE%" (
    set "PS_EXE=%PS_SYSNATIVE%"
) else (
    set "PS_EXE=%PS_SYSTEM32%"
)

start "" "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File "%SCRIPT%"