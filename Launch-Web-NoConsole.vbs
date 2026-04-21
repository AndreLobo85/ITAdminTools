' Launcher silencioso - versao WebView2.
' Forca PowerShell 5.1 x64 via SysNative (evita o WOW64 redirect que
' lancaria a versao 32-bit e podia tornar as queries AD lentas/diferentes).
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
ps1 = scriptDir & "\ITAdminToolkit-Web.ps1"

' Caminhos candidatos em ordem de preferencia
winDir = sh.ExpandEnvironmentStrings("%SystemRoot%")
sysNative = winDir & "\SysNative\WindowsPowerShell\v1.0\powershell.exe"
system32 = winDir & "\System32\WindowsPowerShell\v1.0\powershell.exe"

If fso.FileExists(sysNative) Then
    psExe = sysNative
ElseIf fso.FileExists(system32) Then
    psExe = system32
Else
    psExe = "powershell.exe"
End If

cmd = """" & psExe & """ -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File """ & ps1 & """"
sh.Run cmd, 0, False