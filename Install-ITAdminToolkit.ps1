<#
.SYNOPSIS
    Bootstrap de instalacao do IT Admin Toolkit. Correr no PC do trabalho.

.DESCRIPTION
    1. Verifica / instala WebView2 Runtime
    2. Descarrega o zip da ultima release do GitHub
    3. Extrai para %LOCALAPPDATA%\ITAdminToolkit (ou -InstallPath)
    4. Cria atalho no ambiente de trabalho
    5. Lanca a aplicacao

    One-liner (copiar para PowerShell na maquina alvo):
      iex (irm https://raw.githubusercontent.com/<USER>/<REPO>/main/Install-ITAdminToolkit.ps1)

    Re-correr o one-liner actualiza a instalacao para a versao mais recente.
#>

[CmdletBinding()]
param(
    [string]$Repo = 'AndreLobo85/ITAdminTools',
    [string]$InstallPath = (Join-Path $env:LOCALAPPDATA 'ITAdminToolkit'),
    [switch]$NoShortcut,
    [switch]$NoLaunch
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Host "================================================" -ForegroundColor Cyan
Write-Host " IT Admin Toolkit - Instalador" -ForegroundColor Cyan
Write-Host " Repo: $Repo" -ForegroundColor Cyan
Write-Host " Destino: $InstallPath" -ForegroundColor Cyan
Write-Host "================================================`n" -ForegroundColor Cyan

# ===== 1. WebView2 Runtime =====
function Test-WebView2Runtime {
    foreach ($p in @(
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}',
        'HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}',
        'HKCU:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}'
    )) {
        try { $v = (Get-ItemProperty -Path $p -Name pv -ErrorAction Stop).pv; if ($v) { return $v } } catch {}
    }
    return $null
}

$wvVer = Test-WebView2Runtime
if ($wvVer) {
    Write-Host "[OK] WebView2 Runtime encontrado: $wvVer" -ForegroundColor Green
} else {
    Write-Host "[!] WebView2 Runtime nao encontrado. A descarregar e instalar..." -ForegroundColor Yellow
    $bootUrl = 'https://go.microsoft.com/fwlink/p/?LinkId=2124703'
    $bootExe = Join-Path $env:TEMP 'MicrosoftEdgeWebView2Setup.exe'
    Invoke-WebRequest -Uri $bootUrl -OutFile $bootExe -UseBasicParsing
    Start-Process -FilePath $bootExe -ArgumentList '/silent /install' -Wait
    Remove-Item $bootExe -ErrorAction SilentlyContinue
    $wvVer = Test-WebView2Runtime
    if (-not $wvVer) { throw "Falhou a instalar WebView2 Runtime. Instala manualmente em https://go.microsoft.com/fwlink/p/?LinkId=2124703" }
    Write-Host "[OK] WebView2 Runtime instalado: $wvVer" -ForegroundColor Green
}

# ===== 2. Verificar versao remota =====
Write-Host "`n[1/3] A verificar ultima release..." -ForegroundColor Yellow
$apiUrl = "https://api.github.com/repos/$Repo/releases/latest"
try {
    $release = Invoke-RestMethod -Uri $apiUrl -UseBasicParsing -Headers @{ 'User-Agent' = 'ITAdminToolkit-Installer' }
} catch {
    throw "Nao foi possivel contactar GitHub ($apiUrl). Verifica o repo ou se ha uma release publicada.`n$($_.Exception.Message)"
}
$tag = $release.tag_name
Write-Host "    Release: $tag ($($release.name))" -ForegroundColor Green

# Encontrar asset zip
$asset = $release.assets | Where-Object { $_.name -like 'ITAdminToolkit-*.zip' } | Select-Object -First 1
if (-not $asset) { throw "Release $tag nao tem zip 'ITAdminToolkit-*.zip'." }
$downloadUrl = $asset.browser_download_url

# Check se ja temos esta versao instalada
$installedVer = $null
$verFile = Join-Path $InstallPath 'version.json'
if (Test-Path $verFile) {
    try { $installedVer = "v$((Get-Content $verFile -Raw | ConvertFrom-Json).version)" } catch {}
}
if ($installedVer -eq $tag) {
    Write-Host "`n[OK] Ja tens a versao mais recente ($tag). Nada para fazer." -ForegroundColor Green
    if (-not $NoLaunch) { Start-Process (Join-Path $InstallPath 'ITAdminToolkit\Launch-Web-NoConsole.vbs') }
    return
}

# ===== 3. Download =====
Write-Host "`n[2/3] A descarregar $($asset.name)..." -ForegroundColor Yellow
$zipPath = Join-Path $env:TEMP $asset.name
Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -UseBasicParsing
Write-Host "    Descarregado: $([Math]::Round((Get-Item $zipPath).Length / 1MB, 2)) MB" -ForegroundColor Green

# ===== 4. Extrair =====
Write-Host "`n[3/3] A extrair..." -ForegroundColor Yellow
if (Test-Path $InstallPath) {
    # Preservar logs do user e config anteriores, substituir app
    $oldAppDir = Join-Path $InstallPath 'ITAdminToolkit'
    if (Test-Path $oldAppDir) { Remove-Item $oldAppDir -Recurse -Force }
} else {
    New-Item $InstallPath -ItemType Directory -Force | Out-Null
}
Expand-Archive -Path $zipPath -DestinationPath $InstallPath -Force
Remove-Item $zipPath -ErrorAction SilentlyContinue

$appRoot = Join-Path $InstallPath 'ITAdminToolkit'
$launcher = Join-Path $appRoot 'Launch-Web-NoConsole.vbs'
if (-not (Test-Path $launcher)) { throw "Launcher nao encontrado apos extraccao: $launcher" }

Write-Host "    Instalado em: $appRoot" -ForegroundColor Green

# ===== 5. Shortcut =====
if (-not $NoShortcut) {
    try {
        $desktop = [Environment]::GetFolderPath('Desktop')
        $lnkPath = Join-Path $desktop 'IT Admin Toolkit.lnk'
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($lnkPath)
        $shortcut.TargetPath = $launcher
        $shortcut.WorkingDirectory = $appRoot
        $shortcut.IconLocation = (Join-Path $appRoot 'webui\assets\nb-mark.svg')
        $shortcut.Description = 'IT Admin Toolkit - novobanco'
        $shortcut.Save()
        Write-Host "    Atalho criado no ambiente de trabalho" -ForegroundColor Green
    } catch {
        Write-Host "    (atalho nao criado: $($_.Exception.Message))" -ForegroundColor DarkYellow
    }
}

# ===== Done =====
Write-Host "`n================================================" -ForegroundColor Cyan
Write-Host " Instalacao concluida: $tag" -ForegroundColor Cyan
Write-Host " Lanca com: $launcher" -ForegroundColor Cyan
Write-Host "================================================`n" -ForegroundColor Cyan

if (-not $NoLaunch) {
    Start-Process $launcher
}
