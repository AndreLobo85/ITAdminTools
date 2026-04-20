<#
.SYNOPSIS
    Script de BUILD (para ti, dev). Gera um .zip com TUDO incluido
    (DLLs WebView2 + React + fonts + codigo) pronto para subir como
    GitHub Release asset.

.DESCRIPTION
    Produz: release/ITAdminToolkit-v<version>.zip

    Passos:
     1. Garante que lib/ e webui/assets/vendor, /fonts estao povoadas
        (corre Setup-WebView2.ps1 e Setup-OfflineAssets.ps1 se faltarem)
     2. Le version.json para determinar a tag
     3. Zipa tudo (excepto .git, release/, logs)
     4. Output: release/ITAdminToolkit-v<version>.zip

    Fluxo completo:
        .\Build-Release.ps1
        git tag v1.0.0
        git push origin v1.0.0
        gh release create v1.0.0 release\ITAdminToolkit-v1.0.0.zip \
            --title 'v1.0.0' --notes 'Initial release'
#>

param([switch]$SkipSetup)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ReleaseDir = Join-Path $ScriptDir 'release'
New-Item $ReleaseDir -ItemType Directory -Force | Out-Null

# Ler versao
$verJson = Get-Content (Join-Path $ScriptDir 'version.json') -Raw | ConvertFrom-Json
$version = $verJson.version
Write-Host "================================================" -ForegroundColor Cyan
Write-Host " Build release ITAdminToolkit v$version" -ForegroundColor Cyan
Write-Host "================================================`n" -ForegroundColor Cyan

# Garantir que temos DLLs e assets
if (-not $SkipSetup) {
    $libDll = Join-Path $ScriptDir 'lib\Microsoft.Web.WebView2.WinForms.dll'
    if (-not (Test-Path $libDll)) {
        Write-Host "A correr Setup-WebView2..." -ForegroundColor Yellow
        & (Join-Path $ScriptDir 'Setup-WebView2.ps1')
    }
    $reactJs = Join-Path $ScriptDir 'webui\assets\vendor\react.production.min.js'
    if (-not (Test-Path $reactJs)) {
        Write-Host "A correr Setup-OfflineAssets..." -ForegroundColor Yellow
        & (Join-Path $ScriptDir 'Setup-OfflineAssets.ps1')
    }
}

# Estruturar stage
$stageDir = Join-Path $env:TEMP "ITAdminToolkit-build-$((Get-Date).Ticks)"
New-Item $stageDir -ItemType Directory -Force | Out-Null
$appDir = Join-Path $stageDir 'ITAdminToolkit'
New-Item $appDir -ItemType Directory -Force | Out-Null

# Ficheiros / pastas a incluir
$includes = @(
    'ITAdminToolkit-Web.ps1',
    'ITAdminToolkit-Web.bat',
    'Launch-Web-NoConsole.vbs',
    'Setup-WebView2.ps1',
    'Setup-OfflineAssets.ps1',
    'Install-ITAdminToolkit.ps1',
    'Update-App.ps1',
    'version.json',
    'README.md',
    'tools',
    'webui',
    'lib'
)
foreach ($item in $includes) {
    $src = Join-Path $ScriptDir $item
    if (Test-Path $src) {
        Copy-Item $src -Destination $appDir -Recurse -Force
        Write-Host "  + $item" -ForegroundColor Green
    } else {
        Write-Host "  ! em falta (skip): $item" -ForegroundColor DarkYellow
    }
}

# Zip
$zipName = "ITAdminToolkit-v$version.zip"
$zipPath = Join-Path $ReleaseDir $zipName
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path "$appDir" -DestinationPath $zipPath -CompressionLevel Optimal

Remove-Item $stageDir -Recurse -Force -ErrorAction SilentlyContinue

$size = [Math]::Round((Get-Item $zipPath).Length / 1MB, 2)
Write-Host "`nDone: $zipPath ($size MB)`n" -ForegroundColor Cyan
Write-Host "Passos seguintes:" -ForegroundColor Yellow
Write-Host "  git add -A && git commit -m `"release v$version`"" -ForegroundColor White
Write-Host "  git tag v$version && git push --tags" -ForegroundColor White
Write-Host "  gh release create v$version `"$zipPath`" --title `"v$version`" --notes-from-tag" -ForegroundColor White
