<#
.SYNOPSIS
    Script de setup: descarrega as DLLs Microsoft.Web.WebView2 do NuGet
    e extrai-as para ./lib/ para que o host ITAdminToolkit-Web.ps1 as use.

.DESCRIPTION
    Corre uma unica vez antes do primeiro arranque. Nao precisa de admin.
    Sem internet nao funciona (esta e a unica operacao que precisa de net).
#>

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$LibDir = Join-Path $ScriptDir 'lib'
$TmpDir = Join-Path $env:TEMP "wv2-setup-$((Get-Date).Ticks)"

# Versao do pacote. Actualiza aqui se quiseres outra.
$Version = '1.0.2957.106'
$PackageUrl = "https://www.nuget.org/api/v2/package/Microsoft.Web.WebView2/$Version"

Write-Host "================================================" -ForegroundColor Cyan
Write-Host " WebView2 DLLs setup" -ForegroundColor Cyan
Write-Host " Versao: $Version" -ForegroundColor Cyan
Write-Host " Destino: $LibDir" -ForegroundColor Cyan
Write-Host "================================================`n" -ForegroundColor Cyan

New-Item -Path $LibDir -ItemType Directory -Force | Out-Null
New-Item -Path $TmpDir -ItemType Directory -Force | Out-Null

try {
    $nupkgPath = Join-Path $TmpDir "wv2.zip"
    Write-Host "A descarregar do NuGet..." -ForegroundColor Yellow
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $PackageUrl -OutFile $nupkgPath -UseBasicParsing

    Write-Host "A extrair..." -ForegroundColor Yellow
    $extractDir = Join-Path $TmpDir "extracted"
    Expand-Archive -Path $nupkgPath -DestinationPath $extractDir -Force

    # Copiar DLLs managed (net462 funciona em PowerShell 5.1)
    $managed = @(
        'lib\net462\Microsoft.Web.WebView2.Core.dll',
        'lib\net462\Microsoft.Web.WebView2.WinForms.dll',
        'lib\net462\Microsoft.Web.WebView2.Wpf.dll'
    )
    foreach ($rel in $managed) {
        $src = Join-Path $extractDir $rel
        if (Test-Path $src) {
            Copy-Item -Path $src -Destination $LibDir -Force
            Write-Host "  + $(Split-Path $rel -Leaf)" -ForegroundColor Green
        } else {
            Write-Host "  ! em falta: $rel" -ForegroundColor Yellow
        }
    }

    # WebView2Loader nativo para cada arquitectura. PowerShell carrega o certo.
    foreach ($arch in @('x64', 'x86', 'arm64')) {
        $archDir = Join-Path $LibDir "runtimes\win-$arch\native"
        New-Item -Path $archDir -ItemType Directory -Force | Out-Null
        $src = Join-Path $extractDir "runtimes\win-$arch\native\WebView2Loader.dll"
        if (Test-Path $src) {
            Copy-Item -Path $src -Destination $archDir -Force
            Write-Host "  + WebView2Loader.dll ($arch)" -ForegroundColor Green
        }
    }

    # Tambem copiar Loader para raiz do lib/ para fallback (arquitectura actual)
    $nativeArch = if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' }
    $loaderSrc = Join-Path $LibDir "runtimes\win-$nativeArch\native\WebView2Loader.dll"
    if (Test-Path $loaderSrc) {
        Copy-Item -Path $loaderSrc -Destination $LibDir -Force
        Write-Host "  + WebView2Loader.dll na raiz (arquitectura actual: $nativeArch)" -ForegroundColor Green
    }

    Write-Host "`nSetup concluido. Agora podes lancar ITAdminToolkit-Web.bat" -ForegroundColor Cyan
}
catch {
    Write-Host "`nERRO: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    if (Test-Path $TmpDir) { Remove-Item $TmpDir -Recurse -Force -ErrorAction SilentlyContinue }
}
