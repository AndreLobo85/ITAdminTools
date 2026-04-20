<#
.SYNOPSIS
    Descarrega React, ReactDOM, Babel e fonts (Inter + JetBrains Mono)
    para webui/assets/, para que a app corra 100% offline.

.DESCRIPTION
    Corre uma unica vez apos Setup-WebView2.ps1. Se o PC do trabalho nao
    tiver internet, os bundles ficam em cache local e a aplicacao arranca
    sem tentar atingir unpkg.com ou fonts.googleapis.com.

    Apos correr, index.html e actualizado para referenciar os ficheiros
    locais.
#>

$ErrorActionPreference = 'Stop'
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$VendorDir  = Join-Path $ScriptDir 'webui\assets\vendor'
$FontsDir   = Join-Path $ScriptDir 'webui\assets\fonts'
$IndexHtml  = Join-Path $ScriptDir 'webui\index.html'

New-Item $VendorDir -ItemType Directory -Force | Out-Null
New-Item $FontsDir  -ItemType Directory -Force | Out-Null

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$script:ChromeUA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'

function Download-File {
    param([string]$Url, [string]$OutPath)
    Write-Host "  <- $Url" -ForegroundColor Yellow
    Invoke-WebRequest -Uri $Url -OutFile $OutPath -UseBasicParsing -Headers @{ 'User-Agent' = $script:ChromeUA }
}

Write-Host "================================================" -ForegroundColor Cyan
Write-Host " Offline assets - React + Babel + Fonts" -ForegroundColor Cyan
Write-Host "================================================`n" -ForegroundColor Cyan

# ========== React + ReactDOM + Babel (production) ==========
$libs = @(
    @{ Name='react.production.min.js';     Url='https://unpkg.com/react@18.3.1/umd/react.production.min.js' },
    @{ Name='react-dom.production.min.js'; Url='https://unpkg.com/react-dom@18.3.1/umd/react-dom.production.min.js' },
    @{ Name='babel.min.js';                Url='https://unpkg.com/@babel/standalone@7.29.0/babel.min.js' }
)
foreach ($lib in $libs) {
    $dst = Join-Path $VendorDir $lib.Name
    Download-File -Url $lib.Url -OutPath $dst
    Write-Host "  OK $($lib.Name)" -ForegroundColor Green
}

# ========== Google Fonts (Inter + JetBrains Mono) ==========
Write-Host ""
Write-Host "A descarregar CSS Google Fonts..." -ForegroundColor Yellow
$cssUrl = 'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap'
$cssRaw = Invoke-WebRequest -Uri $cssUrl -UseBasicParsing -Headers @{ 'User-Agent' = $script:ChromeUA } |
    Select-Object -ExpandProperty Content

# O Google Fonts devolve varios blocos @font-face, um por subset (latin, latin-ext, cyrillic, etc).
# So precisamos do subset latin. Filtramos pelas URLs .woff2 que aparecem no CSS.
$woffUrls = [regex]::Matches($cssRaw, "url\((https://fonts\.gstatic\.com/[^)]+\.woff2)\)") |
    ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique

Write-Host "Encontradas $($woffUrls.Count) fontes woff2. A descarregar..." -ForegroundColor Yellow
$urlToLocal = @{}
foreach ($u in $woffUrls) {
    $name = [IO.Path]::GetFileName($u)
    $dst  = Join-Path $FontsDir $name
    Download-File -Url $u -OutPath $dst
    Write-Host "  OK $name" -ForegroundColor Green
    $urlToLocal[$u] = "assets/fonts/$name"
}

# Reescrever o CSS para apontar aos caminhos locais
$localCss = $cssRaw
foreach ($k in $urlToLocal.Keys) { $localCss = $localCss.Replace($k, $urlToLocal[$k]) }
$cssPath = Join-Path $FontsDir 'fonts.css'
Set-Content -Path $cssPath -Value $localCss -Encoding UTF8
Write-Host "`n  OK fonts.css escrito" -ForegroundColor Green

# ========== Patch index.html para usar assets locais ==========
Write-Host "`nA patchar index.html..." -ForegroundColor Yellow
$html = Get-Content -Path $IndexHtml -Raw -Encoding UTF8

# Substituir links
$html = $html -replace '<link rel="preconnect" href="https://fonts\.googleapis\.com"/>', ''
$html = $html -replace '<link rel="preconnect" href="https://fonts\.gstatic\.com" crossorigin/>', ''
$html = $html -replace '<link href="https://fonts\.googleapis\.com/css2\?[^"]+" rel="stylesheet"/>', '<link href="assets/fonts/fonts.css" rel="stylesheet"/>'
$html = $html -replace 'https://unpkg\.com/react@18\.3\.1/umd/react\.development\.js', 'assets/vendor/react.production.min.js'
$html = $html -replace 'https://unpkg\.com/react-dom@18\.3\.1/umd/react-dom\.development\.js', 'assets/vendor/react-dom.production.min.js'
$html = $html -replace 'https://unpkg\.com/@babel/standalone@7\.29\.0/babel\.min\.js', 'assets/vendor/babel.min.js'
# Remover integrity/crossorigin deixados para tras
$html = $html -replace '\s*integrity="[^"]*"', ''
$html = $html -replace '\s*crossorigin="anonymous"', ''
$html = $html -replace '\s*crossorigin\s*(?=/>)', ''

Set-Content -Path $IndexHtml -Value $html -Encoding UTF8
Write-Host "  OK index.html actualizado" -ForegroundColor Green

Write-Host "`nSetup offline concluido. App agora corre sem internet." -ForegroundColor Cyan
