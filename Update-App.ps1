<#
.SYNOPSIS
    Actualiza a instalacao existente do IT Admin Toolkit. Equivalente a re-correr
    o Install-ITAdminToolkit.ps1 no mesmo InstallPath.

.DESCRIPTION
    - Le a versao local em ./version.json
    - Consulta a ultima release no GitHub
    - Se diferente, descarrega e substitui a instalacao
    - Relança a app
#>

param(
    [string]$Repo = 'AndreLobo85/ITAdminTools'
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
# Se este script estiver dentro de ITAdminToolkit/, instala-se no seu parent
$installRoot = Split-Path -Parent $ScriptDir

# Chama o installer que ja sabe fazer tudo (idempotente)
$installerUrl = "https://raw.githubusercontent.com/$Repo/main/Install-ITAdminToolkit.ps1"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$bootstrap = (Invoke-WebRequest -Uri $installerUrl -UseBasicParsing).Content

# Executa com InstallPath actual
$sb = [ScriptBlock]::Create($bootstrap)
& $sb -Repo $Repo -InstallPath $installRoot
