<#
.SYNOPSIS
    Consulta Get-MailboxStatistics para um ou mais UPNs via ExchangeOnlineManagement.
    Pre-valida que o utilizador tem role Exchange Admin (ou superior) ACTIVA em PIM
    via Microsoft Graph antes de fazer Connect-ExchangeOnline.

.DESCRIPTION
    Spawn via Invoke-ScriptExternal -PreferPwsh pelo host. Os modulos Microsoft.Graph
    e ExchangeOnlineManagement estao tipicamente instalados em PS 7 (PowerShell\Modules)
    e nao em PS 5.1 (WindowsPowerShell\Modules) quando o Install-Module foi feito em
    pwsh.exe. Correr via pwsh.exe garante que os modulos sao encontrados.
#>

param(
    [string]$adminUpn = "",
    [string]$upns = "",
    [switch]$includeRecov
)

$ErrorActionPreference = 'Continue'

if (-not $upns) {
    "[ERR] Indique pelo menos um UPN via -upns."
    return
}

$upnList = $upns -split '[\r\n,;]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique

"[DIAG] Host: PS $($PSVersionTable.PSVersion)  $(if ([IntPtr]::Size -eq 8) {'x64'} else {'x86'})"
"[DIAG] PSModulePath prefixes:"
foreach ($p in ($env:PSModulePath -split ';')) { if ($p) { "        $p" } }
""

# ---- 1. Validacao de role (Microsoft Graph) ----
"[INFO] A validar roles activas do administrador (Microsoft Graph)..."
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    "[ERR] Modulo Microsoft.Graph.Authentication nao instalado neste host PowerShell."
    "      Para instalar (na consola onde estas): Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
    "      Se ja instalaste numa versao diferente de PS, abre essa consola e corre novamente."
    return
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

$graphScopes = @('RoleManagement.Read.Directory','Directory.Read.All','User.Read')
$graphParams = @{ Scopes = $graphScopes; NoWelcome = $true; ErrorAction = 'Stop' }
try {
    Connect-MgGraph @graphParams | Out-Null
} catch {
    "[ERR] Connect-MgGraph falhou: $($_.Exception.Message)"
    return
}

# transitiveMemberOf/microsoft.graph.directoryRole devolve so roles ACTIVAS
# (PIM eligible nao activadas nao aparecem aqui)
try {
    $resp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.directoryRole' -ErrorAction Stop
    $activeRoles = @($resp.value | ForEach-Object { $_.displayName })
} catch {
    "[ERR] Falha a ler roles via Graph: $($_.Exception.Message)"
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    return
}

$allowedRoles = @(
    'Exchange Administrator',
    'Exchange Recipient Administrator',
    'Global Administrator',
    'Global Reader'
)
$matched = @($activeRoles | Where-Object { $_ -in $allowedRoles })

if ($activeRoles.Count -eq 0) {
    "[ERR] Nao tens nenhuma directory role ACTIVA neste momento."
    "      Activa em PIM: https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade"
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    return
}

"[INFO] Roles activas detectadas: $($activeRoles -join ', ')"

if ($matched.Count -eq 0) {
    "[ERR] Nenhuma role com acesso a Exchange Online esta ACTIVA."
    "      Roles aceites: $($allowedRoles -join ', ')"
    "      Activa Exchange Administrator em PIM antes de tentar de novo."
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    return
}

"[OK] Role permitida activa: $($matched -join ', ')"

# ---- 2. Connect Exchange Online ----
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    "[ERR] Modulo ExchangeOnlineManagement nao instalado."
    "      Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    return
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop

$alreadyConn = $null
try { $alreadyConn = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected' | Select-Object -First 1 } catch {}
if (-not $alreadyConn) {
    "[INFO] A ligar a Exchange Online (popup de autenticacao vai aparecer)..."
    $exoParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
    if ($adminUpn) { $exoParams.UserPrincipalName = $adminUpn }
    try {
        Connect-ExchangeOnline @exoParams
    } catch {
        "[ERR] Connect-ExchangeOnline falhou: $($_.Exception.Message)"
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
        return
    }
    "[OK] Ligado a Exchange Online"
} else {
    "[INFO] Sessao Exchange Online ja activa, a reutilizar."
}

# ---- 3. Iterar UPNs ----
""
"[INFO] Exchange Online - $($upnList.Count) mailbox(es)"
""
foreach ($u in $upnList) {
    try {
        $stats = Get-MailboxStatistics -Identity $u -ErrorAction Stop
        "[OK] $u"
        "  DisplayName         : $($stats.DisplayName)"
        "  StorageLimitStatus  : $($stats.StorageLimitStatus)"
        "  TotalItemSize       : $($stats.TotalItemSize)"
        "  TotalDeletedItemSize: $($stats.TotalDeletedItemSize)"
        "  ItemCount           : $($stats.ItemCount)"
        "  DeletedItemCount    : $($stats.DeletedItemCount)"
        "  LastLogonTime       : $($stats.LastLogonTime)"
        if ($includeRecov) {
            try {
                $rec = Get-MailboxStatistics -Identity $u -FolderScope RecoverableItems -ErrorAction Stop
                "  RecoverableItems    : $($rec.TotalItemSize)"
            } catch {
                "  RecoverableItems    : (erro: $($_.Exception.Message))"
            }
        }
    } catch {
        "[ERR] $u :: $($_.Exception.Message)"
    }
    ""
}