<#
.SYNOPSIS
    Consulta Get-MailboxStatistics para um ou mais UPNs via ExchangeOnlineManagement.

.DESCRIPTION
    Replica exactamente o fluxo manual do utilizador:
      1. Connect-ExchangeOnline  (abre popup Microsoft para seleccionar conta admin)
      2. Get-MailboxStatistics <UPN> | Select-Object userprincipalname,StorageLimitStatus,
         TotalItemSize,TotalDeletedItemSize,ItemCount,DeletedItemCount,RecoverableItems

    Correr via pwsh.exe (PS 7) onde o modulo ExchangeOnlineManagement esta instalado.
#>

param(
    [string]$upns = "",
    [switch]$includeRecov,
    [switch]$useDeviceCode
)

$ErrorActionPreference = 'Continue'

if (-not $upns) {
    "[ERR] Indique pelo menos um UPN via -upns."
    return
}

$upnList = $upns -split '[\r\n,;]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique

"[DIAG] Host: PS $($PSVersionTable.PSVersion)  $(if ([IntPtr]::Size -eq 8) {'x64'} else {'x86'})"
""

# ---- 1. Connect Exchange Online ----
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    "[ERR] Modulo ExchangeOnlineManagement nao instalado."
    "      Para instalar: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    return
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop

$alreadyConn = $null
try { $alreadyConn = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected' | Select-Object -First 1 } catch {}

if (-not $alreadyConn) {
    "[INFO] Connect-ExchangeOnline  (popup da Microsoft vai aparecer)..."
    $exoParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
    if ($useDeviceCode) {
        $exoParams['Device'] = $true
        "[INFO] Device code activo - URL + codigo vai aparecer aqui."
    }
    $exoConnected = $false
    try {
        Connect-ExchangeOnline @exoParams
        $exoConnected = $true
    } catch {
        $errMsg = $_.Exception.Message
        "[ERR] Connect-ExchangeOnline falhou: $errMsg"
        # Auto-fallback: se o popup foi fechado/perdido, tentar device code
        if (-not $useDeviceCode -and ($errMsg -match 'canceled|cancelled|timeout|not found|could not|closed|user closed')) {
            ""
            "[INFO] Auto-fallback para device code. Abre o URL que vai aparecer num browser e cola o codigo."
            try {
                $exoParams['Device'] = $true
                Connect-ExchangeOnline @exoParams
                $exoConnected = $true
            } catch {
                "[ERR] Device code tambem falhou: $($_.Exception.Message)"
            }
        }
    }
    if (-not $exoConnected) { return }
    "[OK] Ligado a Exchange Online"
} else {
    "[INFO] Sessao Exchange Online ja activa, a reutilizar."
}

""
"[INFO] A consultar $($upnList.Count) mailbox(es)..."
""

# ---- 2. Get-MailboxStatistics por UPN ----
# Equivalente a:
#   Get-MailboxStatistics <upn> | Select-Object userprincipalname,StorageLimitStatus,
#     TotalItemSize,TotalDeletedItemSize,ItemCount,DeletedItemCount,RecoverableItems
foreach ($u in $upnList) {
    try {
        $stats = Get-MailboxStatistics -Identity $u -ErrorAction Stop
        "[OK] $u"
        "  DisplayName          : $($stats.DisplayName)"
        "  UserPrincipalName    : $u"
        "  StorageLimitStatus   : $($stats.StorageLimitStatus)"
        "  TotalItemSize        : $($stats.TotalItemSize)"
        "  TotalDeletedItemSize : $($stats.TotalDeletedItemSize)"
        "  ItemCount            : $($stats.ItemCount)"
        "  DeletedItemCount     : $($stats.DeletedItemCount)"
        "  LastLogonTime        : $($stats.LastLogonTime)"
        if ($includeRecov) {
            try {
                $rec = Get-MailboxStatistics -Identity $u -FolderScope RecoverableItems -ErrorAction Stop
                "  RecoverableItems     : $($rec.TotalItemSize)"
            } catch {
                "  RecoverableItems     : (erro: $($_.Exception.Message))"
            }
        }
    } catch {
        "[ERR] $u :: $($_.Exception.Message)"
    }
    ""
}