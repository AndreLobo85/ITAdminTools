<#
.SYNOPSIS
    Audit de um site SharePoint Online via PnP.PowerShell.

.DESCRIPTION
    Spawn via Invoke-ScriptExternal -PreferPwsh. Reusa a sessao MSAL
    estabelecida pela janela pwsh.exe do botao "Ligar" (Connect-PnPOnline
    -Interactive). Silent auth via token cache apos primeiro Connect.

    Equivalente ao que se faria manualmente:
      Connect-PnPOnline -Url <site>
      Get-PnPWeb | ...
      Get-PnPGroup | ...
      Get-PnPMicrosoft365GroupOwner | ...
#>

param(
    [Parameter(Mandatory)][string]$siteUrl,
    [switch]$includeOwners
)

$ErrorActionPreference = 'Continue'

"[DIAG] Host: PS $($PSVersionTable.PSVersion)"
""

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    "[ERR] Modulo PnP.PowerShell nao instalado neste PowerShell."
    "      Install-Module PnP.PowerShell -Scope CurrentUser"
    return
}
Import-Module PnP.PowerShell -ErrorAction Stop

"[INFO] Connect-PnPOnline -Url '$siteUrl' -Interactive ..."
try {
    Connect-PnPOnline -Url $siteUrl -Interactive -ErrorAction Stop
} catch {
    "[ERR] Connect-PnPOnline falhou: $($_.Exception.Message)"
    "      Se o popup nao apareceu, abre uma consola pwsh e corre:"
    "      Connect-PnPOnline -Url '$siteUrl' -DeviceLogin"
    return
}

try {
    $ctx = Get-PnPContext
    $connectedAs = try { (Get-PnPProperty -ClientObject $ctx.Web -Property CurrentUser).LoginName } catch { '?' }
    "[OK] Ligado como: $connectedAs"
} catch {
    "[WARN] Nao consegui ler CurrentUser: $($_.Exception.Message)"
}

# Site basics
try {
    $web = Get-PnPWeb -Includes 'Title','WebTemplate','Configuration','AssociatedOwnerGroup','AssociatedMemberGroup','AssociatedVisitorGroup' -ErrorAction Stop
    "[INFO] Site URL      : $siteUrl"
    "[INFO] Site Title    : $($web.Title)"
    "[INFO] Site Template : $($web.WebTemplate)#$($web.Configuration)"
    ""
} catch {
    "[ERR] Get-PnPWeb falhou: $($_.Exception.Message)"
    return
}

# Helper local para listar membros de um grupo SP
function Emit-Group {
    param($g, [string]$label)
    if (-not $g) { return }
    $members = @()
    try {
        $members = Get-PnPGroupMember -Identity $g.Id -ErrorAction Stop
    } catch {
        "[WARN] Nao consegui ler membros de $label : $($_.Exception.Message)"
        return
    }
    "[AUDIT] $label ($($g.Title)) - $($members.Count) membro(s)"
    foreach ($m in $members) {
        "  - $($m.Title) <$($m.LoginName)>  [$($m.PrincipalType)]"
    }
    ""
}

Emit-Group $web.AssociatedOwnerGroup   'Owner group'
Emit-Group $web.AssociatedMemberGroup  'Member group'
Emit-Group $web.AssociatedVisitorGroup 'Visitor group'

# Todos os grupos SP
try {
    $allGroups = Get-PnPGroup -ErrorAction Stop
    if ($allGroups.Count -gt 0) {
        "[AUDIT] Todos os grupos SP ($($allGroups.Count)):"
        foreach ($g in $allGroups) {
            $memCount = 0
            try { $memCount = @(Get-PnPGroupMember -Identity $g.Id -ErrorAction Stop).Count } catch {}
            "  [$memCount]  $($g.Title)"
        }
        ""
    }
} catch {
    "[WARN] Get-PnPGroup falhou: $($_.Exception.Message)"
}

# M365 Group owners + members (se o site e group-connected)
if ($includeOwners) {
    try {
        $site = Get-PnPSite -Includes 'GroupId' -ErrorAction SilentlyContinue
        if ($site -and $site.GroupId -and $site.GroupId -ne [guid]::Empty) {
            "[AUDIT] M365 Group GroupId: $($site.GroupId)"
            try {
                $owners = Get-PnPMicrosoft365GroupOwner -Identity $site.GroupId -ErrorAction Stop
                "[AUDIT] M365 Group Owners ($($owners.Count)):"
                foreach ($o in $owners) { "  - $($o.DisplayName) <$($o.UserPrincipalName)>" }
                ""
            } catch { "[WARN] Get-PnPMicrosoft365GroupOwner: $($_.Exception.Message)" }
            try {
                $members = Get-PnPMicrosoft365GroupMember -Identity $site.GroupId -ErrorAction Stop
                "[AUDIT] M365 Group Members ($($members.Count)):"
                foreach ($m in $members) { "  - $($m.DisplayName) <$($m.UserPrincipalName)>" }
                ""
            } catch { "[WARN] Get-PnPMicrosoft365GroupMember: $($_.Exception.Message)" }
        } else {
            "[INFO] Site nao esta connected a M365 Group."
        }
    } catch {
        "[WARN] Get-PnPSite GroupId falhou: $($_.Exception.Message)"
    }
}

"[OK] Concluido."
