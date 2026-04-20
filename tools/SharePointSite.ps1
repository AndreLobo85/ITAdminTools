# ============================================================
# SharePointSite.ps1 - Auditar membros + permissoes de um site SPO
# via PnP.PowerShell. Usa Connect-PnPOnline -Interactive (MFA popup).
# Exporta: SP_Invoke-SiteAudit, SP_Format-Report
# ============================================================

function SP_Test-ModuleAvailable {
    try { return $null -ne (Get-Module -ListAvailable -Name PnP.PowerShell) } catch { return $false }
}

function SP_Ensure-Connected {
    param([string]$SiteUrl, [string]$TenantAdminUrl)

    # Verificar se ja existe conexao activa para este site
    try {
        $ctx = Get-PnPContext -ErrorAction SilentlyContinue
        if ($ctx -and $ctx.Url -like "*$([uri]::new($SiteUrl).Host)*") {
            return  # ja ligado
        }
    } catch { }

    Import-Module PnP.PowerShell -ErrorAction Stop
    $params = @{ Url = $SiteUrl; Interactive = $true; ErrorAction = 'Stop' }
    Connect-PnPOnline @params
}

function SP_Invoke-SiteAudit {
    param(
        [string]$SiteUrl,
        [string]$TenantAdminUrl = '',
        [bool]$IncludeOwners = $true
    )
    if (-not (SP_Test-ModuleAvailable)) {
        throw "Modulo PnP.PowerShell nao instalado. Corre: Install-Module PnP.PowerShell -Scope CurrentUser"
    }
    if (-not $SiteUrl) { throw 'Indique o URL do site.' }

    SP_Ensure-Connected -SiteUrl $SiteUrl -TenantAdminUrl $TenantAdminUrl

    $result = [ordered]@{
        SiteUrl       = $SiteUrl
        ConnectedAs   = ''
        SiteTitle     = ''
        SiteTemplate  = ''
        OwnerGroup    = $null
        MemberGroup   = $null
        VisitorGroup  = $null
        AllGroups     = @()
        M365GroupOwners  = @()
        M365GroupMembers = @()
        Error            = $null
    }

    try {
        $ctx = Get-PnPContext
        $result.ConnectedAs = try { (Get-PnPProperty -ClientObject $ctx.Web -Property CurrentUser).LoginName } catch { '' }

        # Site basics
        $web = Get-PnPWeb -Includes 'Title','WebTemplate','Configuration','AssociatedOwnerGroup','AssociatedMemberGroup','AssociatedVisitorGroup'
        $result.SiteTitle = "$($web.Title)"
        $result.SiteTemplate = "$($web.WebTemplate)#$($web.Configuration)"

        # Helper: resolve members dum SP group
        function _expand { param($g)
            if (-not $g) { return @() }
            $members = @()
            try {
                $members = Get-PnPGroupMember -Identity $g.Id -ErrorAction Stop |
                    ForEach-Object {
                        [PSCustomObject]@{
                            LoginName=$_.LoginName; Title=$_.Title; Email=$_.Email; PrincipalType="$($_.PrincipalType)"
                        }
                    }
            } catch { }
            @{ Id=$g.Id; Title=$g.Title; Members=$members }
        }

        $result.OwnerGroup   = _expand $web.AssociatedOwnerGroup
        $result.MemberGroup  = _expand $web.AssociatedMemberGroup
        $result.VisitorGroup = _expand $web.AssociatedVisitorGroup

        # Todos os grupos SP do site
        try {
            $result.AllGroups = Get-PnPGroup -ErrorAction Stop | ForEach-Object {
                $g = $_
                $mem = @()
                try {
                    $mem = Get-PnPGroupMember -Identity $g.Id -ErrorAction Stop |
                        ForEach-Object { "$($_.Title) <$($_.LoginName)>" }
                } catch { }
                [PSCustomObject]@{ Id=$g.Id; Title=$g.Title; MemberCount=$mem.Count; Members=$mem }
            }
        } catch { }

        # M365 group (se o site for group-connected)
        if ($IncludeOwners) {
            try {
                $site = Get-PnPSite -Includes 'GroupId' -ErrorAction SilentlyContinue
                if ($site -and $site.GroupId -and $site.GroupId -ne [guid]::Empty) {
                    try {
                        $owners = Get-PnPMicrosoft365GroupOwner -Identity $site.GroupId -ErrorAction Stop
                        $result.M365GroupOwners = $owners | ForEach-Object { "$($_.DisplayName) <$($_.UserPrincipalName)>" }
                    } catch { }
                    try {
                        $members = Get-PnPMicrosoft365GroupMember -Identity $site.GroupId -ErrorAction Stop
                        $result.M365GroupMembers = $members | ForEach-Object { "$($_.DisplayName) <$($_.UserPrincipalName)>" }
                    } catch { }
                }
            } catch { }
        }
    } catch {
        $result.Error = $_.Exception.Message
    }

    return [PSCustomObject]$result
}

function SP_Format-Report {
    param($Info)
    if (-not $Info) { return @('(sem info)') }
    $lines = @()
    $lines += "[OK] Ligado: $($Info.ConnectedAs)"
    $lines += "[INFO] Site: $($Info.SiteUrl)"
    $lines += "[INFO] Titulo: $($Info.SiteTitle)"
    $lines += "[INFO] Template: $($Info.SiteTemplate)"
    if ($Info.Error) { $lines += "[WARN] Erros parciais: $($Info.Error)" }
    $lines += ''

    function _emitGroup {
        param($g, $label)
        if (-not $g) { return }
        $lines += "[AUDIT] $label ($($g.Title)) - $($g.Members.Count) membro(s)"
        foreach ($m in $g.Members) {
            $lines += "  - $($m.Title) <$($m.LoginName)>  [$($m.PrincipalType)]"
        }
        $lines += ''
    }

    # emit via local wrapper (capturamos $lines por reference semantics do array in-place)
    if ($Info.OwnerGroup)   { $lines += "[AUDIT] Owner group ($($Info.OwnerGroup.Title)) - $($Info.OwnerGroup.Members.Count) membro(s)"; foreach ($m in $Info.OwnerGroup.Members) { $lines += "  - $($m.Title) <$($m.LoginName)>  [$($m.PrincipalType)]" }; $lines += '' }
    if ($Info.MemberGroup)  { $lines += "[AUDIT] Member group ($($Info.MemberGroup.Title)) - $($Info.MemberGroup.Members.Count) membro(s)"; foreach ($m in $Info.MemberGroup.Members) { $lines += "  - $($m.Title) <$($m.LoginName)>  [$($m.PrincipalType)]" }; $lines += '' }
    if ($Info.VisitorGroup) { $lines += "[AUDIT] Visitor group ($($Info.VisitorGroup.Title)) - $($Info.VisitorGroup.Members.Count) membro(s)"; foreach ($m in $Info.VisitorGroup.Members) { $lines += "  - $($m.Title) <$($m.LoginName)>  [$($m.PrincipalType)]" }; $lines += '' }

    if ($Info.AllGroups.Count -gt 0) {
        $lines += "[AUDIT] Todos os grupos SP ($($Info.AllGroups.Count)):"
        foreach ($g in $Info.AllGroups) {
            $lines += "  [$($g.MemberCount)]  $($g.Title)"
        }
        $lines += ''
    }

    if ($Info.M365GroupOwners.Count -gt 0) {
        $lines += "[AUDIT] M365 Group Owners ($($Info.M365GroupOwners.Count)):"
        foreach ($o in $Info.M365GroupOwners) { $lines += "  - $o" }
        $lines += ''
    }
    if ($Info.M365GroupMembers.Count -gt 0) {
        $lines += "[AUDIT] M365 Group Members ($($Info.M365GroupMembers.Count)):"
        foreach ($m in $Info.M365GroupMembers) { $lines += "  - $m" }
        $lines += ''
    }

    $lines += '[OK] Concluido.'
    return $lines
}
