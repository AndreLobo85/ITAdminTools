# ============================================================
# UserInfo.ps1 - Ferramenta: diagnostico de user AD (sem RSAT)
# Baseado em get-diag-info-aw2.ps1 (@Vitor Rodrigues / @Higino Antunes)
# Exporta: New-UserInfoTab
# ============================================================

# ----- Constantes userAccountControl -----
$script:UAC_ACCOUNTDISABLE       = 0x000002
$script:UAC_DONT_EXPIRE_PASSWORD = 0x010000
$script:UAC_PASSWORD_EXPIRED     = 0x800000
$script:UAC_LOCKOUT              = 0x000010

function UI_Safe-FromFileTime {
    param($Value)
    try {
        if (-not $Value) { return '' }
        $ft = [int64]$Value
        if ($ft -eq 0 -or $ft -eq [int64]::MaxValue) { return 'Nao expira' }
        return [datetime]::FromFileTime($ft).ToString('yyyy-MM-dd HH:mm')
    } catch { return '' }
}

function UI_Get-UserInfo {
    param(
        [string]$Username = '',
        [string]$Email = '',
        [ScriptBlock]$PumpUI = $null
    )

    if (-not $Username -and -not $Email) {
        throw 'Indique username ou email.'
    }

    # Resolver dominio + DC sem chamar GetCurrentDomain() / DomainControllers[0]
    # (essas APIs enumeram internamente TODOS os DCs do dominio e podem pendurar
    # varios segundos em dominios grandes). Usar env vars e instantaneo.
    $dnsDomain = $env:USERDNSDOMAIN
    if (-not $dnsDomain) {
        # Fallback: deriva do AD quando nao estamos num dominio logonservers
        $dnsDomain = ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()).Name
    }
    $server = $env:LOGONSERVER
    if ($server) { $server = $server -replace '^\\\\' }
    if (-not $server) {
        # Fallback: qualquer DC do dominio via DNS locator
        $server = $dnsDomain
    }
    # Distinguished name: novobanco.local -> DC=novobanco,DC=local
    $baseDN = 'DC=' + ($dnsDomain -replace '\.', ',DC=')
    # Precisamos de $domainObj para compatibilidade com codigo abaixo
    $domainObj = [PSCustomObject]@{ distinguishedName = $baseDN }

    $searcher = New-Object System.DirectoryServices.DirectorySearcher
    $searcher.PageSize = 1
    $searcher.SearchScope = 'subtree'
    $searcher.ClientTimeout = [TimeSpan]::FromSeconds(30)
    $searcher.ServerTimeLimit = [TimeSpan]::FromSeconds(30)
    $searcher.Filter = if ($Email) {
        "(|(mail=$Email)(proxyAddresses=smtp:$Email))"
    } else {
        "(samaccountname=$Username)"
    }

    @('name','SamAccountName','DisplayName','Description','mail','o',
      'lastLogonTimestamp','extensionAttribute13','extensionAttribute14','MemberOf',
      'pwdLastSet','msDS-UserPasswordExpiryTimeComputed','DistinguishedName',
      'Department','Manager','Title','Organization','proxyAddresses','homeMDB',
      'UserPrincipalName','whenCreated','AccountExpires','mailNickName',
      'msExchRemoteRecipientType','userAccountControl','badPwdCount') | ForEach-Object {
        [void]$searcher.PropertiesToLoad.Add($_)
    }

    $searcher.SearchRoot = "LDAP://$server/$baseDN"
    $results = $searcher.FindAll()
    if ($results.Count -eq 0) { return $null }
    $result = $results[0]
    # Mimic the $DC variable used downstream
    $DC = [PSCustomObject]@{ Name = $server }

    # LockedOut + badPwdCount - apenas do DC actual (iterar todos os DCs em dominios
    # grandes era inviavel, causava timeouts). Usamos as propriedades ja carregadas
    # no resultado da pesquisa, evitando qualquer 2a chamada LDAP.
    $lockedOut = @(); $badPwdCount = @(); $dcNames = @()
    try {
        $uacVal = 0
        if ($result.Properties['userAccountControl'] -and $result.Properties['userAccountControl'].Count -gt 0) {
            $uacVal = [int]$result.Properties['userAccountControl'][0]
        }
        $bpw = 0
        if ($result.Properties['badPwdCount'] -and $result.Properties['badPwdCount'].Count -gt 0) {
            $bpw = [int]$result.Properties['badPwdCount'][0]
        }
        $lockedOut   += [bool]($uacVal -band $script:UAC_LOCKOUT)
        $badPwdCount += $bpw
        $dcNames     += $DC.Name
    } catch { }

    # Manager DisplayName
    $managerName = ''
    try {
        $mgrDN = $result.Properties['manager'] | Select-Object -First 1
        if ($mgrDN) {
            $s2 = New-Object System.DirectoryServices.DirectorySearcher
            $s2.Filter = "(&(objectCategory=person)(objectClass=user)(distinguishedName=$mgrDN))"
            $s2.SearchRoot = "LDAP://$server/$($domainObj.distinguishedName)"
            $r2 = $s2.FindAll()
            if ($r2.Count -gt 0) {
                $managerName = "$(($r2[0].Properties['DisplayName'] | Select-Object -First 1))"
            }
        }
    } catch { }

    $proxyAddresses = $result.Properties.Item('proxyAddresses')
    $pa = @($proxyAddresses -match 'smtp:[\w@\.]+')

    $grupos = (($result.Properties.memberof | Sort-Object | ForEach-Object {
        $_ -replace 'CN=([^,]+),.+', '$1'
    }) -join '; ')

    $goffice = if ($result.Properties.Item('memberof') -like '*GO365PRO*') {
        $x1 = @($result.Properties.Item('memberof') -like '*GO365PRO*')[0].Split(',')
        ($x1[0].Split('='))[1]
    } else { 'Nao tem licenca!' }

    # Usar directamente do search result (evita 2a chamada LDAP via [adsi] bind)
    $uac = 0
    if ($result.Properties['userAccountControl'] -and $result.Properties['userAccountControl'].Count -gt 0) {
        $uac = [int]$result.Properties['userAccountControl'][0]
    }

    $whenCreated = ''
    try {
        $wc = $result.Properties.Item('whenCreated') | Select-Object -First 1
        if ($wc) { $whenCreated = ([datetime]$wc).ToLocalTime().ToString('yyyy-MM-dd HH:mm') }
    } catch { }

    $accountExpires = $result.Properties.Item('AccountExpires') | Select-Object -First 1
    $accountExpiry = if ($null -eq $accountExpires -or $accountExpires -eq 0 -or $accountExpires -eq [int64]::MaxValue) {
        'Nao expira'
    } else {
        UI_Safe-FromFileTime $accountExpires
    }

    $pwdExpiry = if ([bool]($uac -band $script:UAC_DONT_EXPIRE_PASSWORD)) {
        'Nao expira'
    } else {
        UI_Safe-FromFileTime ($result.Properties.Item('msDS-UserPasswordExpiryTimeComputed') | Select-Object -First 1)
    }

    $lastLogon     = UI_Safe-FromFileTime ($result.Properties.Item('lastLogonTimestamp') | Select-Object -First 1)
    $pwdLastSet    = UI_Safe-FromFileTime ($result.Properties.Item('pwdLastSet') | Select-Object -First 1)

    return [PSCustomObject]@{
        Name               = "$($result.Properties.Item('name'))"
        SamAccountName     = "$($result.Properties.Item('SamAccountName'))"
        Alias              = "$($result.Properties.Item('mailNickName'))"
        Description        = "$($result.Properties.Item('Description'))"
        DisplayName        = "$($result.Properties.Item('DisplayName'))"
        DistinguishedName  = "$($result.Properties.Item('DistinguishedName'))"
        Department         = "$($result.Properties.Item('Department'))"
        Title              = "$($result.Properties.Item('Title'))"
        ManagerDN          = "$($result.Properties.Item('Manager'))"
        ManagerName        = $managerName
        Outsourcer         = "$($result.Properties.Item('Organization'))"
        Organization       = "$($result.Properties.Item('o'))"
        Mail               = "$($result.Properties.Item('mail'))"
        AlternativeEmails  = ($pa -join ', ')
        Database           = ($result.Properties['homeMDB'] -replace 'CN=(\w+),.+', '$1')
        LastLogon          = $lastLogon
        Enabled            = -not [bool]($uac -band $script:UAC_ACCOUNTDISABLE)
        PasswordExpired    = [bool]($uac -band $script:UAC_PASSWORD_EXPIRED)
        PasswordLastSet    = $pwdLastSet
        PasswordExpiry     = $pwdExpiry
        LockedOutPerDC     = ($lockedOut -join ', ')
        BadPwdCountPerDC   = ($badPwdCount -join ', ')
        DCsQueried         = ($dcNames -join ', ')
        WhenCreated        = $whenCreated
        AccountExpiry      = $accountExpiry
        MsExchRemoteRecipientType = "$($result.Properties.Item('msExchRemoteRecipientType'))"
        ExtensionAttr13    = "$($result.Properties.Item('extensionAttribute13'))"
        ExtensionAttr14    = "$($result.Properties.Item('extensionAttribute14'))"
        GrupoSSPR          = (($result.Properties.Item('memberof') -like '*GO365SSPRNR*').Count -eq 1)
        GrupoCondAccessL   = (($result.Properties.Item('memberof') -like '*GO365CALNR*').Count -eq 1)
        GrupoLicOffice     = $goffice
        UserPrincipalName  = "$($result.Properties.Item('userPrincipalName'))"
        AcedeAoPAM         = if ($grupos -match 'GPAMNR') { 'Sim' } else { 'Nao' }
        Grupos             = $grupos
    }
}

function UI_Format-Report {
    param($Info)
    if (-not $Info) { return '(user nao encontrado)' }
    $sep = ('-' * 60)
    $lines = @()
    $lines += $sep
    $lines += 'Name                 : ' + $Info.Name
    $lines += 'SamAccountName       : ' + $Info.SamAccountName
    $lines += 'Alias/mailNickName   : ' + $Info.Alias
    $lines += 'Description          : ' + $Info.Description
    $lines += 'DisplayName          : ' + $Info.DisplayName
    $lines += 'DistinguishedName    : ' + $Info.DistinguishedName
    $lines += 'Estrutura            : ' + $Info.Department
    $lines += 'Funcao               : ' + $Info.Title
    $lines += 'Manager DN           : ' + $Info.ManagerDN
    $lines += 'Manager Name         : ' + $Info.ManagerName
    $lines += 'Outsourcer           : ' + $Info.Outsourcer
    $lines += 'Organization         : ' + $Info.Organization
    $lines += 'mail                 : ' + $Info.Mail
    $lines += 'alternativos         : ' + $Info.AlternativeEmails
    $lines += 'Database             : ' + $Info.Database
    $lines += $sep
    $lines += 'Ultimo Logon         : ' + $Info.LastLogon
    $lines += 'enabled              : ' + $Info.Enabled
    $lines += 'PasswordExpired      : ' + $Info.PasswordExpired
    $lines += 'PasswordLastSet      : ' + $Info.PasswordLastSet
    $lines += 'Password expira em   : ' + $Info.PasswordExpiry
    $lines += 'LockedOut (por DC)   : ' + $Info.LockedOutPerDC
    $lines += 'badPwdCount (por DC) : ' + $Info.BadPwdCountPerDC
    $lines += 'DCs consultados      : ' + $Info.DCsQueried
    $lines += 'Criado em            : ' + $Info.WhenCreated
    $lines += 'User expira em       : ' + $Info.AccountExpiry
    $lines += $sep
    $lines += 'msExchRemoteRecipientType : ' + $Info.MsExchRemoteRecipientType
    $lines += 'extensionAttribute13      : ' + $Info.ExtensionAttr13
    $lines += 'extensionAttribute14      : ' + $Info.ExtensionAttr14
    $lines += 'Grupo SSPR (password)     : ' + $Info.GrupoSSPR
    $lines += 'Grupo Cond Access L       : ' + $Info.GrupoCondAccessL
    $lines += 'Grupo Lic Office          : ' + $Info.GrupoLicOffice
    $lines += 'UserPrincipalName         : ' + $Info.UserPrincipalName
    $lines += $sep
    $lines += 'Acede ao PAM?        : ' + $Info.AcedeAoPAM
    $lines += 'Grupos               : ' + $Info.Grupos
    $lines += $sep
    return ($lines -join "`r`n")
}

function New-UserInfoTab {
    $script:UI_LastInfo = $null

    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = 'User Info'
    $tab.BackColor = [System.Drawing.Color]::White

    $panelTop = New-Object System.Windows.Forms.Panel
    $panelTop.Dock = 'Top'; $panelTop.Height = 130
    $panelTop.Padding = '12,12,12,12'
    $panelTop.BackColor = [System.Drawing.Color]::White
    $tab.Controls.Add($panelTop)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = 'Diagnostico de User AD'
    $lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $script:PaletteBlue.Dark
    $lblTitle.Location = New-Object System.Drawing.Point(12, 4)
    $lblTitle.Size = New-Object System.Drawing.Size(600, 26)
    $panelTop.Controls.Add($lblTitle)

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = 'Username (SamAccountName):'
    $lblUser.Location = New-Object System.Drawing.Point(12, 40); $lblUser.Size = New-Object System.Drawing.Size(210, 22)
    $panelTop.Controls.Add($lblUser)
    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location = New-Object System.Drawing.Point(228, 38); $txtUser.Size = New-Object System.Drawing.Size(220, 24)
    $panelTop.Controls.Add($txtUser)

    $lblEmail = New-Object System.Windows.Forms.Label
    $lblEmail.Text = 'OU Email:'
    $lblEmail.Location = New-Object System.Drawing.Point(470, 40); $lblEmail.Size = New-Object System.Drawing.Size(70, 22)
    $panelTop.Controls.Add($lblEmail)
    $txtEmail = New-Object System.Windows.Forms.TextBox
    $txtEmail.Location = New-Object System.Drawing.Point(546, 38); $txtEmail.Size = New-Object System.Drawing.Size(320, 24)
    $panelTop.Controls.Add($txtEmail)

    $btnSearch = New-StyledButton -Text 'Pesquisar' -X 228 -Y 76 -BackColor $script:PaletteBlue.Dark -ForeColor 'White' -Bold $true
    $panelTop.Controls.Add($btnSearch)
    $btnSave = New-StyledButton -Text 'Guardar (.txt)' -X 378 -Y 76 -BackColor $script:PaletteGreen.Dark -ForeColor 'White' -Bold $true -Width 140
    $btnSave.Enabled = $false; $panelTop.Controls.Add($btnSave)
    $btnClear = New-StyledButton -Text 'Limpar' -X 528 -Y 76 -Width 100
    $panelTop.Controls.Add($btnClear)

    $panelStatus = New-Object System.Windows.Forms.Panel
    $panelStatus.Dock = 'Bottom'; $panelStatus.Height = 30; $panelStatus.Padding = '12,6,12,6'
    $panelStatus.BackColor = [System.Drawing.Color]::White
    $tab.Controls.Add($panelStatus)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = 'Pronto.'; $lblStatus.Dock = 'Fill'
    $panelStatus.Controls.Add($lblStatus)

    $txtOutput = New-Object System.Windows.Forms.TextBox
    $txtOutput.Multiline = $true; $txtOutput.ReadOnly = $true
    $txtOutput.ScrollBars = 'Both'; $txtOutput.WordWrap = $false
    $txtOutput.Font = New-Object System.Drawing.Font('Consolas', 10)
    $txtOutput.Dock = 'Fill'
    $txtOutput.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
    $tab.Controls.Add($txtOutput); $txtOutput.BringToFront()

    $btnSearch.Add_Click({
        $user  = $txtUser.Text.Trim()
        $email = $txtEmail.Text.Trim()
        if (-not $user -and -not $email) {
            [System.Windows.Forms.MessageBox]::Show('Indique username ou email.', 'Aviso', 'OK', 'Warning') | Out-Null
            return
        }
        $btnSearch.Enabled = $false; $btnSave.Enabled = $false; $btnClear.Enabled = $false
        $txtOutput.Text = ''
        $lblStatus.Text = 'A consultar AD...'
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $pump = { [System.Windows.Forms.Application]::DoEvents() }
            $info = UI_Get-UserInfo -Username $user -Email $email -PumpUI $pump
            if (-not $info) {
                $txtOutput.Text = '(user nao encontrado)'
                $lblStatus.Text = 'Sem resultados.'
            } else {
                $script:UI_LastInfo = $info
                $txtOutput.Text = UI_Format-Report $info
                $btnSave.Enabled = $true
                $lblStatus.Text = "OK: $($info.SamAccountName) ($($info.DisplayName))"
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro:`n$($_.Exception.Message)", 'Erro', 'OK', 'Error') | Out-Null
            $lblStatus.Text = "Erro: $($_.Exception.Message)"
        } finally {
            $btnSearch.Enabled = $true; $btnClear.Enabled = $true
        }
    }.GetNewClosure())

    $btnClear.Add_Click({
        $txtUser.Text = ''; $txtEmail.Text = ''; $txtOutput.Text = ''
        $script:UI_LastInfo = $null; $btnSave.Enabled = $false
        $lblStatus.Text = 'Pronto.'
    }.GetNewClosure())

    $btnSave.Add_Click({
        if (-not $script:UI_LastInfo) { return }
        $sf = New-Object System.Windows.Forms.SaveFileDialog
        $sf.Filter = 'Texto (*.txt)|*.txt'
        $sf.FileName = "UserInfo_$($script:UI_LastInfo.SamAccountName)_$((Get-Date).ToString('yyyyMMdd_HHmmss')).txt"
        if ($sf.ShowDialog() -ne 'OK') { return }
        try {
            Set-Content -Path $sf.FileName -Value $txtOutput.Text -Encoding UTF8
            $lblStatus.Text = "Guardado: $($sf.FileName)"
            $res = [System.Windows.Forms.MessageBox]::Show("Guardado:`n$($sf.FileName)`n`nAbrir?", 'OK', 'YesNo', 'Information')
            if ($res -eq 'Yes') { Start-Process -FilePath $sf.FileName }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro:`n$($_.Exception.Message)", 'Erro', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())

    return $tab
}
