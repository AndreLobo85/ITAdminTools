#--------------------------------------------------
#
# Versao 1.0 sem RSAT - @Vitor Rodrigues (2020-03-31)
# Versao 1.1 sem RSAT - @Vitor Rodrigues (2020-04-02): PasswdLastSet Date, badPwdCount
# Versao 1.2 sem RSAT - @Higino Antunes (2020-11-16)
# Versao 1.3 sem RSAT - @Higino Antunes (2020-11-13): Department
# Versao 1.4 sem RSAT - @Higino Antunes (2021-04-06): Manager + bugfix data expiracao
# Versao 1.5 sem RSAT - @Higino Antunes (2021-04-16)
# Versao 2.0 sem RSAT - @Vitor Rodrigues (2021-04-20): proxyAddresses, homeMDB, ManagerName, grupos
# Versao 2.1 sem RSAT - @Higino Antunes (2021-04-26): whenCreated, UserPrincipalName
# Versao 2.1 sem RSAT - @Higino Antunes (2021-04-28): AccountExpires
# Versao 2.2 sem RSAT - @Higino Antunes (2021-05-05): mailNickName
# Versao 2.3 sem RSAT - @Higino Antunes (2021-06-24): Title
# Versao 2.4 sem RSAT - @Higino Antunes (2023-01-09): SamAccountName
#
#--------------------------------------------------

param ($email = "", $userName = "")

Trap {"Error: $_"; Break;}

if (($email -eq "") -and ($userName -eq "")) {
"
.\get-diag-info-aw2.ps1 -userName B26467  -> Info por username
.\get-diag-info-aw2.ps1 -email vitor.macieira.rodrigues@novobanco.pt  -> Info por email
"
    return
}

$ACCOUNTDISABLE       = 0x000002
$DONT_EXPIRE_PASSWORD = 0x010000
$PASSWORD_EXPIRED     = 0x800000
$LOCKOUT              = 0x000010

$D = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$Domain = [ADSI]"LDAP://$D"
$Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Searcher.PageSize = 1
$Searcher.SearchScope = "subtree"

if ($email -ne "") {
    $Searcher.Filter = "(|(mail="+$email+")(proxyAddresses=smtp:"+$email+"))"
} else {
    $Searcher.Filter = "(samaccountname="+$userName+")"
}

$Searcher.PropertiesToLoad.Add("name") > $Null
$Searcher.PropertiesToLoad.Add("SamAccountName") > $Null
$Searcher.PropertiesToLoad.Add("DisplayName") > $Null
$Searcher.PropertiesToLoad.Add("Description") > $Null
$Searcher.PropertiesToLoad.Add("mail") > $Null
$Searcher.PropertiesToLoad.Add("o") > $Null
$Searcher.PropertiesToLoad.Add("enabled") > $Null
$Searcher.PropertiesToLoad.Add("lastLogonTimestamp") > $Null
$Searcher.PropertiesToLoad.Add("PasswordExpired") > $Null
$Searcher.PropertiesToLoad.Add("extensionAttribute13") > $Null
$Searcher.PropertiesToLoad.Add("extensionAttribute14") > $Null
$Searcher.PropertiesToLoad.Add("MemberOf") > $Null
$Searcher.PropertiesToLoad.Add("pwdLastSet") > $Null
$Searcher.PropertiesToLoad.Add("msDS-UserPasswordExpiryTimeComputed") > $Null
$Searcher.PropertiesToLoad.Add("DistinguishedName") > $Null
$Searcher.PropertiesToLoad.Add("PasswordNeverExpires") > $Null
$Searcher.PropertiesToLoad.Add("Department") > $Null
$Searcher.PropertiesToLoad.Add("Manager") > $Null
$Searcher.PropertiesToLoad.Add("Title") > $Null
$Searcher.PropertiesToLoad.Add("Organization") > $Null
$Searcher.PropertiesToLoad.Add("proxyAddresses") > $Null
$Searcher.PropertiesToLoad.Add("homeMDB") > $Null
$Searcher.PropertiesToLoad.Add("memberof") > $Null
$Searcher.PropertiesToLoad.Add("UserPrincipalName") > $Null
$Searcher.PropertiesToLoad.Add("whenCreated") > $Null
$Searcher.PropertiesToLoad.Add("AccountExpires") > $Null
$Searcher.PropertiesToLoad.Add("mailNickName") > $Null
$Searcher.PropertiesToLoad.Add("msExchRemoteRecipientType") > $Null

$arrUsers = @{}

$DC = $D.DomainControllers[0]

$Server = $DC.Name
$Searcher.SearchRoot = "LDAP://$Server/" + $Domain.distinguishedName
$Results = $Searcher.FindAll()

if ($Result = $Results[0]) {
    $LockedOut = @()
    $badPwdCount = @()

    $dcs = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().DomainControllers

    foreach ($dc in $dcs) {
        $Searcher.SearchRoot = "LDAP://$dc/" + $Domain.distinguishedName
        $ResultT = $Searcher.FindAll()

        $LockedOut += [bool](([adsi]$ResultT.Properties.adspath[0]).userAccountControl[0] -band $LOCKOUT)
        $badPwdCount += ([adsi]$ResultT.Properties.adspath[0]).badPwdCount
    }

    try {
        $S2 = New-Object System.DirectoryServices.DirectorySearcher
        $S2.Filter = "(&(objectCategory=person)(objectClass=user)(distinguishedName="+($result.Properties["manager"])+"))"
        $S2.SearchRoot = "LDAP://$Server/" + $Domain.distinguishedName
        $Results2 = $S2.FindAll()

        if ($Results2 -ne "") {
            $manager = $Results2[0].Properties["DisplayName"]
        }
    } catch {}

    $proxyAddresses = $Result.Properties.Item("proxyAddresses")
    $pa = $proxyAddresses -match 'smtp:[\w@\.]+'

    $grupos = ""

    $Result.Properties.memberof | sort | foreach {
        $grupos += ($_ -replace "CN=([^,]+),.+", '$1') + "; "
    }
    $grupos = ($grupos -replace "; $","")

    if ($Result.Properties.Item("memberof") -like "*GO365PRO*") {
        $x1 = ($Result.Properties.Item("memberof") -like "*GO365PRO*").split(",")
        ($lixo, $goffice) = $x1[0].split("=")
    } else {
        $goffice = "Nao tem licenca!"
    }

    ""
    "---------------------------------------------"
    "Name                 : " + $Result.Properties.Item("name")
    "SamAccountName       : " + $Result.Properties.Item("SamAccountName")
    "Alias/mailNickName   : " + $Result.Properties.Item("mailNickName")
    "Description          : " + $Result.Properties.Item("Description")
    "DisplayName          : " + $Result.Properties.Item("DisplayName")
    "DistinguishedName    : " + $Result.Properties.Item("DistinguishedName")
    "Estrutura            : " + $Result.Properties.Item("Department")
    "Funcao               : " + $Result.Properties.Item("Title")
    "Manager              : " + $Result.Properties.Item("Manager")
    "Manager Name         : " + $manager
    "Outsourcer           : " + $Result.Properties.Item("Organization")
    "Organization         : " + $Result.Properties.Item("o")
    "mail                 : " + $Result.Properties.Item("mail")
    "alternativos         : " + $pa
    "Database             : " + ($Result.Properties["homeMDB"] -replace "CN=(\w+),.+", '$1')
    "---------------------------------------------"
    "Ultimo Logon         : " + [datetime]::fromfiletime($Result.Properties.Item("lastLogonTimestamp")[0]).ToString('yyyy-MM-dd HH:mm')
    "enabled              : " + (-not (([adsi]$Result.Properties.adspath[0]).userAccountControl[0] -band $ACCOUNTDISABLE))
    "PasswordExpired      : " + [bool](([adsi]$Result.Properties.adspath[0]).userAccountControl[0] -band $PASSWORD_EXPIRED)
    "PasswordLastSet      : " + [datetime]::fromfiletime($Result.Properties.Item("pwdLastSet")[0]).ToString('yyyy-MM-dd HH:mm')
    If ([bool](([adsi]$Result.Properties.adspath[0]).userAccountControl[0] -band $DONT_EXPIRE_PASSWORD)) {"Password expira em   : Nao expira" }
    else {"Password expira em   : " + [datetime]::fromfiletime($Result.Properties.Item("msDS-UserPasswordExpiryTimeComputed")[0]).ToString('yyyy-MM-dd HH:mm') }
    "LockedOut            : " + $LockedOut
    "badPwdCount          : " + $badPwdCount
    "Criado em            : " + $Result.Properties.Item("whenCreated").ToLocalTime().ToString("yyyy-MM-dd HH:mm")
    if (($Result.Properties.Item("AccountExpires") -like 0)){
        "User expira  em      : NAO EXPIRA"
    }
    elseif (($Result.Properties.Item("AccountExpires") -eq [int64]::MaxValue)) {
        "User expira  em      : NAO EXPIRA"
    }
    else {"User expira em       : " + [datetime]::fromfiletime($Result.Properties.Item("AccountExpires")[0]).ToString('yyyy-MM-dd HH:mm') }
    "---------------------------------------------"
    "msExchRemoteRecipientType : " + $Result.Properties.Item("msExchRemoteRecipientType")
    "extensionAttribute13      : " + $Result.Properties.Item("extensionAttribute13")
    "extensionAttribute14      : " + $Result.Properties.Item("extensionAttribute14")
    "Grupo SSPR (password)     : " + (($Result.Properties.Item("memberof") -like "*GO365SSPRNR*").Count -eq 1)
    "Grupo Cond Access L       : " + (($Result.Properties.Item("memberof") -like "*GO365CALNR*").Count -eq 1)
    "Grupo Lic Office          : " + $goffice
    "User Princpal Name        : " + $Result.Properties.Item("userPrincipalName")
    "---------------------------------------------"
    "Acede ao PAM?        : " + $(if ($grupos -match "GPAMNR") {"Sim"} else {"Nao"})
    "Grupos               : " + $grupos
    "---------------------------------------------"
}
