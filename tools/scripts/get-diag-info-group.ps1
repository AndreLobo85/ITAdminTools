#--------------------------------------------------
#
# Versao 1.0 sem RSAT - @Vitor Rodrigues (2021-05-04)
# Versao 1.1 sem RSAT - @Higino Antunes  (2022-09-29) - adicao de campo DisplayName
# Versao 1.2 sem RSAT - @Higino Antunes  (2024-06-03) - GroupType nativo, Group Scope, Group Category
#--------------------------------------------------

#requires -version 2

param ($groupName = "")

Trap {"Error: $_"; Break;}

if ($groupName -eq "") {
"
.\get-diag-info-group.ps1 -groupName GPAMNR  -> Info de um grupo por nome
"
    return
}

$ACCOUNTDISABLE                  = 0x000002
$ADS_GROUP_TYPE_SECURITY_ENABLED = 0x80000000

$groupNameIn = $groupName

$ScriptName = $MyInvocation.MyCommand.Name
$ADS_GROUP_TYPE_SECURITY_ENABLED = 0x80000000
$PageSize = 250

$ADS_ESCAPEDMODE_ON = 2
$ADS_SETTYPE_DN = 4
$ADS_FORMAT_X500_DN = 7
$Pathname = new-object -comobject "Pathname"
[Void] $Pathname.GetType().InvokeMember("EscapedMode", "SetProperty", $NULL, $Pathname, $ADS_ESCAPEDMODE_ON)

function Get-EscapedPath {
  param(
    [String] $distinguishedName
  )
  [Void] $Pathname.GetType().InvokeMember("Set", "InvokeMethod", $NULL, $Pathname, ($distinguishedName, $ADS_SETTYPE_DN))
  $Pathname.GetType().InvokeMember("Retrieve", "InvokeMethod", $NULL, $Pathname, $ADS_FORMAT_X500_DN)
}

function Get-SearchResultProperty {
  param(
    [System.DirectoryServices.ResultPropertyCollection] $properties,
    [String] $propertyName
  )
  if ( $properties[$propertyName] ) {
    $properties[$propertyName][0]
  }
  else {
    ""
  }
}

function Get-DirEntryProperty {
  param(
    [System.DirectoryServices.DirectoryEntry] $dirEntry,
    [String] $propertyName
  )
  if ( $dirEntry.$propertyName ) {
    $dirEntry.$propertyName[0]
  }
  else {
    ""
  }
}

write-progress $ScriptName "Enumerating groups"
$domain = [ADSI] ""
$searcher = [ADSISearcher] "(objectClass=group)"
$searcher.SearchRoot = $domain
$searcher.PageSize = $PageSize
$searcher.SearchScope = "subtree"
$searcher.PropertiesToLoad.AddRange(@("name","grouptype","distinguishedname","description","managedby","member","info","whencreated","memberof","whenchanged","DisplayName", "GroupScope", "GroupCategory"))
$searcher.Filter = "(name="+$groupNameIn+")"
$searchResults = $searcher.FindAll()
$groupCounter = 0
$groupCount = $searchResults.Count
foreach ( $searchResult in $searchResults ) {
  $properties = $searchResult.Properties
  $domainName = "BESP"
  $groupName = Get-SearchResultProperty $properties "name"
  $groupDescription = Get-SearchResultProperty $properties "description"
  $groupDisplayName = Get-SearchResultProperty $properties "DisplayName"
  $groupScope = Get-SearchResultProperty $properties "GroupScope"
  $groupCategory = Get-SearchResultProperty $properties "GroupCategory"
  $groupDN = Get-SearchResultProperty $properties "GroupScope"
  $groupManagedBy = Get-SearchResultProperty $properties "managedby"
  $groupType = Get-SearchResultProperty $properties "grouptype"
  if ($groupType -eq 2) {
    $groupScope = "Global"
    $groupCategory = "Distribution"
  }
  elseif ($groupType -eq 8) {
    $groupScope = "Global"
    $groupCategory = "Distribution"
  }
  elseif ($groupType -eq -2147483640) {
    $groupScope = "Universal"
    $groupCategory = "Security"
  }
  elseif ($groupType -eq -2147483643 ) {
    $groupScope = "DomainLocal"
    $groupCategory = "Security"
  }
  elseif ($groupType -eq -2147483644) {
    $groupScope = "DomainLocal"
    $groupCategory = "Security"
  }
  elseif ($groupType -eq -2147483646) {
    $groupScope = "Global"
    $groupCategory = "Security"
  }
  $member = $properties["member"]
  $info = Get-SearchResultProperty $properties "info"
  $pGroupName = ""
  if (($memberof = Get-SearchResultProperty $properties "memberof") -ne $null) {
    foreach ( $pGroup in ($memberof | sort)) {
        $pGroupName += $pGroup.split(",")[0].split("=")[1] + ", "
    }
    if ($pGroupName.Length -ge 2) {
        $pGroupName = $pGroupName.substring(0,$pGroupName.Length-2)
    }
  }

    ""
    "----------------------------------------------"
    "Domain             : " + $domainName
    "Group Name         : " + $groupName
    "Description        : " + $groupDescription
    "Display Name       : " + $groupDisplayName
    "Distinguished Name : " + $groupDN
    "----------------------------------------------"
    "Group Type         : " + $GroupType
    "When Created       : " + (Get-SearchResultProperty $properties "whencreated")
    "When Changed       : " + (Get-SearchResultProperty $properties "whenchanged")
    "Group Scope        : " + $GroupScope
    "Group Category     : " + $GroupCategory
    "----------------------------------------------"
    "MemberOf           : " + $pGroupName
    "----------------------------------------------"
    "Info               : " + $info
    "----------------------------------------------"
    "Members            : "

  if ( $member ) {
    $memberCounter = 0
    $memberCount = ($member | measure-object).Count
    foreach ( $memberDN in ($member | sort) ) {
      $memberDirEntry = [ADSI] "LDAP://$(Get-EscapedPath $memberDN)"

      if ((Get-DirEntryProperty $memberDirEntry "class") -eq "u") {
          (Get-DirEntryProperty $memberDirEntry "class") + " - " + (Get-DirEntryProperty $memberDirEntry "samaccountname") + " - " + (-not (([adsi]$memberDirEntry).userAccountControl[0] -band $ACCOUNTDISABLE)) + " - " + (Get-DirEntryProperty $memberDirEntry "displayname")
      } else {
         if ((Get-DirEntryProperty $memberDirEntry "class") -eq "f") {
            $SIDText = ($memberDN.Split(","))[0].SubString(3)
            $SID = New-Object System.Security.Principal.SecurityIdentifier $SIDText
            (Get-DirEntryProperty $memberDirEntry "class") + " - " + $SID.Translate([System.Security.Principal.NTAccount]).Value
         } else {
            (Get-DirEntryProperty $memberDirEntry "class") + " - " + (Get-DirEntryProperty $memberDirEntry "samaccountname")
         }
      }
      $memberCounter++
    }
    "----------------------------------------------"
  }

  $groupCounter++
  if ( ($groupCounter % $PageSize) -eq 0 ) {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
  }
}
$searchResults.Dispose()
