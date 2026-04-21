<#
.SYNOPSIS
    Pesquisa e auditoria de grupos AD via ADSI (sem RSAT).

.DESCRIPTION
    Aceita:
      - Nome exacto (GPAMNR)
      - Wildcard (GNORMA*, *-ADMINS)
      - Lista de sufixos separados por virgula (NR,NF) -> converte em *NR, *NF

    Opcional:
      - Expansao recursiva de grupos aninhados (com deteccao de ciclos)
      - Filtro de apenas users activos
      - Exportacao Excel 3-folhas (Resumo / Detalhe / Por Grupo) ou CSV,
        com auto-abertura no fim

    Tudo via System.DirectoryServices (ADSI). Nao requer RSAT nem
    Get-ADGroup. Funciona em qualquer PC no dominio.

.EXAMPLE
    audit-groups.ps1 -groupName GPAMNR
    audit-groups.ps1 -groupName 'GNORMA*' -expand -activeOnly
    audit-groups.ps1 -groupName 'NR,NF' -expand -export
#>

param(
    [string]$groupName = "",
    [switch]$expand,
    [switch]$activeOnly,
    [switch]$export,
    [int]$maxGroups = 100
)

$ErrorActionPreference = 'Continue'
Trap {"Error: $_"; Break;}

if (-not $groupName) {
"
.\audit-groups.ps1 -groupName GPAMNR                 -> Info de um grupo exacto
.\audit-groups.ps1 -groupName 'GNORMA*'              -> Todos os comecados por GNORMA
.\audit-groups.ps1 -groupName '*-ADMINS'             -> Todos os terminados em -ADMINS
.\audit-groups.ps1 -groupName 'NR,NF' -expand        -> Todos os terminados em NR ou NF, expandido
.\audit-groups.ps1 -groupName 'NR,NF' -expand -activeOnly -export   -> Idem + so users activos + Excel
"
    return
}

$ACCOUNTDISABLE = 0x000002

# -------- Parsing do input -------
$filterPieces = @()
$inputTrimmed = $groupName.Trim()
$hasComma = $inputTrimmed -match ','
$hasWildcard = $inputTrimmed -match '\*|\?'

if ($hasComma -and -not $hasWildcard) {
    # Lista de sufixos: NR,NF -> *NR, *NF
    $parts = $inputTrimmed -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    foreach ($p in $parts) { $filterPieces += "(name=*$p)" }
    "[INFO] Modo sufixo: $($parts.Count) sufixo(s) -> $($parts -join ', ')"
} elseif ($hasComma) {
    # Lista de padroes mistos
    $parts = $inputTrimmed -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    foreach ($p in $parts) { $filterPieces += "(name=$p)" }
    "[INFO] Modo padroes multiplos: $($parts.Count) padroes"
} else {
    # Nome unico (exacto ou com wildcard)
    $filterPieces += "(name=$inputTrimmed)"
}

$nameFilter = if ($filterPieces.Count -eq 1) { $filterPieces[0] } else { '(|' + ($filterPieces -join '') + ')' }
$ldapFilter = "(&(objectClass=group)$nameFilter)"

"[INFO] LDAP filter: $ldapFilter"
"[INFO] Opcoes: expand=$($expand.IsPresent), activeOnly=$($activeOnly.IsPresent), export=$($export.IsPresent)"

# -------- Pesquisa inicial --------
$domain = [ADSI] ""
$searcher = [ADSISearcher] $ldapFilter
$searcher.SearchRoot = $domain
$searcher.PageSize = 250
$searcher.SearchScope = "subtree"
$searcher.SizeLimit = $maxGroups + 1
$searcher.PropertiesToLoad.AddRange(@("name","grouptype","distinguishedname","description","managedby","member","memberof","whencreated","whenchanged","displayname")) | Out-Null

$results = $searcher.FindAll()
$groupCount = $results.Count

if ($groupCount -eq 0) {
    ""
    "[WARN] Nenhum grupo encontrado."
    $searcher.Dispose()
    return
}

""
"[OK] $groupCount grupo(s) encontrado(s)"
if ($groupCount -gt $maxGroups) {
    "[WARN] Cap a $maxGroups. Restringe o padrao para ver tudo (ou usa -maxGroups)."
}
""

# -------- Helpers --------
function Get-Prop {
    param($props, [string]$name)
    if ($props[$name]) { $props[$name][0] } else { "" }
}

function Resolve-GroupTypeText {
    param([int64]$t)
    $map = @{
        2            = 'Global | Distribution'
        8            = 'Global | Distribution (universal)'
        -2147483640  = 'Universal | Security'
        -2147483643  = 'DomainLocal | Security'
        -2147483644  = 'DomainLocal | Security'
        -2147483646  = 'Global | Security'
    }
    if ($map.ContainsKey($t)) { $map[$t] } else { "Tipo=$t" }
}

# Estrutura para resultado (quando expand e export sao usados)
$auditRows = New-Object System.Collections.Generic.List[object]

function Add-Row {
    param(
        [string]$TargetGroup,
        [string]$Path,
        [int]$Depth,
        [string]$MemberType,
        [string]$ParentGroup,
        [string]$SamAccount,
        [string]$DisplayName,
        [string]$Email,
        [object]$Enabled
    )
    $null = $auditRows.Add([PSCustomObject]@{
        TargetGroup = $TargetGroup
        Path        = $Path
        Depth       = $Depth
        MemberType  = $MemberType
        ParentGroup = $ParentGroup
        SamAccount  = $SamAccount
        DisplayName = $DisplayName
        Email       = $Email
        Enabled     = $Enabled
    })
}

function Expand-Group {
    param(
        [string]$GroupDN,
        [string]$GroupName,
        [string]$TargetGroup,
        [string]$PathPrefix,
        [int]$Depth,
        [System.Collections.Generic.HashSet[string]]$Seen
    )

    $currentPath = if ($PathPrefix) { "$PathPrefix > $GroupName" } else { $GroupName }

    if (-not $Seen.Add($GroupDN)) {
        "  $('  ' * $Depth)[LOOP] Referencia ciclica para $GroupName - abortar ramo"
        Add-Row -TargetGroup $TargetGroup -Path $currentPath -Depth $Depth -MemberType 'LoopDetected' `
            -ParentGroup $GroupName -SamAccount '' -DisplayName '(ref. ciclica)' -Email '' -Enabled $null
        return
    }

    $groupEntry = [ADSI]"LDAP://$GroupDN"
    $memberDNs = @($groupEntry.Properties["member"])

    if ($memberDNs.Count -eq 0) {
        "  $('  ' * $Depth)(grupo vazio)"
        Add-Row -TargetGroup $TargetGroup -Path $currentPath -Depth $Depth -MemberType 'EmptyGroup' `
            -ParentGroup $GroupName -SamAccount '' -DisplayName '(vazio)' -Email '' -Enabled $null
        return
    }

    foreach ($memberDN in $memberDNs) {
        try {
            $memberEntry = [ADSI]"LDAP://$memberDN"
            $cls = $memberEntry.SchemaClassName
            if (-not $cls) { $cls = "$($memberEntry.objectClass | Select-Object -Last 1)" }
            $sam = "$($memberEntry.sAMAccountName)"
            $disp = "$($memberEntry.displayName)"
            $mail = "$($memberEntry.mail)"

            if ($cls -eq 'user') {
                $uac = [int64]"$($memberEntry.userAccountControl)"
                $isEnabled = -not ($uac -band $ACCOUNTDISABLE)
                if ($activeOnly -and -not $isEnabled) { continue }
                $state = if ($isEnabled) { 'OK' } else { 'disabled' }
                "  $('  ' * $Depth)user  $sam  $disp  $state"
                Add-Row -TargetGroup $TargetGroup -Path $currentPath -Depth $Depth -MemberType 'User' `
                    -ParentGroup $GroupName -SamAccount $sam -DisplayName $disp -Email $mail -Enabled $isEnabled
            } elseif ($cls -eq 'group') {
                "  $('  ' * $Depth)[ grupo: $sam ]"
                Add-Row -TargetGroup $TargetGroup -Path $currentPath -Depth $Depth -MemberType 'NestedGroup' `
                    -ParentGroup $GroupName -SamAccount $sam -DisplayName "(grupo: $sam)" -Email '' -Enabled $null
                Expand-Group -GroupDN $memberDN -GroupName $sam -TargetGroup $TargetGroup `
                    -PathPrefix $currentPath -Depth ($Depth + 1) -Seen $Seen
            } else {
                "  $('  ' * $Depth)other:$cls  $sam  $disp"
                Add-Row -TargetGroup $TargetGroup -Path $currentPath -Depth $Depth -MemberType "Other:$cls" `
                    -ParentGroup $GroupName -SamAccount $sam -DisplayName $disp -Email $mail -Enabled $null
            }
        } catch {
            "  $('  ' * $Depth)[ERR] nao consegui ler $($memberDN): $($_.Exception.Message)"
        }
    }
}

# -------- Processar cada grupo encontrado --------
$idx = 0
foreach ($result in $results) {
    if ($idx -ge $maxGroups) { break }
    $idx++

    $props = $result.Properties
    $gName = Get-Prop $props 'name'
    $gDN = Get-Prop $props 'distinguishedname'
    $gDesc = Get-Prop $props 'description'
    $gDisp = Get-Prop $props 'displayname'
    $gType = [int64](Get-Prop $props 'grouptype')
    $gCreated = Get-Prop $props 'whencreated'
    $gChanged = Get-Prop $props 'whenchanged'
    $gManagedBy = Get-Prop $props 'managedby'

    $memberOfList = @()
    if ($props['memberof']) {
        $memberOfList = $props['memberof'] | Sort-Object | ForEach-Object { ($_ -replace 'CN=([^,]+),.+', '$1') }
    }

    "===================================================================="
    "Grupo         : $gName   ($idx/$groupCount)"
    "Display Name  : $gDisp"
    "Descricao     : $gDesc"
    "DN            : $gDN"
    "Tipo          : " + (Resolve-GroupTypeText $gType)
    "Criado em     : $gCreated"
    "Alterado em   : $gChanged"
    "Managed By    : $gManagedBy"
    "MemberOf      : " + ($memberOfList -join '; ')
    "--------------------------------------------------------------------"

    if ($expand) {
        "Membros (expandido recursivamente):"
        $seen = New-Object 'System.Collections.Generic.HashSet[string]'
        Expand-Group -GroupDN $gDN -GroupName $gName -TargetGroup $gName -PathPrefix '' -Depth 0 -Seen $seen
    } else {
        $directMembers = @($props['member'])
        "Membros directos: $($directMembers.Count) (usa -expand para ver detalhes recursivamente)"
        foreach ($dn in ($directMembers | Select-Object -First 20)) {
            $shortName = ($dn -replace 'CN=([^,]+),.+', '$1')
            "  - $shortName"
        }
        if ($directMembers.Count -gt 20) { "  ... +$($directMembers.Count - 20) mais" }
    }
    ""
}

$searcher.Dispose()

# -------- Sumario final --------
if ($expand) {
    ""
    "===================================================================="
    "SUMARIO"
    "===================================================================="
    $usersActive   = @($auditRows | Where-Object { $_.MemberType -eq 'User' -and $_.Enabled })
    $usersInactive = @($auditRows | Where-Object { $_.MemberType -eq 'User' -and $_.Enabled -eq $false })
    $distinctUsers = ($auditRows | Where-Object MemberType -eq 'User' | Select-Object -ExpandProperty SamAccount -Unique).Count
    $nested = ($auditRows | Where-Object MemberType -eq 'NestedGroup').Count
    $loops  = ($auditRows | Where-Object MemberType -eq 'LoopDetected').Count

    "Grupos alvo     : $groupCount"
    "Aninhamentos    : $nested"
    "Users distintos : $distinctUsers"
    "Users activos   : $($usersActive.Count)"
    "Users inactivos : $($usersInactive.Count)"
    if ($loops -gt 0) { "Loops detectados: $loops" }
    "Total linhas    : $($auditRows.Count)"
}

# -------- Export --------
if ($export -and $auditRows.Count -gt 0) {
    ""
    "[INFO] A exportar..."
    $timestamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $baseName = "AuditGruposAD_$timestamp"
    $downloads = [Environment]::GetFolderPath('MyDocuments')
    if (-not (Test-Path $downloads)) { $downloads = $env:USERPROFILE }

    $useExcel = $false
    try { $null = New-Object -ComObject Excel.Application; $useExcel = $true } catch {}

    if ($useExcel) {
        $outPath = Join-Path $downloads "$baseName.xlsx"
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        try {
            $wb = $excel.Workbooks.Add()
            while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

            # Folha 1: Resumo
            $ws1 = $wb.Worksheets.Item(1); $ws1.Name = 'Resumo'
            $ws1.Cells.Item(1,1) = 'Auditoria de Grupos AD'
            $ws1.Cells.Item(1,1).Font.Size = 16
            $ws1.Cells.Item(1,1).Font.Bold = $true
            $ws1.Cells.Item(2,1) = 'Padrao:';        $ws1.Cells.Item(2,2) = $groupName
            $ws1.Cells.Item(3,1) = 'Data:';          $ws1.Cells.Item(3,2) = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            $ws1.Cells.Item(4,1) = 'Grupos alvo:';   $ws1.Cells.Item(4,2) = $groupCount
            $ws1.Cells.Item(5,1) = 'Linhas totais:'; $ws1.Cells.Item(5,2) = $auditRows.Count
            $ws1.Cells.Item(6,1) = 'Users distintos:';
            $ws1.Cells.Item(6,2) = ($auditRows | Where-Object MemberType -eq 'User' | Select-Object -ExpandProperty SamAccount -Unique).Count
            for ($r=2;$r-le 6;$r++) { $ws1.Cells.Item($r,1).Font.Bold=$true }

            # Folha 2: Detalhe
            $ws2 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws1); $ws2.Name = 'Detalhe'
            $headers = @('GrupoAlvo','Caminho','Nivel','Tipo','GrupoPai','SamAccount','Nome','Email','Ativo')
            for ($c=0; $c -lt $headers.Count; $c++) { $ws2.Cells.Item(1,$c+1) = $headers[$c] }
            $row = 2
            foreach ($r in $auditRows) {
                $ws2.Cells.Item($row,1) = $r.TargetGroup
                $ws2.Cells.Item($row,2) = $r.Path
                $ws2.Cells.Item($row,3) = $r.Depth
                $ws2.Cells.Item($row,4) = $r.MemberType
                $ws2.Cells.Item($row,5) = $r.ParentGroup
                $ws2.Cells.Item($row,6) = $r.SamAccount
                $ws2.Cells.Item($row,7) = $r.DisplayName
                $ws2.Cells.Item($row,8) = $r.Email
                $ws2.Cells.Item($row,9) = if ($null -ne $r.Enabled) { if ($r.Enabled) {'Sim'} else {'Nao'} } else { '' }
                $row++
            }
            if ($auditRows.Count -gt 0) {
                $rng = $ws2.Range($ws2.Cells.Item(1,1), $ws2.Cells.Item($row-1, $headers.Count))
                $lo = $ws2.ListObjects.Add(1, $rng, $null, 1)
                $lo.TableStyle = 'TableStyleMedium2'
            }
            $ws2.Columns.AutoFit() | Out-Null

            $ws1.Activate()
            $wb.SaveAs($outPath, 51) # xlOpenXMLWorkbook
            $wb.Close($false)
            "[OK] Exportado: $outPath"
            Start-Process $outPath
        } finally {
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            [GC]::Collect() | Out-Null
        }
    } else {
        $outPath = Join-Path $downloads "$baseName.csv"
        $auditRows | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
        "[WARN] Excel nao disponivel, exportado como CSV: $outPath"
        Start-Process $outPath
    }
}