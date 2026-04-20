# ============================================================
# GroupInfo.ps1 - Ferramenta: diagnostico de grupo AD (sem RSAT)
# Baseado em get-diag-info-group.ps1 (@Vitor Rodrigues / @Higino Antunes)
# Exporta: New-GroupInfoTab
# ============================================================

$script:GI_UAC_ACCOUNTDISABLE = 0x000002

function GI_Get-DirEntryProperty {
    param([System.DirectoryServices.DirectoryEntry]$dirEntry, [string]$propertyName)
    try {
        if ($dirEntry.$propertyName) { return "$($dirEntry.$propertyName[0])" }
    } catch { }
    return ''
}

function GI_Get-SearchResultProperty {
    param([System.DirectoryServices.ResultPropertyCollection]$properties, [string]$propertyName)
    if ($properties[$propertyName]) { return "$($properties[$propertyName][0])" }
    return ''
}

function GI_Resolve-GroupType {
    param([int64]$gt)
    switch ($gt) {
        2            { @{Scope='Global';      Category='Distribution'} }
        8            { @{Scope='Global';      Category='Distribution'} }
        -2147483640  { @{Scope='Universal';   Category='Security'} }
        -2147483643  { @{Scope='DomainLocal'; Category='Security'} }
        -2147483644  { @{Scope='DomainLocal'; Category='Security'} }
        -2147483646  { @{Scope='Global';      Category='Security'} }
        default      { @{Scope='Unknown';     Category='Unknown'} }
    }
}

function GI_Get-GroupInfo {
    param(
        [string]$GroupName,
        [ScriptBlock]$OnProgress = $null,
        [ScriptBlock]$PumpUI = $null
    )

    if (-not $GroupName) { throw 'Indique o nome do grupo.' }

    $ADS_ESCAPEDMODE_ON = 2
    $ADS_SETTYPE_DN     = 4
    $ADS_FORMAT_X500_DN = 7
    $pathname = New-Object -ComObject 'Pathname'
    [void]$pathname.GetType().InvokeMember('EscapedMode', 'SetProperty', $null, $pathname, $ADS_ESCAPEDMODE_ON)
    $getEscapedPath = {
        param([string]$dn)
        [void]$pathname.GetType().InvokeMember('Set', 'InvokeMethod', $null, $pathname, ($dn, $ADS_SETTYPE_DN))
        return $pathname.GetType().InvokeMember('Retrieve', 'InvokeMethod', $null, $pathname, $ADS_FORMAT_X500_DN)
    }

    $domain = [ADSI]''
    $searcher = [ADSISearcher]'(objectClass=group)'
    $searcher.SearchRoot = $domain
    $searcher.PageSize = 250
    $searcher.SearchScope = 'subtree'
    [void]$searcher.PropertiesToLoad.AddRange(@(
        'name','grouptype','distinguishedname','description','managedby',
        'member','info','whencreated','memberof','whenchanged','DisplayName'
    ))
    $searcher.Filter = "(name=$GroupName)"
    $searchResults = $searcher.FindAll()
    if ($searchResults.Count -eq 0) { return $null }

    $r = $searchResults[0]
    $props = $r.Properties
    $groupType = [int64]((GI_Get-SearchResultProperty $props 'grouptype'))
    $typeInfo  = GI_Resolve-GroupType $groupType

    $memberofRaw = $props['memberof']
    $memberOfList = if ($memberofRaw) {
        (($memberofRaw | Sort-Object | ForEach-Object { $_.Split(',')[0].Split('=')[1] }) -join ', ')
    } else { '' }

    # Enumerar membros
    $members = New-Object System.Collections.Generic.List[object]
    $memberDNs = @($props['member'])
    $total = $memberDNs.Count
    $i = 0
    foreach ($memberDN in ($memberDNs | Sort-Object)) {
        $i++
        if ($OnProgress) { & $OnProgress $i $total $memberDN }
        if ($PumpUI -and ($i % 10 -eq 0)) { & $PumpUI }
        try {
            $escDN = & $getEscapedPath $memberDN
            $mDE = [ADSI]"LDAP://$escDN"
            $class = GI_Get-DirEntryProperty $mDE 'class'
            if ($class -eq 'u') {
                $sam = GI_Get-DirEntryProperty $mDE 'samaccountname'
                $dname = GI_Get-DirEntryProperty $mDE 'displayname'
                $enabled = ''
                try {
                    $enabled = -not [bool](([adsi]$mDE).userAccountControl[0] -band $script:GI_UAC_ACCOUNTDISABLE)
                } catch { }
                $members.Add([PSCustomObject]@{ Class='user'; SamAccount=$sam; DisplayName=$dname; Enabled=$enabled; DN=$memberDN })
            }
            elseif ($class -eq 'f') {
                $sidText = ($memberDN.Split(','))[0].Substring(3)
                $sidName = ''
                try { $sidName = (New-Object System.Security.Principal.SecurityIdentifier($sidText)).Translate([System.Security.Principal.NTAccount]).Value } catch { $sidName = $sidText }
                $members.Add([PSCustomObject]@{ Class='foreignSecurityPrincipal'; SamAccount=$sidName; DisplayName=''; Enabled=$null; DN=$memberDN })
            }
            else {
                $sam = GI_Get-DirEntryProperty $mDE 'samaccountname'
                $members.Add([PSCustomObject]@{ Class=$class; SamAccount=$sam; DisplayName=''; Enabled=$null; DN=$memberDN })
            }
        } catch {
            $members.Add([PSCustomObject]@{ Class='error'; SamAccount=''; DisplayName="(erro: $($_.Exception.Message))"; Enabled=$null; DN=$memberDN })
        }
    }

    return [PSCustomObject]@{
        GroupName         = GI_Get-SearchResultProperty $props 'name'
        Description       = GI_Get-SearchResultProperty $props 'description'
        DisplayName       = GI_Get-SearchResultProperty $props 'DisplayName'
        DistinguishedName = GI_Get-SearchResultProperty $props 'distinguishedname'
        GroupType         = $groupType
        GroupScope        = $typeInfo.Scope
        GroupCategory     = $typeInfo.Category
        ManagedBy         = GI_Get-SearchResultProperty $props 'managedby'
        WhenCreated       = GI_Get-SearchResultProperty $props 'whencreated'
        WhenChanged       = GI_Get-SearchResultProperty $props 'whenchanged'
        Info              = GI_Get-SearchResultProperty $props 'info'
        MemberOf          = $memberOfList
        Members           = $members
    }
}

function GI_Format-Report {
    param($Info)
    if (-not $Info) { return '(grupo nao encontrado)' }
    $sep = ('-' * 60)
    $lines = @()
    $lines += $sep
    $lines += 'Group Name         : ' + $Info.GroupName
    $lines += 'Description        : ' + $Info.Description
    $lines += 'Display Name       : ' + $Info.DisplayName
    $lines += 'Distinguished Name : ' + $Info.DistinguishedName
    $lines += $sep
    $lines += 'Group Type (code)  : ' + $Info.GroupType
    $lines += 'Group Scope        : ' + $Info.GroupScope
    $lines += 'Group Category     : ' + $Info.GroupCategory
    $lines += 'Managed By         : ' + $Info.ManagedBy
    $lines += 'When Created       : ' + $Info.WhenCreated
    $lines += 'When Changed       : ' + $Info.WhenChanged
    $lines += $sep
    $lines += 'MemberOf           : ' + $Info.MemberOf
    $lines += $sep
    $lines += 'Info               : ' + $Info.Info
    $lines += $sep
    $lines += "Members ($($Info.Members.Count)):"
    foreach ($m in $Info.Members) {
        $enStr = if ($null -ne $m.Enabled) { "$($m.Enabled)" } else { '' }
        $lines += "  $($m.Class) - $($m.SamAccount) - $enStr - $($m.DisplayName)"
    }
    $lines += $sep
    return ($lines -join "`r`n")
}

function GI_Export-ToExcel {
    param($Info, [string]$OutputPath)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.ScreenUpdating = $false
    try {
        $wb = $excel.Workbooks.Add()
        while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
        $ws = $wb.Worksheets.Item(1); $ws.Name = 'GroupInfo'

        $ws.Cells.Item(1,1) = "Grupo AD: $($Info.GroupName)"
        $ws.Cells.Item(1,1).Font.Size = 16; $ws.Cells.Item(1,1).Font.Bold = $true
        $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1,4)).Merge() | Out-Null

        $row = 3
        $kv = @(
            @('Description', $Info.Description),
            @('Display Name', $Info.DisplayName),
            @('Distinguished Name', $Info.DistinguishedName),
            @('Group Type (code)', $Info.GroupType),
            @('Group Scope', $Info.GroupScope),
            @('Group Category', $Info.GroupCategory),
            @('Managed By', $Info.ManagedBy),
            @('When Created', $Info.WhenCreated),
            @('When Changed', $Info.WhenChanged),
            @('Member Of', $Info.MemberOf),
            @('Info', $Info.Info),
            @('Total Members', $Info.Members.Count)
        )
        foreach ($pair in $kv) {
            $ws.Cells.Item($row,1) = $pair[0]
            $ws.Cells.Item($row,1).Font.Bold = $true
            $ws.Cells.Item($row,2) = "$($pair[1])"
            $row++
        }
        $row += 2

        # Tabela de membros
        $hdrs = @('Class','SamAccount','DisplayName','Enabled','DN')
        $startRow = $row
        for ($c = 0; $c -lt $hdrs.Count; $c++) { $ws.Cells.Item($startRow, $c+1) = $hdrs[$c] }
        $row++
        foreach ($m in $Info.Members) {
            $ws.Cells.Item($row,1) = $m.Class
            $ws.Cells.Item($row,2) = $m.SamAccount
            $ws.Cells.Item($row,3) = $m.DisplayName
            $ws.Cells.Item($row,4) = if ($null -ne $m.Enabled) { if ($m.Enabled) {'Sim'} else {'Nao'} } else { '' }
            $ws.Cells.Item($row,5) = $m.DN
            if ($null -ne $m.Enabled -and -not $m.Enabled) {
                $r = $ws.Range($ws.Cells.Item($row,1), $ws.Cells.Item($row,5))
                $r.Font.Color = Convert-RgbToBgr 0x808080
                $r.Font.Italic = $true
            }
            $row++
        }
        if ($Info.Members.Count -gt 0) {
            $rng = $ws.Range($ws.Cells.Item($startRow,1), $ws.Cells.Item($row-1, $hdrs.Count))
            $lo = $ws.ListObjects.Add(1, $rng, $null, 1); $lo.Name = 'tblMembers'; $lo.TableStyle = 'TableStyleMedium2'
        }
        $ws.Columns.AutoFit() | Out-Null

        if ([System.IO.Path]::GetExtension($OutputPath) -ne '.xlsx') {
            $OutputPath = [System.IO.Path]::ChangeExtension($OutputPath, '.xlsx')
        }
        $wb.SaveAs($OutputPath, 51); $wb.Close($false)
    }
    finally {
        $excel.ScreenUpdating = $true; $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect() | Out-Null; [GC]::WaitForPendingFinalizers() | Out-Null
    }
}

function New-GroupInfoTab {
    $script:GI_LastInfo = $null

    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = 'Group Info'
    $tab.BackColor = [System.Drawing.Color]::White

    $panelTop = New-Object System.Windows.Forms.Panel
    $panelTop.Dock = 'Top'; $panelTop.Height = 130
    $panelTop.Padding = '12,12,12,12'
    $panelTop.BackColor = [System.Drawing.Color]::White
    $tab.Controls.Add($panelTop)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = 'Diagnostico de Grupo AD'
    $lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $script:PaletteBlue.Dark
    $lblTitle.Location = New-Object System.Drawing.Point(12, 4); $lblTitle.Size = New-Object System.Drawing.Size(600, 26)
    $panelTop.Controls.Add($lblTitle)

    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = 'Nome do grupo:'
    $lblName.Location = New-Object System.Drawing.Point(12, 40); $lblName.Size = New-Object System.Drawing.Size(120, 22)
    $panelTop.Controls.Add($lblName)
    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(140, 38); $txtName.Size = New-Object System.Drawing.Size(320, 24)
    $panelTop.Controls.Add($txtName)

    $btnSearch = New-StyledButton -Text 'Pesquisar' -X 140 -Y 76 -BackColor $script:PaletteBlue.Dark -ForeColor 'White' -Bold $true
    $panelTop.Controls.Add($btnSearch)
    $btnExport = New-StyledButton -Text 'Exportar Excel' -X 290 -Y 76 -BackColor $script:PaletteGreen.Dark -ForeColor 'White' -Bold $true -Width 140
    $btnExport.Enabled = $false; $panelTop.Controls.Add($btnExport)
    $btnSave = New-StyledButton -Text 'Guardar (.txt)' -X 440 -Y 76 -Width 140
    $btnSave.Enabled = $false; $panelTop.Controls.Add($btnSave)
    $btnClear = New-StyledButton -Text 'Limpar' -X 590 -Y 76 -Width 100
    $panelTop.Controls.Add($btnClear)

    $panelStatus = New-Object System.Windows.Forms.Panel
    $panelStatus.Dock = 'Bottom'; $panelStatus.Height = 52; $panelStatus.Padding = '12,6,12,6'
    $panelStatus.BackColor = [System.Drawing.Color]::White
    $tab.Controls.Add($panelStatus)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = 'Pronto.'; $lblStatus.Location = New-Object System.Drawing.Point(12, 6)
    $lblStatus.Size = New-Object System.Drawing.Size(1100, 20); $lblStatus.Anchor = 'Top, Left, Right'
    $panelStatus.Controls.Add($lblStatus)
    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Location = New-Object System.Drawing.Point(12, 28); $progress.Size = New-Object System.Drawing.Size(1100, 16)
    $progress.Anchor = 'Top, Left, Right'; $progress.Style = 'Continuous'
    $panelStatus.Controls.Add($progress)

    $txtOutput = New-Object System.Windows.Forms.TextBox
    $txtOutput.Multiline = $true; $txtOutput.ReadOnly = $true
    $txtOutput.ScrollBars = 'Both'; $txtOutput.WordWrap = $false
    $txtOutput.Font = New-Object System.Drawing.Font('Consolas', 10)
    $txtOutput.Dock = 'Fill'
    $txtOutput.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
    $tab.Controls.Add($txtOutput); $txtOutput.BringToFront()

    $btnSearch.Add_Click({
        $name = $txtName.Text.Trim()
        if (-not $name) { [System.Windows.Forms.MessageBox]::Show('Indique o nome do grupo.', 'Aviso', 'OK', 'Warning') | Out-Null; return }
        $btnSearch.Enabled = $false; $btnExport.Enabled = $false; $btnSave.Enabled = $false; $btnClear.Enabled = $false
        $txtOutput.Text = ''; $progress.Value = 0
        $lblStatus.Text = 'A consultar AD...'
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $progressCb = {
                param($i, $total, $dn)
                $pct = if ($total -gt 0) { [int](($i / $total) * 100) } else { 0 }
                $progress.Value = [Math]::Min(100, [Math]::Max(0, $pct))
                $lblStatus.Text = "A ler membro $i/$total"
                [System.Windows.Forms.Application]::DoEvents()
            }.GetNewClosure()
            $pump = { [System.Windows.Forms.Application]::DoEvents() }
            $info = GI_Get-GroupInfo -GroupName $name -OnProgress $progressCb -PumpUI $pump
            if (-not $info) {
                $txtOutput.Text = '(grupo nao encontrado)'
                $lblStatus.Text = 'Sem resultados.'; $progress.Value = 0
            } else {
                $script:GI_LastInfo = $info
                $txtOutput.Text = GI_Format-Report $info
                $btnExport.Enabled = $true; $btnSave.Enabled = $true
                $progress.Value = 100
                $lblStatus.Text = "OK: $($info.GroupName) | $($info.Members.Count) membros"
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro:`n$($_.Exception.Message)", 'Erro', 'OK', 'Error') | Out-Null
            $lblStatus.Text = "Erro: $($_.Exception.Message)"
        } finally {
            $btnSearch.Enabled = $true; $btnClear.Enabled = $true
        }
    }.GetNewClosure())

    $btnClear.Add_Click({
        $txtName.Text = ''; $txtOutput.Text = ''; $script:GI_LastInfo = $null
        $btnExport.Enabled = $false; $btnSave.Enabled = $false
        $lblStatus.Text = 'Pronto.'; $progress.Value = 0
    }.GetNewClosure())

    $btnSave.Add_Click({
        if (-not $script:GI_LastInfo) { return }
        $sf = New-Object System.Windows.Forms.SaveFileDialog
        $sf.Filter = 'Texto (*.txt)|*.txt'
        $sf.FileName = "GroupInfo_$($script:GI_LastInfo.GroupName)_$((Get-Date).ToString('yyyyMMdd_HHmmss')).txt"
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

    $btnExport.Add_Click({
        if (-not $script:GI_LastInfo) { return }
        $sf = New-Object System.Windows.Forms.SaveFileDialog
        $sf.Filter = 'Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv'
        $sf.FileName = "GroupInfo_$($script:GI_LastInfo.GroupName)_$((Get-Date).ToString('yyyyMMdd_HHmmss')).xlsx"
        if ($sf.ShowDialog() -ne 'OK') { return }
        $out = $sf.FileName
        try {
            if ($sf.FilterIndex -eq 1) {
                if (-not (Test-ExcelAvailable)) {
                    [System.Windows.Forms.MessageBox]::Show('Excel nao disponivel. Use Guardar (.txt) ou CSV.', 'Aviso', 'OK', 'Warning') | Out-Null
                    return
                }
                GI_Export-ToExcel -Info $script:GI_LastInfo -OutputPath $out
            } else {
                $out = Export-ResultsToCsv -Results $script:GI_LastInfo.Members -OutputPath $out
            }
            $lblStatus.Text = "Exportado: $out"
            $res = [System.Windows.Forms.MessageBox]::Show("Guardado:`n$out`n`nAbrir?", 'OK', 'YesNo', 'Information')
            if ($res -eq 'Yes') { Start-Process -FilePath $out }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro:`n$($_.Exception.Message)", 'Erro', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())

    return $tab
}
