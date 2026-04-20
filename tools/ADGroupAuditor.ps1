# ============================================================
# ADGroupAuditor.ps1 - Ferramenta: auditar grupos AD por sufixo
# Exporta: New-ADGroupAuditorTab
# ============================================================

function ADG_Find-GroupsBySuffix {
    param([string[]]$Suffixes, [string]$SearchBase = $null)
    $filters = $Suffixes | ForEach-Object { "Name -like '*$($_.Trim())'" }
    $filter = $filters -join ' -or '
    $params = @{ Filter = $filter; Properties = 'Description','whenCreated','DistinguishedName','GroupCategory','GroupScope' }
    if ($SearchBase) { $params.SearchBase = $SearchBase }
    Get-ADGroup @params | Sort-Object Name
}

function ADG_Find-GroupsByName {
    <#
    Procura grupos por nome. Aceita nomes exactos ou wildcards com '*'.
    Ex: "HR_Lisboa", "HR_*", "*_NF"
    #>
    param([string[]]$Names, [string]$SearchBase = $null)
    $filters = $Names | ForEach-Object {
        $n = $_.Trim()
        if ($n -match '\*') { "Name -like '$n'" } else { "Name -eq '$n'" }
    }
    $filter = $filters -join ' -or '
    $params = @{ Filter = $filter; Properties = 'Description','whenCreated','DistinguishedName','GroupCategory','GroupScope' }
    if ($SearchBase) { $params.SearchBase = $SearchBase }
    Get-ADGroup @params | Sort-Object Name
}

function ADG_Expand-GroupTree {
    param(
        [Parameter(Mandatory)] [string]$GroupName,
        [Parameter(Mandatory)] [string]$TargetGroup,
        [string]$PathPrefix = '', [int]$Depth = 0,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [System.Collections.Generic.HashSet[string]]$Seen,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [System.Collections.Generic.List[object]]$Rows,
        [ScriptBlock]$PumpUI = $null
    )
    $currentPath = if ($PathPrefix) { "$PathPrefix > $GroupName" } else { $GroupName }

    if (-not $Seen.Add($GroupName)) {
        $Rows.Add([PSCustomObject]@{
            TargetGroup=$TargetGroup; Path=$currentPath; Depth=$Depth
            MemberType='LoopDetected'; ParentGroup=$GroupName; SamAccountName=''
            DisplayName='(referencia ciclica detectada)'; Email=''; Enabled=$null
        })
        return
    }

    try { $members = @(Get-ADGroupMember -Identity $GroupName -ErrorAction Stop) }
    catch {
        $Rows.Add([PSCustomObject]@{
            TargetGroup=$TargetGroup; Path=$currentPath; Depth=$Depth
            MemberType='Error'; ParentGroup=$GroupName; SamAccountName=''
            DisplayName="(erro: $($_.Exception.Message))"; Email=''; Enabled=$null
        })
        return
    }

    if ($members.Count -eq 0) {
        $Rows.Add([PSCustomObject]@{
            TargetGroup=$TargetGroup; Path=$currentPath; Depth=$Depth
            MemberType='EmptyGroup'; ParentGroup=$GroupName; SamAccountName=''
            DisplayName='(grupo vazio)'; Email=''; Enabled=$null
        })
        return
    }

    $userMembers  = @($members | Where-Object objectClass -eq 'user')
    $groupMembers = @($members | Where-Object objectClass -eq 'group')
    $otherMembers = @($members | Where-Object { $_.objectClass -ne 'user' -and $_.objectClass -ne 'group' })

    if ($PumpUI) { & $PumpUI }
    $userDetails = Get-ADUserDetailsBatch -UserMembers $userMembers -PumpUI $PumpUI
    if ($PumpUI) { & $PumpUI }

    $counter = 0
    foreach ($m in $userMembers) {
        $u = $userDetails[$m.distinguishedName]
        if ($u) {
            $Rows.Add([PSCustomObject]@{
                TargetGroup=$TargetGroup; Path=$currentPath; Depth=$Depth
                MemberType='User'; ParentGroup=$GroupName; SamAccountName=$u.SamAccountName
                DisplayName=$u.DisplayName; Email=$u.EmailAddress; Enabled=$u.Enabled
                Title=$u.Title; Department=$u.Department
            })
        } else {
            $Rows.Add([PSCustomObject]@{
                TargetGroup=$TargetGroup; Path=$currentPath; Depth=$Depth
                MemberType='UserError'; ParentGroup=$GroupName; SamAccountName=$m.SamAccountName
                DisplayName='(erro a ler user)'; Email=''; Enabled=$null
            })
        }
        $counter++
        if ($PumpUI -and ($counter % 25 -eq 0)) { & $PumpUI }
    }

    foreach ($m in $otherMembers) {
        $Rows.Add([PSCustomObject]@{
            TargetGroup=$TargetGroup; Path=$currentPath; Depth=$Depth
            MemberType="Other:$($m.objectClass)"; ParentGroup=$GroupName
            SamAccountName=$m.SamAccountName; DisplayName=$m.Name; Email=''; Enabled=$null
        })
    }

    foreach ($m in $groupMembers) {
        $Rows.Add([PSCustomObject]@{
            TargetGroup=$TargetGroup; Path=$currentPath; Depth=$Depth
            MemberType='NestedGroup'; ParentGroup=$GroupName
            SamAccountName=$m.SamAccountName; DisplayName="(grupo: $($m.Name))"; Email=''; Enabled=$null
        })
        if ($PumpUI) { & $PumpUI }
        ADG_Expand-GroupTree -GroupName $m.SamAccountName -TargetGroup $TargetGroup `
            -PathPrefix $currentPath -Depth ($Depth + 1) -Seen $Seen -Rows $Rows -PumpUI $PumpUI
    }
}

function ADG_Invoke-Audit {
    param(
        [string[]]$Terms,
        [ValidateSet('Suffix','Name')] [string]$SearchMode = 'Suffix',
        [bool]$ActiveOnly,
        [ScriptBlock]$OnProgress,
        [ScriptBlock]$PumpUI = $null
    )
    $groups = if ($SearchMode -eq 'Name') {
        ADG_Find-GroupsByName -Names $Terms
    } else {
        ADG_Find-GroupsBySuffix -Suffixes $Terms
    }
    $allRows = [System.Collections.Generic.List[object]]::new()

    for ($i = 0; $i -lt $groups.Count; $i++) {
        $g = $groups[$i]
        if ($OnProgress) { & $OnProgress $i $groups.Count $g.Name }
        $groupRows = [System.Collections.Generic.List[object]]::new()
        $seen = [System.Collections.Generic.HashSet[string]]::new()
        try {
            ADG_Expand-GroupTree -GroupName $g.Name -TargetGroup $g.Name -Seen $seen -Rows $groupRows -PumpUI $PumpUI
        } catch {
            $groupRows.Add([PSCustomObject]@{
                TargetGroup=$g.Name; Path=$g.Name; Depth=0; MemberType='Error'
                ParentGroup=$g.Name; SamAccountName=''; DisplayName="(erro: $($_.Exception.Message))"
                Email=''; Enabled=$null
            })
        }
        foreach ($r in $groupRows) {
            $r | Add-Member -NotePropertyName GroupDescription -NotePropertyValue $g.Description -Force
            $r | Add-Member -NotePropertyName GroupCategory -NotePropertyValue $g.GroupCategory -Force
            $r | Add-Member -NotePropertyName GroupScope -NotePropertyValue $g.GroupScope -Force
            if ($ActiveOnly -and $r.MemberType -eq 'User' -and -not $r.Enabled) { continue }
            $allRows.Add($r)
        }
    }
    return , $allRows
}

function ADG_Export-ToExcel {
    param([System.Collections.IList]$Results, [string]$OutputPath, [string]$Suffixes)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.ScreenUpdating=$false
    try {
        $wb = $excel.Workbooks.Add()
        while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

        # Resumo
        $ws1 = $wb.Worksheets.Item(1); $ws1.Name = 'Resumo'
        $ws1.Cells.Item(1,1) = 'Auditoria de Grupos AD'
        $ws1.Cells.Item(1,1).Font.Size = 16; $ws1.Cells.Item(1,1).Font.Bold = $true
        $ws1.Range($ws1.Cells.Item(1,1), $ws1.Cells.Item(1,6)).Merge() | Out-Null
        $uniqueTargets = $Results | Select-Object -ExpandProperty TargetGroup -Unique
        $uniqueUsers   = $Results | Where-Object MemberType -eq 'User' | Select-Object -ExpandProperty SamAccountName -Unique
        $ws1.Cells.Item(2,1)='Sufixos:';      $ws1.Cells.Item(2,2)=$Suffixes
        $ws1.Cells.Item(3,1)='Data:';         $ws1.Cells.Item(3,2)=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $ws1.Cells.Item(4,1)='Grupos-alvo:';  $ws1.Cells.Item(4,2)=$uniqueTargets.Count
        $ws1.Cells.Item(5,1)='Users distintos:'; $ws1.Cells.Item(5,2)=$uniqueUsers.Count
        $ws1.Cells.Item(6,1)='Total linhas:'; $ws1.Cells.Item(6,2)=$Results.Count
        for ($r=2; $r -le 6; $r++) { $ws1.Cells.Item($r,1).Font.Bold = $true }

        $summary = $Results | Group-Object TargetGroup | ForEach-Object {
            $first = $_.Group[0]
            $users = @($_.Group | Where-Object MemberType -eq 'User')
            $nested = @($_.Group | Where-Object MemberType -eq 'NestedGroup')
            [PSCustomObject]@{
                Grupo=$_.Name; Descricao=$first.GroupDescription
                Scope="$($first.GroupScope)"; Categoria="$($first.GroupCategory)"
                GruposAninhados=$nested.Count
                UsersTotais=($users | Select-Object -ExpandProperty SamAccountName -Unique).Count
                UsersActivos=($users | Where-Object { $_.Enabled } | Select-Object -ExpandProperty SamAccountName -Unique).Count
                UsersDesativados=($users | Where-Object { $null -ne $_.Enabled -and -not $_.Enabled } | Select-Object -ExpandProperty SamAccountName -Unique).Count
            }
        } | Sort-Object Grupo

        $startRow = 8
        $hdrs = @('Grupo','Descricao','Scope','Categoria','GruposAninhados','UsersTotais','UsersActivos','UsersDesativados')
        for ($c = 0; $c -lt $hdrs.Count; $c++) { $ws1.Cells.Item($startRow, $c+1) = $hdrs[$c] }
        $row = $startRow + 1
        foreach ($s in $summary) {
            $ws1.Cells.Item($row,1)=$s.Grupo; $ws1.Cells.Item($row,2)=$s.Descricao
            $ws1.Cells.Item($row,3)=$s.Scope; $ws1.Cells.Item($row,4)=$s.Categoria
            $ws1.Cells.Item($row,5)=$s.GruposAninhados; $ws1.Cells.Item($row,6)=$s.UsersTotais
            $ws1.Cells.Item($row,7)=$s.UsersActivos; $ws1.Cells.Item($row,8)=$s.UsersDesativados
            $row++
        }
        if ($summary.Count -gt 0) {
            $rng = $ws1.Range($ws1.Cells.Item($startRow,1), $ws1.Cells.Item($row-1, $hdrs.Count))
            $lo = $ws1.ListObjects.Add(1, $rng, $null, 1); $lo.Name='tblResumo'; $lo.TableStyle='TableStyleMedium2'
        }
        $ws1.Columns.AutoFit() | Out-Null
        $ws1.Application.ActiveWindow.SplitRow = $startRow
        $ws1.Application.ActiveWindow.FreezePanes = $true

        # Detalhe
        $ws2 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws1); $ws2.Name = 'Detalhe'
        $dh = @('GrupoAlvo','Descricao','Caminho','Nivel','Tipo','GrupoPai','SamAccount','Nome','Email','Ativo','Titulo','Departamento')
        for ($c = 0; $c -lt $dh.Count; $c++) { $ws2.Cells.Item(1, $c+1) = $dh[$c] }
        $row = 2
        foreach ($r in $Results) {
            $ws2.Cells.Item($row,1)=$r.TargetGroup; $ws2.Cells.Item($row,2)=$r.GroupDescription
            $ws2.Cells.Item($row,3)=$r.Path; $ws2.Cells.Item($row,4)=$r.Depth
            $ws2.Cells.Item($row,5)=$r.MemberType; $ws2.Cells.Item($row,6)=$r.ParentGroup
            $ws2.Cells.Item($row,7)=$r.SamAccountName; $ws2.Cells.Item($row,8)=$r.DisplayName
            $ws2.Cells.Item($row,9)=$r.Email
            $ws2.Cells.Item($row,10)=if ($null -ne $r.Enabled) { if ($r.Enabled) {'Sim'} else {'Nao'} } else { '' }
            $ws2.Cells.Item($row,11)=if ($r.PSObject.Properties.Name -contains 'Title') { $r.Title } else { '' }
            $ws2.Cells.Item($row,12)=if ($r.PSObject.Properties.Name -contains 'Department') { $r.Department } else { '' }
            $row++
        }
        if ($Results.Count -gt 0) {
            $rng = $ws2.Range($ws2.Cells.Item(1,1), $ws2.Cells.Item($row-1, $dh.Count))
            $lo = $ws2.ListObjects.Add(1, $rng, $null, 1); $lo.Name='tblDetalhe'; $lo.TableStyle='TableStyleMedium2'
        }
        $ws2.Columns.AutoFit() | Out-Null
        $ws2.Application.ActiveWindow.SplitRow = 1
        $ws2.Application.ActiveWindow.FreezePanes = $true

        # Por Grupo (hierarquico)
        $ws3 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws2); $ws3.Name = 'Por Grupo'
        $ws3.Cells.Item(1,1) = 'Vista hierarquica: grupo-alvo -> sub-grupos -> users'
        $ws3.Cells.Item(1,1).Font.Bold = $true; $ws3.Cells.Item(1,1).Font.Size = 14
        $ws3.Range($ws3.Cells.Item(1,1), $ws3.Cells.Item(1,7)).Merge() | Out-Null

        $targetBg=0x305496; $targetFg=0xFFFFFF; $nestedBg=0x8EA9DB; $disabledFg=0x808080; $headerBg=0xD9E1F2
        $hh = @('Tipo','Caminho/Nome','SamAccount','Email','Ativo','Titulo','Departamento')
        $row = 3
        $grouped = $Results | Group-Object TargetGroup | Sort-Object Name

        foreach ($grp in $grouped) {
            $firstItem = $grp.Group[0]
            $headerText = "GRUPO: $($grp.Name)"
            if ($firstItem.GroupDescription) { $headerText += "    |    $($firstItem.GroupDescription)" }
            $ws3.Cells.Item($row,1) = $headerText
            $rng = $ws3.Range($ws3.Cells.Item($row,1), $ws3.Cells.Item($row, $hh.Count))
            $rng.Merge() | Out-Null
            $rng.Font.Bold=$true; $rng.Font.Size=11; $rng.Font.Color=Convert-RgbToBgr $targetFg
            $rng.Interior.Color=Convert-RgbToBgr $targetBg; $rng.HorizontalAlignment=-4131
            $row++
            for ($c = 0; $c -lt $hh.Count; $c++) { $ws3.Cells.Item($row, $c+1) = $hh[$c] }
            $hdrRng = $ws3.Range($ws3.Cells.Item($row,1), $ws3.Cells.Item($row, $hh.Count))
            $hdrRng.Font.Bold=$true; $hdrRng.Interior.Color=Convert-RgbToBgr $headerBg; $hdrRng.Borders.LineStyle=1
            $row++
            foreach ($item in $grp.Group) {
                $indent = '  ' * $item.Depth
                $ws3.Cells.Item($row,1)=$item.MemberType
                $ws3.Cells.Item($row,2)="$indent$($item.Path)"
                $ws3.Cells.Item($row,3)=$item.SamAccountName; $ws3.Cells.Item($row,4)=$item.Email
                $ws3.Cells.Item($row,5)=if ($null -ne $item.Enabled) { if ($item.Enabled) {'Sim'} else {'Nao'} } else { '' }
                $ws3.Cells.Item($row,6)=if ($item.PSObject.Properties.Name -contains 'Title') { $item.Title } else { '' }
                $ws3.Cells.Item($row,7)=if ($item.PSObject.Properties.Name -contains 'Department') { $item.Department } else { '' }
                $rowRng = $ws3.Range($ws3.Cells.Item($row,1), $ws3.Cells.Item($row, $hh.Count))
                $rowRng.Borders.LineStyle = 1
                switch ($item.MemberType) {
                    'NestedGroup' {
                        $rowRng.Interior.Color=Convert-RgbToBgr $nestedBg; $rowRng.Font.Bold=$true
                        $ws3.Cells.Item($row,2)="$indent[ $($item.DisplayName) ]"
                    }
                    'EmptyGroup' { $rowRng.Font.Italic=$true; $rowRng.Font.Color=Convert-RgbToBgr $disabledFg }
                    'Error' { $rowRng.Font.Color=Convert-RgbToBgr 0xC00000; $rowRng.Font.Italic=$true }
                    'LoopDetected' { $rowRng.Font.Color=Convert-RgbToBgr 0xC00000; $rowRng.Font.Italic=$true }
                    'User' {
                        if ($null -ne $item.Enabled -and -not $item.Enabled) {
                            $rowRng.Font.Color=Convert-RgbToBgr $disabledFg; $rowRng.Font.Italic=$true
                        }
                    }
                }
                $row++
            }
            $row++
        }
        $ws3.Columns.AutoFit() | Out-Null
        $ws3.Application.ActiveWindow.SplitRow = 1
        $ws3.Application.ActiveWindow.FreezePanes = $true
        $ws1.Activate()

        if ([System.IO.Path]::GetExtension($OutputPath) -ne '.xlsx') {
            $OutputPath = [System.IO.Path]::ChangeExtension($OutputPath, '.xlsx')
        }
        $wb.SaveAs($OutputPath, 51)
        $wb.Close($false)
    }
    finally {
        $excel.ScreenUpdating=$true; $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect() | Out-Null; [GC]::WaitForPendingFinalizers() | Out-Null
    }
}

# ============================================================
# FUNCAO EXPORTADA
# ============================================================

function New-ADGroupAuditorTab {
    $script:ADG_LastResults = $null

    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = 'Group Auditor'

    $panelTop = New-Object System.Windows.Forms.Panel
    $panelTop.Dock = 'Top'; $panelTop.Height = 160; $panelTop.Padding = '12,12,12,12'
    $tab.Controls.Add($panelTop)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = 'Auditor de Grupos AD (por sufixo)'
    $lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $script:PaletteBlue.Dark
    $lblTitle.Location = New-Object System.Drawing.Point(12, 4)
    $lblTitle.Size = New-Object System.Drawing.Size(600, 26)
    $panelTop.Controls.Add($lblTitle)

    if (-not $script:ADAvailable) {
        $lblWarn = New-Object System.Windows.Forms.Label
        $lblWarn.Text = 'Modulo ActiveDirectory indisponivel. Esta ferramenta so funciona em maquinas com RSAT (tipicamente um DC ou com o modulo instalado).'
        $lblWarn.ForeColor = [System.Drawing.Color]::DarkOrange
        $lblWarn.Location = New-Object System.Drawing.Point(12, 36)
        $lblWarn.Size = New-Object System.Drawing.Size(1050, 40)
        $panelTop.Controls.Add($lblWarn)
        return $tab
    }

    # --- Modo de pesquisa (radio) ---
    $lblMode = New-Object System.Windows.Forms.Label
    $lblMode.Text = 'Modo:'
    $lblMode.Location = New-Object System.Drawing.Point(12, 40); $lblMode.Size = New-Object System.Drawing.Size(50, 22)
    $panelTop.Controls.Add($lblMode)

    $rdoSuffix = New-Object System.Windows.Forms.RadioButton
    $rdoSuffix.Text = 'Por sufixo'
    $rdoSuffix.Checked = $true
    $rdoSuffix.Location = New-Object System.Drawing.Point(66, 40); $rdoSuffix.Size = New-Object System.Drawing.Size(110, 22)
    $panelTop.Controls.Add($rdoSuffix)

    $rdoName = New-Object System.Windows.Forms.RadioButton
    $rdoName.Text = 'Por nome do grupo'
    $rdoName.Location = New-Object System.Drawing.Point(180, 40); $rdoName.Size = New-Object System.Drawing.Size(160, 22)
    $panelTop.Controls.Add($rdoName)

    # --- Input (partilhado, label muda consoante modo) ---
    $lblSuffix = New-Object System.Windows.Forms.Label
    $lblSuffix.Text = 'Sufixos (separados por virgula):'
    $lblSuffix.Location = New-Object System.Drawing.Point(12, 72); $lblSuffix.Size = New-Object System.Drawing.Size(220, 22)
    $panelTop.Controls.Add($lblSuffix)

    $txtSuffix = New-Object System.Windows.Forms.TextBox
    $txtSuffix.Text = 'NF,NR'
    $txtSuffix.Location = New-Object System.Drawing.Point(238, 70); $txtSuffix.Size = New-Object System.Drawing.Size(300, 24)
    $panelTop.Controls.Add($txtSuffix)

    $lblEx = New-Object System.Windows.Forms.Label
    $lblEx.Text = '(ex: NF,NR => todos os grupos terminados em "NF" ou "NR")'
    $lblEx.Location = New-Object System.Drawing.Point(550, 74); $lblEx.Size = New-Object System.Drawing.Size(500, 22)
    $lblEx.ForeColor = [System.Drawing.Color]::Gray
    $panelTop.Controls.Add($lblEx)

    # Trocar label/exemplo ao alternar modo
    $rdoSuffix.Add_CheckedChanged({
        if ($rdoSuffix.Checked) {
            $lblSuffix.Text = 'Sufixos (separados por virgula):'
            $lblEx.Text = '(ex: NF,NR => todos os grupos terminados em "NF" ou "NR")'
            $txtSuffix.Text = 'NF,NR'
        }
    }.GetNewClosure())
    $rdoName.Add_CheckedChanged({
        if ($rdoName.Checked) {
            $lblSuffix.Text = 'Nome(s) do grupo (virgulas; aceita *):'
            $lblEx.Text = '(ex: HR_Lisboa ou HR_* ou *_Admins)'
            $txtSuffix.Text = ''
            $txtSuffix.Focus() | Out-Null
        }
    }.GetNewClosure())

    $chkActive = New-Object System.Windows.Forms.CheckBox
    $chkActive.Text = 'Apenas users activos (Enabled = True)'
    $chkActive.Location = New-Object System.Drawing.Point(238, 100)
    $chkActive.Size = New-Object System.Drawing.Size(300, 22)
    $panelTop.Controls.Add($chkActive)

    $btnSearch = New-StyledButton -Text 'Pesquisar' -X 238 -Y 130 -BackColor $script:PaletteBlue.Dark -ForeColor 'White' -Bold $true
    $panelTop.Controls.Add($btnSearch)
    $btnExport = New-StyledButton -Text 'Exportar Excel' -X 388 -Y 130 -BackColor $script:PaletteGreen.Dark -ForeColor 'White' -Bold $true
    $btnExport.Enabled = $false; $panelTop.Controls.Add($btnExport)
    $btnClear = New-StyledButton -Text 'Limpar' -X 538 -Y 130 -Width 100
    $panelTop.Controls.Add($btnClear)

    $panelStatus = New-Object System.Windows.Forms.Panel
    $panelStatus.Dock = 'Bottom'; $panelStatus.Height = 52; $panelStatus.Padding = '12,6,12,6'
    $tab.Controls.Add($panelStatus)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = 'Pronto.'
    $lblStatus.Location = New-Object System.Drawing.Point(12, 6)
    $lblStatus.Size = New-Object System.Drawing.Size(1100, 20)
    $lblStatus.Anchor = 'Top, Left, Right'
    $panelStatus.Controls.Add($lblStatus)

    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Location = New-Object System.Drawing.Point(12, 28)
    $progress.Size = New-Object System.Drawing.Size(1100, 16)
    $progress.Anchor = 'Top, Left, Right'
    $progress.Style = 'Continuous'
    $panelStatus.Controls.Add($progress)

    $grid = New-StyledDataGridView
    $tab.Controls.Add($grid); $grid.BringToFront()

    $btnClear.Add_Click({
        $grid.DataSource = $null; $script:ADG_LastResults = $null
        $btnExport.Enabled = $false; $lblStatus.Text = 'Pronto.'; $progress.Value = 0
    }.GetNewClosure())

    $btnSearch.Add_Click({
        $raw = $txtSuffix.Text.Trim()
        $mode = if ($rdoName.Checked) { 'Name' } else { 'Suffix' }
        $msg = if ($mode -eq 'Name') { 'Indique pelo menos um nome de grupo.' } else { 'Indique pelo menos um sufixo.' }
        if (-not $raw) { [System.Windows.Forms.MessageBox]::Show($msg, 'Aviso', 'OK', 'Warning') | Out-Null; return }
        $terms = $raw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        if ($terms.Count -eq 0) { return }

        $btnSearch.Enabled=$false; $btnExport.Enabled=$false; $btnClear.Enabled=$false
        $grid.DataSource=$null; $progress.Value=0
        $lblStatus.Text='A procurar grupos...'
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $progressCb = {
                param($i, $total, $groupName)
                $pct = if ($total -gt 0) { [int](($i / $total) * 100) } else { 0 }
                $progress.Value = [Math]::Min(100, [Math]::Max(0, $pct))
                $lblStatus.Text = "A expandir ($($i+1)/$total): $groupName"
                [System.Windows.Forms.Application]::DoEvents()
            }.GetNewClosure()
            $pumpCb = { [System.Windows.Forms.Application]::DoEvents() }

            $results = ADG_Invoke-Audit -Terms $terms -SearchMode $mode -ActiveOnly $chkActive.Checked -OnProgress $progressCb -PumpUI $pumpCb
            $script:ADG_LastResults = $results

            if ($results.Count -eq 0) {
                $lblStatus.Text = 'Nenhum grupo encontrado.'; $progress.Value = 0
            } else {
                $dt = New-Object System.Data.DataTable
                $cols = 'GrupoAlvo','Descricao','Caminho','Nivel','Tipo','GrupoPai','SamAccount','Nome','Email','Ativo','Titulo','Departamento'
                foreach ($c in $cols) { [void]$dt.Columns.Add($c) }
                foreach ($r in $results) {
                    $row = $dt.NewRow()
                    $row['GrupoAlvo']=[string]$r.TargetGroup; $row['Descricao']=[string]$r.GroupDescription
                    $row['Caminho']=[string]$r.Path; $row['Nivel']=[string]$r.Depth
                    $row['Tipo']=[string]$r.MemberType; $row['GrupoPai']=[string]$r.ParentGroup
                    $row['SamAccount']=[string]$r.SamAccountName; $row['Nome']=[string]$r.DisplayName
                    $row['Email']=[string]$r.Email
                    $row['Ativo']=if ($null -ne $r.Enabled) { if ($r.Enabled) {'Sim'} else {'Nao'} } else { '' }
                    $row['Titulo']=if ($r.PSObject.Properties.Name -contains 'Title') { [string]$r.Title } else { '' }
                    $row['Departamento']=if ($r.PSObject.Properties.Name -contains 'Department') { [string]$r.Department } else { '' }
                    $dt.Rows.Add($row)
                }
                $grid.DataSource = $dt; $progress.Value = 100
                $targets = ($results | Select-Object -ExpandProperty TargetGroup -Unique).Count
                $users = ($results | Where-Object MemberType -eq 'User' | Select-Object -ExpandProperty SamAccountName -Unique).Count
                $nested = ($results | Where-Object MemberType -eq 'NestedGroup').Count
                $lblStatus.Text = "Concluido: $targets grupos | $nested aninhamentos | $users users | $($results.Count) linhas."
                $btnExport.Enabled = $true
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Erro:`n$($_.Exception.Message)", 'Erro', 'OK', 'Error') | Out-Null
            $lblStatus.Text = "Erro: $($_.Exception.Message)"
        }
        finally {
            $btnSearch.Enabled = $true; $btnClear.Enabled = $true
        }
    }.GetNewClosure())

    $btnExport.Add_Click({
        if (-not $script:ADG_LastResults -or $script:ADG_LastResults.Count -eq 0) { return }
        $sf = New-Object System.Windows.Forms.SaveFileDialog
        $sf.Filter = 'Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv'
        $sf.FileName = "AuditGruposAD_$((Get-Date).ToString('yyyyMMdd_HHmmss')).xlsx"
        if ($sf.ShowDialog() -ne 'OK') { return }

        $out = $sf.FileName
        $btnExport.Enabled = $false; $btnSearch.Enabled = $false
        $lblStatus.Text = 'A exportar...'
        [System.Windows.Forms.Application]::DoEvents()
        try {
            if ($sf.FilterIndex -eq 1) {
                if (-not (Test-ExcelAvailable)) {
                    $res = [System.Windows.Forms.MessageBox]::Show('Excel nao disponivel. Exportar CSV?', 'Aviso', 'YesNo', 'Question')
                    if ($res -eq 'Yes') { $out = Export-ResultsToCsv -Results $script:ADG_LastResults -OutputPath $out }
                    else { $lblStatus.Text = 'Cancelado.'; return }
                } else {
                    ADG_Export-ToExcel -Results $script:ADG_LastResults -OutputPath $out -Suffixes $txtSuffix.Text
                }
            } else {
                $out = Export-ResultsToCsv -Results $script:ADG_LastResults -OutputPath $out
            }
            $lblStatus.Text = "Exportado: $out"
            $res = [System.Windows.Forms.MessageBox]::Show("Guardado:`n$out`n`nAbrir?", 'OK', 'YesNo', 'Information')
            if ($res -eq 'Yes') { Start-Process -FilePath $out }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Erro:`n$($_.Exception.Message)", 'Erro', 'OK', 'Error') | Out-Null
        }
        finally {
            $btnExport.Enabled = ($script:ADG_LastResults.Count -gt 0); $btnSearch.Enabled = $true
        }
    }.GetNewClosure())

    return $tab
}
