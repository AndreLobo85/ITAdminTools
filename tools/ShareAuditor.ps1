# ============================================================
# ShareAuditor.ps1 - Ferramenta: auditoria de shares SMB/NTFS
# Exporta: New-ShareAuditorTab  (funcao chamada pelo toolkit)
# ============================================================

# ----- Funcoes de dominio (prefixo SA_) -----

function SA_Resolve-ShareToLocal {
    param([string]$Unc)
    if ($Unc -notmatch '^\\\\') {
        return @{ Server = $env:COMPUTERNAME; ShareName = $null; LocalPath = $Unc }
    }
    $parts = $Unc.TrimStart('\') -split '\\', 3
    $server = $parts[0]; $share = $parts[1]
    $sub = if ($parts.Count -ge 3) { $parts[2] } else { '' }
    try {
        $smb = Get-WmiObject -Class Win32_Share -ComputerName $server -Filter "Name='$share'" -ErrorAction Stop
        $local = if ($sub) { Join-Path $smb.Path $sub } else { $smb.Path }
        return @{ Server = $server; ShareName = $share; LocalPath = $local }
    } catch {
        return @{ Server = $server; ShareName = $share; LocalPath = $Unc }
    }
}

function SA_Get-SharePermission {
    param([string]$Server, [string]$ShareName)
    if (-not $ShareName) { return @() }
    try {
        $acl = Get-WmiObject -Class Win32_LogicalShareSecuritySetting -ComputerName $Server -Filter "Name='$ShareName'" -ErrorAction Stop
        $sd  = $acl.GetSecurityDescriptor().Descriptor
        return $sd.DACL | ForEach-Object {
            $trustee = if ($_.Trustee.Domain) { "$($_.Trustee.Domain)\$($_.Trustee.Name)" } else { $_.Trustee.Name }
            [PSCustomObject]@{
                Source = 'Share'; Identity = $trustee; AccessMask = $_.AccessMask
                AceType = if ($_.AceType -eq 0) { 'Allow' } else { 'Deny' }
                IsInherited = $false; Path = "\\$Server\$ShareName"
            }
        }
    } catch { @() }
}

function SA_Get-NtfsPermission {
    param([string]$Path, [bool]$ExplicitOnly = $false)
    try {
        $access = (Get-Acl -LiteralPath $Path -ErrorAction Stop).Access
        if ($ExplicitOnly) { $access = $access | Where-Object { -not $_.IsInherited } }
        $access | ForEach-Object {
            [PSCustomObject]@{
                Source = 'NTFS'; Identity = $_.IdentityReference.Value
                AccessMask = $_.FileSystemRights; AceType = $_.AccessControlType
                IsInherited = $_.IsInherited; Path = $Path
            }
        }
    } catch { @() }
}

function SA_Expand-ADGroupMember {
    param(
        [string]$Identity,
        [System.Collections.Generic.HashSet[string]]$Seen,
        [ScriptBlock]$PumpUI = $null
    )
    if ($Identity -match '^(NT AUTHORITY|BUILTIN|CREATOR|Everyone|S-1-)') {
        return @([PSCustomObject]@{ Group = $Identity; Member = '(principal local/builtin)'; DisplayName = ''; Enabled = $null; Email = ''; Type = 'Builtin' })
    }
    if (-not $script:ADAvailable) {
        return @([PSCustomObject]@{ Group = $Identity; Member = '(AD nao disponivel)'; DisplayName = ''; Enabled = $null; Email = ''; Type = 'NoAD' })
    }
    $name = ($Identity -split '\\')[-1]
    try {
        $obj = Get-ADObject -Filter "SamAccountName -eq '$name'" -Properties objectClass -ErrorAction Stop | Select-Object -First 1
        if (-not $obj) {
            return @([PSCustomObject]@{ Group = $Identity; Member = '(nao encontrado no AD)'; DisplayName = ''; Enabled = $null; Email = ''; Type = 'Unknown' })
        }
        if ($obj.objectClass -eq 'user') {
            $u = Get-ADUser $obj -Properties DisplayName, Enabled, EmailAddress
            return @([PSCustomObject]@{ Group = '(direto)'; Member = $u.SamAccountName; DisplayName = $u.DisplayName; Enabled = $u.Enabled; Email = $u.EmailAddress; Type = 'User' })
        }
        if ($obj.objectClass -eq 'group') {
            if (-not $Seen.Add($name)) { return @() }
            if ($PumpUI) { & $PumpUI }
            $members = @(Get-ADGroupMember -Identity $name -Recursive -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' })
            if ($members.Count -eq 0) {
                return @([PSCustomObject]@{ Group = $name; Member = '(grupo vazio)'; DisplayName = ''; Enabled = $null; Email = ''; Type = 'EmptyGroup' })
            }
            $userDetails = Get-ADUserDetailsBatch -UserMembers $members -PumpUI $PumpUI
            $out = New-Object System.Collections.Generic.List[object]
            $counter = 0
            foreach ($m in $members) {
                $u = $userDetails[$m.distinguishedName]
                if ($u) {
                    $out.Add([PSCustomObject]@{ Group = $name; Member = $u.SamAccountName; DisplayName = $u.DisplayName; Enabled = $u.Enabled; Email = $u.EmailAddress; Type = 'User' })
                } else {
                    $out.Add([PSCustomObject]@{ Group = $name; Member = $m.SamAccountName; DisplayName = '(erro a ler user)'; Enabled = $null; Email = ''; Type = 'UserError' })
                }
                $counter++
                if ($PumpUI -and ($counter % 50 -eq 0)) { & $PumpUI }
            }
            return , $out.ToArray()
        }
    } catch {
        return @([PSCustomObject]@{ Group = $Identity; Member = "(erro: $($_.Exception.Message))"; DisplayName = ''; Enabled = $null; Email = ''; Type = 'Error' })
    }
}

function SA_Get-FolderTree {
    param([string]$Root, [int]$MaxDepth)
    $list = [System.Collections.Generic.List[string]]::new()
    $list.Add($Root)
    if ($MaxDepth -le 0) { return $list }
    try {
        Get-ChildItem -LiteralPath $Root -Directory -Recurse -Depth $MaxDepth -Force -ErrorAction SilentlyContinue |
            ForEach-Object { $list.Add($_.FullName) }
    } catch {}
    return $list
}

function SA_Invoke-Audit {
    param(
        [string]$SharePath, [bool]$Recurse, [int]$Depth, [bool]$OnlyExplicit,
        [ScriptBlock]$OnProgress, [ScriptBlock]$PumpUI = $null
    )
    $info = SA_Resolve-ShareToLocal -Unc $SharePath
    $folders = if ($Recurse) { SA_Get-FolderTree -Root $SharePath -MaxDepth $Depth } else { @($SharePath) }
    $results = New-Object System.Collections.Generic.List[object]

    for ($i = 0; $i -lt $folders.Count; $i++) {
        $folder = $folders[$i]
        $isRoot = ($i -eq 0)
        if ($OnProgress) { & $OnProgress $i $folders.Count $folder }

        $sharePerms = if ($isRoot) { SA_Get-SharePermission -Server $info.Server -ShareName $info.ShareName } else { @() }
        $ntfsAll    = SA_Get-NtfsPermission -Path $folder -ExplicitOnly:$false
        $ntfsPerms  = if (!$isRoot -and $OnlyExplicit) { @($ntfsAll | Where-Object { -not $_.IsInherited }) } else { @($ntfsAll) }

        $allPerms = @($sharePerms) + @($ntfsPerms)
        if (-not $allPerms -or $allPerms.Count -eq 0) {
            if (!$isRoot -and $OnlyExplicit -and $ntfsAll.Count -gt 0) {
                $results.Add([PSCustomObject]@{
                    Folder = $folder; IsRoot = $false; Principal = '(apenas herda da pasta-mae)'
                    Permissions = '(sem ACEs explicitas nesta pasta)'; Inherited = 'Sim'
                    Group = ''; Member = ''; DisplayName = ''; Enabled = $null; Email = ''; Type = 'InheritedOnly'
                })
            }
            continue
        }

        $identities = $allPerms | Select-Object -ExpandProperty Identity -Unique
        foreach ($id in $identities) {
            if ($PumpUI) { & $PumpUI }
            $seen = New-Object 'System.Collections.Generic.HashSet[string]'
            $expanded = SA_Expand-ADGroupMember -Identity $id -Seen $seen -PumpUI $PumpUI

            $principalPerms = @($allPerms | Where-Object Identity -eq $id)
            $perms = ($principalPerms | ForEach-Object {
                $inh = if ($_.IsInherited) { 'Herd' } else { 'Expl' }
                "$($_.Source):$($_.AceType):$($_.AccessMask):$inh"
            }) -join ' | '

            $inheritedCount = @($principalPerms | Where-Object { $_.IsInherited }).Count
            $totalCount = $principalPerms.Count
            $inheritedValue = if ($inheritedCount -eq 0) { 'Nao' } elseif ($inheritedCount -eq $totalCount) { 'Sim' } else { 'Parcial' }

            foreach ($e in $expanded) {
                $results.Add([PSCustomObject]@{
                    Folder = $folder; IsRoot = $isRoot; Principal = $id; Permissions = $perms
                    Inherited = $inheritedValue; Group = $e.Group; Member = $e.Member
                    DisplayName = $e.DisplayName; Enabled = $e.Enabled; Email = $e.Email; Type = $e.Type
                })
            }
        }
    }
    return , $results
}

function SA_Export-ToExcel {
    param([System.Collections.IList]$Results, [string]$OutputPath, [string]$SharePath)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.ScreenUpdating = $false
    try {
        $wb = $excel.Workbooks.Add()
        while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

        # ---------- Resumo ----------
        $ws1 = $wb.Worksheets.Item(1); $ws1.Name = 'Resumo'
        $ws1.Cells.Item(1,1) = 'Auditoria de Share'
        $ws1.Cells.Item(1,1).Font.Size = 16; $ws1.Cells.Item(1,1).Font.Bold = $true
        $ws1.Range($ws1.Cells.Item(1,1), $ws1.Cells.Item(1,4)).Merge() | Out-Null
        $ws1.Cells.Item(2,1) = 'Share auditado:';  $ws1.Cells.Item(2,2) = $SharePath
        $ws1.Cells.Item(3,1) = 'Data:';            $ws1.Cells.Item(3,2) = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $ws1.Cells.Item(4,1) = 'Total de linhas:'; $ws1.Cells.Item(4,2) = $Results.Count
        $ws1.Cells.Item(5,1) = 'Pastas:';          $ws1.Cells.Item(5,2) = ($Results | Select-Object -ExpandProperty Folder -Unique).Count
        $ws1.Cells.Item(6,1) = 'Users distintos:'; $ws1.Cells.Item(6,2) = ($Results | Where-Object Type -eq 'User' | Select-Object -ExpandProperty Member -Unique).Count
        for ($r = 2; $r -le 6; $r++) { $ws1.Cells.Item($r,1).Font.Bold = $true }

        $summary = $Results | Group-Object Folder, Principal | ForEach-Object {
            $first = $_.Group[0]
            [PSCustomObject]@{
                Pasta = $first.Folder; Raiz = if ($first.IsRoot) {'Sim'} else {''}
                Principal = $first.Principal; Herdada = $first.Inherited
                Permissoes = $first.Permissions; NumMembros = ($_.Group | Where-Object Type -eq 'User').Count
            }
        }
        $startRow = 8
        $headers = @('Pasta','Raiz','Principal','Herdada','Permissoes','NumMembros')
        for ($c = 0; $c -lt $headers.Count; $c++) { $ws1.Cells.Item($startRow, $c+1) = $headers[$c] }
        $row = $startRow + 1
        foreach ($s in $summary) {
            $ws1.Cells.Item($row,1)=$s.Pasta; $ws1.Cells.Item($row,2)=$s.Raiz
            $ws1.Cells.Item($row,3)=$s.Principal; $ws1.Cells.Item($row,4)=$s.Herdada
            $ws1.Cells.Item($row,5)=$s.Permissoes; $ws1.Cells.Item($row,6)=$s.NumMembros
            $row++
        }
        if ($summary.Count -gt 0) {
            $rng = $ws1.Range($ws1.Cells.Item($startRow,1), $ws1.Cells.Item($row-1, $headers.Count))
            $lo = $ws1.ListObjects.Add(1, $rng, $null, 1); $lo.Name = 'tblResumo'; $lo.TableStyle = 'TableStyleMedium2'
        }
        $ws1.Columns.AutoFit() | Out-Null
        $ws1.Range('A1').Select() | Out-Null
        $ws1.Application.ActiveWindow.SplitRow = $startRow
        $ws1.Application.ActiveWindow.FreezePanes = $true

        # ---------- Detalhe ----------
        $ws2 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws1); $ws2.Name = 'Detalhe'
        $dh = @('Pasta','Raiz','Principal','Herdada','Grupo','Membro','Nome','Ativo','Email','Tipo','Permissoes')
        for ($c = 0; $c -lt $dh.Count; $c++) { $ws2.Cells.Item(1, $c+1) = $dh[$c] }
        $row = 2
        foreach ($r in $Results) {
            $ws2.Cells.Item($row,1)=$r.Folder
            $ws2.Cells.Item($row,2)=if ($r.IsRoot) {'Sim'} else {''}
            $ws2.Cells.Item($row,3)=$r.Principal; $ws2.Cells.Item($row,4)=$r.Inherited
            $ws2.Cells.Item($row,5)=$r.Group; $ws2.Cells.Item($row,6)=$r.Member
            $ws2.Cells.Item($row,7)=$r.DisplayName
            $ws2.Cells.Item($row,8)=if ($null -ne $r.Enabled) { if ($r.Enabled) {'Sim'} else {'Nao'} } else { '' }
            $ws2.Cells.Item($row,9)=$r.Email; $ws2.Cells.Item($row,10)=$r.Type
            $ws2.Cells.Item($row,11)=$r.Permissions
            $row++
        }
        if ($Results.Count -gt 0) {
            $rng = $ws2.Range($ws2.Cells.Item(1,1), $ws2.Cells.Item($row-1, $dh.Count))
            $lo = $ws2.ListObjects.Add(1, $rng, $null, 1); $lo.Name = 'tblDetalhe'; $lo.TableStyle = 'TableStyleMedium2'
        }
        $ws2.Columns.AutoFit() | Out-Null
        $ws2.Application.ActiveWindow.SplitRow = 1
        $ws2.Application.ActiveWindow.FreezePanes = $true

        # ---------- Por Pasta (hierarquico) ----------
        $ws3 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws2); $ws3.Name = 'Por Pasta'
        $ws3.Cells.Item(1,1) = 'Vista hierarquica: pasta -> principal -> membros'
        $ws3.Cells.Item(1,1).Font.Bold = $true; $ws3.Cells.Item(1,1).Font.Size = 14
        $ws3.Range($ws3.Cells.Item(1,1), $ws3.Cells.Item(1,8)).Merge() | Out-Null

        $folderBg=0x4472C4; $folderFg=0xFFFFFF; $principalBg=0xE7EFF8; $headerBg=0xD9E1F2
        $hh = @('Principal','Herdada','Permissoes','Grupo','Membro','Nome','Ativo','Email')
        $row = 3; $grouped = $Results | Group-Object Folder

        foreach ($grp in $grouped) {
            $ws3.Cells.Item($row,1) = "PASTA: $($grp.Name)"
            $rng = $ws3.Range($ws3.Cells.Item($row,1), $ws3.Cells.Item($row,8))
            $rng.Merge() | Out-Null
            $rng.Font.Bold=$true; $rng.Font.Color=Convert-RgbToBgr $folderFg
            $rng.Interior.Color=Convert-RgbToBgr $folderBg; $rng.HorizontalAlignment=-4131
            $row++
            for ($c = 0; $c -lt $hh.Count; $c++) { $ws3.Cells.Item($row, $c+1) = $hh[$c] }
            $hdrRng = $ws3.Range($ws3.Cells.Item($row,1), $ws3.Cells.Item($row, $hh.Count))
            $hdrRng.Font.Bold = $true; $hdrRng.Interior.Color = Convert-RgbToBgr $headerBg; $hdrRng.Borders.LineStyle = 1
            $row++
            $byPrincipal = $grp.Group | Group-Object Principal
            foreach ($p in $byPrincipal) {
                $fr = $p.Group[0]
                $ws3.Cells.Item($row,1)=$fr.Principal; $ws3.Cells.Item($row,2)=$fr.Inherited
                $ws3.Cells.Item($row,3)=$fr.Permissions
                $pRng = $ws3.Range($ws3.Cells.Item($row,1), $ws3.Cells.Item($row,8))
                $pRng.Interior.Color = Convert-RgbToBgr $principalBg
                $pRng.Font.Bold = $true; $pRng.Borders.LineStyle = 1
                if ($fr.Inherited -eq 'Sim') {
                    $ws3.Cells.Item($row,2).Font.Color = Convert-RgbToBgr 0x808080
                    $ws3.Cells.Item($row,2).Font.Italic = $true
                } elseif ($fr.Inherited -eq 'Parcial') {
                    $ws3.Cells.Item($row,2).Font.Color = Convert-RgbToBgr 0xB7950B
                }
                $row++
                foreach ($m in $p.Group) {
                    $ws3.Cells.Item($row,4)=$m.Group; $ws3.Cells.Item($row,5)=$m.Member
                    $ws3.Cells.Item($row,6)=$m.DisplayName
                    $ws3.Cells.Item($row,7)=if ($null -ne $m.Enabled) { if ($m.Enabled) {'Sim'} else {'Nao'} } else { '' }
                    $ws3.Cells.Item($row,8)=$m.Email
                    $mRng = $ws3.Range($ws3.Cells.Item($row,4), $ws3.Cells.Item($row,8))
                    $mRng.Borders.LineStyle = 1
                    if ($null -ne $m.Enabled -and -not $m.Enabled) {
                        $mRng.Font.Color = Convert-RgbToBgr 0x808080; $mRng.Font.Italic = $true
                    }
                    $row++
                }
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
        $excel.ScreenUpdating = $true; $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect() | Out-Null; [GC]::WaitForPendingFinalizers() | Out-Null
    }
}

# ============================================================
# FUNCAO EXPORTADA: constroi e devolve o TabPage
# ============================================================

function New-ShareAuditorTab {
    # Estado isolado por tab (captured via closures em $script: com prefixo unico)
    $script:SA_LastResults = $null

    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = 'Share Auditor'
    $tab.Padding = New-Object System.Windows.Forms.Padding(0)

    # --- Painel topo ---
    $panelTop = New-Object System.Windows.Forms.Panel
    $panelTop.Dock = 'Top'; $panelTop.Height = 160
    $panelTop.Padding = '12,12,12,12'
    $tab.Controls.Add($panelTop)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = 'Auditor de Permissoes de Share'
    $lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $script:PaletteBlue.Mid
    $lblTitle.Location = New-Object System.Drawing.Point(12, 4)
    $lblTitle.Size = New-Object System.Drawing.Size(600, 26)
    $panelTop.Controls.Add($lblTitle)

    $lblPath = New-Object System.Windows.Forms.Label
    $lblPath.Text = 'Caminho do share:'
    $lblPath.Location = New-Object System.Drawing.Point(12, 36)
    $lblPath.Size = New-Object System.Drawing.Size(120, 22)
    $panelTop.Controls.Add($lblPath)

    $txtPath = New-Object System.Windows.Forms.TextBox
    $txtPath.Location = New-Object System.Drawing.Point(140, 34)
    $txtPath.Size = New-Object System.Drawing.Size(780, 24)
    $txtPath.Anchor = 'Top, Left, Right'
    $panelTop.Controls.Add($txtPath)

    $btnBrowse = New-StyledButton -Text 'Procurar...' -X 928 -Y 33 -Width 130 -Height 26
    $btnBrowse.Anchor = 'Top, Right'
    $panelTop.Controls.Add($btnBrowse)

    $chkRecurse = New-Object System.Windows.Forms.CheckBox
    $chkRecurse.Text = 'Incluir subpastas'; $chkRecurse.Checked = $true
    $chkRecurse.Location = New-Object System.Drawing.Point(140, 68)
    $chkRecurse.Size = New-Object System.Drawing.Size(160, 22)
    $panelTop.Controls.Add($chkRecurse)

    $lblDepth = New-Object System.Windows.Forms.Label
    $lblDepth.Text = 'Profundidade:'
    $lblDepth.Location = New-Object System.Drawing.Point(310, 70); $lblDepth.Size = New-Object System.Drawing.Size(90, 22)
    $panelTop.Controls.Add($lblDepth)

    $numDepth = New-Object System.Windows.Forms.NumericUpDown
    $numDepth.Minimum = 0; $numDepth.Maximum = 20; $numDepth.Value = 3
    $numDepth.Location = New-Object System.Drawing.Point(400, 68)
    $numDepth.Size = New-Object System.Drawing.Size(60, 24)
    $panelTop.Controls.Add($numDepth)

    $chkOnlyExpl = New-Object System.Windows.Forms.CheckBox
    $chkOnlyExpl.Text = 'Filtrar ACEs herdadas nas subpastas (por defeito mostra tudo)'
    $chkOnlyExpl.Checked = $false
    $chkOnlyExpl.Location = New-Object System.Drawing.Point(480, 68)
    $chkOnlyExpl.Size = New-Object System.Drawing.Size(480, 22)
    $panelTop.Controls.Add($chkOnlyExpl)

    $btnAudit = New-StyledButton -Text 'Auditar' -X 140 -Y 100 -BackColor $script:PaletteBlue.Mid -ForeColor 'White' -Bold $true
    $panelTop.Controls.Add($btnAudit)
    $btnExport = New-StyledButton -Text 'Exportar Excel' -X 290 -Y 100 -BackColor $script:PaletteGreen.Dark -ForeColor 'White' -Bold $true
    $btnExport.Enabled = $false
    $panelTop.Controls.Add($btnExport)
    $btnClear = New-StyledButton -Text 'Limpar' -X 440 -Y 100 -Width 100
    $panelTop.Controls.Add($btnClear)

    # --- Painel status ---
    $panelStatus = New-Object System.Windows.Forms.Panel
    $panelStatus.Dock = 'Bottom'; $panelStatus.Height = 52
    $panelStatus.Padding = '12,6,12,6'
    $tab.Controls.Add($panelStatus)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = 'Pronto.'
    $lblStatus.Location = New-Object System.Drawing.Point(12, 6)
    $lblStatus.Size = New-Object System.Drawing.Size(1060, 20)
    $lblStatus.Anchor = 'Top, Left, Right'
    $panelStatus.Controls.Add($lblStatus)

    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Location = New-Object System.Drawing.Point(12, 28)
    $progress.Size = New-Object System.Drawing.Size(1060, 16)
    $progress.Anchor = 'Top, Left, Right'
    $progress.Style = 'Continuous'
    $panelStatus.Controls.Add($progress)

    # --- Grid ---
    $grid = New-StyledDataGridView
    $tab.Controls.Add($grid); $grid.BringToFront()

    # --- Eventos ---
    $btnBrowse.Add_Click({
        $fb = New-Object System.Windows.Forms.FolderBrowserDialog
        $fb.Description = 'Escolha a pasta ou share'; $fb.ShowNewFolderButton = $false
        if ($fb.ShowDialog() -eq 'OK') { $txtPath.Text = $fb.SelectedPath }
    }.GetNewClosure())

    $btnClear.Add_Click({
        $grid.DataSource = $null
        $script:SA_LastResults = $null
        $btnExport.Enabled = $false
        $lblStatus.Text = 'Pronto.'; $progress.Value = 0
    }.GetNewClosure())

    $btnAudit.Add_Click({
        $path = $txtPath.Text.Trim()
        if (-not $path) { [System.Windows.Forms.MessageBox]::Show('Indique o caminho do share.', 'Aviso', 'OK', 'Warning') | Out-Null; return }
        if (-not (Test-Path -LiteralPath $path)) { [System.Windows.Forms.MessageBox]::Show("Caminho nao acessivel:`n$path", 'Erro', 'OK', 'Error') | Out-Null; return }

        $btnAudit.Enabled = $false; $btnExport.Enabled = $false; $btnClear.Enabled = $false; $btnBrowse.Enabled = $false
        $grid.DataSource = $null; $progress.Value = 0
        $lblStatus.Text = 'A auditar...'
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $progressCb = {
                param($i, $total, $folder)
                $pct = if ($total -gt 0) { [int](($i / $total) * 100) } else { 0 }
                $progress.Value = [Math]::Min(100, [Math]::Max(0, $pct))
                $lblStatus.Text = "A processar ($($i+1)/$total): $folder"
                [System.Windows.Forms.Application]::DoEvents()
            }.GetNewClosure()
            $pumpCb = { [System.Windows.Forms.Application]::DoEvents() }

            $results = SA_Invoke-Audit -SharePath $path -Recurse $chkRecurse.Checked `
                -Depth ([int]$numDepth.Value) -OnlyExplicit $chkOnlyExpl.Checked `
                -OnProgress $progressCb -PumpUI $pumpCb
            $script:SA_LastResults = $results

            if ($results.Count -eq 0) {
                $lblStatus.Text = 'Nenhuma permissao encontrada.'; $progress.Value = 0
            } else {
                $dt = New-Object System.Data.DataTable
                $cols = 'Folder','IsRoot','Principal','Herdada','Permissions','Group','Member','DisplayName','Enabled','Email','Type'
                foreach ($c in $cols) { [void]$dt.Columns.Add($c) }
                foreach ($r in $results) {
                    $row = $dt.NewRow()
                    $row['Folder']=[string]$r.Folder; $row['IsRoot']=if ($r.IsRoot) {'Sim'} else {''}
                    $row['Principal']=[string]$r.Principal; $row['Herdada']=[string]$r.Inherited
                    $row['Permissions']=[string]$r.Permissions; $row['Group']=[string]$r.Group
                    $row['Member']=[string]$r.Member; $row['DisplayName']=[string]$r.DisplayName
                    $row['Enabled']=if ($null -ne $r.Enabled) { if ($r.Enabled) {'Sim'} else {'Nao'} } else { '' }
                    $row['Email']=[string]$r.Email; $row['Type']=[string]$r.Type
                    $dt.Rows.Add($row)
                }
                $grid.DataSource = $dt; $progress.Value = 100
                $folders = ($results | Select-Object -ExpandProperty Folder -Unique).Count
                $users = ($results | Where-Object Type -eq 'User' | Select-Object -ExpandProperty Member -Unique).Count
                $lblStatus.Text = "Concluido: $($results.Count) linhas | $folders pastas | $users users distintos."
                $btnExport.Enabled = $true
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Erro:`n$($_.Exception.Message)", 'Erro', 'OK', 'Error') | Out-Null
            $lblStatus.Text = "Erro: $($_.Exception.Message)"
        }
        finally {
            $btnAudit.Enabled = $true; $btnClear.Enabled = $true; $btnBrowse.Enabled = $true
        }
    }.GetNewClosure())

    $btnExport.Add_Click({
        if (-not $script:SA_LastResults -or $script:SA_LastResults.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Nao ha resultados.', 'Aviso', 'OK', 'Warning') | Out-Null
            return
        }
        $sf = New-Object System.Windows.Forms.SaveFileDialog
        $sf.Filter = 'Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv'
        $sf.FileName = "AuditShare_$((Get-Date).ToString('yyyyMMdd_HHmmss')).xlsx"
        if ($sf.ShowDialog() -ne 'OK') { return }

        $out = $sf.FileName
        $btnExport.Enabled = $false; $btnAudit.Enabled = $false
        $lblStatus.Text = 'A exportar...'
        [System.Windows.Forms.Application]::DoEvents()
        try {
            if ($sf.FilterIndex -eq 1) {
                if (-not (Test-ExcelAvailable)) {
                    $res = [System.Windows.Forms.MessageBox]::Show('Excel nao disponivel. Exportar CSV?', 'Aviso', 'YesNo', 'Question')
                    if ($res -eq 'Yes') { $out = Export-ResultsToCsv -Results $script:SA_LastResults -OutputPath $out }
                    else { $lblStatus.Text = 'Cancelado.'; return }
                } else {
                    SA_Export-ToExcel -Results $script:SA_LastResults -OutputPath $out -SharePath $txtPath.Text
                }
            } else {
                $out = Export-ResultsToCsv -Results $script:SA_LastResults -OutputPath $out
            }
            $lblStatus.Text = "Exportado: $out"
            $res = [System.Windows.Forms.MessageBox]::Show("Guardado em:`n$out`n`nAbrir?", 'OK', 'YesNo', 'Information')
            if ($res -eq 'Yes') { Start-Process -FilePath $out }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Erro:`n$($_.Exception.Message)", 'Erro', 'OK', 'Error') | Out-Null
        }
        finally {
            $btnExport.Enabled = ($script:SA_LastResults.Count -gt 0); $btnAudit.Enabled = $true
        }
    }.GetNewClosure())

    return $tab
}
