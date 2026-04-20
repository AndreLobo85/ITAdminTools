# ============================================================
# MailboxStats.ps1 - Ferramenta Exchange Online
# Consulta Get-MailboxStatistics para um ou mais UPNs
# Exporta: New-MailboxStatsTab
# ============================================================

# ----- Funcoes de dominio (prefixo MBX_) -----

function MBX_Test-ModuleAvailable {
    try {
        return $null -ne (Get-Module -ListAvailable -Name ExchangeOnlineManagement)
    } catch { return $false }
}

function MBX_Get-ConnectionInfo {
    # Devolve $null se nao ligado; caso contrario, o objecto de ligacao
    try {
        $info = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if ($info) { return ($info | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1) }
    } catch { }
    return $null
}

function MBX_Get-Stats {
    param([string]$Upn, [bool]$IncludeRecoverable = $false)
    try {
        $stats = Get-MailboxStatistics -Identity $Upn -ErrorAction Stop

        $recoverable = ''
        if ($IncludeRecoverable) {
            try {
                $folder = Get-MailboxFolderStatistics -Identity $Upn -FolderScope RecoverableItems -ErrorAction Stop |
                          Where-Object { $_.Name -eq 'Recoverable Items' } | Select-Object -First 1
                if ($folder) { $recoverable = "$($folder.FolderSize)" }
            } catch { $recoverable = '(erro)' }
        }

        [PSCustomObject]@{
            UserPrincipalName    = $Upn
            DisplayName          = "$($stats.DisplayName)"
            StorageLimitStatus   = "$($stats.StorageLimitStatus)"
            TotalItemSize        = if ($stats.TotalItemSize)        { "$($stats.TotalItemSize)" }        else { '' }
            TotalDeletedItemSize = if ($stats.TotalDeletedItemSize) { "$($stats.TotalDeletedItemSize)" } else { '' }
            ItemCount            = [int64]$stats.ItemCount
            DeletedItemCount     = [int64]$stats.DeletedItemCount
            LastLogonTime        = if ($stats.LastLogonTime) { $stats.LastLogonTime.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
            RecoverableItems     = $recoverable
            Status               = 'OK'
        }
    } catch {
        [PSCustomObject]@{
            UserPrincipalName    = $Upn
            DisplayName          = ''
            StorageLimitStatus   = ''
            TotalItemSize        = ''
            TotalDeletedItemSize = ''
            ItemCount            = ''
            DeletedItemCount     = ''
            LastLogonTime        = ''
            RecoverableItems     = ''
            Status               = "Erro: $($_.Exception.Message)"
        }
    }
}

function MBX_Export-ToExcel {
    param([System.Collections.IList]$Results, [string]$OutputPath)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.ScreenUpdating = $false
    try {
        $wb = $excel.Workbooks.Add()
        while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
        $ws = $wb.Worksheets.Item(1); $ws.Name = 'MailboxStats'

        $ws.Cells.Item(1,1) = 'Exchange Online - Mailbox Statistics'
        $ws.Cells.Item(1,1).Font.Size = 16; $ws.Cells.Item(1,1).Font.Bold = $true
        $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1,10)).Merge() | Out-Null
        $ws.Cells.Item(2,1) = 'Data:'; $ws.Cells.Item(2,2) = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $ws.Cells.Item(3,1) = 'Total mailboxes:'; $ws.Cells.Item(3,2) = $Results.Count
        for ($r = 2; $r -le 3; $r++) { $ws.Cells.Item($r,1).Font.Bold = $true }

        $hdrs = @('UserPrincipalName','DisplayName','StorageLimitStatus','TotalItemSize','TotalDeletedItemSize','ItemCount','DeletedItemCount','LastLogonTime','RecoverableItems','Status')
        $startRow = 5
        for ($c = 0; $c -lt $hdrs.Count; $c++) { $ws.Cells.Item($startRow, $c+1) = $hdrs[$c] }

        $row = $startRow + 1
        foreach ($r in $Results) {
            $ws.Cells.Item($row,1)  = $r.UserPrincipalName
            $ws.Cells.Item($row,2)  = $r.DisplayName
            $ws.Cells.Item($row,3)  = $r.StorageLimitStatus
            $ws.Cells.Item($row,4)  = $r.TotalItemSize
            $ws.Cells.Item($row,5)  = $r.TotalDeletedItemSize
            $ws.Cells.Item($row,6)  = $r.ItemCount
            $ws.Cells.Item($row,7)  = $r.DeletedItemCount
            $ws.Cells.Item($row,8)  = $r.LastLogonTime
            $ws.Cells.Item($row,9)  = $r.RecoverableItems
            $ws.Cells.Item($row,10) = $r.Status
            if ($r.Status -ne 'OK') {
                $errRng = $ws.Range($ws.Cells.Item($row,1), $ws.Cells.Item($row,10))
                $errRng.Font.Color = Convert-RgbToBgr 0xC00000
                $errRng.Font.Italic = $true
            }
            $row++
        }

        if ($Results.Count -gt 0) {
            $rng = $ws.Range($ws.Cells.Item($startRow,1), $ws.Cells.Item($row-1, $hdrs.Count))
            $lo = $ws.ListObjects.Add(1, $rng, $null, 1)
            $lo.Name = 'tblMailboxStats'; $lo.TableStyle = 'TableStyleMedium2'
        }
        $ws.Columns.AutoFit() | Out-Null
        $ws.Application.ActiveWindow.SplitRow = $startRow
        $ws.Application.ActiveWindow.FreezePanes = $true

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

# ============================================================
# FUNCAO EXPORTADA
# ============================================================

function New-MailboxStatsTab {
    $script:MBX_LastResults = $null
    $script:MBX_Connected = $false

    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = 'Mailbox Stats'
    $tab.BackColor = [System.Drawing.Color]::White

    $panelTop = New-Object System.Windows.Forms.Panel
    $panelTop.Dock = 'Top'; $panelTop.Height = 210
    $panelTop.Padding = '12,12,12,12'
    $panelTop.BackColor = [System.Drawing.Color]::White
    $tab.Controls.Add($panelTop)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = 'Exchange Online - Mailbox Statistics'
    $lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(79, 129, 189)
    $lblTitle.Location = New-Object System.Drawing.Point(12, 4)
    $lblTitle.Size = New-Object System.Drawing.Size(700, 26)
    $panelTop.Controls.Add($lblTitle)

    # ----- Deteccao do modulo EXO -----
    if (-not (MBX_Test-ModuleAvailable)) {
        $lblWarn = New-Object System.Windows.Forms.Label
        $lblWarn.Text = "Modulo ExchangeOnlineManagement nao esta instalado.`n`n" +
                        "Instale com (PowerShell como admin ou CurrentUser):`n" +
                        "    Install-Module ExchangeOnlineManagement -Scope CurrentUser`n`n" +
                        "Depois reinicie a aplicacao."
        $lblWarn.ForeColor = [System.Drawing.Color]::DarkOrange
        $lblWarn.Font = New-Object System.Drawing.Font('Segoe UI', 10)
        $lblWarn.Location = New-Object System.Drawing.Point(12, 40)
        $lblWarn.Size = New-Object System.Drawing.Size(1050, 150)
        $panelTop.Controls.Add($lblWarn)
        return $tab
    }

    # ----- Linha 1: botoes de ligacao -----
    $lblConn = New-Object System.Windows.Forms.Label
    $lblConn.Text = 'Admin UPN:'
    $lblConn.Location = New-Object System.Drawing.Point(12, 42); $lblConn.Size = New-Object System.Drawing.Size(90, 22)
    $panelTop.Controls.Add($lblConn)

    $txtAdminUpn = New-Object System.Windows.Forms.TextBox
    $txtAdminUpn.Location = New-Object System.Drawing.Point(108, 40); $txtAdminUpn.Size = New-Object System.Drawing.Size(260, 24)
    $panelTop.Controls.Add($txtAdminUpn)

    $btnConnect = New-StyledButton -Text 'Ligar ao Exchange Online' -X 378 -Y 38 `
        -BackColor ([System.Drawing.Color]::FromArgb(79, 129, 189)) -ForeColor 'White' -Bold $true -Width 180
    $panelTop.Controls.Add($btnConnect)

    $btnDisconnect = New-StyledButton -Text 'Desligar' -X 568 -Y 38 -Width 110
    $btnDisconnect.Enabled = $false
    $panelTop.Controls.Add($btnDisconnect)

    $lblConnStatus = New-Object System.Windows.Forms.Label
    $lblConnStatus.Text = 'Nao ligado.'
    $lblConnStatus.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $lblConnStatus.ForeColor = [System.Drawing.Color]::DarkOrange
    $lblConnStatus.Location = New-Object System.Drawing.Point(690, 42); $lblConnStatus.Size = New-Object System.Drawing.Size(380, 22)
    $panelTop.Controls.Add($lblConnStatus)

    # ----- Linha 2: UPNs a consultar -----
    $lblUpns = New-Object System.Windows.Forms.Label
    $lblUpns.Text = 'UPN(s) a consultar (um por linha, ou separados por virgula):'
    $lblUpns.Location = New-Object System.Drawing.Point(12, 74); $lblUpns.Size = New-Object System.Drawing.Size(400, 22)
    $panelTop.Controls.Add($lblUpns)

    $txtUpns = New-Object System.Windows.Forms.TextBox
    $txtUpns.Multiline = $true; $txtUpns.ScrollBars = 'Vertical'
    $txtUpns.Location = New-Object System.Drawing.Point(12, 96); $txtUpns.Size = New-Object System.Drawing.Size(720, 64)
    $txtUpns.Font = New-Object System.Drawing.Font('Consolas', 9)
    $panelTop.Controls.Add($txtUpns)

    $chkRecov = New-Object System.Windows.Forms.CheckBox
    $chkRecov.Text = 'Incluir RecoverableItems (query adicional, mais lento)'
    $chkRecov.Checked = $true
    $chkRecov.Location = New-Object System.Drawing.Point(748, 96); $chkRecov.Size = New-Object System.Drawing.Size(330, 22)
    $panelTop.Controls.Add($chkRecov)

    # ----- Linha 3: botoes de accao -----
    $btnQuery = New-StyledButton -Text 'Consultar' -X 12 -Y 168 `
        -BackColor ([System.Drawing.Color]::FromArgb(48, 84, 150)) -ForeColor 'White' -Bold $true -Width 130
    $btnQuery.Enabled = $false
    $panelTop.Controls.Add($btnQuery)

    $btnExport = New-StyledButton -Text 'Exportar Excel' -X 152 -Y 168 `
        -BackColor $script:PaletteGreen.Dark -ForeColor 'White' -Bold $true -Width 140
    $btnExport.Enabled = $false
    $panelTop.Controls.Add($btnExport)

    $btnClear = New-StyledButton -Text 'Limpar' -X 302 -Y 168 -Width 100
    $panelTop.Controls.Add($btnClear)

    # --- Painel status ---
    $panelStatus = New-Object System.Windows.Forms.Panel
    $panelStatus.Dock = 'Bottom'; $panelStatus.Height = 52; $panelStatus.Padding = '12,6,12,6'
    $panelStatus.BackColor = [System.Drawing.Color]::White
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

    # --- Grid ---
    $grid = New-StyledDataGridView
    $tab.Controls.Add($grid); $grid.BringToFront()

    # Detectar se ja ha sessao activa (caso o user tenha ligado previamente noutra ferramenta)
    $existing = MBX_Get-ConnectionInfo
    if ($existing) {
        $script:MBX_Connected = $true
        $lblConnStatus.Text = "Ligado como: $($existing.UserPrincipalName)"
        $lblConnStatus.ForeColor = [System.Drawing.Color]::Green
        $btnConnect.Enabled = $false; $btnDisconnect.Enabled = $true; $btnQuery.Enabled = $true
        if ($existing.UserPrincipalName) { $txtAdminUpn.Text = $existing.UserPrincipalName }
    }

    # ============================================================
    # EVENTOS
    # ============================================================

    $btnConnect.Add_Click({
        $btnConnect.Enabled = $false
        $lblConnStatus.Text = 'A ligar... (pode abrir popup de autenticacao)'
        $lblConnStatus.ForeColor = [System.Drawing.Color]::DarkBlue
        [System.Windows.Forms.Application]::DoEvents()

        try {
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
            $params = @{ ShowBanner = $false; ErrorAction = 'Stop' }
            if ($txtAdminUpn.Text.Trim()) { $params.UserPrincipalName = $txtAdminUpn.Text.Trim() }
            Connect-ExchangeOnline @params

            $info = MBX_Get-ConnectionInfo
            if ($info) {
                $script:MBX_Connected = $true
                $lblConnStatus.Text = "Ligado como: $($info.UserPrincipalName)"
                $lblConnStatus.ForeColor = [System.Drawing.Color]::Green
                $btnDisconnect.Enabled = $true
                $btnQuery.Enabled = $true
                if (-not $txtAdminUpn.Text -and $info.UserPrincipalName) { $txtAdminUpn.Text = $info.UserPrincipalName }
                $lblStatus.Text = 'Ligacao estabelecida.'
            } else {
                throw 'Nao foi possivel confirmar a ligacao apos Connect-ExchangeOnline.'
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Erro a ligar:`n$($_.Exception.Message)", 'Erro', 'OK', 'Error') | Out-Null
            $lblConnStatus.Text = 'Falha na ligacao.'
            $lblConnStatus.ForeColor = [System.Drawing.Color]::Red
            $btnConnect.Enabled = $true
        }
    }.GetNewClosure())

    $btnDisconnect.Add_Click({
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
            $script:MBX_Connected = $false
            $lblConnStatus.Text = 'Nao ligado.'; $lblConnStatus.ForeColor = [System.Drawing.Color]::DarkOrange
            $btnConnect.Enabled = $true; $btnDisconnect.Enabled = $false; $btnQuery.Enabled = $false
            $lblStatus.Text = 'Desligado.'
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro a desligar:`n$($_.Exception.Message)", 'Erro', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())

    $btnClear.Add_Click({
        $grid.DataSource = $null; $script:MBX_LastResults = $null
        $btnExport.Enabled = $false; $lblStatus.Text = 'Pronto.'; $progress.Value = 0
    }.GetNewClosure())

    $btnQuery.Add_Click({
        if (-not $script:MBX_Connected) {
            [System.Windows.Forms.MessageBox]::Show('Ligue primeiro ao Exchange Online.', 'Aviso', 'OK', 'Warning') | Out-Null
            return
        }
        $raw = $txtUpns.Text
        if (-not $raw -or -not $raw.Trim()) {
            [System.Windows.Forms.MessageBox]::Show('Indique pelo menos um UPN.', 'Aviso', 'OK', 'Warning') | Out-Null
            return
        }
        # Aceita uma por linha OU separadas por virgula / ponto-virgula
        $upns = $raw -split '[\r\n,;]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique
        if ($upns.Count -eq 0) { return }

        $btnQuery.Enabled = $false; $btnExport.Enabled = $false; $btnClear.Enabled = $false
        $btnDisconnect.Enabled = $false
        $grid.DataSource = $null; $progress.Value = 0
        $lblStatus.Text = "A consultar $($upns.Count) mailbox(es)..."
        [System.Windows.Forms.Application]::DoEvents()

        try {
            $results = New-Object System.Collections.Generic.List[object]
            $includeRecov = $chkRecov.Checked

            for ($i = 0; $i -lt $upns.Count; $i++) {
                $upn = $upns[$i]
                $pct = [int](($i / [Math]::Max(1, $upns.Count)) * 100)
                $progress.Value = [Math]::Min(100, [Math]::Max(0, $pct))
                $lblStatus.Text = "A consultar ($($i+1)/$($upns.Count)): $upn"
                [System.Windows.Forms.Application]::DoEvents()

                $row = MBX_Get-Stats -Upn $upn -IncludeRecoverable $includeRecov
                $results.Add($row)
            }

            $script:MBX_LastResults = $results
            $progress.Value = 100

            $dt = New-Object System.Data.DataTable
            $cols = 'UserPrincipalName','DisplayName','StorageLimitStatus','TotalItemSize','TotalDeletedItemSize','ItemCount','DeletedItemCount','LastLogonTime','RecoverableItems','Status'
            foreach ($c in $cols) { [void]$dt.Columns.Add($c) }
            foreach ($r in $results) {
                $row = $dt.NewRow()
                foreach ($c in $cols) { $row[$c] = [string]$r.$c }
                $dt.Rows.Add($row)
            }
            $grid.DataSource = $dt
            $okCount = ($results | Where-Object Status -eq 'OK').Count
            $errCount = $results.Count - $okCount
            $lblStatus.Text = "Concluido: $okCount OK | $errCount erro(s) | $($results.Count) total."
            $btnExport.Enabled = $true
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Erro:`n$($_.Exception.Message)", 'Erro', 'OK', 'Error') | Out-Null
            $lblStatus.Text = "Erro: $($_.Exception.Message)"
        }
        finally {
            $btnQuery.Enabled = $true; $btnClear.Enabled = $true; $btnDisconnect.Enabled = $true
        }
    }.GetNewClosure())

    $btnExport.Add_Click({
        if (-not $script:MBX_LastResults -or $script:MBX_LastResults.Count -eq 0) { return }
        $sf = New-Object System.Windows.Forms.SaveFileDialog
        $sf.Filter = 'Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv'
        $sf.FileName = "MailboxStats_$((Get-Date).ToString('yyyyMMdd_HHmmss')).xlsx"
        if ($sf.ShowDialog() -ne 'OK') { return }
        $out = $sf.FileName
        $btnExport.Enabled = $false; $btnQuery.Enabled = $false
        $lblStatus.Text = 'A exportar...'
        [System.Windows.Forms.Application]::DoEvents()
        try {
            if ($sf.FilterIndex -eq 1) {
                if (-not (Test-ExcelAvailable)) {
                    $res = [System.Windows.Forms.MessageBox]::Show('Excel nao disponivel. Exportar CSV?', 'Aviso', 'YesNo', 'Question')
                    if ($res -eq 'Yes') { $out = Export-ResultsToCsv -Results $script:MBX_LastResults -OutputPath $out }
                    else { $lblStatus.Text = 'Cancelado.'; return }
                } else {
                    MBX_Export-ToExcel -Results $script:MBX_LastResults -OutputPath $out
                }
            } else {
                $out = Export-ResultsToCsv -Results $script:MBX_LastResults -OutputPath $out
            }
            $lblStatus.Text = "Exportado: $out"
            $res = [System.Windows.Forms.MessageBox]::Show("Guardado:`n$out`n`nAbrir?", 'OK', 'YesNo', 'Information')
            if ($res -eq 'Yes') { Start-Process -FilePath $out }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Erro:`n$($_.Exception.Message)", 'Erro', 'OK', 'Error') | Out-Null
        }
        finally {
            $btnExport.Enabled = ($script:MBX_LastResults.Count -gt 0); $btnQuery.Enabled = $true
        }
    }.GetNewClosure())

    return $tab
}
