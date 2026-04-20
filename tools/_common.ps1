# ============================================================
# _common.ps1 - Helpers partilhados por todas as ferramentas
# ============================================================
# Dot-sourced por ITAdminToolkit.ps1 ANTES dos tools, para que
# as ferramentas possam usar $script:ADAvailable, caches, helpers.

# ----- Deteccao do modulo ActiveDirectory -----
$script:ADAvailable = $false
try {
    if (Get-Module -ListAvailable -Name ActiveDirectory) {
        Import-Module ActiveDirectory -ErrorAction Stop
        $script:ADAvailable = $true
    }
} catch { $script:ADAvailable = $false }

# ----- Cache global de users (partilhada entre ferramentas) -----
$script:UserDetailsCache = @{}

# ============================================================
# FUNCOES PARTILHADAS
# ============================================================

function Test-ExcelAvailable {
    try { $null = New-Object -ComObject Excel.Application; return $true } catch { return $false }
}

function Convert-RgbToBgr {
    param([int]$Rgb)
    $r = ($Rgb -shr 16) -band 0xFF
    $g = ($Rgb -shr 8) -band 0xFF
    $b =  $Rgb -band 0xFF
    return ($b -shl 16) -bor ($g -shl 8) -bor $r
}

function Export-ResultsToCsv {
    param([System.Collections.IList]$Results, [string]$OutputPath)
    if ([System.IO.Path]::GetExtension($OutputPath) -ne '.csv') {
        $OutputPath = [System.IO.Path]::ChangeExtension($OutputPath, '.csv')
    }
    $Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    return $OutputPath
}

function Get-ADUserDetailsBatch {
    <#
    Recebe objectos de Get-ADGroupMember (com .distinguishedName e .SamAccountName)
    e devolve hashtable DN -> user ADObject completo. Chunks de 100 DNs por query LDAP.
    Usa e actualiza $script:UserDetailsCache para evitar queries repetidas entre
    ferramentas e entre chamadas.
    #>
    param(
        [object[]]$UserMembers,
        [ScriptBlock]$PumpUI = $null
    )
    $result = @{}
    if (-not $UserMembers -or $UserMembers.Count -eq 0) { return $result }

    # Reaproveitar cache
    $toFetch = @()
    foreach ($u in $UserMembers) {
        if ($script:UserDetailsCache.ContainsKey($u.distinguishedName)) {
            $result[$u.distinguishedName] = $script:UserDetailsCache[$u.distinguishedName]
        } else {
            $toFetch += $u
        }
    }
    if ($toFetch.Count -eq 0) { return $result }

    $chunkSize = 100
    for ($i = 0; $i -lt $toFetch.Count; $i += $chunkSize) {
        if ($PumpUI) { & $PumpUI }
        $end = [Math]::Min($i + $chunkSize - 1, $toFetch.Count - 1)
        $chunk = $toFetch[$i..$end]
        $parts = foreach ($u in $chunk) {
            $dn = $u.distinguishedName -replace '\\', '\5c' -replace '\(', '\28' -replace '\)', '\29'
            "(distinguishedName=$dn)"
        }
        $filter = '(|' + ($parts -join '') + ')'
        try {
            Get-ADUser -LDAPFilter $filter -Properties DisplayName, Enabled, EmailAddress, Title, Department -ErrorAction Stop |
                ForEach-Object {
                    $result[$_.DistinguishedName] = $_
                    $script:UserDetailsCache[$_.DistinguishedName] = $_
                }
        } catch {
            foreach ($u in $chunk) {
                try {
                    $full = Get-ADUser -Identity $u.distinguishedName -Properties DisplayName, Enabled, EmailAddress, Title, Department -ErrorAction Stop
                    $result[$u.distinguishedName] = $full
                    $script:UserDetailsCache[$u.distinguishedName] = $full
                } catch { }
            }
        }
    }
    return $result
}

# ----- Paletas de cores reutilizaveis -----
$script:PaletteBlue = @{
    Dark     = [System.Drawing.Color]::FromArgb(48, 84, 150)
    Mid      = [System.Drawing.Color]::FromArgb(68, 114, 196)
    Light    = [System.Drawing.Color]::FromArgb(217, 225, 242)
    VLight   = [System.Drawing.Color]::FromArgb(245, 248, 253)
}

$script:PaletteGreen = @{
    Dark = [System.Drawing.Color]::FromArgb(84, 130, 53)
}

# ============================================================
# TEMA "Corporate Dark" (design novobanco - teal)
# Mapeado do design em IT Admin Toolkit.html / styles.css
# ============================================================
$script:Theme = @{
    # Background e painéis
    Bg           = [System.Drawing.Color]::FromArgb(12, 16, 20)    # #0C1014
    Bg2          = [System.Drawing.Color]::FromArgb(17, 24, 32)    # #111820
    Panel        = [System.Drawing.Color]::FromArgb(20, 28, 38)    # #141C26
    Panel2       = [System.Drawing.Color]::FromArgb(26, 36, 49)    # #1A2431

    # Borders
    Border       = [System.Drawing.Color]::FromArgb(31, 43, 58)    # #1F2B3A
    BorderStrong = [System.Drawing.Color]::FromArgb(43, 59, 79)    # #2B3B4F

    # Text
    Text         = [System.Drawing.Color]::FromArgb(231, 238, 246) # #E7EEF6
    TextDim      = [System.Drawing.Color]::FromArgb(138, 155, 178) # #8A9BB2
    TextFaint    = [System.Drawing.Color]::FromArgb(90, 107, 130)  # #5A6B82

    # Accent (novobanco teal)
    Accent       = [System.Drawing.Color]::FromArgb(0, 161, 154)   # #00A19A
    Accent600    = [System.Drawing.Color]::FromArgb(0, 143, 137)   # #008F89
    Accent400    = [System.Drawing.Color]::FromArgb(38, 189, 183)  # #26BDB7
    Accent300    = [System.Drawing.Color]::FromArgb(77, 208, 203)  # #4DD0CB

    # Semânticos
    Ok           = [System.Drawing.Color]::FromArgb(48, 164, 108)  # #30A46C
    Warn         = [System.Drawing.Color]::FromArgb(245, 165, 36)  # #F5A524
    Danger       = [System.Drawing.Color]::FromArgb(229, 72, 77)   # #E5484D

    # Cores para categoria M365 (azul do design)
    M365         = [System.Drawing.Color]::FromArgb(91, 141, 239)  # #5B8DEF

    # Fonts
    FontSans     = 'Segoe UI'                                       # fallback de 'Inter'
    FontMono     = 'Consolas'                                       # fallback de 'JetBrains Mono'
}

# ----- Paleta ciclica para sub-tabs (adaptada para fundo escuro) -----
$script:SubTabPalette = @(
    $script:Theme.Accent,                                           # teal (1a ferramenta de cada categoria)
    [System.Drawing.Color]::FromArgb(91, 141, 239),                 # azul M365
    [System.Drawing.Color]::FromArgb(165, 89, 199),                 # roxo
    [System.Drawing.Color]::FromArgb(245, 165, 36),                 # ambar
    [System.Drawing.Color]::FromArgb(48, 164, 108),                 # verde ok
    [System.Drawing.Color]::FromArgb(77, 208, 203),                 # teal claro
    [System.Drawing.Color]::FromArgb(229, 72, 77),                  # vermelho
    [System.Drawing.Color]::FromArgb(138, 155, 178)                 # cinza-azulado
)

# ============================================================
# APPLY-DARKTHEME: retematiza recursivamente uma arvore de controlos
# Preserva cores explicitas ja definidas pelos tools (ex: titulos azuis,
# avisos laranja) - so actualiza valores que estejam nos defaults do WinForms.
# ============================================================
function Apply-DarkTheme {
    param(
        [Parameter(Mandatory)] [System.Windows.Forms.Control]$Control,
        [hashtable]$Theme = $script:Theme
    )

    $defaultWhite = [System.Drawing.Color]::White.ToArgb()
    $defaultCtlBg = [System.Drawing.SystemColors]::Control.ToArgb()
    $defaultCtlTx = [System.Drawing.SystemColors]::ControlText.ToArgb()
    $defaultBlack = [System.Drawing.Color]::Black.ToArgb()

    foreach ($child in $Control.Controls) {
        $type = $child.GetType().Name

        switch ($type) {
            'TabPage' {
                $child.BackColor = $Theme.Panel
                $child.ForeColor = $Theme.Text
            }
            'Panel' {
                if ($child.BackColor.ToArgb() -eq $defaultWhite -or
                    $child.BackColor.ToArgb() -eq $defaultCtlBg) {
                    $child.BackColor = $Theme.Panel
                    $child.ForeColor = $Theme.Text
                }
            }
            'Label' {
                # Preservar ForeColor se foi customizado (titulos, avisos)
                $isDefault = ($child.ForeColor.ToArgb() -eq $defaultCtlTx) -or
                             ($child.ForeColor.ToArgb() -eq $defaultBlack)
                if ($isDefault) { $child.ForeColor = $Theme.Text }
                $child.BackColor = [System.Drawing.Color]::Transparent
            }
            'TextBox' {
                $child.BackColor   = $Theme.Bg2
                $child.ForeColor   = $Theme.Text
                $child.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            }
            'CheckBox' {
                $child.ForeColor = $Theme.Text
                $child.BackColor = [System.Drawing.Color]::Transparent
                $child.FlatStyle = 'Flat'
            }
            'RadioButton' {
                $child.ForeColor = $Theme.Text
                $child.BackColor = [System.Drawing.Color]::Transparent
                $child.FlatStyle = 'Flat'
            }
            'NumericUpDown' {
                $child.BackColor   = $Theme.Bg2
                $child.ForeColor   = $Theme.Text
                $child.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            }
            'DataGridView' {
                $child.BackgroundColor = $Theme.Bg
                $child.GridColor       = $Theme.Border
                $child.BorderStyle     = 'None'
                $child.DefaultCellStyle.BackColor          = $Theme.Panel
                $child.DefaultCellStyle.ForeColor          = $Theme.Text
                $child.DefaultCellStyle.SelectionBackColor = $Theme.Accent
                $child.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
                $child.AlternatingRowsDefaultCellStyle.BackColor = $Theme.Bg2
                $child.AlternatingRowsDefaultCellStyle.ForeColor = $Theme.Text
                $child.ColumnHeadersDefaultCellStyle.BackColor = $Theme.Bg2
                $child.ColumnHeadersDefaultCellStyle.ForeColor = $Theme.TextDim
                $child.EnableHeadersVisualStyles = $false
                $child.RowHeadersVisible = $false
            }
            'ProgressBar' { }
            'Button'      { }
        }

        if ($child.Controls.Count -gt 0) {
            Apply-DarkTheme -Control $child -Theme $Theme
        }
    }
}

# ----- Helper: pintar tabs com cores por indice (owner-draw) -----
function Enable-ColoredTabs {
    <#
    Aplica owner-draw a um TabControl com cores por tab. A paleta pode ser fornecida
    (uma cor por tab) ou ciclica (SubTabPalette). O tab seleccionado usa a cor cheia
    e os restantes a versao mais escura.
    #>
    param(
        [Parameter(Mandatory)] [System.Windows.Forms.TabControl]$TabControl,
        [System.Drawing.Color[]]$Colors = $script:SubTabPalette,
        [int]$TabWidth = 140,
        [int]$TabHeight = 28
    )

    $TabControl.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
    $TabControl.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
    $TabControl.ItemSize = New-Object System.Drawing.Size($TabWidth, $TabHeight)
    # Guardar paleta no Tag para o handler poder recupera-la
    $TabControl.Tag = @{ Colors = $Colors }

    $TabControl.Add_DrawItem({
        param($sender, $e)
        $idx = $e.Index
        if ($idx -lt 0) { return }
        $palette = $sender.Tag.Colors
        if (-not $palette -or $palette.Count -eq 0) { return }

        $baseColor = $palette[$idx % $palette.Count]
        $isSelected = ($sender.SelectedIndex -eq $idx)

        $bgColor = if ($isSelected) {
            $baseColor
        } else {
            [System.Drawing.Color]::FromArgb(
                [Math]::Max(0, [int]$baseColor.R - 60),
                [Math]::Max(0, [int]$baseColor.G - 60),
                [Math]::Max(0, [int]$baseColor.B - 60))
        }

        $rect = $e.Bounds
        if ($isSelected) {
            $rect = [System.Drawing.Rectangle]::new($rect.X, $rect.Y - 2, $rect.Width, $rect.Height + 2)
        }

        $brush = New-Object System.Drawing.SolidBrush($bgColor)
        $e.Graphics.FillRectangle($brush, $rect)
        $brush.Dispose()

        $textBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = [System.Drawing.StringAlignment]::Center
        $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
        $textRect = [System.Drawing.RectangleF]::new([float]$rect.X, [float]$rect.Y, [float]$rect.Width, [float]$rect.Height)
        $e.Graphics.DrawString($sender.TabPages[$idx].Text, $sender.Font, $textBrush, $textRect, $sf)
        $textBrush.Dispose(); $sf.Dispose()
    })
}

# ----- Helper: criar botao estilizado (primario teal ou secundario outlined) -----
function New-StyledButton {
    param(
        [string]$Text,
        [System.Drawing.Color]$BackColor = [System.Drawing.Color]::Empty,
        [System.Drawing.Color]$ForeColor = [System.Drawing.Color]::Empty,
        [int]$X, [int]$Y, [int]$Width = 140, [int]$Height = 32,
        [bool]$Bold = $false
    )
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Location = New-Object System.Drawing.Point($X, $Y)
    $btn.Size = New-Object System.Drawing.Size($Width, $Height)
    $btn.FlatStyle = 'Flat'

    # Se cores nao foram explicitamente fornecidas, aplicar estilo secundario (outlined dark)
    if ($BackColor -eq [System.Drawing.Color]::Empty) {
        $btn.BackColor = $script:Theme.Panel2
        $btn.ForeColor = $script:Theme.Text
        $btn.FlatAppearance.BorderColor = $script:Theme.BorderStrong
    } else {
        $btn.BackColor = $BackColor
        $btn.ForeColor = if ($ForeColor -eq [System.Drawing.Color]::Empty) { [System.Drawing.Color]::White } else { $ForeColor }
        $btn.FlatAppearance.BorderColor = $BackColor
    }
    $btn.FlatAppearance.BorderSize = 1

    $style = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $btn.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 9, $style)
    $btn.Cursor = 'Hand'
    return $btn
}

# ----- Helper: criar chip/pill (label estilizada) -----
function New-StatusPill {
    param(
        [string]$Text,
        [System.Drawing.Color]$Color = $script:Theme.Ok,
        [int]$X, [int]$Y
    )
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "  *  $Text"
    $lbl.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 8, [System.Drawing.FontStyle]::Bold)
    $lbl.ForeColor = $Color
    $bgR = [int][Math]::Min(255, [Math]::Floor([int]$Color.R / 4) + 18)
    $bgG = [int][Math]::Min(255, [Math]::Floor([int]$Color.G / 4) + 18)
    $bgB = [int][Math]::Min(255, [Math]::Floor([int]$Color.B / 4) + 18)
    $lbl.BackColor = [System.Drawing.Color]::FromArgb($bgR, $bgG, $bgB)
    $lbl.TextAlign = 'MiddleLeft'
    $lbl.AutoSize = $true
    $lbl.Padding = New-Object System.Windows.Forms.Padding(8, 3, 10, 3)
    $lbl.Location = New-Object System.Drawing.Point($X, $Y)
    return $lbl
}

# ----- Helper: criar DataGridView estilizado (tema dark) -----
function New-StyledDataGridView {
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.AutoSizeColumnsMode = 'AllCells'
    $grid.RowHeadersVisible = $false
    $grid.BorderStyle = 'None'
    $grid.BackgroundColor = $script:Theme.Bg
    $grid.GridColor = $script:Theme.Border
    $grid.DefaultCellStyle.BackColor = $script:Theme.Panel
    $grid.DefaultCellStyle.ForeColor = $script:Theme.Text
    $grid.DefaultCellStyle.SelectionBackColor = $script:Theme.Accent
    $grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $grid.DefaultCellStyle.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 9)
    $grid.AlternatingRowsDefaultCellStyle.BackColor = $script:Theme.Bg2
    $grid.AlternatingRowsDefaultCellStyle.ForeColor = $script:Theme.Text
    $grid.ColumnHeadersDefaultCellStyle.BackColor = $script:Theme.Bg2
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = $script:Theme.TextDim
    $grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 9, [System.Drawing.FontStyle]::Bold)
    $grid.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersHeight = 32
    $grid.RowTemplate.Height = 26
    return $grid
}
