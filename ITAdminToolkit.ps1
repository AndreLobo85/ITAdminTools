<#
.SYNOPSIS
    IT Admin Toolkit - Dashboard-first app inspirada no design Claude (novobanco teal).
.DESCRIPTION
    Entry: Dashboard com hero + 2 grandes areas (ON PREM / 365).
    Click numa area -> ToolView com sidebar + runner.
    Click num tool -> carrega a ferramenta no runner.
    Back button -> volta ao Dashboard.
#>

$ErrorActionPreference = 'Stop'

# ----- Log + trap -----
$script:LogPath = Join-Path $env:TEMP "ITAdminToolkit-$((Get-Date).ToString('yyyyMMdd_HHmmss')).log"
function Write-AppLog {
    param([string]$Message)
    try { Add-Content -Path $script:LogPath -Value "[$((Get-Date).ToString('HH:mm:ss.fff'))] $Message" -Encoding UTF8 } catch { }
}
Write-AppLog "=== IT Admin Toolkit arranque ==="

trap {
    $err = $_
    $msg = "Erro critico: $($err.Exception.Message)`n`nStack:`n$($err.ScriptStackTrace)`n`nLog: $script:LogPath"
    Write-AppLog "TRAP: $msg"
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($msg, 'IT Admin Toolkit - Erro', 'OK', 'Error') | Out-Null
    } catch { }
    exit 1
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ToolsDir  = Join-Path $ScriptDir 'tools'

# ----- Helpers partilhados -----
Write-AppLog "A carregar _common.ps1"
. (Join-Path $ToolsDir '_common.ps1')
Write-AppLog "A carregar _splash.ps1"
. (Join-Path $ToolsDir '_splash.ps1')

# ----- Splash -----
try { Show-Splash -DurationMs 1800 } catch { Write-AppLog "Splash ERRO: $($_.Exception.Message)" }
Write-AppLog "Splash terminado"

# ----- Ferramentas -----
foreach ($tf in 'ShareAuditor.ps1','ADGroupAuditor.ps1','UserInfo.ps1','GroupInfo.ps1','MailboxStats.ps1') {
    Write-AppLog "Carregar $tf"
    . (Join-Path $ToolsDir $tf)
}

# ============================================================
# CONFIGURACAO DE CATEGORIAS + AREAS
# ============================================================
$Categories = @(
    @{ Name='Active Directory';        OnPrem=$true;  Color=[System.Drawing.Color]::FromArgb(0,120,212);   Area='ONPREM'; Icon='AD';
       Tools=@(
         @{ Name='User Info';     Desc='Diagnostico de user AD (Username ou Email)';             Factory={ New-UserInfoTab } },
         @{ Name='Group Info';    Desc='Diagnostico de grupo AD + lista de membros';             Factory={ New-GroupInfoTab } },
         @{ Name='Group Auditor'; Desc='Auditar grupos AD por sufixo ou nome (recursivo)';       Factory={ New-ADGroupAuditorTab } }
       ) },
    @{ Name='Fileshare';               OnPrem=$true;  Color=[System.Drawing.Color]::FromArgb(230,145,56);  Area='ONPREM'; Icon='FS';
       Tools=@(
         @{ Name='Share Auditor'; Desc='Auditar permissoes SMB/NTFS de um share';                Factory={ New-ShareAuditorTab } }
       ) },
    @{ Name='Exchange On-Premise';     OnPrem=$true;  Color=[System.Drawing.Color]::FromArgb(192,80,77);   Area='ONPREM'; Icon='EX'; Tools=@() },
    @{ Name='Exchange Online';         OnPrem=$false; Color=[System.Drawing.Color]::FromArgb(79,129,189);  Area='M365';   Icon='EO';
       Tools=@(
         @{ Name='Mailbox Stats'; Desc='Exchange Online - statistics de mailbox';                Factory={ New-MailboxStatsTab } }
       ) },
    @{ Name='SharePoint Online';       OnPrem=$false; Color=[System.Drawing.Color]::FromArgb(128,100,162); Area='M365';   Icon='SP'; Tools=@() },
    @{ Name='Microsoft 365 / Entra';   OnPrem=$false; Color=[System.Drawing.Color]::FromArgb(47,165,165);  Area='M365';   Icon='M3'; Tools=@() }
)

$Areas = @(
    @{ Id='ONPREM'; Name='ON PREM'; Desc='Active Directory, Fileshares, Exchange on-prem. Requer acesso directo ao dominio.'; Accent=$script:Theme.Accent; Scope='[ON-PREM]' },
    @{ Id='M365';   Name='365';     Desc='Exchange Online, SharePoint Online, Entra ID. Corre da sua maquina com modulos cloud.'; Accent=$script:Theme.M365;    Scope='[CLOUD]' }
)

function Get-AreaCategories { param($AreaId) $Categories | Where-Object { $_.Area -eq $AreaId } }
function Get-AreaToolCount  { param($AreaId)
    $tc = 0; foreach ($c in (Get-AreaCategories $AreaId)) { $tc += $c.Tools.Count }; return $tc
}

# ============================================================
# FORM principal
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = 'IT Admin Toolkit - novobanco'
$form.Size = New-Object System.Drawing.Size(1280, 860)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(1024, 700)
$form.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 9)
$form.BackColor = $script:Theme.Bg
$form.ForeColor = $script:Theme.Text

# ============================================================
# TITLEBAR (48px)
# ============================================================
$titlebar = New-Object System.Windows.Forms.Panel
$titlebar.Dock = 'Top'; $titlebar.Height = 48; $titlebar.BackColor = $script:Theme.Panel

$tbBorder = New-Object System.Windows.Forms.Panel
$tbBorder.Dock = 'Bottom'; $tbBorder.Height = 1; $tbBorder.BackColor = $script:Theme.Border
$titlebar.Controls.Add($tbBorder)

$tbMark = New-Object System.Windows.Forms.Label
$tbMark.Text = 'N'
$tbMark.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 11, [System.Drawing.FontStyle]::Bold)
$tbMark.ForeColor = [System.Drawing.Color]::White; $tbMark.BackColor = $script:Theme.Accent
$tbMark.TextAlign = 'MiddleCenter'
$tbMark.Size = New-Object System.Drawing.Size(24,24)
$tbMark.Location = New-Object System.Drawing.Point(14,12)
$titlebar.Controls.Add($tbMark)

$tbName = New-Object System.Windows.Forms.Label
$tbName.Text = 'IT Admin Toolkit'
$tbName.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 10, [System.Drawing.FontStyle]::Bold)
$tbName.ForeColor = $script:Theme.Text; $tbName.AutoSize = $true
$tbName.Location = New-Object System.Drawing.Point(46,15)
$titlebar.Controls.Add($tbName)

$tbSuffix = New-Object System.Windows.Forms.Label
$tbSuffix.Text = 'NOVOBANCO'
$tbSuffix.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 8)
$tbSuffix.ForeColor = $script:Theme.TextDim; $tbSuffix.AutoSize = $true
$tbSuffix.Location = New-Object System.Drawing.Point(170,18)
$titlebar.Controls.Add($tbSuffix)

# Breadcrumb (area atual)
$tbCrumb = New-Object System.Windows.Forms.Label
$tbCrumb.Text = ''
$tbCrumb.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 9)
$tbCrumb.ForeColor = $script:Theme.TextDim; $tbCrumb.AutoSize = $true
$tbCrumb.Location = New-Object System.Drawing.Point(260,18)
$titlebar.Controls.Add($tbCrumb)

# Pills (lado direito)
$statusText = if ($script:ADAvailable) { 'AD ONLINE' } else { 'AD OFFLINE' }
$statusColor = if ($script:ADAvailable) { $script:Theme.Ok } else { $script:Theme.Warn }
$tbStatus = New-StatusPill -Text $statusText -Color $statusColor -X 0 -Y 13
$tbStatus.Anchor = 'Top, Right'
$tbStatus.Location = New-Object System.Drawing.Point(($form.Width - 380),13)
$titlebar.Controls.Add($tbStatus)

$tbEnv = New-Object System.Windows.Forms.Label
$tbEnv.Text = "  $env:USERDOMAIN  "
$tbEnv.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 8, [System.Drawing.FontStyle]::Bold)
$tbEnv.ForeColor = $script:Theme.Warn
$tbEnv.BackColor = [System.Drawing.Color]::FromArgb(40,32,12)
$tbEnv.TextAlign = 'MiddleCenter'; $tbEnv.AutoSize = $true
$tbEnv.Padding = New-Object System.Windows.Forms.Padding(6,3,6,3)
$tbEnv.Anchor = 'Top, Right'
$tbEnv.Location = New-Object System.Drawing.Point(($form.Width - 250),13)
$titlebar.Controls.Add($tbEnv)

$tbUser = New-Object System.Windows.Forms.Label
$tbUser.Text = "$env:USERNAME"
$tbUser.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 9)
$tbUser.ForeColor = $script:Theme.TextDim; $tbUser.TextAlign = 'MiddleLeft'; $tbUser.AutoSize = $true
$tbUser.Anchor = 'Top, Right'
$tbUser.Location = New-Object System.Drawing.Point(($form.Width - 130),17)
$titlebar.Controls.Add($tbUser)

$tbAvatar = New-Object System.Windows.Forms.Label
$avInitial = if ($env:USERNAME -and $env:USERNAME.Length -gt 0) { $env:USERNAME.Substring(0,1).ToUpper() } else { '?' }
$tbAvatar.Text = $avInitial
$tbAvatar.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 9, [System.Drawing.FontStyle]::Bold)
$tbAvatar.ForeColor = [System.Drawing.Color]::White; $tbAvatar.BackColor = $script:Theme.Accent600
$tbAvatar.TextAlign = 'MiddleCenter'
$tbAvatar.Size = New-Object System.Drawing.Size(24,24)
$tbAvatar.Anchor = 'Top, Right'
$tbAvatar.Location = New-Object System.Drawing.Point(($form.Width - 160),12)
$titlebar.Controls.Add($tbAvatar)

$form.Controls.Add($titlebar)

# ============================================================
# STATUS BAR inferior
# ============================================================
$statusBar = New-Object System.Windows.Forms.Panel
$statusBar.Dock = 'Bottom'; $statusBar.Height = 24; $statusBar.BackColor = $script:Theme.Panel
$sbBorder = New-Object System.Windows.Forms.Panel
$sbBorder.Dock = 'Top'; $sbBorder.Height = 1; $sbBorder.BackColor = $script:Theme.Border
$statusBar.Controls.Add($sbBorder)

$lblStatusHost = New-Object System.Windows.Forms.Label
$lblStatusHost.Text = "  Host: $env:COMPUTERNAME    User: $env:USERDOMAIN\$env:USERNAME"
$lblStatusHost.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 8)
$lblStatusHost.ForeColor = $script:Theme.TextFaint
$lblStatusHost.TextAlign = 'MiddleLeft'; $lblStatusHost.Dock = 'Fill'
$statusBar.Controls.Add($lblStatusHost)
$form.Controls.Add($statusBar)

# ============================================================
# MAIN CONTAINER (troca entre dashboard e tool view)
# ============================================================
$main = New-Object System.Windows.Forms.Panel
$main.Dock = 'Fill'; $main.BackColor = $script:Theme.Bg
$form.Controls.Add($main)
$main.BringToFront()

# Estado de navegacao
$script:CurrentView     = 'Dashboard'
$script:CurrentArea     = $null
$script:CurrentToolHost = $null   # TabControl que aloja o tab da ferramenta activa

# ============================================================
# DASHBOARD PANEL
# ============================================================
$dashPanel = New-Object System.Windows.Forms.Panel
$dashPanel.Dock = 'Fill'; $dashPanel.BackColor = $script:Theme.Bg
$dashPanel.AutoScroll = $true
$main.Controls.Add($dashPanel)

# Hero
$hero = New-Object System.Windows.Forms.Panel
$hero.Location = New-Object System.Drawing.Point(48,36)
$hero.Size = New-Object System.Drawing.Size(1180,96)
$hero.BackColor = $script:Theme.Bg
$dashPanel.Controls.Add($hero)

$heroEyebrow = New-Object System.Windows.Forms.Label
$heroEyebrow.Text = '---  SYSTEMS ADMINISTRATION'
$heroEyebrow.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 8, [System.Drawing.FontStyle]::Bold)
$heroEyebrow.ForeColor = $script:Theme.Accent
$heroEyebrow.AutoSize = $true
$heroEyebrow.Location = New-Object System.Drawing.Point(0,0)
$hero.Controls.Add($heroEyebrow)

$hour = (Get-Date).Hour
$greet = if ($hour -lt 12) { 'Bom dia' } elseif ($hour -lt 19) { 'Boa tarde' } else { 'Boa noite' }
$heroTitle = New-Object System.Windows.Forms.Label
$heroTitle.Text = "$greet, $env:USERNAME."
$heroTitle.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 26, [System.Drawing.FontStyle]::Bold)
$heroTitle.ForeColor = $script:Theme.Text
$heroTitle.AutoSize = $true
$heroTitle.Location = New-Object System.Drawing.Point(0,22)
$hero.Controls.Add($heroTitle)

$heroSub = New-Object System.Windows.Forms.Label
$heroSub.Text = 'Escolha uma area para comecar.'
$heroSub.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 11)
$heroSub.ForeColor = $script:Theme.TextDim
$heroSub.AutoSize = $true
$heroSub.Location = New-Object System.Drawing.Point(0,68)
$hero.Controls.Add($heroSub)

# Stats (canto direito do hero)
$totalTools = 0; foreach ($c in $Categories) { $totalTools += $c.Tools.Count }
$statsY = 10
$statX = 780
$stats = @(
    @{ Val=$totalTools;          Lbl='SCRIPTS' },
    @{ Val=$Areas.Count;         Lbl='AREAS' },
    @{ Val='< 1s';               Lbl='AVG RUNTIME' }
)
for ($i=0; $i -lt $stats.Count; $i++) {
    $s = $stats[$i]
    $statVal = New-Object System.Windows.Forms.Label
    $statVal.Text = "$($s.Val)"
    $statVal.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 22, [System.Drawing.FontStyle]::Bold)
    $statVal.ForeColor = $script:Theme.Text
    $statVal.AutoSize = $true
    $statVal.Location = New-Object System.Drawing.Point(($statX + $i*130), $statsY)
    $hero.Controls.Add($statVal)

    $statLbl = New-Object System.Windows.Forms.Label
    $statLbl.Text = $s.Lbl
    $statLbl.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 8)
    $statLbl.ForeColor = $script:Theme.TextDim
    $statLbl.AutoSize = $true
    $statLbl.Location = New-Object System.Drawing.Point(($statX + $i*130), ($statsY+42))
    $hero.Controls.Add($statLbl)
}

# =========================================================
# AREA CARDS (2 grandes)
# =========================================================
$cardsTop = 160
$cardW = 568
$cardH = 240
$cardGap = 24
$cardMarginL = 48

for ($ai = 0; $ai -lt $Areas.Count; $ai++) {
    $area = $Areas[$ai]
    $x = $cardMarginL + $ai * ($cardW + $cardGap)

    $card = New-Object System.Windows.Forms.Panel
    $card.Location = New-Object System.Drawing.Point($x, $cardsTop)
    $card.Size = New-Object System.Drawing.Size($cardW, $cardH)
    $card.BackColor = $script:Theme.Panel
    $card.Cursor = 'Hand'
    $card.Tag = $area
    $dashPanel.Controls.Add($card)

    # Border (painel 1px em volta)
    $card.Padding = New-Object System.Windows.Forms.Padding(1)

    # Tag/pill no topo
    $areaTag = New-Object System.Windows.Forms.Label
    $areaTag.Text = "  *  $($area.Scope)  $($area.Name)"
    $areaTag.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 8, [System.Drawing.FontStyle]::Bold)
    $areaTag.ForeColor = $area.Accent
    $areaTag.BackColor = [System.Drawing.Color]::FromArgb(
        [int][Math]::Min(255, [int]($area.Accent.R / 5) + 14),
        [int][Math]::Min(255, [int]($area.Accent.G / 5) + 14),
        [int][Math]::Min(255, [int]($area.Accent.B / 5) + 14))
    $areaTag.TextAlign = 'MiddleLeft'
    $areaTag.AutoSize = $true
    $areaTag.Padding = New-Object System.Windows.Forms.Padding(8,4,10,4)
    $areaTag.Location = New-Object System.Drawing.Point(24,24)
    $card.Controls.Add($areaTag)

    # Mini grid decoration (3x3 squares top-right)
    for ($gy = 0; $gy -lt 3; $gy++) {
        for ($gx = 0; $gx -lt 3; $gx++) {
            $sq = New-Object System.Windows.Forms.Panel
            $sq.Size = New-Object System.Drawing.Size(14,14)
            $sq.Location = New-Object System.Drawing.Point(($cardW - 110 + $gx*18), (24 + $gy*18))
            $sq.BackColor = if ($gx -eq 0 -and $gy -eq 1) { $area.Accent } else {
                [System.Drawing.Color]::FromArgb(
                    [int][Math]::Min(255, [int]($area.Accent.R / 4) + 20),
                    [int][Math]::Min(255, [int]($area.Accent.G / 4) + 20),
                    [int][Math]::Min(255, [int]($area.Accent.B / 4) + 20))
            }
            $card.Controls.Add($sq)
        }
    }

    # Nome grande
    $cardName = New-Object System.Windows.Forms.Label
    $cardName.Text = $area.Name
    $cardName.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 36, [System.Drawing.FontStyle]::Bold)
    $cardName.ForeColor = $script:Theme.Text
    $cardName.AutoSize = $true
    $cardName.BackColor = [System.Drawing.Color]::Transparent
    $cardName.Location = New-Object System.Drawing.Point(20,62)
    $card.Controls.Add($cardName)

    # Desc
    $cardDesc = New-Object System.Windows.Forms.Label
    $cardDesc.Text = $area.Desc
    $cardDesc.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 10)
    $cardDesc.ForeColor = $script:Theme.TextDim
    $cardDesc.AutoSize = $false
    $cardDesc.BackColor = [System.Drawing.Color]::Transparent
    $cardDesc.Location = New-Object System.Drawing.Point(24, 120)
    $cardDesc.Size = New-Object System.Drawing.Size(420, 44)
    $card.Controls.Add($cardDesc)

    # Count mono (bottom-left)
    $toolCount = Get-AreaToolCount $area.Id
    $catCount = (Get-AreaCategories $area.Id).Count
    $cardCount = New-Object System.Windows.Forms.Label
    $cardCount.Text = "$toolCount SCRIPTS    $catCount CATEGORIAS"
    $cardCount.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 9, [System.Drawing.FontStyle]::Bold)
    $cardCount.ForeColor = $script:Theme.TextDim
    $cardCount.BackColor = [System.Drawing.Color]::Transparent
    $cardCount.AutoSize = $true
    $cardCount.Location = New-Object System.Drawing.Point(24, ($cardH - 42))
    $card.Controls.Add($cardCount)

    # Arrow circle (bottom-right)
    $cardArrow = New-Object System.Windows.Forms.Label
    $cardArrow.Text = '>'
    $cardArrow.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 18, [System.Drawing.FontStyle]::Bold)
    $cardArrow.ForeColor = $area.Accent
    $cardArrow.BackColor = [System.Drawing.Color]::FromArgb(
        [int][Math]::Min(255, [int]($area.Accent.R / 5) + 14),
        [int][Math]::Min(255, [int]($area.Accent.G / 5) + 14),
        [int][Math]::Min(255, [int]($area.Accent.B / 5) + 14))
    $cardArrow.TextAlign = 'MiddleCenter'
    $cardArrow.Size = New-Object System.Drawing.Size(40,40)
    $cardArrow.Location = New-Object System.Drawing.Point(($cardW - 60), ($cardH - 56))
    $card.Controls.Add($cardArrow)

    # Hover + click
    $hoverIn  = { $this.BackColor = $script:Theme.Panel2 }.GetNewClosure()
    $hoverOut = { $this.BackColor = $script:Theme.Panel }.GetNewClosure()
    $card.Add_MouseEnter($hoverIn)
    $card.Add_MouseLeave($hoverOut)
    # Tambem para os filhos (para hover nao ser perdido quando o rato passa sobre labels)
    foreach ($ch in $card.Controls) {
        $ch.Add_MouseEnter($hoverIn)
        $ch.Add_MouseLeave($hoverOut)
    }

    $clickHandler = {
        param($sender, $e)
        $p = $sender
        while ($p -and -not $p.Tag) { $p = $p.Parent }
        if ($p -and $p.Tag) { Show-ToolView -AreaId $p.Tag.Id }
    }.GetNewClosure()
    $card.Add_Click($clickHandler)
    foreach ($ch in $card.Controls) { $ch.Add_Click($clickHandler) }
}

# =========================================================
# QUICK RUN GRID (abaixo das area cards)
# =========================================================
$qrPanel = New-Object System.Windows.Forms.Panel
$qrPanel.Location = New-Object System.Drawing.Point(48, 430)
$qrPanel.Size = New-Object System.Drawing.Size(1180, 220)
$qrPanel.BackColor = $script:Theme.Panel
$dashPanel.Controls.Add($qrPanel)

$qrTitle = New-Object System.Windows.Forms.Label
$qrTitle.Text = '-  QUICK RUN'
$qrTitle.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 9, [System.Drawing.FontStyle]::Bold)
$qrTitle.ForeColor = $script:Theme.Accent
$qrTitle.BackColor = [System.Drawing.Color]::Transparent
$qrTitle.AutoSize = $true
$qrTitle.Location = New-Object System.Drawing.Point(18,14)
$qrPanel.Controls.Add($qrTitle)

# Flat list de todas as ferramentas com metadata
$allTools = @()
foreach ($cat in $Categories) {
    foreach ($t in $cat.Tools) {
        $allTools += @{
            Name=$t.Name; Desc=$t.Desc; Factory=$t.Factory
            CatName=$cat.Name; CatIcon=$cat.Icon; AreaId=$cat.Area
            Accent=$cat.Color
        }
    }
}

$qrCols = 4
$qrItemW = 270; $qrItemH = 72; $qrStartX = 18; $qrStartY = 46; $qrGap = 10
for ($i = 0; $i -lt $allTools.Count; $i++) {
    $t = $allTools[$i]
    $col = $i % $qrCols
    $row = [Math]::Floor($i / $qrCols)
    $x = $qrStartX + $col * ($qrItemW + $qrGap)
    $y = $qrStartY + $row * ($qrItemH + $qrGap)

    $qi = New-Object System.Windows.Forms.Panel
    $qi.Size = New-Object System.Drawing.Size($qrItemW, $qrItemH)
    $qi.Location = New-Object System.Drawing.Point($x, $y)
    $qi.BackColor = $script:Theme.Panel2
    $qi.Cursor = 'Hand'
    $qi.Tag = $t
    $qrPanel.Controls.Add($qi)

    # Icon box
    $qIcon = New-Object System.Windows.Forms.Label
    $qIcon.Text = $t.CatIcon
    $qIcon.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 10, [System.Drawing.FontStyle]::Bold)
    $qIcon.ForeColor = $t.Accent
    $qIcon.BackColor = [System.Drawing.Color]::FromArgb(
        [int][Math]::Min(255, [int]($t.Accent.R / 5) + 14),
        [int][Math]::Min(255, [int]($t.Accent.G / 5) + 14),
        [int][Math]::Min(255, [int]($t.Accent.B / 5) + 14))
    $qIcon.TextAlign = 'MiddleCenter'
    $qIcon.Size = New-Object System.Drawing.Size(32,32)
    $qIcon.Location = New-Object System.Drawing.Point(12,20)
    $qi.Controls.Add($qIcon)

    # Name
    $qName = New-Object System.Windows.Forms.Label
    $qName.Text = $t.Name
    $qName.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 10, [System.Drawing.FontStyle]::Bold)
    $qName.ForeColor = $script:Theme.Text
    $qName.BackColor = [System.Drawing.Color]::Transparent
    $qName.AutoSize = $true
    $qName.Location = New-Object System.Drawing.Point(56, 16)
    $qi.Controls.Add($qName)

    $qCat = New-Object System.Windows.Forms.Label
    $qCat.Text = $t.CatName.ToUpper()
    $qCat.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 7)
    $qCat.ForeColor = $script:Theme.TextFaint
    $qCat.BackColor = [System.Drawing.Color]::Transparent
    $qCat.AutoSize = $true
    $qCat.Location = New-Object System.Drawing.Point(56, 40)
    $qi.Controls.Add($qCat)

    $qiHoverIn  = { $this.BackColor = $script:Theme.Panel }.GetNewClosure()
    $qiHoverOut = { $this.BackColor = $script:Theme.Panel2 }.GetNewClosure()
    $qi.Add_MouseEnter($qiHoverIn); $qi.Add_MouseLeave($qiHoverOut)
    foreach ($ch in $qi.Controls) { $ch.Add_MouseEnter($qiHoverIn); $ch.Add_MouseLeave($qiHoverOut) }

    $qClick = {
        param($sender, $e)
        $p = $sender; while ($p -and -not $p.Tag) { $p = $p.Parent }
        if ($p -and $p.Tag) {
            Show-ToolView -AreaId $p.Tag.AreaId -PreselectTool $p.Tag.Name
        }
    }.GetNewClosure()
    $qi.Add_Click($qClick)
    foreach ($ch in $qi.Controls) { $ch.Add_Click($qClick) }
}

# ============================================================
# TOOL VIEW PANEL (sidebar + runner)
# ============================================================
$toolPanel = New-Object System.Windows.Forms.Panel
$toolPanel.Dock = 'Fill'; $toolPanel.BackColor = $script:Theme.Bg; $toolPanel.Visible = $false
$main.Controls.Add($toolPanel)

# Sidebar esquerda 320px
$sidebar = New-Object System.Windows.Forms.Panel
$sidebar.Dock = 'Left'; $sidebar.Width = 320
$sidebar.BackColor = $script:Theme.Panel
$toolPanel.Controls.Add($sidebar)

$sbBorderR = New-Object System.Windows.Forms.Panel
$sbBorderR.Dock = 'Right'; $sbBorderR.Width = 1; $sbBorderR.BackColor = $script:Theme.Border
$sidebar.Controls.Add($sbBorderR)

# Back button
$btnBack = New-Object System.Windows.Forms.Button
$btnBack.Text = '  < Dashboard'
$btnBack.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 9)
$btnBack.FlatStyle = 'Flat'
$btnBack.BackColor = $script:Theme.Panel2
$btnBack.ForeColor = $script:Theme.TextDim
$btnBack.FlatAppearance.BorderColor = $script:Theme.Border
$btnBack.Size = New-Object System.Drawing.Size(130, 30)
$btnBack.Location = New-Object System.Drawing.Point(16, 16)
$btnBack.TextAlign = 'MiddleLeft'
$btnBack.Cursor = 'Hand'
$btnBack.Add_Click({ Show-Dashboard })
$sidebar.Controls.Add($btnBack)

$lblAreaBadge = New-Object System.Windows.Forms.Label
$lblAreaBadge.Text = ''
$lblAreaBadge.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 8, [System.Drawing.FontStyle]::Bold)
$lblAreaBadge.AutoSize = $false
$lblAreaBadge.Size = New-Object System.Drawing.Size(110, 22)
$lblAreaBadge.Location = New-Object System.Drawing.Point(16, 56)
$lblAreaBadge.TextAlign = 'MiddleCenter'
$sidebar.Controls.Add($lblAreaBadge)

$lblSbTitle = New-Object System.Windows.Forms.Label
$lblSbTitle.Text = ''
$lblSbTitle.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 18, [System.Drawing.FontStyle]::Bold)
$lblSbTitle.ForeColor = $script:Theme.Text
$lblSbTitle.AutoSize = $false
$lblSbTitle.Size = New-Object System.Drawing.Size(280, 30)
$lblSbTitle.Location = New-Object System.Drawing.Point(16, 82)
$sidebar.Controls.Add($lblSbTitle)

# Search
$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 10)
$txtSearch.BackColor = $script:Theme.Bg2
$txtSearch.ForeColor = $script:Theme.Text
$txtSearch.BorderStyle = 'FixedSingle'
$txtSearch.Size = New-Object System.Drawing.Size(288, 26)
$txtSearch.Location = New-Object System.Drawing.Point(16, 124)
$sidebar.Controls.Add($txtSearch)

# Lista de tools scrollable
$toolListPanel = New-Object System.Windows.Forms.Panel
$toolListPanel.Location = New-Object System.Drawing.Point(8, 164)
$toolListPanel.Size = New-Object System.Drawing.Size(304, 640)
$toolListPanel.AutoScroll = $true
$toolListPanel.BackColor = $script:Theme.Panel
$sidebar.Controls.Add($toolListPanel)

# Runner (direita)
$runner = New-Object System.Windows.Forms.Panel
$runner.Dock = 'Fill'; $runner.BackColor = $script:Theme.Bg
$toolPanel.Controls.Add($runner)

# Header do runner
$runnerHead = New-Object System.Windows.Forms.Panel
$runnerHead.Dock = 'Top'; $runnerHead.Height = 90
$runnerHead.BackColor = $script:Theme.Bg
$runnerHead.Padding = New-Object System.Windows.Forms.Padding(32, 20, 32, 12)
$runner.Controls.Add($runnerHead)

$runnerHeadBorder = New-Object System.Windows.Forms.Panel
$runnerHeadBorder.Dock = 'Bottom'; $runnerHeadBorder.Height = 1; $runnerHeadBorder.BackColor = $script:Theme.Border
$runnerHead.Controls.Add($runnerHeadBorder)

$lblRunnerIcon = New-Object System.Windows.Forms.Label
$lblRunnerIcon.Text = ''
$lblRunnerIcon.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 14, [System.Drawing.FontStyle]::Bold)
$lblRunnerIcon.ForeColor = $script:Theme.Accent
$lblRunnerIcon.BackColor = [System.Drawing.Color]::FromArgb(
    [int][Math]::Min(255, [int]($script:Theme.Accent.R / 5) + 14),
    [int][Math]::Min(255, [int]($script:Theme.Accent.G / 5) + 14),
    [int][Math]::Min(255, [int]($script:Theme.Accent.B / 5) + 14))
$lblRunnerIcon.TextAlign = 'MiddleCenter'
$lblRunnerIcon.Size = New-Object System.Drawing.Size(44,44)
$lblRunnerIcon.Location = New-Object System.Drawing.Point(32, 20)
$runnerHead.Controls.Add($lblRunnerIcon)

$lblRunnerTitle = New-Object System.Windows.Forms.Label
$lblRunnerTitle.Text = 'Seleccione uma ferramenta'
$lblRunnerTitle.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 16, [System.Drawing.FontStyle]::Bold)
$lblRunnerTitle.ForeColor = $script:Theme.Text
$lblRunnerTitle.AutoSize = $true
$lblRunnerTitle.Location = New-Object System.Drawing.Point(92, 22)
$runnerHead.Controls.Add($lblRunnerTitle)

$lblRunnerDesc = New-Object System.Windows.Forms.Label
$lblRunnerDesc.Text = 'Escolha uma ferramenta na barra lateral para comecar.'
$lblRunnerDesc.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 9)
$lblRunnerDesc.ForeColor = $script:Theme.TextDim
$lblRunnerDesc.AutoSize = $true
$lblRunnerDesc.Location = New-Object System.Drawing.Point(92, 50)
$runnerHead.Controls.Add($lblRunnerDesc)

# Host da ferramenta (fill)
$toolHost = New-Object System.Windows.Forms.Panel
$toolHost.Dock = 'Fill'; $toolHost.BackColor = $script:Theme.Bg
$runner.Controls.Add($toolHost)

# Placeholder inicial
$phLabel = New-Object System.Windows.Forms.Label
$phLabel.Text = "Clique numa ferramenta na barra lateral`npara carregar a interface."
$phLabel.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 11)
$phLabel.ForeColor = $script:Theme.TextFaint
$phLabel.TextAlign = 'MiddleCenter'; $phLabel.Dock = 'Fill'; $phLabel.BackColor = $script:Theme.Bg
$toolHost.Controls.Add($phLabel)

# ============================================================
# FUNCOES DE NAVEGACAO
# ============================================================
function Show-Dashboard {
    $script:CurrentView = 'Dashboard'
    $script:CurrentArea = $null
    $toolPanel.Visible = $false
    $dashPanel.Visible = $true
    $dashPanel.BringToFront()
    $tbCrumb.Text = ''
}

function Show-ToolView {
    param([string]$AreaId, [string]$PreselectTool = $null)
    $script:CurrentView = 'Tool'
    $area = $Areas | Where-Object Id -eq $AreaId | Select-Object -First 1
    if (-not $area) { return }
    $script:CurrentArea = $area

    # Atualizar sidebar
    $lblAreaBadge.Text = "  $($area.Scope)  "
    $lblAreaBadge.ForeColor = $area.Accent
    $lblAreaBadge.BackColor = [System.Drawing.Color]::FromArgb(
        [int][Math]::Min(255, [int]($area.Accent.R / 5) + 14),
        [int][Math]::Min(255, [int]($area.Accent.G / 5) + 14),
        [int][Math]::Min(255, [int]($area.Accent.B / 5) + 14))
    $lblSbTitle.Text = $area.Name
    $tbCrumb.Text = "/  $($area.Name)"

    # Reconstruir lista de tools (filtrada pelo search)
    Build-ToolList -AreaId $AreaId -Filter $txtSearch.Text

    # Reset runner
    Clear-Runner

    # Preseleccionar se pedido
    if ($PreselectTool) {
        $tool = $allTools | Where-Object { $_.Name -eq $PreselectTool -and $_.AreaId -eq $AreaId } | Select-Object -First 1
        if ($tool) { Load-Tool $tool }
    }

    $dashPanel.Visible = $false
    $toolPanel.Visible = $true
    $toolPanel.BringToFront()
}

function Clear-Runner {
    $toolHost.Controls.Clear()
    if ($script:CurrentToolHost) {
        try { $script:CurrentToolHost.Dispose() } catch { }
        $script:CurrentToolHost = $null
    }
    $lblRunnerTitle.Text = 'Seleccione uma ferramenta'
    $lblRunnerDesc.Text  = 'Escolha uma ferramenta na barra lateral para comecar.'
    $lblRunnerIcon.Text  = ''

    $ph = New-Object System.Windows.Forms.Label
    $ph.Text = "Clique numa ferramenta na barra lateral`npara carregar a interface."
    $ph.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 11)
    $ph.ForeColor = $script:Theme.TextFaint
    $ph.TextAlign = 'MiddleCenter'; $ph.Dock = 'Fill'; $ph.BackColor = $script:Theme.Bg
    $toolHost.Controls.Add($ph)
}

function Load-Tool {
    param($Tool)
    Write-AppLog "Load-Tool: $($Tool.Name)"
    $toolHost.Controls.Clear()

    $lblRunnerIcon.Text  = $Tool.CatIcon
    $lblRunnerIcon.ForeColor = $Tool.Accent
    $lblRunnerIcon.BackColor = [System.Drawing.Color]::FromArgb(
        [int][Math]::Min(255, [int]($Tool.Accent.R / 5) + 14),
        [int][Math]::Min(255, [int]($Tool.Accent.G / 5) + 14),
        [int][Math]::Min(255, [int]($Tool.Accent.B / 5) + 14))
    $lblRunnerTitle.Text = $Tool.Name
    $lblRunnerDesc.Text  = $Tool.Desc

    try {
        $tabPage = & $Tool.Factory
        $tc = New-Object System.Windows.Forms.TabControl
        $tc.Dock = 'Fill'
        $tc.Appearance = 'Normal'
        $tc.ItemSize = New-Object System.Drawing.Size(0,1)
        $tc.SizeMode = 'Fixed'
        $tc.TabPages.Add($tabPage) | Out-Null
        $toolHost.Controls.Add($tc)
        $script:CurrentToolHost = $tc

        try { Apply-DarkTheme -Control $tabPage -Theme $script:Theme }
        catch { Write-AppLog "Apply-DarkTheme ERRO: $($_.Exception.Message)" }
    } catch {
        Write-AppLog "Factory ERRO $($Tool.Name): $($_.Exception.Message)"
        $err = New-Object System.Windows.Forms.Label
        $err.Text = "Erro a carregar '$($Tool.Name)':`n$($_.Exception.Message)"
        $err.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 10)
        $err.ForeColor = $script:Theme.Danger
        $err.BackColor = $script:Theme.Bg
        $err.TextAlign = 'MiddleCenter'; $err.Dock = 'Fill'
        $toolHost.Controls.Add($err)
    }
}

function Build-ToolList {
    param([string]$AreaId, [string]$Filter = '')
    $toolListPanel.Controls.Clear()
    $y = 0
    foreach ($cat in (Get-AreaCategories $AreaId)) {
        $matching = $cat.Tools | Where-Object {
            -not $Filter -or $_.Name -like "*$Filter*" -or $_.Desc -like "*$Filter*"
        }
        if (-not $matching -or $matching.Count -eq 0) { continue }

        # Category header
        $catLbl = New-Object System.Windows.Forms.Label
        $catLbl.Text = $cat.Name.ToUpper()
        $catLbl.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 8, [System.Drawing.FontStyle]::Bold)
        $catLbl.ForeColor = $script:Theme.TextFaint
        $catLbl.AutoSize = $false
        $catLbl.Size = New-Object System.Drawing.Size(288, 20)
        $catLbl.Location = New-Object System.Drawing.Point(8, $y)
        $toolListPanel.Controls.Add($catLbl)
        $y += 24

        foreach ($tool in $matching) {
            $item = New-Object System.Windows.Forms.Panel
            $item.Size = New-Object System.Drawing.Size(288, 44)
            $item.Location = New-Object System.Drawing.Point(8, $y)
            $item.BackColor = $script:Theme.Panel
            $item.Cursor = 'Hand'
            $item.Tag = @{ Tool=$tool; CatIcon=$cat.Icon; CatName=$cat.Name; Accent=$cat.Color; AreaId=$AreaId }
            $toolListPanel.Controls.Add($item)

            $icon = New-Object System.Windows.Forms.Label
            $icon.Text = $cat.Icon
            $icon.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 8, [System.Drawing.FontStyle]::Bold)
            $icon.ForeColor = $cat.Color
            $icon.BackColor = [System.Drawing.Color]::FromArgb(
                [int][Math]::Min(255, [int]($cat.Color.R / 5) + 14),
                [int][Math]::Min(255, [int]($cat.Color.G / 5) + 14),
                [int][Math]::Min(255, [int]($cat.Color.B / 5) + 14))
            $icon.TextAlign = 'MiddleCenter'
            $icon.Size = New-Object System.Drawing.Size(26,26)
            $icon.Location = New-Object System.Drawing.Point(10, 9)
            $item.Controls.Add($icon)

            $nm = New-Object System.Windows.Forms.Label
            $nm.Text = $tool.Name
            $nm.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 9, [System.Drawing.FontStyle]::Bold)
            $nm.ForeColor = $script:Theme.Text
            $nm.BackColor = [System.Drawing.Color]::Transparent
            $nm.AutoSize = $true
            $nm.Location = New-Object System.Drawing.Point(44, 6)
            $item.Controls.Add($nm)

            $cn = New-Object System.Windows.Forms.Label
            $cn.Text = $cat.Name
            $cn.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 7)
            $cn.ForeColor = $script:Theme.TextFaint
            $cn.BackColor = [System.Drawing.Color]::Transparent
            $cn.AutoSize = $true
            $cn.Location = New-Object System.Drawing.Point(44, 24)
            $item.Controls.Add($cn)

            $hIn  = { $this.BackColor = $script:Theme.Panel2 }.GetNewClosure()
            $hOut = { $this.BackColor = $script:Theme.Panel }.GetNewClosure()
            $item.Add_MouseEnter($hIn); $item.Add_MouseLeave($hOut)
            foreach ($ch in $item.Controls) { $ch.Add_MouseEnter($hIn); $ch.Add_MouseLeave($hOut) }

            $clk = {
                param($sender, $e)
                $p = $sender; while ($p -and -not $p.Tag) { $p = $p.Parent }
                if ($p -and $p.Tag) {
                    $t = $p.Tag
                    Load-Tool @{ Name=$t.Tool.Name; Desc=$t.Tool.Desc; Factory=$t.Tool.Factory; CatIcon=$t.CatIcon; CatName=$t.CatName; Accent=$t.Accent; AreaId=$t.AreaId }
                }
            }.GetNewClosure()
            $item.Add_Click($clk)
            foreach ($ch in $item.Controls) { $ch.Add_Click($clk) }

            $y += 48
        }

        $y += 8
    }

    if ($y -eq 0) {
        $none = New-Object System.Windows.Forms.Label
        $none.Text = 'Sem ferramentas nesta area (ainda).'
        $none.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 9)
        $none.ForeColor = $script:Theme.TextFaint
        $none.AutoSize = $false
        $none.Size = New-Object System.Drawing.Size(288, 60)
        $none.Location = New-Object System.Drawing.Point(8, 8)
        $none.TextAlign = 'MiddleCenter'
        $toolListPanel.Controls.Add($none)
    }
}

# Wire search
$txtSearch.Add_TextChanged({
    if ($script:CurrentArea) {
        Build-ToolList -AreaId $script:CurrentArea.Id -Filter $txtSearch.Text
    }
})

# ============================================================
# ARRANQUE
# ============================================================
Show-Dashboard
Write-AppLog "Arrancar Application.Run"
try {
    [System.Windows.Forms.Application]::Run($form)
    Write-AppLog "Application.Run devolveu"
} catch {
    Write-AppLog "Application.Run ERRO: $($_.Exception.Message)"
    throw
}
