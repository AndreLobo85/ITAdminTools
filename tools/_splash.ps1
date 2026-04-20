# ============================================================
# _splash.ps1 - Splash de arranque (dark + teal novobanco)
# Baseado no design IT Admin Toolkit.html
# ============================================================

function Show-Splash {
    param([int]$DurationMs = 1800)

    $splash = New-Object System.Windows.Forms.Form
    $splash.Text = 'IT Admin Toolkit'
    $splash.FormBorderStyle = 'None'
    $splash.StartPosition = 'CenterScreen'
    $splash.Size = New-Object System.Drawing.Size(520, 380)
    $splash.BackColor = $script:Theme.Bg
    $splash.ShowInTaskbar = $false
    $splash.TopMost = $true

    # Logo teal (quadrado com N ao centro - approximacao do design NB)
    $logo = New-Object System.Windows.Forms.Panel
    $logo.Size = New-Object System.Drawing.Size(108, 108)
    $logo.Location = New-Object System.Drawing.Point((($splash.Width - 108) / 2), 48)
    $logo.BackColor = $script:Theme.Accent
    $splash.Controls.Add($logo)

    $logoLetter = New-Object System.Windows.Forms.Label
    $logoLetter.Text = 'N'
    $logoLetter.Dock = 'Fill'
    $logoLetter.TextAlign = 'MiddleCenter'
    $logoLetter.ForeColor = [System.Drawing.Color]::White
    $logoLetter.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 52, [System.Drawing.FontStyle]::Bold)
    $logo.Controls.Add($logoLetter)

    # Label "NOVOBANCO" (mono, letter-spacing simulado com espacos)
    $lblBrand = New-Object System.Windows.Forms.Label
    $lblBrand.Text = 'N O V O B A N C O'
    $lblBrand.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 9, [System.Drawing.FontStyle]::Regular)
    $lblBrand.ForeColor = $script:Theme.TextDim
    $lblBrand.AutoSize = $false
    $lblBrand.TextAlign = 'MiddleCenter'
    $lblBrand.Size = New-Object System.Drawing.Size($splash.Width, 18)
    $lblBrand.Location = New-Object System.Drawing.Point(0, 166)
    $splash.Controls.Add($lblBrand)

    # Titulo principal
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = 'IT Admin Toolkit'
    $lblTitle.Font = New-Object System.Drawing.Font($script:Theme.FontSans, 18, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $script:Theme.Text
    $lblTitle.AutoSize = $false
    $lblTitle.TextAlign = 'MiddleCenter'
    $lblTitle.Size = New-Object System.Drawing.Size($splash.Width, 32)
    $lblTitle.Location = New-Object System.Drawing.Point(0, 186)
    $splash.Controls.Add($lblTitle)

    # Checklist (6 items que vao aparecendo)
    $checks = @(
        'Kerberos ticket',
        'Domain controller',
        'Active Directory module',
        'Exchange Online module',
        'Role-based access',
        'UI modules'
    )

    $checkLabels = @()
    for ($i = 0; $i -lt $checks.Count; $i++) {
        $c = New-Object System.Windows.Forms.Label
        $c.Text = "  o  $($checks[$i])"
        $c.Font = New-Object System.Drawing.Font($script:Theme.FontMono, 8)
        $c.ForeColor = $script:Theme.TextFaint
        $c.AutoSize = $false
        $c.TextAlign = 'MiddleLeft'
        $c.Size = New-Object System.Drawing.Size(320, 16)
        $c.Location = New-Object System.Drawing.Point(100, (230 + ($i * 18)))
        $c.Visible = $false
        $splash.Controls.Add($c)
        $checkLabels += $c
    }

    # Barra de progresso fina (bottom)
    $progressBar = New-Object System.Windows.Forms.Panel
    $progressBar.BackColor = $script:Theme.Border
    $progressBar.Size = New-Object System.Drawing.Size(420, 2)
    $progressBar.Location = New-Object System.Drawing.Point((($splash.Width - 420) / 2), 352)
    $splash.Controls.Add($progressBar)

    $progressFill = New-Object System.Windows.Forms.Panel
    $progressFill.BackColor = $script:Theme.Accent
    $progressFill.Size = New-Object System.Drawing.Size(0, 2)
    $progressFill.Location = New-Object System.Drawing.Point(0, 0)
    $progressBar.Controls.Add($progressFill)

    # Estado partilhado entre o handler do timer e o resto do codigo.
    # Usar indexer syntax ($state['Key']) evita problemas de resolucao de
    # propriedade em closures de event handlers WinForms.
    $state = @{
        CheckIdx = 0
        Elapsed  = 0
        TickMs   = 30
    }
    $perCheckMs = [int]($DurationMs / ($checks.Count + 2))
    $okColor = $script:Theme.Ok

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 30
    $timer.Add_Tick({
        $state['Elapsed'] = $state['Elapsed'] + $state['TickMs']
        $pct = [Math]::Min(1.0, [double]$state['Elapsed'] / [double]$DurationMs)
        $progressFill.Size = New-Object System.Drawing.Size([int](420 * $pct), 2)

        $targetIdx = [int]([Math]::Floor([double]$state['Elapsed'] / [double]$perCheckMs))
        $limit = [Math]::Min($targetIdx, $checkLabels.Count)
        while ($state['CheckIdx'] -lt $limit) {
            $idx = $state['CheckIdx']
            $cl = $checkLabels[$idx]
            $cl.Visible = $true
            $cl.Text = "  +  $($checks[$idx])"
            $cl.ForeColor = $okColor
            $state['CheckIdx'] = $idx + 1
        }

        if ($state['Elapsed'] -ge $DurationMs) {
            $timer.Stop()
            $splash.Close()
        }
    }.GetNewClosure())

    $timer.Start()
    [void]$splash.ShowDialog()
    $timer.Dispose()
    $splash.Dispose()
}
