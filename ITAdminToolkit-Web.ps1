<#
.SYNOPSIS
    IT Admin Toolkit (versao WebView2) - Design Claude pixel-perfect.

.DESCRIPTION
    Cria uma janela WinForms com controlo WebView2 em Dock=Fill que carrega
    o design HTML do Claude (webui/index.html). O JS comunica com os scripts
    PowerShell via postMessage.

.REQUIREMENTS
    - Windows 10/11 ou Server 2016+
    - Microsoft Edge WebView2 Runtime (verificado em runtime)
    - DLLs em ./lib/ (correr Setup-WebView2.ps1 primeiro)
    - Powershell 5.1+
#>

$ErrorActionPreference = 'Stop'

# ========== Logging + trap ==========
$script:LogPath = Join-Path $env:TEMP "ITAdminToolkit-Web-$((Get-Date).ToString('yyyyMMdd_HHmmss')).log"
function Write-AppLog { param([string]$m) try { Add-Content -Path $script:LogPath -Value "[$((Get-Date).ToString('HH:mm:ss.fff'))] $m" -Encoding UTF8 } catch {} }
Write-AppLog "=== Arranque (Web) ==="
trap {
    $e = $_
    $msg = "Erro critico: $($e.Exception.Message)`n`n$($e.ScriptStackTrace)`n`nLog: $script:LogPath"
    Write-AppLog "TRAP: $msg"
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($msg, 'IT Admin Toolkit - Erro', 'OK', 'Error') | Out-Null
    } catch {}
    exit 1
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ToolsDir  = Join-Path $ScriptDir 'tools'
$LibDir    = Join-Path $ScriptDir 'lib'
$WebUiDir  = Join-Path $ScriptDir 'webui'

# ========== Validar que DLLs existem ==========
$coreDll  = Join-Path $LibDir 'Microsoft.Web.WebView2.Core.dll'
$wfDll    = Join-Path $LibDir 'Microsoft.Web.WebView2.WinForms.dll'
if (-not (Test-Path $coreDll) -or -not (Test-Path $wfDll)) {
    [System.Windows.Forms.MessageBox]::Show(
        "DLLs WebView2 nao encontradas em:`n$LibDir`n`nCorre primeiro Setup-WebView2.ps1",
        'Setup em falta', 'OK', 'Error') | Out-Null
    exit 1
}

# ========== Load WinForms + WebView2 ==========
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# DLL dir para o WebView2Loader nativo ser encontrado
[System.Environment]::CurrentDirectory = $LibDir
Add-Type -Path $coreDll
Add-Type -Path $wfDll
Write-AppLog "WebView2 DLLs carregadas"

# ========== Carregar ferramentas ==========
foreach ($f in '_common.ps1','UserInfo.ps1','GroupInfo.ps1','ADGroupAuditor.ps1','ShareAuditor.ps1','MailboxStats.ps1','SharePointSite.ps1') {
    Write-AppLog "Carregar $f"
    . (Join-Path $ToolsDir $f)
}

# ========== Dispatcher: recebe (toolId, params) e devolve @{ lines = [...] } ==========
function Invoke-ToolRequest {
    param([string]$ToolId, [hashtable]$Params)

    Write-AppLog "RUN_TOOL: $ToolId  params=$($Params | ConvertTo-Json -Compress -Depth 5)"

    switch ($ToolId) {
        'UserInfo' {
            $username = if ($Params.username) { [string]$Params.username } else { '' }
            $email    = if ($Params.email) { [string]$Params.email } else { '' }
            if (-not $username -and -not $email) { throw 'Indique username ou email.' }
            $info = UI_Get-UserInfo -Username $username -Email $email
            if (-not $info) { return @{ lines = @('[WARN] User nao encontrado.') } }
            $report = UI_Format-Report $info
            return @{ lines = ($report -split "`r?`n") }
        }

        'GroupInfo' {
            $g = [string]$Params.groupName
            if (-not $g) { throw 'Indique o nome do grupo.' }
            $info = GI_Get-GroupInfo -GroupName $g
            if (-not $info) { return @{ lines = @('[WARN] Grupo nao encontrado.') } }
            $report = GI_Format-Report $info
            return @{ lines = ($report -split "`r?`n") }
        }

        'ADGroupAuditor' {
            $mode = if ($Params.mode) { [string]$Params.mode } else { 'Suffix' }
            $raw  = [string]$Params.terms
            $active = [bool]$Params.activeOnly
            if (-not $raw) { throw 'Indique termos.' }
            $terms = $raw -split '[,;\r\n]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            $results = ADG_Invoke-Audit -Terms $terms -SearchMode $mode -ActiveOnly $active -OnProgress $null -PumpUI $null
            $lines = @()
            $lines += "[INFO] Total linhas: $($results.Count)"
            $targets = $results | Select-Object -ExpandProperty TargetGroup -Unique
            $lines += "[INFO] Grupos-alvo encontrados: $($targets.Count)"
            $users = $results | Where-Object MemberType -eq 'User' | Select-Object -ExpandProperty SamAccountName -Unique
            $lines += "[INFO] Users distintos: $($users.Count)"
            $lines += ''
            foreach ($g in $targets) {
                $lines += "[OK] === $g ==="
                $groupRows = $results | Where-Object TargetGroup -eq $g
                foreach ($r in $groupRows) {
                    $indent = '  ' * $r.Depth
                    $enStr = if ($null -ne $r.Enabled) { if ($r.Enabled) { 'OK' } else { 'disabled' } } else { '' }
                    $lines += "$indent  $($r.MemberType)  $($r.SamAccountName)  $($r.DisplayName)  $enStr"
                }
                $lines += ''
            }
            return @{ lines = $lines }
        }

        'ShareAuditor' {
            $path = [string]$Params.path
            $recurse = [bool]$Params.recurse
            $depth = if ($Params.depth) { [int]$Params.depth } else { 3 }
            $onlyExpl = [bool]$Params.onlyExplicit
            if (-not $path) { throw 'Indique o caminho do share.' }
            if (-not (Test-Path -LiteralPath $path)) { throw "Caminho nao acessivel: $path" }
            $results = SA_Invoke-Audit -SharePath $path -Recurse $recurse -Depth $depth -OnlyExplicit $onlyExpl -OnProgress $null -PumpUI $null
            $lines = @()
            $lines += "[INFO] Linhas: $($results.Count)"
            $folders = ($results | Select-Object -ExpandProperty Folder -Unique).Count
            $users   = ($results | Where-Object Type -eq 'User' | Select-Object -ExpandProperty Member -Unique).Count
            $lines += "[INFO] Pastas: $folders | Users distintos: $users"
            $lines += ''
            $byFolder = $results | Group-Object Folder
            foreach ($gf in $byFolder) {
                $lines += "[OK] Pasta: $($gf.Name)"
                $byP = $gf.Group | Group-Object Principal
                foreach ($p in $byP) {
                    $fr = $p.Group[0]
                    $lines += "  Principal: $($fr.Principal)  [$($fr.Inherited)]"
                    $lines += "  Permissoes: $($fr.Permissions)"
                    foreach ($m in $p.Group) {
                        if ($m.Type -eq 'User') {
                            $en = if ($null -ne $m.Enabled) { if ($m.Enabled) { 'OK' } else { 'disabled' } } else { '' }
                            $lines += "    - $($m.Member)  $($m.DisplayName)  $en"
                        }
                    }
                }
                $lines += ''
            }
            return @{ lines = $lines }
        }

        'MailboxStats' {
            $adminUpn = [string]$Params.adminUpn
            $raw = [string]$Params.upns
            $includeRecov = [bool]$Params.includeRecov
            if (-not $raw) { throw 'Indique pelo menos um UPN.' }

            # Connect se ainda nao ligado
            $conn = $null
            try { $conn = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected' | Select-Object -First 1 } catch {}
            if (-not $conn) {
                Import-Module ExchangeOnlineManagement -ErrorAction Stop
                $p = @{ ShowBanner = $false; ErrorAction = 'Stop' }
                if ($adminUpn) { $p.UserPrincipalName = $adminUpn }
                Connect-ExchangeOnline @p
            }

            $upns = $raw -split '[\r\n,;]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique
            $lines = @()
            $lines += "[INFO] Exchange Online - $($upns.Count) mailbox(es)"
            $lines += ''
            foreach ($u in $upns) {
                $r = MBX_Get-Stats -Upn $u -IncludeRecoverable $includeRecov
                if ($r.Status -eq 'OK') {
                    $lines += "[OK] $u"
                    $lines += "  DisplayName         : $($r.DisplayName)"
                    $lines += "  StorageLimitStatus  : $($r.StorageLimitStatus)"
                    $lines += "  TotalItemSize       : $($r.TotalItemSize)"
                    $lines += "  TotalDeletedItemSize: $($r.TotalDeletedItemSize)"
                    $lines += "  ItemCount           : $($r.ItemCount)"
                    $lines += "  DeletedItemCount    : $($r.DeletedItemCount)"
                    $lines += "  LastLogonTime       : $($r.LastLogonTime)"
                    if ($includeRecov) { $lines += "  RecoverableItems    : $($r.RecoverableItems)" }
                } else {
                    $lines += "[ERR] $u :: $($r.Status)"
                }
                $lines += ''
            }
            return @{ lines = $lines }
        }

        'SharePointSite' {
            $siteUrl = [string]$Params.siteUrl
            $adminUrl = [string]$Params.tenantAdminUrl
            $includeOwners = if ($null -ne $Params.includeOwners) { [bool]$Params.includeOwners } else { $true }
            if (-not $siteUrl) { throw 'Indique o URL do site.' }
            $info = SP_Invoke-SiteAudit -SiteUrl $siteUrl -TenantAdminUrl $adminUrl -IncludeOwners $includeOwners
            return @{ lines = (SP_Format-Report $info) }
        }

        default { throw "Ferramenta desconhecida: $ToolId" }
    }
}

# ========== FORM + WebView2 ==========
$form = New-Object System.Windows.Forms.Form
$form.Text = 'IT Admin Toolkit - novobanco'
$form.Size = New-Object System.Drawing.Size(1360, 900)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(1024, 720)
$form.BackColor = [System.Drawing.Color]::FromArgb(12,16,20)

$wv = New-Object Microsoft.Web.WebView2.WinForms.WebView2
$wv.Dock = 'Fill'
$form.Controls.Add($wv)

$userDataFolder = Join-Path $env:LOCALAPPDATA 'ITAdminToolkit\WebView2UserData'
if (-not (Test-Path $userDataFolder)) { New-Item $userDataFolder -ItemType Directory -Force | Out-Null }

$script:WebViewReady = $false
$script:WebViewCtrl = $wv

# Handler de inicializacao
$wv.add_CoreWebView2InitializationCompleted({
    param($sender, $e)
    if (-not $e.IsSuccess) {
        Write-AppLog "Init ERRO: $($e.InitializationException.Message)"
        return
    }
    Write-AppLog "CoreWebView2 inicializado"

    # Expor webui local como https://app.local/
    $sender.CoreWebView2.SetVirtualHostNameToFolderMapping(
        'app.local', $WebUiDir, [Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind]::Allow
    )

    # DevTools no arranque (util para debug; podes remover depois)
    # $sender.CoreWebView2.OpenDevToolsWindow()

    # Handler de mensagens do JS
    $sender.CoreWebView2.add_WebMessageReceived({
        param($s, $args)
        $raw = $args.TryGetWebMessageAsString()
        if (-not $raw) { return }
        Write-AppLog "WebMsg recv: $raw"
        $msg = $null
        try { $msg = $raw | ConvertFrom-Json } catch {
            Write-AppLog "JSON parse falhou: $($_.Exception.Message)"
            return
        }
        $reply = $null
        try {
            switch ($msg.type) {
                'RUN_TOOL' {
                    $params = @{}
                    if ($msg.params) {
                        foreach ($prop in $msg.params.PSObject.Properties) {
                            $params[$prop.Name] = $prop.Value
                        }
                    }
                    $result = Invoke-ToolRequest -ToolId $msg.toolId -Params $params
                    $reply = @{ id = $msg.id; ok = $true; result = $result }
                }
                'GET_CONTEXT' {
                    $reply = @{ id = $msg.id; ok = $true; result = @{
                        host = $env:COMPUTERNAME; user = "$env:USERDOMAIN\$env:USERNAME"
                        adAvailable = [bool]$script:ADAvailable
                    }}
                }
                default { throw "tipo desconhecido: $($msg.type)" }
            }
        } catch {
            Write-AppLog "Dispatcher ERRO: $($_.Exception.Message)"
            $reply = @{ id = $msg.id; ok = $false; error = $_.Exception.Message }
        }
        if ($reply) {
            $json = $reply | ConvertTo-Json -Depth 10 -Compress
            $s.PostWebMessageAsJson($json)
        }
    })

    # Navegar
    $sender.CoreWebView2.Navigate('https://app.local/index.html')
    $script:WebViewReady = $true
    Write-AppLog "Navegacao iniciada"
})

# Iniciar CoreWebView2 com environment customizado (userDataFolder).
# Cria o environment sincronamente para evitar Task.ContinueWith (PS nao
# desambigua as overloads). Depois passa-o a EnsureCoreWebView2Async.
Write-AppLog "A criar CoreWebView2Environment em $userDataFolder"
$envTask = [Microsoft.Web.WebView2.Core.CoreWebView2Environment]::CreateAsync($null, $userDataFolder)
$envTask.Wait()
$wvEnv = $envTask.Result
Write-AppLog "Environment criado. EnsureCoreWebView2Async..."
$wv.EnsureCoreWebView2Async($wvEnv) | Out-Null

Write-AppLog "Application.Run..."
[System.Windows.Forms.Application]::Run($form)
Write-AppLog "Application.Run devolveu"
