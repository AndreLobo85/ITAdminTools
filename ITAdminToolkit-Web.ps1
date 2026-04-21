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

# ========== Runspace runner: executa scripts em runspace separado (cancelavel) ==========
$script:RunningPS = $null
$script:RunningRS = $null

function Format-PSErrorRecord {
    param($ErrorRecord)
    $out = @()
    $msg = if ($ErrorRecord.Exception) { $ErrorRecord.Exception.Message } else { "$ErrorRecord" }
    $out += "[ERR] $msg"
    if ($ErrorRecord.InvocationInfo -and $ErrorRecord.InvocationInfo.PositionMessage) {
        foreach ($l in ($ErrorRecord.InvocationInfo.PositionMessage -split "`r?`n")) {
            if ($l.Trim()) { $out += "    $l" }
        }
    }
    if ($ErrorRecord.CategoryInfo) {
        $out += "    + CategoryInfo          : " + $ErrorRecord.CategoryInfo.ToString()
    }
    if ($ErrorRecord.FullyQualifiedErrorId) {
        $out += "    + FullyQualifiedErrorId : " + $ErrorRecord.FullyQualifiedErrorId
    }
    if ($ErrorRecord.Exception -and $ErrorRecord.Exception.InnerException) {
        $out += "    + InnerException        : " + $ErrorRecord.Exception.InnerException.Message
    }
    return $out
}

function Stop-RunningScript {
    if ($script:RunningPS) {
        Write-AppLog "CANCEL_TOOL: a parar runspace activo"
        try { $script:RunningPS.BeginStop($null, $null) | Out-Null } catch { Write-AppLog "Stop falhou: $($_.Exception.Message)" }
    } else {
        Write-AppLog "CANCEL_TOOL: nenhum script em execucao"
    }
}

function Push-StreamLine {
    param([string]$Line)
    if (-not $script:CurrentReqId) { return }
    if (-not $script:WebViewCtrl -or -not $script:WebViewCtrl.CoreWebView2) { return }
    try {
        $payload = @{ id = $script:CurrentReqId; type = 'TOOL_LINE'; line = $Line }
        $json = $payload | ConvertTo-Json -Compress
        $script:WebViewCtrl.CoreWebView2.PostWebMessageAsJson($json)
    } catch {
        Write-AppLog "Push-StreamLine falhou: $($_.Exception.Message)"
    }
}

function Invoke-ScriptInRunspace {
    [CmdletBinding(DefaultParameterSetName='File')]
    param(
        [Parameter(Mandatory, ParameterSetName='File')][string]$ScriptPath,
        [Parameter(Mandatory, ParameterSetName='Block')][scriptblock]$ScriptBlock,
        [hashtable]$Parameters = @{}
    )

    # Garantir que so corre 1 script de cada vez
    if ($script:RunningPS) {
        try { $script:RunningPS.Stop() } catch {}
        try { $script:RunningPS.Dispose() } catch {}
        $script:RunningPS = $null
    }
    if ($script:RunningRS) {
        try { $script:RunningRS.Close() } catch {}
        try { $script:RunningRS.Dispose() } catch {}
        $script:RunningRS = $null
    }

    # MTA por defeito (AD/ADSI funciona melhor em MTA; STA so e necessario
    # para popups WPF/WinForms criados dentro do runspace e o EXO ja gere isso).
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'MTA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()

    $ps = [powershell]::Create()
    $ps.Runspace = $rs

    if ($PSCmdlet.ParameterSetName -eq 'File') {
        # Wrapper scriptblock: invoca o ficheiro via `& $Path @BoundParams`.
        # Isto comporta-se igual a correr o script na consola PS, e o
        # splatting garante que os parametros fazem bind ao param() block
        # do script. (AddCommand com path absoluto nem sempre resolve o .ps1
        # como ExternalScript e pode falhar silenciosamente no parameter
        # binding.)
        $wrapper = @'
param($__Path, $__BoundParams)
if (-not (Test-Path -LiteralPath $__Path)) {
    throw "Script nao encontrado no runspace: $__Path"
}
if ($__BoundParams -is [hashtable] -and $__BoundParams.Count -gt 0) {
    & $__Path @__BoundParams
} else {
    & $__Path
}
'@
        [void]$ps.AddScript($wrapper)
        [void]$ps.AddParameter('__Path', $ScriptPath)
        [void]$ps.AddParameter('__BoundParams', $Parameters)
    } else {
        # Executa um scriptblock literal (texto). Re-cria o scriptblock dentro
        # do runspace para evitar capturar variaveis do scope-pai.
        [void]$ps.AddScript($ScriptBlock.ToString())
        foreach ($k in $Parameters.Keys) {
            [void]$ps.AddParameter($k, $Parameters[$k])
        }
    }

    $script:RunningPS = $ps
    $script:RunningRS = $rs

    $tag = if ($PSCmdlet.ParameterSetName -eq 'File') { $ScriptPath } else { '<scriptblock>' }
    Write-AppLog "Runspace start: $tag  params=$($Parameters | ConvertTo-Json -Compress -Depth 5)"

    $output = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
    $async  = $ps.BeginInvoke($null, $output)

    # Loop nao-bloqueante: pumpa UI ate completar e faz streaming live
    # de cada nova linha de output para a UI (TOOL_LINE).
    $lastIdx = 0
    $errLastIdx = 0
    while (-not $async.IsCompleted) {
        [System.Windows.Forms.Application]::DoEvents()
        # Stream output novo
        while ($lastIdx -lt $output.Count) {
            Push-StreamLine ("$($output[$lastIdx])")
            $lastIdx++
        }
        # Stream erros novos
        while ($errLastIdx -lt $ps.Streams.Error.Count) {
            foreach ($eline in (Format-PSErrorRecord $ps.Streams.Error[$errLastIdx])) {
                Push-StreamLine $eline
            }
            $errLastIdx++
        }
        Start-Sleep -Milliseconds 60
    }

    # Flush remanescente
    while ($lastIdx -lt $output.Count) {
        Push-StreamLine ("$($output[$lastIdx])")
        $lastIdx++
    }
    while ($errLastIdx -lt $ps.Streams.Error.Count) {
        foreach ($eline in (Format-PSErrorRecord $ps.Streams.Error[$errLastIdx])) {
            Push-StreamLine $eline
        }
        $errLastIdx++
    }

    $lines = @()
    foreach ($o in $output) { $lines += "$o" }

    $threwTerminating = $null
    try {
        [void]$ps.EndInvoke($async)
    } catch {
        $threwTerminating = $_
        Write-AppLog "EndInvoke threw: $($_.Exception.Message)"
    }

    # Erros nao-terminantes (ja foram streamed mas vao tambem no resultado final)
    if ($ps.Streams.Error -and $ps.Streams.Error.Count -gt 0) {
        $lines += ''
        foreach ($err in $ps.Streams.Error) {
            $lines += (Format-PSErrorRecord $err)
        }
    }
    if ($ps.Streams.Warning) {
        foreach ($w in $ps.Streams.Warning) {
            $line = "[WARN] $w"
            $lines += $line
            Push-StreamLine $line
        }
    }
    if ($ps.Streams.Information) {
        foreach ($i in $ps.Streams.Information) {
            $line = "$i"
            $lines += $line
            Push-StreamLine $line
        }
    }

    $state = $ps.InvocationStateInfo.State
    $cancelled = ($state -eq 'Stopped' -or $state -eq 'Stopping')

    if ($threwTerminating -and -not $cancelled) {
        $lines += ''
        $errLines = Format-PSErrorRecord $threwTerminating
        foreach ($el in $errLines) { $lines += $el; Push-StreamLine $el }
    }

    if ($cancelled) {
        $lines += ''
        $lines += '[WARN] Execucao cancelada pelo utilizador (Stop).'
        Push-StreamLine '[WARN] Execucao cancelada pelo utilizador (Stop).'
    }

    # Fallback diagnostico: se o script terminou normalmente mas nao produziu
    # NADA, o utilizador fica a olhar para um terminal vazio sem saber porque.
    # Isto normalmente significa que o script nao encontrou dados (ex: user
    # inexistente no AD) e retornou silenciosamente. Dar-lhe uma pista util.
    if (-not $cancelled -and -not $threwTerminating -and
        ($ps.Streams.Error.Count -eq 0) -and ($output.Count -eq 0)) {
        $diag = @()
        $diag += '[WARN] O script terminou sem produzir output e sem erros.'
        if ($PSCmdlet.ParameterSetName -eq 'File') {
            $diag += "       Script: $ScriptPath"
            if ($Parameters -and $Parameters.Count -gt 0) {
                $pStr = ($Parameters.GetEnumerator() | ForEach-Object { "-$($_.Key) '$($_.Value)'" }) -join ' '
                $diag += "       Parametros passados ao script: $pStr"
            } else {
                $diag += '       (nenhum parametro passado)'
            }
            $diag += '       Causas provaveis:'
            $diag += '       - O valor pesquisado nao existe no AD (ex: samaccountname errado)'
            $diag += '       - O script fez return cedo porque os parametros nao fizeram bind'
            $diag += '       - Sem ligacao ao dominio (tenta "nltest /dsgetdc:" numa consola PS)'
        }
        foreach ($d in $diag) { $lines += $d; Push-StreamLine $d }
    }

    Write-AppLog "Runspace done: state=$state  lines=$($lines.Count)  errors=$($ps.Streams.Error.Count)  outputCount=$($output.Count)"

    try { $ps.Dispose() } catch {}
    try { $rs.Close(); $rs.Dispose() } catch {}
    $script:RunningPS = $null
    $script:RunningRS = $null

    return ,$lines
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
            $scriptPath = Join-Path $ToolsDir 'scripts\get-diag-info-aw2.ps1'
            if (-not (Test-Path $scriptPath)) { throw "Script nao encontrado: $scriptPath" }
            $p = @{}
            if ($username) { $p['userName'] = $username }
            if ($email)    { $p['email']    = $email }
            $lines = Invoke-ScriptInRunspace -ScriptPath $scriptPath -Parameters $p
            return @{ lines = $lines }
        }

        'GroupInfo' {
            $g = [string]$Params.groupName
            if (-not $g) { throw 'Indique o nome do grupo.' }
            $scriptPath = Join-Path $ToolsDir 'scripts\get-diag-info-group.ps1'
            if (-not (Test-Path $scriptPath)) { throw "Script nao encontrado: $scriptPath" }
            $lines = Invoke-ScriptInRunspace -ScriptPath $scriptPath -Parameters @{ groupName = $g }
            return @{ lines = $lines }
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

            $upns = $raw -split '[\r\n,;]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique

            # Corre tudo num runspace separado:
            #  1) Valida via Microsoft.Graph que o user tem role Exchange Admin (ou superior) ACTIVA
            #  2) So depois faz Connect-ExchangeOnline
            #  3) Itera os UPNs e devolve estatisticas
            #  Auth popups (Graph + EXO) sao spawnados fora da UI thread,
            #  por isso aparecem visiveis e a UI nao bloqueia.
            $sb = {
                param([string]$AdminUpn, [string[]]$Upns, [bool]$IncludeRecov)

                # ---- 1. Validacao de role (Microsoft Graph) ----
                "[INFO] A validar roles activas do administrador (Microsoft Graph)..."
                if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
                    "[ERR] Modulo Microsoft.Graph.Authentication nao instalado."
                    "      Para instalar: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
                    return
                }
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

                $graphScopes = @('RoleManagement.Read.Directory','Directory.Read.All','User.Read')
                $graphParams = @{ Scopes = $graphScopes; NoWelcome = $true; ErrorAction = 'Stop' }
                try {
                    Connect-MgGraph @graphParams | Out-Null
                } catch {
                    "[ERR] Connect-MgGraph falhou: $($_.Exception.Message)"
                    return
                }

                # transitiveMemberOf/microsoft.graph.directoryRole devolve so roles ACTIVAS
                # (PIM eligible nao activadas nao aparecem aqui)
                try {
                    $resp = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.directoryRole' -ErrorAction Stop
                    $activeRoles = @($resp.value | ForEach-Object { $_.displayName })
                } catch {
                    "[ERR] Falha a ler roles via Graph: $($_.Exception.Message)"
                    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
                    return
                }

                $allowedRoles = @(
                    'Exchange Administrator',
                    'Exchange Recipient Administrator',
                    'Global Administrator',
                    'Global Reader'
                )
                $matched = @($activeRoles | Where-Object { $_ -in $allowedRoles })

                if ($activeRoles.Count -eq 0) {
                    "[ERR] Nao tens nenhuma directory role ACTIVA neste momento."
                    "      Activa em PIM: https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade"
                    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
                    return
                }

                "[INFO] Roles activas detectadas: $($activeRoles -join ', ')"

                if ($matched.Count -eq 0) {
                    "[ERR] Nenhuma role com acesso a Exchange Online esta ACTIVA."
                    "      Roles aceites: $($allowedRoles -join ', ')"
                    "      Activa Exchange Administrator em PIM antes de tentar de novo."
                    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
                    return
                }

                "[OK] Role permitida activa: $($matched -join ', ')"

                # ---- 2. Connect Exchange Online ----
                if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
                    "[ERR] Modulo ExchangeOnlineManagement nao instalado."
                    "      Install-Module ExchangeOnlineManagement -Scope CurrentUser"
                    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
                    return
                }
                Import-Module ExchangeOnlineManagement -ErrorAction Stop

                $alreadyConn = $null
                try { $alreadyConn = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected' | Select-Object -First 1 } catch {}
                if (-not $alreadyConn) {
                    "[INFO] A ligar a Exchange Online (popup de autenticacao vai aparecer)..."
                    $exoParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
                    if ($AdminUpn) { $exoParams.UserPrincipalName = $AdminUpn }
                    try {
                        Connect-ExchangeOnline @exoParams
                    } catch {
                        "[ERR] Connect-ExchangeOnline falhou: $($_.Exception.Message)"
                        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
                        return
                    }
                    "[OK] Ligado a Exchange Online"
                } else {
                    "[INFO] Sessao Exchange Online ja activa, a reutilizar."
                }

                # ---- 3. Iterar UPNs ----
                ""
                "[INFO] Exchange Online - $($Upns.Count) mailbox(es)"
                ""
                foreach ($u in $Upns) {
                    try {
                        $stats = Get-MailboxStatistics -Identity $u -ErrorAction Stop
                        "[OK] $u"
                        "  DisplayName         : $($stats.DisplayName)"
                        "  StorageLimitStatus  : $($stats.StorageLimitStatus)"
                        "  TotalItemSize       : $($stats.TotalItemSize)"
                        "  TotalDeletedItemSize: $($stats.TotalDeletedItemSize)"
                        "  ItemCount           : $($stats.ItemCount)"
                        "  DeletedItemCount    : $($stats.DeletedItemCount)"
                        "  LastLogonTime       : $($stats.LastLogonTime)"
                        if ($IncludeRecov) {
                            try {
                                $rec = Get-MailboxStatistics -Identity $u -FolderScope RecoverableItems -ErrorAction Stop
                                "  RecoverableItems    : $($rec.TotalItemSize)"
                            } catch {
                                "  RecoverableItems    : (erro: $($_.Exception.Message))"
                            }
                        }
                    } catch {
                        "[ERR] $u :: $($_.Exception.Message)"
                    }
                    ""
                }
            }

            $lines = Invoke-ScriptInRunspace -ScriptBlock $sb -Parameters @{
                AdminUpn     = $adminUpn
                Upns         = ,$upns
                IncludeRecov = $includeRecov
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
                    # ID disponivel para Push-StreamLine durante a execucao
                    $script:CurrentReqId = $msg.id
                    try {
                        $result = Invoke-ToolRequest -ToolId $msg.toolId -Params $params
                    } finally {
                        $script:CurrentReqId = $null
                    }
                    $reply = @{ id = $msg.id; ok = $true; result = $result }
                }
                'GET_CONTEXT' {
                    $reply = @{ id = $msg.id; ok = $true; result = @{
                        host = $env:COMPUTERNAME; user = "$env:USERDOMAIN\$env:USERNAME"
                        adAvailable = [bool]$script:ADAvailable
                    }}
                }
                'CANCEL_TOOL' {
                    Stop-RunningScript
                    $reply = @{ id = $msg.id; ok = $true; result = @{ cancelled = $true } }
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
