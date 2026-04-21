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

# Log da versao Core WebView2 carregada (SDK 1.0.2365.x tem bug conhecido
# em WebMessageReceived - WebView2Feedback #4441)
try {
    $coreDllVer = (Get-Item $coreDll).VersionInfo.FileVersion
    $wfDllVer   = (Get-Item $wfDll).VersionInfo.FileVersion
    Write-AppLog "WebView2.Core.dll version: $coreDllVer"
    Write-AppLog "WebView2.WinForms.dll version: $wfDllVer"
    if ($coreDllVer -like '1.0.2365.*') {
        Write-AppLog "AVISO: SDK 1.0.2365.x tem bug conhecido em WebMessageReceived!"
    }
} catch { Write-AppLog "Nao consegui ler versao WebView2 DLL: $($_.Exception.Message)" }

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

    # Diagnosticos visiveis no terminal da app (chegam via streaming)
    $bitness = if ([IntPtr]::Size -eq 8) { '64-bit' } else { '32-bit' }
    Push-StreamLine "[DIAG] Host: PS $($PSVersionTable.PSVersion) $bitness, PID $PID, runspace apt=$($rs.ApartmentState)"
    if ($PSCmdlet.ParameterSetName -eq 'File') {
        Push-StreamLine "[DIAG] Script: $ScriptPath"
        if ($Parameters -and $Parameters.Count -gt 0) {
            $pStr = ($Parameters.GetEnumerator() | ForEach-Object { "-$($_.Key) '$($_.Value)'" }) -join ' '
            Push-StreamLine "[DIAG] Parametros: $pStr"
        }
    }

    # Output collection com input explicito (passar `$null` faz o binder de PS
    # 5.1 resolver para a overload errada e em alguns casos o BeginInvoke
    # devolve sem executar nada)
    $inputColl = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
    $inputColl.Complete()
    $output = New-Object 'System.Management.Automation.PSDataCollection[psobject]'

    Push-StreamLine "[DIAG] A chamar BeginInvoke..."
    try {
        $async = $ps.BeginInvoke($inputColl, $output)
    } catch {
        $msg = "[ERR] BeginInvoke lancou: $($_.Exception.Message)"
        Write-AppLog $msg
        Push-StreamLine $msg
        try { $ps.Dispose() } catch {}
        try { $rs.Close(); $rs.Dispose() } catch {}
        $script:RunningPS = $null
        $script:RunningRS = $null
        return ,@($msg)
    }
    Push-StreamLine "[DIAG] Runspace a correr (state=$($ps.InvocationStateInfo.State))"

    # Loop nao-bloqueante: pumpa UI ate completar e faz streaming live
    # de cada nova linha de output para a UI (TOOL_LINE).
    $lastIdx = 0
    $errLastIdx = 0
    $startT = [datetime]::UtcNow
    $lastPingSec = -1
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
        # Heartbeat cada segundo — prova que o pump de UI esta a correr
        $secElapsed = [int]([datetime]::UtcNow - $startT).TotalSeconds
        if ($secElapsed -gt $lastPingSec -and $secElapsed -ge 2) {
            $lastPingSec = $secElapsed
            Push-StreamLine "[DIAG] ... t=${secElapsed}s state=$($ps.InvocationStateInfo.State) output=$($output.Count) errors=$($ps.Streams.Error.Count)"
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

# ========== Execucao externa: spawn de processo powershell.exe separado =========
# Isola completamente o script do processo da app. Equivalente a abrir uma
# consola nova e correr `powershell.exe -File script.ps1 -userName t05352`.
# Forca PowerShell 5.1 64-bit (SysNative) para evitar WOW64 redirection.
function Invoke-ScriptExternal {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [hashtable]$Parameters = @{},
        [int]$TimeoutSeconds = 90
    )

    $lines = @()

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        return ,@("[ERR] Script nao encontrado: $ScriptPath")
    }

    # Resolve PowerShell.exe 64-bit.
    $psExe = Join-Path $env:SystemRoot 'SysNative\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path $psExe)) {
        $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    }
    if (-not (Test-Path $psExe)) {
        return ,@("[ERR] powershell.exe nao encontrado")
    }

    # Ficheiros temporarios para stdout/stderr. Esta abordagem evita o
    # deadlock classico de redirect com pipes (sem event pump) e tambem
    # o Register-ObjectEvent (cuja -Action depende da event queue do PS
    # principal - que esta bloqueada enquanto esperamos pela saida).
    $tmpBase   = Join-Path $env:TEMP "ITAdminToolkit-run-$([guid]::NewGuid().ToString('N').Substring(0,8))"
    $tmpStdout = "$tmpBase.out.txt"
    $tmpStderr = "$tmpBase.err.txt"

    # Construir argumentos
    $argList = @('-NoProfile','-ExecutionPolicy','Bypass','-NonInteractive','-File',$ScriptPath)
    foreach ($k in $Parameters.Keys) {
        $v = [string]$Parameters[$k]
        $argList += "-$k"
        $argList += $v
    }

    $lines += "[DIAG] Executar processo externo:"
    $lines += "       $psExe " + (($argList | ForEach-Object { if ($_ -match '\s') { "`"$_`"" } else { $_ } }) -join ' ')
    $lines += "[DIAG] stdout -> $tmpStdout"

    $startT = [datetime]::UtcNow
    $exitCode = -1
    try {
        # Start-Process com redirect para ficheiro. Sem deadlock de pipes,
        # sem dependencia da event queue do PS.
        $proc = Start-Process -FilePath $psExe -ArgumentList $argList `
            -NoNewWindow -PassThru `
            -RedirectStandardOutput $tmpStdout `
            -RedirectStandardError  $tmpStderr `
            -ErrorAction Stop
    } catch {
        $lines += "[ERR] Falha a lancar powershell.exe: $($_.Exception.Message)"
        return ,$lines
    }

    try {
        $exited = $proc.WaitForExit($TimeoutSeconds * 1000)
        if (-not $exited) {
            try { $proc.Kill() } catch {}
            $lines += "[ERR] Timeout apos ${TimeoutSeconds}s. Processo terminado."
            return ,$lines
        }
        $exitCode = $proc.ExitCode
    } finally {
        try { $proc.Dispose() } catch {}
    }

    $elapsed = [int]([datetime]::UtcNow - $startT).TotalMilliseconds
    $lines += "[DIAG] Exit code: $exitCode  |  Duracao: ${elapsed}ms"

    # Ler os ficheiros agora que o processo terminou. Sem risco de deadlock.
    $stdout = ''
    $stderr = ''
    try { if (Test-Path $tmpStdout) { $stdout = Get-Content -LiteralPath $tmpStdout -Raw -Encoding UTF8 } } catch {}
    try { if (Test-Path $tmpStderr) { $stderr = Get-Content -LiteralPath $tmpStderr -Raw -Encoding UTF8 } } catch {}

    $stdoutLen = if ($stdout) { $stdout.Length } else { 0 }
    $stderrLen = if ($stderr) { $stderr.Length } else { 0 }
    $lines += "[DIAG] stdout=$stdoutLen chars | stderr=$stderrLen chars"
    $lines += ''

    $bodyCount = 0
    if ($stdout) {
        $stdoutLines = $stdout -split "`r?`n"
        # Remove ultima linha vazia resultante do split final
        if ($stdoutLines.Count -gt 0 -and $stdoutLines[-1] -eq '') {
            $stdoutLines = $stdoutLines[0..($stdoutLines.Count-2)]
        }
        foreach ($line in $stdoutLines) { $lines += $line; $bodyCount++ }
    }
    if ($stderr -and $stderr.Trim()) {
        $lines += ''
        $lines += '[DIAG] stderr:'
        foreach ($line in ($stderr -split "`r?`n")) {
            if ($line.Trim()) { $lines += "  $line" }
        }
    }

    if ($bodyCount -eq 0 -and (-not $stderr -or -not $stderr.Trim())) {
        $lines += '[WARN] Processo externo terminou com exit 0 mas sem qualquer output nem erro.'
        $lines += "       Ver ficheiros: $tmpStdout  e  $tmpStderr"
        $lines += '       Se estao vazios, o script em si nao emitiu nada no ambiente do spawn.'
    }

    # NAO apagamos os ficheiros temporarios - permitem post-mortem
    Write-AppLog "External done: exit=$exitCode  stdout=$stdoutLen  stderr=$stderrLen  tmp=$tmpBase"
    return ,$lines
}

# ========== Execucao inline: corre script na thread principal (UI freeze) ==========
# Usado para scripts curtos (AD lookups ~3s). Nao tem cancelamento, mas e
# MUITO mais fiavel que o runspace - nao depende de DoEvents nem de
# pumps de mensagens WebView2 no meio do BeginInvoke.
function Invoke-ScriptInline {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [hashtable]$Parameters = @{}
    )

    $lines = @()
    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        $msg = "[ERR] Script nao encontrado: $ScriptPath"
        Write-AppLog $msg
        return ,@($msg)
    }

    $pStr = if ($Parameters -and $Parameters.Count -gt 0) {
        ($Parameters.GetEnumerator() | ForEach-Object { "-$($_.Key) '$($_.Value)'" }) -join ' '
    } else { '(nenhum)' }
    Write-AppLog "Inline start: $ScriptPath  params=$pStr"

    $prevEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        # `*>&1` redirige todos os streams (output+error+warning+verbose+info)
        # para o success stream, que capturamos em $raw. Mesmo comportamento
        # que correr `& .\script.ps1 -userName t05352 *>&1` numa consola.
        $raw = & $ScriptPath @Parameters *>&1
        foreach ($item in $raw) {
            if ($item -is [System.Management.Automation.ErrorRecord]) {
                foreach ($el in (Format-PSErrorRecord $item)) { $lines += $el }
            } elseif ($item -is [System.Management.Automation.WarningRecord]) {
                $lines += "[WARN] $($item.Message)"
            } elseif ($item -is [System.Management.Automation.InformationRecord]) {
                $lines += "$item"
            } else {
                $lines += "$item"
            }
        }
    } catch {
        foreach ($el in (Format-PSErrorRecord $_)) { $lines += $el }
    } finally {
        $ErrorActionPreference = $prevEAP
    }

    if ($lines.Count -eq 0) {
        $lines += '[WARN] O script terminou sem output e sem erros.'
        $lines += "       Script: $ScriptPath"
        $lines += "       Parametros: $pStr"
        $lines += '       Causas provaveis:'
        $lines += '       - O valor pesquisado nao existe no AD'
        $lines += '       - Sem ligacao ao dominio (tenta "nltest /dsgetdc:" numa consola PS)'
    }

    Write-AppLog "Inline done: lines=$($lines.Count)"
    return ,$lines
}

# ========== Dispatcher: recebe (toolId, params) e devolve @{ lines = [...] } ==========
function Invoke-ToolRequest {
    param([string]$ToolId, [hashtable]$Params)

    Write-AppLog "RUN_TOOL: $ToolId  params=$($Params | ConvertTo-Json -Compress -Depth 5)"

    switch ($ToolId) {
        'SelfTest' {
            # Teste isolado da ligacao WebView2 <-> PowerShell.
            # Nao toca em AD, nao spawn processos, nao toca em ficheiros.
            # Se isto devolver linhas a UI, o bridge funciona. Se nao devolver,
            # o problema NAO esta nos scripts AD.
            $bits = if ([IntPtr]::Size -eq 8) { '64-bit' } else { '32-bit' }
            $lines = @(
                "[OK] Bridge PS <-> JS: OK",
                "[INFO] Se estas a ver estas linhas, a reply do PowerShell chega a UI.",
                "",
                "[DIAG] Host PowerShell $($PSVersionTable.PSVersion) $bits",
                "[DIAG] PID: $PID",
                "[DIAG] User: $env:USERDOMAIN\$env:USERNAME",
                "[DIAG] Machine: $env:COMPUTERNAME",
                "[DIAG] ApartmentState: $([System.Threading.Thread]::CurrentThread.GetApartmentState())",
                "[DIAG] App log: $script:LogPath",
                "[DIAG] Hora servidor: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))",
                "",
                "[INFO] Se UserInfo/GroupInfo nao devolvem nada mas este teste devolve,",
                "       o problema esta no spawn de powershell.exe ou na leitura",
                "       dos ficheiros stdout. Partilha o app log para diagnostico."
            )
            return @{ lines = $lines }
        }

        'UserInfo' {
            $username = if ($Params.username) { [string]$Params.username } else { '' }
            $email    = if ($Params.email) { [string]$Params.email } else { '' }
            if (-not $username -and -not $email) { throw 'Indique username ou email.' }
            $scriptPath = Join-Path $ToolsDir 'scripts\get-diag-info-aw2.ps1'
            $p = @{}
            if ($username) { $p['userName'] = $username }
            if ($email)    { $p['email']    = $email }
            # Spawn powershell.exe 64-bit em processo separado (isolado da app)
            $lines = Invoke-ScriptExternal -ScriptPath $scriptPath -Parameters $p -TimeoutSeconds 90
            return @{ lines = $lines }
        }

        'GroupInfo' {
            $g = [string]$Params.groupName
            if (-not $g) { throw 'Indique o nome do grupo.' }
            $scriptPath = Join-Path $ToolsDir 'scripts\get-diag-info-group.ps1'
            $lines = Invoke-ScriptExternal -ScriptPath $scriptPath -Parameters @{ groupName = $g } -TimeoutSeconds 60
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

# Limpar cache do WebView2 se a versao mudou.
# WebView2 guarda JS/CSS/HTML agressivamente; sem isto, user pode estar a
# correr assets cacheados de uma versao antiga enquanto os ficheiros do
# disco ja sao novos.
try {
    $versionFile = Join-Path $ScriptDir 'version.json'
    $curVer = '?'
    if (Test-Path $versionFile) {
        $vObj = Get-Content $versionFile -Raw | ConvertFrom-Json
        $curVer = [string]$vObj.version
    }
    $lastVerMarker = Join-Path $userDataFolder '.last-version'
    $lastVer = if (Test-Path $lastVerMarker) { (Get-Content $lastVerMarker -Raw).Trim() } else { '' }
    if ($lastVer -ne $curVer) {
        Write-AppLog "Versao mudou ($lastVer -> $curVer). A limpar cache WebView2."
        $cacheDirs = @(
            (Join-Path $userDataFolder 'EBWebView\Default\Cache'),
            (Join-Path $userDataFolder 'EBWebView\Default\Code Cache'),
            (Join-Path $userDataFolder 'EBWebView\Default\Service Worker')
        )
        foreach ($d in $cacheDirs) {
            if (Test-Path $d) {
                try { Remove-Item $d -Recurse -Force -ErrorAction SilentlyContinue } catch {}
            }
        }
        Set-Content -Path $lastVerMarker -Value $curVer -Encoding UTF8 -Force
    }
} catch {
    Write-AppLog "Cache clear falhou: $($_.Exception.Message)"
}

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

    # Injectar APP_VERSION antes de qualquer <script type=text/babel> correr
    try {
        $verPath = Join-Path $ScriptDir 'version.json'
        $appVer = '?'
        if (Test-Path $verPath) {
            $verObj = Get-Content $verPath -Raw | ConvertFrom-Json
            if ($verObj.version) { $appVer = [string]$verObj.version }
        }
        $injectJs = "window.APP_VERSION = '$appVer';"
        [void]$sender.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync($injectJs)
        Write-AppLog "APP_VERSION injectado: $appVer"
    } catch {
        Write-AppLog "Injeccao APP_VERSION falhou: $($_.Exception.Message)"
    }

    # DevTools no arranque (DEBUG MODE - v1.0.12: esta a investigar-se porque
    # que a comunicacao PS<->JS nao funciona no PC do utilizador)
    try {
        $sender.CoreWebView2.OpenDevToolsWindow()
        Write-AppLog "DevTools aberto (debug mode)"
    } catch {
        Write-AppLog "OpenDevToolsWindow falhou: $($_.Exception.Message)"
    }

    # Garantir que web messages estao habilitados (default e $true, mas em
    # ambientes corporativos algumas politicas podem ter desligado)
    try {
        $wmEnabledBefore = $sender.CoreWebView2.Settings.IsWebMessageEnabled
        $sender.CoreWebView2.Settings.IsWebMessageEnabled = $true
        Write-AppLog "IsWebMessageEnabled before=$wmEnabledBefore, forced=$($sender.CoreWebView2.Settings.IsWebMessageEnabled)"
    } catch {
        Write-AppLog "IsWebMessageEnabled falhou: $($_.Exception.Message)"
    }

    # Log de NavigationCompleted para confirmar que o documento carregou
    try {
        $sender.CoreWebView2.add_NavigationCompleted({
            param($navSrc, $navArgs)
            try {
                Write-AppLog "NavigationCompleted: success=$($navArgs.IsSuccess) httpStatus=$($navArgs.HttpStatusCode) webErr=$($navArgs.WebErrorStatus)"
                # Push PS->JS test message para verificar o outbound channel
                $navSrc.ExecuteScriptAsync("console.log('[PS->JS] NavigationCompleted seen by PS');") | Out-Null
            } catch { Write-AppLog "NavCompleted handler erro: $($_.Exception.Message)" }
        })
        Write-AppLog "NavigationCompleted handler registado"
    } catch {
        Write-AppLog "add_NavigationCompleted falhou: $($_.Exception.Message)"
    }

    # Handler de mensagens do JS - usar delegate tipado explicito em vez de
    # scriptblock implicito (conversao implicita pode falhar silenciosamente
    # em alguns builds PS 5.1)
    Write-AppLog "A registar WebMessageReceived handler (delegate explicito)..."
    try {
        $sender.CoreWebView2.add_WebMessageReceived(
            [EventHandler[Microsoft.Web.WebView2.Core.CoreWebView2WebMessageReceivedEventArgs]]{
                param($src, $evArgs)
                # LOG MUITO CEDO — garante que sabemos que o handler disparou
                try { Write-AppLog "WebMsg handler: FIRED" } catch {}
        # WebView2 entrega postMessage(string) via TryGetWebMessageAsString()
        # e postMessage(object) via WebMessageAsJson. Tentamos ambos para
        # nao depender de como o JS enviou.
        $raw = $null
        try { $raw = $evArgs.TryGetWebMessageAsString() } catch {}
        if ([string]::IsNullOrEmpty($raw)) {
            try {
                $rawJson = $evArgs.WebMessageAsJson
                if (-not [string]::IsNullOrEmpty($rawJson)) {
                    Write-AppLog "WebMsg: fallback para WebMessageAsJson"
                    # WebMessageAsJson devolve a propria JSON string
                    $raw = $rawJson
                }
            } catch {
                Write-AppLog "WebMsg: WebMessageAsJson THREW: $($_.Exception.Message)"
            }
        }
        if ([string]::IsNullOrEmpty($raw)) {
            Write-AppLog "WebMsg: nem string nem JSON - a ignorar"
            return
        }
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
            Write-AppLog "Dispatcher ERRO: $($_.Exception.Message)`n  at: $($_.ScriptStackTrace)"
            $reply = @{ id = $msg.id; ok = $false; error = "Dispatcher: $($_.Exception.Message)" }
        }

        # Safety net: garantir que SEMPRE enviamos uma resposta, mesmo que tudo falhe
        if (-not $reply) {
            Write-AppLog "REPLY nulo apos dispatch - criando resposta de erro default"
            $reply = @{ id = $msg.id; ok = $false; error = 'Dispatcher terminou sem produzir resposta' }
        }

        try {
            $json = $reply | ConvertTo-Json -Depth 10 -Compress
            $shortPreview = if ($json.Length -gt 300) { $json.Substring(0, 300) + '...' } else { $json }
            Write-AppLog "REPLY ($($json.Length) chars) id=$($msg.id): $shortPreview"
            $src.PostWebMessageAsJson($json)
            Write-AppLog "REPLY posted OK (id=$($msg.id))"
        } catch {
            Write-AppLog "REPLY FALHOU (id=$($msg.id)): $($_.Exception.Message)"
            # Tentar enviar erro simples sem serializar $reply inteiro
            try {
                $fallbackJson = "{""id"":$($msg.id),""ok"":false,""error"":""PostWebMessageAsJson falhou: $($_.Exception.Message -replace '\"','\\\"')""}"
                $src.PostWebMessageAsJson($fallbackJson)
                Write-AppLog "REPLY fallback posted"
            } catch {
                Write-AppLog "REPLY fallback TAMBEM falhou: $($_.Exception.Message)"
            }
        }
        })
        Write-AppLog "WebMessageReceived handler REGISTADO com sucesso"
    } catch {
        Write-AppLog "FALHOU a registar WebMessageReceived: $($_.Exception.Message)"
        Write-AppLog "  Tipo: $($_.Exception.GetType().FullName)"
    }

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
