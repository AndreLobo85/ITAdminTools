<#
.SYNOPSIS
    REPL de longa duracao para Exchange Online.

.DESCRIPTION
    Processo pwsh.exe mantido vivo pela app apos o utilizador clicar "Ligar a
    Exchange Online". Faz Connect-ExchangeOnline uma vez, depois fica em loop
    a ler comandos do stdin e a executar na sessao EXO activa.

    Protocolo stdin/stdout:
      - Cada linha de stdin e um JSON:
          { id: <int>, type: 'run', script: '<ps code>' }
          { id: <int>, type: 'exit' }

      - Resposta e sempre uma sequencia de linhas em stdout, terminadas por:
          <<<DONE id=N ok=true|false>>>

      - Durante arranque, emite <<<READY upn=... tenant=...>>> ou
        <<<READY_ERROR msg=...>>>

    Os markers <<<...>>> sao o mecanismo de sincronizacao entre o app e este
    REPL.
#>

param(
    [switch]$useDeviceCode
)

$ErrorActionPreference = 'Continue'

function Emit-Line { param([string]$Line); [Console]::Out.WriteLine($Line); [Console]::Out.Flush() }

Emit-Line "[DIAG] EXO REPL a iniciar (PS $($PSVersionTable.PSVersion))..."

# ---- Load module ----
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Emit-Line "[ERR] Modulo ExchangeOnlineManagement nao instalado."
    Emit-Line "      Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    Emit-Line "<<<READY_ERROR msg=module-not-installed>>>"
    exit 1
}
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Emit-Line "[DIAG] ExchangeOnlineManagement carregado."
} catch {
    Emit-Line "[ERR] Falha a importar modulo: $($_.Exception.Message)"
    Emit-Line "<<<READY_ERROR msg=import-failed>>>"
    exit 1
}

# ---- Connect ----
$exoParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
if ($useDeviceCode) {
    $exoParams['Device'] = $true
    Emit-Line "[INFO] Device code activo - URL + codigo vai aparecer."
}

Emit-Line "[INFO] Connect-ExchangeOnline (popup Microsoft vai abrir)..."
try {
    Connect-ExchangeOnline @exoParams
} catch {
    $errMsg = $_.Exception.Message
    Emit-Line "[ERR] Connect-ExchangeOnline falhou: $errMsg"
    if (-not $useDeviceCode -and ($errMsg -match 'canceled|cancelled|timeout|not found|could not|closed')) {
        Emit-Line "[INFO] Auto-fallback para device code..."
        try {
            $exoParams['Device'] = $true
            Connect-ExchangeOnline @exoParams
        } catch {
            Emit-Line "[ERR] Device code tambem falhou: $($_.Exception.Message)"
            Emit-Line "<<<READY_ERROR msg=connect-failed>>>"
            exit 1
        }
    } else {
        Emit-Line "<<<READY_ERROR msg=connect-failed>>>"
        exit 1
    }
}

# Identify connected session
$conn = $null
try { $conn = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected' | Select-Object -First 1 } catch {}
$upn = if ($conn) { $conn.UserPrincipalName } else { '?' }
$tenant = if ($conn) { $conn.TenantId } else { '?' }
Emit-Line "[OK] Ligado como: $upn"
Emit-Line "<<<READY upn=$upn tenant=$tenant>>>"

# ---- REPL loop ----
while ($true) {
    $line = [Console]::In.ReadLine()
    if ($null -eq $line) { break }  # stdin fechado pelo pai
    if (-not $line.Trim()) { continue }

    $req = $null
    try { $req = $line | ConvertFrom-Json } catch {
        Emit-Line "[ERR] JSON invalido: $line"
        Emit-Line "<<<DONE id=0 ok=false>>>"
        continue
    }

    $reqId = if ($null -ne $req.id) { $req.id } else { 0 }

    if ($req.type -eq 'exit') {
        Emit-Line "[INFO] Exit solicitado. A desconectar..."
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
        Emit-Line "<<<DONE id=$reqId ok=true>>>"
        break
    }

    if ($req.type -eq 'ping') {
        Emit-Line "[OK] pong"
        Emit-Line "<<<DONE id=$reqId ok=true>>>"
        continue
    }

    if ($req.type -eq 'run') {
        $script = [string]$req.script
        if (-not $script) {
            Emit-Line "[ERR] 'run' sem campo 'script'"
            Emit-Line "<<<DONE id=$reqId ok=false>>>"
            continue
        }

        $ok = $true
        try {
            $sb = [scriptblock]::Create($script)
            $raw = & $sb *>&1
            foreach ($item in $raw) {
                if ($item -is [System.Management.Automation.ErrorRecord]) {
                    Emit-Line "[ERR] $($item.Exception.Message)"
                } elseif ($item -is [System.Management.Automation.WarningRecord]) {
                    Emit-Line "[WARN] $($item.Message)"
                } else {
                    Emit-Line "$item"
                }
            }
        } catch {
            $ok = $false
            Emit-Line "[ERR] $($_.Exception.Message)"
        }
        Emit-Line "<<<DONE id=$reqId ok=$(if ($ok) {'true'} else {'false'})>>>"
        continue
    }

    Emit-Line "[ERR] Tipo desconhecido: $($req.type)"
    Emit-Line "<<<DONE id=$reqId ok=false>>>"
}

Emit-Line "[INFO] REPL a terminar."
