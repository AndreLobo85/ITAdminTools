# IT Admin Toolkit

Aplicação gráfica unificada para administração IT. Organiza ferramentas por **categoria** (AD, Fileshare, Exchange, SharePoint, M365), cada uma num separador próprio. Extensível: adicionar uma ferramenta nova é um ficheiro em `tools/` + uma linha no registo.

## Como correr

- Duplo-clique em `Launch-NoConsole.vbs` (recomendado — sem consola)
- Ou `ITAdminToolkit.bat`
- Ou clique direito em `ITAdminToolkit.ps1` → **Run with PowerShell**

## On-Premise vs Cloud

Cada categoria está marcada com `[On-Prem]` ou `[Cloud]` no separador.

| Tipo | Onde correr | Exemplos |
|------|-------------|----------|
| **[On-Prem]** | Tipicamente num **servidor da rede** (DC, file server, Exchange server) ou máquina com acesso directo aos recursos e módulos RSAT/ExchangeMgmtShell instalados | AD, Fileshare, Exchange On-Premise |
| **[Cloud]** | Pode correr na **máquina local do admin**. Requer módulos PowerShell de cloud (`ExchangeOnlineManagement`, `PnP.PowerShell`, `Microsoft.Graph`) | Exchange Online, SharePoint Online, Entra ID |

A app não falha se um módulo não estiver disponível — cada ferramenta detecta as suas dependências e mostra um aviso amigável se não puder correr no ambiente actual.

## Ferramentas incluídas

### [On-Prem] Active Directory
- **Group Auditor** — procura grupos AD por sufixo (ex: `NF`, `NR`), expande recursivamente sub-grupos, lista users finais. Exporta Excel com 3 folhas.

### [On-Prem] Fileshare
- **Share Auditor** — permissões SMB + NTFS de um share, opcionalmente recursivo, expande grupos AD nas ACEs. Exporta Excel com 3 folhas.

### Categorias preparadas (vazias)
- `[On-Prem] Exchange On-Premise`
- `[Cloud] Exchange Online`
- `[Cloud] SharePoint Online`
- `[Cloud] Microsoft 365 / Entra ID`

## Estrutura de ficheiros

```
ITAdminToolkit/
├── ITAdminToolkit.ps1          # App principal (registo de ferramentas + UI)
├── ITAdminToolkit.bat          # Launcher com consola oculta
├── Launch-NoConsole.vbs        # Launcher 100% silencioso
├── README.md
└── tools/
    ├── _common.ps1             # Helpers partilhados (AD, Excel, cache, UI)
    ├── ShareAuditor.ps1        # Ferramenta: New-ShareAuditorTab
    └── ADGroupAuditor.ps1      # Ferramenta: New-ADGroupAuditorTab
```

## Adicionar uma nova ferramenta (exemplo: "Mailbox Permissions" para Exchange Online)

**1. Criar `tools/MailboxPermissions.ps1`:**

```powershell
function MBX_Invoke-Audit { ... }
function MBX_Export-ToExcel { ... }

function New-MailboxPermissionsTab {
    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = 'Mailbox Permissions'
    # ... construir UI, wire handlers ...
    return $tab
}
```

**2. Em `ITAdminToolkit.ps1`**, dot-source o ficheiro junto aos outros:

```powershell
. (Join-Path $ToolsDir 'MailboxPermissions.ps1')
```

**3. Registar na categoria** no array `$Categories`:

```powershell
@{
    Name   = 'Exchange Online'
    OnPrem = $false
    Tools  = @(
        @{ Name = 'Mailbox Permissions'; Factory = { New-MailboxPermissionsTab } }
    )
}
```

A próxima execução da app vai mostrar o novo separador automaticamente.

## Helpers partilhados (`tools/_common.ps1`)

Cada ferramenta pode usar sem importar nada extra:

| Função | Para quê |
|--------|----------|
| `$script:ADAvailable` | Flag booleana — módulo AD disponível? |
| `$script:UserDetailsCache` | Cache AD partilhada entre ferramentas (evita queries repetidas) |
| `Get-ADUserDetailsBatch` | Batch query de detalhes de users (100 DNs por chamada LDAP) |
| `Test-ExcelAvailable` | Detecta se Excel está instalado |
| `Export-ResultsToCsv` | Export CSV genérico |
| `Convert-RgbToBgr` | Conversão de cor para Excel COM |
| `New-StyledButton`, `New-StyledDataGridView` | Controlos WinForms com estilo consistente |
| `$script:PaletteBlue`, `$script:PaletteGreen` | Cores standard |

## Requisitos

| | |
|-|-|
| Windows + PowerShell 5.1+ | Runtime |
| Módulo **ActiveDirectory** (RSAT) | Opcional — necessário para ferramentas de AD/Fileshare |
| **Microsoft Excel** | Opcional — para exportar `.xlsx`. Sem Excel → fallback CSV |

Instalar RSAT (admin, uma vez):
```powershell
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

Para ferramentas cloud (quando forem adicionadas), instalar conforme necessário:
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module PnP.PowerShell -Scope CurrentUser
Install-Module Microsoft.Graph -Scope CurrentUser
```
