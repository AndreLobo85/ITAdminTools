// Script catalog — mapped to the REAL scripts shared by the user
// ON PREM: UserInfo, GroupInfo, ADGroupAuditor, ShareAuditor
// M365:    MailboxStats
// _common.ps1 is a shared library (not listed as a tool)

const ONPREM_MODULES = [
  {
    id: 'UserInfo',
    file: 'UserInfo.ps1',
    cat: 'Active Directory',
    icon: 'users',
    name: { pt: 'Diagnóstico de User AD', en: 'AD User Diagnostics' },
    desc: {
      pt: 'Consulta completa de um utilizador no AD (sem RSAT): grupos, password, lockout por DC, manager, licenças, atributos de extensão.',
      en: 'Full AD user lookup (no RSAT required): groups, password state, per-DC lockout, manager, licensing, extension attributes.',
    },
    params: [
      { id: 'username', label: { pt: 'Username (SamAccountName)', en: 'Username (SamAccountName)' }, placeholder: 'j.silva' },
      { id: 'email', label: { pt: 'OU Email', en: 'OR Email' }, placeholder: 'joao.silva@novobanco.pt' },
    ],
    requireOneOf: ['username', 'email'],
    runtime: 3.2,
    output: (p) => {
      const who = p.username || (p.email ? p.email.split('@')[0].replace('.', '') : 'j.silva');
      const sam = p.username || who;
      const sep = '------------------------------------------------------------';
      return [
        `PS> .\\UserInfo.ps1  ${p.username ? `-Username ${p.username}` : `-Email ${p.email}`}`,
        `[INFO] A consultar AD em todos os DCs do dominio...`,
        sep,
        `Name                 : ${sam.replace('.', ' ')}`,
        `SamAccountName       : ${sam}`,
        `Alias/mailNickName   : ${sam}`,
        `DisplayName          : ${sam.replace('.', ' ').replace(/\b\w/g, c => c.toUpperCase())}`,
        `DistinguishedName    : CN=${sam},OU=Users,OU=NB,DC=nbdomain,DC=local`,
        `Estrutura            : Sistemas de Informação`,
        `Funcao               : Sysadmin N2`,
        `Manager Name         : Higino Antunes`,
        `mail                 : ${p.email || sam + '@novobanco.pt'}`,
        `alternativos         : smtp:${sam}@nbdomain.local, smtp:${sam}@novobanco.onmicrosoft.com`,
        `Database             : MBX-PT-03`,
        sep,
        `Ultimo Logon         : 2026-04-20 08:12`,
        `enabled              : True`,
        `PasswordExpired      : False`,
        `PasswordLastSet      : 2026-02-14 09:41`,
        `Password expira em   : 2026-05-15 09:41`,
        `LockedOut (por DC)   : False, False, False, False`,
        `badPwdCount (por DC) : 0, 0, 1, 0`,
        `DCs consultados      : DC01, DC02, DC03, DC04`,
        `Criado em            : 2019-06-03 14:22`,
        `User expira em       : Nao expira`,
        sep,
        `Grupo Lic Office          : GO365PRO-E3`,
        `UserPrincipalName         : ${sam}@novobanco.pt`,
        sep,
        `Acede ao PAM?        : Nao`,
        `Grupos               : Domain Users; NB-VPN-Access; NB-FileShare-IT-RO; NB-Exchange-Mailbox-Standard; GO365PRO-E3; GO365SSPRNR`,
        sep,
        `[OK] Consulta concluida (4 DCs consultados em 2.8s)`,
      ];
    },
  },
  {
    id: 'GroupInfo',
    file: 'GroupInfo.ps1',
    cat: 'Active Directory',
    icon: 'shield',
    name: { pt: 'Diagnóstico de Grupo AD', en: 'AD Group Diagnostics' },
    desc: {
      pt: 'Lê um grupo AD via ADSI (sem RSAT): scope, categoria, managedBy, memberOf e todos os membros (users com estado Enabled).',
      en: 'Reads an AD group via ADSI (no RSAT): scope, category, managedBy, memberOf and every member (with user Enabled state).',
    },
    params: [
      { id: 'groupName', label: { pt: 'Nome do grupo', en: 'Group name' }, placeholder: 'NB-IT-Sysadmins', required: true },
    ],
    runtime: 2.8,
    output: (p) => {
      const g = p.groupName || 'NB-IT-Sysadmins';
      const sep = '------------------------------------------------------------';
      return [
        `PS> .\\GroupInfo.ps1 -GroupName "${g}"`,
        `[INFO] A consultar LDAP...`,
        sep,
        `Group Name         : ${g}`,
        `Description        : Acesso administrativo a servidores on-prem`,
        `Display Name       : ${g}`,
        `Distinguished Name : CN=${g},OU=Groups,OU=NB,DC=nbdomain,DC=local`,
        sep,
        `Group Type (code)  : -2147483646`,
        `Group Scope        : Global`,
        `Group Category     : Security`,
        `Managed By         : CN=Higino Antunes,OU=Users,OU=NB,DC=nbdomain,DC=local`,
        `When Created       : 2018-09-14 11:03:22`,
        `When Changed       : 2026-04-18 16:41:08`,
        sep,
        `MemberOf           : NB-Admin-Tier1, NB-PAM-Eligible, NB-ConditionalAccess-Admins`,
        sep,
        `Members (7):`,
        `  user - j.ferreira       - True  - Joao Ferreira`,
        `  user - v.rodrigues      - True  - Vitor Rodrigues`,
        `  user - h.antunes        - True  - Higino Antunes`,
        `  user - p.costa          - True  - Pedro Costa`,
        `  user - a.martins        - False - Ana Martins`,
        `  user - svc-backup       - True  - Service Backup`,
        `  foreignSecurityPrincipal - CORP\\admin-sync - - -`,
        sep,
        `[OK] 7 membros lidos em 1.9s | 1 desativado`,
      ];
    },
  },
  {
    id: 'ADGroupAuditor',
    file: 'ADGroupAuditor.ps1',
    cat: 'Active Directory',
    icon: 'search',
    name: { pt: 'Auditor de Grupos AD', en: 'AD Group Auditor' },
    desc: {
      pt: 'Descobre grupos por sufixo ou nome (com wildcards) e expande recursivamente membros e sub-grupos. Exporta Excel com 3 folhas (Resumo / Detalhe / Por Grupo).',
      en: 'Finds groups by suffix or name (with wildcards) and recursively expands members and sub-groups. Exports 3-sheet Excel workbook (Summary / Detail / By Group).',
    },
    params: [
      { id: 'mode', label: { pt: 'Modo de pesquisa', en: 'Search mode' }, type: 'select', options: [
        { v: 'Suffix', l: { pt: 'Por sufixo', en: 'By suffix' } },
        { v: 'Name',   l: { pt: 'Por nome (aceita *)', en: 'By name (wildcards)' } },
      ], default: 'Suffix' },
      { id: 'terms', label: { pt: 'Termos (separados por vírgula)', en: 'Terms (comma-separated)' }, placeholder: 'NF,NR  ou  HR_*', required: true },
      { id: 'activeOnly', label: { pt: 'Apenas users activos (Enabled=True)', en: 'Active users only (Enabled=True)' }, type: 'check', default: false },
    ],
    runtime: 6.4,
    output: (p) => {
      const terms = (p.terms || 'NF,NR').split(',').map(t=>t.trim()).filter(Boolean);
      const mode = p.mode === 'Name' ? 'nome' : 'sufixo';
      const groups = mode === 'sufixo'
        ? ['NB-FileShare-HR-NF','NB-FileShare-IT-NF','NB-App-SAP-NR','NB-Remote-NF','NB-Backup-NR']
        : ['HR_Lisboa','HR_Porto','HR_Admins'];
      const lines = [
        `PS> .\\ADGroupAuditor.ps1 -Mode ${p.mode||'Suffix'} -Terms ${terms.join(',')}${p.activeOnly?' -ActiveOnly':''}`,
        `[INFO] A procurar grupos por ${mode}: ${terms.join(', ')}`,
        `[OK] ${groups.length} grupos-alvo encontrados`,
        ``,
        `  Grupo                       Scope        Category   Membros  Aninh.`,
        `  --------------------------  -----------  ---------  -------  ------`,
      ];
      groups.forEach((g, i) => {
        const mem = 8 + i*3;
        lines.push(`  ${g.padEnd(27)} ${['Global','DomainLocal','Universal'][i%3].padEnd(12)} Security    ${String(mem).padStart(5)}     ${i%2===0?1:0}`);
      });
      lines.push(``);
      lines.push(`[...] A expandir (1/${groups.length}): ${groups[0]}`);
      lines.push(`[...] Batch LDAP de 100 DNs...`);
      lines.push(`[...] A expandir (${groups.length}/${groups.length}): ${groups[groups.length-1]}`);
      lines.push(``);
      const totalLinhas = 127;
      const usersDist = 84;
      lines.push(`[OK] Concluido: ${groups.length} grupos | ${totalLinhas} linhas | ${usersDist} users distintos${p.activeOnly?' (apenas activos)':''}`);
      lines.push(`[INFO] Disponivel para exportar: Excel (Resumo + Detalhe + Por Grupo) ou CSV`);
      return lines;
    },
  },
  {
    id: 'ShareAuditor',
    file: 'ShareAuditor.ps1',
    cat: 'File Shares',
    icon: 'folder',
    name: { pt: 'Auditor de Permissões Share', en: 'Share Permissions Auditor' },
    desc: {
      pt: 'Lê ACLs SMB + NTFS de um share, opcionalmente recursivo. Expande grupos AD até aos users finais. Exporta Excel 3-folhas (Resumo / Detalhe / Por Pasta).',
      en: 'Reads SMB + NTFS ACLs on a share, optionally recursive. Expands AD groups to final users. Exports 3-sheet Excel (Summary / Detail / By Folder).',
    },
    params: [
      { id: 'path', label: { pt: 'Caminho do share (UNC)', en: 'Share path (UNC)' }, placeholder: '\\\\fileserver\\compliance', required: true },
      { id: 'recurse', label: { pt: 'Incluir subpastas', en: 'Include subfolders' }, type: 'check', default: true },
      { id: 'depth', label: { pt: 'Profundidade', en: 'Depth' }, type: 'number', default: 3, min: 0, max: 20 },
      { id: 'onlyExplicit', label: { pt: 'Filtrar ACEs herdadas nas subpastas', en: 'Filter inherited ACEs on subfolders' }, type: 'check', default: false },
    ],
    runtime: 8.2,
    output: (p) => {
      const path = p.path || '\\\\fileserver\\compliance';
      const lines = [
        `PS> .\\ShareAuditor.ps1 -SharePath "${path}" -Recurse:${p.recurse??true} -Depth ${p.depth??3}${p.onlyExplicit?' -OnlyExplicit':''}`,
        `[INFO] A resolver share...`,
        `[OK] Resolved: \\\\FILESRV01 -> D:\\Shares\\Compliance`,
        `[...] A processar (1/12): ${path}`,
        `[...] A ler ACLs NTFS + SMB`,
        `[...] A expandir grupos AD (batch LDAP x100)...`,
        `[...] A processar (12/12): ${path}\\Audits\\2026`,
        ``,
        `  Pasta                              Raiz  Principal                      Herdada`,
        `  ---------------------------------  ----  -----------------------------  -------`,
        `  ${path}                            Sim   NB-Compliance-Admins           Nao`,
        `  ${path}                            Sim   NB-Compliance-Readers          Nao`,
        `  ${path}\\Audits                    -     NB-Internal-Audit              Nao`,
        `  ${path}\\Contracts                 -     NB-Legal-Team                  Nao`,
        `  ${path}\\Contracts\\Signed         -     j.silva                        Parcial`,
        ``,
        `[OK] Concluido: 203 linhas | 12 pastas | 47 users distintos`,
        `[INFO] Exportar: Excel (Resumo + Detalhe + Por Pasta) ou CSV`,
      ];
      return lines;
    },
  },
];

const M365_MODULES = [
  {
    id: 'MailboxStats',
    file: 'MailboxStats.ps1',
    cat: 'Exchange Online',
    icon: 'mail',
    name: { pt: 'Mailbox Statistics', en: 'Mailbox Statistics' },
    desc: {
      pt: 'Consulta Get-MailboxStatistics para um ou mais UPNs via ExchangeOnlineManagement. Opcionalmente inclui RecoverableItems.',
      en: 'Runs Get-MailboxStatistics for one or more UPNs via ExchangeOnlineManagement. Optionally includes RecoverableItems.',
    },
    params: [
      { id: 'adminUpn', label: { pt: 'Admin UPN (ligação EXO)', en: 'Admin UPN (EXO connection)' }, placeholder: 'admin@novobanco.onmicrosoft.com' },
      { id: 'upns', label: { pt: 'UPN(s) a consultar (um por linha ou separados por vírgula)', en: 'UPN(s) to query (one per line or comma-separated)' }, type: 'textarea', placeholder: 'joao.silva@novobanco.pt\nana.martins@novobanco.pt', required: true },
      { id: 'includeRecov', label: { pt: 'Incluir RecoverableItems (mais lento)', en: 'Include RecoverableItems (slower)' }, type: 'check', default: true },
    ],
    runtime: 5.5,
    output: (p) => {
      const raw = p.upns || 'joao.silva@novobanco.pt\nana.martins@novobanco.pt';
      const upns = raw.split(/[\r\n,;]+/).map(s=>s.trim()).filter(Boolean);
      const lines = [
        `PS> Connect-ExchangeOnline${p.adminUpn ? ` -UserPrincipalName ${p.adminUpn}`:''} -ShowBanner:$false`,
        `[OK] Ligado: ${p.adminUpn || 'admin@novobanco.onmicrosoft.com'}`,
      ];
      upns.forEach((u, i) => {
        lines.push(`[...] A consultar (${i+1}/${upns.length}): ${u}`);
      });
      lines.push(``);
      lines.push(`  UPN                            DisplayName       Size       Items   Quota    LastLogon           Recov.`);
      lines.push(`  -----------------------------  ----------------  ---------  ------  -------  ------------------  ------`);
      const samples = [
        { dn:'Joao Silva',    sz:'18.42 GB', it:42817, q:'BelowLimit', ll:'2026-04-20 08:12', rc:'1.2 GB' },
        { dn:'Ana Martins',   sz:'34.08 GB', it:71204, q:'WarningIssued', ll:'2026-04-20 09:47', rc:'2.8 GB' },
        { dn:'Pedro Costa',   sz:'8.91 GB',  it:19405, q:'BelowLimit', ll:'2026-04-19 17:30', rc:'0.4 GB' },
      ];
      upns.slice(0,3).forEach((u, i) => {
        const s = samples[i % samples.length];
        lines.push(`  ${u.padEnd(30)} ${s.dn.padEnd(17)} ${s.sz.padEnd(10)} ${String(s.it).padStart(6)}  ${s.q.padEnd(7)}  ${s.ll}  ${p.includeRecov!==false?s.rc:'-'}`);
      });
      lines.push(``);
      lines.push(`[OK] Concluido: ${upns.length} OK | 0 erro(s) | ${upns.length} total`);
      lines.push(`[INFO] Exportar: Excel ou CSV`);
      return lines;
    },
  },
  {
    id: 'SharePointSite',
    file: 'SharePointSite.ps1',
    cat: 'SharePoint Online',
    icon: 'folder',
    name: { pt: 'Site Members', en: 'Site Members' },
    desc: {
      pt: 'Lista membros, grupos SharePoint e permissões de um site (via PnP.PowerShell, MFA interactive).',
      en: 'Lists members, SharePoint groups and permissions of a site (via PnP.PowerShell, interactive MFA).',
    },
    params: [
      { id: 'siteUrl', label: { pt: 'URL do site', en: 'Site URL' }, placeholder: 'https://tenant.sharepoint.com/sites/mysite', required: true },
      { id: 'tenantAdminUrl', label: { pt: 'Admin URL (opcional)', en: 'Admin URL (optional)' }, placeholder: 'https://tenant-admin.sharepoint.com' },
      { id: 'includeOwners', label: { pt: 'Incluir owners do M365 Group', en: 'Include M365 Group owners' }, type: 'check', default: true },
    ],
    runtime: 6.2,
    output: (p) => {
      const site = p.siteUrl || 'https://tenant.sharepoint.com/sites/mysite';
      return [
        `PS> Connect-PnPOnline -Url ${site} -Interactive`,
        `[OK] Ligado ao site ${site}`,
        `[INFO] Site Title: Compliance Team Site`,
        `[INFO] Template: STS#3`,
        ``,
        `[AUDIT] SharePoint groups:`,
        `  Compliance Members          12 users`,
        `  Compliance Owners            3 users`,
        `  Compliance Visitors         28 users`,
        ``,
        `[AUDIT] M365 Group owners:`,
        `  joao.silva@novobanco.pt`,
        `  ana.martins@novobanco.pt`,
        ``,
        `[OK] Exportar: Excel ou CSV`,
      ];
    },
  },
];

window.ONPREM_MODULES = ONPREM_MODULES;
window.M365_MODULES = M365_MODULES;
