// ToolView — sidebar list + runner
const { useState: useToolState, useMemo, useEffect: useToolEffect } = React;

// Determina que servico M365 um modulo requer (retorna null se nenhum).
// Serve para: (a) desactivar Run quando nao ha sessao, (b) mostrar barra
// de ligacao apropriada no topo.
function moduleRequiresService(mod) {
  if (!mod) return null;
  if (mod.cat === 'Exchange Online') return 'exo';
  if (mod.cat === 'SharePoint Online') return 'spo';
  return null;
}

const M365ConnectionBar = ({ lang, service, status, onConnect, onDisconnect, connecting, disconnecting, pendingLines }) => {
  const label = service === 'exo' ? 'Exchange Online' : 'SharePoint Online';
  const svcStatus = status ? status[service] : null;
  const connected = !!(svcStatus && svcStatus.connected);
  const upn = svcStatus && svcStatus.info ? svcStatus.info.upn : null;

  return (
    <div className="m365-conn-bar">
      <div className="m365-conn-svc">
        <div className={"m365-dot " + (connected ? 'ok' : 'off')}/>
        <div>
          <div className="m365-conn-title">{label}</div>
          <div className="m365-conn-sub">
            {connected
              ? (lang === 'pt' ? `Ligado como ${upn}` : `Connected as ${upn}`)
              : (connecting
                  ? (lang === 'pt' ? 'A ligar...' : 'Connecting...')
                  : (lang === 'pt' ? 'Não ligado' : 'Not connected'))}
          </div>
        </div>
      </div>
      <div className="m365-conn-actions">
        {connected ? (
          <button className="btn btn--ghost" onClick={onDisconnect} disabled={disconnecting}>
            <Icon name="power" size={14}/>
            {disconnecting ? (lang==='pt'?'A desligar...':'Disconnecting...') : (lang==='pt'?'Desligar':'Disconnect')}
          </button>
        ) : (
          <button className="btn btn--primary" onClick={onConnect} disabled={connecting}>
            <Icon name="power" size={14}/>
            {connecting ? (lang==='pt'?'A ligar...':'Connecting...') : (lang==='pt'?'Ligar':'Connect')}
          </button>
        )}
      </div>
      {(connecting || (pendingLines && pendingLines.length > 0)) && (
        <div className="m365-conn-log">
          {pendingLines && pendingLines.map((l, i) => (<div key={i}>{l}</div>))}
        </div>
      )}
    </div>
  );
};

const ToolView = ({ area, lang, initialScriptId, onBack }) => {
  const s = window.STRINGS[lang];
  const modules = area === 'onprem' ? window.ONPREM_MODULES : window.M365_MODULES;
  const areaLabel = area === 'onprem' ? s.onprem : s.m365;

  const [selectedId, setSelectedId] = useToolState(initialScriptId || modules[0].id);
  const [query, setQuery] = useToolState('');

  const selected = modules.find(m => m.id === selectedId) || modules[0];

  // M365 connection state (polling + manual refresh)
  const [m365Status, setM365Status] = useToolState(null);
  const [connecting, setConnecting] = useToolState(null);  // 'exo' | 'spo' | null
  const [disconnecting, setDisconnecting] = useToolState(null);
  const [connectLog, setConnectLog] = useToolState([]);

  const requiredService = area === 'm365' ? moduleRequiresService(selected) : null;

  const refreshM365Status = async () => {
    if (!(window.bridge && window.bridge.available && window.bridge.m365Status)) return;
    try {
      const st = await window.bridge.m365Status();
      setM365Status(st);
    } catch {}
  };

  useToolEffect(() => {
    if (area !== 'm365') return;
    refreshM365Status();
    const t = setInterval(refreshM365Status, 15000);
    return () => clearInterval(t);
  }, [area]);

  const connectSvc = async (svc) => {
    if (!window.bridge || !window.bridge.available) return;
    setConnecting(svc);
    setConnectLog([lang === 'pt' ? '[INFO] A iniciar pwsh.exe e Connect...' : '[INFO] Starting pwsh.exe and Connect...']);
    try {
      const r = await window.bridge.m365Connect(svc);
      if (r && r.lines) setConnectLog(r.lines);
      await refreshM365Status();
    } catch (e) {
      setConnectLog(prev => [...prev, '[ERR] ' + (e.message || e)]);
    } finally {
      setConnecting(null);
    }
  };

  const disconnectSvc = async (svc) => {
    if (!window.bridge || !window.bridge.available) return;
    setDisconnecting(svc);
    try {
      await window.bridge.m365Disconnect(svc);
      setConnectLog([]);
      await refreshM365Status();
    } catch (e) {
      console.error('disconnect failed', e);
    } finally {
      setDisconnecting(null);
    }
  };

  const svcConnected = requiredService && m365Status && m365Status[requiredService] && m365Status[requiredService].connected;

  const grouped = useMemo(() => {
    const q = query.trim().toLowerCase();
    const filtered = modules.filter(m => {
      if (!q) return true;
      return (m.name[lang] + ' ' + m.cat + ' ' + m.desc[lang]).toLowerCase().includes(q);
    });
    const byCat = {};
    filtered.forEach(m => {
      byCat[m.cat] = byCat[m.cat] || [];
      byCat[m.cat].push(m);
    });
    return byCat;
  }, [query, lang, modules]);

  return (
    <div className="tool-view fade-in">
      <aside className="tool-sidebar" style={{ '--accent-area': area === 'm365' ? '#5B8DEF' : 'var(--nb)' }}>
        <button className="back-btn" onClick={onBack}>
          <Icon name="chevron" size={14}/> {s.back}
        </button>
        <div className="area-badge"><span/>{areaLabel}</div>
        <h2 className="sidebar-title">
          {area === 'onprem'
            ? (lang === 'pt' ? 'Ferramentas on-premises' : 'On-premises tools')
            : (lang === 'pt' ? 'Ferramentas Microsoft 365' : 'Microsoft 365 tools')}
        </h2>
        <div className="sidebar-sub">
          {modules.length} {s.scripts} · {Object.keys(grouped).length} {s.categories}
        </div>

        <div className="search-box">
          <Icon name="search" size={14}/>
          <input type="text" placeholder={s.search}
                 value={query} onChange={e => setQuery(e.target.value)}/>
        </div>

        <div className="sidebar-list">
          {Object.keys(grouped).length === 0 && (
            <div style={{ padding: 20, textAlign: 'center', color: 'var(--text-faint)', fontSize: 12 }}>
              {s.noResults}
            </div>
          )}
          {Object.entries(grouped).map(([cat, items]) => (
            <div className="cat-group" key={cat}>
              <div className="cat-label">{cat}</div>
              {items.map(m => (
                <button key={m.id}
                        className={"script-item" + (m.id === selectedId ? " active" : "")}
                        onClick={() => setSelectedId(m.id)}>
                  <div className="script-icon"><Icon name={m.icon} size={14}/></div>
                  <div className="script-body">
                    <div className="script-name">{m.name[lang]}</div>
                    <div className="script-cat">{m.file || m.id+'.ps1'}</div>
                  </div>
                </button>
              ))}
            </div>
          ))}
        </div>
      </aside>

      <div className="tool-view-main">
        {requiredService && (
          <M365ConnectionBar
            lang={lang}
            service={requiredService}
            status={m365Status}
            onConnect={() => connectSvc(requiredService)}
            onDisconnect={() => disconnectSvc(requiredService)}
            connecting={connecting === requiredService}
            disconnecting={disconnecting === requiredService}
            pendingLines={connecting === requiredService ? connectLog : []}
          />
        )}
        <Runner
          script={selected}
          area={area}
          lang={lang}
          onBack={onBack}
          requiresService={requiredService}
          serviceConnected={!!svcConnected}
        />
      </div>
    </div>
  );
};

window.ToolView = ToolView;
