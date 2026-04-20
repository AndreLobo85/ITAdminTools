// ToolView — sidebar list + runner
const { useState: useToolState, useMemo } = React;

const ToolView = ({ area, lang, initialScriptId, onBack }) => {
  const s = window.STRINGS[lang];
  const modules = area === 'onprem' ? window.ONPREM_MODULES : window.M365_MODULES;
  const areaLabel = area === 'onprem' ? s.onprem : s.m365;

  const [selectedId, setSelectedId] = useToolState(initialScriptId || modules[0].id);
  const [query, setQuery] = useToolState('');

  const selected = modules.find(m => m.id === selectedId) || modules[0];

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

      <Runner script={selected} area={area} lang={lang} onBack={onBack}/>
    </div>
  );
};

window.ToolView = ToolView;
