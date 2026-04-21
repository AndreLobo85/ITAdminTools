// App — shell, splash, titlebar, routing, tweaks
const { useState: useAppState, useEffect: useAppEffect } = React;

const TWEAK_DEFAULTS = /*EDITMODE-BEGIN*/{
  "theme": "corporate",
  "lang": "pt"
}/*EDITMODE-END*/;

const App = () => {
  const [booting, setBooting] = useAppState(true);
  const [view, setView] = useAppState(() => {
    try { return JSON.parse(localStorage.getItem('nbit_view')) || { screen: 'dashboard' }; }
    catch { return { screen: 'dashboard' }; }
  });
  const [theme, setTheme] = useAppState(TWEAK_DEFAULTS.theme);
  const [lang, setLang] = useAppState(TWEAK_DEFAULTS.lang);
  const [tweaksOpen, setTweaksOpen] = useAppState(false);
  const [editMode, setEditMode] = useAppState(false);

  useAppEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
  }, [theme]);

  useAppEffect(() => {
    localStorage.setItem('nbit_view', JSON.stringify(view));
  }, [view]);

  // Edit mode integration
  useAppEffect(() => {
    const handler = (e) => {
      if (e.data?.type === '__activate_edit_mode') { setEditMode(true); setTweaksOpen(true); }
      if (e.data?.type === '__deactivate_edit_mode') { setEditMode(false); setTweaksOpen(false); }
    };
    window.addEventListener('message', handler);
    window.parent.postMessage({ type: '__edit_mode_available' }, '*');
    return () => window.removeEventListener('message', handler);
  }, []);

  const persist = (patch) => {
    window.parent.postMessage({ type: '__edit_mode_set_keys', edits: patch }, '*');
  };

  const s = window.STRINGS[lang];

  const goDashboard = () => setView({ screen: 'dashboard' });
  const enterArea = (area) => setView({ screen: 'tool', area });
  const quickRun = (area, scriptId) => setView({ screen: 'tool', area, scriptId });

  return (
    <>
      {booting && <Splash lang={lang} onDone={() => setBooting(false)}/>}
      <div className="app-bg"/>
      <div className="app">
        <div className="titlebar">
          <div className="tb-left">
            <div className="tb-brand" onClick={goDashboard} style={{ cursor: 'pointer' }}>
              <div className="tb-brand-mark">n</div>
              <span className="tb-brand-name">novobanco</span>
              <span className="tb-brand-suffix">IT·ADMIN</span>
            </div>
            <span className="tb-sep"/>
            <div className="tb-crumb">
              {view.screen === 'dashboard' ? (
                <><Icon name="grid" size={13}/> <strong>{lang==='pt'?'Consola':'Console'}</strong></>
              ) : (
                <>
                  <span style={{cursor:'pointer'}} onClick={goDashboard}>{lang==='pt'?'Consola':'Console'}</span>
                  <Icon name="chevron" size={11}/>
                  <strong>{view.area === 'onprem' ? s.onprem : s.m365}</strong>
                </>
              )}
            </div>
          </div>
          <div className="tb-right">
            <div className="tb-status"><span className="dot"/>{s.connected}</div>
            <div className="tb-env">{s.prod.toUpperCase()}</div>
            <button className="btn btn--ghost" style={{padding:'6px 10px',fontSize:12}}
                    onClick={() => setTweaksOpen(o => !o)}>
              <Icon name="settings" size={14}/> {s.tweaks}
            </button>
            <div className="tb-user">
              <div className="avatar">JF</div>
              <span>{s.user}</span>
            </div>
          </div>
        </div>

        {view.screen === 'dashboard' && (
          <Dashboard lang={lang}
                     onEnterArea={enterArea}
                     onQuickRun={quickRun}/>
        )}
        {view.screen === 'tool' && (
          <ToolView key={view.area + (view.scriptId||'')}
                    area={view.area} lang={lang}
                    initialScriptId={view.scriptId}
                    onBack={goDashboard}/>
        )}

        <div className="app-version-badge" title={lang==='pt'?'Versão da aplicação':'App version'}>
          v{window.APP_VERSION || '?'}
        </div>
      </div>

      <div className={"tweaks" + (tweaksOpen ? " open" : "")}>
        <h4>
          <span>{s.tweaks}</span>
          <button className="tweak-close" onClick={() => setTweaksOpen(false)}><Icon name="x" size={12}/></button>
        </h4>
        <div className="tweak-group">
          <label>{s.theme}</label>
          <div className="seg">
            {['corporate','terminal','light'].map(t => (
              <button key={t} className={theme===t?'active':''}
                      onClick={() => { setTheme(t); persist({ theme: t }); }}>
                {t === 'corporate' ? (lang==='pt'?'Corporate':'Corporate')
                 : t === 'terminal' ? 'Terminal'
                 : (lang==='pt'?'Light':'Light')}
              </button>
            ))}
          </div>
        </div>
        <div className="tweak-group">
          <label>{s.language}</label>
          <div className="seg seg-2">
            <button className={lang==='pt'?'active':''} onClick={() => { setLang('pt'); persist({ lang: 'pt' }); }}>Português</button>
            <button className={lang==='en'?'active':''} onClick={() => { setLang('en'); persist({ lang: 'en' }); }}>English</button>
          </div>
        </div>
        <div style={{
          marginTop: 10, padding: '10px 12px',
          background: 'var(--bg-2)', borderRadius: 6,
          fontSize: 11, color: 'var(--text-dim)',
          fontFamily: 'var(--mono)', lineHeight: 1.5,
          border: '1px solid var(--border)'
        }}>
          {lang === 'pt'
            ? 'Dica: Cmd/Ctrl+K abre pesquisa rápida (em breve)'
            : 'Tip: Cmd/Ctrl+K opens quick search (soon)'}
        </div>
      </div>
    </>
  );
};

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);
