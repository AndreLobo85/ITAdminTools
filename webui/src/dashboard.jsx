// Dashboard — entry screen with ON PREM / M365 areas + quick actions
const Dashboard = ({ lang, onEnterArea, onQuickRun, recent }) => {
  const s = window.STRINGS[lang];
  const onprem = window.ONPREM_MODULES;
  const m365 = window.M365_MODULES;
  const onpremCats = new Set(onprem.map(m => m.cat));
  const m365Cats = new Set(m365.map(m => m.cat));
  const totalScripts = onprem.length + m365.length;

  const quick = [
    { m: onprem.find(x=>x.id==='UserInfo'), area: 'onprem' },
    { m: onprem.find(x=>x.id==='GroupInfo'), area: 'onprem' },
    { m: onprem.find(x=>x.id==='ADGroupAuditor'), area: 'onprem' },
    { m: onprem.find(x=>x.id==='ShareAuditor'), area: 'onprem' },
    { m: m365.find(x=>x.id==='MailboxStats'), area: 'm365' },
  ];

  const activity = lang === 'pt' ? [
    { ok: true, name: 'UserInfo · j.silva', time: 'há 3m' },
    { ok: true, name: 'MailboxStats · a.martins@novobanco.pt', time: 'há 8m' },
    { ok: true, name: 'ADGroupAuditor · sufixo NF,NR (47 users)', time: 'há 14m' },
    { ok: false, name: 'ShareAuditor · \\\\FILESRV01\\compliance (WARN)', time: 'há 27m' },
    { ok: true, name: 'GroupInfo · NB-IT-Sysadmins', time: 'há 41m' },
    { ok: true, name: 'UserInfo · p.costa@novobanco.pt', time: 'há 1h' },
  ] : [
    { ok: true, name: 'UserInfo · j.silva', time: '3m ago' },
    { ok: true, name: 'MailboxStats · a.martins@novobanco.pt', time: '8m ago' },
    { ok: true, name: 'ADGroupAuditor · suffix NF,NR (47 users)', time: '14m ago' },
    { ok: false, name: 'ShareAuditor · \\\\FILESRV01\\compliance (WARN)', time: '27m ago' },
    { ok: true, name: 'GroupInfo · NB-IT-Sysadmins', time: '41m ago' },
    { ok: true, name: 'UserInfo · p.costa@novobanco.pt', time: '1h ago' },
  ];

  return (
    <div className="main fade-in">
      <div className="dash-hero">
        <div>
          <div className="dash-hello">{s.welcome} · {new Date().toLocaleDateString(lang === 'pt' ? 'pt-PT' : 'en-GB', { weekday: 'long', day: 'numeric', month: 'long' })}</div>
          <h1 className="dash-title">
            IT Admin <span className="accent">Toolkit</span>
          </h1>
          <div className="dash-sub">{s.subtitle}</div>
        </div>
        <div className="dash-stats">
          <div className="dash-stat">
            <div className="dash-stat-val">{totalScripts}</div>
            <div className="dash-stat-label">{lang === 'pt' ? 'Scripts disponíveis' : 'Scripts available'}</div>
          </div>
          <div className="dash-stat">
            <div className="dash-stat-val">147</div>
            <div className="dash-stat-label">{lang === 'pt' ? 'Execuções hoje' : 'Runs today'}</div>
          </div>
          <div className="dash-stat">
            <div className="dash-stat-val">99.8<span style={{fontSize:14,color:'var(--text-dim)'}}>%</span></div>
            <div className="dash-stat-label">{lang === 'pt' ? 'Taxa de sucesso' : 'Success rate'}</div>
          </div>
        </div>
      </div>

      <div className="areas">
        <div className="area area--onprem" onClick={() => onEnterArea('onprem')}>
          <div className="area-mini">
            {Array.from({length: 9}).map((_,i)=>(<span key={i}/>))}
          </div>
          <div>
            <div className="area-tag"><span className="pulse"/>{s.onprem}</div>
            <div className="area-name">ON&nbsp;PREM</div>
            <div className="area-desc">{s.onpremDesc}</div>
          </div>
          <div className="area-foot">
            <div className="area-count">
              <span><strong>{onprem.length}</strong> {s.scripts}</span>
              <span><strong>{onpremCats.size}</strong> {s.categories}</span>
            </div>
            <div className="area-arrow"><Icon name="chevron" size={20}/></div>
          </div>
        </div>

        <div className="area area--m365" onClick={() => onEnterArea('m365')}>
          <div className="area-mini">
            {Array.from({length: 9}).map((_,i)=>(<span key={i}/>))}
          </div>
          <div>
            <div className="area-tag"><span className="pulse"/>{s.m365}</div>
            <div className="area-name">365</div>
            <div className="area-desc">{s.m365Desc}</div>
          </div>
          <div className="area-foot">
            <div className="area-count">
              <span><strong>{m365.length}</strong> {s.scripts}</span>
              <span><strong>{m365Cats.size}</strong> {s.categories}</span>
            </div>
            <div className="area-arrow"><Icon name="chevron" size={20}/></div>
          </div>
        </div>
      </div>

      <div className="dash-lower">
        <div className="panel">
          <div className="panel-head">
            <div className="panel-title"><span className="dot"/>{s.quickRun}</div>
            <div className="panel-hint">⌘K</div>
          </div>
          <div className="quick-grid">
            {quick.filter(x=>x.m).map((x,i) => (
              <button key={i} className="quick-item" onClick={() => onQuickRun(x.area, x.m.id)}>
                <div className="q-icon"><Icon name={x.m.icon} size={16}/></div>
                <div>
                  <div className="q-name">{x.m.name[lang]}</div>
                  <div className="q-cat">{x.m.cat}</div>
                </div>
              </button>
            ))}
          </div>
        </div>

        <div className="panel">
          <div className="panel-head">
            <div className="panel-title"><span className="dot"/>{s.activity}</div>
            <div className="panel-hint"><Icon name="history" size={12}/></div>
          </div>
          <div className="activity-list">
            {activity.map((a,i) => (
              <div className="activity-row" key={i}>
                <span className={"a-mark " + (a.ok ? 'ok' : 'warn')}/>
                <span className="a-name">{a.name}</span>
                <span className="a-time">{a.time}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};

window.Dashboard = Dashboard;
