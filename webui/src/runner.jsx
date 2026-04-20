// Runner — runs a script, shows terminal output
const { useState, useEffect, useRef } = React;

const TerminalView = ({ lines, running, onClear, script, lang, params }) => {
  const s = window.STRINGS[lang];
  const bodyRef = useRef(null);
  const [copied, setCopied] = useState(false);
  const [visibleLines, setVisibleLines] = useState([]);
  const [runStart, setRunStart] = useState(null);
  const [elapsed, setElapsed] = useState(0);

  useEffect(() => {
    if (!lines || lines.length === 0) { setVisibleLines([]); return; }
    setVisibleLines([]);
    setRunStart(Date.now());
    let i = 0;
    const interval = setInterval(() => {
      i++;
      setVisibleLines(lines.slice(0, i));
      if (i >= lines.length) clearInterval(interval);
    }, 180);
    return () => clearInterval(interval);
  }, [lines]);

  useEffect(() => {
    if (!running) return;
    const t = setInterval(() => setElapsed((Date.now() - runStart) / 1000), 80);
    return () => clearInterval(t);
  }, [running, runStart]);

  useEffect(() => {
    if (bodyRef.current) bodyRef.current.scrollTop = bodyRef.current.scrollHeight;
  }, [visibleLines]);

  const classifyLine = (line) => {
    if (!line) return '';
    if (line.startsWith('PS>')) return 'l-cmd';
    if (line.startsWith('[OK]')) return 'l-ok';
    if (line.startsWith('[WARN]')) return 'l-warn';
    if (line.startsWith('[ERR]')) return 'l-err';
    if (line.startsWith('[INFO]') || line.startsWith('[AUDIT]') || line.startsWith('[...]')) return 'l-info';
    if (line.match(/^\s{2}[A-Z]/)) return 'l-kv';
    return '';
  };

  const copyAll = () => {
    navigator.clipboard?.writeText(visibleLines.join('\n'));
    setCopied(true);
    setTimeout(() => setCopied(false), 1400);
  };

  return (
    <div className="terminal">
      <div className="term-head">
        <div className="term-head-left">
          <div className="term-dots">
            <div className="term-dot d1"/><div className="term-dot d2"/><div className="term-dot d3"/>
          </div>
          <span className="term-title">
            {script ? `PowerShell 7.4 — ${script.file || script.id+'.ps1'}` : 'PowerShell 7.4'}
          </span>
        </div>
        <div className="term-head-right">
          <button className="term-icon-btn" onClick={copyAll} title={copied ? s.copied : s.copy}>
            <Icon name={copied ? 'check' : 'copy'} size={14}/>
          </button>
          <button className="term-icon-btn" onClick={onClear} title={s.clear}>
            <Icon name="x" size={14}/>
          </button>
        </div>
      </div>
      <div className="term-body" ref={bodyRef}>
        {visibleLines.length === 0 && !running ? (
          <div className="term-empty">
            <div className="icon"><Icon name="terminal" size={22}/></div>
            <div>
              {lang === 'pt'
                ? 'Preenche os parâmetros e clica em Executar.'
                : 'Fill parameters and click Run.'}
            </div>
          </div>
        ) : (
          <>
            {visibleLines.map((line, i) => (
              <div key={i} className={classifyLine(line)}>{line || '\u00A0'}</div>
            ))}
            {running && <div className="l-cmd">PS&gt; <span className="term-caret"/></div>}
          </>
        )}
      </div>
      <div className="term-foot">
        <span>
          {running ? (
            <span className="term-foot-pill run">● {s.running.toUpperCase()}</span>
          ) : visibleLines.length > 0 ? (
            <span className="term-foot-pill ok">● {s.success.toUpperCase()} · exit 0</span>
          ) : (
            <span className="term-foot-pill idle">● IDLE</span>
          )}
        </span>
        <span>
          {visibleLines.length > 0 && (
            <>{s.duration}: {elapsed.toFixed(2)}s · {visibleLines.length} lines</>
          )}
        </span>
      </div>
    </div>
  );
};

const Runner = ({ script, area, lang, onBack }) => {
  const s = window.STRINGS[lang];
  const [formValues, setFormValues] = useState(() => {
    const init = {};
    script.params?.forEach(p => {
      init[p.id] = p.default !== undefined ? p.default : '';
    });
    return init;
  });
  const [running, setRunning] = useState(false);
  const [stopping, setStopping] = useState(false);
  const [outputLines, setOutputLines] = useState([]);
  const [confirmingDanger, setConfirmingDanger] = useState(false);

  // Reset when script changes
  useEffect(() => {
    const init = {};
    script.params?.forEach(p => {
      init[p.id] = p.default !== undefined ? p.default : '';
    });
    setFormValues(init);
    setOutputLines([]);
    setConfirmingDanger(false);
    setStopping(false);
  }, [script.id]);

  const canRun = !running && (() => {
    // required fields
    for (const p of script.params) {
      if (p.required) {
        const v = formValues[p.id];
        if (!v && v !== 0) return false;
        if (typeof v === 'string' && !v.trim()) return false;
      }
    }
    // requireOneOf — at least one of these must be filled
    if (script.requireOneOf && script.requireOneOf.length) {
      const anyFilled = script.requireOneOf.some(id => {
        const v = formValues[id];
        return v && String(v).trim();
      });
      if (!anyFilled) return false;
    }
    return true;
  })();

  const doRun = async () => {
    setRunning(true);
    setStopping(false);
    setOutputLines([]);
    setConfirmingDanger(false);
    try {
      let out;
      if (window.bridge && window.bridge.available) {
        // Execucao real via host PowerShell (WebView2 bridge)
        const result = await window.bridge.runTool(script.id, formValues);
        out = (result && result.lines) ? result.lines : (Array.isArray(result) ? result : [String(result)]);
      } else {
        // Fallback: simulacao embutida
        await new Promise(r => setTimeout(r, 400));
        out = script.output(formValues);
      }
      setOutputLines(out);
      setTimeout(() => { setRunning(false); setStopping(false); }, out.length * 180 + 200);
    } catch (err) {
      setOutputLines(['[ERR] ' + (err && err.message ? err.message : err)]);
      setRunning(false);
      setStopping(false);
    }
  };

  const stop = async () => {
    if (!running || stopping) return;
    setStopping(true);
    if (window.bridge && window.bridge.available && window.bridge.cancelTool) {
      try {
        await window.bridge.cancelTool();
      } catch (err) {
        // se o cancelamento falhar, mostra mas mantem o estado
        setOutputLines(prev => [...prev, '[WARN] Cancelamento falhou: ' + (err.message || err)]);
        setStopping(false);
      }
      // o doRun() ainda esta a aguardar pela resposta do runspace; quando o
      // runspace parar, o resultado (com '[WARN] Execucao cancelada') chega.
    } else {
      // sem bridge: simulacao apenas para de "renderizar"
      setRunning(false);
      setStopping(false);
    }
  };

  const run = () => {
    if (script.danger && !confirmingDanger) {
      setConfirmingDanger(true);
      return;
    }
    doRun();
  };

  return (
    <div className="runner fade-in" key={script.id}>
      <div className="runner-head">
        <div className="runner-title">
          <div className="runner-icon"><Icon name={script.icon} size={22}/></div>
          <div>
            <h1 className="runner-name">{script.name[lang]}</h1>
            <div className="runner-desc">{script.desc[lang]}</div>
            <div className="runner-path">scripts/{area}/{script.file || script.id+'.ps1'}</div>
          </div>
        </div>
      </div>

      <div className="runner-body">
        <div>
          <div className="form-section">
            <h3>{s.params}</h3>
            {script.danger && (
              <div className="danger-banner">
                <Icon name="shield" size={18}/>
                <div>
                  <strong>{s.danger}</strong>
                  {s.dangerDesc}
                </div>
              </div>
            )}
            {script.requireOneOf && (
              <div style={{
                padding: '8px 12px', marginBottom: 12,
                background: 'var(--bg-2)', border: '1px dashed var(--border)',
                borderRadius: 6, fontSize: 11.5, color: 'var(--text-dim)',
                fontFamily: 'var(--mono)'
              }}>
                {lang === 'pt'
                  ? `Preenche pelo menos um: ${script.requireOneOf.join(' ou ')}`
                  : `Fill at least one of: ${script.requireOneOf.join(' or ')}`}
              </div>
            )}
            {script.params.map(p => (
              <div className="field" key={p.id}>
                {p.type === 'check' ? (
                  <label className="check-row">
                    <input type="checkbox"
                           checked={!!formValues[p.id]}
                           onChange={e => setFormValues({...formValues, [p.id]: e.target.checked})}/>
                    {p.label[lang]}
                  </label>
                ) : p.type === 'select' ? (
                  <>
                    <div className="field-label">
                      <span>{p.label[lang]}</span>
                      {p.required && <span className="req">* {s.required}</span>}
                    </div>
                    <select value={formValues[p.id] ?? p.default ?? ''}
                            onChange={e => setFormValues({...formValues, [p.id]: e.target.value})}>
                      {p.options.map(o => (
                        <option key={o.v} value={o.v}>{o.l[lang]}</option>
                      ))}
                    </select>
                  </>
                ) : p.type === 'textarea' ? (
                  <>
                    <div className="field-label">
                      <span>{p.label[lang]}</span>
                      {p.required && <span className="req">* {s.required}</span>}
                    </div>
                    <textarea rows={4}
                              value={formValues[p.id] || ''}
                              onChange={e => setFormValues({...formValues, [p.id]: e.target.value})}
                              placeholder={p.placeholder}
                              spellCheck={false}/>
                  </>
                ) : p.type === 'number' ? (
                  <>
                    <div className="field-label">
                      <span>{p.label[lang]}</span>
                      {p.required && <span className="req">* {s.required}</span>}
                    </div>
                    <input type="number"
                           value={formValues[p.id] ?? p.default ?? 0}
                           min={p.min} max={p.max}
                           onChange={e => setFormValues({...formValues, [p.id]: e.target.value})}
                           style={{ maxWidth: 140 }}/>
                  </>
                ) : (
                  <>
                    <div className="field-label">
                      <span>{p.label[lang]}</span>
                      {p.required && <span className="req">* {s.required}</span>}
                    </div>
                    <input type="text"
                           value={formValues[p.id] || ''}
                           onChange={e => setFormValues({...formValues, [p.id]: e.target.value})}
                           placeholder={p.placeholder}
                           spellCheck={false}/>
                  </>
                )}
              </div>
            ))}

            <div className="run-actions">
              {confirmingDanger ? (
                <>
                  <button className="btn btn--danger" onClick={doRun}>
                    <Icon name="power" size={16}/> {s.confirm}
                  </button>
                  <button className="btn btn--ghost" onClick={() => setConfirmingDanger(false)}>
                    {s.cancel}
                  </button>
                </>
              ) : running ? (
                <button className="btn btn--danger" onClick={stop} disabled={stopping}>
                  <Icon name="stop" size={14}/>
                  {stopping ? s.stopping + '...' : s.stop}
                </button>
              ) : (
                <button className="btn btn--primary" onClick={run} disabled={!canRun}>
                  <Icon name="play" size={14}/>
                  {s.run}
                </button>
              )}
              <button className="btn btn--ghost" onClick={() => setOutputLines([])} disabled={outputLines.length === 0 || running}>
                <Icon name="x" size={14}/> {s.clear}
              </button>
            </div>
          </div>
        </div>

        <div>
          <div className="form-section">
            <h3>{s.output}</h3>
          </div>
          <TerminalView lines={outputLines} running={running}
                        onClear={() => setOutputLines([])}
                        script={script} lang={lang} params={formValues}/>
        </div>
      </div>
    </div>
  );
};

window.Runner = Runner;
