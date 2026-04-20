// Splash — loading screen with spinning logo + module checks
const { useState, useEffect } = React;

const Splash = ({ lang, onDone }) => {
  const s = window.STRINGS[lang];
  const checks = lang === 'pt' ? [
    'A validar credenciais Kerberos',
    'A ligar a DC01.nbdomain.local',
    'A autenticar no tenant Entra ID',
    'A carregar módulos ON PREM',
    'A carregar módulos Microsoft 365',
    'A verificar permissões RBAC',
  ] : [
    'Validating Kerberos credentials',
    'Connecting to DC01.nbdomain.local',
    'Authenticating against Entra ID tenant',
    'Loading ON PREM modules',
    'Loading Microsoft 365 modules',
    'Verifying RBAC permissions',
  ];
  const [step, setStep] = useState(0);
  const [done, setDone] = useState(false);
  useEffect(() => {
    if (step < checks.length) {
      const t = setTimeout(() => setStep(step + 1), 380);
      return () => clearTimeout(t);
    }
    const t = setTimeout(() => setDone(true), 450);
    const t2 = setTimeout(onDone, 1000);
    return () => { clearTimeout(t); clearTimeout(t2); };
  }, [step]);

  return (
    <div className={"splash" + (done ? " done" : "")}>
      <div className="splash-inner">
        <div className="splash-logo-wrap">
          <div className="splash-ring" />
          <div className="splash-ring-2" />
          <img src="assets/nb-mark.svg" className="splash-logo" alt="novobanco" />
        </div>
        <div style={{ textAlign: 'center' }}>
          <div className="splash-brand">novobanco · IT Operations</div>
          <div className="splash-title" style={{ marginTop: 6 }}>
            {s.booting}<span className="dots"></span>
          </div>
        </div>
        <div className="splash-progress" />
        <div className="splash-checks">
          {checks.map((c, i) => (
            <div key={i} className={"splash-check" + (i < step ? " show ok" : i === step ? " show" : "")}>
              <span className="mark">
                {i < step ? (
                  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"><path d="M5 12l5 5L20 7"/></svg>
                ) : i === step ? (
                  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3"><circle cx="12" cy="12" r="8" strokeDasharray="10 30" style={{ animation: 'ringSpin 0.8s linear infinite', transformOrigin: 'center' }}/></svg>
                ) : null}
              </span>
              <span>{c}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};

window.Splash = Splash;
