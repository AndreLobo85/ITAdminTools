// ============================================================
// bridge.js — Ponte JS ↔ PowerShell via WebView2
// ============================================================
// Quando a app corre dentro do WebView2, window.chrome.webview existe.
// Este bridge envia mensagens para o host PS e resolve promessas com
// o resultado. Se nao houver host, desactiva-se (simulacao continua a
// funcionar normalmente).
// ============================================================
(function () {
  const hasHost = !!(window.chrome && window.chrome.webview);
  if (!hasHost) {
    console.info('[bridge] Sem host WebView2 — a usar simulacoes.');
    window.bridge = { available: false };
    return;
  }

  const pending = new Map();
  let seq = 0;

  window.chrome.webview.addEventListener('message', (evt) => {
    // evt.data pode vir como objecto (PostWebMessageAsJson) ou string
    let msg = evt.data;
    if (typeof msg === 'string') { try { msg = JSON.parse(msg); } catch {} }
    if (!msg || msg.id == null) return;
    const p = pending.get(msg.id);
    if (!p) return;
    pending.delete(msg.id);
    if (msg.ok) p.resolve(msg.result);
    else p.reject(new Error(msg.error || 'erro desconhecido'));
  });

  function send(type, payload, opts = {}) {
    const timeoutMs = opts.timeoutMs ?? 180000; // 3 min default
    return new Promise((resolve, reject) => {
      const id = ++seq;
      const timer = timeoutMs > 0 ? setTimeout(() => {
        if (pending.has(id)) {
          pending.delete(id);
          reject(new Error('timeout (' + Math.round(timeoutMs/1000) + 's) a executar ' + type));
        }
      }, timeoutMs) : null;
      pending.set(id, {
        resolve: (v) => { if (timer) clearTimeout(timer); resolve(v); },
        reject:  (e) => { if (timer) clearTimeout(timer); reject(e); }
      });
      window.chrome.webview.postMessage({ id, type, ...payload });
    });
  }

  window.bridge = {
    available: true,
    // Sem timeout do lado JS: o utilizador pode parar via cancelTool().
    // O host PowerShell ja pode ser cancelado manualmente atraves do botao Stop.
    runTool: (toolId, params) => send('RUN_TOOL', { toolId, params }, { timeoutMs: 0 }),
    cancelTool: () => send('CANCEL_TOOL', {}, { timeoutMs: 5000 }),
    getContext: () => send('GET_CONTEXT', {}, { timeoutMs: 5000 })
  };

  console.info('[bridge] Ligado ao host PowerShell.');
})();
