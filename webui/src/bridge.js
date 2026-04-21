// ============================================================
// bridge.js — Ponte JS <-> PowerShell via WebView2
// ============================================================
// Quando a app corre dentro do WebView2, window.chrome.webview existe.
// Este bridge envia mensagens para o host PS e resolve promessas com
// o resultado. Se nao houver host, desactiva-se (simulacao continua a
// funcionar normalmente).
// ============================================================
(function () {
  const BRIDGE_VERSION = 'v1.0.19';
  console.info('[bridge] loading ' + BRIDGE_VERSION);

  const hasHost = !!(window.chrome && window.chrome.webview);
  console.info('[bridge] hasHost =', hasHost);

  if (!hasHost) {
    console.warn('[bridge] Sem host WebView2 - a usar simulacoes.');
    window.bridge = { available: false };
    return;
  }

  const pending = new Map();
  let seq = 0;

  window.chrome.webview.addEventListener('message', (evt) => {
    let msg = evt.data;
    if (typeof msg === 'string') {
      try { msg = JSON.parse(msg); } catch (e) { console.warn('[bridge] JSON parse failed', e); }
    }
    if (!msg || msg.id == null) return;

    const p = pending.get(msg.id);
    if (!p) return;

    if (msg.type === 'TOOL_LINE') {
      if (typeof p.onLine === 'function') {
        try { p.onLine(msg.line); } catch (e) { console.error('[bridge] onLine:', e); }
      }
      return;
    }

    pending.delete(msg.id);
    if (msg.ok) p.resolve(msg.result);
    else p.reject(new Error(msg.error || 'erro desconhecido'));
  });

  function send(type, payload, opts = {}) {
    const timeoutMs = opts.timeoutMs ?? 180000;
    const onLine    = opts.onLine ?? null;
    return new Promise((resolve, reject) => {
      const id = ++seq;
      const timer = timeoutMs > 0 ? setTimeout(() => {
        if (pending.has(id)) {
          pending.delete(id);
          reject(new Error('timeout (' + Math.round(timeoutMs/1000) + 's) a executar ' + type));
        }
      }, timeoutMs) : null;
      pending.set(id, {
        onLine,
        resolve: (v) => { if (timer) clearTimeout(timer); resolve(v); },
        reject:  (e) => { if (timer) clearTimeout(timer); reject(e); }
      });
      try {
        // STRING (JSON.stringify), nao object. WebView2 entrega objects via
        // WebMessageAsJson e strings via TryGetWebMessageAsString - o PS
        // handler usa o metodo string.
        const payloadStr = JSON.stringify({ id, type, ...payload });
        window.chrome.webview.postMessage(payloadStr);
      } catch (e) {
        console.error('[bridge] postMessage THREW', e);
        pending.delete(id);
        if (timer) clearTimeout(timer);
        reject(e);
      }
    });
  }

  window.bridge = {
    available: true,
    version: BRIDGE_VERSION,
    runTool: (toolId, params, onLine) =>
      send('RUN_TOOL', { toolId, params }, { timeoutMs: 0, onLine }),
    cancelTool: () => send('CANCEL_TOOL', {}, { timeoutMs: 5000 }),
    getContext: () => send('GET_CONTEXT', {}, { timeoutMs: 5000 })
  };

  console.info('[bridge] Ligado ao host PowerShell ' + BRIDGE_VERSION);
})();