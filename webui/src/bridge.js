// ============================================================
// bridge.js — Ponte JS <-> PowerShell via WebView2
// ============================================================
// Quando a app corre dentro do WebView2, window.chrome.webview existe.
// Este bridge envia mensagens para o host PS e resolve promessas com
// o resultado. Se nao houver host, desactiva-se (simulacao continua a
// funcionar normalmente).
// ============================================================
(function () {
  const BRIDGE_VERSION = 'v1.0.16';
  console.info('[bridge] loading ' + BRIDGE_VERSION);

  const hasHost = !!(window.chrome && window.chrome.webview);
  console.info('[bridge] hasHost =', hasHost);
  console.info('[bridge] window.chrome =', !!window.chrome);
  console.info('[bridge] window.chrome.webview =', !!(window.chrome && window.chrome.webview));

  if (!hasHost) {
    console.warn('[bridge] Sem host WebView2 - a usar simulacoes.');
    window.bridge = { available: false };
    return;
  }

  const pending = new Map();
  let seq = 0;

  window.chrome.webview.addEventListener('message', (evt) => {
    console.info('[bridge] RAW message received', evt.data);
    let msg = evt.data;
    if (typeof msg === 'string') {
      try { msg = JSON.parse(msg); } catch (e) {
        console.warn('[bridge] JSON parse failed for string msg', e);
      }
    }
    if (!msg) { console.warn('[bridge] msg is null/undefined'); return; }
    if (msg.id == null) { console.warn('[bridge] msg has no id', msg); return; }

    const p = pending.get(msg.id);
    if (!p) {
      console.warn('[bridge] no pending request for id=' + msg.id, 'pending ids:', Array.from(pending.keys()));
      return;
    }

    if (msg.type === 'TOOL_LINE') {
      if (typeof p.onLine === 'function') {
        try { p.onLine(msg.line); } catch (e) { console.error('[bridge] onLine handler threw:', e); }
      }
      return;
    }

    console.info('[bridge] final reply for id=' + msg.id, 'ok=' + msg.ok);
    pending.delete(msg.id);
    if (msg.ok) p.resolve(msg.result);
    else p.reject(new Error(msg.error || 'erro desconhecido'));
  });

  function send(type, payload, opts = {}) {
    const timeoutMs = opts.timeoutMs ?? 180000;
    const onLine    = opts.onLine ?? null;
    return new Promise((resolve, reject) => {
      const id = ++seq;
      console.info('[bridge] SEND id=' + id + ' type=' + type, payload);
      const timer = timeoutMs > 0 ? setTimeout(() => {
        if (pending.has(id)) {
          pending.delete(id);
          console.warn('[bridge] timeout id=' + id + ' type=' + type);
          reject(new Error('timeout (' + Math.round(timeoutMs/1000) + 's) a executar ' + type));
        }
      }, timeoutMs) : null;
      pending.set(id, {
        onLine,
        resolve: (v) => { if (timer) clearTimeout(timer); resolve(v); },
        reject:  (e) => { if (timer) clearTimeout(timer); reject(e); }
      });
      try {
        // IMPORTANTE: enviar como STRING (JSON.stringify), nao object.
        // WebView2 entrega objects via WebMessageAsJson (propriedade) e
        // strings via TryGetWebMessageAsString (metodo). O nosso handler
        // PS usa TryGetWebMessageAsString, portanto a mensagem tem de ser
        // string. Enviar object retornava null no PS e o handler saia
        // silenciosamente sem log.
        const payloadStr = JSON.stringify({ id, type, ...payload });
        window.chrome.webview.postMessage(payloadStr);
        console.info('[bridge] postMessage OK id=' + id + ' (' + payloadStr.length + ' chars)');
      } catch (e) {
        console.error('[bridge] postMessage THREW for id=' + id, e);
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

  console.info('[bridge] Ligado ao host PowerShell. window.bridge =', window.bridge);

  // Auto-ping ao arranque - se a comunicacao funciona, isto aparece
  // como SEND + final reply no console.
  setTimeout(() => {
    console.info('[bridge] AUTO-PING a executar getContext() no arranque...');
    window.bridge.getContext()
      .then(ctx => console.info('[bridge] AUTO-PING OK:', ctx))
      .catch(err => console.error('[bridge] AUTO-PING FAILED:', err.message));
  }, 1500);
})();