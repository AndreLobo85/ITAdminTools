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

  function send(type, payload) {
    return new Promise((resolve, reject) => {
      const id = ++seq;
      pending.set(id, { resolve, reject });
      window.chrome.webview.postMessage({ id, type, ...payload });
      // timeout de seguranca: 60s
      setTimeout(() => {
        if (pending.has(id)) {
          pending.delete(id);
          reject(new Error('timeout (60s) a executar ' + type));
        }
      }, 60000);
    });
  }

  window.bridge = {
    available: true,
    runTool: (toolId, params) => send('RUN_TOOL', { toolId, params }),
    getContext: () => send('GET_CONTEXT', {})
  };

  console.info('[bridge] Ligado ao host PowerShell.');
})();
