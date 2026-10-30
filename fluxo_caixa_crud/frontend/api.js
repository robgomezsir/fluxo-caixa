const pending = new Map();

function uid() {
  return 'r' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

function initApi() {
  if (window.chrome && window.chrome.webview) {
    window.chrome.webview.addEventListener('message', e => {
      let msg;
      try { msg = JSON.parse(e.data); } catch { return; }
      const reqId = msg.id ?? msg.Id;
      const resolver = pending.get(reqId);
      if (!resolver) return;
      pending.delete(reqId);
      const ok = msg.success ?? msg.Success;
      const data = msg.data ?? msg.Data;
      const error = msg.error ?? msg.Error;
      ok ? resolver.resolve(data) : resolver.reject(new Error(error || 'Erro'));
    });
  }
}

function api(action, entity, payload = {}) {
  return new Promise((resolve, reject) => {
    const id = uid();
    pending.set(id, { resolve, reject });
    const req = JSON.stringify({ id, action, entity: entity || '', payload });
    if (window.chrome && window.chrome.webview) {
      window.chrome.webview.postMessage(req);
    } else {
      pending.delete(id);
      reject(new Error('WebView2 não disponível. Abra pelo Excel.'));
    }
    setTimeout(() => {
      if (pending.has(id)) {
        pending.delete(id);
        reject(new Error('Timeout na comunicação com Excel.'));
      }
    }, 30000);
  });
}

window.FCApi = {
  init: initApi,
  ping: () => api('ping', ''),
  meta: () => api('meta', ''),
  resumo: () => api('resumo', ''),
  list: (entity, payload) => api('list', entity, payload),
  get: (entity, payload) => api('get', entity, payload),
  create: (entity, payload) => api('create', entity, payload),
  update: (entity, payload) => api('update', entity, payload),
  upsert: (entity, payload) => api('upsert', entity, payload),
  remove: (entity, payload) => api('delete', entity, payload),
  rubricas: sheet => api('list', 'rubricas', { sheet }),
  download: () => api('download', ''),
  saveWorkbook: () => api('save', ''),
  showExcelBase: () => api('show_excel', ''),
  configList: (entity, payload) => api('list', entity, payload || {}),
  configAdd: (entity, payload) => api('create', entity, payload),
  configRemove: (entity, payload) => api('delete', entity, payload)
};

initApi();
