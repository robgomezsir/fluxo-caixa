const $ = s => document.querySelector(s);
const PAGE = 8;
let meta = null;
let unitModules = [];
let currentModule = 'unit:ALS';
let currentUnitTab = 'lancamentos';
let editingId = null;
let pendingDelete = null;
let modalOnConfirm = null;
let toastT;

const CONSOLIDATED = [
  { id: 'capex', label: 'CAPEX', entity: 'capex', type: 'tabular' },
  { id: 'novos', label: 'Novos Negócios', entity: 'novos_neg', type: 'tabular' },
  { id: 'premissas', label: 'Premissas (Notas)', entity: 'premissas_anot', type: 'tabular' },
  { id: 'bonus', label: 'Bônus Caixa', entity: 'bonus', type: 'tabular' }
];

const PREMISSAS = [
  { id: 'prem_cal', label: 'Premissas Calibradas', sheet: 'premissas calibradas', type: 'premissas_matrix' },
  { id: 'prem_neg', label: 'Premissas de Negócio', sheet: 'PREMISSAS DE NEGÓCIO', type: 'premissas_matrix' },
  { id: 'prem_capex', label: 'CAPEX Mensal', sheet: 'CAPEX_0403', type: 'premissas_matrix' }
];

const RESUMO = { id: 'resumo', label: 'Resumo Geral', type: 'resumo' };
const CONFIG = { id: 'config', label: 'Configurações', type: 'config' };

const UNIT_TABS = [
  { id: 'lancamentos', label: 'Lançamentos' },
  { id: 'capex', label: 'CAPEX' },
  { id: 'novos', label: 'Novos Neg.' },
  { id: 'premissas', label: 'Notas' }
];

const TABULAR_SCHEMAS = {
  capex: [
    { key: 'unidade', label: 'Unidade', req: true, type: 'select', optionsKey: 'unidades', lockOnUnit: true },
    { key: 'fluxo', label: 'Fluxo?', type: 'select', optionsKey: 'simNao' },
    { key: 'pctTerral', label: '% Terral', type: 'percent' },
    { key: 'nomeProjeto', label: 'Nome do Projeto', req: true, grow: true },
    { key: 'valor', label: 'Valor', type: 'money', req: true },
    { key: 'origemCapex', label: 'Origem CAPEX', type: 'select', optionsKey: 'origensCapex' },
    { key: 'classificacao', label: 'Classificação', type: 'select', optionsKey: 'classificacoesCapex' },
    { key: 'centroCusto', label: 'Centro de Custo', type: 'select', optionsKey: 'centrosCusto' },
    { key: 'observacao', label: 'Observação', type: 'textarea', grow: true }
  ],
  novos_neg: [
    { key: 'unidade', label: 'Unidade', req: true, type: 'select', optionsKey: 'unidades', lockOnUnit: true },
    { key: 'codigo', label: 'Código', req: true },
    { key: 'descricao', label: 'Descrição', grow: true, req: true },
    { key: 'valor', label: 'Valor', type: 'money', req: true },
    { key: 'payback', label: 'Payback' }
  ],
  premissas_anot: [
    { key: 'unidade', label: 'Unidade', req: true, type: 'select', optionsKey: 'unidades', lockOnUnit: true },
    { key: 'criterio', label: 'Critério', grow: true, req: true },
    { key: 'responsavel', label: 'Responsável', type: 'select', optionsKey: 'responsaveis' },
    { key: 'nota', label: 'Anotação', type: 'textarea', grow: true }
  ],
  bonus: [
    { key: 'undPertencente', label: 'Und. Pertencente', req: true, type: 'select', optionsKey: 'unidades' },
    { key: 'filialPagante', label: 'Filial Pagante' },
    { key: 'centroCusto', label: 'Centro de Custo', type: 'select', optionsKey: 'centrosCusto' },
    { key: 'projeto', label: 'Projeto' },
    { key: 'valorProvisao', label: 'Valor Provisão', type: 'money' },
    { key: 'marco', label: 'Março', type: 'money' },
    { key: 'abril', label: 'Abril', type: 'money' },
    { key: 'maio', label: 'Maio', type: 'money' }
  ]
};

const UNIT_TABULAR = {
  capex: { entity: 'capex', unitField: 'unidade' },
  novos: { entity: 'novos_neg', unitField: 'unidade' },
  premissas: { entity: 'premissas_anot', unitField: 'unidade' }
};

function buildUnitModules(m) {
  return (m.matrix || [])
    .filter(x => {
      const e = (x.entity || '').toLowerCase();
      return e.startsWith('matrix_') && !e.includes('premissas') && !e.includes('capex0403');
    })
    .map(x => ({
      id: 'unit:' + x.sheet.trim(),
      label: (x.label || x.sheet).trim(),
      sheet: x.sheet.trim(),
      code: (x.code || x.sheet).trim(),
      nome: (x.nome || '').trim(),
      type: 'unit'
    }))
    .sort((a, b) => a.label.localeCompare(b.label, 'pt-BR'));
}

function unitLabel(sheet) {
  const m = unitModules.find(u => u.sheet === sheet);
  return m ? m.label : sheet;
}

function isUnitModule(id) {
  return String(id).startsWith('unit:');
}

function getUnitSheet(id) {
  return id ? id.replace(/^unit:/, '') : '';
}

function fieldOptions(f) {
  if (Array.isArray(f.options) && f.options.length) return f.options;
  if (f.optionsKey && meta && Array.isArray(meta[f.optionsKey])) return meta[f.optionsKey];
  return [];
}

function brl(cents) {
  return (Number(cents || 0) / 100).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

/** Valor decimal do Excel (0,42 = 42%) → "42,00%" */
function pct(ratio) {
  const n = Number(ratio);
  if (!Number.isFinite(n)) return '0,00%';
  return n.toLocaleString('pt-BR', { style: 'percent', minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/** Texto digitado (42 ou 42,00 ou 42,00%) → razão decimal 0,42 */
function pctRatioFromInput(str) {
  let s = String(str || '').trim().replace(/%/g, '').replace(/\s/g, '');
  if (!s) return 0;
  if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  if (!Number.isFinite(n)) return 0;
  return n / 100;
}

function formatFieldValue(f, value) {
  if (value == null || value === '') return '—';
  if (f.type === 'money') return brl(value);
  if (f.type === 'percent') return pct(value);
  return esc(value);
}

function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function normMes(s) {
  return String(s || '').toUpperCase().trim()
    .replace(/Ç/g, 'C').replace(/Á/g, 'A').replace(/Ã/g, 'A')
    .replace(/É/g, 'E').replace(/Í/g, 'I').replace(/Ó/g, 'O').replace(/Ú/g, 'U');
}

function mesMatch(a, b) {
  const na = normMes(a);
  const nb = normMes(b);
  if (!na || !nb) return false;
  if (na === nb) return true;
  return na.startsWith(nb.slice(0, 3)) || nb.startsWith(na.slice(0, 3));
}

function findMatrixVal(cache, rub, mes, tipo) {
  const hit = cache.find(x =>
    x.rubrica === rub &&
    mesMatch(x.mes, mes) &&
    String(x.tipo || '').toLowerCase() === String(tipo || '').toLowerCase() &&
    !x.readonly && !x.ReadOnly
  );
  return hit ? Number(hit.valorCents || 0) : 0;
}

function launchCents(row) {
  const c = row.valorCents ?? row.ValorCents;
  if (c != null && c !== '') return Number(c);
  if (row.valor != null) return Math.round(Number(row.valor) * 100);
  return 0;
}

function mesSortIndex(mes, mesesList) {
  const nm = normMes(mes);
  const idx = (mesesList || []).findIndex(m => mesMatch(m, mes));
  return idx >= 0 ? idx : 999;
}

function sortLaunches(list, mesesList) {
  return list.slice().sort((a, b) => {
    const rub = String(a.rubrica || '').localeCompare(String(b.rubrica || ''), 'pt-BR');
    if (rub) return rub;
    const ma = mesSortIndex(a.mes, mesesList);
    const mb = mesSortIndex(b.mes, mesesList);
    if (ma !== mb) return ma - mb;
    const ta = String(a.tipo || '').localeCompare(String(b.tipo || ''), 'pt-BR');
    if (ta) return ta;
    return String(a.coluna || '').localeCompare(String(b.coluna || ''));
  });
}

function launchesWithValue(allRows) {
  return sortLaunches(
    (allRows || []).filter(r => {
      if (r.readonly || r.ReadOnly) return false;
      return Math.abs(launchCents(r)) > 0;
    }),
    meta?.meses
  );
}

function rubricaNome(r) {
  return typeof r === 'string' ? r : (r.nome || r.Nome || '');
}

function rubricaNatureza(r) {
  if (typeof r === 'object' && r) return r.natureza || r.Natureza || classifyNaturezaLocal(rubricaNome(r), r.secao || r.Secao);
  return classifyNaturezaLocal(r);
}

function launchNatureza(r) {
  return r.natureza || r.Natureza || classifyNaturezaLocal(r.rubrica, r.secao || r.Secao);
}

function launchRubrica(r) {
  return String(r.rubrica || r.Rubrica || '').trim();
}

function classifyNaturezaLocal(name, sectionColA) {
  const r = String(name || '').trim();
  if (r.startsWith('(+)')) return 'Receita';
  if (r.startsWith('(-)')) return 'Despesa';
  if (r.startsWith('(=)')) return 'Resultado';
  const s = String(sectionColA || '').trim();
  if (s) {
    if (/receita/i.test(s)) return 'Receita';
    if (/despesa/i.test(s)) return 'Despesa';
  }
  return 'Outros';
}

function naturezaBadge(n) {
  const cls = n === 'Receita' ? 'nat-rec' : n === 'Despesa' ? 'nat-desp' : n === 'Resultado' ? 'nat-res' : 'nat-out';
  return `<span class="nat-badge ${cls}">${esc(n)}</span>`;
}

function tipoBadge(t) {
  const tipo = String(t || '').trim();
  const cls = tipo === 'Previsto' ? 'tipo-prev' : 'tipo-real';
  return `<span class="tipo-badge ${cls}">${esc(tipo)}</span>`;
}

function normalizeTipo(t) {
  const s = String(t || '').trim();
  if (/prev/i.test(s)) return 'Previsto';
  if (/real/i.test(s)) return 'Realizado';
  return s;
}

function launchCol(r) {
  return String(r.coluna || r.Coluna || '').trim().toUpperCase();
}

function tiposDisponiveisParaMes(mes, columnLayout) {
  const tipos = (columnLayout || [])
    .filter(c => mesMatch(c.mes, mes))
    .map(c => normalizeTipo(c.tipo));
  return [...new Set(tipos)];
}

function resolveColumnForMesTipo(mes, tipo, columnLayout) {
  return (columnLayout || []).find(c =>
    mesMatch(c.mes, mes) && normalizeTipo(c.tipo) === normalizeTipo(tipo)
  );
}

function columnForMes(mes, columnLayout) {
  const hits = (columnLayout || []).filter(c => mesMatch(c.mes, mes));
  return hits.length === 1 ? hits[0] : null;
}

function tipoFromColumn(coluna, columnLayout, fallbackTipo) {
  const col = String(coluna || '').toUpperCase();
  const hit = (columnLayout || []).find(c => String(c.coluna || '').toUpperCase() === col);
  return hit ? normalizeTipo(hit.tipo) : normalizeTipo(fallbackTipo);
}

/** Tipo é por coluna/mês — resolve só pelo mês (uma coluna por mês). */
function alignMatrixPayload(payload, columnLayout) {
  if (!columnLayout.length) return true;
  const col = columnForMes(payload.mes, columnLayout);
  if (!col) {
    toast(`Mês ${payload.mes} não encontrado nesta aba.`, true);
    return false;
  }
  payload.coluna = String(col.coluna).toUpperCase();
  payload.mes = String(col.mes || payload.mes).trim();
  return true;
}

function tipoDoMes(mes, columnLayout) {
  const col = columnForMes(mes, columnLayout);
  return col ? normalizeTipo(col.tipo) : '';
}

function tipoUnicoParaMes(mes, columnLayout) {
  const tipos = tiposDisponiveisParaMes(mes, columnLayout);
  return tipos.length === 1 ? tipos[0] : '';
}

/** Ao mudar o mês, carrega o tipo atual do cabeçalho da coluna. */
function suggestTipoForMes(mes, columnLayout) {
  const sel = $('#mTipo');
  if (!sel) return;
  const tipo = tipoDoMes(mes, columnLayout);
  if (tipo) setSelectValue(sel, tipo);
}

function launchKey(r) {
  return `${launchRubrica(r)}\0${r.mes}\0${normalizeTipo(r.tipo || r.Tipo)}`;
}

function launchTableHeadHtml() {
  return `<colgroup>
    <col class="col-nat" style="width:6%">
    <col class="col-rubrica" style="width:38%">
    <col class="col-mes" style="width:9%">
    <col class="col-tipo" style="width:9%">
    <col class="col-valor" style="width:17%">
    <col class="col-acoes" style="width:21%">
  </colgroup><thead><tr>
    <th class="col-nat">Natureza</th>
    <th class="col-rubrica">Rubrica</th>
    <th class="col-mes">Mês</th>
    <th class="col-tipo">Tipo</th>
    <th class="col-valor">Valor</th>
    <th class="col-acoes">Ações</th>
  </tr></thead>`;
}

function syncTableColumnWidths(table) {
  if (!table) return;
  const tbody = table.querySelector('tbody');
  const thead = table.querySelector('thead');
  const tfoot = table.querySelector('tfoot');
  const ths = table.querySelectorAll('thead th');
  if (!ths.length || !tbody) return;

  const scrollBar = tbody.offsetWidth - tbody.clientWidth;
  const pad = scrollBar > 0 ? `${scrollBar}px` : '';
  if (thead) thead.style.paddingRight = pad;
  if (tfoot) tfoot.style.paddingRight = pad;

  const bodyWidth = tbody.clientWidth;
  const theadTr = thead?.querySelector('tr');
  if (theadTr) theadTr.style.width = bodyWidth > 0 ? `${bodyWidth}px` : '100%';

  const cols = table.querySelectorAll('colgroup col');
  const widths = Array.from(ths).map((th, i) => {
    const col = cols[i];
    const w = col?.style?.width || '';
    if (w.endsWith('px')) return w;
    if (w.endsWith('%') && bodyWidth > 0) {
      return `${Math.round(bodyWidth * parseFloat(w) / 100)}px`;
    }
    return `${th.offsetWidth}px`;
  });

  const syncRows = (rows, cells = 'td') => {
    rows.forEach(tr => {
      if (tr.querySelector(`${cells}[colspan]`)) return;
      tr.style.width = bodyWidth > 0 ? `${bodyWidth}px` : '100%';
      tr.querySelectorAll(cells).forEach((cell, i) => {
        if (widths[i]) cell.style.width = widths[i];
      });
    });
  };

  syncRows(table.querySelectorAll('tbody tr'), 'td');
  if (tfoot) syncRows(tfoot.querySelectorAll('tr'), 'td');
  if (theadTr) syncRows([theadTr], 'th');
}

function syncAllTscrollTables(root) {
  const scope = root && root.querySelectorAll ? root : document;
  scope.querySelectorAll('.tscroll table').forEach(syncTableColumnWidths);
}

function scheduleTscrollSync(root) {
  requestAnimationFrame(() => {
    initColumnResize(root);
    syncAllTscrollTables(root);
    requestAnimationFrame(() => syncAllTscrollTables(root));
  });
}

const COL_WIDTH_STORAGE = 'fc_col_widths_v2';

function migrateColWidthStorage() {
  try {
    const all = JSON.parse(localStorage.getItem(COL_WIDTH_STORAGE) || '{}');
    const legacy = JSON.parse(localStorage.getItem('fc_col_widths_v1') || '{}');
    let changed = false;
    Object.keys(legacy).forEach(k => {
      if (!(k in all)) { all[k] = legacy[k]; changed = true; }
    });
    Object.keys(all).forEach(k => {
      if (k.startsWith('launch-') && !k.startsWith('launch-crud-')) {
        delete all[k];
        changed = true;
      }
    });
    if (changed) localStorage.setItem(COL_WIDTH_STORAGE, JSON.stringify(all));
  } catch { /* ignore */ }
}

function loadColWidths(key) {
  if (!key) return null;
  try {
    const all = JSON.parse(localStorage.getItem(COL_WIDTH_STORAGE) || '{}');
    return Array.isArray(all[key]) ? all[key] : null;
  } catch { return null; }
}

function saveColWidths(key, widths) {
  if (!key || !widths?.length) return;
  try {
    const all = JSON.parse(localStorage.getItem(COL_WIDTH_STORAGE) || '{}');
    all[key] = widths.map(w => Math.round(w));
    localStorage.setItem(COL_WIDTH_STORAGE, JSON.stringify(all));
  } catch { /* ignore */ }
}

function colWidthPx(table, col, index) {
  const w = col.style.width;
  if (w.endsWith('px')) return parseFloat(w);
  if (w.endsWith('%')) return table.offsetWidth * parseFloat(w) / 100;
  const th = table.querySelectorAll('thead th')[index];
  return th?.offsetWidth || 80;
}

function applyColWidths(table, widthsPx) {
  const cols = table.querySelectorAll('colgroup col');
  if (!cols.length || !widthsPx?.length || widthsPx.length !== cols.length) return false;
  widthsPx.forEach((w, i) => {
    if (cols[i] && w > 0) cols[i].style.width = `${Math.round(w)}px`;
  });
  return true;
}

function readColWidthsPx(table) {
  const cols = table.querySelectorAll('colgroup col');
  return Array.from(cols).map((c, i) => colWidthPx(table, c, i));
}

function ensureColsPx(table) {
  const cols = table.querySelectorAll('colgroup col');
  if (!cols.length) return [];
  const widths = readColWidthsPx(table);
  cols.forEach((c, i) => { c.style.width = `${Math.round(widths[i])}px`; });
  return widths;
}

function startColResize(table, colIndex, startX) {
  const cols = [...table.querySelectorAll('colgroup col')];
  if (!cols.length || colIndex < 0 || colIndex >= cols.length) return;

  ensureColsPx(table);
  const minW = 52;
  const partnerIndex = colIndex < cols.length - 1 ? colIndex + 1 : colIndex - 1;
  if (partnerIndex < 0) return;

  const startCol = colWidthPx(table, cols[colIndex], colIndex);
  const startPartner = colWidthPx(table, cols[partnerIndex], partnerIndex);
  const pairTotal = startCol + startPartner;
  const invert = colIndex > partnerIndex;

  document.body.classList.add('col-resizing');

  const onMove = (ev) => {
    let dx = ev.clientX - startX;
    if (invert) dx = -dx;
    let newCol = Math.max(minW, Math.min(pairTotal - minW, startCol + dx));
    let newPartner = pairTotal - newCol;
    cols[colIndex].style.width = `${Math.round(newCol)}px`;
    cols[partnerIndex].style.width = `${Math.round(newPartner)}px`;
  };

  const onUp = () => {
    document.body.classList.remove('col-resizing');
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
    saveColWidths(table.dataset.colResize, readColWidthsPx(table));
    syncTableColumnWidths(table);
  };

  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}

function initColumnResize(root) {
  const scope = root && root.querySelectorAll ? root : document;
  scope.querySelectorAll('table[data-col-resize]').forEach(table => {
    const key = table.dataset.colResize;
    const saved = loadColWidths(key);
    if (saved?.length) {
      const applied = applyColWidths(table, saved);
      if (!applied && key) {
        try {
          const all = JSON.parse(localStorage.getItem(COL_WIDTH_STORAGE) || '{}');
          delete all[key];
          localStorage.setItem(COL_WIDTH_STORAGE, JSON.stringify(all));
        } catch { /* ignore */ }
      }
    }

    table.querySelectorAll('thead th').forEach((th, index) => {
      if (th.querySelector('.col-resize-handle')) return;
      th.classList.add('th-resizable');
      const handle = document.createElement('span');
      handle.className = 'col-resize-handle';
      handle.title = 'Arraste para ajustar a largura';
      handle.setAttribute('role', 'separator');
      handle.setAttribute('aria-orientation', 'vertical');
      handle.addEventListener('mousedown', e => {
        e.preventDefault();
        e.stopPropagation();
        startColResize(table, index, e.clientX);
      });
      handle.addEventListener('dblclick', e => {
        e.preventDefault();
        e.stopPropagation();
        if (!key) return;
        try {
          const all = JSON.parse(localStorage.getItem(COL_WIDTH_STORAGE) || '{}');
          delete all[key];
          localStorage.setItem(COL_WIDTH_STORAGE, JSON.stringify(all));
        } catch { /* ignore */ }
        table.querySelectorAll('colgroup col').forEach(c => { c.style.width = ''; });
        scheduleTscrollSync(table.closest('.tscroll') || table.parentElement);
      });
      th.appendChild(handle);
    });
  });
}

function tableColClass(f) {
  if (f.type === 'money') return 'col-money';
  if (f.type === 'percent') return 'col-pct';
  if (f.type === 'number') return 'col-num';
  if (f.grow) return 'col-grow';
  if (f.type === 'select') return 'col-select';
  return 'col-text';
}

function tabularColWidths(cols) {
  const acoesPct = 19;
  const avail = 100 - acoesPct;
  const weights = cols.map(c => {
    if (c.key === 'origemCapex' || c.key === 'classificacao') return 2.7;
    if (c.key === 'fluxo') return 0.6;
    if (c.type === 'money') return 1;
    if (c.type === 'percent') return 0.7;
    if (c.grow) return 2.4;
    if (c.type === 'select') return 1.45;
    return 1;
  });
  const sum = weights.reduce((a, b) => a + b, 0);
  return weights.map(w => `${((w / sum) * avail).toFixed(1)}%`);
}

function tabularTableHeadHtml(cols) {
  const widths = tabularColWidths(cols);
  const colgroup = cols.map((c, i) =>
    `<col class="${tableColClass(c)}" style="width:${widths[i]}">`
  ).join('') + `<col class="col-acoes" style="width:19%">`;
  const headers = cols.map(c =>
    `<th class="${tableColClass(c)}">${esc(c.label)}</th>`
  ).join('');
  return `<colgroup>${colgroup}</colgroup><thead><tr>${headers}<th class="col-acoes">Ações</th></tr></thead>`;
}

function launchRowHtml(r, editingKey) {
  const cents = launchCents(r);
  const nat = launchNatureza(r);
  const rub = launchRubrica(r);
  const mes = String(r.mes || '').trim();
  const tipo = String(r.tipo || r.Tipo || '').trim();
  const key = launchKey(r);
  const editing = editingKey && key === editingKey;
  const rubEnc = encodeURIComponent(rub);
  return `<tr class="row-${nat.toLowerCase()}${editing ? ' row-editing' : ''}">
    <td class="col-nat">${naturezaBadge(nat)}</td>
    <td class="col-rubrica">${esc(rub)}</td>
    <td class="col-mes">${esc(mes)}</td>
    <td class="col-tipo">${tipoBadge(tipo)}</td>
    <td class="col-valor val-${nat.toLowerCase()}">${brl(cents)}</td>
    <td class="col-acoes"><div class="acoes">
      <button type="button" class="act-btn" data-launch-edit data-rub="${rubEnc}" data-mes="${esc(mes)}" data-tipo="${esc(tipo)}" data-col="${esc(launchCol(r))}">Editar</button>
      <button type="button" class="act-btn del" data-launch-del data-rub="${rubEnc}" data-mes="${esc(mes)}" data-tipo="${esc(tipo)}" data-col="${esc(launchCol(r))}">Zerar</button>
    </div></td>
  </tr>`;
}

function renderGroupedLaunchRows(list, grouped, editingKey) {
  if (!grouped) return list.map(r => launchRowHtml(r, editingKey)).join('');
  const order = ['Receita', 'Despesa', 'Resultado', 'Outros'];
  let html = '';
  order.forEach(nat => {
    const group = list.filter(r => launchNatureza(r) === nat);
    if (!group.length) return;
    html += `<tr class="group-row group-${nat.toLowerCase()}"><td colspan="6">
      ${naturezaBadge(nat)}<span class="group-count">${group.length} lançamento${group.length > 1 ? 's' : ''}</span>
    </td></tr>`;
    html += group.map(r => launchRowHtml(r, editingKey)).join('');
  });
  return html;
}

function renderFilterBar(meses) {
  const mesOpts = (meses || []).map(m => `<option value="${esc(m)}">${esc(m)}</option>`).join('');
  return `<div class="toolbar-filters">
    <div class="field filter-field">
      <label class="flabel">Natureza</label>
      <div class="chip-filter" id="natFilter">
        <button type="button" class="chip-fbtn on" data-nat="">Todos</button>
        <button type="button" class="chip-fbtn nat-rec" data-nat="Receita">Receita</button>
        <button type="button" class="chip-fbtn nat-desp" data-nat="Despesa">Despesa</button>
        <button type="button" class="chip-fbtn nat-res" data-nat="Resultado">Resultado</button>
        <button type="button" class="chip-fbtn nat-out" data-nat="Outros">Outros</button>
      </div>
    </div>
    <div class="filter-sep" aria-hidden="true"></div>
    <div class="field filter-field">
      <label class="flabel">Tipo</label>
      <div class="chip-filter" id="tipoFilter">
        <button type="button" class="chip-fbtn on" data-tipo="">Todos</button>
        <button type="button" class="chip-fbtn tipo-real" data-tipo="Realizado">Realizado</button>
        <button type="button" class="chip-fbtn tipo-prev" data-tipo="Previsto">Previsto</button>
      </div>
    </div>
    <div class="filter-sep" aria-hidden="true"></div>
    <div class="field filter-field filter-mes">
      <label class="flabel">Mês</label>
      <select id="mesFilter" class="filter-select">
        <option value="">Todos</option>${mesOpts}
      </select>
    </div>
    <div class="filter-sep" aria-hidden="true"></div>
    <div class="field filter-field filter-rubrica">
      <label class="flabel">Rubrica</label>
      <select id="rubricaFilter" class="filter-select">
        <option value="">Todas</option>
      </select>
    </div>
  </div>`;
}

function renderRubricaOptions(rubs) {
  const groups = { Receita: [], Despesa: [], Resultado: [], Outros: [] };
  rubs.forEach(r => {
    const n = rubricaNome(r);
    if (!n) return;
    const nat = rubricaNatureza(r);
    (groups[nat] || groups.Outros).push(n);
  });
  return ['Receita', 'Despesa', 'Resultado', 'Outros']
    .filter(g => groups[g].length)
    .map(g => `<optgroup label="${esc(g)}">${groups[g].map(n => `<option value="${esc(n)}">${esc(n)}</option>`).join('')}</optgroup>`)
    .join('');
}

function openModal(opts = {}) {
  const box = $('#modalBox');
  $('#modalTitle').textContent = opts.title || 'Confirmar ação';
  $('#modalBody').innerHTML = opts.html || `<p>${esc(opts.text || '')}</p>`;
  $('#modalOk').textContent = opts.confirmText || 'Confirmar';
  $('#modalOk').className = 'btn ' + (opts.danger ? 'btn-danger' : 'btn-accent');
  $('#modalCancel').textContent = opts.cancelText || 'Cancelar';
  $('#modalCancel').style.display = opts.hideCancel ? 'none' : '';
  box.classList.toggle('modal-wide', !!opts.wide);
  modalOnConfirm = opts.onConfirm || null;
  pendingDelete = opts.pendingDelete || null;
  $('#modal').classList.add('show');
}

function closeModal() {
  $('#modal').classList.remove('show');
  $('#modalBox').classList.remove('modal-wide');
  modalOnConfirm = null;
  pendingDelete = null;
}

function onModalOk() {
  if (modalOnConfirm) {
    const fn = modalOnConfirm;
    modalOnConfirm = null;
    fn();
    return;
  }
  if (pendingDelete) confirmDelete();
  else closeModal();
}

function confirmDetailHtml(items) {
  const rows = items.filter(x => x.value != null && x.value !== '').map(x =>
    `<dt>${esc(x.label)}</dt><dd>${esc(x.value)}</dd>`
  ).join('');
  return rows ? `<div class="confirm-detail"><dl>${rows}</dl></div>` : '';
}

function buildGradeTableHtml(matrixCache, rub, meses) {
  let totalReal = 0;
  let totalPrev = 0;
  let countReal = 0;
  let countPrev = 0;
  const rows = meses.map(m => {
    const real = findMatrixVal(matrixCache, rub, m, 'Realizado');
    const prev = findMatrixVal(matrixCache, rub, m, 'Previsto');
    totalReal += real;
    totalPrev += prev;
    if (Math.abs(real) > 0) countReal++;
    if (Math.abs(prev) > 0) countPrev++;
    return `<tr>
      <td class="col-mes">${esc(m)}</td>
      <td class="col-money num ${real ? 'has-val' : 'empty'}">${real ? brl(real) : '—'}</td>
      <td class="col-money num ${prev ? 'has-val' : 'empty'}">${prev ? brl(prev) : '—'}</td>
    </tr>`;
  }).join('');
  return `<div class="tscroll grade-scroll">
    <table class="grade-table" data-col-resize="grade">
      <colgroup>
        <col class="col-mes" style="width:28%">
        <col class="col-money" style="width:36%">
        <col class="col-money" style="width:36%">
      </colgroup>
      <thead><tr><th class="col-mes">Mês</th><th class="col-money">Realizado</th><th class="col-money">Previsto</th></tr></thead>
      <tbody>${rows}</tbody>
      <tfoot><tr>
        <td class="col-mes">Total (${countReal} / ${countPrev} meses)</td>
        <td class="col-money num">${brl(totalReal)}</td>
        <td class="col-money num">${brl(totalPrev)}</td>
      </tr></tfoot>
    </table>
  </div>`;
}

function toast(msg, danger) {
  const t = $('#toast');
  $('#toastText').textContent = msg;
  t.classList.toggle('danger', !!danger);
  t.classList.add('show');
  clearTimeout(toastT);
  toastT = setTimeout(() => t.classList.remove('show'), 2400);
}

function scale() {
  const pad = 48;
  const s = Math.min((window.innerWidth - pad) / 1280, (window.innerHeight - pad) / 800, 1.2);
  $('#canvas').style.transform = `scale(${s})`;
}

function findModule(id) {
  if (CONFIG.id === id) return CONFIG;
  if (isUnitModule(id)) return unitModules.find(u => u.id === id);
  return CONSOLIDATED.find(c => c.id === id) || PREMISSAS.find(p => p.id === id) || (RESUMO.id === id ? RESUMO : null);
}

async function refreshMeta() {
  meta = await FCApi.meta();
  unitModules = buildUnitModules(meta);
  renderNav();
}

async function init() {
  migrateColWidthStorage();
  scale();
  window.addEventListener('resize', scale);
  $('#btnSaveWorkbook').onclick = async () => {
    try {
      await FCApi.saveWorkbook();
      toast('Planilha salva');
    } catch (e) { toast(e.message, true); }
  };
  $('#btnRefresh').onclick = () => loadModule(currentModule, currentUnitTab);
  $('#modalCancel').onclick = closeModal;
  $('#modalOk').onclick = onModalOk;
  $('#modal').onclick = e => { if (e.target === $('#modal')) closeModal(); };
  try {
    await FCApi.ping();
    meta = await FCApi.meta();
    unitModules = buildUnitModules(meta);
    if (unitModules.length) currentModule = unitModules[0].id;
    renderNav();
    loadModule(currentModule);
  } catch (e) {
    $('#content').innerHTML = `<div class="empty-state">${esc(e.message)}</div>`;
  }
}

function renderUnitPicker() {
  const picker = $('#unitPicker');
  if (!picker) return;
  const cur = isUnitModule(currentModule) ? currentModule : '';
  picker.innerHTML = `<option value="">Selecione a unidade…</option>${unitModules.map(u =>
    `<option value="${esc(u.id)}"${u.id === cur ? ' selected' : ''}>${esc(u.label)}</option>`
  ).join('')}`;
}

function bindUnitPicker() {
  const picker = $('#unitPicker');
  if (!picker || picker.dataset.bound) return;
  picker.dataset.bound = '1';
  picker.onchange = () => {
    if (!picker.value) return;
    currentUnitTab = 'lancamentos';
    loadModule(picker.value);
  };
}

function renderNav() {
  renderUnitPicker();
  bindUnitPicker();
  const sections = [
    {
      id: 'consolidados',
      title: 'Consolidados',
      items: CONSOLIDATED.map(c => ({ id: c.id, label: c.label, icon: '◆' }))
    },
    {
      id: 'premissas',
      title: 'Premissas globais',
      items: PREMISSAS.map(p => ({ id: p.id, label: p.label, icon: '◇' }))
    },
    {
      id: 'visao',
      title: 'Visão',
      items: [{ id: RESUMO.id, label: RESUMO.label, icon: '◎' }]
    },
    {
      id: 'sistema',
      title: 'Sistema',
      items: [{ id: CONFIG.id, label: CONFIG.label, icon: '⚙' }]
    }
  ];

  $('#nav').innerHTML = sections.filter(sec => sec.items.length).map((sec, i) => `
    <div class="nav-section nav-section-${sec.id}${i > 0 ? ' nav-section-divider' : ''}">
      <div class="nav-section-title">${esc(sec.title)}</div>
      ${sec.items.map(it => `
        <button class="nav-item ${it.id === currentModule ? 'on' : ''}" data-mod="${esc(it.id)}" title="${esc(it.label)}">
          <span class="nav-icon">${it.icon}</span>
          <span class="nav-label">${esc(it.label)}</span>
        </button>`).join('')}
    </div>`).join('');

  $('#nav').onclick = e => {
    const b = e.target.closest('.nav-item');
    if (!b) return;
    currentUnitTab = 'lancamentos';
    loadModule(b.dataset.mod);
  };
}

function renderUnitTabs(sheet) {
  const bar = $('#unitTabs');
  bar.hidden = false;
  bar.innerHTML = UNIT_TABS.map(t =>
    `<button class="unit-tab ${t.id === currentUnitTab ? 'on' : ''}" data-utab="${t.id}">${esc(t.label)}</button>`
  ).join('');
  bar.onclick = e => {
    const b = e.target.closest('[data-utab]');
    if (!b) return;
    currentUnitTab = b.dataset.utab;
    loadModule(currentModule, currentUnitTab);
  };
}

function hideUnitChrome() {
  $('#unitTabs').hidden = true;
}

async function loadModule(id, unitTab) {
  if (unitTab) currentUnitTab = unitTab;
  currentModule = id;
  editingId = null;
  renderNav();

  const mod = findModule(id);
  if (!mod) return;

  if (mod.type === 'unit') {
    renderUnitTabs(mod.sheet);
    $('#pageTitle').textContent = mod.label;
    const picker = $('#unitPicker');
    if (picker) picker.value = mod.id;
  } else {
    hideUnitChrome();
    $('#pageTitle').textContent = mod.label;
    const picker = $('#unitPicker');
    if (picker && mod.type !== 'unit') picker.value = '';
  }

  $('#content').innerHTML = '<div class="loading">Carregando dados…</div>';
  try {
    if (mod.type === 'unit') await renderUnitModule(mod);
    else if (mod.type === 'tabular') await renderTabular(mod);
    else if (mod.type === 'premissas_matrix') await renderPremissasMatrix(mod);
    else if (mod.type === 'resumo') await renderResumo();
    else if (mod.type === 'config') await renderConfiguracoes();
  } catch (e) {
    $('#content').innerHTML = `<div class="empty-state">${esc(e.message)}</div>`;
    toast(e.message, true);
  }
}

async function renderUnitModule(mod) {
  if (currentUnitTab === 'lancamentos') {
    await renderUnitMatrix(mod.sheet);
    return;
  }
  const map = UNIT_TABULAR[currentUnitTab];
  if (!map) {
    $('#content').innerHTML = '<div class="empty-state">Seção não disponível.</div>';
    return;
  }
  await renderTabular(
    { id: map.entity, entity: map.entity, type: 'tabular', label: UNIT_TABS.find(t => t.id === currentUnitTab)?.label },
    { unitFilter: mod.sheet, unitField: map.unitField, lockUnit: true }
  );
}

async function renderTabular(mod, opts = {}) {
  if (!meta) meta = await FCApi.meta();
  const schema = TABULAR_SCHEMAS[mod.entity];
  if (!schema) throw new Error('Formulário não configurado: ' + mod.entity);

  let rows = await FCApi.list(mod.entity);
  if (!Array.isArray(rows)) throw new Error('Resposta inválida do servidor.');

  if (opts.unitFilter) {
    const uf = opts.unitField || 'unidade';
    const needle = opts.unitFilter.trim().toLowerCase();
    rows = rows.filter(r => String(r[uf] || '').trim().toLowerCase() === needle);
  }

  const formSchema = schema.map(f => {
    if (opts.lockUnit && f.lockOnUnit) return { ...f, readonly: true, defaultValue: opts.unitFilter };
    return f;
  });

  const formHtml = formSchema.map(f => fieldHtml(f, opts.unitFilter)).join('');
  const tableCols = formSchema.filter(f => f.type !== 'textarea' && !f.lockOnUnit).slice(0, 5);

  $('#content').innerHTML = `
    <div class="top-section">
      <div class="form-card glass">
        <div class="card-head"><span class="dot"></span> ${editingId ? 'Editar registro' : 'Novo registro'}${opts.unitFilter ? ` — <em>${esc(opts.unitFilter)}</em>` : ''}</div>
        <form id="crudForm" class="form-bar">${formHtml}</form>
        <div class="form-actions">
          <span class="form-msg" id="formMsg">Preencha os campos obrigatórios.</span>
          <button type="button" class="btn" id="btnClear">Limpar</button>
          <button type="button" class="btn btn-accent" id="btnSave">${editingId ? 'Atualizar' : 'Salvar'}</button>
        </div>
      </div>
    </div>
    <div class="table-card glass">
      <div class="table-toolbar">
        <div class="table-title">Registros <span class="count">${rows.length}</span></div>
        <div class="tsearch"><input id="search" placeholder="Buscar…"></div>
      </div>
      <div class="tscroll">
        <table data-col-resize="tabular-${mod.entity}">${tabularTableHeadHtml(tableCols)}<tbody id="tbody"></tbody></table>
      </div>
    </div>`;

  if (opts.lockUnit && opts.unitFilter) {
    formSchema.filter(f => f.lockOnUnit).forEach(f => {
      const el = document.querySelector(`[data-k="${f.key}"]`);
      if (el) setSelectValue(el, opts.unitFilter);
    });
  }

  const state = { rows, filter: '' };
  function paint() {
    const list = state.rows.filter(r => {
      if (!state.filter) return true;
      return JSON.stringify(r).toLowerCase().includes(state.filter);
    });
    const tbody = $('#tbody');
    if (!list.length) {
      tbody.innerHTML = `<tr><td colspan="${tableCols.length + 1}"><div class="empty-state">Nenhum registro${opts.unitFilter ? ' para esta unidade' : ''}</div></td></tr>`;
      scheduleTscrollSync($('#content'));
      return;
    }
    tbody.innerHTML = list.map(r => `<tr>
      ${tableCols.map(c => `<td class="${tableColClass(c)}">${formatFieldValue(c, r[c.key])}</td>`).join('')}
      <td class="col-acoes"><div class="acoes">
        <button class="act-btn" data-edit="${esc(r.id)}">Editar</button>
        <button class="act-btn del" data-del="${esc(r.id)}">Excluir</button>
      </div></td></tr>`).join('');
    scheduleTscrollSync($('#content'));
  }
  paint();

  $('#search').oninput = e => { state.filter = e.target.value.trim().toLowerCase(); paint(); };
  $('#btnClear').onclick = () => {
    if (editingId || document.querySelector('#crudForm input:not([disabled]), #crudForm select:not([disabled]), #crudForm textarea')?.value) {
      openModal({
        title: 'Limpar formulário',
        text: 'Descartar os dados preenchidos no formulário?',
        confirmText: 'Limpar',
        onConfirm: () => { closeModal(); editingId = null; loadModule(currentModule, currentUnitTab); }
      });
    } else {
      editingId = null;
      loadModule(currentModule, currentUnitTab);
    }
  };
  $('#btnSave').onclick = () => saveTabular(mod, formSchema, opts);
  $('#tbody').onclick = e => {
    const edit = e.target.closest('[data-edit]');
    const del = e.target.closest('[data-del]');
    if (edit) fillForm(mod, formSchema, state.rows.find(x => x.id === edit.dataset.edit), opts);
    if (del) askDelete(mod, del.dataset.del);
  };
  bindMoneyInputs();
  bindPercentInputs();
}

function fieldHtml(f, unitDefault) {
  const cls = 'field' + (f.grow ? ' grow' : '');
  if (f.type === 'textarea') return `<div class="${cls} field-textarea"><label class="flabel">${esc(f.label)}</label><textarea data-k="${f.key}" rows="1"></textarea></div>`;
  if (f.type === 'select') {
    const opts = fieldOptions(f);
    const reqMark = f.req ? ' <span class="req">*</span>' : '';
    const blank = f.req ? '' : '<option value="">— Selecione —</option>';
    const ro = f.readonly ? ' disabled' : '';
    const def = f.defaultValue || unitDefault || '';
    return `<div class="${cls}"><label class="flabel">${esc(f.label)}${reqMark}</label><select data-k="${f.key}"${ro}>${blank}${opts.map(o => `<option value="${esc(o)}"${def && o === def ? ' selected' : ''}>${esc(o)}</option>`).join('')}</select></div>`;
  }
  if (f.type === 'money') return `<div class="${cls}"><label class="flabel">${esc(f.label)}${f.req ? ' <span class="req">*</span>' : ''}</label><input data-k="${f.key}" data-type="money" placeholder="R$ 0,00"></div>`;
  if (f.type === 'percent') return `<div class="${cls}"><label class="flabel">${esc(f.label)}${f.req ? ' <span class="req">*</span>' : ''}</label><input data-k="${f.key}" data-type="percent" placeholder="0,00%" inputmode="decimal"></div>`;
  return `<div class="${cls}"><label class="flabel">${esc(f.label)}${f.req ? ' <span class="req">*</span>' : ''}</label><input data-k="${f.key}" ${f.type === 'number' ? 'data-type="number"' : ''}></div>`;
}

function bindPercentInputs(root) {
  (root || document).querySelectorAll('[data-type="percent"]').forEach(el => {
    el.addEventListener('input', e => {
      const ratio = pctRatioFromInput(e.target.value);
      e.target.dataset.ratio = String(ratio);
      e.target.value = pct(ratio);
    });
  });
}

function bindMoneyInputs(root) {
  (root || document).querySelectorAll('[data-type="money"]').forEach(el => {
    el.addEventListener('input', e => {
      let d = e.target.value.replace(/\D/g, '');
      const cents = d ? parseInt(d, 10) : 0;
      e.target.dataset.cents = cents;
      e.target.value = brl(cents);
    });
  });
}

function readForm(schema) {
  const data = {};
  schema.forEach(f => {
    const el = document.querySelector(`[data-k="${f.key}"]`);
    if (!el) return;
    if (f.type === 'money') data[f.key] = parseInt(el.dataset.cents || '0', 10);
    else if (f.type === 'percent') data[f.key] = el.dataset.ratio != null ? parseFloat(el.dataset.ratio) : pctRatioFromInput(el.value);
    else if (f.type === 'number') data[f.key] = parseFloat(el.value) || 0;
    else data[f.key] = el.value.trim();
  });
  if (editingId) data.id = editingId;
  return data;
}

function setSelectValue(el, value) {
  if (!el || value == null || value === '') return;
  const v = String(value);
  el.value = v;
  if (el.value !== v) {
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    el.appendChild(opt);
    el.value = v;
  }
}

function fillForm(mod, schema, row, opts = {}) {
  if (!row) return;
  editingId = row.id;
  loadModule(currentModule, currentUnitTab).then(() => {
    schema.forEach(f => {
      const el = document.querySelector(`[data-k="${f.key}"]`);
      if (!el || row[f.key] == null) return;
      if (f.type === 'money') { el.dataset.cents = row[f.key]; el.value = brl(row[f.key]); }
      else if (f.type === 'percent') { el.dataset.ratio = row[f.key]; el.value = pct(row[f.key]); }
      else if (f.type === 'select') setSelectValue(el, row[f.key]);
      else el.value = row[f.key];
    });
    if (opts.lockUnit && opts.unitFilter) {
      schema.filter(f => f.lockOnUnit).forEach(f => {
        const el = document.querySelector(`[data-k="${f.key}"]`);
        if (el) setSelectValue(el, opts.unitFilter);
      });
    }
    $('#btnSave').textContent = 'Atualizar';
  });
}

async function saveTabular(mod, schema, opts = {}) {
  const data = readForm(schema);
  if (opts.lockUnit && opts.unitFilter) {
    const uf = schema.find(f => f.lockOnUnit);
    if (uf) data[uf.key] = opts.unitFilter;
  }
  for (const f of schema.filter(x => x.req)) {
    if ((f.type === 'money' && !(data[f.key] > 0)) || (data[f.key] == null || data[f.key] === '') && f.type !== 'money' && f.type !== 'percent') {
      $('#formMsg').classList.add('show');
      return;
    }
  }

  const action = editingId ? 'Atualizar registro' : 'Criar registro';
  const detail = schema
    .filter(f => f.type !== 'textarea' && data[f.key] != null && data[f.key] !== '')
    .slice(0, 6)
    .map(f => ({
      label: f.label,
      value: f.type === 'money' ? brl(data[f.key]) : f.type === 'percent' ? pct(data[f.key]) : String(data[f.key])
    }));

  openModal({
    title: action,
    html: `<p>Confirma ${editingId ? 'a atualização' : 'a criação'} deste registro na planilha?</p>${confirmDetailHtml(detail)}`,
    confirmText: editingId ? 'Atualizar' : 'Salvar',
    onConfirm: async () => {
      closeModal();
      try {
        if (editingId) await FCApi.update(mod.entity, data);
        else await FCApi.create(mod.entity, data);
        toast(editingId ? 'Registro atualizado' : 'Registro criado');
        loadModule(currentModule, currentUnitTab);
      } catch (e) { toast(e.message, true); }
    }
  });
}

function askDelete(mod, id) {
  openModal({
    title: 'Excluir registro',
    text: 'Esta ação remove o registro da planilha Excel. Deseja continuar?',
    confirmText: 'Excluir',
    danger: true,
    pendingDelete: { mod, id }
  });
}

async function confirmDelete() {
  if (!pendingDelete) return closeModal();
  try {
    await FCApi.remove(pendingDelete.mod.entity, { id: pendingDelete.id });
    toast('Registro excluído', true);
    closeModal();
    loadModule(currentModule, currentUnitTab);
  } catch (e) { toast(e.message, true); closeModal(); }
}

async function renderUnitMatrix(sheet) {
  if (!meta) meta = await FCApi.meta();
  const meses = meta?.meses || [];
  const tipos = meta?.tipos || ['Realizado', 'Previsto'];

  $('#content').innerHTML = `
    <div class="matrix-page" id="matrixPage">
      <div class="form-card glass matrix-form-compact" id="matrixFormCard">
        <div class="card-head"><span class="dot"></span> Lançamento</div>
        <div class="form-bar">
          <div class="field grow"><label class="flabel">Rubrica</label><select id="mRubrica"></select></div>
          <div class="field"><label class="flabel">Mês</label><select id="mMes">${meses.map(m => `<option>${esc(m)}</option>`).join('')}</select></div>
          <div class="field"><label class="flabel">Tipo</label><select id="mTipo">${tipos.map(t => `<option>${esc(t)}</option>`).join('')}</select></div>
          <div class="field field-valor"><label class="flabel">Valor <span class="req">*</span></label><input id="mValor" data-type="money"></div>
        </div>
        <div class="form-actions">
          <button type="button" class="btn" id="mClear">Limpar</button>
          <button class="btn btn-accent" id="mSave">Salvar lançamento</button>
          <button class="btn btn-danger btn-soft" id="mDel">Zerar valor</button>
          <button class="btn btn-soft" id="btnToggleGrade">Ver grade mensal</button>
        </div>
      </div>
      <div class="table-card glass launches-panel">
        <div class="table-toolbar">
          <div class="table-toolbar-head">
            <div class="table-title" id="launchPanelTitle">Lançamentos com valor — ${esc(unitLabel(sheet))} <span class="count" id="launchCount">0</span></div>
            <button type="button" class="panel-expand-btn" id="btnExpandLaunches" title="Expandir lista" aria-label="Expandir lista" aria-expanded="false">
              <svg class="icon-expand" viewBox="0 0 24 24" aria-hidden="true"><path d="M8 3H3v5M16 3h5v5M3 16v5h5M21 16v5h-5M8 8L3 3M16 8l5-5M8 16l-5 5M16 16l5 5"/></svg>
              <svg class="icon-collapse" viewBox="0 0 24 24" aria-hidden="true"><path d="M8 3v5H3M16 3v5h5M3 16h5v5M21 16h-5v5M8 8L3 3M16 8l5-5M8 16l-5 5M16 16l5 5"/></svg>
            </button>
          </div>
          ${renderFilterBar(meses)}
        </div>
        <div id="launchPanel" class="launch-panel-body">
          <div class="tscroll launch-scroll">
            <table data-col-resize="launch-crud-${sheet}">${launchTableHeadHtml()}<tbody id="mTbody"></tbody></table>
          </div>
        </div>
      </div>
    </div>`;

  bindMoneyInputs($('#content'));
  let matrixCache = [];
  let allLaunches = [];
  let rubricaFilter = '';
  let natureFilter = '';
  let tipoFilter = '';
  let mesFilter = '';
  let showGradeView = false;
  let launchesExpanded = false;
  let matrixEditingKey = null;
  let matrixEditingOrig = null;
  let matrixColumnLayout = [];

  function matrixPayload(valorCents) {
    const rubrica = $('#mRubrica').value;
    const mes = $('#mMes').value;
    const tipo = normalizeTipo($('#mTipo').value);
    const payload = { sheet, rubrica, mes, tipo, valorCents };

    if (!alignMatrixPayload(payload, matrixColumnLayout)) return null;

    if (!matrixEditingOrig?.coluna) return payload;

    const samePhysicalCol = String(payload.coluna || '').toUpperCase()
      === String(matrixEditingOrig.coluna || '').toUpperCase();

    if (!(samePhysicalCol && rubrica === matrixEditingOrig.rubrica) && matrixEditingOrig.coluna)
      payload.colunaOrigem = matrixEditingOrig.coluna;

    return payload;
  }

  function findLaunchRow(rub, mes, tipo, col) {
    const rubNorm = decodeURIComponent(String(rub || '')).trim();
    const colNorm = String(col || '').trim().toUpperCase();
    if (colNorm) {
      return allLaunches.find(r =>
        launchRubrica(r) === rubNorm && launchCol(r) === colNorm
      );
    }
    const tipoNorm = normalizeTipo(tipo);
    return allLaunches.find(r =>
      launchRubrica(r) === rubNorm &&
      mesMatch(r.mes, mes) &&
      normalizeTipo(r.tipo || r.Tipo) === tipoNorm
    );
  }

  function clearLaunchForm() {
    matrixEditingKey = null;
    matrixEditingOrig = null;
    $('#mValor').value = '';
    delete $('#mValor').dataset.cents;
    $('#mSave').textContent = 'Salvar lançamento';
    $('#matrixFormCard')?.classList.remove('form-editing');
    if (!showGradeView) paintList();
  }

  function fillLaunchForm(r) {
    setSelectValue($('#mRubrica'), launchRubrica(r));
    setSelectValue($('#mMes'), r.mes);
    setSelectValue($('#mTipo'), normalizeTipo(r.tipo || r.Tipo));
    const col = launchCol(r);
    const cents = launchCents(r);
    $('#mValor').dataset.cents = String(cents);
    $('#mValor').value = brl(cents);
    matrixEditingKey = launchKey(r);
    matrixEditingOrig = {
      rubrica: launchRubrica(r),
      mes: String(r.mes || '').trim(),
      tipo: normalizeTipo(r.tipo || r.Tipo),
      coluna: col || null,
      valorCents: cents
    };
    $('#mSave').textContent = 'Atualizar lançamento';
    $('#matrixFormCard')?.classList.add('form-editing');
    $('#matrixFormCard')?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    if (!showGradeView) paintList();
  }

  function confirmDeleteLaunch(r) {
    const rub = launchRubrica(r);
    const mes = r.mes;
    const tipo = String(r.tipo || r.Tipo || '').trim();
    const payload = {
      sheet,
      rubrica: rub,
      mes,
      tipo,
      valorCents: 0
    };
    const col = launchCol(r);
    if (col) payload.coluna = col;
    openModal({
      title: 'Zerar lançamento',
      html: `<p>Remover o valor de <strong>${esc(mes)}</strong> (${esc(tipo)}) para a rubrica abaixo?</p>${confirmDetailHtml([
        { label: 'Rubrica', value: rub },
        { label: 'Valor atual', value: brl(launchCents(r)) }
      ])}`,
      confirmText: 'Zerar',
      danger: true,
      onConfirm: async () => {
        closeModal();
        try {
          await FCApi.remove('matrix', payload);
          toast('Lançamento zerado', true);
          if (matrixEditingKey === launchKey(r)) clearLaunchForm();
          await reloadSheetData();
          if (showGradeView) paintGradeView();
        } catch (e) { toast(e.message, true); }
      }
    });
  }

  function handleLaunchTableClick(e) {
    if (showGradeView) return;
    const editBtn = e.target.closest('[data-launch-edit]');
    const delBtn = e.target.closest('[data-launch-del]');
    if (editBtn) {
      const row = findLaunchRow(editBtn.dataset.rub, editBtn.dataset.mes, editBtn.dataset.tipo, editBtn.dataset.col);
      if (row) fillLaunchForm(row);
      return;
    }
    if (delBtn) {
      const row = findLaunchRow(delBtn.dataset.rub, delBtn.dataset.mes, delBtn.dataset.tipo, delBtn.dataset.col);
      if (row) confirmDeleteLaunch(row);
    }
  }

  function updateLaunchesExpand() {
    const page = $('#matrixPage');
    const btn = $('#btnExpandLaunches');
    if (!page || !btn) return;
    page.classList.toggle('launches-expanded', launchesExpanded);
    btn.classList.toggle('is-expanded', launchesExpanded);
    btn.setAttribute('aria-expanded', launchesExpanded ? 'true' : 'false');
    const label = launchesExpanded ? 'Contrair e exibir lançamento' : 'Expandir lista';
    btn.title = label;
    btn.setAttribute('aria-label', label);
  }

  $('#btnExpandLaunches').onclick = () => {
    launchesExpanded = !launchesExpanded;
    updateLaunchesExpand();
    scheduleTscrollSync($('#content'));
  };

  async function loadColumnLayout() {
    try {
      const cols = await FCApi.list('matrix_columns', { sheet });
      matrixColumnLayout = Array.isArray(cols) ? cols : [];
    } catch {
      matrixColumnLayout = [];
    }
  }

  async function loadRubrics() {
    const rubs = await FCApi.rubricas(sheet);
    const arr = Array.isArray(rubs) ? rubs : [];
    $('#mRubrica').innerHTML = renderRubricaOptions(arr);
    const rf = $('#rubricaFilter');
    if (rf) rf.innerHTML = `<option value="">Todas</option>${renderRubricaOptions(arr)}`;
    await reloadSheetData();
  }

  async function loadAllLaunches() {
    allLaunches = await FCApi.list('matrix', { sheet });
    if (!Array.isArray(allLaunches)) allLaunches = [];
    if (!showGradeView) paintList();
  }

  async function refreshMatrix() {
    const rub = $('#mRubrica').value;
    matrixCache = await FCApi.list('matrix', { sheet, rubrica: rub });
    if (showGradeView) paintGradeView();
  }

  async function reloadSheetData() {
    await loadColumnLayout();
    await loadAllLaunches();
    await refreshMatrix();
  }

  function paintGradeView() {
    const rub = $('#mRubrica').value;
    const nat = classifyNaturezaLocal(rub);
    $('#launchPanelTitle').innerHTML = `Grade mensal — ${naturezaBadge(nat)} <span id="mRubTitle">${esc(rub || '—')}</span> · ${esc(unitLabel(sheet))}`;
    $('#launchPanel').innerHTML = buildGradeTableHtml(matrixCache, rub, meta?.meses || []);
    document.querySelector('.toolbar-filters')?.style.setProperty('display', 'none');
    scheduleTscrollSync($('#content'));
  }

  function paintListView() {
    $('#launchPanelTitle').innerHTML = `Lançamentos com valor — ${esc(unitLabel(sheet))} <span class="count" id="launchCount">0</span>`;
    $('#launchPanel').innerHTML = `<div class="tscroll launch-scroll">
      <table data-col-resize="launch-crud-${sheet}">${launchTableHeadHtml()}<tbody id="mTbody"></tbody></table>
    </div>`;
    scheduleTscrollSync($('#content'));
    document.querySelector('.toolbar-filters')?.style.setProperty('display', '');
    paintList();
  }

  function toggleGradeView() {
    showGradeView = !showGradeView;
    const btn = $('#btnToggleGrade');
    btn.textContent = showGradeView ? 'Ver lançamentos' : 'Ver grade mensal';
    btn.classList.toggle('btn-accent', showGradeView);
    if (showGradeView) {
      refreshMatrix().then(() => paintGradeView());
    } else {
      paintListView();
    }
  }

  function paintList() {
    const withVal = launchesWithValue(allLaunches);
    let list = withVal;
    if (natureFilter) list = list.filter(r => launchNatureza(r) === natureFilter);
    if (tipoFilter) list = list.filter(r => String(r.tipo || r.Tipo || '').trim() === tipoFilter);
    if (mesFilter) list = list.filter(r => mesMatch(r.mes, mesFilter));
    if (rubricaFilter) list = list.filter(r => launchRubrica(r) === rubricaFilter);
    const countEl = $('#launchCount');
    if (countEl) countEl.textContent = String(withVal.length);
    const tbody = $('#mTbody');
    if (!tbody) return;
    if (!list.length) {
      tbody.innerHTML = `<tr><td colspan="6"><div class="empty-state">${rubricaFilter || natureFilter || tipoFilter || mesFilter ? 'Nenhum lançamento corresponde aos filtros' : 'Nenhum lançamento com valor'}</div></td></tr>`;
      scheduleTscrollSync($('#content'));
      return;
    }
    tbody.innerHTML = renderGroupedLaunchRows(list, !natureFilter && !tipoFilter && !mesFilter && !rubricaFilter, matrixEditingKey);
    scheduleTscrollSync($('#content'));
  }

  $('#rubricaFilter').onchange = e => {
    rubricaFilter = e.target.value;
    paintList();
  };

  $('#mesFilter').onchange = e => {
    mesFilter = e.target.value;
    paintList();
  };

  $('#natFilter').onclick = e => {
    const b = e.target.closest('[data-nat]');
    if (!b) return;
    natureFilter = b.dataset.nat;
    document.querySelectorAll('#natFilter .chip-fbtn').forEach(x => x.classList.toggle('on', x === b));
    paintList();
  };

  $('#tipoFilter').onclick = e => {
    const b = e.target.closest('[data-tipo]');
    if (!b) return;
    tipoFilter = b.dataset.tipo;
    document.querySelectorAll('#tipoFilter .chip-fbtn').forEach(x => x.classList.toggle('on', x === b));
    paintList();
  };

  $('#btnToggleGrade').onclick = toggleGradeView;
  $('#matrixPage').addEventListener('click', handleLaunchTableClick);
  $('#mClear').onclick = () => clearLaunchForm();
  $('#mRubrica').onchange = () => {
    refreshMatrix().then(() => {
      if (showGradeView) paintGradeView();
    });
  };
  $('#mMes').onchange = () => suggestTipoForMes($('#mMes').value, matrixColumnLayout);

  $('#mSave').onclick = () => {
    const payload = matrixPayload(parseInt($('#mValor').dataset.cents || '0', 10));
    if (!payload) return;
    if (!payload.valorCents) {
      toast('Informe o valor do lançamento', true);
      return;
    }
    const isUpdate = !!matrixEditingKey;
    const tipoHeaderAtual = tipoDoMes(payload.mes, matrixColumnLayout);
    const tipoColunaMuda = tipoHeaderAtual && normalizeTipo(payload.tipo) !== tipoHeaderAtual;
    const detalhes = [
      { label: 'Natureza', value: launchNatureza({ rubrica: payload.rubrica, natureza: classifyNaturezaLocal(payload.rubrica) }) },
      { label: 'Rubrica', value: payload.rubrica },
      { label: 'Mês', value: payload.mes },
      { label: 'Tipo', value: payload.tipo },
      { label: 'Valor', value: brl(payload.valorCents) }
    ];
    if (tipoColunaMuda) {
      detalhes.push({ label: 'Tipo do mês', value: `${tipoHeaderAtual} → ${payload.tipo} (todas as rubricas)` });
    }
    openModal({
      title: isUpdate ? 'Atualizar lançamento' : 'Salvar lançamento',
      html: `<p>Confirma ${isUpdate ? 'a atualização' : 'o lançamento'} na aba <strong>${esc(sheet)}</strong>?</p>${confirmDetailHtml(detalhes)}`,
      confirmText: isUpdate ? 'Atualizar' : 'Salvar',
      onConfirm: async () => {
        closeModal();
        try {
          await FCApi.upsert('matrix', payload);
          toast(isUpdate ? 'Lançamento atualizado' : 'Lançamento salvo');
          clearLaunchForm();
          await reloadSheetData();
          if (showGradeView) paintGradeView();
        } catch (e) { toast(e.message, true); }
      }
    });
  };
  $('#mDel').onclick = () => {
    const payload = matrixPayload(0);
    if (!payload) return;
    if (!payload.rubrica) {
      toast('Selecione a rubrica', true);
      return;
    }
    openModal({
      title: 'Zerar valor',
      text: `Remover o valor de ${payload.mes} (${payload.tipo}) para "${payload.rubrica}"?`,
      confirmText: 'Zerar',
      danger: true,
      onConfirm: async () => {
        closeModal();
        try {
          await FCApi.remove('matrix', payload);
          toast('Valor zerado', true);
          clearLaunchForm();
          await reloadSheetData();
          if (showGradeView) paintGradeView();
        } catch (e) { toast(e.message, true); }
      }
    });
  };

  await loadColumnLayout();
  await loadRubrics();
}

async function renderPremissasMatrix(mod) {
  meta = meta || await FCApi.meta();
  const meses = meta.meses || [];
  const sheet = mod.sheet;
  let premRows = [];
  let premEditing = null;
  let premColumnLayout = [];

  $('#content').innerHTML = `
    <div class="form-card glass" id="premFormCard">
      <div class="card-head"><span class="dot"></span> ${esc(mod.label)}</div>
      <div class="form-bar">
        <div class="field grow"><label class="flabel">Premissa / Rubrica</label><select id="pRubrica"></select></div>
        <div class="field"><label class="flabel">Mês</label><select id="pMes">${meses.map(m => `<option>${esc(m)}</option>`).join('')}</select></div>
        <div class="field"><label class="flabel">Valor %</label><input id="pValor" data-type="percent" placeholder="0,00%" inputmode="decimal"></div>
      </div>
      <div class="form-actions">
        <button type="button" class="btn" id="pClear">Limpar</button>
        <button class="btn btn-accent" id="pSave">Salvar</button>
      </div>
    </div>
    <div class="table-card glass">
      <div class="table-toolbar"><div class="table-title">Valores cadastrados <span class="count" id="premCount">0</span></div></div>
      <div class="tscroll"><table data-col-resize="prem-${sheet}">
        <colgroup><col style="width:44%"><col style="width:16%"><col style="width:16%"><col class="col-acoes" style="width:24%"></colgroup>
        <thead><tr><th>Premissa</th><th>Mês</th><th>Valor %</th><th class="col-acoes">Ações</th></tr></thead>
        <tbody id="premTbody"></tbody>
      </table></div>
    </div>`;

  bindPercentInputs($('#content'));

  function clearPremForm() {
    premEditing = null;
    $('#pValor').value = '';
    delete $('#pValor').dataset.ratio;
    $('#pSave').textContent = 'Salvar';
    $('#premFormCard')?.classList.remove('form-editing');
    paintPremTable();
  }

  function fillPremForm(r) {
    setSelectValue($('#pRubrica'), launchRubrica(r));
    setSelectValue($('#pMes'), r.mes);
    const ratio = launchCents(r) / 100;
    $('#pValor').dataset.ratio = String(ratio);
    $('#pValor').value = pct(ratio);
    premEditing = {
      rubrica: launchRubrica(r),
      mes: String(r.mes || '').trim(),
      coluna: launchCol(r) || null
    };
    $('#pSave').textContent = 'Atualizar';
    $('#premFormCard')?.classList.add('form-editing');
    paintPremTable();
  }

  function premPayload(valorCents) {
    const mes = $('#pMes').value;
    const tipo = tipoUnicoParaMes(mes, premColumnLayout) || 'Realizado';
    const payload = {
      sheet,
      rubrica: $('#pRubrica').value,
      mes,
      tipo,
      valorCents
    };
    if (!premEditing) return payload;
    const same = payload.rubrica === premEditing.rubrica && mesMatch(payload.mes, premEditing.mes);
    if (premEditing.coluna) {
      if (same) payload.coluna = premEditing.coluna;
      else payload.colunaOrigem = premEditing.coluna;
    }
    return payload;
  }

  function paintPremTable() {
    const withVal = (premRows || []).filter(r => Math.abs(launchCents(r)) > 0);
    $('#premCount').textContent = String(withVal.length);
    const tbody = $('#premTbody');
    if (!withVal.length) {
      tbody.innerHTML = '<tr><td colspan="4"><div class="empty-state">Nenhum valor cadastrado</div></td></tr>';
      scheduleTscrollSync($('#content'));
      return;
    }
    tbody.innerHTML = withVal.map(r => {
      const rub = launchRubrica(r);
      const mes = String(r.mes || '').trim();
      const rubEnc = encodeURIComponent(rub);
      const editing = premEditing && premEditing.rubrica === rub && mesMatch(mes, premEditing.mes);
      return `<tr class="${editing ? 'row-editing' : ''}">
        <td class="col-grow">${esc(rub)}</td>
        <td>${esc(mes)}</td>
        <td class="col-pct">${pct(launchCents(r) / 100)}</td>
        <td class="col-acoes"><div class="acoes">
          <button type="button" class="act-btn" data-prem-edit data-rub="${rubEnc}" data-mes="${esc(mes)}">Editar</button>
          <button type="button" class="act-btn del" data-prem-del data-rub="${rubEnc}" data-mes="${esc(mes)}">Zerar</button>
        </div></td>
      </tr>`;
    }).join('');
    scheduleTscrollSync($('#content'));
  }

  async function loadPremData() {
    premRows = await FCApi.list('matrix', { sheet });
    if (!Array.isArray(premRows)) premRows = [];
    paintPremTable();
  }

  async function loadR() {
    try {
      const cols = await FCApi.list('matrix_columns', { sheet });
      premColumnLayout = Array.isArray(cols) ? cols : [];
    } catch {
      premColumnLayout = [];
    }
    const rubs = await FCApi.rubricas(sheet);
    $('#pRubrica').innerHTML = (Array.isArray(rubs) ? rubs : []).map(r => `<option>${esc(rubricaNome(r))}</option>`).join('');
    await loadPremData();
  }

  $('#content').onclick = e => {
    const edit = e.target.closest('[data-prem-edit]');
    const del = e.target.closest('[data-prem-del]');
    if (edit) {
      const row = premRows.find(r => launchRubrica(r) === decodeURIComponent(edit.dataset.rub) && mesMatch(r.mes, edit.dataset.mes));
      if (row) fillPremForm(row);
    }
    if (del) {
      const row = premRows.find(r => launchRubrica(r) === decodeURIComponent(del.dataset.rub) && mesMatch(r.mes, del.dataset.mes));
      if (!row) return;
      const payload = { sheet, rubrica: launchRubrica(row), mes: row.mes, tipo: 'Realizado', valorCents: 0 };
      if (launchCol(row)) payload.coluna = launchCol(row);
      openModal({
        title: 'Zerar premissa',
        text: `Remover o valor de ${row.mes} para "${launchRubrica(row)}"?`,
        confirmText: 'Zerar',
        danger: true,
        onConfirm: async () => {
          closeModal();
          try {
            await FCApi.remove('matrix', payload);
            toast('Valor zerado', true);
            if (premEditing && premEditing.rubrica === launchRubrica(row) && mesMatch(row.mes, premEditing.mes)) clearPremForm();
            await loadPremData();
          } catch (err) { toast(err.message, true); }
        }
      });
    }
  };

  $('#pClear').onclick = () => clearPremForm();
  $('#pSave').onclick = () => {
    const el = $('#pValor');
    const val = el.dataset.ratio != null ? parseFloat(el.dataset.ratio) : pctRatioFromInput(el.value);
    const payload = premPayload(Math.round(val * 100));
    const isUpdate = !!premEditing;
    openModal({
      title: isUpdate ? 'Atualizar premissa' : 'Salvar premissa',
      html: `<p>Confirma ${isUpdate ? 'a atualização' : 'o cadastro'} na aba <strong>${esc(sheet)}</strong>?</p>${confirmDetailHtml([
        { label: 'Premissa', value: payload.rubrica },
        { label: 'Mês', value: payload.mes },
        { label: 'Valor %', value: pct(val) }
      ])}`,
      confirmText: isUpdate ? 'Atualizar' : 'Salvar',
      onConfirm: async () => {
        closeModal();
        try {
          await FCApi.upsert('matrix', payload);
          toast(isUpdate ? 'Premissa atualizada' : 'Premissa salva');
          clearPremForm();
          await loadPremData();
        } catch (e) { toast(e.message, true); }
      }
    });
  };
  await loadR();
}

async function renderConfiguracoes() {
  if (!meta) meta = await FCApi.meta();
  const units = (meta.matrix || [])
    .filter(x => x.entity && x.entity.startsWith('matrix_') &&
      !x.entity.includes('premissas') && !x.entity.includes('capex0403'))
    .map(x => x.sheet);

  $('#content').innerHTML = `
    <div class="config-page">
      <div class="config-row">
        <div class="form-card glass config-block">
          <div class="card-head"><span class="dot"></span> Planilha Excel</div>
          <p class="config-desc">Exibe a janela do Excel com as abas e células da base de dados (planilha oculta ao abrir o Fluxo).</p>
          <button type="button" class="btn btn-soft" id="cfgShowExcel">Mostrar Base</button>
        </div>
        <div class="form-card glass config-block">
          <div class="card-head"><span class="dot"></span> Download</div>
          <p class="config-desc">Salve uma cópia de backup do workbook atual (.xlsx) em qualquer pasta do computador.</p>
          <button class="btn btn-accent" id="cfgDownload">Baixar planilha</button>
        </div>
      </div>

      <div class="config-row">
        <div class="form-card glass config-block">
          <div class="card-head"><span class="dot"></span> Responsáveis</div>
          <div class="config-form">
            <div class="field grow"><label class="flabel">Nome</label><input id="cfgRespNome" placeholder="Novo responsável"></div>
            <button class="btn btn-accent" id="cfgRespAdd">Incluir</button>
          </div>
          <div class="tscroll config-scroll"><table data-col-resize="config-responsaveis">
            <colgroup><col class="col-grow" style="width:48%"><col class="col-text" style="width:32%"><col class="col-acoes" style="width:20%"></colgroup>
            <thead><tr><th class="col-grow">Nome</th><th class="col-text">Origem</th><th class="col-acoes">Ações</th></tr></thead>
            <tbody id="cfgRespBody"></tbody></table></div>
        </div>

        <div class="form-card glass config-block">
          <div class="card-head"><span class="dot"></span> Unidades</div>
          <div class="config-form">
            <div class="field grow"><label class="flabel">Nome da aba</label><input id="cfgUnitNome" placeholder="Ex.: NOVA UND"></div>
            <div class="field"><label class="flabel">Modelo</label><select id="cfgUnitTpl"><option value="ALS">ALS</option></select></div>
            <button class="btn btn-accent" id="cfgUnitAdd">Incluir</button>
          </div>
          <div class="tscroll config-scroll"><table data-col-resize="config-unidades">
            <colgroup><col class="col-grow" style="width:78%"><col class="col-acoes" style="width:22%"></colgroup>
            <thead><tr><th class="col-grow">Unidade (aba)</th><th class="col-acoes">Ações</th></tr></thead>
            <tbody id="cfgUnitBody"></tbody></table></div>
        </div>
      </div>

      <div class="form-card glass config-block">
        <div class="card-head"><span class="dot"></span> Rubricas</div>
        <div class="config-form">
          <div class="field"><label class="flabel">Unidade</label><select id="cfgRubSheet"></select></div>
          <div class="field grow"><label class="flabel">Rubrica</label><input id="cfgRubNome" placeholder="Ex.: (+) Nova receita"></div>
          <div class="field"><label class="flabel">Seção (col. A)</label><input id="cfgRubSecao" placeholder="Opcional"></div>
          <button class="btn btn-accent" id="cfgRubAdd">Incluir</button>
        </div>
        <div class="tscroll config-scroll"><table data-col-resize="config-rubricas">
          <colgroup><col class="col-grow" style="width:42%"><col class="col-text" style="width:16%"><col class="col-text" style="width:22%"><col class="col-acoes" style="width:20%"></colgroup>
          <thead><tr><th class="col-grow">Rubrica</th><th class="col-text">Natureza</th><th class="col-text">Seção</th><th class="col-acoes">Ações</th></tr></thead>
          <tbody id="cfgRubBody"></tbody></table></div>
      </div>
    </div>`;

  $('#cfgUnitTpl').innerHTML = units.map(u => `<option value="${esc(u)}">${esc(u)}</option>`).join('');
  $('#cfgRubSheet').innerHTML = units.map(u => `<option value="${esc(u)}">${esc(u)}</option>`).join('');

  function confirmRemove(label, name, onConfirm) {
    openModal({
      title: `Excluir ${label}`,
      html: `<p>Confirma a exclusão de <strong>${esc(name)}</strong>?</p>`,
      confirmText: 'Excluir',
      danger: true,
      onConfirm: async () => {
        closeModal();
        try { await onConfirm(); } catch (e) { toast(e.message, true); }
      }
    });
  }

  async function paintResponsaveis() {
    const rows = await FCApi.configList('config_responsaveis');
    const tbody = $('#cfgRespBody');
    if (!rows.length) {
      tbody.innerHTML = '<tr><td colspan="3"><div class="empty-state">Nenhum responsável cadastrado</div></td></tr>';
      scheduleTscrollSync($('#content'));
      return;
    }
    tbody.innerHTML = rows.map(r => {
      const origem = r.origem === 'config' ? 'Configuração' : r.origem === 'planilha' ? 'Planilha' : 'Config. + Planilha';
      const del = r.removivel ? `<button class="act-btn del" data-resp="${esc(r.nome)}">Excluir</button>` : '';
      return `<tr><td class="col-grow">${esc(r.nome)}</td><td class="col-text">${esc(origem)}</td><td class="col-acoes acoes">${del}</td></tr>`;
    }).join('');
    scheduleTscrollSync($('#content'));
    tbody.querySelectorAll('[data-resp]').forEach(btn => {
      btn.onclick = () => confirmRemove('Responsável', btn.dataset.resp, async () => {
        await FCApi.configRemove('config_responsaveis', { nome: btn.dataset.resp });
        toast('Responsável removido');
        await paintResponsaveis();
        await refreshMeta();
      });
    });
  }

  async function paintUnidades() {
    const rows = await FCApi.configList('config_unidades');
    const tbody = $('#cfgUnitBody');
    if (!rows.length) {
      tbody.innerHTML = '<tr><td colspan="2"><div class="empty-state">Nenhuma unidade</div></td></tr>';
      scheduleTscrollSync($('#content'));
      return;
    }
    tbody.innerHTML = rows.map(r => `<tr>
      <td class="col-grow">${esc(r.label || r.nome)}</td>
      <td class="col-acoes acoes"><button class="act-btn del" data-unit="${esc(r.nome)}">Excluir</button></td>
    </tr>`).join('');
    scheduleTscrollSync($('#content'));
    tbody.querySelectorAll('[data-unit]').forEach(btn => {
      btn.onclick = () => confirmRemove('Unidade', btn.dataset.unit, async () => {
        await FCApi.configRemove('config_unidades', { nome: btn.dataset.unit });
        toast('Unidade removida');
        await refreshMeta();
        await renderConfiguracoes();
      });
    });
  }

  async function paintRubricas() {
    const sheet = $('#cfgRubSheet').value;
    const rows = await FCApi.configList('config_rubricas', { sheet });
    const tbody = $('#cfgRubBody');
    if (!rows.length) {
      tbody.innerHTML = '<tr><td colspan="4"><div class="empty-state">Nenhuma rubrica nesta unidade</div></td></tr>';
      scheduleTscrollSync($('#content'));
      return;
    }
    tbody.innerHTML = rows.map(r => `<tr>
      <td class="col-grow">${esc(r.nome)}</td>
      <td class="col-text">${naturezaBadge(r.natureza)}</td>
      <td class="col-text">${esc(r.secao || '—')}</td>
      <td class="col-acoes acoes"><button class="act-btn del" data-rub="${esc(r.nome)}">Excluir</button></td>
    </tr>`).join('');
    scheduleTscrollSync($('#content'));
    tbody.querySelectorAll('[data-rub]').forEach(btn => {
      btn.onclick = () => confirmRemove('Rubrica', btn.dataset.rub, async () => {
        await FCApi.configRemove('config_rubricas', { sheet, nome: btn.dataset.rub });
        toast('Rubrica removida');
        await paintRubricas();
      });
    });
  }

  $('#cfgDownload').onclick = async () => {
    try {
      const r = await FCApi.download();
      toast(`Salvo: ${r.fileName || 'planilha'}`);
    } catch (e) { toast(e.message, true); }
  };

  $('#cfgRespAdd').onclick = async () => {
    const nome = $('#cfgRespNome').value.trim();
    if (!nome) return toast('Informe o nome do responsável', true);
    try {
      await FCApi.configAdd('config_responsaveis', { nome });
      $('#cfgRespNome').value = '';
      toast('Responsável incluído');
      await paintResponsaveis();
      await refreshMeta();
    } catch (e) { toast(e.message, true); }
  };

  $('#cfgUnitAdd').onclick = async () => {
    const nome = $('#cfgUnitNome').value.trim();
    if (!nome) return toast('Informe o nome da unidade', true);
    try {
      await FCApi.configAdd('config_unidades', { nome, templateSheet: $('#cfgUnitTpl').value });
      $('#cfgUnitNome').value = '';
      toast('Unidade incluída');
      await refreshMeta();
      await renderConfiguracoes();
    } catch (e) { toast(e.message, true); }
  };

  $('#cfgRubAdd').onclick = async () => {
    const sheet = $('#cfgRubSheet').value;
    const nome = $('#cfgRubNome').value.trim();
    const secao = $('#cfgRubSecao').value.trim();
    if (!nome) return toast('Informe o nome da rubrica', true);
    try {
      await FCApi.configAdd('config_rubricas', { sheet, nome, secao: secao || undefined });
      $('#cfgRubNome').value = '';
      $('#cfgRubSecao').value = '';
      toast('Rubrica incluída');
      await paintRubricas();
    } catch (e) { toast(e.message, true); }
  };

  $('#cfgRubSheet').onchange = paintRubricas;

  $('#cfgShowExcel').onclick = async () => {
    try {
      await FCApi.showExcelBase();
      toast('Planilha Excel exibida');
    } catch (e) { toast(e.message, true); }
  };

  await paintResponsaveis();
  await paintUnidades();
  await paintRubricas();
  scheduleTscrollSync($('#content'));
}

async function renderResumo() {
  const data = await FCApi.resumo();
  const prevReal = data.prevReal || [];
  let resumoFilter = '';

  const cards = (data.cards || []).map(c => `
    <div class="stat-card glass">
      <div class="stat-label">${esc(c.titulo)}</div>
      <div class="stat-value">${c.valorCents != null ? brl(c.valorCents) : esc(c.valor)}</div>
    </div>`).join('');

  const alerts = (data.alerts || []).map(a => `<li>${esc(a.sheet)} — ${esc(a.cell)}</li>`).join('');

  $('#content').innerHTML = `
    <div class="stats">${cards}</div>
    <div class="table-card glass resumo-panel">
      <div class="table-toolbar">
        <div class="table-title">Previsto x Realizado por unidade <span class="count" id="resumoCount">${prevReal.length}</span></div>
        <div class="tsearch resumo-search"><input id="resumoSearch" type="search" placeholder="Buscar unidade…" autocomplete="off"></div>
      </div>
      <div class="tscroll resumo-scroll"><table data-col-resize="resumo">
        <colgroup><col class="col-grow" style="width:46%"><col class="col-money" style="width:27%"><col class="col-money" style="width:27%"></colgroup>
        <thead><tr><th class="col-grow">Unidade (aba)</th><th class="col-money">Previsto</th><th class="col-money">Realizado</th></tr></thead>
        <tbody id="resumoTbody"></tbody></table></div>
    </div>
    ${data.alertCount ? `<div class="alert-box glass"><strong>${data.alertCount} alerta(s) #REF!</strong><ul>${alerts}</ul></div>` : ''}`;

  function paintResumoTable() {
    const q = resumoFilter.trim().toLowerCase();
    const list = q
      ? prevReal.filter(p => {
          const hay = `${p.label || ''} ${p.unidade || ''}`.toLowerCase();
          return hay.includes(q);
        })
      : prevReal;
    const countEl = $('#resumoCount');
    if (countEl) {
      countEl.textContent = q && list.length !== prevReal.length
        ? `${list.length} / ${prevReal.length}`
        : String(prevReal.length);
    }
    const tbody = $('#resumoTbody');
    if (!tbody) return;
    if (!list.length) {
      tbody.innerHTML = `<tr><td colspan="3"><div class="empty-state">${q ? 'Nenhuma unidade corresponde à busca' : 'Sem dados'}</div></td></tr>`;
      scheduleTscrollSync($('#content'));
      return;
    }
    tbody.innerHTML = list.map(p => `<tr>
      <td class="col-grow">${esc(p.label || p.unidade)}</td><td class="col-money">${brl(p.previstoCents)}</td><td class="col-money">${brl(p.realizadoCents)}</td></tr>`).join('');
    scheduleTscrollSync($('#content'));
  }

  $('#resumoSearch').oninput = e => {
    resumoFilter = e.target.value;
    paintResumoTable();
  };

  paintResumoTable();
}

window.addEventListener('resize', () => scheduleTscrollSync());

init();
