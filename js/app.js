import * as db from './db.js';
import { parseWorkbook, parseCSV, parsePDF, pdfToLines, assignIds } from './parsers.js';
import { CATEGORIES, CAT, categorize, merchantKey } from './categories.js';

const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const money = (n, sign = false) => {
  const s = Math.abs(n).toLocaleString('en-CA', { style: 'currency', currency: 'CAD' });
  return (n < 0 ? '−' : sign && n > 0 ? '+' : '') + s;
};
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const monthName = (ym) => `${MESES[Number(ym.slice(5, 7)) - 1]} ${ym.slice(0, 4)}`;
const dayName = (iso) => new Date(iso + 'T12:00:00').toLocaleDateString('es', { weekday: 'long', day: 'numeric', month: 'long' });

const state = {
  txns: [],
  accounts: [],
  rules: {},
  month: null,
  view: 'resumen',
  filterCat: '',
  filterAcct: '',
  search: '',
  lastResults: [],
};

// ---------- carga ----------

async function load() {
  [state.txns, state.accounts, state.rules] = await Promise.all([db.getAll('txns'), db.getAll('accounts'), db.kvGet('rules')]);
  state.rules = state.rules || {};
  const months = monthsWithData();
  if (!state.month || !months.includes(state.month)) state.month = months[months.length - 1] || new Date().toISOString().slice(0, 7);
}

function monthsWithData() {
  return [...new Set(state.txns.map((t) => t.date.slice(0, 7)))].sort();
}

const acctName = (id) => state.accounts.find((a) => a.id === id)?.name || id;
const shortAcct = (id) => {
  const a = state.accounts.find((x) => x.id === id);
  if (!a) return id;
  return a.name.replace(/Scotia(bank)? |American Express |Mastercard |Line of Credit |Visa Infinite /g, '').replace(/\s+/g, ' ').trim();
};

// ---------- importar ----------

async function readFile(file) {
  const name = file.name;
  const lower = name.toLowerCase();
  if (lower.endsWith('.csv')) return parseCSV(await file.text(), name);
  if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
    const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array' });
    const sheets = wb.SheetNames.map((n) => ({ name: n, rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) }));
    return parseWorkbook(sheets, name);
  }
  if (lower.endsWith('.pdf')) {
    const pdfjs = await import('../vendor/pdf.min.mjs');
    pdfjs.GlobalWorkerOptions.workerSrc = new URL('../vendor/pdf.worker.min.mjs', import.meta.url).href;
    const lines = await pdfToLines(pdfjs, new Uint8Array(await file.arrayBuffer()));
    return parsePDF(lines, name);
  }
  throw new Error('Formato no soportado. Usa CSV, Excel o PDF.');
}

async function importFiles(files) {
  const results = [];
  toast('Leyendo archivos…', 0);
  for (const file of files) {
    try {
      const r = assignIds(await readFile(file));
      const existing = new Set(state.txns.map((t) => t.id));
      let acct = state.accounts.find((a) => a.id === r.account.id) || { ...r.account };
      const isCapOne = r.account.id.startsWith('capone-');
      const isStatement = Boolean(r.meta.periodEnd);
      const source = isStatement ? 'statement' : isCapOne ? 'activity' : 'export';

      let txns = r.txns;
      // Capital One: el estado mensual manda sobre la lista de "movimientos recientes".
      if (isCapOne && isStatement) {
        const from = shiftDate(r.meta.periodStart, -3);
        const stale = state.txns.filter((t) => t.account === acct.id && t.source === 'activity' && t.date >= from && t.date <= r.meta.periodEnd);
        if (stale.length) await db.delMany('txns', stale.map((t) => t.id));
        state.txns = state.txns.filter((t) => !stale.includes(t));
        for (const t of stale) existing.delete(t.id);
      }
      if (isCapOne && !isStatement && acct.statementEnd) txns = txns.filter((t) => t.date > acct.statementEnd);

      const fresh = txns.filter((t) => !existing.has(t.id)).map((t) => ({ ...t, source, cat: categorize(t, state.rules) }));
      await db.putMany('txns', fresh);
      state.txns.push(...fresh);

      const m = r.meta;
      if (m.cutDay) acct.cutDay = m.cutDay;
      if (m.dueDate) acct.dueDate = m.dueDate;
      if (m.statementBalance != null) acct.statementBalance = m.statementBalance;
      if (m.minPayment != null) acct.minPayment = m.minPayment;
      if (m.periodEnd && (!acct.statementEnd || m.periodEnd > acct.statementEnd)) acct.statementEnd = m.periodEnd;
      if (m.balance != null) acct.balance = m.balance;
      acct.name = acct.name || r.account.name;
      acct.lastImport = new Date().toISOString();
      await db.putMany('accounts', [acct]);
      state.accounts = state.accounts.filter((a) => a.id !== acct.id).concat(acct);

      results.push({ ok: true, file: file.name, account: acct.name, added: fresh.length, dup: txns.length - fresh.length, pending: m.skippedPending || 0 });
    } catch (e) {
      console.error(e);
      results.push({ ok: false, file: file.name, error: e.message || String(e) });
    }
  }
  state.lastResults = results;
  const added = results.reduce((a, r) => a + (r.added || 0), 0);
  const months = monthsWithData();
  if (added) state.month = months[months.length - 1];
  hideToast();
  render();
  toast(added ? `Listo: ${added} movimientos nuevos` : 'No había movimientos nuevos');
}

function shiftDate(iso, days) {
  const d = new Date(iso + 'T12:00:00');
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}

// ---------- cálculos ----------

function monthTxns() {
  return state.txns.filter((t) => t.date.startsWith(state.month));
}

function summary(txns) {
  const byCat = {};
  let gasto = 0, ingreso = 0;
  for (const t of txns) {
    const c = CAT[t.cat] || CAT.otros;
    if (c.kind === 'gasto') { gasto += t.amount; byCat[c.id] = (byCat[c.id] || 0) + t.amount; }
    else if (c.kind === 'ingreso') ingreso -= t.amount;
  }
  const cats = Object.entries(byCat).filter(([, v]) => v > 0.004).sort((a, b) => b[1] - a[1]);
  return { gasto, ingreso, cats, intereses: byCat.intereses || 0 };
}

// ---------- vistas ----------

function render() {
  const months = monthsWithData();
  $('#monthLabel').textContent = state.view === 'cuentas' ? 'Cuentas' : monthName(state.month);
  $('#prevMonth').disabled = state.view === 'cuentas' || !months.length || state.month <= months[0];
  $('#nextMonth').disabled = state.view === 'cuentas' || !months.length || state.month >= months[months.length - 1];
  $('#prevMonth').style.visibility = $('#nextMonth').style.visibility = state.view === 'cuentas' ? 'hidden' : 'visible';
  document.querySelectorAll('.tabs button').forEach((b) => b.classList.toggle('active', b.dataset.view === state.view));
  document.querySelectorAll('.view').forEach((v) => v.classList.toggle('active', v.id === `view-${state.view}`));
  if (state.view === 'resumen') renderResumen();
  if (state.view === 'movs') renderMovs();
  if (state.view === 'cuentas') renderCuentas();
}

function emptyState() {
  return `<div class="empty"><div class="big">📥</div><p>Todavía no hay movimientos.</p>
    <p class="small">Baja el estado de cuenta de tu banco (Excel, CSV o PDF) y súbelo aquí.</p>
    <button class="btn" data-action="import">Importar estados de cuenta</button></div>`;
}

function renderResumen() {
  const el = $('#view-resumen');
  if (!state.txns.length) { el.innerHTML = emptyState(); return; }
  const s = summary(monthTxns());
  const net = s.ingreso - s.gasto;
  const accts = [...new Set(monthTxns().map((t) => t.account))].map(shortAcct);
  el.innerHTML = `
    <div class="kpis">
      <div class="kpi big"><div class="label">Gastaste en ${esc(MESES[Number(state.month.slice(5)) - 1])}</div><div class="value">${money(s.gasto)}</div></div>
      <div class="kpi"><div class="label">Ingresos</div><div class="value pos">${money(s.ingreso)}</div></div>
      <div class="kpi"><div class="label">Diferencia</div><div class="value ${net >= 0 ? 'pos' : 'neg'}">${money(net, true)}</div></div>
    </div>
    ${s.intereses > 0.5 ? `<div class="alert" data-cat="intereses">⚠️ Este mes pagaste <strong>${money(s.intereses)}</strong> en intereses y cargos del banco. Es dinero que se va sin comprar nada.</div>` : ''}
    <div class="card">
      <h2>¿A dónde se fue el dinero?</h2>
      <div class="donut-wrap">${donut(s.cats, s.gasto)}</div>
      ${s.cats.map(([id, v]) => {
        const c = CAT[id];
        const pct = s.gasto ? (v / s.gasto) * 100 : 0;
        return `<div class="catrow" data-cat="${id}">
          <div class="ic">${c.icon}</div>
          <div class="nm">${esc(c.name)} <span class="pct">${pct.toFixed(0)}%</span></div>
          <div class="amt">${money(v)}</div>
          <div class="bar"><i style="width:${Math.max(pct, 1.5)}%;background:${c.color}"></i></div>
        </div>`;
      }).join('') || '<p class="muted">Sin gastos este mes.</p>'}
    </div>
    <p class="muted small">Incluye: ${esc(accts.join(', ')) || '—'}. Los pagos a tus tarjetas, transferencias e inversiones no cuentan como gasto.</p>`;
}

function donut(cats, total) {
  const R = 80, r = 52, C = 100;
  if (!total) return '';
  let a0 = -Math.PI / 2;
  const paths = cats.map(([id, v]) => {
    const frac = v / total;
    const a1 = a0 + frac * Math.PI * 2 - (cats.length > 1 ? 0.012 : 0);
    const large = a1 - a0 > Math.PI ? 1 : 0;
    const p = (rad, ang) => `${C + rad * Math.cos(ang)} ${C + rad * Math.sin(ang)}`;
    const d = frac >= 0.9999
      ? `M ${C} ${C - R} A ${R} ${R} 0 1 1 ${C - 0.01} ${C - R} L ${C - 0.01} ${C - r} A ${r} ${r} 0 1 0 ${C} ${C - r} Z`
      : `M ${p(R, a0)} A ${R} ${R} 0 ${large} 1 ${p(R, a1)} L ${p(r, a1)} A ${r} ${r} 0 ${large} 0 ${p(r, a0)} Z`;
    a0 += frac * Math.PI * 2;
    return `<path d="${d}" fill="${CAT[id].color}"><title>${esc(CAT[id].name)}</title></path>`;
  });
  return `<svg viewBox="0 0 200 200" role="img" aria-label="Gastos por categoría">${paths.join('')}
    <text x="100" y="96" text-anchor="middle" font-size="12" fill="currentColor" opacity=".6">Total</text>
    <text x="100" y="116" text-anchor="middle" font-size="17" font-weight="700" fill="currentColor">${esc(money(total).replace('.00', ''))}</text></svg>`;
}

function renderMovs() {
  const el = $('#view-movs');
  if (!state.txns.length) { el.innerHTML = emptyState(); return; }
  const q = state.search.trim().toLowerCase();
  const list = monthTxns()
    .filter((t) => !state.filterCat || t.cat === state.filterCat)
    .filter((t) => !state.filterAcct || t.account === state.filterAcct)
    .filter((t) => !q || t.desc.toLowerCase().includes(q))
    .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : b.amount - a.amount));
  const accts = [...new Set(monthTxns().map((t) => t.account))];
  const total = list.reduce((a, t) => a + ((CAT[t.cat]?.kind === 'gasto') ? t.amount : 0), 0);
  let html = `<div class="filters">
      <input type="search" id="search" placeholder="Buscar comercio" value="${esc(state.search)}">
      <select id="fAcct"><option value="">Todas las cuentas</option>${accts.map((a) => `<option value="${a}" ${a === state.filterAcct ? 'selected' : ''}>${esc(shortAcct(a))}</option>`).join('')}</select>
    </div>`;
  if (state.filterCat) html += `<p><button class="chip" data-action="clearcat">${CAT[state.filterCat].icon} ${esc(CAT[state.filterCat].name)} · ${money(total)} ✕</button></p>`;
  let day = '';
  for (const t of list) {
    if (t.date !== day) { day = t.date; html += `<div class="day">${esc(dayName(day))}</div>`; }
    const c = CAT[t.cat] || CAT.otros;
    const cls = c.kind === 'ingreso' ? 'pos' : '';
    html += `<div class="tx ${c.kind === 'mov' ? 'mov' : ''}" data-id="${t.id}">
      <div class="ic">${c.icon}</div>
      <div class="d"><div class="t">${esc(t.desc)}</div><div class="s">${esc(c.name)} · ${esc(shortAcct(t.account))}</div></div>
      <div class="a ${cls}">${money(-t.amount, true)}</div>
    </div>`;
  }
  if (!list.length) html += '<p class="empty">No hay movimientos con ese filtro.</p>';
  el.innerHTML = html;
}

function renderCuentas() {
  const el = $('#view-cuentas');
  const res = state.lastResults.length ? `<div class="card"><h2>Última importación</h2>${state.lastResults.map((r) => r.ok
    ? `<div class="result">✅ <b>${esc(r.account)}</b>: ${r.added} nuevos${r.dup ? `, ${r.dup} ya estaban` : ''}${r.pending ? `, ${r.pending} pendientes ignorados` : ''}</div>`
    : `<div class="result err">❌ ${esc(r.file)}: ${esc(r.error)}</div>`).join('')}</div>` : '';
  const accts = [...state.accounts].sort((a, b) => a.name.localeCompare(b.name));
  el.innerHTML = `
    <div class="card">
      <button class="btn" data-action="import">Importar estados de cuenta</button>
      <p class="muted small">Acepta Excel de Scotiabank y Amex, CSV de Canadian Tire y PDF de Capital One. Puedes subir varios a la vez, y si subes uno repetido no se duplica nada.</p>
    </div>
    ${res}
    <div class="card"><h2>Tus cuentas</h2>
      ${accts.map((a) => {
        const ts = state.txns.filter((t) => t.account === a.id);
        const dates = ts.map((t) => t.date).sort();
        const extra = [a.cutDay ? `corte día ${a.cutDay}` : '', a.dueDate ? `pago ${new Date(a.dueDate + 'T12:00').toLocaleDateString('es', { day: 'numeric', month: 'short' })}` : ''].filter(Boolean).join(' · ');
        return `<div class="acct"><div class="n">${esc(a.name)}</div><div class="r">${ts.length} movs.</div>
          <div class="muted small">${dates.length ? `${dates[0]} a ${dates[dates.length - 1]}` : ''}</div><div class="muted small r">${esc(extra)}</div></div>`;
      }).join('') || '<p class="muted">Aún no hay cuentas. Importa un estado de cuenta para empezar.</p>'}
    </div>
    <div class="card"><h2>Privacidad y respaldo</h2>
      <p class="small muted">Tus datos se guardan solo en este teléfono. No se envían a ningún servidor. Mientras llega el respaldo automático en Google Drive, puedes guardar una copia a mano.</p>
      <button class="btn secondary" data-action="backup">Guardar copia de respaldo</button>
      <button class="btn secondary" data-action="restore">Restaurar copia</button>
      <button class="btn danger" data-action="wipe">Borrar todos los datos</button>
    </div>`;
}

// ---------- hoja para cambiar categoría ----------

function openTxn(id) {
  const t = state.txns.find((x) => x.id === id);
  if (!t) return;
  const key = merchantKey(t.desc);
  const same = state.txns.filter((x) => merchantKey(x.desc) === key).length;
  $('#sheetPanel').innerHTML = `
    <h3>${esc(t.desc)}</h3>
    <div class="muted small">${esc(dayName(t.date))} · ${esc(acctName(t.account))} · <b style="white-space:nowrap">${money(-t.amount, true)}</b></div>
    <div class="catgrid">${CATEGORIES.map((c) => `<button data-setcat="${c.id}" class="${c.id === t.cat ? 'sel' : ''}"><span>${c.icon}</span>${esc(c.name)}</button>`).join('')}</div>
    <label class="check"><input type="checkbox" id="remember" checked> Usar esta categoría siempre para «${esc(key)}»${same > 1 ? ` (${same} movimientos)` : ''}</label>
    <button class="btn secondary" data-close>Cerrar</button>`;
  $('#sheetPanel').dataset.id = id;
  $('#sheet').classList.remove('hidden');
}

async function setCategory(id, cat) {
  const t = state.txns.find((x) => x.id === id);
  const remember = $('#remember')?.checked;
  const changed = [];
  if (remember) {
    const key = merchantKey(t.desc);
    state.rules[key] = cat;
    await db.kvSet('rules', state.rules);
    for (const x of state.txns) if (merchantKey(x.desc) === key && x.cat !== cat) { x.cat = cat; changed.push(x); }
  }
  if (t.cat !== cat) { t.cat = cat; if (!changed.includes(t)) changed.push(t); }
  await db.putMany('txns', changed);
  closeSheet();
  render();
  toast(changed.length > 1 ? `${changed.length} movimientos cambiados a ${CAT[cat].name}` : `Cambiado a ${CAT[cat].name}`);
}

function closeSheet() { $('#sheet').classList.add('hidden'); }

// ---------- respaldo manual ----------

async function backup() {
  const data = { app: 'finanzas', version: 1, exportedAt: new Date().toISOString(), txns: state.txns, accounts: state.accounts, rules: state.rules };
  const blob = new Blob([JSON.stringify(data)], { type: 'application/json' });
  const name = `finanzas-respaldo-${new Date().toISOString().slice(0, 10)}.json`;
  const file = new File([blob], name, { type: 'application/json' });
  if (navigator.canShare && navigator.canShare({ files: [file] })) {
    try { await navigator.share({ files: [file], title: name }); return; } catch (e) { if (e.name === 'AbortError') return; }
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

async function restore(file) {
  try {
    const data = JSON.parse(await file.text());
    if (data.app !== 'finanzas' || !Array.isArray(data.txns)) throw new Error('Este archivo no es una copia de Finanzas.');
    await db.putMany('txns', data.txns);
    await db.putMany('accounts', data.accounts || []);
    await db.kvSet('rules', { ...(state.rules || {}), ...(data.rules || {}) });
    await load();
    render();
    toast(`Copia restaurada: ${data.txns.length} movimientos`);
  } catch (e) {
    toast(e.message || 'No pude leer la copia');
  }
}

async function wipe() {
  if (!confirm('¿Borrar todos los movimientos, cuentas y reglas de este teléfono? No se puede deshacer.')) return;
  await Promise.all([db.clear('txns'), db.clear('accounts'), db.clear('kv')]);
  state.lastResults = [];
  await load();
  render();
  toast('Datos borrados');
}

// ---------- utilidades de UI ----------

let toastTimer;
function toast(msg, ms = 2600) {
  const el = $('#toast');
  el.textContent = msg;
  el.classList.remove('hidden');
  clearTimeout(toastTimer);
  if (ms) toastTimer = setTimeout(hideToast, ms);
}
function hideToast() { $('#toast').classList.add('hidden'); }

function goMonth(delta) {
  const months = monthsWithData();
  const i = months.indexOf(state.month) + delta;
  if (i >= 0 && i < months.length) { state.month = months[i]; render(); window.scrollTo(0, 0); }
}

// ---------- eventos ----------

document.addEventListener('click', (e) => {
  const tab = e.target.closest('.tabs button');
  if (tab) { state.view = tab.dataset.view; render(); window.scrollTo(0, 0); return; }
  if (e.target.closest('[data-close]')) { closeSheet(); return; }
  const setc = e.target.closest('[data-setcat]');
  if (setc) { setCategory($('#sheetPanel').dataset.id, setc.dataset.setcat); return; }
  const act = e.target.closest('[data-action]')?.dataset.action;
  if (act === 'import') $('#fileInput').click();
  if (act === 'backup') backup();
  if (act === 'restore') $('#restoreInput').click();
  if (act === 'wipe') wipe();
  if (act === 'clearcat') { state.filterCat = ''; render(); }
  if (act) return;
  const cat = e.target.closest('[data-cat]');
  if (cat) { state.filterCat = cat.dataset.cat; state.view = 'movs'; render(); window.scrollTo(0, 0); return; }
  const tx = e.target.closest('.tx');
  if (tx) openTxn(tx.dataset.id);
});

document.addEventListener('input', (e) => {
  if (e.target.id === 'search') {
    state.search = e.target.value;
    const pos = e.target.selectionStart;
    renderMovs();
    const s = $('#search'); s.focus(); s.setSelectionRange(pos, pos);
  }
});
document.addEventListener('change', (e) => {
  if (e.target.id === 'fAcct') { state.filterAcct = e.target.value; renderMovs(); }
});
$('#fileInput').addEventListener('change', async (e) => {
  const files = [...e.target.files];
  e.target.value = '';
  if (files.length) { state.view = 'cuentas'; await importFiles(files); }
});
$('#restoreInput').addEventListener('change', (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) restore(f); });
$('#prevMonth').addEventListener('click', () => goMonth(-1));
$('#nextMonth').addEventListener('click', () => goMonth(1));

// ---------- inicio ----------

(async () => {
  if (navigator.storage?.persist) navigator.storage.persist().catch(() => {});
  await load();
  render();
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
})();
