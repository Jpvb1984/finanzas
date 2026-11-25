import * as db from './db.js';
import { parseWorkbook, parseCSV, parsePDF, pdfToLines, assignIds } from './parsers.js';
import { CATEGORIES, CAT, categorize, merchantKey } from './categories.js';
import { cardInfo, fixedExpenses, monthlyPicture, simulate, minimumsOnly, buildICS, isDebtAccount, todayISO, daysBetween, allocateLump, cashProjection, nextOccurrence, reviewSubscriptions } from './debts.js';
import * as gdrive from './gdrive.js';

const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const money = (n, sign = false) => {
  const s = Math.abs(n).toLocaleString('en-CA', { style: 'currency', currency: 'CAD' });
  return (n < 0 ? '−' : sign && n > 0 ? '+' : '') + s;
};
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const monthName = (ym) => { const m = MESES[Number(ym.slice(5, 7)) - 1]; return `${m[0].toUpperCase()}${m.slice(1)} ${ym.slice(0, 4)}`; };
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
  fixedOv: {},
  settings: {},
  debtHist: {},
  lump: 0,
};

// ---------- carga ----------

async function load() {
  [state.txns, state.accounts, state.rules] = await Promise.all([db.getAll('txns'), db.getAll('accounts'), db.kvGet('rules')]);
  state.rules = state.rules || {};
  state.fixedOv = (await db.kvGet('fixed')) || {};
  state.settings = (await db.kvGet('settings')) || {};
  state.debtHist = (await db.kvGet('debtHist')) || {};
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
      if (m.balance != null) { acct.balance = m.balance; acct.balanceAt = new Date().toISOString(); }
      if (m.limit != null) acct.limit = m.limit;
      if (m.limitApprox != null) acct.limitApprox = m.limitApprox;
      if (m.apr != null) acct.apr = m.apr;
      if (/line of credit|scotialine/i.test(acct.name)) acct.kind = 'linea';
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
  await snapshotDebt();
  if (added) scheduleBackup();
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

const TITLES = { cuentas: 'Cuentas', deudas: 'Deudas y pagos', asesoria: 'Asesoría' };
function render() {
  const months = monthsWithData();
  const noMonth = Boolean(TITLES[state.view]);
  $('#monthLabel').textContent = noMonth ? TITLES[state.view] : monthName(state.month);
  $('#prevMonth').disabled = noMonth || !months.length || state.month <= months[0];
  $('#nextMonth').disabled = noMonth || !months.length || state.month >= months[months.length - 1];
  $('#prevMonth').style.visibility = $('#nextMonth').style.visibility = noMonth ? 'hidden' : 'visible';
  document.querySelectorAll('.tabs button').forEach((b) => b.classList.toggle('active', b.dataset.view === state.view));
  document.querySelectorAll('.view').forEach((v) => v.classList.toggle('active', v.id === `view-${state.view}`));
  if (state.view === 'resumen') renderResumen();
  if (state.view === 'movs') renderMovs();
  if (state.view === 'cuentas') renderCuentas();
  if (state.view === 'deudas') renderDeudas();
  if (state.view === 'asesoria') renderAsesoria();
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
  const isLatest = state.month === latestMonth();
  const budgets = state.settings.budgets || {};
  const trend = categoryTrends();
  el.innerHTML = `${isLatest ? backupBanner() + staleBanner() + weeklyBanner() + upcomingBanner() + budgetBanner() : ''}
    <div class="kpis">
      <div class="kpi big"><div class="label">Gastaste en ${esc(MESES[Number(state.month.slice(5)) - 1])}</div><div class="value">${money(s.gasto)}</div></div>
      <div class="kpi"><div class="label">Ingresos</div><div class="value pos">${money(s.ingreso)}</div></div>
      <div class="kpi"><div class="label">Diferencia</div><div class="value ${net >= 0 ? 'pos' : 'neg'}">${money(net, true)}</div></div>
    </div>
    ${s.intereses > 0.5 ? `<div class="alert" data-cat="intereses">⚠️ Este mes pagaste <strong>${money(s.intereses)}</strong> en intereses y cargos del banco. Es dinero que se va sin comprar nada.</div>` : ''}
    <div class="card">
      <div class="dhead"><h2>¿A dónde se fue el dinero?</h2><button class="chip" data-action="budgets">🎯 Presupuesto</button></div>
      <div class="donut-wrap">${donut(s.cats, s.gasto)}</div>
      ${s.cats.map(([id, v]) => {
        const c = CAT[id];
        const pct = s.gasto ? (v / s.gasto) * 100 : 0;
        const b = budgets[id];
        const bp = b ? v / b : null;
        const tr = trend[id];
        return `<div class="catrow" data-cat="${id}">
          <div class="ic">${c.icon}</div>
          <div class="nm">${esc(c.name)} <span class="pct">${pct.toFixed(0)}%</span>${tr ? ` <span class="trend ${tr > 0 ? 'up' : 'down'}">${tr > 0 ? '▲' : '▼'} ${Math.abs(Math.round(tr * 100))}%</span>` : ''}</div>
          <div class="amt">${money(v)}${b ? `<div class="small muted">de ${money(b)}</div>` : ''}</div>
          <div class="bar">${b ? `<i style="width:${Math.min(100, bp * 100)}%;background:${bp >= 1 ? 'var(--danger)' : bp >= 0.8 ? '#f59f00' : '#37b24d'}"></i>` : `<i style="width:${Math.max(pct, 1.5)}%;background:${c.color}"></i>`}</div>
        </div>`;
      }).join('') || '<p class="muted">Sin gastos este mes.</p>'}
    </div>
    ${fixedCard()}
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
    ${updateCard()}
    ${driveCard()}
    <div class="card"><h2>Privacidad y copia manual</h2>
      <p class="small muted">Tus datos viven en este teléfono y, si lo activas, en tu propio Google Drive. No pasan por ningún otro servidor. También puedes guardar una copia como archivo.</p>
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
    ${t.amount > 0 ? `<label class="check"><input type="checkbox" id="oneOff" ${t.oneOff ? 'checked' : ''}> Es un gasto de una sola vez (no cuenta en el mes típico)</label>` : ''}
    ${t.amount > 0 ? `<label class="check"><input type="checkbox" id="isFixed" data-key="${esc(key)}" ${isFixedKey(key) ? 'checked' : ''}> Es un gasto fijo (se repite cada mes)</label>` : ''}
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
  scheduleBackup();
  toast(changed.length > 1 ? `${changed.length} movimientos cambiados a ${CAT[cat].name}` : `Cambiado a ${CAT[cat].name}`);
}

function closeSheet() { $('#sheet').classList.add('hidden'); }

// ---------- respaldo manual ----------

async function backup() {
  const data = { app: 'finanzas', version: 1, exportedAt: new Date().toISOString(), txns: state.txns, accounts: state.accounts, rules: state.rules, fixed: state.fixedOv, settings: state.settings };
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
    await restoreData(JSON.parse(await file.text()));
  } catch (e) {
    toast(e.message || 'No pude leer la copia');
  }
}

async function restoreData(data) {
  {
    if (data.app !== 'finanzas' || !Array.isArray(data.txns)) throw new Error('Este archivo no es una copia de Finanzas.');
    await db.putMany('txns', data.txns);
    await db.putMany('accounts', data.accounts || []);
    await db.kvSet('rules', { ...(state.rules || {}), ...(data.rules || {}) });
    if (data.fixed) await db.kvSet('fixed', { ...state.fixedOv, ...data.fixed });
    if (data.settings) await db.kvSet('settings', { ...data.settings, ...state.settings, goals: state.settings.goals?.length ? state.settings.goals : data.settings.goals, budgets: Object.keys(state.settings.budgets || {}).length ? state.settings.budgets : data.settings.budgets });
    if (data.debtHist) await db.kvSet('debtHist', { ...data.debtHist, ...state.debtHist });
    await load();
    render();
    toast(`Copia restaurada: ${data.txns.length} movimientos`);
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

// ---------- deudas ----------

const fmtDate = (iso, opts = { day: 'numeric', month: 'short' }) => new Date(iso + 'T12:00:00').toLocaleDateString('es', opts);
const inDays = (n) => (n === 0 ? 'hoy' : n === 1 ? 'mañana' : `en ${n} días`);
const est = (src) => (src && !['estado', 'archivo', 'tuyo', 'estado + movimientos'].includes(src) ? `<span class="est" title="${esc(src)}">*</span>` : '');
const pct = (n) => `${(Math.round(n * 100) / 100).toLocaleString('es')}%`;

function isFixedKey(key) {
  return fixedList().some((f) => (f.keys || [f.key]).includes(key));
}
function fixedList() { return fixedExpenses(state.txns, state.fixedOv); }
function allCards() {
  const today = todayISO();
  return state.accounts.filter(isDebtAccount).map((a) => cardInfo(a, state.txns, today));
}

function upcomingBanner() {
  const soon = allCards().filter((c) => c.daysToDue != null && c.daysToDue <= 10 && (c.balance ?? 1) > 0).sort((a, b) => a.daysToDue - b.daysToDue);
  if (!soon.length) return '';
  return `<div class="alert warn" data-action="goto" data-view="deudas"><b>📅 Pagos próximos</b>
    ${soon.map((c) => `<div class="small">${esc(shortCard(c.name))}: <b>${inDays(c.daysToDue)}</b> (${fmtDate(c.dueDate)})${c.minPayment ? ` · mínimo ${money(c.minPayment)}${est(c.minSrc)}` : ''}</div>`).join('')}
  </div>`;
}
const shortCard = (n) => String(n).replace(/Scotia(bank)? |American Express |Visa Infinite |Line of Credit /g, '').replace(/\s+/g, ' ').trim();

function fixedCard() {
  const list = fixedList();
  if (!list.length) return '';
  const total = list.reduce((s, f) => s + f.perMonth, 0);
  const rows = list.map((f) => {
    const paid = f.dates.filter((d) => d.startsWith(state.month)).length;
    const status = paid ? `<span class="pos">✓ pagado${paid > 1 ? ` ×${paid}` : ''}</span>` : `<span class="muted">pendiente · ~día ${f.day}</span>`;
    return `<div class="frow"><div class="ic">${CAT[f.cat]?.icon || '📌'}</div><div><div>${esc(f.name)}</div><div class="small">${status} · ${esc(f.freq)}</div></div><div class="amt">${money(f.perMonth)}<div class="small muted">al mes</div></div></div>`;
  }).join('');
  return `<div class="card"><h2>Gastos fijos · ${money(total)} al mes</h2>${rows}
    <p class="small muted">Se detectan solos (renta, préstamos, seguros, teléfono, suscripciones). Toca cualquier movimiento para marcarlo o desmarcarlo como fijo.</p></div>`;
}

const STATUS = {
  usar: ['🟢', 'Úsala', 'ok'],
  cuidado: ['🟡', 'Con cuidado', 'mid'],
  quieta: ['🔴', 'Déjala quieta', 'bad'],
  prestamo: ['🏦', 'Préstamo', 'mid'],
};

function renderDeudas() {
  const el = $('#view-deudas');
  const cards = allCards();
  if (!cards.length) { el.innerHTML = emptyState(); return; }
  const withBal = cards.filter((c) => c.balance > 0);
  const total = withBal.reduce((s, c) => s + c.balance, 0);
  const interest = withBal.reduce((s, c) => s + (c.monthlyInterest || 0), 0);
  const mins = cards.filter((c) => c.minPayment).reduce((s, c) => s + c.minPayment, 0);
  const anyEst = cards.some((c) => ['estimado'].includes(c.balanceSrc) || c.aprSrc === 'típica' || c.dueSrc === 'estimada' || c.balance == null || c.apr == null);
  const usable = cards.filter((c) => c.status === 'usar').sort((a, b) => (b.graceIfBuyToday || 0) - (a.graceIfBuyToday || 0));
  const upcoming = cards.filter((c) => c.dueDate && c.kind !== 'prestamo').sort((a, b) => a.dueDate.localeCompare(b.dueDate));

  const loanFixed = fixedList().filter((f) => f.cat === 'prestamos');
  const hasLoan = cards.some((c) => c.kind === 'prestamo');

  el.innerHTML = `
    <div class="kpis">
      <div class="kpi big"><div class="label">Debes en total</div><div class="value neg">${money(total)}</div></div>
      <div class="kpi"><div class="label">Intereses al mes</div><div class="value neg">~${money(interest)}</div></div>
      <div class="kpi"><div class="label">Mínimos del mes</div><div class="value">${money(mins)}</div></div>
    </div>
    ${debtChart()}
    <div class="card hl">
      <h2>¿Con qué pago hoy?</h2>
      ${usable.length
        ? `<p>Usa <b>${esc(shortCard(usable[0].name))}</b>: si compras hoy, tienes ${usable[0].graceIfBuyToday ? `<b>${usable[0].graceIfBuyToday} días</b>` : 'tiempo'} para pagarlo sin intereses.</p>`
        : `<p><b>Usa tu tarjeta de débito.</b> Todas tus tarjetas de crédito tienen saldo con intereses o están muy usadas. Lo que compres con ellas te cuesta más.</p>`}
    </div>
    <div class="card"><h2>Próximos pagos</h2>
      ${upcoming.map((c) => `<div class="due ${c.daysToDue <= 5 ? 'urgent' : ''}">
          <div class="when"><b>${fmtDate(c.dueDate)}</b>${est(c.dueSrc)}<div class="small">${inDays(c.daysToDue)}</div></div>
          <div><div class="n">${esc(shortCard(c.name))}</div>
            <div class="small">Mínimo <b>${c.minPayment != null ? money(c.minPayment) : '?'}</b>${est(c.minSrc)}${c.statementBalance ? ` · sin intereses: ${money(c.statementBalance)}` : ''}</div></div>
        </div>`).join('') || '<p class="muted">Sin fechas todavía. Edita cada tarjeta para poner su fecha de pago.</p>'}
      <button class="btn secondary" data-action="ics" style="margin-top:12px">📅 Agregar avisos al Calendario del iPhone</button>
      <p class="small muted">Te avisa 3 días antes y 1 día antes de cada pago, aunque no abras la app.</p>
    </div>
    ${cards.map(cardBlock).join('')}
    ${!hasLoan && loanFixed.length ? `<div class="card"><h2>¿Y tu préstamo?</h2><p class="small">Veo pagos de préstamo de ${money(loanFixed[0].amount)} ${esc(loanFixed[0].freq)}. Si me dices cuánto debes y a qué tasa, lo incluyo en el plan.</p>
      <button class="btn secondary" data-action="addloan" data-pay="${loanFixed[0].perMonth.toFixed(2)}">Agregar préstamo</button></div>` : ''}
    <button class="btn secondary" data-action="addloan" data-pay="">+ Agregar otra deuda</button>
    ${anyEst ? '<p class="small muted"><span class="est">*</span> Dato estimado o típico: tócale "Editar datos" y escribe el real que ves en tu banco.</p>' : ''}`;
}

function cardBlock(c) {
  const [ic, label, cls] = STATUS[c.status] || STATUS.cuidado;
  const util = c.util != null ? Math.min(c.util, 1) : null;
  const rows = [
    ['Saldo', c.balance != null ? money(c.balance) + est(c.balanceSrc) : '<span class="neg">falta</span>'],
    c.kind !== 'prestamo' ? ['Límite', c.limit ? money(c.limit) : '—'] : null,
    ['Tasa anual', c.apr != null ? pct(c.apr) + (c.aprSrc === 'típica' ? '<span class="est">*</span>' : '') : '<span class="neg">falta</span>'],
    ['Interés al mes', c.monthlyInterest ? `~${money(c.monthlyInterest)}` : '—'],
    c.kind !== 'prestamo' ? ['Corte', c.cutDay ? `día ${c.cutDay}${c.cutSrc === 'intereses' ? '<span class="est">*</span>' : ''}` : '—'] : null,
    ['Pago', c.dueDate ? `${fmtDate(c.dueDate)}${est(c.dueSrc)}` : '—'],
  ].filter(Boolean);
  return `<div class="card debt">
    <div class="dhead"><div class="n">${esc(c.name)}</div><span class="badge ${cls}">${ic} ${label}</span></div>
    ${c.why ? `<p class="small">${esc(c.why)}</p>` : ''}
    ${util != null ? `<div class="ubar"><i style="width:${Math.max(util * 100, 2)}%" class="${util >= 0.7 ? 'bad' : util >= 0.3 ? 'mid' : 'ok'}"></i></div><div class="small muted">${Math.round(c.util * 100)}% del límite usado</div>` : ''}
    <div class="grid2">${rows.map(([k, v]) => `<div class="muted small">${k}</div><div class="r">${v}</div>`).join('')}</div>
    <button class="btn secondary sm" data-action="editcard" data-id="${esc(c.id)}">Editar datos</button>
  </div>`;
}

function numField(id, label, val, hint = '', step = '0.01') {
  return `<label class="field"><span>${label}</span><input id="${id}" type="number" inputmode="decimal" step="${step}" value="${val ?? ''}" placeholder="${esc(hint)}"></label>`;
}

function openCardEditor(id, pay) {
  const a = id ? state.accounts.find((x) => x.id === id) : null;
  const c = a ? cardInfo(a, state.txns, todayISO()) : null;
  const u = a?.user || {};
  const isLoan = !a || a.kind === 'prestamo';
  $('#sheetPanel').innerHTML = `
    <h3>${a ? esc(a.name) : 'Nueva deuda'}</h3>
    <p class="small muted">Escribe lo que ves en la app de tu banco. Lo que dejes vacío se calcula de tus archivos.</p>
    ${!a ? `<label class="field"><span>Nombre</span><input id="f-name" type="text" placeholder="Préstamo Scotiabank"></label>` : ''}
    ${numField('f-balance', 'Saldo actual (lo que debes hoy)', u.balance, c?.balance != null ? `ahora: ${c.balance}` : '')}
    ${!isLoan ? numField('f-limit', 'Límite de crédito', u.limit, c?.limit ? `ahora: ${c.limit}` : '') : ''}
    ${numField('f-apr', 'Tasa de interés anual (%)', u.apr, c?.apr != null ? `ahora: ${c.apr}` : 'ej. 20.99')}
    ${isLoan ? numField('f-payment', 'Pago al mes', u.payment ?? (pay || ''), '') : numField('f-min', 'Pago mínimo de este mes', u.minPayment, c?.minPayment != null ? `ahora: ${c.minPayment}` : '')}
    ${!isLoan ? numField('f-stmt', 'Saldo del último estado (para no pagar intereses)', u.statementBalance, c?.statementBalance != null ? `ahora: ${c.statementBalance}` : '') : ''}
    ${!isLoan ? numField('f-cut', 'Día de corte (1 a 31)', u.cutDay, c?.cutDay ? `ahora: ${c.cutDay}` : '', '1') : ''}
    ${numField('f-due', 'Día de pago (1 a 31)', u.dueDay, c?.dueDate ? `ahora: ${Number(c.dueDate.slice(8))}` : '', '1')}
    <button class="btn" data-action="savecard">Guardar</button>
    ${a && a.kind === 'prestamo' ? `<button class="btn danger" data-action="delloan" data-id="${esc(a.id)}">Borrar esta deuda</button>` : ''}
    <button class="btn secondary" data-close>Cancelar</button>`;
  $('#sheetPanel').dataset.card = id || '';
  $('#sheet').classList.remove('hidden');
}

async function saveCard() {
  const id = $('#sheetPanel').dataset.card;
  const v = (sel) => { const el = $(sel); if (!el || el.value === '') return null; const n = Number(el.value); return isFinite(n) ? n : null; };
  let a = id ? state.accounts.find((x) => x.id === id) : null;
  if (!a) {
    const name = ($('#f-name')?.value || '').trim() || 'Préstamo';
    a = { id: 'manual-' + Date.now().toString(36), name, kind: 'prestamo', bank: '', manual: true };
  }
  const prev = a.user || {};
  const u = {};
  const set = (k, val) => { if (val != null) u[k] = val; };
  set('balance', v('#f-balance')); set('limit', v('#f-limit')); set('apr', v('#f-apr'));
  set('minPayment', v('#f-min')); set('payment', v('#f-payment')); set('statementBalance', v('#f-stmt'));
  const cut = v('#f-cut'), due = v('#f-due');
  if (cut >= 1 && cut <= 31) u.cutDay = Math.round(cut);
  if (due >= 1 && due <= 31) u.dueDay = Math.round(due);
  u.balanceAt = u.balance != null && u.balance !== prev.balance ? new Date().toISOString() : prev.balanceAt;
  a.user = u;
  await db.putMany('accounts', [a]);
  state.accounts = state.accounts.filter((x) => x.id !== a.id).concat(a);
  closeSheet();
  render();
  toast('Guardado');
  await snapshotDebt();
  scheduleBackup();
}

async function deleteLoan(id) {
  if (!confirm('¿Borrar esta deuda?')) return;
  await db.delMany('accounts', [id]);
  state.accounts = state.accounts.filter((x) => x.id !== id);
  closeSheet(); render();
}

async function downloadICS() {
  const cards = allCards().filter((c) => c.dueDate);
  if (!cards.length) { toast('Primero pon las fechas de pago de tus tarjetas'); return; }
  const blob = new Blob([buildICS(cards)], { type: 'text/calendar' });
  const url = URL.createObjectURL(blob);
  // En iPhone, abrir el .ics muestra "Añadir todo" al Calendario.
  const a = document.createElement('a');
  a.href = url; a.download = 'pagos-tarjetas.ics';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
  toast('Abre el archivo y toca "Añadir todo"');
}

// ---------- asesoría ----------

function monthLabelFromNow(n) {
  const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() + n);
  return `${MESES[d.getMonth()]} ${d.getFullYear()}`;
}
const yearsMonths = (m) => (m === Infinity ? 'nunca' : m < 12 ? `${m} meses` : `${Math.floor(m / 12)} año${m >= 24 ? 's' : ''}${m % 12 ? ` y ${m % 12} meses` : ''}`);

function adviceModel() {
  const fixed = fixedList();
  const pic = monthlyPicture(state.txns, fixed);
  const cards = allCards();
  const hasLoanAcct = cards.some((c) => c.kind === 'prestamo');
  const fijos = fixed.filter((f) => !(hasLoanAcct && f.cat === 'prestamos')).reduce((s, f) => s + f.perMonth, 0);
  const income = state.settings.income ?? pic.income;
  const debts = cards.filter((c) => c.balance > 0).map((c) => ({
    id: c.id, name: shortCard(c.name), balance: c.balance, apr: c.apr ?? 12, aprMissing: c.apr == null,
    min: c.minPayment ?? Math.max(10, Math.round(c.balance * 0.03)), fixedPayment: c.kind === 'prestamo', status: c.status,
  }));
  const mins = debts.reduce((s, d) => s + d.min, 0);
  const libre = income - fijos - pic.variable - mins;
  const strategy = state.settings.strategy || 'avalancha';
  const plan = simulate(debts, mins + Math.max(0, libre), strategy);
  const only = minimumsOnly(debts);
  const cutCats = ['delivery', 'resto', 'alcohol', 'compras'];
  const recorte = cutCats.reduce((s, k) => s + (pic.varByCat[k] || 0), 0) / 2;
  const maxApr = Math.max(0, ...debts.map((d) => d.apr));
  const extraInv = pic.invest > 0 && maxApr > 8 ? pic.invest : 0;
  const plan2 = simulate(debts, mins + Math.max(0, libre + recorte + extraInv), strategy);
  return { fixed, pic, cards, fijos, income, debts, mins, libre, strategy, plan, only, recorte, maxApr, extraInv, plan2, libre2: libre + recorte + extraInv };
}

// Acciones concretas con fecha y monto, como te las diría un asesor.
function concreteActions(m) {
  const today = todayISO();
  const acts = [];
  const target = m.plan.order?.find((d) => d.status !== 'usar') || m.plan.order?.[0];
  const soon = m.cards.filter((c) => c.dueDate && c.daysToDue <= 35 && c.kind !== 'prestamo').sort((a, b) => a.dueDate.localeCompare(b.dueDate));
  for (const c of soon) {
    const when = `${fmtDate(c.dueDate)} (${inDays(c.daysToDue)})`;
    if (c.status === 'usar' && c.statementBalance > 0) {
      acts.push({ date: c.dueDate, html: `<b>Antes del ${when}: paga completo ${money(c.statementBalance)} de ${esc(shortCard(c.name))}.</b> Hoy no te cobra intereses; si dejas aunque sea un poco, empieza a cobrar ~${pct(c.apr ?? 20)} sobre todo el saldo.${c.dueSrc === 'estimada' ? ' Confirma la fecha exacta en la app del banco.' : ''}` });
    } else if (c.minPayment) {
      const isTarget = target && target.id === c.id;
      acts.push({ date: c.dueDate, html: `<b>Antes del ${when}: paga por lo menos ${money(c.minPayment)} de ${esc(shortCard(c.name))}</b>${isTarget ? `, y todo lo extra que puedas: es la deuda más cara (${pct(c.apr)}).` : '.'}` });
    } else if (c.balance == null) {
      acts.push({ date: c.dueDate, html: `<b>${esc(shortCard(c.name))}: escribe el saldo y el pago mínimo en Deudas</b> para saber cuánto pagar antes del ${when}.` });
    }
  }
  // Próximo sueldo: cuánto apartar y a dónde mandar el resto
  const pi = m.pic.payInfo;
  if (pi?.every) {
    let next = pi.last; while (next <= today) next = shiftDateISO(next, pi.every);
    const after = shiftDateISO(next, pi.every);
    const fixedDue = m.fixed.filter((f) => f.account && isDebit(f.account)).map((f) => ({ f, ...nextOccurrence(f, next) }))
      .filter((x) => x.next >= next && x.next < after);
    const fixedSum = fixedDue.reduce((s, x) => s + x.f.amount, 0);
    const perCheckExtra = m.libre > 0 ? m.libre * 12 / 26 : 0;
    acts.push({ date: next, html: `<b>El ${fmtDate(next)} entra tu sueldo (~${money(pi.amount)}).</b> ${fixedDue.length ? `Aparta ${money(fixedSum)} para ${fixedDue.map((x) => esc(shortFixed(x.f.name))).join(', ')} antes del ${fmtDate(after)}. ` : ''}${perCheckExtra > 10 && target ? `Manda ${money(perCheckExtra)} extra a ${esc(target.name)}.` : 'Por ahora no queda sobrante: cubre fijos y mínimos, y no uses tarjetas con saldo.'}` });
  }
  const cm = cashModel(m);
  if (cm && cm.low.bal < 0) {
    const before = cm.evs.filter((e) => e.date <= cm.low.date && e.amount < -300).map((e) => `${esc(shortFixed(e.label))} (${money(-e.amount)})`);
    acts.push({ date: shiftDateISO(cm.low.date, -1), html: `<b>Ojo con tu cuenta de cheques: el ${fmtDate(cm.low.date)} quedaría en ${money(cm.low.bal)}.</b> Antes de esa fecha salen ${before.join(', ')}. Paga primero la renta y lo que evita intereses; si no alcanza, paga una parte ahora y el resto cuando entre el sueldo, sin pasarte de la fecha límite.` });
  }
  acts.sort((a, b) => a.date.localeCompare(b.date));
  // Recordatorios sin fecha
  const quietas = m.cards.filter((c) => c.status === 'quieta' && c.kind !== 'linea');
  if (quietas.length) acts.push({ html: `<b>No uses ${quietas.map((c) => esc(shortCard(c.name))).join(' ni ')}</b> hasta dejarlas en cero. Para el día a día, débito${m.cards.some((c) => c.status === 'usar') ? ` o ${esc(shortCard(m.cards.find((c) => c.status === 'usar').name))} (pagándola completa)` : ''}.` });
  for (const [cat, spent, budget] of budgetStatus().filter((b) => b[1] >= b[2] * 0.8)) {
    acts.push({ html: `<b>${esc(CAT[cat].name)}: llevas ${money(spent)} de ${money(budget)}</b> este mes. ${spent >= budget ? 'Ya te pasaste: frena hasta el próximo mes.' : `Te quedan ${money(budget - spent)}.`}` });
  }
  for (const w of weekStatus().filter((x) => x.spent >= x.limit * 0.8)) {
    acts.push({ html: `<b>${esc(w.name)}: esta semana llevas ${money(w.spent)} de ${money(w.limit)}.</b> ${w.spent >= w.limit ? 'Ya te pasaste: cocina en casa hasta el lunes.' : `Te quedan ${money(w.limit - w.spent)} hasta el domingo.`}` });
  }
  const stale = staleAccounts();
  if (stale.length) acts.push({ html: `<b>Actualiza tus datos:</b> ${stale.map((s) => `${esc(shortAcct(s.id))} (${s.days} días)`).join(', ')}. Mientras más al día, mejores consejos.` });
  return acts;
}

function shiftDateISO(iso, d) { const x = new Date(iso + 'T12:00:00Z'); x.setUTCDate(x.getUTCDate() + d); return x.toISOString().slice(0, 10); }
const isDebit = (id) => state.accounts.find((a) => a.id === id)?.kind === 'debito';
const shortFixed = (n) => String(n).replace(/^\d+[.,]\d+\s+/, '').replace(/\s\+\d+/g, '').replace(/\s+\d[\d\s]*$/, '').replace(/\s+(ltd|inc\.?|holdings)\b.*$/i, '').slice(0, 34);

function staleAccounts() {
  const today = todayISO();
  return state.accounts.filter((a) => !a.manual).map((a) => {
    const last = state.txns.filter((t) => t.account === a.id).map((t) => t.date).sort().pop() || a.lastImport?.slice(0, 10);
    return { id: a.id, last, days: last ? daysBetween(last, today) : 999 };
  }).filter((x) => x.days > 7).sort((a, b) => b.days - a.days);
}

function renderAsesoria() {
  const el = $('#view-asesoria');
  if (!state.txns.length) { el.innerHTML = emptyState(); return; }
  const m = adviceModel();
  const { pic, cards, debts, plan, only, plan2, strategy, libre, libre2, recorte, extraInv, maxApr, mins, fijos, income } = m;
  const quietas = cards.filter((c) => c.status === 'quieta' && c.kind !== 'linea');
  const line = cards.find((c) => c.kind === 'linea');
  const first = plan.order?.[0];
  const missing = debts.filter((d) => d.aprMissing).map((d) => d.name).concat(cards.filter((c) => c.balanceSrc === 'estimado' || c.balance == null).map((c) => shortCard(c.name)));
  const topVar = Object.entries(pic.varByCat).sort((a, b) => b[1] - a[1]).slice(0, 3);

  const steps = [];
  if (libre < 0) steps.push(`<b>Hoy sale más de lo que entra: te faltan ${money(-libre)} al mes.</b> Antes de cualquier plan hay que recortar. Donde más se va el dinero del día a día: ${topVar.map(([k, v]) => `${CAT[k].name} (${money(v)})`).join(', ')}. Si algo de eso fue de una sola vez, márcalo así en Movimientos.`);
  steps.push(`<b>Paga siempre el mínimo de todas a tiempo</b> (${money(mins)} al mes). Un pago tarde suma cargos, sube la tasa y baja tu puntaje de crédito.`);
  if (first && libre > 0) steps.push(`<b>Todo lo que sobra (${money(libre)} al mes) va a ${esc(first.name)}</b>, ${strategy === 'avalancha' ? `porque es la de tasa más alta (${pct(first.apr)})` : 'porque es la de saldo más chico'}. Cuando la termines, ese dinero pasa a la siguiente.`);
  if (quietas.length) steps.push(`<b>Deja quietas ${quietas.map((c) => esc(shortCard(c.name))).join(', ')}</b>. Mientras tengan saldo, cada compra cobra interés desde el primer día.`);
  if (extraInv) steps.push(`<b>Considera pausar lo que mandas a inversión (${money(extraInv)} al mes)</b> y usarlo en deudas. Ninguna inversión normal te da ${pct(maxApr)} seguro, que es lo que te ahorras al pagar esa deuda. La excepción es si tu empleador iguala tu aporte.`);
  if (recorte > 20) steps.push(`<b>Recorta a la mitad comida a domicilio, restaurantes, licor y compras</b>: liberas unos ${money(recorte)} al mes.`);
  if (line && maxApr > (line.apr ?? 99)) steps.push(`<b>Pregunta en Scotiabank</b> si puedes pasar la deuda de la tarjeta más cara a tu ScotiaLine (${pct(line.apr)}). Pagarías menos interés, pero solo sirve si después no vuelves a usar esa tarjeta.`);
  else if (line && line.apr == null) steps.push(`<b>Escribe la tasa de tu ScotiaLine</b> en Deudas. Si es más baja que la de tus tarjetas, conviene usarla para pagar la tarjeta más cara.`);
  steps.push(`<b>Guarda un colchón pequeño</b> (unos $1,000) para emergencias, así no vuelves a usar las tarjetas cuando algo pase.`);

  const acts = concreteActions(m);
  el.innerHTML = `
    ${missing.length ? `<div class="alert warn" data-action="goto" data-view="deudas">Faltan datos reales de: ${esc([...new Set(missing)].join(', '))}. Los consejos usan estimaciones hasta que los escribas en <b>Deudas</b>.</div>` : ''}
    <div class="card hl"><h2>Qué hacer ahora</h2><ul class="steps acts">${acts.map((a) => `<li>${a.html}</li>`).join('')}</ul></div>
    ${leaksCard(m)}
    ${weeklyCard()}
    ${lumpCard(cards)}
    <div class="card"><h2>Tu mes típico</h2>
      <div class="grid2 big">
        <div>Entra</div><div class="r pos">${money(income)}</div>
        <div>Gastos fijos</div><div class="r">−${money(fijos)}</div>
        <div>Gastos del día a día</div><div class="r">−${money(pic.variable)}</div>
        <div>Mínimos de deudas</div><div class="r">−${money(mins)}</div>
        <div class="tot">Queda libre</div><div class="r tot ${libre >= 0 ? 'pos' : 'neg'}">${money(libre)}</div>
      </div>
      <p class="small muted">Con tus últimos 30 días (${fmtDate(pic.start)} a ${fmtDate(pic.end)})${pic.incomeNote ? `, ${esc(pic.incomeNote)}` : ''}. No cuenta intereses, inversión ni gastos marcados como "de una sola vez".</p>
      <label class="field inline"><span>¿Tu ingreso mensual es otro?</span><input id="f-income" type="number" inputmode="decimal" value="${state.settings.income ?? ''}" placeholder="${Math.round(pic.income)}"><button class="btn sm" data-action="saveincome">OK</button></label>
    </div>
    ${cashCard(m)}
    ${debts.length ? `<div class="card"><h2>Tu camino para salir de deudas</h2>
      <div class="seg"><button data-action="strategy" data-s="avalancha" class="${strategy === 'avalancha' ? 'sel' : ''}">Pagar menos intereses</button><button data-action="strategy" data-s="bolaDeNieve" class="${strategy === 'bolaDeNieve' ? 'sel' : ''}">Ver avances rápido</button></div>
      ${plan.feasible && libre >= 0 ? `<p>Con este plan terminas en <b>${yearsMonths(plan.months)}</b> (${monthLabelFromNow(plan.months)}) y pagas <b>${money(plan.totalInterest)}</b> de intereses.</p>
        <ol class="order">${plan.order.map((d) => `<li><b>${esc(d.name)}</b> · ${money(d.balance)} al ${pct(d.apr)}${d.aprMissing ? '<span class="est">*</span>' : ''}<div class="small muted">Pagada en ${monthLabelFromNow(d.paidOffMonth)}</div></li>`).join('')}</ol>`
        : `<p class="neg">Con lo que entra hoy no alcanza para bajar las deudas más allá de los mínimos. Primero hay que cerrar ese hueco.</p>
          <p class="small">Cuando sobre dinero, este es el orden para atacarlas:</p>
          <ol class="order">${plan.order.map((d) => `<li><b>${esc(d.name)}</b> · ${money(d.balance)} al ${pct(d.apr)}${d.aprMissing ? '<span class="est">*</span>' : ''}</li>`).join('')}</ol>`}
      <p class="small">Pagando solo los mínimos: <b>${yearsMonths(only.months)}</b> y ${money(only.totalInterest)} de intereses.</p>
      ${(recorte > 20 || extraInv) ? (libre2 > 0 && plan2.feasible
        ? `<p class="small">Si recortas${extraInv ? ' y pausas la inversión' : ''}, te quedan ${money(libre2)} libres al mes: terminas en <b>${yearsMonths(plan2.months)}</b> y te ahorras ${money(Math.max(0, (plan.feasible && libre >= 0 ? plan.totalInterest : only.totalInterest) - plan2.totalInterest))} en intereses.</p>`
        : `<p class="small">Aun recortando${extraInv ? ' y pausando la inversión' : ''} (${money(recorte + extraInv)} al mes) faltarían ${money(-libre2)} al mes. Revisa los gastos fijos más grandes.</p>`) : ''}
    </div>` : ''}
    <div class="card"><h2>Reglas para tomar buenas decisiones</h2><ol class="steps">${steps.map((s) => `<li>${s}</li>`).join('')}</ol></div>
    ${subsCard(m.fixed)}
    ${goalsCard()}
    <p class="small muted">Orientación general calculada con tus números, no asesoría financiera profesional. Antes de mover deudas entre productos, confirma tasas y condiciones con tu banco.</p>`;
}

// ---------- tengo $X para pagar ----------

function lumpCard(cards) {
  const amt = state.lump || 0;
  const res = amt > 0 ? allocateLump(amt, cards.map((c) => ({ ...c, name: shortCard(c.name) }))) : null;
  return `<div class="card"><h2>¿Tienes dinero para pagar deudas?</h2>
    <label class="field inline"><span>¿Cuánto puedes pagar hoy?</span><input id="f-lump" type="number" inputmode="decimal" value="${amt || ''}" placeholder="ej. 1500"><button class="btn sm" data-action="lump">Ver</button></label>
    ${res ? `${res.rows.map((r) => `<div class="frow"><div class="ic">💳</div><div><b>${esc(r.name)}</b><div class="small muted">${esc(r.why)}</div></div><div class="amt">${money(r.amount)}${r.saved > 1 ? `<div class="small pos">ahorras ~${money(r.saved)}${r.why.startsWith('Pagar el estado') ? '' : '/año'}</div>` : ''}</div></div>`).join('')}
      ${res.left > 1 ? `<p class="small">Te sobran ${money(res.left)}: guárdalos como colchón de emergencia.</p>` : ''}
      <button class="btn secondary" data-action="applylump">Ya hice estos pagos: actualizar saldos</button>` : '<p class="small muted">Te digo a qué tarjeta va cada dólar para ahorrar más intereses.</p>'}
  </div>`;
}

async function applyLump() {
  const res = allocateLump(state.lump, allCards());
  const now = new Date().toISOString();
  for (const r of res.rows) {
    const a = state.accounts.find((x) => x.id === r.id);
    const c = allCards().find((x) => x.id === r.id);
    if (!a || !c) continue;
    a.user = { ...(a.user || {}) };
    if (c.balance != null) { a.user.balance = Math.max(0, Math.round((c.balance - r.amount) * 100) / 100); a.user.balanceAt = now; }
    if (c.statementBalance) a.user.statementBalance = Math.max(0, Math.round((c.statementBalance - r.amount) * 100) / 100);
    if (c.minPayment) a.user.minPayment = Math.max(0, Math.round((c.minPayment - r.amount) * 100) / 100);
    await db.putMany('accounts', [a]);
  }
  state.lump = 0;
  await snapshotDebt();
  render();
  scheduleBackup();
  toast('Saldos actualizados. ¡Bien hecho!');
}

// ---------- proyección de efectivo ----------

function cashModel(m) {
  const chq = state.accounts.find((a) => a.kind === 'debito' && a.balance != null);
  const pi = m.pic.payInfo;
  if (!chq || !pi) return null;
  const lastDate = state.txns.filter((t) => t.account === chq.id).map((t) => t.date).sort().pop() || todayISO();
  const today = todayISO();
  const paydays = [];
  if (pi.every) { let d = pi.last; for (let i = 0; i < 6; i++) { d = shiftDateISO(d, pi.every); paydays.push(d); } }
  const fixed = m.fixed.filter((f) => isDebit(f.account)).map((f) => ({ ...f, ...nextOccurrence(f, lastDate) }));
  const cardPayments = m.cards.filter((c) => c.dueDate && c.kind !== 'prestamo').map((c) => ({
    date: c.dueDate, name: shortCard(c.name), amount: c.status === 'usar' && c.statementBalance ? c.statementBalance : (c.minPayment || 0),
  })).filter((c) => c.amount > 0);
  // Gasto del día a día que sale directo de la cuenta de cheques, por semana
  const win = state.txns.filter((t) => t.account === chq.id && t.amount > 0 && !t.oneOff && CAT[t.cat]?.kind === 'gasto' && daysBetween(t.date, lastDate) <= 30);
  const fixedKeys = new Set(m.fixed.flatMap((f) => f.keys));
  const weekly = win.filter((t) => !fixedKeys.has(merchantKey(t.desc))).reduce((s, t) => s + t.amount, 0) / 30 * 7;
  const weeks = []; for (let d = shiftDateISO(lastDate, 7); daysBetween(lastDate, d) <= 35; d = shiftDateISO(d, 7)) weeks.push(d);
  const proj = cashProjection({ start: chq.balance, startDate: lastDate, paydays, payAmount: pi.amount, fixed: fixed.map((f) => ({ ...f })), cardPayments });
  // agrega el gasto semanal estimado
  let bal = chq.balance, low = { bal, date: lastDate };
  const evs = [...proj.events, ...weeks.map((d) => ({ date: d, label: 'Día a día (estimado)', amount: -weekly }))].sort((a, b) => a.date.localeCompare(b.date) || b.amount - a.amount);
  for (const e of evs) { bal += e.amount; e.bal = bal; if (bal < low.bal) low = { bal, date: e.date }; }
  return { chq, lastDate, evs, low };
}

function cashCard(m) {
  const cm = cashModel(m);
  if (!cm) return '';
  const { chq, lastDate, evs, low } = cm;
  return `<div class="card"><h2>Tu cuenta de cheques en las próximas 5 semanas</h2>
    <p>Hoy tienes <b>${money(chq.balance)}</b> (al ${fmtDate(lastDate)}). ${low.bal < 0
      ? `<span class="neg"><b>Ojo: el ${fmtDate(low.date)} quedarías en ${money(low.bal)}.</b></span> Mueve pagos o recorta antes de esa fecha.`
      : `Lo más bajo sería <b>${money(low.bal)}</b> el ${fmtDate(low.date)}.`}</p>
    <div class="cash">${evs.map((e) => `<div class="crow ${e.bal < 0 ? 'neg' : ''}"><div class="small muted">${fmtDate(e.date)}</div><div class="small">${esc(shortFixed(e.label))}</div><div class="small r ${e.amount > 0 ? 'pos' : ''}">${money(e.amount, true)}</div><div class="small r"><b>${money(e.bal)}</b></div></div>`).join('')}</div>
    <p class="small muted">Incluye sueldos, gastos fijos de esta cuenta, pagos de tarjetas (completo en las que no cobran interés, mínimo en las demás) y tu gasto semanal típico.</p>
  </div>`;
}

// ---------- suscripciones ----------

function subsCard(fixed) {
  const { subs, monthly } = reviewSubscriptions(fixed);
  if (!subs.length) return '';
  return `<div class="card"><h2>Suscripciones · ${money(monthly)} al mes (${money(monthly * 12)} al año)</h2>
    ${subs.map((s) => `<div class="frow"><div class="ic">🔁</div><div>${esc(shortFixed(s.name))}<div class="small muted">${esc(shortAcct(s.account))}${s.flags.length ? ` · <span class="neg">${esc(s.flags.join('; '))}</span>` : ''}</div></div><div class="amt">${money(s.perMonth)}</div></div>`).join('')}
    <p class="small muted">¿Usas todas? Cancelar una de $15 son $180 al año para tus deudas.</p></div>`;
}

// ---------- metas de ahorro ----------

function goalsCard() {
  const goals = state.settings.goals || [];
  return `<div class="card"><h2>Metas de ahorro</h2>
    ${goals.map((g, i) => {
      const p = g.target ? Math.min(1, (g.saved || 0) / g.target) : 0;
      return `<div class="goal" data-action="editgoal" data-i="${i}"><div class="dhead"><b>${esc(g.name)}</b><span class="small">${money(g.saved || 0)} de ${money(g.target)}</span></div>
        <div class="ubar"><i class="ok" style="width:${Math.max(p * 100, 2)}%"></i></div></div>`;
    }).join('') || '<p class="small muted">Aún no tienes metas. Te recomiendo empezar con un colchón de emergencia de $1,000.</p>'}
    <button class="btn secondary sm" data-action="editgoal" data-i="-1">+ Agregar meta</button></div>`;
}

function openGoal(i) {
  const goals = state.settings.goals || [];
  const g = i >= 0 ? goals[i] : { name: goals.length ? '' : 'Colchón de emergencia', target: goals.length ? '' : 1000, saved: 0 };
  $('#sheetPanel').innerHTML = `<h3>${i >= 0 ? 'Editar meta' : 'Nueva meta'}</h3>
    <label class="field"><span>Nombre</span><input id="g-name" type="text" value="${esc(g.name)}"></label>
    ${numField('g-target', 'Cuánto quieres juntar', g.target)}
    ${numField('g-saved', 'Cuánto llevas', g.saved)}
    <button class="btn" data-action="savegoal" data-i="${i}">Guardar</button>
    ${i >= 0 ? `<button class="btn danger" data-action="delgoal" data-i="${i}">Borrar meta</button>` : ''}
    <button class="btn secondary" data-close>Cancelar</button>`;
  $('#sheet').classList.remove('hidden');
}

async function saveGoal(i, del) {
  const goals = [...(state.settings.goals || [])];
  if (del) goals.splice(i, 1);
  else {
    const g = { name: $('#g-name').value.trim() || 'Meta', target: Number($('#g-target').value) || 0, saved: Number($('#g-saved').value) || 0 };
    if (i >= 0) goals[i] = g; else goals.push(g);
  }
  state.settings.goals = goals;
  await saveSettings();
  closeSheet(); renderAsesoria();
}

// ---------- presupuestos ----------

const latestMonth = () => monthsWithData().pop();
function budgetStatus(month = latestMonth()) {
  const b = state.settings.budgets || {};
  const spent = {};
  for (const t of state.txns) if (t.date.startsWith(month) && CAT[t.cat]?.kind === 'gasto') spent[t.cat] = (spent[t.cat] || 0) + t.amount;
  return Object.entries(b).filter(([, v]) => v > 0).map(([cat, v]) => [cat, spent[cat] || 0, v]);
}

function budgetBanner() {
  if (state.month !== latestMonth()) return '';
  const over = budgetStatus().filter(([, s, b]) => s >= b * 0.8).sort((a, b) => b[1] / b[2] - a[1] / a[2]);
  if (!over.length) return '';
  const more = over.length - 4;
  return `<div class="alert ${over.some(([, s, b]) => s >= b) ? '' : 'warn'}"><b>🎯 Presupuesto</b>${over.slice(0, 4).map(([c, s, b]) => `<div class="small">${CAT[c].icon} ${esc(CAT[c].name)}: ${money(s)} de ${money(b)} (${Math.round(s / b * 100)}%)${s >= b ? ' · <b>te pasaste</b>' : ''}</div>`).join('')}${more > 0 ? `<div class="small">y ${more} más</div>` : ''}</div>`;
}

function openBudgets() {
  const b = state.settings.budgets || {};
  const month = latestMonth();
  const spent = {};
  for (const t of state.txns) if (t.date.startsWith(month) && CAT[t.cat]?.kind === 'gasto') spent[t.cat] = (spent[t.cat] || 0) + t.amount;
  const cats = CATEGORIES.filter((c) => c.kind === 'gasto' && !['intereses'].includes(c.id));
  $('#sheetPanel').innerHTML = `<h3>Presupuesto por mes</h3>
    <p class="small muted">Pon un límite a las categorías que quieres controlar. Te aviso al llegar al 80%. Deja vacío lo que no quieras limitar.</p>
    <button class="btn secondary sm" data-action="suggestbudget">Sugerirme montos (20% menos en lo recortable)</button>
    ${cats.map((c) => `<label class="field inline2"><span>${c.icon} ${esc(c.name)}<small class="muted"> · ${esc(MESES[Number(month.slice(5)) - 1])}: ${money(spent[c.id] || 0)}</small></span><input type="number" inputmode="decimal" data-budget="${c.id}" value="${b[c.id] ?? ''}" placeholder="—"></label>`).join('')}
    <button class="btn" data-action="savebudget">Guardar</button>
    <button class="btn secondary" data-close>Cancelar</button>`;
  $('#sheetPanel').dataset.spent = JSON.stringify(spent);
  $('#sheet').classList.remove('hidden');
}

function suggestBudgets() {
  const spent = JSON.parse($('#sheetPanel').dataset.spent || '{}');
  const cut = new Set(['delivery', 'resto', 'alcohol', 'compras', 'otros', 'personal', 'hogar']);
  const skip = new Set(['vivienda', 'prestamos', 'seguros', 'servicios', 'suscripciones', 'tramites', 'servprof', 'donaciones', 'efectivo']);
  document.querySelectorAll('[data-budget]').forEach((inp) => {
    const v = spent[inp.dataset.budget];
    if (!v || skip.has(inp.dataset.budget)) return;
    inp.value = cut.has(inp.dataset.budget) ? Math.max(10, Math.round(v * 0.8 / 10) * 10) : Math.ceil(v * 1.1 / 50) * 50;
  });
}

async function saveBudgets() {
  const b = {};
  document.querySelectorAll('[data-budget]').forEach((inp) => { const v = Number(inp.value); if (inp.value !== '' && v > 0) b[inp.dataset.budget] = v; });
  state.settings.budgets = b;
  await saveSettings();
  closeSheet(); render();
  toast('Presupuesto guardado');
}

// ---------- historial de deuda ----------

async function snapshotDebt() {
  const cards = allCards();
  const total = cards.reduce((s, c) => s + (c.balance > 0 ? c.balance : 0), 0);
  if (!total) return;
  const hist = (await db.kvGet('debtHist')) || {};
  hist[todayISO()] = Math.round(total * 100) / 100;
  state.debtHist = hist;
  await db.kvSet('debtHist', hist);
}

function debtChart() {
  const m = adviceModel();
  const hist = Object.entries(state.debtHist || {}).sort();
  const today = todayISO();
  const pts = hist.map(([d, v]) => ({ t: daysBetween(today, d) / 30.4, v }));
  const plan = m.libre >= 0 && m.plan.feasible ? m.plan : (m.libre2 > 0 && m.plan2.feasible ? m.plan2 : null);
  const proj = plan ? plan.trace.map((v, i) => ({ t: i, v })) : [];
  const all = [...pts, ...proj];
  if (!all.length) return '';
  const W = 320, H = 150, P = 28;
  const tMin = Math.min(0, ...all.map((p) => p.t)), tMax = Math.max(1, ...all.map((p) => p.t));
  const vMax = Math.max(...all.map((p) => p.v)) * 1.1 || 1;
  const x = (t) => P + (t - tMin) / (tMax - tMin) * (W - P - 8);
  const y = (v) => 8 + (1 - v / vMax) * (H - 30);
  const line = (ps) => ps.map((p, i) => `${i ? 'L' : 'M'}${x(p.t).toFixed(1)},${y(p.v).toFixed(1)}`).join(' ');
  const goal = plan ? `Libre de deudas en <b>${monthLabelFromNow(plan.months)}</b> si sigues el plan${plan === m.plan2 ? ' con los recortes' : ''}.` : 'Con lo que entra hoy el plan no alcanza para bajarlas. Revisa Asesoría.';
  return `<div class="card"><h2>Tu deuda en el tiempo</h2>
    <svg viewBox="0 0 ${W} ${H}" class="chart" role="img" aria-label="Deuda total en el tiempo">
      <line x1="${P}" x2="${W - 8}" y1="${y(0)}" y2="${y(0)}" class="axis"/>
      <text x="${P - 4}" y="${y(vMax / 1.1) + 4}" text-anchor="end" class="lbl">${Math.round(vMax / 1.1 / 1000)}k</text>
      <text x="${P - 4}" y="${y(0) + 4}" text-anchor="end" class="lbl">0</text>
      ${proj.length ? `<path d="${line(proj)}" class="proj"/>` : ''}
      ${pts.length > 1 ? `<path d="${line(pts)}" class="real"/>` : ''}
      ${pts.map((p) => `<circle cx="${x(p.t)}" cy="${y(p.v)}" r="3.5" class="dot"/>`).join('')}
      <text x="${x(0)}" y="${H - 4}" text-anchor="middle" class="lbl">hoy</text>
      ${plan ? `<text x="${x(plan.months)}" y="${H - 4}" text-anchor="end" class="lbl">${monthLabelFromNow(plan.months)}</text>` : ''}
    </svg>
    <p class="small">${goal}</p>
    <p class="small muted">Los puntos son tu deuda real cada vez que actualizas; la línea punteada es el plan. Si los puntos van por debajo de la línea, vas adelantado.</p></div>`;
}

// ---------- recordatorio para actualizar ----------

const BANK_STEPS = [
  ['Scotiabank (las 3 cuentas)', 'En scotiabank.com, entra a cada cuenta y descarga las transacciones en Excel.'],
  ['American Express', 'En la web de Amex: Estados y actividad → Descargar → Excel.'],
  ['Canadian Tire', 'En la web de Triangle: Transacciones → Descargar CSV.'],
  ['Capital One', 'Descarga el estado en PDF cuando salga, y entre estados imprime "Posted transactions" como PDF.'],
];

function updateCard() {
  const stale = staleAccounts();
  const freq = state.settings.reminder || 'semanal';
  return `<div class="card"><h2>Mantén tus datos al día</h2>
    ${stale.length ? `<p class="small neg">Hace más de una semana que no actualizas: ${stale.map((s) => esc(shortAcct(s.id))).join(', ')}.</p>` : '<p class="small pos">Tus datos están al día.</p>'}
    <ol class="steps small">${BANK_STEPS.map(([b, s]) => `<li><b>${esc(b)}:</b> ${esc(s)}</li>`).join('')}<li>Aquí, toca <b>Importar</b> y súbelos todos juntos. Lo repetido se ignora.</li></ol>
    <div class="seg"><button data-action="reminder" data-f="semanal" class="${freq === 'semanal' ? 'sel' : ''}">Aviso semanal</button><button data-action="reminder" data-f="diario" class="${freq === 'diario' ? 'sel' : ''}">Aviso diario</button></div>
    <button class="btn secondary" data-action="icsupdate">📅 Agregar el aviso al Calendario</button>
    <p class="small muted">${freq === 'semanal' ? 'Cada domingo a las 7 PM' : 'Cada día a las 8 PM'} te llega un aviso con estos pasos. Toma 5 minutos.</p></div>`;
}

function staleBanner() {
  const s = staleAccounts();
  if (!s.length) return '';
  return `<div class="alert warn" data-action="goto" data-view="cuentas">🔄 Hace ${s[0].days} días que no actualizas ${esc(shortAcct(s[0].id))}${s.length > 1 ? ` y ${s.length - 1} cuenta${s.length > 2 ? 's' : ''} más` : ''}. Toca para ver qué bajar.</div>`;
}

function downloadUpdateICS() {
  const freq = state.settings.reminder || 'semanal';
  const steps = BANK_STEPS.map(([b, s]) => `${b}: ${s}`).join('\\n') + '\\nLuego abre Finanzas → Cuentas → Importar.';
  const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+Z$/, 'Z');
  const d = new Date(); d.setDate(d.getDate() + ((7 - d.getDay()) % 7 || 7));
  const day = freq === 'semanal' ? d : new Date();
  const ds = `${day.getFullYear()}${String(day.getMonth() + 1).padStart(2, '0')}${String(day.getDate()).padStart(2, '0')}`;
  const ics = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Finanzas//ES', 'METHOD:PUBLISH',
    'BEGIN:VEVENT', 'UID:actualizar@finanzas', `DTSTAMP:${stamp}`, `DTSTART:${ds}T${freq === 'semanal' ? '190000' : '200000'}`, 'DURATION:PT15M',
    `RRULE:FREQ=${freq === 'semanal' ? 'WEEKLY;BYDAY=SU' : 'DAILY'}`, 'SUMMARY:🔄 Actualiza Finanzas', `DESCRIPTION:${steps}`,
    'BEGIN:VALARM', 'ACTION:DISPLAY', 'DESCRIPTION:Actualiza tus gastos en Finanzas', 'TRIGGER:PT0M', 'END:VALARM', 'END:VEVENT', 'END:VCALENDAR'].join('\r\n');
  saveFile(new Blob([ics], { type: 'text/calendar' }), 'actualizar-finanzas.ics');
  toast('Abre el archivo y toca "Añadir"');
}

function saveFile(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

async function saveSettings() {
  await db.kvSet('settings', state.settings);
  scheduleBackup();
}

// ---------- respaldo en Google Drive ----------

let backupTimer;
function scheduleBackup() {
  if (!state.settings.driveClientId) return;
  state.settings.backupDirty = true;
  db.kvSet('settings', state.settings);
  if (!gdrive.getToken()) return;
  clearTimeout(backupTimer);
  backupTimer = setTimeout(() => driveBackup(true), 2500);
}

function backupData() {
  return { app: 'finanzas', version: 2, exportedAt: new Date().toISOString(), txns: state.txns, accounts: state.accounts, rules: state.rules, fixed: state.fixedOv, settings: { ...state.settings, backupDirty: undefined }, debtHist: state.debtHist || {} };
}

async function driveBackup(silent) {
  const id = state.settings.driveClientId;
  if (!id) { toast('Primero pega tu Client ID de Google'); return; }
  if (!gdrive.getToken()) { gdrive.connect(id, 'backup'); return; }
  try {
    if (!silent) toast('Respaldando en tu Google Drive…', 0);
    await gdrive.upload(backupData());
    state.settings.lastBackup = new Date().toISOString();
    state.settings.backupDirty = false;
    await db.kvSet('settings', state.settings);
    if (!silent) toast('Respaldo guardado en tu Google Drive');
    if (state.view === 'cuentas') renderCuentas();
  } catch (e) {
    if (e.code === 'auth' && !silent) gdrive.connect(id, 'backup');
    else if (!silent) toast(e.message);
  }
}

async function driveRestore() {
  const id = state.settings.driveClientId;
  if (!gdrive.getToken()) { gdrive.connect(id, 'restore'); return; }
  try {
    toast('Buscando tu respaldo…', 0);
    const r = await gdrive.download();
    if (!r) { toast('No encontré respaldo en tu Drive'); return; }
    if (!confirm(`¿Traer el respaldo del ${new Date(r.modifiedTime).toLocaleString('es')}? Se junta con lo que ya tienes en este teléfono.`)) { hideToast(); return; }
    await restoreData(r.data);
  } catch (e) { toast(e.message); }
}

function driveCard() {
  const s = state.settings;
  const connected = Boolean(s.driveClientId);
  const last = s.lastBackup ? new Date(s.lastBackup).toLocaleString('es', { day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit' }) : null;
  return `<div class="card"><h2>Respaldo en tu Google Drive</h2>
    ${connected ? `<p class="small">${last ? `Último respaldo: <b>${last}</b>.` : 'Aún no hay respaldo.'} ${s.backupDirty ? '<span class="neg">Hay cambios sin respaldar.</span>' : ''}</p>
      <button class="btn" data-action="drivebackup">Respaldar ahora</button>
      <button class="btn secondary" data-action="driverestore">Traer respaldo de Drive</button>
      <p class="small muted">Se guarda como "Finanzas - respaldo.json" en tu Drive. La app solo puede ver ese archivo, nada más de tu Drive. Mientras la app tenga permiso (dura una hora), cada cambio se respalda solo.</p>`
    : `<p class="small muted">Para activarlo necesitas un "Client ID" de Google (te guío paso a paso). Pégalo aquí:</p>`}
    <label class="field"><span>Client ID de Google</span><input id="f-client" type="text" autocomplete="off" value="${esc(s.driveClientId || '')}" placeholder="xxxx.apps.googleusercontent.com"></label>
    <button class="btn secondary sm" data-action="saveclient">Guardar Client ID</button></div>`;
}

async function saveClientId() {
  const v = ($('#f-client').value || '').trim();
  if (v && !/\.apps\.googleusercontent\.com$/.test(v)) { toast('Ese no parece un Client ID de Google'); return; }
  state.settings.driveClientId = v || undefined;
  await db.kvSet('settings', state.settings);
  renderCuentas();
  toast(v ? 'Listo. Toca "Respaldar ahora"' : 'Client ID borrado');
}

function backupBanner() {
  const s = state.settings;
  if (!s.driveClientId || !s.backupDirty) return '';
  const days = s.lastBackup ? daysBetween(s.lastBackup.slice(0, 10), todayISO()) : 99;
  if (days < 2 && s.lastBackup) return '';
  return `<div class="alert warn" data-action="drivebackup">☁️ Tienes cambios sin respaldar${s.lastBackup ? ` desde hace ${days} días` : ''}. Toca para respaldar en tu Google Drive.</div>`;
}

async function saveIncome() {
  const v = Number($('#f-income').value);
  if ($('#f-income').value === '' ) delete state.settings.income; else if (isFinite(v) && v > 0) state.settings.income = v;
  await db.kvSet('settings', state.settings);
  renderAsesoria();
  toast('Ingreso actualizado');
}

// Cambio contra el promedio de meses anteriores (solo si hay al menos 2 meses completos antes)
function categoryTrends() {
  const months = monthsWithData().filter((m) => m < state.month);
  const full = months.filter((m) => state.txns.filter((t) => t.date.startsWith(m)).length >= 40);
  if (full.length < 2) return {};
  const cur = summary(monthTxns()).cats;
  const out = {};
  for (const [id, v] of cur) {
    const prev = full.map((m) => state.txns.filter((t) => t.date.startsWith(m) && t.cat === id).reduce((s, t) => s + t.amount, 0));
    const avg = prev.reduce((a, b) => a + b, 0) / prev.length;
    if (avg > 20 && Math.abs(v / avg - 1) >= 0.15) out[id] = v / avg - 1;
  }
  return out;
}

// ---------- fugas de dinero ----------

function leaks(m) {
  const end = m.pic.end, start = m.pic.start;
  const win = state.txns.filter((t) => t.date > start && t.date <= end && t.amount > 0);
  const sumBy = (fn) => win.filter(fn).reduce((s, t) => s + t.amount, 0);
  const out = [];
  const protection = sumBy((t) => /credit card protec|balance protection|payment protection|card protection/i.test(t.raw || t.desc));
  if (protection > 5) out.push({ save: protection, icon: '🛡️', html: `<b>Seguro de protección de tarjeta: ${money(protection)} al mes.</b> Es un seguro opcional que el banco cobra sobre tu saldo. Llama y cancélalo.` });
  const fees = sumBy((t) => t.cat === 'intereses' && /monthly fee|service charge|odp|overdraft|annual fee|late fee|nsf/i.test(t.raw || t.desc));
  if (fees > 3) out.push({ save: fees, icon: '🏦', html: `<b>Comisiones del banco: ${money(fees)} al mes.</b> Pregunta qué saldo mínimo evita la cuota mensual, o cambia a una cuenta sin comisión.` });
  const eat = sumBy((t) => t.cat === 'resto' || t.cat === 'delivery');
  const eatGoal = (state.settings.weekly?.eatout ?? 60) * 52 / 12;
  if (eat > eatGoal + 20) out.push({ save: eat - eatGoal, icon: '🍔', html: `<b>Comida fuera y a domicilio: ${money(eat)} al mes.</b> Con un límite de ${money(state.settings.weekly?.eatout ?? 60)} por semana gastarías ${money(eatGoal)}.` });
  if (m.extraInv) out.push({ save: m.extraInv, icon: '📈', html: `<b>Inversión mientras pagas ${pct(m.maxApr)} de interés: ${money(m.extraInv)} al mes.</b> Pausarla hasta pagar las tarjetas rinde más que cualquier inversión.` });
  const { subs } = reviewSubscriptions(m.fixed);
  const dup = subs.filter((s) => s.flags.some((f) => f.includes('dos veces')));
  if (dup.length) { const c = Math.min(...dup.map((d) => d.perMonth)); out.push({ save: c, icon: '🔁', html: `<b>Suscripción repetida: ${esc(shortFixed(dup[0].name))}</b> aparece en ${dup.length} cuentas. Si es la misma, cancela una.` }); }
  const phones = m.fixed.filter((f) => f.cat === 'servicios' && /telus|koodo|rogers|fido|bell|freedom|virgin|chatr|public mobile|fizz/i.test(f.name));
  if (phones.length >= 2) { const c = Math.min(...phones.map((p) => p.perMonth)); out.push({ save: c, icon: '📱', html: `<b>${phones.length} planes de celular</b> (${phones.map((p) => `${esc(shortFixed(p.name).split(' ')[0])} ${money(p.perMonth)}`).join(', ')}). Si alguno es tuyo y no lo usas, cancélalo o junta las líneas.` }); }
  const alc = sumBy((t) => t.cat === 'alcohol');
  if (alc > 60) out.push({ save: alc / 2, icon: '🍷', html: `<b>Licor: ${money(alc)} al mes.</b> A la mitad son ${money(alc / 2)} más para tus deudas.` });
  const housing = m.fixed.filter((f) => f.cat === 'vivienda' || f.cat === 'prestamos').reduce((s, f) => s + f.perMonth, 0);
  const ratio = m.income ? housing / m.income : 0;
  const notes = [];
  if (ratio > 0.5) notes.push({ icon: '🏠', html: `<b>Renta y préstamo se llevan el ${Math.round(ratio * 100)}% de tu sueldo</b> (${money(housing)} de ${money(m.income)}). Lo sano es menos del 40%. No es para hoy, pero al renovar contrato vale la pena buscar algo más barato o un roomie.` });
  out.sort((a, b) => b.save - a.save);
  return { items: out, total: out.reduce((s, x) => s + x.save, 0), notes };
}

function leaksCard(m) {
  const { items, total, notes } = leaks(m);
  if (!items.length && !notes.length) return '';
  return `<div class="card leaks"><h2>Fugas de dinero · puedes recuperar ~${money(total)} al mes</h2>
    ${items.map((x) => `<div class="frow"><div class="ic">${x.icon}</div><div class="small">${x.html}</div><div class="amt pos">+${money(x.save)}<div class="small muted">al mes</div></div></div>`).join('')}
    ${notes.map((x) => `<div class="frow"><div class="ic">${x.icon}</div><div class="small">${x.html}</div><div></div></div>`).join('')}
    <p class="small muted">Calculado con tus últimos 30 días.</p></div>`;
}

// ---------- límites semanales ----------

const WEEKLY = [
  { id: 'eatout', name: 'Comida fuera y a domicilio', icon: '🍔', cats: ['resto', 'delivery'], def: 60 },
  { id: 'compras', name: 'Compras y ropa', icon: '🛍️', cats: ['compras'], def: 50 },
  { id: 'alcohol', name: 'Licor', icon: '🍷', cats: ['alcohol'], def: 25 },
];

function weekStatus() {
  const lim = state.settings.weekly || {};
  const ref = state.txns.map((t) => t.date).sort().pop() || todayISO();
  const d = new Date(ref + 'T12:00:00Z');
  const monday = shiftDateISO(ref, -((d.getUTCDay() + 6) % 7));
  return WEEKLY.map((w) => {
    const limit = lim[w.id] ?? (w.id === 'eatout' ? w.def : null);
    if (!limit) return null;
    const spent = state.txns.filter((t) => t.date >= monday && t.date <= ref && w.cats.includes(t.cat) && t.amount > 0).reduce((s, t) => s + t.amount, 0);
    return { ...w, limit, spent, monday, ref };
  }).filter(Boolean);
}

function weeklyBanner() {
  const st = weekStatus().filter((w) => w.spent >= w.limit * 0.8);
  if (!st.length) return '';
  return `<div class="alert ${st.some((w) => w.spent >= w.limit) ? '' : 'warn'}" data-action="weekly"><b>⏱️ Esta semana</b> <span class="small muted">(desde el ${fmtDate(st[0].monday)})</span>
    ${st.map((w) => `<div class="small">${w.icon} ${esc(w.name)}: <b>${money(w.spent)}</b> de ${money(w.limit)}${w.spent >= w.limit ? ` · <b>te pasaste por ${money(w.spent - w.limit)}</b>` : ` · te quedan ${money(w.limit - w.spent)}`}</div>`).join('')}</div>`;
}

function weeklyCard() {
  const st = weekStatus();
  return `<div class="card"><div class="dhead"><h2>Límites de esta semana</h2><button class="chip" data-action="weekly">Cambiar</button></div>
    ${st.map((w) => { const p = Math.min(1, w.spent / w.limit); return `<div class="goal"><div class="dhead"><span>${w.icon} ${esc(w.name)}</span><span class="small"><b>${money(w.spent)}</b> de ${money(w.limit)}</span></div><div class="ubar"><i class="${p >= 1 ? 'bad' : p >= 0.8 ? 'mid' : 'ok'}" style="width:${Math.max(p * 100, 2)}%"></i></div></div>`; }).join('') || '<p class="small muted">Sin límites. Toca Cambiar.</p>'}
    <p class="small muted">La semana va de lunes a domingo y se cuenta con lo último que importaste. Importa seguido (el aviso del Calendario te ayuda) para que la alerta llegue a tiempo.</p></div>`;
}

function openWeekly() {
  const lim = state.settings.weekly || {};
  $('#sheetPanel').innerHTML = `<h3>Límites por semana</h3>
    <p class="small muted">Te aviso al llegar al 80% y cuando te pases. Deja vacío para no limitar.</p>
    ${WEEKLY.map((w) => `<label class="field inline2"><span>${w.icon} ${esc(w.name)}<small class="muted">sugerido: ${money(w.def)}</small></span><input type="number" inputmode="decimal" data-weekly="${w.id}" value="${lim[w.id] ?? (w.id === 'eatout' && lim.eatout === undefined ? w.def : '')}" placeholder="—"></label>`).join('')}
    <button class="btn" data-action="saveweekly">Guardar</button>
    <button class="btn secondary" data-close>Cancelar</button>`;
  $('#sheet').classList.remove('hidden');
}

async function saveWeekly() {
  const w = {};
  document.querySelectorAll('[data-weekly]').forEach((i) => { w[i.dataset.weekly] = i.value === '' ? 0 : Number(i.value) || 0; });
  state.settings.weekly = w;
  await saveSettings();
  closeSheet(); render();
  toast('Límites guardados');
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
  if (act === 'editcard') openCardEditor(e.target.closest('[data-action]').dataset.id);
  if (act === 'savecard') saveCard();
  if (act === 'addloan') openCardEditor(null, e.target.closest('[data-action]').dataset.pay);
  if (act === 'delloan') deleteLoan(e.target.closest('[data-action]').dataset.id);
  if (act === 'ics') downloadICS();
  if (act === 'goto') { state.view = e.target.closest('[data-action]').dataset.view; render(); window.scrollTo(0, 0); }
  if (act === 'strategy') { state.settings.strategy = e.target.closest('[data-action]').dataset.s; db.kvSet('settings', state.settings); renderAsesoria(); }
  if (act === 'saveincome') saveIncome();
  const ds = e.target.closest('[data-action]')?.dataset || {};
  if (act === 'lump') { state.lump = Number($('#f-lump').value) || 0; renderAsesoria(); }
  if (act === 'applylump') applyLump();
  if (act === 'editgoal') openGoal(Number(ds.i));
  if (act === 'savegoal') saveGoal(Number(ds.i), false);
  if (act === 'delgoal') saveGoal(Number(ds.i), true);
  if (act === 'budgets') openBudgets();
  if (act === 'suggestbudget') suggestBudgets();
  if (act === 'savebudget') saveBudgets();
  if (act === 'reminder') { state.settings.reminder = ds.f; saveSettings(); renderCuentas(); }
  if (act === 'icsupdate') downloadUpdateICS();
  if (act === 'drivebackup') driveBackup(false);
  if (act === 'driverestore') driveRestore();
  if (act === 'saveclient') saveClientId();
  if (act === 'weekly') openWeekly();
  if (act === 'saveweekly') saveWeekly();
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
  if (e.target.id === 'oneOff') {
    const t = state.txns.find((x) => x.id === $('#sheetPanel').dataset.id);
    if (t) { t.oneOff = e.target.checked; db.putMany('txns', [t]); scheduleBackup(); toast(t.oneOff ? 'Marcado como gasto de una sola vez' : 'Ya cuenta en el mes típico'); }
  }
  if (e.target.id === 'isFixed') {
    state.fixedOv[e.target.dataset.key] = e.target.checked;
    db.kvSet('fixed', state.fixedOv);
    toast(e.target.checked ? 'Marcado como gasto fijo' : 'Ya no es gasto fijo');
    render();
  }
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
  const back = gdrive.handleRedirect();
  await load();
  render();
  if (back?.error) toast(back.error === 'access_denied' ? 'No se dio permiso en Google' : 'No se pudo conectar con Google');
  else if (back?.ok) { state.view = 'cuentas'; render(); if (back.pending === 'restore') driveRestore(); else driveBackup(false); }
  else if (!Object.keys(state.debtHist).length) snapshotDebt();
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
})();
