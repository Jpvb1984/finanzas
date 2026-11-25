import * as db from './db.js';
import { parseWorkbook, parseCSV, parsePDF, pdfToLines, assignIds } from './parsers.js';
import { CATEGORIES, CAT, categorize, merchantKey } from './categories.js';
import { cardInfo, fixedExpenses, monthlyPicture, simulate, minimumsOnly, buildICS, isDebtAccount, todayISO, daysBetween } from './debts.js';

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
};

// ---------- carga ----------

async function load() {
  [state.txns, state.accounts, state.rules] = await Promise.all([db.getAll('txns'), db.getAll('accounts'), db.kvGet('rules')]);
  state.rules = state.rules || {};
  state.fixedOv = (await db.kvGet('fixed')) || {};
  state.settings = (await db.kvGet('settings')) || {};
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
  el.innerHTML = `${upcomingBanner()}
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
    const data = JSON.parse(await file.text());
    if (data.app !== 'finanzas' || !Array.isArray(data.txns)) throw new Error('Este archivo no es una copia de Finanzas.');
    await db.putMany('txns', data.txns);
    await db.putMany('accounts', data.accounts || []);
    await db.kvSet('rules', { ...(state.rules || {}), ...(data.rules || {}) });
    if (data.fixed) await db.kvSet('fixed', { ...state.fixedOv, ...data.fixed });
    if (data.settings) await db.kvSet('settings', { ...state.settings, ...data.settings });
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

function renderAsesoria() {
  const el = $('#view-asesoria');
  if (!state.txns.length) { el.innerHTML = emptyState(); return; }
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
  const libre2 = libre + recorte + extraInv;
  const quietas = cards.filter((c) => c.status === 'quieta' && c.kind !== 'linea');
  const line = cards.find((c) => c.kind === 'linea');
  const first = plan.order?.[0];
  const missing = debts.filter((d) => d.aprMissing).map((d) => d.name).concat(cards.filter((c) => c.balanceSrc === 'estimado' || c.balance == null).map((c) => shortCard(c.name)));
  const topVar = Object.entries(pic.varByCat).sort((a, b) => b[1] - a[1]).slice(0, 3);

  const steps = [];
  if (libre < 0) steps.push(`<b>Hoy sale más de lo que entra: te faltan ${money(-libre)} al mes.</b> Antes de cualquier plan hay que recortar. Donde más se va el dinero variable: ${topVar.map(([k, v]) => `${CAT[k].name} (${money(v)})`).join(', ')}.`);
  steps.push(`<b>Paga siempre el mínimo de todas a tiempo</b> (${money(mins)} al mes). Un pago tarde suma cargos, sube la tasa y baja tu puntaje de crédito. Los avisos del Calendario te ayudan con esto.`);
  if (first && libre > 0) steps.push(`<b>Todo lo que sobra (${money(libre)} al mes) va a ${esc(first.name)}</b>, ${strategy === 'avalancha' ? `porque es la de tasa más alta (${pct(first.apr)})` : 'porque es la de saldo más chico'}. Cuando la termines, ese dinero pasa a la siguiente de la lista.`);
  if (quietas.length) steps.push(`<b>Deja quietas ${quietas.map((c) => esc(shortCard(c.name))).join(', ')}</b>. Mientras tengan saldo, cada compra cobra interés desde el primer día. Para el día a día usa débito.`);
  if (extraInv) steps.push(`<b>Considera pausar lo que mandas a inversión (${money(extraInv)} al mes)</b> y usarlo en deudas. Ninguna inversión normal te da ${pct(maxApr)} seguro, que es lo que te ahorras al pagar esa deuda. La excepción es si tu empleador iguala tu aporte.`);
  if (recorte > 20) steps.push(`<b>Recorta a la mitad comida a domicilio, restaurantes, licor y compras</b>: liberas unos ${money(recorte)} al mes.`);
  if (line && maxApr > (line.apr ?? 99)) steps.push(`<b>Pregunta en Scotiabank</b> si puedes pasar la deuda de la tarjeta más cara a tu ScotiaLine (${pct(line.apr)}). Pagarías menos interés, pero solo sirve si después no vuelves a usar esa tarjeta.`);
  else if (line && line.apr == null) steps.push(`<b>Escribe la tasa de tu ScotiaLine</b> en Deudas. Si es más baja que la de tus tarjetas, conviene usarla para pagar la tarjeta más cara.`);
  steps.push(`<b>Guarda un colchón pequeño</b> (unos $1,000 en tu cuenta de cheques) para emergencias, así no vuelves a usar las tarjetas cuando algo pase.`);

  el.innerHTML = `
    ${missing.length ? `<div class="alert warn" data-action="goto" data-view="deudas">Faltan datos reales de: ${esc([...new Set(missing)].join(', '))}. El plan usa estimaciones hasta que los escribas en <b>Deudas</b>.</div>` : ''}
    <div class="card"><h2>Tu mes típico</h2>
      <div class="grid2 big">
        <div>Entra</div><div class="r pos">${money(income)}</div>
        <div>Gastos fijos</div><div class="r">−${money(fijos)}</div>
        <div>Gastos del día a día</div><div class="r">−${money(pic.variable)}</div>
        <div>Mínimos de deudas</div><div class="r">−${money(mins)}</div>
        <div class="tot">Queda libre</div><div class="r tot ${libre >= 0 ? 'pos' : 'neg'}">${money(libre)}</div>
      </div>
      <p class="small muted">Calculado con tus últimos 30 días (${fmtDate(pic.start)} a ${fmtDate(pic.end)})${pic.incomeNote ? `, ${esc(pic.incomeNote)}` : ''}. Intereses e inversión van aparte.</p>
      <label class="field inline"><span>¿Tu ingreso mensual es otro?</span><input id="f-income" type="number" inputmode="decimal" value="${state.settings.income ?? ''}" placeholder="${Math.round(pic.income)}"><button class="btn sm" data-action="saveincome">OK</button></label>
    </div>
    <div class="card hl"><h2>Qué hacer</h2><ol class="steps">${steps.map((s) => `<li>${s}</li>`).join('')}</ol></div>
    ${debts.length ? `<div class="card"><h2>Tu camino para salir de deudas</h2>
      <div class="seg"><button data-action="strategy" data-s="avalancha" class="${strategy === 'avalancha' ? 'sel' : ''}">Pagar menos intereses</button><button data-action="strategy" data-s="bolaDeNieve" class="${strategy === 'bolaDeNieve' ? 'sel' : ''}">Ver avances rápido</button></div>
      ${plan.feasible && libre >= 0 ? `<p>Con este plan terminas en <b>${yearsMonths(plan.months)}</b> (${monthLabelFromNow(plan.months)}) y pagas <b>${money(plan.totalInterest)}</b> de intereses.</p>
        <ol class="order">${plan.order.map((d) => `<li><b>${esc(d.name)}</b> · ${money(d.balance)} al ${pct(d.apr)}${d.aprMissing ? '<span class="est">*</span>' : ''}<div class="small muted">Pagada en ${monthLabelFromNow(d.paidOffMonth)}</div></li>`).join('')}</ol>`
        : `<p class="neg">Con lo que entra hoy no alcanza para bajar las deudas más allá de los mínimos. Primero hay que cerrar ese hueco.</p>
          <p class="small">Cuando sobre dinero, este es el orden para atacarlas:</p>
          <ol class="order">${plan.order.map((d) => `<li><b>${esc(d.name)}</b> · ${money(d.balance)} al ${pct(d.apr)}${d.aprMissing ? '<span class="est">*</span>' : ''}</li>`).join('')}</ol>`}
      <p class="small">Pagando solo los mínimos: <b>${yearsMonths(only.months)}</b> y ${money(only.totalInterest)} de intereses.</p>
      ${(recorte > 20 || extraInv) ? (libre2 > 0 && plan2.feasible
        ? `<p class="small">Si recortas lo de arriba${extraInv ? ' y pausas la inversión' : ''}, te quedan ${money(libre2)} libres al mes: terminas en <b>${yearsMonths(plan2.months)}</b> y te ahorras ${money(Math.max(0, (plan.feasible && libre >= 0 ? plan.totalInterest : only.totalInterest) - plan2.totalInterest))} en intereses.</p>`
        : `<p class="small">Aun recortando lo de arriba${extraInv ? ' y pausando la inversión' : ''} (${money(recorte + extraInv)} al mes) faltarían ${money(-libre2)} al mes. Revisa los gastos fijos más grandes.</p>`) : ''}
    </div>` : ''}
    <p class="small muted">Esto es orientación general calculada con tus números, no asesoría financiera profesional. Antes de mover deudas entre productos, confirma tasas y condiciones con tu banco.</p>`;
}

async function saveIncome() {
  const v = Number($('#f-income').value);
  if ($('#f-income').value === '' ) delete state.settings.income; else if (isFinite(v) && v > 0) state.settings.income = v;
  await db.kvSet('settings', state.settings);
  renderAsesoria();
  toast('Ingreso actualizado');
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
  await load();
  render();
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
})();
