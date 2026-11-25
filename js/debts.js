// Cálculos de deudas, gastos fijos y plan de pago. Todo se calcula en el teléfono.
import { CAT, merchantKey } from './categories.js';
import { shiftISO } from './parsers.js';

export const GRACE_DAYS = 21; // días típicos entre el corte y la fecha de pago en Canadá

// Tasas típicas para cuando el archivo no la trae. Se muestran como "típica, confírmala".
const TYPICAL_APR = [
  [/momentum visa infinite/i, 20.99],
  [/cobalt/i, 21.99],
  [/canadian tire|triangle/i, 21.99],
  [/capital one/i, 21.9],
  [/scotia.*(visa|mastercard|amex)/i, 20.99],
];

export const todayISO = (d = new Date()) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
export const daysBetween = (a, b) => Math.round((new Date(b + 'T12:00:00Z') - new Date(a + 'T12:00:00Z')) / 86400000);

function addMonthsISO(iso, n, day) {
  const [y, m] = iso.split('-').map(Number);
  const d = new Date(Date.UTC(y, m - 1 + n, 1));
  const last = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0)).getUTCDate();
  d.setUTCDate(Math.min(day ?? Number(iso.slice(8)), last));
  return d.toISOString().slice(0, 10);
}

// Último corte (en o antes de hoy) para un día de corte dado.
function lastCutFor(cutDay, today) {
  let c = addMonthsISO(today, 0, cutDay);
  if (c > today) c = addMonthsISO(today, -1, cutDay);
  return c;
}

export const isDebtAccount = (a) => a.kind === 'credito' || a.kind === 'linea' || a.kind === 'prestamo';
export const isLine = (a) => a.kind === 'linea' || /line of credit|scotialine|línea/i.test(a.name || '');

// Junta lo que dicen los archivos, lo que escribió el usuario y estimaciones.
export function cardInfo(a, txns, today = todayISO()) {
  const u = a.user || {};
  const mine = txns.filter((t) => t.account === a.id);
  const info = { id: a.id, name: a.name, kind: isLine(a) ? 'linea' : a.kind, notes: [] };

  // Intereses cobrados en los últimos 45 días
  const interestTx = mine.filter((t) => t.cat === 'intereses' && /interest|interés/i.test(t.raw || t.desc) && daysBetween(t.date, today) <= 45);
  info.interestRecent = interestTx.reduce((s, t) => s + t.amount, 0);
  const lastInterest = interestTx.map((t) => t.date).sort().pop();

  // Tasa anual
  if (u.apr != null) { info.apr = u.apr; info.aprSrc = 'tuyo'; }
  else if (a.apr != null) { info.apr = a.apr; info.aprSrc = 'estado'; }
  else {
    const typ = TYPICAL_APR.find(([re]) => re.test(a.name || ''));
    if (typ && !isLine(a)) { info.apr = typ[1]; info.aprSrc = 'típica'; }
    else { info.apr = null; info.aprSrc = null; }
  }

  // Día de corte y último corte
  info.cutDay = u.cutDay ?? a.cutDay ?? (lastInterest ? Number(lastInterest.slice(8)) : null);
  info.cutSrc = u.cutDay != null ? 'tuyo' : a.cutDay != null ? 'estado' : lastInterest ? 'intereses' : null;
  info.lastCut = info.cutDay ? lastCutFor(info.cutDay, today) : null;

  // Fecha de pago
  let due = null, dueSrc = null;
  if (u.dueDay) { due = addMonthsISO(today, 0, u.dueDay); dueSrc = 'tuyo'; if (due < today) due = addMonthsISO(today, 1, u.dueDay); }
  else if (a.dueDate) { due = a.dueDate; dueSrc = 'estado'; }
  else if (info.lastCut) { due = shiftISO(info.lastCut, GRACE_DAYS); dueSrc = 'estimada'; }
  if (due && due < today && dueSrc === 'estimada') due = shiftISO(addMonthsISO(info.lastCut, 1, info.cutDay), GRACE_DAYS);
  for (let i = 0; due && due < today && i < 24; i++) due = addMonthsISO(due, 1);
  info.dueDate = due; info.dueSrc = dueSrc;
  info.daysToDue = due ? daysBetween(today, due) : null;

  // Límite
  info.limit = u.limit ?? a.limit ?? a.limitApprox ?? null;

  // Saldo actual (lo más reciente gana: lo escrito por ti o lo del archivo)
  const fileBal = a.balance != null ? { v: a.balance, at: a.balanceAt || a.lastImport, src: 'archivo' } : null;
  let fromStatement = null;
  if (a.statementBalance != null && a.statementEnd && a.id.startsWith('capone-')) {
    const after = mine.filter((t) => t.date > a.statementEnd).reduce((s, t) => s + t.amount, 0);
    fromStatement = { v: a.statementBalance + after, at: a.lastImport, src: 'estado + movimientos' };
  }
  const userBal = u.balance != null ? { v: u.balance, at: u.balanceAt, src: 'tuyo' } : null;
  const cands = [fileBal, fromStatement, userBal].filter(Boolean).sort((x, y) => String(y.at).localeCompare(String(x.at)));
  if (cands.length) { info.balance = cands[0].v; info.balanceSrc = cands[0].src; }
  else if (info.interestRecent > 0 && info.apr) {
    // Sin saldo en el archivo: lo estimamos con los intereses del mes
    const purchaseInt = interestTx.filter((t) => /purchase|interest charges$/i.test((t.raw || t.desc).trim())).reduce((s, t) => s + t.amount, 0) || info.interestRecent;
    info.balance = Math.round(purchaseInt / (info.apr / 100 / 12));
    info.balanceSrc = 'estimado';
  } else { info.balance = null; info.balanceSrc = null; }

  info.statementBalance = u.statementBalance ?? a.statementBalance ?? null;
  // Pago mínimo
  if (u.minPayment != null) { info.minPayment = u.minPayment; info.minSrc = 'tuyo'; }
  else if (a.minPayment != null && a.dueDate && a.dueDate >= today) { info.minPayment = a.minPayment; info.minSrc = 'estado'; }
  else if (info.balance) {
    const base = info.statementBalance ?? info.balance;
    info.minPayment = Math.min(base, Math.max(10, Math.round(base * 0.03)));
    info.minSrc = 'estimado';
  } else { info.minPayment = null; info.minSrc = null; }
  if (a.kind === 'prestamo') { info.minPayment = u.payment ?? a.payment ?? info.minPayment; info.minSrc = 'tuyo'; }

  info.util = info.balance != null && info.limit ? info.balance / info.limit : null;
  info.monthlyInterest = info.balance && info.apr ? info.balance * (info.apr / 100 / 12) : null;
  info.carrying = info.interestRecent > 0.5;

  // Semáforo: ¿la uso o la dejo quieta?
  if (a.kind === 'prestamo') { info.status = 'prestamo'; }
  else if (info.kind === 'linea') { info.status = 'quieta'; info.why = 'Es una línea de crédito: úsala solo para pagar deudas más caras, nunca para compras.'; }
  else if (info.carrying) { info.status = 'quieta'; info.why = 'Le estás pagando intereses. Mientras tenga saldo, cada compra nueva cobra interés desde el primer día.'; }
  else if (info.util != null && info.util >= 0.7) { info.status = 'quieta'; info.why = `Está al ${Math.round(info.util * 100)}% del límite. Usarla más baja tu puntaje de crédito.`; }
  else if (info.util != null && info.util >= 0.3) { info.status = 'cuidado'; info.why = `Va al ${Math.round(info.util * 100)}% del límite. Lo sano es usar menos del 30%.`; }
  else {
    info.status = 'usar';
    info.why = info.statementBalance ? 'No vemos intereses. Úsala solo si pagas el saldo completo cada mes.' : 'No vemos intereses. Úsala y págala completa cada mes.';
  }
  // Días de gracia si compras hoy: hasta el pago del próximo corte
  if (info.cutDay) {
    const nextCut = addMonthsISO(info.lastCut, 1, info.cutDay);
    info.graceIfBuyToday = daysBetween(today, nextCut) + GRACE_DAYS;
  }
  return info;
}

// ---------- gastos fijos ----------

const FIXED_CATS = new Set(['vivienda', 'prestamos', 'seguros', 'servicios', 'suscripciones']);

// "Marca" del comercio: la primera palabra con sentido. Junta "Spotify P496A" con "Spotify P4830".
const NOISE = new Set(['google', 'www', 'the', 'canada', 'inc', 'ltd', 'pos', 'purchase', 'apple', 'com']);
export function brand(desc) {
  const words = merchantKey(desc).replace(/[^a-z\s&]/g, ' ').split(/\s+/).filter((w) => w.length >= 3 && !NOISE.has(w));
  return words[0] || merchantKey(desc);
}

export function fixedExpenses(txns, overrides = {}) {
  const groups = [];
  for (const t of txns) {
    if (!(t.amount > 0)) continue;
    const key = merchantKey(t.desc);
    const ov = overrides[key];
    if (ov === false) continue;
    if (!(ov === true || FIXED_CATS.has(t.cat))) continue;
    const b = brand(t.desc);
    // Mismo comercio, misma cuenta y monto parecido (±20%) = el mismo gasto fijo
    const same = (x, y) => x === y || (Math.min(x.length, y.length) >= 4 && (x.startsWith(y) || y.startsWith(x)));
    let g = groups.find((x) => same(x.brand, b) && x.account === t.account && Math.abs(x.items[0].amount - t.amount) <= 0.2 * x.items[0].amount);
    if (!g) { g = { brand: b, account: t.account, keys: new Set(), cat: t.cat, items: [] }; groups.push(g); }
    g.keys.add(key);
    g.items.push(t);
  }
  const out = [];
  for (const g of groups) {
    g.items.sort((a, b) => a.date.localeCompare(b.date));
    const last = g.items[g.items.length - 1];
    const gaps = g.items.slice(1).map((t, i) => daysBetween(g.items[i].date, t.date)).filter((d) => d > 3);
    const med = gaps.length ? gaps.sort((a, b) => a - b)[Math.floor(gaps.length / 2)] : null;
    let freq = 'mensual', perMonth = last.amount;
    if (med && med >= 12 && med <= 16) { freq = 'cada 2 semanas'; perMonth = last.amount * 26 / 12; }
    else if (med && med >= 6 && med <= 8) { freq = 'semanal'; perMonth = last.amount * 52 / 12; }
    out.push({ key: [...g.keys][0], keys: [...g.keys], name: last.desc, cat: last.cat, amount: last.amount, freq, perMonth, day: Number(last.date.slice(8)), lastDate: last.date, account: last.account, dates: g.items.map((t) => t.date) });
  }
  return out.sort((a, b) => b.perMonth - a.perMonth);
}

// ---------- ingresos y gastos típicos ----------

// Mira los últimos 30 días con datos para estimar el mes típico.
export function monthlyPicture(txns, fixed) {
  if (!txns.length) return null;
  const end = txns.map((t) => t.date).sort().pop();
  const start = shiftISO(end, -30);
  const win = txns.filter((t) => t.date > start && t.date <= end);
  const fixedKeys = new Set(fixed.flatMap((f) => f.keys || [f.key]));

  // Sueldo: si llega cada 2 semanas, el mes típico es pago × 26 / 12
  let pay = txns.filter((t) => t.cat === 'sueldo' && t.amount < -100);
  if (!pay.length) {
    // Sin categoría "Sueldo": buscamos depósitos grandes que se repiten del mismo origen
    const cand = new Map();
    for (const t of txns) if (CAT[t.cat]?.kind === 'ingreso' && t.cat !== 'reembolsos' && t.amount < -500) {
      const k = merchantKey(t.desc); if (!cand.has(k)) cand.set(k, []); cand.get(k).push(t);
    }
    pay = [...cand.values()].filter((l) => l.length >= 2).flat();
  }
  const byKey = new Map();
  for (const t of pay) { const k = merchantKey(t.desc); if (!byKey.has(k)) byKey.set(k, []); byKey.get(k).push(t); }
  let income = 0, incomeNote = '';
  for (const [, list] of byKey) {
    list.sort((a, b) => a.date.localeCompare(b.date));
    const gaps = list.slice(1).map((t, i) => daysBetween(list[i].date, t.date));
    const avg = -list.reduce((s, t) => s + t.amount, 0) / list.length;
    if (gaps.length && gaps.every((g) => g >= 12 && g <= 16)) { income += avg * 26 / 12; incomeNote = `sueldo cada 2 semanas de ~${Math.round(avg)}`; }
    else income += -win.filter((t) => list.includes(t)).reduce((s, t) => s + t.amount, 0);
  }
  if (!income) income = -win.filter((t) => CAT[t.cat]?.kind === 'ingreso' && t.cat !== 'reembolsos').reduce((s, t) => s + t.amount, 0);

  let variable = 0, interest = 0, invest = 0;
  const varByCat = {};
  for (const t of win) {
    const c = CAT[t.cat];
    if (!c) continue;
    // Reembolsos (ej. seguro médico) restan del gasto del día a día
    if (t.cat === 'reembolsos' && t.amount < 0) { variable += t.amount; continue; }
    if (t.cat === 'inversion' && t.amount > 0) invest += t.amount;
    if (c.kind !== 'gasto') continue;
    if (t.cat === 'intereses') { interest += t.amount; continue; }
    if (t.cat === 'reembolsos') continue;
    if (fixedKeys.has(merchantKey(t.desc))) continue;
    variable += t.amount;
    varByCat[t.cat] = (varByCat[t.cat] || 0) + t.amount;
  }
  return { start, end, income, incomeNote, variable, interest, invest, varByCat };
}

// ---------- plan para salir de deudas ----------

export function simulate(debts, monthlyPay, strategy = 'avalancha') {
  let ds = debts.filter((d) => d.balance > 0.5).map((d) => ({ ...d, bal: d.balance, paidOffMonth: null, interestPaid: 0 }));
  if (!ds.length) return { months: 0, totalInterest: 0, order: [], feasible: true };
  const order = [...ds].sort((a, b) => strategy === 'bolaDeNieve' ? a.bal - b.bal : (b.apr ?? 0) - (a.apr ?? 0) || a.bal - b.bal);
  let month = 0, totalInterest = 0;
  while (ds.some((d) => d.bal > 0.5) && month < 600) {
    month++;
    for (const d of ds) if (d.bal > 0) { const i = d.bal * ((d.apr ?? 0) / 100 / 12); d.bal += i; d.interestPaid += i; totalInterest += i; }
    let budget = monthlyPay;
    for (const d of ds) if (d.bal > 0) { const p = Math.min(d.bal, d.min); d.bal -= p; budget -= p; }
    if (budget < -0.01 && month === 1) return { months: Infinity, totalInterest: Infinity, order, feasible: false };
    for (const o of order) {
      const d = ds.find((x) => x.id === o.id);
      if (budget <= 0) break;
      if (d.bal > 0) { const p = Math.min(d.bal, budget); d.bal -= p; budget -= p; }
    }
    for (const d of ds) if (d.bal <= 0.5 && d.paidOffMonth == null) { d.bal = 0; d.paidOffMonth = month; }
  }
  if (month >= 600) return { months: Infinity, totalInterest: Infinity, order, feasible: false };
  return { months: month, totalInterest, order: order.map((o) => ds.find((d) => d.id === o.id)), feasible: true };
}

// Solo pagando el mínimo (3% del saldo, mínimo $10) cada mes
export function minimumsOnly(debts) {
  let total = 0, maxMonths = 0;
  for (const d of debts) {
    let bal = d.balance, m = 0;
    if (!(bal > 0)) continue;
    while (bal > 0.5 && m < 600) {
      m++;
      const i = bal * ((d.apr ?? 0) / 100 / 12);
      bal += i; total += i;
      const p = d.fixedPayment ? Math.min(bal, d.min) : Math.min(bal, Math.max(10, bal * 0.03, i + 1));
      bal -= p;
    }
    maxMonths = Math.max(maxMonths, m);
  }
  return { months: maxMonths, totalInterest: total };
}

// ---------- calendario (.ics) ----------

export function buildICS(cards) {
  const pad = (n) => String(n).padStart(2, '0');
  const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+Z$/, 'Z');
  const lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Finanzas//ES', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH', 'X-WR-CALNAME:Pagos de tarjetas'];
  for (const c of cards) {
    if (c.dueDate) {
      const d = c.dueDate.replace(/-/g, '');
      const day = Number(c.dueDate.slice(8));
      lines.push('BEGIN:VEVENT', `UID:pago-${c.id}@finanzas`, `DTSTAMP:${stamp}`, `DTSTART;VALUE=DATE:${d}`,
        `RRULE:FREQ=MONTHLY;BYMONTHDAY=${day > 28 ? -1 : day}`, `SUMMARY:💳 Pagar ${c.name}`,
        `DESCRIPTION:Fecha límite de pago de ${c.name}. Revisa el monto en la app Finanzas.`,
        'BEGIN:VALARM', 'ACTION:DISPLAY', `DESCRIPTION:En 3 días vence ${c.name}`, 'TRIGGER:-P2DT15H', 'END:VALARM',
        'BEGIN:VALARM', 'ACTION:DISPLAY', `DESCRIPTION:Mañana vence ${c.name}`, 'TRIGGER:-PT15H', 'END:VALARM',
        'END:VEVENT');
    }
    if (c.cutDay && c.lastCut) {
      const next = addMonthsISO(c.lastCut, 1, c.cutDay).replace(/-/g, '');
      lines.push('BEGIN:VEVENT', `UID:corte-${c.id}@finanzas`, `DTSTAMP:${stamp}`, `DTSTART;VALUE=DATE:${next}`,
        `RRULE:FREQ=MONTHLY;BYMONTHDAY=${c.cutDay > 28 ? -1 : c.cutDay}`, `SUMMARY:✂️ Corte ${c.name}`,
        'DESCRIPTION:Fecha de corte. Lo que compres después de hoy se paga en el siguiente estado.', 'END:VEVENT');
    }
  }
  lines.push('END:VCALENDAR');
  return lines.join('\r\n');
}
