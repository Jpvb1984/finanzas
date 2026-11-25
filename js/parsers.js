// Lectores de estados de cuenta. Cada lector recibe datos ya leídos
// (filas de Excel, texto CSV o líneas de PDF) y devuelve:
//   { account: {id, name, kind: 'credito'|'debito'}, txns: [...], meta: {...} }
// En cada movimiento, amount > 0 significa dinero que sale (gasto o cargo)
// y amount < 0 significa dinero que entra (pago, ingreso o reembolso).

const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

const pad = (n) => String(n).padStart(2, '0');
const clean = (s) => (s == null ? '' : String(s)).replace(/\s+/g, ' ').trim();
const num = (v) => {
  if (typeof v === 'number') return v;
  const s = clean(v).replace(/[$,\s]/g, '');
  if (!s) return NaN;
  return Number(s.replace(/^\((.*)\)$/, '-$1'));
};

export function toISODate(v, fallbackYear) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
  if (typeof v === 'number') {
    // número de serie de Excel
    const d = new Date(Math.round((v - 25569) * 86400000));
    return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  }
  const s = clean(v);
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
  m = s.match(/^(\d{1,2}) ([A-Za-z]{3})[a-z]* (\d{4})$/); // 22 Nov 2026
  if (m) return `${m[3]}-${pad(MONTHS[m[2].toLowerCase()])}-${pad(m[1])}`;
  m = s.match(/^([A-Za-z]{3})[a-z]*\.? (\d{1,2}),? (\d{4})$/); // Nov 22, 2026
  if (m) return `${m[3]}-${pad(MONTHS[m[1].toLowerCase()])}-${pad(m[2])}`;
  m = s.match(/^([A-Za-z]{3})[a-z]*\.? (\d{1,2})$/); // Nov 22 (sin año)
  if (m && fallbackYear) {
    const mo = MONTHS[m[1].toLowerCase()];
    const y = typeof fallbackYear === 'function' ? fallbackYear(mo) : fallbackYear;
    return `${y}-${pad(mo)}-${pad(m[2])}`;
  }
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // 11/22/2026 (mes/día)
  if (m) return `${m[3]}-${pad(m[1])}-${pad(m[2])}`;
  return null;
}

const last4 = (s) => {
  const d = String(s || '').replace(/\D/g, '');
  return d.slice(-4);
};

// Nombre legible a partir del nombre del archivo de Scotiabank:
// "Scotia_Momentum_Visa_Infinite__Card_1234_010126.xlsx" -> "Scotia Momentum Visa Infinite Card 1234"
function nameFromFile(fileName) {
  return clean(String(fileName || '').replace(/\.[^.]+$/, '').replace(/_\d{6}$/, '').replace(/_+/g, ' '));
}

// ---------- Excel ----------

export function parseWorkbook(sheets, fileName) {
  // sheets: [{name, rows: [[...], ...]}]
  for (const sh of sheets) {
    const headerIdx = sh.rows.findIndex((r) => r && r.some((c) => clean(c) === 'Date') && r.some((c) => clean(c) === 'Description'));
    if (headerIdx < 0) continue;
    const header = sh.rows[headerIdx].map(clean);
    if (header.includes('Date Processed')) return parseAmex(sh.rows, headerIdx, header);
    if (header.includes('Type of Transaction')) return parseScotia(sh.rows, headerIdx, header, fileName);
  }
  throw new Error('No reconozco este archivo de Excel.');
}

function parseScotia(rows, headerIdx, header, fileName) {
  const col = (n) => header.indexOf(n);
  const cDate = col('Date'), cDesc = col('Description'), cSub = col('Sub-description'), cStatus = col('Status'),
    cAmt = col('Amount'), cCard = col('Associated card number'), cBal = col('Balance');
  const isCredit = cCard >= 0;
  const body = rows.slice(headerIdx + 1).filter((r) => r && r[cDate]);
  const fileDigits = (String(fileName || '').match(/_(\d{4})_\d{6}/) || [])[1];
  const digits = (isCredit && body.length && last4(body[0][cCard])) || fileDigits || '0000';
  const account = { id: `scotia-${digits}`, name: nameFromFile(fileName) || `Scotiabank ${digits}`, kind: isCredit ? 'credito' : 'debito', bank: 'Scotiabank' };
  const txns = [];
  let skippedPending = 0;
  for (const r of body) {
    if (cStatus >= 0 && clean(r[cStatus]).toLowerCase() === 'pending') { skippedPending++; continue; }
    const amt = num(r[cAmt]);
    if (!isFinite(amt) || amt === 0) continue;
    const desc = clean(r[cDesc]);
    const sub = cSub >= 0 ? clean(r[cSub]) : '';
    // En la cuenta de cheques el comercio real está en Sub-description.
    const merchant = !isCredit && sub ? sub.replace(/^[AFO]pos\s+/i, '') : desc;
    txns.push({
      date: toISODate(r[cDate]),
      desc: merchant || desc,
      raw: `${desc} ${sub}`.trim(),
      amount: isCredit ? amt : -amt,
    });
  }
  const meta = {};
  if (cBal >= 0 && body.length) meta.balance = num(body[0][cBal]);
  return { account, txns, meta: { ...meta, skippedPending } };
}

function parseAmex(rows, headerIdx, header) {
  const title = rows.slice(0, headerIdx).flat().map(clean).find((c) => /American Express/i.test(c)) || 'American Express';
  const acctCell = rows.slice(0, headerIdx).flat().map(clean).find((c) => /^X{3,}/i.test(c)) || '';
  const digits = acctCell.replace(/\D/g, '').slice(-5) || '0000';
  const cardName = clean(title.split('/')[0]);
  const account = { id: `amex-${digits}`, name: cardName, kind: 'credito', bank: 'American Express' };
  const period = title.match(/(\d{1,2} \w{3} \d{4}) to (\d{1,2} \w{3} \d{4})/);
  const meta = {};
  if (period) { meta.periodStart = toISODate(period[1]); meta.periodEnd = toISODate(period[2]); meta.cutDay = Number(meta.periodEnd.slice(8)); }
  const c = (n) => header.indexOf(n);
  const cDate = c('Date'), cDesc = c('Description'), cAmt = c('Amount'), cMer = c('Merchant');
  const txns = [];
  for (const r of rows.slice(headerIdx + 1)) {
    if (!r || !r[cDate]) continue;
    const date = toISODate(r[cDate]);
    const amt = num(r[cAmt]);
    if (!date || !isFinite(amt) || amt === 0) continue;
    const desc = clean(r[cDesc]);
    txns.push({ date, desc, raw: `${desc} ${cMer >= 0 ? clean(r[cMer]) : ''}`.trim(), amount: amt });
  }
  return { account, txns, meta };
}

// ---------- CSV ----------

export function parseCSVText(text) {
  const rows = [];
  let row = [], field = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; }
      else field += ch;
    } else if (ch === '"') q = true;
    else if (ch === ',') { row.push(field); field = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else field += ch;
  }
  if (field || row.length) { row.push(field); rows.push(row); }
  return rows;
}

export function parseCSV(text, fileName) {
  const rows = parseCSVText(text.replace(/^﻿/, ''));
  const headerIdx = rows.findIndex((r) => r.map((c) => clean(c).toUpperCase()).includes('TRANSACTION DATE'));
  if (headerIdx >= 0 && /MY ACCOUNT TRANSACTIONS/i.test(rows[0]?.[0] || '')) return parseCanadianTire(rows, headerIdx);
  throw new Error('No reconozco este archivo CSV.');
}

function parseCanadianTire(rows, headerIdx) {
  const H = rows[headerIdx].map((c) => clean(c).toUpperCase());
  const c = (n) => H.indexOf(n);
  const cDate = c('TRANSACTION DATE'), cType = c('TYPE'), cDesc = c('DESCRIPTION'), cAmt = c('AMOUNT'), cCat = c('CATEGORY');
  const account = { id: 'ctfs', name: 'Canadian Tire Mastercard', kind: 'credito', bank: 'Canadian Tire' };
  const meta = {};
  const sumIdx = rows.findIndex((r) => clean(r[0]) === 'Start Date');
  if (sumIdx >= 0 && rows[sumIdx + 1]) meta.balance = num(rows[sumIdx + 1][2]);
  const txns = [];
  for (const r of rows.slice(headerIdx + 1)) {
    if (!r || !r[cDate]) continue;
    const date = toISODate(r[cDate]);
    const amt = num(r[cAmt]);
    if (!date || !isFinite(amt) || amt === 0) continue;
    const desc = clean(r[cDesc]);
    txns.push({ date, desc, raw: `${clean(r[cType])} ${desc} ${cCat >= 0 ? clean(r[cCat]) : ''}`, amount: amt });
  }
  return { account, txns, meta };
}

// ---------- PDF ----------

// Convierte un PDF (pdf.js) en líneas de texto. Las columnas se separan con \t.
export async function pdfToLines(pdfjsLib, data) {
  const doc = await pdfjsLib.getDocument({ data, isEvalSupported: false }).promise;
  const lines = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const content = await page.getTextContent();
    const rowsByY = [];
    for (const it of content.items) {
      if (!it.str || !it.str.trim()) continue;
      const x = it.transform[4], y = it.transform[5];
      let row = rowsByY.find((r) => Math.abs(r.y - y) < 2.5);
      if (!row) { row = { y, items: [] }; rowsByY.push(row); }
      row.items.push({ x, w: it.width, s: it.str });
    }
    rowsByY.sort((a, b) => b.y - a.y);
    for (const r of rowsByY) {
      r.items.sort((a, b) => a.x - b.x);
      let s = '', end = null;
      for (const it of r.items) {
        if (end != null) s += it.x - end > 12 ? '\t' : (it.x - end > 0.5 ? ' ' : '');
        s += it.s;
        end = it.x + it.w;
      }
      lines.push(s.replace(/ {2,}/g, '\t').trim());
    }
  }
  return lines;
}

export function parsePDF(lines, fileName, today = new Date()) {
  const text = lines.join('\n');
  if (/Capital One/i.test(text) && /Statement Period/i.test(text)) return parseCapOneStatement(lines, text);
  if (/Posted Transactions Since Your Last Statement/i.test(text)) return parseCapOneActivity(lines, text, today);
  throw new Error('No reconozco este PDF.');
}

const AMT_RE = /(-\s?)?\$([\d,]+\.\d{2})\s*$/;

function parseCapOneStatement(lines, text) {
  const digits = (text.match(/ending in (\d{4})/i) || [])[1] || '0000';
  const account = { id: `capone-${digits}`, name: `Capital One Mastercard ${digits}`, kind: 'credito', bank: 'Capital One' };
  const per = text.match(/Statement Period:\s*(\w{3} \d{1,2}, \d{4})\s*-\s*(\w{3} \d{1,2}, \d{4})/i);
  const meta = {};
  let yearFor = () => new Date().getFullYear();
  if (per) {
    meta.periodStart = toISODate(per[1]);
    meta.periodEnd = toISODate(per[2]);
    meta.cutDay = Number(meta.periodEnd.slice(8));
    const [sy, sm] = meta.periodStart.split('-').map(Number);
    const ey = Number(meta.periodEnd.slice(0, 4));
    // Una compra del 13 de oct puede aparecer en un estado que empieza el 14 de oct.
    yearFor = (mo) => (sy !== ey ? (mo >= sm - 1 ? sy : ey) : sy);
  }
  const due = text.match(/Payment Due Date[\s\S]{0,200}?(\w{3} \d{1,2}, \d{4})/i) || text.match(/Due Date[\s\S]{0,120}?(\w{3} \d{1,2}, \d{4})/i);
  if (due) meta.dueDate = toISODate(due[1]);
  const nb = text.match(/New Balance\s*=\s*\$([\d,]+\.\d{2})/i) || text.match(/New Balance[\s\S]{0,80}?\$([\d,]+\.\d{2})/i);
  const mp = text.match(/\$[\d,]+\.\d{2}\s+\$([\d,]+\.\d{2})\s+\w{3} \d{1,2}, \d{4}/);
  if (mp) meta.minPayment = num(mp[1]);
  if (nb) meta.statementBalance = num(nb[1]);

  const txns = [];
  const re = /^([A-Z][a-z]{2} \d{1,2})\s+([A-Z][a-z]{2} \d{1,2})\s+(.+?)\s+(-\s?)?\$([\d,]+\.\d{2})$/;
  for (const raw of lines) {
    const line = raw.replace(/\t/g, '  ');
    const m = line.match(re);
    if (!m) continue;
    const desc = clean(m[3]);
    if (/^Total/i.test(desc)) continue;
    const amt = num(m[5]) * (m[4] ? -1 : 1);
    txns.push({ date: toISODate(m[1], yearFor), desc, raw: desc, amount: amt });
  }
  return { account, txns, meta };
}

function parseCapOneActivity(lines, text, today) {
  const digits = (text.match(/Account Ending in\s*\W*(\d{4})/i) || [])[1] || '0000';
  const account = { id: `capone-${digits}`, name: `Capital One Mastercard ${digits}`, kind: 'credito', bank: 'Capital One' };
  const ty = today.getFullYear(), tm = today.getMonth() + 1;
  const yearFor = (mo) => (mo > tm ? ty - 1 : ty);
  const txns = [];
  for (const line of lines) {
    const m = line.match(/^([A-Z][a-z]{2} \d{1,2})\t(.+)$/);
    if (!m) continue;
    const am = m[2].match(AMT_RE);
    if (!am) continue;
    const cols = m[2].replace(AMT_RE, '').split('\t').map(clean).filter(Boolean);
    // columnas: descripción, categoría, tarjeta
    const desc = cols[0] || '';
    const cat = cols.length >= 3 ? cols[1] : '';
    txns.push({ date: toISODate(m[1], yearFor), desc, raw: `${desc} ${cat}`.trim(), amount: num(am[2]) * (am[1] ? -1 : 1) });
  }
  return { account, txns, meta: {} };
}

// ---------- Identificadores ----------

function fnv(s) {
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 0x01000193); }
  return (h >>> 0).toString(16).padStart(8, '0');
}

// Da a cada movimiento un id estable, para que importar el mismo archivo
// dos veces (o dos archivos que se traslapan) no duplique nada.
export function assignIds(result) {
  const seen = new Map();
  for (const t of result.txns) {
    const key = `${result.account.id}|${t.date}|${t.amount.toFixed(2)}|${t.desc.toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 12)}`;
    const n = (seen.get(key) || 0) + 1;
    seen.set(key, n);
    t.id = fnv(key) + fnv(key + '#' + n);
    t.account = result.account.id;
  }
  return result;
}
