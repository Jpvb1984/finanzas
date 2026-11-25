// Categorías y reglas automáticas.
// kind: 'gasto' cuenta como gasto, 'ingreso' como ingreso,
// 'mov' es un movimiento entre tus propias cuentas (no cuenta en ninguno).

export const CATEGORIES = [
  { id: 'super', name: 'Supermercado', icon: '🛒', color: '#2f9e44', kind: 'gasto' },
  { id: 'resto', name: 'Restaurantes y café', icon: '🍔', color: '#e8590c', kind: 'gasto' },
  { id: 'delivery', name: 'Comida a domicilio', icon: '🛵', color: '#f08c00', kind: 'gasto' },
  { id: 'auto', name: 'Gasolina y auto', icon: '⛽', color: '#1971c2', kind: 'gasto' },
  { id: 'parking', name: 'Estacionamiento y multas', icon: '🅿️', color: '#4263eb', kind: 'gasto' },
  { id: 'transporte', name: 'Transporte', icon: '🚕', color: '#3b5bdb', kind: 'gasto' },
  { id: 'vivienda', name: 'Vivienda', icon: '🏠', color: '#5f3dc4', kind: 'gasto' },
  { id: 'servicios', name: 'Teléfono e internet', icon: '📱', color: '#7048e8', kind: 'gasto' },
  { id: 'suscripciones', name: 'Suscripciones', icon: '🔁', color: '#9c36b5', kind: 'gasto' },
  { id: 'salud', name: 'Salud y farmacia', icon: '💊', color: '#c2255c', kind: 'gasto' },
  { id: 'personal', name: 'Cuidado personal', icon: '💈', color: '#d6336c', kind: 'gasto' },
  { id: 'compras', name: 'Compras y ropa', icon: '🛍️', color: '#f76707', kind: 'gasto' },
  { id: 'hogar', name: 'Hogar y tiendas', icon: '🧰', color: '#a61e4d', kind: 'gasto' },
  { id: 'alcohol', name: 'Licor', icon: '🍷', color: '#862e9c', kind: 'gasto' },
  { id: 'seguros', name: 'Seguros', icon: '🛡️', color: '#495057', kind: 'gasto' },
  { id: 'prestamos', name: 'Préstamos', icon: '🏦', color: '#343a40', kind: 'gasto' },
  { id: 'intereses', name: 'Intereses y cargos', icon: '⚠️', color: '#e03131', kind: 'gasto' },
  { id: 'donaciones', name: 'Iglesia y donaciones', icon: '🙏', color: '#0c8599', kind: 'gasto' },
  { id: 'tramites', name: 'Trámites y gobierno', icon: '📄', color: '#5c940d', kind: 'gasto' },
  { id: 'servprof', name: 'Servicios profesionales', icon: '💼', color: '#087f5b', kind: 'gasto' },
  { id: 'efectivo', name: 'Retiro de efectivo', icon: '💵', color: '#868e96', kind: 'gasto' },
  { id: 'otros', name: 'Otros gastos', icon: '❓', color: '#adb5bd', kind: 'gasto' },
  { id: 'sueldo', name: 'Sueldo', icon: '💰', color: '#2b8a3e', kind: 'ingreso' },
  { id: 'reembolsos', name: 'Reembolsos', icon: '↩️', color: '#37b24d', kind: 'ingreso' },
  { id: 'otrosingresos', name: 'Otros ingresos', icon: '➕', color: '#69db7c', kind: 'ingreso' },
  { id: 'pagotarjeta', name: 'Pago de tarjeta', icon: '💳', color: '#adb5bd', kind: 'mov' },
  { id: 'transferencia', name: 'Transferencias', icon: '🔄', color: '#adb5bd', kind: 'mov' },
  { id: 'inversion', name: 'Ahorro e inversión', icon: '📈', color: '#1098ad', kind: 'mov' },
];

export const CAT = Object.fromEntries(CATEGORIES.map((c) => [c.id, c]));

// Reglas en orden: la primera que coincide gana.
// Se buscan en el texto completo del movimiento, en minúsculas.
const RULES = [
  // Pagos a tarjetas y movimientos entre cuentas
  [/crd\. card bill payment|autopay pymt|scotiabank transit|ctfs\.com\/payments|^payment |payment - thank you|bill payment capital one|miscellaneous payment american express|bill payment american express|bill payment (canadian tire|ctfs)/, 'pagotarjeta'],
  [/interac e-transfer|e-transfer|customer transfer|transfer to|transfer from/, 'transferencia'],
  [/wealthsimple|questrade|tfsa|rrsp|investment/, 'inversion'],
  // Intereses, comisiones y cargos del banco
  [/interest charge|monthly fees|service charge|overdraft|\bodp\b|annual fee|\bnsf\b|credit card protec|balance protection|late fee/, 'intereses'],
  [/loan payment|loans /, 'prestamos'],
  // Ingresos
  [/payroll|salary|pay ?roll|direct dep/, 'sueldo'],
  [/claim ins|manulife|refund|reembolso|correction/, 'reembolsos'],
  [/abm deposit|deposit/, 'otrosingresos'],
  // Gastos
  [/rent\/lease|strata|\brent\b/, 'vivienda'],
  [/insurance corporation of bc|icbc|optiom|apollo insurance|insurance/, 'seguros'],
  [/costco|superstore|safeway|save.?on|no frills|walmart|t&t|whole foods|sobeys|freshco|grocery/, 'super'],
  [/uber.*eats|ubereats|doordash|skip ?the ?dishes|fantuan/, 'delivery'],
  [/mcdonald|tim hortons|starbucks|subway|pizza|thai express|boston pizza|flatbread|restaura|cafe|coffee|sushi|burger|a&w|wendy|kfc|dairy queen|chipotle|coke_|nayax|g&a robot/, 'resto'],
  [/\bchv\d|chevr|chevron|petro|esso|shell |husky|\bmobil\b|co-op gas|7-eleven fuel|car wash/, 'auto'],
  [/impark|paybyphone|indigo park|easypark|parking|city of van.*(bylaw|rpp)|bylaw fine|city of vancouver/, 'parking'],
  [/uber canada\/ubertrip|uber trip|lyft|translink|compass|evo car|modo/, 'transporte'],
  [/spotify|youtube|netflix|disney|apple\.com\/bill|google \*|anthropic|openai|chatgpt|smartsheet|uber \*one|uberonemem|one membership|amazon prime|koho/, 'suscripciones'],
  [/telus|koodo|rogers|fido|\bbell\b|freedom mobile|shaw|tsi internet|internet/, 'servicios'],
  [/pocketpills|pharma|london drugs|shoppers drug|rexall|health|rehab|dental|braces|physio|clinic|medical/, 'salud'],
  [/barber|salon|hair|\bspa\b/, 'personal'],
  [/h ?& ?m |hm richmond|mountain warehouse|winners|old navy|uniqlo|zara|marshalls|sport chek|amazon|affirm|daiso|dollarama|staples|best buy/, 'compras'],
  [/canadian tire|home depot|ikea|rona|home stores|\blowe/, 'hogar'],
  [/liquor|bcl |bc liquor|wine|beer/, 'alcohol'],
  [/charity|church|parish|kofc|donation/, 'donaciones'],
  [/embassy|driver services|passport|service canada|city of|recreatio|government|csra/, 'tramites'],
  [/accounti|tax and|lawyer|notary/, 'servprof'],
  [/abm withdrawal|\batm\b/, 'efectivo'],
];

// Clave del comercio para recordar las correcciones del usuario.
export function merchantKey(desc) {
  return String(desc || '')
    .toLowerCase()
    .replace(/^(fpos|apos|opos)\s+/, '')
    .replace(/^(sq|tst|sp|nbx|pp)\s?\*\s?/, '')
    .replace(/[0-9#*\-_.,/'()]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .split(' ')
    .slice(0, 3)
    .join(' ');
}

export function categorize(txn, userRules = {}) {
  const key = merchantKey(txn.desc);
  if (userRules[key]) return userRules[key];
  const text = `${txn.raw || ''} ${txn.desc || ''}`.toLowerCase();
  for (const [re, cat] of RULES) if (re.test(text)) return cat;
  return txn.amount < 0 ? 'otrosingresos' : 'otros';
}
