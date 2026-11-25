// Respaldo en tu propio Google Drive. La app solo puede ver los archivos que ella misma crea
// (permiso drive.file). El archivo se llama "Finanzas - respaldo.json" y lo ves en tu Drive.
const SCOPE = 'https://www.googleapis.com/auth/drive.file';
const FILE_NAME = 'Finanzas - respaldo.json';
const TOKEN_KEY = 'finanzas.gtoken';

const redirectUri = () => location.origin + location.pathname.replace(/index\.html$/, '');

function store() { try { return window.localStorage; } catch { return null; } }

export function getToken() {
  try {
    const t = JSON.parse(store()?.getItem(TOKEN_KEY) || 'null');
    if (t && t.exp > Date.now() + 60000) return t.token;
  } catch {}
  return null;
}

// Va a la pantalla de Google y vuelve a la app con el permiso.
export function connect(clientId, pending) {
  const state = Math.random().toString(36).slice(2);
  try { sessionStorage.setItem('finanzas.gstate', state); sessionStorage.setItem('finanzas.gpending', pending || ''); } catch {}
  try { store()?.setItem('finanzas.gpending', pending || ''); store()?.setItem('finanzas.gstate', state); } catch {}
  const p = new URLSearchParams({
    client_id: clientId,
    redirect_uri: redirectUri(),
    response_type: 'token',
    scope: SCOPE,
    include_granted_scopes: 'true',
    state,
  });
  location.href = `https://accounts.google.com/o/oauth2/v2/auth?${p}`;
}

// Al volver de Google, guarda el permiso y limpia la dirección. Devuelve la acción pendiente.
export function handleRedirect() {
  if (!location.hash.includes('access_token') && !location.hash.includes('error=')) return null;
  const h = new URLSearchParams(location.hash.slice(1));
  history.replaceState(null, '', location.pathname + location.search);
  const expected = (() => { try { return sessionStorage.getItem('finanzas.gstate') || store()?.getItem('finanzas.gstate'); } catch { return null; } })();
  const pending = (() => { try { return sessionStorage.getItem('finanzas.gpending') || store()?.getItem('finanzas.gpending') || ''; } catch { return ''; } })();
  if (h.get('error')) return { error: h.get('error') };
  if (!expected || h.get('state') !== expected) return { error: 'state' };
  const token = h.get('access_token');
  const exp = Date.now() + Number(h.get('expires_in') || 3600) * 1000;
  try { store()?.setItem(TOKEN_KEY, JSON.stringify({ token, exp })); } catch {}
  return { ok: true, pending };
}

async function api(url, opts = {}) {
  const token = getToken();
  if (!token) throw Object.assign(new Error('Falta conectar con Google'), { code: 'auth' });
  const res = await fetch(url, { ...opts, headers: { ...(opts.headers || {}), Authorization: `Bearer ${token}` } });
  if (res.status === 401) { try { store()?.removeItem(TOKEN_KEY); } catch {} throw Object.assign(new Error('El permiso de Google venció'), { code: 'auth' }); }
  if (!res.ok) throw new Error(`Google Drive respondió ${res.status}`);
  return res;
}

async function findFile() {
  const q = encodeURIComponent(`name='${FILE_NAME}' and trashed=false`);
  const r = await api(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,modifiedTime)&orderBy=modifiedTime desc&pageSize=1`);
  const j = await r.json();
  return j.files?.[0] || null;
}

export async function upload(data) {
  const body = JSON.stringify(data);
  const f = await findFile();
  if (f) {
    await api(`https://www.googleapis.com/upload/drive/v3/files/${f.id}?uploadType=media`, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body });
    return f.id;
  }
  const boundary = 'finanzas' + Math.random().toString(36).slice(2);
  const meta = { name: FILE_NAME, mimeType: 'application/json', description: 'Respaldo de la app Finanzas. No lo edites a mano.' };
  const multipart = `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n--${boundary}\r\nContent-Type: application/json\r\n\r\n${body}\r\n--${boundary}--`;
  const r = await api('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', { method: 'POST', headers: { 'Content-Type': `multipart/related; boundary=${boundary}` }, body: multipart });
  return (await r.json()).id;
}

export async function download() {
  const f = await findFile();
  if (!f) return null;
  const r = await api(`https://www.googleapis.com/drive/v3/files/${f.id}?alt=media`);
  return { data: await r.json(), modifiedTime: f.modifiedTime };
}
