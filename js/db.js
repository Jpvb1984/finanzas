// Almacenamiento local en el teléfono (IndexedDB). Nada sale del dispositivo.
const DB_NAME = 'finanzas';
const VERSION = 1;
let dbp;

function open() {
  if (dbp) return dbp;
  dbp = new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains('txns')) db.createObjectStore('txns', { keyPath: 'id' });
      if (!db.objectStoreNames.contains('accounts')) db.createObjectStore('accounts', { keyPath: 'id' });
      if (!db.objectStoreNames.contains('kv')) db.createObjectStore('kv');
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbp;
}

async function tx(store, mode, fn) {
  const db = await open();
  return new Promise((resolve, reject) => {
    const t = db.transaction(store, mode);
    const s = t.objectStore(store);
    let out;
    Promise.resolve(fn(s)).then((v) => (out = v));
    t.oncomplete = () => resolve(out);
    t.onerror = () => reject(t.error);
    t.onabort = () => reject(t.error);
  });
}

const reqP = (r) => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });

export const getAll = (store) => tx(store, 'readonly', (s) => reqP(s.getAll()));
export const putMany = (store, items) => tx(store, 'readwrite', (s) => { for (const it of items) s.put(it); });
export const delMany = (store, keys) => tx(store, 'readwrite', (s) => { for (const k of keys) s.delete(k); });
export const clear = (store) => tx(store, 'readwrite', (s) => s.clear());
export const kvGet = (key) => tx('kv', 'readonly', (s) => reqP(s.get(key)));
export const kvSet = (key, val) => tx('kv', 'readwrite', (s) => s.put(val, key));
