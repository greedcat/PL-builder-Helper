// ─────────────────────────────────────────────────────────────
// BACKEND API CONNECTION
// Talks to the small Node/Express server in server/ which stores
// the shared config document in MongoDB.
// ─────────────────────────────────────────────────────────────
// The builder's own server keeps the shared config now (it has MongoDB);
// the separate config server on Railway is retired.
const DB_DEFAULT_BASE  = location.origin;
const DB_FIELDS        = ['DEST_LIST', 'CARTON_KEYWORDS', 'CBM_KEYWORDS', 'WEIGHT_KEYWORDS', 'NO_NEED_COL'];

// One backend, named here. It used to be editable in the settings panel,
// which only ever offered a way to point the app at nothing.
function getDbBaseUrl() {
  return DB_DEFAULT_BASE.replace(/\/+$/, '');
}

function isDbConfigured() {
  return !!getDbBaseUrl();
}

async function dbCheckConnection() {
  const res = await fetch(getDbBaseUrl() + '/api/health');
  if (!res.ok) throw new Error(`Health check failed (${res.status})`);
  return res.json();
}

// Fetches the shared config document. Returns null if not found.
async function loadConfigFromDB() {
  const res = await fetch(getDbBaseUrl() + '/api/config');
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Failed to load config (${res.status}): ${text}`);
  }
  const doc = await res.json();
  return doc && Object.keys(doc).length ? doc : null;
}

// Lists the shared config server above does not know (it keeps only the
// five in DB_FIELDS), kept by this app's own server instead. Where that
// server has no database yet, a list is kept in this browser, so what was
// saved still works here, and is merged in once the server has one.
const OWN_FIELDS = ['JOB_KEYWORDS'];
const ownLocalKey = key => `pl-builder.${key}`;

function readLocalList(key) {
  try { const v = JSON.parse(localStorage.getItem(ownLocalKey(key)) || 'null'); return Array.isArray(v) ? v : null; }
  catch (_) { return null; }
}
function writeLocalList(key, values) {
  try { localStorage.setItem(ownLocalKey(key), JSON.stringify(values)); } catch (_) { /* private window */ }
}

// The list as saved: the server's, plus anything kept only in this browser.
async function loadOwnConfig(key) {
  let server = null;
  try {
    const res = await fetch('/api/config');
    if (res.ok) { const doc = await res.json(); if (Array.isArray(doc[key])) server = doc[key]; }
  } catch (_) { /* no server list */ }
  const local = readLocalList(key) || [];
  if (!server && !local.length) return null;
  const base = server || getCfgList(key) || [];
  const seen = new Set(base.map(v => String(v).toLowerCase()));
  return [...base, ...local.filter(v => !seen.has(String(v).toLowerCase()))];
}

// Returns 'server' or 'browser': where the list ended up.
async function saveOwnConfig(key, values) {
  try {
    const res = await fetch('/api/config', { method: 'PUT', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ [key]: values }) });
    if (res.ok) { writeLocalList(key, []); return 'server'; }
  } catch (_) { /* fall through to this browser */ }
  writeLocalList(key, values);
  return 'browser';
}

// Persists one or more config lists to the shared document.
async function saveConfigToDB(fields) {
  const body = {};
  for (const key of DB_FIELDS) {
    if (fields[key]) body[key] = fields[key];
  }
  const res = await fetch(getDbBaseUrl() + '/api/config', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Failed to save config (${res.status}): ${text}`);
  }
  return res.json();
}
