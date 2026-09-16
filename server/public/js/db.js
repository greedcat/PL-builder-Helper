// ─────────────────────────────────────────────────────────────
// BACKEND API CONNECTION
// Talks to the small Node/Express server in server/ which stores
// the shared config document in MongoDB.
// ─────────────────────────────────────────────────────────────
const DB_DEFAULT_BASE  = 'https://pl-builder-helper-server-production.up.railway.app';
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
