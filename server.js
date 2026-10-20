// ─────────────────────────────────────────────────────────────
// PL Builder Helper — the one server.
// Serves the browser app from public/ and the small config API.
// MongoDB is optional: without MONGO_URI the app still runs and
// the config endpoints answer 503, so every page keeps working
// on its built-in fallback lists.
// ─────────────────────────────────────────────────────────────
require('dotenv').config();
const path = require('path');
const express = require('express');
const cors = require('cors');
const { MongoClient } = require('mongodb');

const PORT       = process.env.PORT || 3000;
const MONGO_URI  = process.env.MONGO_URI;
const MONGO_DB   = process.env.MONGO_DB;
const COLLECTION = 'PL_cols';
const DOC_ID     = 'pl-builder-config';
const FIELDS     = ['DEST_LIST', 'CARTON_KEYWORDS', 'CBM_KEYWORDS', 'WEIGHT_KEYWORDS', 'NO_NEED_COL'];

// fba-app puts each PL made here into its container's SharePoint folder.
// Both are needed; without them the PL is only downloaded.
const FBA_APP_URL    = (process.env.FBA_APP_URL || '').replace(/\/+$/, '');
const PL_BUILDER_KEY = process.env.PL_BUILDER_KEY || '';

const MATCH_HELPER_DB     = 'match_helper';
const MATCH_HELPER_DOC_ID = 'container-match-config';

const app = express();
let collection     = null; // PL config collection, null until Mongo connects
let matchHelperCol = null;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.get('/api/health', (req, res) => res.json({ ok: true, db: !!collection }));

// Every data endpoint goes through this gate so a missing database reads as
// one consistent, non-fatal condition on the client.
function requireDb(res, col) {
  if (!col) { res.status(503).json({ error: 'Database not configured' }); return false; }
  return true;
}

app.get('/api/config', async (req, res) => {
  if (!requireDb(res, collection)) return;
  try {
    const doc = await collection.findOne({ _id: DOC_ID });
    res.json(doc || {});
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to load config' });
  }
});

app.put('/api/config', async (req, res) => {
  if (!requireDb(res, collection)) return;
  try {
    const set = {};
    for (const key of FIELDS) {
      if (Array.isArray(req.body[key])) set[key] = req.body[key];
    }
    if (Object.keys(set).length === 0) {
      return res.status(400).json({ error: 'No valid config fields provided' });
    }
    await collection.updateOne({ _id: DOC_ID }, { $set: set }, { upsert: true });
    const doc = await collection.findOne({ _id: DOC_ID });
    res.json(doc);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to save config' });
  }
});

app.get('/api/match-helper/dest-order', async (req, res) => {
  if (!requireDb(res, matchHelperCol)) return;
  try {
    const doc = await matchHelperCol.findOne({ _id: MATCH_HELPER_DOC_ID });
    res.json({ destOrder: (doc && doc.destOrder) || [] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to load destination order' });
  }
});

app.put('/api/match-helper/dest-order', async (req, res) => {
  if (!requireDb(res, matchHelperCol)) return;
  try {
    if (!Array.isArray(req.body.destOrder)) {
      return res.status(400).json({ error: 'destOrder must be an array' });
    }
    const destOrder = req.body.destOrder.map(s => String(s).trim()).filter(Boolean);
    await matchHelperCol.updateOne(
      { _id: MATCH_HELPER_DOC_ID },
      { $set: { destOrder, updatedAt: new Date() } },
      { upsert: true }
    );
    res.json({ destOrder });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to save destination order' });
  }
});

// The generated PL, passed on to fba-app, which finds or makes the
// container's folder (Containers Info / client / file-container) and puts
// the file there in place of any earlier one. The key stays on this server.
app.get('/api/sharepoint-pl', (req, res) => res.json({ enabled: !!(FBA_APP_URL && PL_BUILDER_KEY) }));

app.post('/api/sharepoint-pl', express.raw({ type: () => true, limit: '20mb' }), async (req, res) => {
  if (!FBA_APP_URL || !PL_BUILDER_KEY) {
    return res.status(503).json({ error: 'SharePoint upload is not set up on this server' });
  }
  if (!req.body || !req.body.length) return res.status(400).json({ error: 'no file' });
  const q = new URLSearchParams();
  for (const k of ['container', 'file_no', 'customer', 'name']) q.set(k, String(req.query[k] || ''));
  try {
    const up = await fetch(`${FBA_APP_URL}/api/pl-builder/upload?${q}`, {
      method: 'POST',
      headers: { 'x-pl-builder-key': PL_BUILDER_KEY,
                 'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
      body: req.body,
    });
    const out = await up.json().catch(() => ({ error: `fba-app answered ${up.status}` }));
    res.status(up.ok ? 200 : (up.status === 401 ? 502 : up.status)).json(
      up.status === 401 ? { error: 'fba-app refused the key (check PL_BUILDER_KEY on both services)' } : out);
  } catch (err) {
    console.error('SharePoint upload failed:', err.message);
    res.status(502).json({ error: 'Could not reach fba-app: ' + err.message });
  }
});

async function connectDb() {
  if (!MONGO_URI || !MONGO_DB) {
    console.warn('MONGO_URI / MONGO_DB not set — serving the app without the config API.');
    return;
  }
  try {
    const client = new MongoClient(MONGO_URI);
    await client.connect();
    collection     = client.db(MONGO_DB).collection(COLLECTION);
    matchHelperCol = client.db(MATCH_HELPER_DB).collection('config');
    console.log(`Connected to MongoDB database "${MONGO_DB}"`);
  } catch (err) {
    console.error('MongoDB connection failed — continuing without the config API:', err.message);
  }
}

app.listen(PORT, () => console.log(`Server listening on port ${PORT}`));
connectDb();
