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

const MATCH_HELPER_DB     = 'match_helper';
const MATCH_HELPER_DOC_ID = 'container-match-config';

if (!MONGO_URI || !MONGO_DB) {
  console.error('Missing MONGO_URI or MONGO_DB in environment.');
  process.exit(1);
}

const app    = express();
const client = new MongoClient(MONGO_URI);
let collection;
let matchHelperCol;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.get('/api/health', (req, res) => res.json({ ok: true }));

app.get('/api/config', async (req, res) => {
  try {
    const doc = await collection.findOne({ _id: DOC_ID });
    res.json(doc || {});
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to load config' });
  }
});

app.put('/api/config', async (req, res) => {
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
  try {
    const doc = await matchHelperCol.findOne({ _id: MATCH_HELPER_DOC_ID });
    res.json({ destOrder: (doc && doc.destOrder) || [] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to load destination order' });
  }
});

app.put('/api/match-helper/dest-order', async (req, res) => {
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

async function start() {
  await client.connect();
  collection = client.db(MONGO_DB).collection(COLLECTION);
  matchHelperCol = client.db(MATCH_HELPER_DB).collection('config');
  console.log(`Connected to MongoDB database "${MONGO_DB}"`);
  app.listen(PORT, () => console.log(`Server listening on port ${PORT}`));
}

start().catch(err => {
  console.error('Failed to start server:', err);
  process.exit(1);
});
