// Where the PL builder keeps its settings. DATA_STORE=pg uses Postgres (PG_URL);
// anything else uses Mongo, as before.
const CONFIG_ID = 'pl-builder-config';
const MATCH_ID = 'container-match-config';
const MATCH_DB = 'match_helper';
const COLLECTION = 'PL_cols';

function mongoStore(env) {
  const store = { ready: false };
  const uri = env.MONGO_URI;
  const dbName = env.MONGO_DB;
  let col = null;
  let matchCol = null;
  if (!uri || !dbName) {
    console.warn('MONGO_URI / MONGO_DB not set — serving the app without the config API.');
  } else {
    const { MongoClient } = require('mongodb');
    const client = new MongoClient(uri);
    client.connect().then(() => {
      col = client.db(dbName).collection(COLLECTION);
      matchCol = client.db(MATCH_DB).collection('config');
      store.ready = true;
      console.log(`Connected to MongoDB database "${dbName}"`);
    }).catch((err) => {
      console.error('MongoDB connection failed — continuing without the config API:', err.message);
    });
    store.close = () => client.close();
  }
  store.getConfig = () => col.findOne({ _id: CONFIG_ID });
  store.saveConfig = async (set) => {
    await col.updateOne({ _id: CONFIG_ID }, { $set: set }, { upsert: true });
    return col.findOne({ _id: CONFIG_ID });
  };
  store.getDestOrder = async () => {
    const doc = await matchCol.findOne({ _id: MATCH_ID });
    return (doc && doc.destOrder) || [];
  };
  store.saveDestOrder = async (destOrder) => {
    await matchCol.updateOne({ _id: MATCH_ID }, { $set: { destOrder, updatedAt: new Date() } }, { upsert: true });
    return destOrder;
  };
  return store;
}

function pgStore(env) {
  const store = { ready: false };
  if (!env.PG_URL) {
    console.warn('PG_URL not set — serving the app without the config API.');
    store.getConfig = store.saveConfig = store.getDestOrder = store.saveDestOrder = async () => {
      throw new Error('Database not configured');
    };
    return store;
  }
  const { Pool } = require('pg');
  const u = new URL(env.PG_URL);
  u.searchParams.delete('sslmode'); // ssl is set in code, never in the URL
  const local = ['localhost', '127.0.0.1', '[::1]', '::1'].includes(u.hostname);
  const schema = env.PG_SCHEMA || 'plbuilder';
  if (!/^[a-z_][a-z0-9_]*$/.test(schema)) throw new Error('bad PG_SCHEMA');
  const pool = new Pool({
    connectionString: u.toString(),
    max: 3,
    ssl: local ? false : { rejectUnauthorized: false },
    options: `-c search_path=${schema}`,
  });
  pool.on('error', (e) => console.error('Postgres pool error:', e.message));
  store.ready = true;
  store.close = () => pool.end();
  const merge = `insert into settings (id, value) values ($1, $2::jsonb)
    on conflict (id) do update set value = settings.value || excluded.value, updated_at = now()
    returning value`;
  store.getConfig = async () => {
    const { rows } = await pool.query('select value from settings where id = $1', [CONFIG_ID]);
    return rows[0] ? { _id: CONFIG_ID, ...rows[0].value } : null;
  };
  store.saveConfig = async (set) => {
    const { rows } = await pool.query(merge, [CONFIG_ID, JSON.stringify(set)]);
    return { _id: CONFIG_ID, ...rows[0].value };
  };
  store.getDestOrder = async () => {
    const { rows } = await pool.query('select value from settings where id = $1', [MATCH_ID]);
    return (rows[0] && rows[0].value.destOrder) || [];
  };
  store.saveDestOrder = async (destOrder) => {
    await pool.query(merge, [MATCH_ID, JSON.stringify({ destOrder })]);
    return destOrder;
  };
  return store;
}

function createStore(env = process.env) {
  return env.DATA_STORE === 'pg' ? pgStore(env) : mongoStore(env);
}

module.exports = { createStore };
