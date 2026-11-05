const { test, before, after } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { Pool } = require('pg');
const { createStore } = require('../lib/store');

const URL_ = process.env.TEST_PG_URL || 'postgres://localhost/fba_test';
const schema = 't_' + Math.random().toString(16).slice(2, 10);
let admin;
let store;

before(async () => {
  admin = new Pool({ connectionString: URL_, max: 1 });
  await admin.query(`create schema ${schema}`);
  await admin.query(`set search_path = ${schema}; ${fs.readFileSync(path.join(__dirname, '../sql/schema.sql'), 'utf8')}`);
  store = createStore({ DATA_STORE: 'pg', PG_URL: URL_, PG_SCHEMA: schema });
});
after(async () => {
  await store.close();
  await admin.query(`drop schema ${schema} cascade`);
  await admin.end();
});

test('nothing saved yet: config is null, dest order is empty', async () => {
  assert.equal(await store.getConfig(), null);
  assert.deepEqual(await store.getDestOrder(), []);
});

test('saveConfig merges the given fields only', async () => {
  await store.saveConfig({ DEST_LIST: ['A'] });
  const saved = await store.saveConfig({ CBM_KEYWORDS: ['cbm'] });
  assert.deepEqual(saved, { _id: 'pl-builder-config', DEST_LIST: ['A'], CBM_KEYWORDS: ['cbm'] });
  assert.deepEqual(await store.getConfig(), saved);
});

test('dest order round trip, and it does not touch the config', async () => {
  assert.deepEqual(await store.saveDestOrder(['YYZ1', 'YOW1']), ['YYZ1', 'YOW1']);
  assert.deepEqual(await store.getDestOrder(), ['YYZ1', 'YOW1']);
  assert.deepEqual(Object.keys(await store.getConfig()).sort(), ['CBM_KEYWORDS', 'DEST_LIST', '_id']);
});

test('pg store without PG_URL is not ready', () => {
  const s = createStore({ DATA_STORE: 'pg' });
  assert.equal(s.ready, false);
});

test('mongo store without MONGO_URI is not ready', () => {
  assert.equal(createStore({}).ready, false);
});
