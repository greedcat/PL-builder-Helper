const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

// The pipeline modules are plain browser scripts sharing one global scope.
// Load them the same way into a sandbox so they can be exercised in Node.
const ctx = { console, fetch: () => Promise.reject(new Error('no network')),
              document: { getElementById: () => null, addEventListener() {} } };
vm.createContext(ctx);
for (const f of ['db.js', 'config.js', 'utils.js', 'detect.js', 'columns.js', 'process.js']) {
  const file = path.join(__dirname, '..', 'public', 'js', f);
  vm.runInContext(fs.readFileSync(file, 'utf8'), ctx, { filename: f });
}
const { getSummaryDF } = ctx;

const H = ['Destination', 'Carton', 'CMB', '_Row', '_Pallet'];
const row = (dest, ctn, cbm) => [dest, ctn, cbm, 0, cbm == null ? null : cbm / 1.7];
const skids = rows => getSummaryDF(H, rows).rows.map(r => [r[0], r[2]]);

test('one small carton still needs one skid', () => {
  assert.deepEqual(skids([row('YYZ3', 1, 0.05)]), [['YYZ3', 1]]);
  assert.deepEqual(skids([row('YYZ3', 1, 0.085)]), [['YYZ3', 1]]);
});

test('a carton with no volume still needs one skid', () => {
  assert.deepEqual(skids([row('YYZ4', 1, null)]), [['YYZ4', 1]]);
  assert.deepEqual(skids([row('YYZ4', 3, null)]), [['YYZ4', 1]]);
});

test('the estimate still rounds volume up and never exceeds the carton count', () => {
  assert.deepEqual(skids([row('YOO1', 10, 1.7)]), [['YOO1', 1]]);
  assert.deepEqual(skids([row('YOO1', 10, 1.8)]), [['YOO1', 2]]);
  assert.deepEqual(skids([row('YOO1', 10, 3.5)]), [['YOO1', 3]]);
  assert.deepEqual(skids([row('YOO1', 2, 8.5)]), [['YOO1', 2]]);
});

test('a destination with no cartons is left at zero', () => {
  assert.deepEqual(skids([row('YHM1', 0, null)]), [['YHM1', 0]]);
});
