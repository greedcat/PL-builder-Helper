const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const ctx = { console, fetch: () => Promise.reject(new Error('no network')),
              document: { getElementById: () => null, addEventListener() {} } };
vm.createContext(ctx);
for (const f of ['db.js', 'config.js', 'utils.js', 'detect.js', 'columns.js', 'process.js']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'public', 'js', f), 'utf8'), ctx, { filename: f });
}
const { parseDelimitedText, tableCellValue, looksLikeTable,
        detectHeader, detectTableEnd, readExcel, modifyDF, getSummaryDF } = ctx;

test('parseDelimitedText splits tabs and lines, and keeps a quoted line break inside its cell', () => {
  assert.deepEqual(parseDelimitedText('a\tb\r\n1\t"two\nlines"\r\n'),
                   [['a', 'b'], ['1', 'two\nlines']]);
  assert.deepEqual(parseDelimitedText('x\t"say ""hi"""\n'), [['x', 'say "hi"']]);
  assert.deepEqual(parseDelimitedText('\t\t\n'), [['', '', '']]);
  assert.deepEqual(parseDelimitedText(''), []);
});

test('tableCellValue types cells the way Excel shows them', () => {
  assert.equal(tableCellValue(' 300 '), 300);
  assert.equal(tableCellValue('3395.8'), 3395.8);
  assert.equal(tableCellValue('1,590'), 1590);
  assert.equal(tableCellValue('-0.5'), -0.5);
  assert.equal(tableCellValue(''), null);
  assert.equal(tableCellValue('   '), null);
  assert.equal(tableCellValue('007'), '007');
  assert.equal(tableCellValue('8FUYSZX26095842A'), '8FUYSZX26095842A');
  assert.equal(tableCellValue('2OOPCDHH'), '2OOPCDHH');
});

test('looksLikeTable wants a tab and at least two lines', () => {
  assert.equal(looksLikeTable('a\tb\n1\t2'), true);
  assert.equal(looksLikeTable('a\tb'), false);
  assert.equal(looksLikeTable('just a value'), false);
  assert.equal(looksLikeTable(null), false);
});

test('a table copied from Excel runs through the whole pipeline', () => {
  const text = fs.readFileSync(path.join(__dirname, 'fixtures', 'pasted-table.tsv'), 'utf8');
  const raw  = parseDelimitedText(text).map(r => r.map(tableCellValue));
  const cleaned = raw.filter(r => r.some(v => v != null));

  const h = detectHeader(cleaned);
  assert.equal(h, 0, 'the first line is the header');
  const [last] = detectTableEnd(cleaned, h);
  assert.equal(last, cleaned.length - 2, 'the TOTAL line is not a load');

  const headers = cleaned[h].map(v => (v == null ? null : String(v)));
  const rows    = cleaned.slice(h + 1, last + 1);
  const re  = readExcel(headers, rows);
  assert.deepEqual(re.matchDict, {
    Destination: 'Amazon', Carton: 'CTNS', FBA_ID: 'Shipment ID',
    'REF ID': 'Reference ID', Weight: 'KGS', CMB: 'CBM',
  });

  const mod = modifyDF(re.headers, re.rows, re.matchDict);
  const sum = getSummaryDF(mod.headers, mod.rows);
  const byDest = Object.fromEntries(sum.rows.map(r => [r[0], { cartons: r[1], skids: r[2] }]));
  // A blank destination continues the row above it, so the six rows after
  // YOW3 are YOW3's. Skids are volume / 1.7, rounded up.
  assert.deepEqual(byDest, {
    XYY1: { cartons: 300,  skids: 2  },   //  3.00 CBM
    YOW3: { cartons: 1068, skids: 14 },   // 22.72 CBM
    XYY4: { cartons: 308,  skids: 16 },   // 26.63 CBM
    YYZ4: { cartons: 261,  skids: 8  },   // 12.85 CBM
    YYZ7: { cartons: 35,   skids: 2  },   //  2.80 CBM
  });
  const total = sum.rows.reduce((n, r) => n + r[1], 0);
  assert.equal(total, 1972, 'matches the sheet\'s own TOTAL');
  const nameIdx = mod.headers.indexOf('Product Name');
  assert.ok(mod.rows.some(r => r[nameIdx] === 'SHOE RACK\nBOOKSHELF'), 'the two-line name survived');
});
