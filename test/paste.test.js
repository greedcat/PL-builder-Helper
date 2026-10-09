const { test } = require('node:test');
const assert = require('node:assert');
const { parseTsv, pastedValue } = require('../public/js/shared/paste.js');

test('parseTsv keeps a quoted multi-line cell in one row', () => {
  const rows = parseTsv('A\tB\tC\r\nx\t"SHOE RACK\nBOOKSHELF"\t230\r\n');
  assert.deepEqual(rows, [['A', 'B', 'C'], ['x', 'SHOE RACK\nBOOKSHELF', 230]]);
});

test('parseTsv keeps blank cells as null so columns stay aligned', () => {
  assert.deepEqual(parseTsv('\tLG79\t\t3'), [[null, 'LG79', null, 3]]);
});

test('parseTsv unescapes doubled quotes inside a quoted cell', () => {
  assert.deepEqual(parseTsv('"12"" TV"\tx'), [['12" TV', 'x']]);
});

test('pastedValue turns figures into numbers and keeps codes as text', () => {
  assert.equal(pastedValue('3395.8'), 3395.8);
  assert.equal(pastedValue('1,590'), 1590);
  assert.equal(pastedValue('00123'), '00123');
  assert.equal(pastedValue('8JN5TTZO'), '8JN5TTZO');
  assert.equal(pastedValue('  '), null);
});
