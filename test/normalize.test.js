const { test } = require('node:test');
const assert = require('node:assert');
const {
  normContainer, normFileNo, normDest, normOrderItem, matchDestIndex, guessColumn
} = require('../public/js/shared/normalize.js');

test('normContainer extracts the ISO 6346 code from prefixed values', () => {
  assert.equal(normContainer('DS-TIIU8097334'), 'TIIU8097334');
  assert.equal(normContainer('  oolu9286063 '), 'OOLU9286063');
  assert.equal(normContainer('OOLU9286063'), 'OOLU9286063');
});

test('normContainer keeps non-ISO values as cleaned alphanumerics', () => {
  assert.equal(normContainer('EXPRESSLINE 18050437052'), 'EXPRESSLINE18050437052');
  assert.equal(normContainer(null), '');
  assert.equal(normContainer(''), '');
});

test('normFileNo keeps the container file number, drops the shipment suffix', () => {
  assert.equal(normFileNo('2067-26033-01'), '2067-26033');
  assert.equal(normFileNo('2067-26033'), '2067-26033');
  assert.equal(normFileNo('  2067-26033-03 '), '2067-26033');
  assert.equal(normFileNo('plain'), 'plain');
  assert.equal(normFileNo(null), '');
});

test('normDest folds the known aliases', () => {
  assert.equal(normDest('私仓派送'), '私人地址');
  assert.equal(normDest('hold at wh'), 'HOLD');
  assert.equal(normDest('UPS ground'), '快递');
  assert.equal(normDest('Self - Pickup'), '自提');
  assert.equal(normDest(' YYZ3 '), 'YYZ3');
});

test('normOrderItem folds the same aliases from the order list side', () => {
  assert.equal(normOrderItem('Hold'), 'HOLD');
  assert.equal(normOrderItem('ups'), '快递');
  assert.equal(normOrderItem('self pickup'), '自提');
  assert.equal(normOrderItem(' YGK1 '), 'YGK1');
});

test('matchDestIndex matches exact names first', () => {
  const order = ['YYZ3', 'YYC4/6', 'YEG1/2', 'YVR3/4'].map(o => o.toUpperCase());
  assert.equal(matchDestIndex('YYZ3', order), 0);
  assert.equal(matchDestIndex('yyc4/6', order), 1);
});

test('matchDestIndex folds YYC/YEG/YVR sub-destinations into the prefix column', () => {
  const order = ['YYZ3', 'YYC4/6', 'YEG1/2', 'YVR3/4'].map(o => o.toUpperCase());
  assert.equal(matchDestIndex('YYC6', order), 1);
  assert.equal(matchDestIndex('YEG2', order), 2);
  assert.equal(matchDestIndex('YVR4', order), 3);
  assert.equal(matchDestIndex('ZZZ9', order), -1);
});

test('matchDestIndex prefers a dedicated column over the prefix fold', () => {
  const order = ['YEG1', 'YEG1/2'].map(o => o.toUpperCase());
  assert.equal(matchDestIndex('YEG1', order), 0);
});

test('guessColumn prefers exact header matches, then substrings', () => {
  const cols = ['Client Name', 'Container No.', 'File No', 'Pallets'];
  assert.equal(guessColumn(cols, ['file no', 'file']), 'File No');
  assert.equal(guessColumn(cols, ['container', '柜号']), 'Container No.');
  assert.equal(guessColumn(cols, ['nothing']), '');
});
