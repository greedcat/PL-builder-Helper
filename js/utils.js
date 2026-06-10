// ─────────────────────────────────────────────────────────────
// TEXT UTILITIES
// ─────────────────────────────────────────────────────────────
function isEmpty(v) {
  return v === null || v === undefined || v === '';
}

function normalizeText(text) {
  if (text == null) return '';
  return String(text).toUpperCase().trim()
    .replace(/\s+/g, '').replace(/-/g, '/').replace(/[^\w一-鿿/]/g, '');
}

function normCBM(s) {
  if (s == null) return '';
  return String(s).normalize('NFKC').toLowerCase().trim()
    .replace(/\s+/g, '').replace(/（/g, '(').replace(/）/g, ')').replace(/[^a-z0-9一-鿿()]+/g, '');
}

function normWeight(s) {
  if (s == null) return '';
  return String(s).normalize('NFKC').toLowerCase().replace(/\s+/g, '').trim();
}

// Drops the internal _Pallet column and any column that is null in every row.
function getVisibleColumns(headers, rows) {
  const cols = headers
    .map((h, ci) => ({ h, ci }))
    .filter(({ h, ci }) => h !== '_Pallet' && rows.some(r => r[ci] != null));
  return {
    headers: cols.map(x => x.h),
    rows:    rows.map(r => cols.map(x => r[x.ci])),
  };
}
