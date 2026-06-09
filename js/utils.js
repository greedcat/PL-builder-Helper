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
