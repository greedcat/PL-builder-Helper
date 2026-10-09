// ─────────────────────────────────────────────────────────────
// PASTED CELLS → ROWS
// Excel puts a copied range on the clipboard as tab-separated text. A cell
// holding a line break (Alt+Enter) comes quoted, with "" for a literal quote,
// so a plain split on newlines would cut "SHOE RACK\nBOOKSHELF" in two.
// ─────────────────────────────────────────────────────────────
function parseTsv(text) {
  const src  = String(text || '').replace(/\r\n?/g, '\n');
  const rows = [];
  let row = [], cell = '', i = 0, quoted = false, atCellStart = true;

  while (i < src.length) {
    const ch = src[i];
    if (quoted) {
      if (ch === '"' && src[i + 1] === '"') { cell += '"'; i += 2; continue; }
      if (ch === '"') { quoted = false; i++; continue; }
      cell += ch; i++; continue;
    }
    if (ch === '"' && atCellStart) { quoted = true; atCellStart = false; i++; continue; }
    if (ch === '\t') { row.push(cell); cell = ''; atCellStart = true; i++; continue; }
    if (ch === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; atCellStart = true; i++; continue; }
    cell += ch; atCellStart = false; i++;
  }
  // Excel ends the copy with a newline; only keep a last row that has content.
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }

  return rows.map(r => r.map(pastedValue));
}

// Numbers must arrive as numbers, or nothing downstream can sum them.
// Thousands separators are what Excel copies for a formatted figure.
function pastedValue(text) {
  const t = text.trim();
  if (t === '') return null;
  const n = t.replace(/,(?=\d{3}(\D|$))/g, '');
  // A code like 00123 is an ID, and turning it into 123 would lose it.
  if (/^0\d/.test(n)) return t;
  return /^-?(\d+\.?\d*|\.\d+)$/.test(n) ? Number(n) : t;
}

// Node (unit tests) — the browser just uses the globals.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { parseTsv, pastedValue };
}
