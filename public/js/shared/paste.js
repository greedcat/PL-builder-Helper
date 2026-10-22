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


// Some apps put only a formatted table on the clipboard, with no plain text
// beside it — Numbers, WPS and some browser spreadsheets among them. Read
// the table's cells straight from its HTML. A merged cell counts once, in
// its top-left position, as it does in a plain-text copy. (Browser only.)
function rowsFromHtml(html) {
  if (!html || typeof DOMParser === 'undefined') return [];
  const doc   = new DOMParser().parseFromString(html, 'text/html');
  const table = doc.querySelector('table');
  if (!table) return [];
  const grid  = [];
  [...table.rows].forEach((tr, r) => {
    grid[r] = grid[r] || [];
    let c = 0;
    for (const td of tr.cells) {
      while (grid[r][c] !== undefined) c++;
      td.querySelectorAll('br').forEach(br => br.replaceWith('\n'));
      const text = td.textContent.replace(/\u00a0/g, ' ').replace(/[ \t]*\n[ \t]*/g, '\n').trim();
      const rs = Math.max(1, td.rowSpan || 1), cs = Math.max(1, td.colSpan || 1);
      for (let dr = 0; dr < rs; dr++) {
        grid[r + dr] = grid[r + dr] || [];
        for (let dc = 0; dc < cs; dc++) grid[r + dr][c + dc] = dr || dc ? null : pastedValue(text);
      }
      c += cs;
    }
  });
  return grid.map(row => Array.from(row, v => (v === undefined ? null : v)));
}

const hasCells = rows => rows.some(r => r.some(v => v != null));

// What a paste carries, best source first: a copied .xlsx file, tab-
// separated text, then an HTML table. Returns { file }, { rows } or { error }.
function clipboardRows(cd) {
  const file = [...(cd.files || [])].find(f => /\.(xlsx|xls|xlsm|csv)$/i.test(f.name));
  if (file) return { file };
  let rows = parseTsv(cd.getData('text/plain'));
  if (!hasCells(rows) || rows.every(r => r.length < 2)) {
    const fromHtml = rowsFromHtml(cd.getData('text/html'));
    if (hasCells(fromHtml)) rows = fromHtml;
  }
  if (hasCells(rows)) return { rows };
  const kinds = [...(cd.types || [])].filter(t => t !== 'Files');
  const what  = cd.files && cd.files.length ? 'a file that is not a spreadsheet'
              : kinds.length ? kinds.join(', ') : 'nothing';
  return { error: `Nothing to read in what was pasted (the clipboard held ${String(what).replace(/[<>&]/g, '')}). `
    + 'Select the cells in Excel, press Ctrl+C / ⌘C, then paste here — or drop the Excel file in the box above.' };
}

// A paste that is a table, not a word typed somewhere.
function looksLikeTable(cd) {
  return !!((cd.files && cd.files.length)
    || (cd.getData('text/plain') || '').includes('\t')
    || /<table/i.test(cd.getData('text/html') || ''));
}

// Node (unit tests) — the browser just uses the globals.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { parseTsv, pastedValue, hasCells, clipboardRows, looksLikeTable };
}
