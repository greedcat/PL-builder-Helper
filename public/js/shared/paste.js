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


// Where each cell of a copied table sits once its row and column spans are
// laid out, as Excel copies a merged cell: one <td rowspan / colspan> in its
// top-left position. cells: rows of { text, rs, cs } → [{ r, c, rs, cs, text }].
function layoutSpans(rows) {
  const taken = [], out = [];
  rows.forEach((tr, r) => {
    taken[r] = taken[r] || [];
    let c = 0;
    for (const td of tr) {
      while (taken[r][c]) c++;
      const rs = Math.max(1, td.rs || 1), cs = Math.max(1, td.cs || 1);
      for (let dr = 0; dr < rs; dr++) {
        taken[r + dr] = taken[r + dr] || [];
        for (let dc = 0; dc < cs; dc++) taken[r + dr][c + dc] = true;
      }
      out.push({ r, c, rs, cs, text: td.text });
      c += cs;
    }
  });
  return out;
}

// The merged blocks of a laid-out table, as SheetJS keeps them in
// ws['!merges']: { s: top-left, e: bottom-right }, each { r, c }.
function mergesOf(cells) {
  return cells.filter(x => x.rs > 1 || x.cs > 1)
    .map(x => ({ s: { r: x.r, c: x.c }, e: { r: x.r + x.rs - 1, c: x.c + x.cs - 1 } }));
}

// A copied HTML table's cells as { text, rs, cs } rows. (Browser only.)
function htmlCells(html) {
  if (!html || typeof DOMParser === 'undefined') return [];
  const doc   = new DOMParser().parseFromString(html, 'text/html');
  const table = doc.querySelector('table');
  if (!table) return [];
  return [...table.rows].map(tr => [...tr.cells].map(td => {
    td.querySelectorAll('br').forEach(br => br.replaceWith('\n'));
    const text = td.textContent.replace(/\u00a0/g, ' ').replace(/[ \t]*\n[ \t]*/g, '\n').trim();
    return { text, rs: td.rowSpan || 1, cs: td.colSpan || 1 };
  }));
}

// Some apps put only a formatted table on the clipboard, with no plain text
// beside it — Numbers, WPS and some browser spreadsheets among them. Read
// the table's cells straight from its HTML. A merged cell counts once, in
// its top-left position, as it does in a plain-text copy.
function rowsFromHtml(html) {
  const cells = layoutSpans(htmlCells(html));
  if (!cells.length) return [];
  const grid = [];
  cells.forEach(x => {
    for (let dr = 0; dr < x.rs; dr++) {
      grid[x.r + dr] = grid[x.r + dr] || [];
      for (let dc = 0; dc < x.cs; dc++) grid[x.r + dr][x.c + dc] = dr || dc ? null : pastedValue(x.text);
    }
  });
  return grid.map(row => Array.from(row, v => (v === undefined ? null : v)));
}

/* Excel's copy carries the merged cells only in its HTML. They are taken
   from there when its table lines up with the rows read (same rows, no
   wider), so the pasted sheet keeps its merges like an uploaded file. */
function mergesFor(rows, cells) {
  if (!cells.length) return [];
  const height = cells.reduce((n, x) => Math.max(n, x.r + x.rs), 0);
  const width  = cells.reduce((n, x) => Math.max(n, x.c + x.cs), 0);
  const cols   = rows.reduce((n, r) => Math.max(n, r.length), 0);
  if (height !== rows.length || width > cols + 1) return [];
  return mergesOf(cells);
}

const hasCells = rows => rows.some(r => r.some(v => v != null));

// What a paste carries, best source first: a copied .xlsx file, tab-
// separated text, then an HTML table. Returns { file }, { rows } or { error }.
function clipboardRows(cd) {
  const file = [...(cd.files || [])].find(f => /\.(xlsx|xls|xlsm|csv)$/i.test(f.name));
  if (file) return { file };
  let rows = parseTsv(cd.getData('text/plain'));
  const html = cd.getData('text/html');
  if (!hasCells(rows) || rows.every(r => r.length < 2)) {
    const fromHtml = rowsFromHtml(html);
    if (hasCells(fromHtml)) rows = fromHtml;
  }
  if (hasCells(rows)) return { rows, merges: mergesFor(rows, layoutSpans(htmlCells(html))) };
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
  module.exports = { parseTsv, pastedValue, hasCells, clipboardRows, looksLikeTable, layoutSpans, mergesOf, mergesFor };
}
