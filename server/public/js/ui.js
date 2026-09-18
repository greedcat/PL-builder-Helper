// ─────────────────────────────────────────────────────────────
// UI
// ─────────────────────────────────────────────────────────────
const dropZone     = document.getElementById('dropZone');
const fileInput    = document.getElementById('excelFile');
const fileBadge    = document.getElementById('fileBadge');
const statusEl     = document.getElementById('status');
const hdrRowInput  = document.getElementById('hdrRowInput');
const lastRowInput = document.getElementById('lastRowInput');
const sheetView    = document.getElementById('plOriginalSheet');
const sheetLegend  = document.getElementById('plSheetLegend');
const previewSec   = document.getElementById('plPreview');
const btnViewSheet = document.getElementById('btnViewSheet');

initConfigFromDB();

// Cache of the most recently read workbook, keyed by File object.
let cachedFile      = null;
let cachedSheetData = null;

// Excel file drop zone
dropZone.addEventListener('click',     () => fileInput.click());
dropZone.addEventListener('dragover',  e  => { e.preventDefault(); dropZone.classList.add('active'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('active'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('active');
  if (e.dataTransfer.files[0]) setExcelFile(e.dataTransfer.files[0], e.dataTransfer);
});
fileInput.addEventListener('change', () => { if (fileInput.files[0]) setExcelFile(fileInput.files[0]); });

function setExcelFile(file, dt) {
  if (dt) { const dtp = new DataTransfer(); dtp.items.add(file); fileInput.files = dtp.files; }
  fileBadge.innerHTML     = `<span>📄</span> ${file.name}`;
  fileBadge.style.display = 'inline-flex';
  dropZone.querySelector('.drop-text').style.display = 'none';
  dropZone.classList.add('pl-drop-compact');

  // A new file invalidates the cached sheet data, row range and previews.
  cachedFile      = null;
  cachedWb        = null;
  cachedSheetData = null;
  cachedSheetName = null;
  activeSheet     = null;
  manualRange     = null;
  roleOverrides   = {};
  columnOverrides = {};
  clearPreviewEdits();
  clearDestRenames();
  clearLearnQueue();
  editHistory = [];
  selectMode      = 'replace';
  hdrRowInput.value  = '';
  lastRowInput.value = '';
  sheetView.style.display   = 'none';
  sheetView.innerHTML       = '';
  sheetLegend.style.display = 'none';
  sheetLegend.innerHTML     = '';
  previewSec.style.display  = 'none';
  hideLearnBar();

  showFileOnLoad();
}

// Opens the sheet grid and builds the packing-list preview as soon as a file
// is chosen, so the detected range and its result are visible without a click.
async function showFileOnLoad() {
  try {
    showStatus('Reading file…', 'info', true);
    const data = await loadSheetData(fileInput.files[0]);

    // SheetJS parses almost anything without throwing, so an unreadable file
    // arrives as an empty sheet rather than an error. Say so plainly.
    if (!data.cleaned.length) {
      sheetView.style.display   = 'none';
      sheetLegend.style.display = 'none';
      setSheetToggleLabel(false);
      showStatus('That file has no readable rows. Check it opens in Excel and is not password protected.', 'error');
      return;
    }

    setSheetToggleLabel(true);
    sheetView.style.display   = 'block';
    sheetLegend.style.display = 'block';
    await refreshSheetView({ rebuild: true });
    await refreshPreview();

    updateLoadStatus(data);
  } catch (err) {
    console.error(err);
    sheetView.style.display   = 'none';
    sheetLegend.style.display = 'none';
    setSheetToggleLabel(false);
    showStatus('Could not read that file: ' + err.message, 'error');
  }
}

// The status must agree with the legend: a range with no matched columns is
// not a success, however many rows it happens to cover. Recomputed whenever
// Rows the range covers that actually carry something. A gap left between
// two blocks of the client's sheet sits inside the range but is not a line
// of the packing list, and counting it said 8 lines where 6 were written.
function countDataRows(data, range) {
  return rangeDataRowIndices(range).filter(i =>
    sliceRow(data.raw[i], range.firstCol, range.lastCol).some(v => !isEmpty(v))).length;
}

// the range or a column mapping changes.
function updateLoadStatus(data) {
  const range   = getRange(data);
  const nData   = countDataRows(data, range);
  const missing = classifyColumns(data, range).missing;

  if (nData < 1) {
    showStatus('Sheet loaded, but no data rows were found below the header. Drag on the grid to set the range.', 'error');
  } else if (missing.length === PL_ROLES.length) {
    showStatus('Sheet loaded, but no columns were recognised. The header row is probably wrong — drag on the grid to select it.', 'error');
  } else if (missing.length) {
    showStatus(`${nData} data row${nData === 1 ? '' : 's'} selected, but no column matched ${missing.join(', ')}. `
             + 'Click that chip to pick its column, or right-click the column in the grid.', 'error');
  } else {
    showStatus(`✓ ${nData} data row${nData === 1 ? '' : 's'} selected, all columns matched.`, 'success');
  }
}

// The container number is usually printed above the table. Fill it in, but
// never overwrite something the user has already typed.
function setSheetToggleLabel(shown) {
  btnViewSheet.textContent = shown ? 'Hide Data Range' : 'View & Select Data Range';
}

// Rebuilds the packing-list preview from the current range.
async function refreshPreview() {
  const file = fileInput.files[0];
  if (!file) return;
  const { modH, modR, sumH, sumR, longDests } = await runPipeline(file);
  renderPreview(modH, modR, sumH, sumR, longDests);
}

function showStatus(msg, type = 'info', spinner = false) {
  statusEl.className     = type;
  statusEl.style.display = 'block';
  statusEl.innerHTML     = spinner ? `<span class="spinner"></span>${msg}` : msg;
}

// ─────────────────────────────────────────────────────────────
// Reads and caches the workbook.
// `raw` keeps every row, so grid row numbers match the real Excel file.
// `cleaned` drops blank rows because the detectors are tuned for that shape;
// `cleanedIdx` maps a cleaned index back to its raw row.
// ─────────────────────────────────────────────────────────────
let cachedWb        = null;
let cachedSheetName = null;

// Which worksheet the builder reads. Reset whenever a new file arrives.
let activeSheet = null;

async function loadWorkbook(file) {
  if (cachedFile === file && cachedWb) return cachedWb;
  const buf = await file.arrayBuffer();
  cachedWb  = XLSX.read(buf, { type: 'array', raw: true });
  cachedFile      = file;
  cachedSheetName = null;
  cachedSheetData = null;
  activeSheet     = null;
  return cachedWb;
}

async function loadSheetData(file) {
  const wbIn = await loadWorkbook(file);
  if (!activeSheet || !wbIn.SheetNames.includes(activeSheet)) activeSheet = wbIn.SheetNames[0];
  if (cachedSheetName === activeSheet && cachedSheetData) return cachedSheetData;

  const raw = XLSX.utils.sheet_to_json(wbIn.Sheets[activeSheet],
                                       { header: 1, defval: null, raw: true });
  // Done once here, not per render: from this point the sheet has a real
  // value in every row of a merged destination.
  fillMergedRows(wbIn.Sheets[activeSheet], raw);

  // Excel's used range runs well past the data — a sheet with nineteen rows
  // of content routinely claims a thousand — and the grid would then draw a
  // screenful of empty rows below the table. Trailing empties only, so every
  // row index still matches the row number in Excel.
  let lastUsed = -1;
  for (let i = raw.length - 1; i >= 0; i--) {
    if (raw[i] && raw[i].some(v => !isEmpty(v))) { lastUsed = i; break; }
  }
  raw.length = Math.max(lastUsed + 1, 1);

  const cleaned    = [];
  const cleanedIdx = [];
  raw.forEach((r, i) => {
    if (r && r.some(v => !isEmpty(v))) { cleaned.push(r); cleanedIdx.push(i); }
  });

  // Excel's used range often stretches far past the data — real files arrive
  // claiming a thousand columns with thirteen filled. Measuring the widest
  // row would then build a million empty cells and hang the page.
  let maxCols = 0;
  for (const r of raw) {
    if (!r) continue;
    for (let c = r.length - 1; c >= maxCols; c--) {
      if (!isEmpty(r[c])) { maxCols = c + 1; break; }
    }
  }
  maxCols = Math.max(maxCols, 1);
  cachedSheetName = activeSheet;
  cachedSheetData = { wbIn, sheetNames: wbIn.SheetNames, sheetName: activeSheet,
                      raw, cleaned, cleanedIdx, maxCols };
  return cachedSheetData;
}

// Switching sheet invalidates everything that described the old one.
async function setActiveSheet(name) {
  activeSheet     = name;
  cachedSheetName = null;
  cachedSheetData = null;
  manualRange     = null;
  roleOverrides   = {};
  columnOverrides = {};
  clearPreviewEdits();
  clearDestRenames();
  clearLearnQueue();
  editHistory = [];
  selectMode      = 'replace';
  hdrRowInput.value  = '';
  lastRowInput.value = '';
  await refreshSheetView({ rebuild: true });
  await refreshPreview();
  updateLoadStatus(await loadSheetData(fileInput.files[0]));
}

// ─────────────────────────────────────────────────────────────
// Data range
// One rectangle drives everything: its first row is the header, the rest are
// data, and its column span limits which columns are read. Coordinates are
// 0-based raw row/column indices, matching the grid.
// ─────────────────────────────────────────────────────────────
let manualRange = null;   // null = fall back to auto-detection

// How a plain drag behaves: replace the selection, or add another block.
let selectMode = 'replace';

// Role → header name the user picked, or null for "not used". A role absent
// from this map is left to auto-detection.
let roleOverrides = {};

// Short names accepted for long destinations: original text → chosen name.
// Nothing is renamed until the user asks for it.
let destRenames = {};

function clearDestRenames() { destRenames = {}; }

// Hand edits made in the preview, keyed "<rowIndex>\u0000<columnName>".
// Loads edits feed the summary; summary edits then override it.
let previewEdits = { loads: new Map(), summary: new Map() };

function hasPreviewEdits() {
  return previewEdits.loads.size > 0 || previewEdits.summary.size > 0;
}

// Structural changes move rows around, so edits keyed by row would land on
// the wrong data. Drop them rather than silently misapply them.
function clearPreviewEdits() {
  previewEdits = { loads: new Map(), summary: new Map() };
}

// Numbers typed into the preview must reach the workbook as numbers.
function coerceCell(text) {
  const t = String(text).trim();
  if (t === '') return null;
  return /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : t;
}

function applyEdits(headers, rows, edits) {
  if (!edits.size) return rows;
  const out = rows.map(r => [...r]);
  // Edits are keyed by the row's stable id, so they follow their own row
  // through the re-sort that a destination rename causes.
  const idIdx = headers.indexOf('_Row');
  const byId  = new Map();
  if (idIdx >= 0) out.forEach(r => byId.set(String(r[idIdx]), r));
  for (const [key, value] of edits) {
    const sep  = key.indexOf('\u0000');
    const id   = key.slice(0, sep);
    const name = key.slice(sep + 1);
    const ci   = headers.indexOf(name);
    const row  = idIdx >= 0 ? byId.get(id) : out[Number(id)];
    if (ci >= 0 && row) row[ci] = value;
  }
  return out;
}

// Absolute column index → 'keep' or 'drop', set by right-clicking a column.
// Overrides the shared NO_NEED_COL list for this file only.
let columnOverrides = {};

// Which columns are dropped from the output, given the ignore list and any
// per-column override. Returns absolute column indices.
function droppedColumns(headers, firstCol) {
  const out = new Set();
  headers.forEach((name, i) => {
    const abs = i + firstCol;
    const ov  = columnOverrides[abs];
    if (ov === 'drop') { out.add(abs); return; }
    if (ov === 'keep') return;
    if (name != null && NO_NEED_COL.includes(name)) out.add(abs);
  });
  return out;
}

function autoRange(data) {
  const { cleaned, cleanedIdx, maxCols } = data;
  if (!cleaned.length) return { hdrRow: 0, lastRow: 0, firstCol: 0, lastCol: maxCols - 1, extras: [], auto: true };
  const h = detectHeader(cleaned);
  const l = detectTableEnd(cleaned, h)[0];
  return {
    hdrRow:   cleanedIdx[h] ?? 0,
    lastRow:  cleanedIdx[Math.max(l, h)] ?? cleanedIdx[h] ?? 0,
    firstCol: 0,
    lastCol:  maxCols - 1,
    extras:   [],
    auto:     true,
  };
}

function getRange(data) {
  if (!manualRange) return autoRange(data);
  const lastRaw = Math.max(data.raw.length - 1, 0);
  const hdrRow  = Math.min(Math.max(manualRange.hdrRow, 0), lastRaw);
  return {
    hdrRow,
    lastRow:  Math.min(Math.max(manualRange.lastRow, hdrRow), lastRaw),
    firstCol: Math.min(Math.max(manualRange.firstCol, 0), data.maxCols - 1),
    lastCol:  Math.min(Math.max(manualRange.lastCol, 0), data.maxCols - 1),
    extras:   (manualRange.extras || [])
      .map(b => ({ r1: Math.min(Math.max(b.r1, 0), lastRaw),
                   r2: Math.min(Math.max(b.r2, 0), lastRaw) }))
      .filter(b => b.r2 >= b.r1),
    auto:     false,
  };
}

// Every data row the range covers, in sheet order: the main block plus any
// extra blocks added with Ctrl or Cmd. That is how a table broken by a gap
// is read as one table.
function rangeDataRowIndices(range) {
  const seen = new Set();
  for (let r = range.hdrRow + 1; r <= range.lastRow; r++) seen.add(r);
  for (const b of range.extras || []) {
    for (let r = b.r1; r <= b.r2; r++) if (r !== range.hdrRow) seen.add(r);
  }
  return [...seen].sort((a, b) => a - b);
}

// Reads one row across the selected column span, padding short rows.
function sliceRow(row, firstCol, lastCol) {
  const out = [];
  for (let c = firstCol; c <= lastCol; c++) out.push(row && row[c] !== undefined ? row[c] : null);
  return out;
}

// Runs detection + processing, returning everything needed to either
// preview or write the packing list.
async function runPipeline(file) {
  const data  = await loadSheetData(file);
  const range = getRange(data);
  const { hdrRow, lastRow, firstCol, lastCol } = range;

  const headers = sliceRow(data.raw[hdrRow] || [], firstCol, lastCol)
    .map(h => (h != null ? String(h) : null));
  const rows = rangeDataRowIndices(range)
    .map(i => sliceRow(data.raw[i], firstCol, lastCol));

  const dropAbs   = droppedColumns(headers, firstCol);
  const dropLocal = new Set([...dropAbs].map(c => c - firstCol));
  const { headers: reH, rows: reR, matchDict } = readExcel(headers, rows, roleOverrides, dropLocal);
  const { headers: modH, rows: modR0 } = modifyDF(reH, reR, matchDict);

  const modR = applyEdits(modH, modR0, previewEdits.loads);

  // An edited CBM must move the pallet estimate with it.
  const cbmIdx = modH.indexOf('CMB');
  const pltIdx = modH.indexOf('_Pallet');
  if (cbmIdx >= 0 && pltIdx >= 0) {
    for (const row of modR) {
      const cbm = parseFloat(row[cbmIdx]);
      row[pltIdx] = isNaN(cbm) ? null : cbm / 1.7;
    }
  }

  // Destinations that match no known code. Each is offered a short name.
  const destIdx = modH.indexOf('Destination');
  const longDests = [];
  if (destIdx >= 0) {
    const seen = new Map();
    for (const row of modR) {
      const v = row[destIdx];
      if (v == null || v === '') continue;
      const text = String(v);
      if (DEST_LIST.some(code => text.includes(code))) continue;
      seen.set(text, (seen.get(text) || 0) + 1);
    }
    for (const [text, count] of seen) {
      longDests.push({ text, count, suggestion: simplifyPrivateAddress(text), applied: destRenames[text] || null });
    }
    // Apply only what the user has accepted, noting which rows it moved so
    // they can be set apart below.
    const renamed = new Set();
    const rowIdIdx = modH.indexOf('_Row');
    for (const row of modR) {
      const v = row[destIdx];
      if (v != null && destRenames[String(v)]) {
        row[destIdx] = destRenames[String(v)];
        if (rowIdIdx >= 0) renamed.add(String(row[rowIdIdx]));
      }
    }
    // A destination that was changed by hand — renamed here or typed over in
    // the preview — goes to the end of the loads, where it can be checked
    // against the paperwork instead of being lost among the codes it now
    // sorts next to.
    const destChanged = (row) => {
      if (rowIdIdx < 0) return false;
      const id = String(row[rowIdIdx]);
      return renamed.has(id) || previewEdits.loads.has(`${id}\u0000Destination`);
    };
    // Renaming can reorder groups, so sort again by destination.
    modR.sort((a, b) => {
      const ca = destChanged(a) ? 1 : 0, cb = destChanged(b) ? 1 : 0;
      if (ca !== cb) return ca - cb;                 // changed ones last
      const x = String(a[destIdx] ?? ''), y = String(b[destIdx] ?? '');
      return x < y ? -1 : x > y ? 1 : 0;
    });
  }

  const { headers: sumH, rows: sumR0 } = getSummaryDF(modH, modR);
  const sumR = applyEdits(sumH, sumR0, previewEdits.summary);

  return { wbIn: data.wbIn, data, range, modH, modR, sumH, sumR, longDests };
}

// ─────────────────────────────────────────────────────────────
// Table rendering helpers
// ─────────────────────────────────────────────────────────────
// splitAfter: row indices that end a destination group, drawn with a thick
// line. `editKey` ('loads' | 'summary') makes the cells editable in place.
function tableHtml(headers, rows, { splitAfter = null, editKey = null, rowIds = null, tints = null } = {}) {
  const edits = editKey ? previewEdits[editKey] : null;
  let html = '<table><thead><tr>';
  headers.forEach(h => {
    html += `<th>${escapeHtml(h ?? '').replace(/\n/g, '<br>')}</th>`;
  });
  // Empty trailing column. It takes whatever width is left over, so the
  // real columns size to their content instead of one of them stretching
  // across half the card.
  html += '<th class="pl-fill" aria-hidden="true"></th>';
  html += '</tr></thead><tbody>';
  rows.forEach((row, ri) => {
    const rowCls = [];
    if (splitAfter && splitAfter.has(ri) && ri < rows.length - 1) rowCls.push('pl-group-end');
    // The destination's colour runs the width of the row, so a group reads
    // as one band rather than one tinted cell.
    if (tints && tints[ri]) rowCls.push(tints[ri]);
    const rowId = rowIds ? rowIds[ri] : ri;
    html += `<tr${rowCls.length ? ` class="${rowCls.join(' ')}"` : ''}>`;
    row.forEach((v, ci) => {
      const name = headers[ci];
      const edited = edits && edits.has(`${rowId}\u0000${name}`);
      const numCls  = typeof tidyNumber(v) === 'number' ? ' pl-num' : '';
      const attrs = editKey
        ? ` tabindex="0" spellcheck="false" class="pl-edit${edited ? ' pl-edited' : ''}${numCls}"`
        + ` data-table="${editKey}" data-row="${rowId}" data-col="${escapeHtml(String(name ?? ''))}"`
        : '';
      const shown = tidyNumber(v);
      const text = shown != null ? String(shown) : '';
      // Long values are clipped to keep one value per line, so carry the
      // whole thing in a tooltip.
      const tip = text.length > 24 ? ` title="${escapeHtml(text)}"` : '';
      const plainCls = numCls.trim();
      html += `<td${attrs}${editKey ? '' : (plainCls ? ` class="${plainCls}"` : '')}${tip}>${
        escapeHtml(text)}</td>`;
    });
    html += '<td class="pl-fill"></td>';
    html += '</tr>';
  });
  html += '</tbody></table>';
  return html;
}

// ─────────────────────────────────────────────────────────────
// Undo and reset. Every change to the preview — a typed cell, a pasted
// block, an emptied cell, an accepted destination name — takes a snapshot
// of what it is about to change first. Undo puts the last one back; Reset
// goes the whole way to the file as it was read.
// ─────────────────────────────────────────────────────────────
let editHistory = [];
const EDIT_HISTORY_MAX = 200;

function editSnapshot() {
  return {
    loads:   new Map(previewEdits.loads),
    summary: new Map(previewEdits.summary),
    renames: { ...destRenames },
  };
}

function pushEditHistory() {
  editHistory.push(editSnapshot());
  if (editHistory.length > EDIT_HISTORY_MAX) editHistory.shift();
}

function applySnapshot(s) {
  previewEdits = { loads: new Map(s.loads), summary: new Map(s.summary) };
  destRenames  = { ...s.renames };
}

function undoLastEdit() {
  if (!editHistory.length) return false;
  applySnapshot(editHistory.pop());
  return true;
}

function resetAllEdits() {
  editHistory = [];
  clearPreviewEdits();
  clearDestRenames();
}

function hasAnyChange() {
  return previewEdits.loads.size > 0 || previewEdits.summary.size > 0
    || Object.keys(destRenames).length > 0;
}

// The two buttons live in the Loads band and are wired once, since the band
// is part of the page rather than of the table that gets re-rendered.
function syncBandTools() {
  const undo  = document.getElementById('btnUndoEdit');
  const reset = document.getElementById('btnResetEdits');
  if (undo)  undo.disabled  = editHistory.length === 0;
  if (reset) reset.disabled = !hasAnyChange() && editHistory.length === 0;
}

(function wireBandTools() {
  const undo  = document.getElementById('btnUndoEdit');
  const reset = document.getElementById('btnResetEdits');
  const redraw = () => refreshPreview().catch(err => {
    console.error(err); showStatus('Error: ' + err.message, 'error');
  });
  if (undo) undo.addEventListener('click', () => {
    if (!undoLastEdit()) return;
    showStatus('Undid the last change.', 'info');
    redraw();
  });
  if (reset) reset.addEventListener('click', () => {
    if (!hasAnyChange() && !editHistory.length) return;
    resetAllEdits();
    showStatus('Reset to the file as it was read.', 'info');
    redraw();
  });
})();

// Commits a cell the user has finished editing. Re-renders so the summary and
// the destination split lines follow the change.
function commitCellEdit(td) {
  const key   = `${td.dataset.row}\u0000${td.dataset.col}`;
  const store = previewEdits[td.dataset.table];
  if (!store) return;
  const value = coerceCell(td.textContent);
  const shown = td.dataset.original === undefined ? null : td.dataset.original;
  if (String(value ?? '') === String(shown ?? '')) return;   // nothing changed
  pushEditHistory();
  store.set(key, value);
  refreshPreview().catch(err => { console.error(err); showStatus('Error: ' + err.message, 'error'); });
}

// ─────────────────────────────────────────────────────────────
// Cells behave the way a spreadsheet's do. A cell is selected first and
// edited second: typing on a selected cell replaces the whole value,
// Backspace empties it, and a double-click opens it with the caret where
// it was clicked so one character can be corrected.
// ─────────────────────────────────────────────────────────────
function isEditing(td) { return td.classList.contains('pl-editing'); }

function caretToEnd(el) {
  const r = document.createRange();
  r.selectNodeContents(el);
  r.collapse(false);
  const s = window.getSelection();
  s.removeAllRanges();
  s.addRange(r);
}

function caretAtPoint(x, y) {
  let r = null;
  if (document.caretRangeFromPoint) r = document.caretRangeFromPoint(x, y);
  else if (document.caretPositionFromPoint) {
    const p = document.caretPositionFromPoint(x, y);
    if (p) { r = document.createRange(); r.setStart(p.offsetNode, p.offset); r.collapse(true); }
  }
  if (!r) return false;
  const s = window.getSelection();
  s.removeAllRanges();
  s.addRange(r);
  return true;
}

// `initial` replaces the value outright — that is a typed character
// landing on a selected cell. `point` puts the caret where the mouse was.
// The editor is its own box sitting on top of the cell, not the cell made
// editable. A cell that stretched to hold what was typed pushed its column
// wider and shoved the table about; this keeps the column exactly as it is
// and drops the box downward over the rows below when the text needs room.
function startCellEdit(td, { initial = null, point = null } = {}) {
  if (isEditing(td)) return;
  dropRange();                     // a cell being typed into is not a block
  td.dataset.original = td.textContent;
  const ed = document.createElement('div');
  ed.className       = 'pl-cell-editor';
  ed.contentEditable = 'true';
  ed.spellcheck      = false;
  ed.textContent     = initial != null ? initial : td.textContent;
  td.classList.add('pl-editing');
  // The cell keeps its text underneath, hidden behind the editor: emptying
  // it let the column shrink to nothing the moment editing began.
  td.appendChild(ed);
  placeEditor(td, ed);
  ed.addEventListener('input', () => placeEditor(td, ed));
  ed.focus({ preventScroll: true });
  if (point && caretAtPoint(point.x, point.y)) return;
  caretToEnd(ed);
}

// The box drops down, unless down is off the end of the table — on the last
// rows it opens upward instead. Either way it is never taller than the room
// it has, so it can always be read in full.
function placeEditor(td, ed) {
  const box = td.closest('.pl-out-table');
  if (!box) return;
  ed.style.top = '-2px';
  ed.style.bottom = 'auto';
  ed.style.maxHeight = '';
  const cell = td.getBoundingClientRect();
  const wrap = box.getBoundingClientRect();
  const below = wrap.bottom - cell.top;
  const above = cell.bottom - wrap.top;
  const want  = ed.scrollHeight + 4;
  if (want > below && above > below) {
    ed.style.top = 'auto';
    ed.style.bottom = '-2px';
    ed.style.maxHeight = `${Math.max(48, Math.floor(above - 4))}px`;
  } else {
    ed.style.maxHeight = `${Math.max(48, Math.floor(below - 4))}px`;
  }
}

function endCellEdit(td, { commit = true } = {}) {
  if (!isEditing(td)) return;
  const ed    = td.querySelector('.pl-cell-editor');
  const typed = ed ? ed.textContent : td.textContent;
  if (ed) ed.remove();
  td.classList.remove('pl-editing');
  td.textContent = commit ? typed : (td.dataset.original ?? '');
  if (commit) commitCellEdit(td);
}

// ─────────────────────────────────────────────────────────────
// Drag a rectangle across the preview and copy it, the way a spreadsheet
// does. The range is geometric — row and column positions in the table —
// so it survives nothing but the current render, which is all it needs.
// ─────────────────────────────────────────────────────────────
let rangeAnchor = null;      // { table, r, c }
let rangeFocus  = null;
let rangeDrag   = false;

function cellRC(td) {
  const tr = td.parentElement;
  return { r: [...tr.parentElement.children].indexOf(tr), c: td.cellIndex };
}

function rangeCells() {
  if (!rangeAnchor || !rangeFocus) return [];
  const t = rangeAnchor.table;
  const r1 = Math.min(rangeAnchor.r, rangeFocus.r), r2 = Math.max(rangeAnchor.r, rangeFocus.r);
  const c1 = Math.min(rangeAnchor.c, rangeFocus.c), c2 = Math.max(rangeAnchor.c, rangeFocus.c);
  const out = [];
  const rows = [...t.querySelectorAll('tbody tr')];
  for (let r = r1; r <= r2; r++) {
    const tr = rows[r];
    if (!tr) continue;
    const line = [];
    for (let c = c1; c <= c2; c++) {
      const td = tr.children[c];
      if (td && !td.classList.contains('pl-fill')) line.push(td);
    }
    if (line.length) out.push(line);
  }
  return out;
}

function clearRange() {
  document.querySelectorAll('#tab-packing-list td.pl-range')
    .forEach(td => td.classList.remove('pl-range'));
}

// Forgets the block as well as its paint. Clearing only the paint left the
// copy handler still holding a rectangle, so Ctrl+C inside an open cell
// copied the old block instead of the text being edited.
function dropRange() {
  rangeAnchor = null;
  rangeFocus  = null;
  rangeDrag   = false;
  clearRange();
}

function paintRange() {
  clearRange();
  for (const line of rangeCells()) for (const td of line) td.classList.add('pl-range');
}

// Only a cell of one of the two preview tables, and never the spacer.
function previewCell(target) {
  const td = target && target.closest ? target.closest('#plPreviewData td, #plPreviewSummary td') : null;
  if (!td || td.classList.contains('pl-fill')) return null;
  return td;
}

previewSec.addEventListener('mousedown', e => {
  const td = previewCell(e.target);
  if (!td) return;
  if (isEditing(td)) return;                   // a click inside an open cell
  if (e.button !== 0) return;
  // Taking the mousedown stops the browser selecting the table's text as
  // the pointer moves; the cell is focused by hand instead.
  e.preventDefault();
  const rc = cellRC(td);
  rangeAnchor = { table: td.closest('table'), r: rc.r, c: rc.c };
  rangeFocus  = { r: rc.r, c: rc.c };
  rangeDrag   = true;
  paintRange();
  if (td.classList.contains('pl-edit')) td.focus({ preventScroll: true });
});

previewSec.addEventListener('mouseover', e => {
  if (!rangeDrag) return;
  const td = previewCell(e.target);
  if (!td || td.closest('table') !== rangeAnchor.table) return;
  rangeFocus = cellRC(td);
  paintRange();
});

document.addEventListener('mouseup', () => { rangeDrag = false; });

// The cell a block was started from, looked up fresh: the preview is
// rebuilt on every change, so a held element reference goes stale.
function anchorCell() {
  if (!rangeAnchor || !rangeAnchor.table.isConnected) return null;
  const tr = [...rangeAnchor.table.querySelectorAll('tbody tr')][rangeAnchor.r];
  return tr ? tr.children[rangeAnchor.c] || null : null;
}

// Writes a pasted block into the edit store in one go. Going through the
// per-cell commit would re-render the preview once per value.
function pasteBlock(startTd, text) {
  const table = startTd.closest('table');
  const trs   = [...table.querySelectorAll('tbody tr')];
  const start = cellRC(startTd);
  const grid  = text.replace(/\r\n?/g, '\n').replace(/\n$/, '')
    .split('\n').map(line => line.split('\t'));
  let filled = 0, unplaced = 0, locked = 0;
  pushEditHistory();                       // the whole block undoes at once
  grid.forEach((line, ri) => {
    const tr = trs[start.r + ri];
    if (!tr) { unplaced += line.length; return; }        // past the last row
    line.forEach((val, ci) => {
      const td = tr.children[start.c + ci];
      if (!td || td.classList.contains('pl-fill')) { unplaced++; return; }
      const store = td.classList.contains('pl-edit') && previewEdits[td.dataset.table];
      if (!store) { locked++; return; }                  // the summary is derived
      store.set(`${td.dataset.row}\u0000${td.dataset.col}`, coerceCell(val));
      filled++;
    });
  });
  return { filled, unplaced, locked };
}

document.addEventListener('paste', e => {
  if (previewSec.style.display !== 'block') return;
  const open = document.querySelector('#tab-packing-list td.pl-editing');
  if (open) return;                          // a cell being typed into takes it
  const startTd = anchorCell();
  if (!startTd) return;
  const text = (e.clipboardData || window.clipboardData).getData('text/plain');
  if (!text) return;
  e.preventDefault();
  const { filled, unplaced, locked } = pasteBlock(startTd, text);
  dropRange();
  if (!filled && !unplaced && !locked) return;
  const notes = [];
  if (unplaced) notes.push(`${unplaced} had no cell to land on`);
  if (locked)   notes.push(`${locked} fell on the read-only summary`);
  showStatus(`Pasted ${filled} cell${filled === 1 ? '' : 's'}`
    + (notes.length ? ' \u00b7 ' + notes.join(' \u00b7 ') : '') + '.',
    notes.length ? 'error' : 'success');
  refreshPreview().catch(err => {
    console.error(err); showStatus('Error: ' + err.message, 'error');
  });
});

// Copy the rectangle as tab-separated text, which is what a spreadsheet
// reads back as cells.
document.addEventListener('copy', e => {
  const lines = rangeCells();
  if (!lines.length) return;
  // One cell counts too. The table does not allow a text selection, so
  // leaving a single cell to the browser copied nothing at all.
  const text = lines.map(line => line.map(td => td.textContent.trim()).join('\t')).join('\n');
  e.clipboardData.setData('text/plain', text);
  e.preventDefault();
  const one = lines.length === 1 && lines[0].length === 1;
  showStatus(one ? 'Copied 1 cell.'
    : `Copied ${lines.length} row${lines.length === 1 ? '' : 's'} \u00d7 `
      + `${lines[0].length} column${lines[0].length === 1 ? '' : 's'}.`, 'success');
});

// One set of listeners on the preview, so re-rendering never loses them.
// A plain click only selects: the browser gives the cell focus because it
// carries a tabindex, and nothing opens until a key or a double-click.
previewSec.addEventListener('dblclick', e => {
  const td = e.target.closest('td.pl-edit');
  if (!td || isEditing(td)) return;
  e.preventDefault();
  startCellEdit(td, { point: { x: e.clientX, y: e.clientY } });
});

previewSec.addEventListener('focusout', e => {
  const td = e.target.closest('td.pl-edit');
  if (!td) return;
  // Opening a cell moves the focus from the cell to the editor inside it,
  // which is a focusout on the cell. Closing on that shut the editor in the
  // same breath as opening it, and no cell could be edited at all.
  if (e.relatedTarget && td.contains(e.relatedTarget)) return;
  endCellEdit(td);
});

previewSec.addEventListener('keydown', e => {
  const td = e.target.closest('td.pl-edit');
  if (!td) return;

  if (isEditing(td)) {
    if (e.key === 'Enter')  { e.preventDefault(); endCellEdit(td); }
    if (e.key === 'Escape') { e.preventDefault(); endCellEdit(td, { commit: false }); }
    return;                                   // every other key edits text
  }

  // Selected, not open. These are the spreadsheet's keys.
  if (e.key === 'Backspace' || e.key === 'Delete') {
    e.preventDefault();
    td.dataset.original = td.textContent;
    td.textContent = '';
    commitCellEdit(td);
    return;
  }
  if (e.key === 'Enter' || e.key === 'F2') { e.preventDefault(); startCellEdit(td); return; }
  if (e.key === 'Escape') { td.blur(); return; }
  if (e.key.length === 1 && !e.metaKey && !e.ctrlKey && !e.altKey) {
    e.preventDefault();
    startCellEdit(td, { initial: e.key });
  }
});

// ─────────────────────────────────────────────────────────────
// Cell classification
// Works out what the pipeline will make of every cell in the sheet, so the
// sheet view can show it. Mirrors readExcel(): the same detectors, and a role
// claims the FIRST column carrying its header name.
// ─────────────────────────────────────────────────────────────
const PL_ROLES = [
  { role: 'Destination', cls: 'dest',   label: 'Destination', color: '#2563eb' },
  { role: 'Carton',      cls: 'carton', label: 'Carton',      color: '#16a34a' },
  { role: 'FBA_ID',      cls: 'fba',    label: 'FBA ID',      color: '#d97706' },
  { role: 'REF ID',      cls: 'po',     label: 'REF ID',      color: '#db2777' },
  { role: 'Weight',      cls: 'weight', label: 'Weight',      color: '#7c3aed' },
  { role: 'CMB',         cls: 'cbm',    label: 'CBM',         color: '#0891b2' },
];

// Works out what the pipeline will make of each column in the selected range.
// Returns one entry per ABSOLUTE column so the grid can tint it directly;
// columns outside the range get no tint.
function classifyColumns(data, range) {
  const { hdrRow, lastRow, firstCol, lastCol } = range;
  const headers = sliceRow(data.raw[hdrRow] || [], firstCol, lastCol)
    .map(h => (h != null ? String(h) : null));
  const rows = rangeDataRowIndices(range)
    .map(i => sliceRow(data.raw[i], firstCol, lastCol));

  const autoDict = {
    'Destination': detectDestinationColumn(headers, rows),
    'Carton':      detectCartonColumn(headers),
    'FBA_ID':      detectFBAColumn(headers, rows),
    'REF ID':      detectAmazonPOColumn(headers, rows),
    'Weight':      detectWeightColumn(headers),
    'CMB':         detectCBMColumn(headers),
  };
  const matchDict = { ...autoDict };
  for (const role of Object.keys(matchDict)) {
    if (Object.prototype.hasOwnProperty.call(roleOverrides, role)) matchDict[role] = roleOverrides[role];
  }

  // Mirror readExcel(): a role claims the FIRST column with that header name.
  const roleByLocal = {};
  const roleCol     = {};
  for (const { role } of PL_ROLES) {
    const name = matchDict[role];
    if (!name) { roleCol[role] = null; continue; }
    const i = headers.indexOf(name);
    roleCol[role] = i >= 0 ? i + firstCol : null;
    if (i >= 0 && roleByLocal[i] === undefined) roleByLocal[i] = role;
  }

  const dropped = droppedColumns(headers, firstCol);
  const cols = [];
  for (let c = 0; c < data.maxCols; c++) {
    if (c < firstCol || c > lastCol) { cols.push({ kind: 'outside', cls: '', label: '' }); continue; }
    const local   = c - firstCol;
    const name    = headers[local];
    const hasData = rows.some(r => !isEmpty(r[local]));
    if (roleByLocal[local] !== undefined) {
      const meta = PL_ROLES.find(r => r.role === roleByLocal[local]);
      cols.push({ kind: 'role', role: meta.role, cls: meta.cls, label: meta.label, color: meta.color, name, hasData });
    } else if (dropped.has(c)) {
      cols.push({ kind: 'ignored', cls: 'ignored', label: 'dropped', name, hasData });
    } else if (hasData) {
      cols.push({ kind: 'extra', cls: 'extra', label: 'extra', name, hasData });
    } else {
      cols.push({ kind: 'empty', cls: 'empty', label: 'empty', name, hasData });
    }
  }

  const missing = PL_ROLES.filter(r => !matchDict[r.role]).map(r => r.label);
  // Columns the user can choose from, in sheet order.
  const choices = [];
  for (let c = firstCol; c <= lastCol; c++) {
    choices.push({ col: c, letter: colLetter(c), name: headers[c - firstCol] });
  }
  return { cols, matchDict, autoDict, roleCol, choices, missing };
}

function renderSheetLegend(info, range, data) {
  const { cols, roleCol, missing } = info;
  const counts = {};
  cols.forEach(c => { if (c.kind !== 'outside') counts[c.kind] = (counts[c.kind] || 0) + 1; });

  const blocks  = (range.extras || []);
  const ref     = `${cellRef(range.hdrRow, range.firstCol)}:${cellRef(range.lastRow, range.lastCol)}`
    + blocks.map(b => ` + ${cellRef(b.r1, range.firstCol)}:${cellRef(b.r2, range.lastCol)}`).join('');
  const hdrNo     = range.hdrRow + 1;
  const firstData = range.hdrRow + 2;
  const lastData  = range.lastRow + 1;
  const nData     = countDataRows(data, range);
  const colSpan   = `${colLetter(range.firstCol)}–${colLetter(range.lastCol)}`;

  const mode = range.auto
    ? '<span class="pl-range-mode pl-range-auto">found automatically</span>'
    : '<span class="pl-range-mode pl-range-manual">you chose this</span>';
  const reset = range.auto ? ''
    : '<button type="button" class="pl-range-reset" id="btnResetRange">Reset to auto</button>';

  // What the blue box actually controls, in the sheet's own row numbers.
  const where = blocks.length
    ? `${blocks.length + 1} blocks of rows`
    : `Rows <strong>${firstData}–${lastData}</strong>`;
  const explain = nData
    ? `Row <strong>${hdrNo}</strong> is read as the column header.
       ${where} become the ${nData} line${nData === 1 ? '' : 's'}
       of the packing list. Only columns <strong>${colSpan}</strong> are read.`
    : `Row <strong>${hdrNo}</strong> is read as the column header, but no data rows are selected below it.`;

  const chips = PL_ROLES.map(r => {
    const abs  = roleCol[r.role];
    const col  = abs != null ? cols[abs] : null;
    const name = col && col.name != null && col.name !== '' ? escapeHtml(col.name) : null;
    const set  = abs != null;
    return `
    <button type="button" class="pl-chip pl-chip-btn${set ? '' : ' pl-chip-unset'}"
            data-role="${escapeHtml(r.role)}" aria-haspopup="menu" aria-expanded="false"
            title="Click to choose which column is the ${escapeHtml(r.label)}">
      <i class="pl-dot" style="background:${set ? r.color : '#cbd5e1'}"></i>${escapeHtml(r.label)}
      <em>${set ? (name || 'no header') : 'not set'}</em>
      ${set ? `<span class="pl-chip-col">${colLetter(abs)}</span>` : ''}
      <span class="pl-chip-caret">▾</span>
    </button>`;
  }).join('');

  const asides = [
    counts.extra   ? `${counts.extra} extra column${counts.extra === 1 ? '' : 's'} carried through` : '',
    counts.ignored ? `${counts.ignored} on the ignore list` : '',
    counts.empty   ? `${counts.empty} empty` : '',
  ].filter(Boolean).join(' · ');

  const warn = missing.length ? `
    <div class="pl-legend-warn">
      <strong>No column matched ${missing.map(escapeHtml).join(', ')}.</strong>
      Those values come out blank or zero. Either the blue box is on the wrong rows,
      or this sheet words that column differently — add its wording under Settings.
    </div>` : '';

  const names = (data && data.sheetNames) || [];
  const picker = names.length > 1 ? `
    <span class="pl-sheet-pick">Sheet
      <select id="plSheetPick">
        ${names.map(n => `<option value="${escapeHtml(n)}"${n === data.sheetName ? ' selected' : ''}>${escapeHtml(n)}</option>`).join('')}
      </select>
    </span>` : '';

  const modeBtns = `
    <span class="pl-mode" role="group" aria-label="What a drag does">
      <button type="button" class="pl-mode-btn${selectMode === 'replace' ? ' pl-mode-on' : ''}"
              data-mode="replace">Select</button>
      <button type="button" class="pl-mode-btn${selectMode === 'add' ? ' pl-mode-on' : ''}"
              data-mode="add">Add block</button>
    </span>`;

  return `
    <div class="pl-range-bar">
      <span class="pl-range-ref">${ref}</span>${mode}${modeBtns}${picker}${reset}
    </div>
    <p class="pl-range-explain">${explain}</p>
    <p class="pl-range-hint">Drag across the grid to change it. The top row of the box is always the header.
      Click a row number or column letter to take the whole line.
      When the table is split by a gap, hold <strong>⌘</strong> or <strong>Ctrl</strong>
      and drag the other block to add it.</p>
    <div class="pl-role-row">
      <span class="pl-role-label">Columns found</span>${chips}
      ${asides ? `<span class="pl-role-aside">${asides}</span>` : ''}
    </div>${warn}`;
}

// ─────────────────────────────────────────────────────────────
// Pop-up menu
// Rendered on <body> at page coordinates so the scrolling grid cannot clip it.
// Items: { label, note, checked, disabled, sep, onPick }
// ─────────────────────────────────────────────────────────────
let plMenuEl = null;

// Tears down the menu element only. Opening a new menu replaces the old one
// without disturbing the column marker the caller just set.
function removePlMenuEl() {
  if (plMenuEl) { plMenuEl.remove(); plMenuEl = null; }
  document.querySelectorAll('.pl-chip-on').forEach(c => c.classList.remove('pl-chip-on'));
}

function closePlMenu() {
  removePlMenuEl();
  // The column marker only ever means "this is what the menu acts on".
  if (sheetGrid) sheetGrid.clearFocus();
}

function openPlMenu(x, y, title, items) {
  removePlMenuEl();
  const menu = document.createElement('div');
  menu.className = 'pl-menu';
  menu.setAttribute('role', 'menu');
  menu.innerHTML =
    (title ? `<div class="pl-menu-title">${title}</div>` : '') +
    items.map((it, i) => it.sep
      ? '<div class="pl-menu-sep"></div>'
      : `<button type="button" class="pl-menu-item${it.checked ? ' pl-menu-on' : ''}"
           data-i="${i}"${it.disabled ? ' disabled' : ''} role="menuitem">
           <span class="pl-menu-tick">${it.checked ? '✓' : ''}</span>
           <span class="pl-menu-label">${it.label}</span>
           ${it.note ? `<span class="pl-menu-note">${it.note}</span>` : ''}
         </button>`).join('');
  document.body.appendChild(menu);

  // Keep it on screen.
  const w = menu.offsetWidth, h = menu.offsetHeight;
  const maxX = window.scrollX + document.documentElement.clientWidth  - w - 8;
  const maxY = window.scrollY + document.documentElement.clientHeight - h - 8;
  menu.style.left = Math.max(window.scrollX + 8, Math.min(x, maxX)) + 'px';
  menu.style.top  = Math.max(window.scrollY + 8, Math.min(y, maxY)) + 'px';

  menu.addEventListener('click', e => {
    const btn = e.target.closest('.pl-menu-item');
    if (!btn || btn.disabled) return;
    const it = items[Number(btn.dataset.i)];
    closePlMenu();
    if (it && it.onPick) it.onPick();
  });
  plMenuEl = menu;
}

document.addEventListener('mousedown', e => {
  if (plMenuEl && !e.target.closest('.pl-menu') && !e.target.closest('.pl-chip-btn')) closePlMenu();
});
document.addEventListener('keydown', e => { if (e.key === 'Escape') closePlMenu(); });

// Builds the menu for one column: which role it fills, and whether it is
// carried into the packing list at all.
function columnMenuItems(abs, info) {
  const { cols, roleCol, choices } = info;
  const meta   = choices.find(c => c.col === abs);
  const name   = meta && meta.name != null && meta.name !== '' ? meta.name : null;
  const isDrop = cols[abs] && cols[abs].kind === 'ignored';
  const items  = [];

  for (const r of PL_ROLES) {
    const isThis = roleCol[r.role] === abs;
    items.push({
      label: `Use as ${escapeHtml(r.label)}`,
      checked: isThis,
      onPick: () => {
        if (isThis) delete roleOverrides[r.role];
        else roleOverrides[r.role] = name;   // null header means "not used"
        if (!isThis && name == null) roleOverrides[r.role] = null;
        reapply();
      },
    });
  }
  items.push({ sep: true });
  items.push({
    label: 'Not a mapped column',
    checked: !PL_ROLES.some(r => roleCol[r.role] === abs),
    onPick: () => {
      for (const r of PL_ROLES) if (roleCol[r.role] === abs) roleOverrides[r.role] = null;
      reapply();
    },
  });
  items.push({ sep: true });
  items.push({
    label: 'Keep in packing list',
    note: isDrop ? '' : 'current',
    checked: !isDrop,
    onPick: () => { columnOverrides[abs] = 'keep'; reapply(); },
  });
  items.push({
    label: 'Drop from packing list',
    note: isDrop ? 'current' : 'this file only',
    checked: isDrop,
    onPick: () => { columnOverrides[abs] = 'drop'; reapply(); },
  });

  // Dropping the same wording on every file is a rule, not a decision to
  // repeat. This writes it to the shared ignore list so the next sheet
  // that uses the words drops the column without being asked.
  if (name != null && String(name).trim() !== '' && !alreadyKnown('NO_NEED_COL', name)) {
    const queued = isQueuedToLearn('NO_NEED_COL', name);
    items.push({
      label: 'Always ignore this wording',
      note: queued ? 'queued' : 'saved when the list is generated',
      checked: queued,
      onPick: () => {
        columnOverrides[abs] = 'drop';
        if (queued) {
          unqueueLearn('NO_NEED_COL', name);
          showStatus(`“${String(name)}” will not be added to the ignore list.`, 'info');
        } else {
          queueLearn('NO_NEED_COL', String(name));
          showStatus(`“${String(name)}” will be added to the shared ignore list `
            + 'once the packing list is generated.', 'info');
        }
        reapply();
      },
    });
  }

  if (columnOverrides[abs] !== undefined) {
    items.push({
      label: 'Back to the ignore list',
      onPick: () => { delete columnOverrides[abs]; reapply(); },
    });
  }
  return items;
}

// Menu for a role chip: which column should fill this role.
function roleMenuItems(role, info) {
  const { roleCol, choices, autoDict } = info;
  const current = roleCol[role];
  const items = [{
    label: 'Detect automatically',
    note: autoDict[role] ? escapeHtml(String(autoDict[role])) : 'nothing found',
    checked: !Object.prototype.hasOwnProperty.call(roleOverrides, role),
    onPick: () => { delete roleOverrides[role]; reapply(); },
  }, {
    label: 'Not used',
    checked: Object.prototype.hasOwnProperty.call(roleOverrides, role) && roleOverrides[role] === null,
    onPick: () => { roleOverrides[role] = null; reapply(); },
  }, { sep: true }];

  for (const c of choices) {
    items.push({
      label: `${c.letter} — ${c.name != null && c.name !== '' ? escapeHtml(String(c.name)) : '(no header)'}`,
      checked: current === c.col,
      disabled: c.name == null || c.name === '',
      onPick: () => { roleOverrides[role] = c.name; reapply(); },
    });
  }
  return items;
}

// Re-runs classification and the preview after a menu choice.
function reapply() {
  hideLearnBar();
  refreshSheetView({ keepSelection: true })
    .then(refreshPreview)
    .then(() => loadSheetData(fileInput.files[0]))
    .then(updateLoadStatus)
    .catch(err => { console.error(err); showStatus('Error: ' + err.message, 'error'); });
}

// ─────────────────────────────────────────────────────────────
// Teaching the detector
// Carton, Weight and CBM are found by matching the header wording against
// a shared list. When the user maps one by hand, that wording is worth
// keeping — but the lists are shared with everyone, so it is offered
// rather than saved silently.
// ─────────────────────────────────────────────────────────────
const ROLE_KEYWORD_LIST = { Carton: 'CARTON_KEYWORDS', Weight: 'WEIGHT_KEYWORDS', CMB: 'CBM_KEYWORDS' };

/* Wordings the user has asked the app to learn, held until the packing
   list is actually written. Nothing reaches the shared lists on the
   strength of a menu click alone: a mapping is only worth keeping once
   it has produced a file. */
let learnQueue = [];

function isQueuedToLearn(listKey, name) {
  return learnQueue.some(q => q.listKey === listKey && q.name === String(name));
}

function queueLearn(listKey, name) {
  if (!isQueuedToLearn(listKey, name)) learnQueue.push({ listKey, name: String(name) });
}

function unqueueLearn(listKey, name) {
  learnQueue = learnQueue.filter(q => !(q.listKey === listKey && q.name === String(name)));
}

function clearLearnQueue() { learnQueue = []; }

// Writes the queue to the shared lists. Called once a packing list has
// been generated, never before.
async function flushLearnQueue() {
  if (!learnQueue.length) return null;
  const byList = new Map();
  for (const q of learnQueue) {
    if (alreadyKnown(q.listKey, q.name)) continue;
    if (!byList.has(q.listKey)) byList.set(q.listKey, [...(getCfgList(q.listKey) || [])]);
    byList.get(q.listKey).push(q.name);
  }
  const names = learnQueue.map(q => q.name);
  learnQueue = [];
  if (!byList.size) return null;
  for (const [listKey, values] of byList) {
    setCfgList(listKey, values);
    await persistConfigList(listKey, values);
  }
  renderConfigPanel();
  return names;
}

function alreadyKnown(listKey, name) {
  const norm = v => String(v).normalize('NFKC').toLowerCase().replace(/\s+/g, '').trim();
  return (getCfgList(listKey) || []).some(v => norm(v) === norm(name));
}

// Header wordings the user has mapped by hand that the lists do not know.
function unlearnedMappings(info) {
  const out = [];
  for (const [role, listKey] of Object.entries(ROLE_KEYWORD_LIST)) {
    if (!Object.prototype.hasOwnProperty.call(roleOverrides, role)) continue;
    const name = roleOverrides[role];
    if (!name || alreadyKnown(listKey, name)) continue;
    const meta = PL_ROLES.find(r => r.role === role);
    out.push({ role, label: meta ? meta.label : role, listKey, name });
  }
  // Columns dropped by hand whose wording is not on the shared ignore list.
  for (const [abs, mode] of Object.entries(columnOverrides)) {
    if (mode !== 'drop') continue;
    const col = info.cols[Number(abs)];
    const name = col && col.name;
    if (!name || alreadyKnown('NO_NEED_COL', name)) continue;
    out.push({ role: null, label: 'always ignored', listKey: 'NO_NEED_COL', name });
  }
  return out;
}

function hideLearnBar() {
  const bar = document.getElementById('plLearnBar');
  if (bar) { bar.hidden = true; bar.innerHTML = ''; }
}

function renderLearnBar(info) {
  const bar = document.getElementById('plLearnBar');
  if (!bar) return;
  const items = unlearnedMappings(info);
  if (!items.length) { bar.hidden = true; bar.innerHTML = ''; return; }
  bar.hidden = false;
  bar.innerHTML = items.map((it, i) => `
    <span class="pl-learn-item">
      <span class="pl-learn-text">Remember <strong>${escapeHtml(String(it.name))}</strong>
        ${it.role ? `as a <strong>${escapeHtml(it.label)}</strong> column` : 'as a column to ignore'}?</span>
      <button type="button" class="pl-learn-btn" data-i="${i}">Remember for next time</button>
    </span>`).join('');

  bar.querySelectorAll('.pl-learn-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const it = items[Number(btn.dataset.i)];
      btn.disabled = true;
      btn.textContent = 'Saving…';
      queueLearn(it.listKey, String(it.name));
      try {
        await flushLearnQueue();
        // The wording is known now, so detection finds it without the override.
        if (it.role) delete roleOverrides[it.role];
        showStatus(`“${String(it.name)}” saved. Remove it under Settings if that was wrong.`, 'success');
      } catch (err) {
        console.error(err);
        showStatus('Could not save to the shared lists: ' + err.message, 'error');
      }
      reapply();
    });
  });
}

// The grid is created once and reused for every file.
let sheetGrid = null;

function ensureGrid() {
  if (sheetGrid) return sheetGrid;
  sheetGrid = createSheetGrid({
    mount: sheetView,
    onSelect: (sel, extras) => {
      if (!sel) return;
      const blocks = (extras || []).map(b => ({ r1: b.r1, r2: b.r2 }));
      const rowsMoved = !manualRange || manualRange.hdrRow !== sel.r1
        || manualRange.lastRow !== sel.r2
        || JSON.stringify(manualRange.extras || []) !== JSON.stringify(blocks);
      manualRange = { hdrRow: sel.r1, lastRow: sel.r2,
                      firstCol: sel.c1, lastCol: sel.c2, extras: blocks };
      hdrRowInput.value  = sel.r1 + 1;
      lastRowInput.value = sel.r2 + 1;
      // Edits are keyed by row, so a different row set would misapply them.
      if (rowsMoved) clearPreviewEdits();
      refreshSheetView({ keepSelection: true })
        .then(refreshPreview)
        .then(() => loadSheetData(fileInput.files[0]))
        .then(updateLoadStatus)
        .catch(err => { console.error(err); showStatus('Error: ' + err.message, 'error'); });
    },
  });
  return sheetGrid;
}

// Re-runs classification for the current range and repaints the legend and
// column tints. Rebuilds the grid body only when the file changed.
async function refreshSheetView({ rebuild = false, keepSelection = false } = {}) {
  const file = fileInput.files[0];
  if (!file) return;
  const data  = await loadSheetData(file);
  const range = getRange(data);
  const grid  = ensureGrid();

  const sel    = { r1: range.hdrRow, r2: range.lastRow, c1: range.firstCol, c2: range.lastCol };
  const blocks = (range.extras || [])
    .map(b => ({ r1: b.r1, r2: b.r2, c1: range.firstCol, c2: range.lastCol }));

  // The grid must be told about the extra blocks too, or a reset leaves the
  // old ones drawn over a range that no longer has them.
  if (rebuild) { grid.render(data.raw, sel, data.maxCols); grid.setExtras(blocks); }
  else if (!keepSelection) { grid.setSelection(sel); grid.setExtras(blocks); }

  const info = classifyColumns(data, range);
  grid.setColumnInfo(info.cols);
  sheetLegend.innerHTML = renderSheetLegend(info, range, data);

  // A chip opens the list of columns that could fill its role.
  sheetLegend.querySelectorAll('.pl-chip-btn').forEach(chip => {
    chip.addEventListener('click', e => {
      e.stopPropagation();
      const role = chip.dataset.role;
      const abs  = info.roleCol[role];
      if (abs != null) grid.focusColumn(abs);
      chip.classList.add('pl-chip-on');
      const r = chip.getBoundingClientRect();
      openPlMenu(r.left + window.scrollX, r.bottom + window.scrollY + 4,
                 `Which column is the ${escapeHtml(role)}?`, roleMenuItems(role, info));
    });
  });

  // Right-clicking a column offers the same choices plus keep / drop.
  grid.onColumnContextMenu((abs, x, y, where) => {
    const meta = info.choices.find(c => c.col === abs);
    const name = meta && meta.name != null && meta.name !== '' ? escapeHtml(String(meta.name)) : '(no header)';
    const items = [];

    // A block added on top of the main selection can be taken back out here.
    if (where && where.extraIndex >= 0) {
      const b = grid.getExtras()[where.extraIndex];
      items.push({
        label: `Remove this block (rows ${b.r1 + 1}–${b.r2 + 1})`,
        onPick: () => grid.removeExtra(where.extraIndex),
      }, { sep: true });
    }

    openPlMenu(x, y, `Column ${colLetter(abs)} — ${name}`,
               items.concat(columnMenuItems(abs, info)));
  });

  sheetLegend.querySelectorAll('.pl-mode-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      selectMode = btn.dataset.mode;
      grid.setSelectMode(selectMode);
      sheetLegend.querySelectorAll('.pl-mode-btn').forEach(b =>
        b.classList.toggle('pl-mode-on', b.dataset.mode === selectMode));
    });
  });
  grid.setSelectMode(selectMode);

  const pick = document.getElementById('plSheetPick');
  if (pick) pick.addEventListener('change', () => {
    setActiveSheet(pick.value)
      .catch(err => { console.error(err); showStatus('Error: ' + err.message, 'error'); });
  });

  const resetBtn = document.getElementById('btnResetRange');
  if (resetBtn) resetBtn.addEventListener('click', () => {
    manualRange        = null;
    clearPreviewEdits();
    hdrRowInput.value  = '';
    lastRowInput.value = '';
    refreshSheetView().then(refreshPreview).catch(err => console.error(err));
  });
}

const clientNameInput  = document.getElementById('clientName');
const containerNoInput = document.getElementById('containerNo');
const fileNoInput      = document.getElementById('fileNo');

function readClientContainer() {
  return {
    cName:         (clientNameInput.value  || '').trim(),
    containerName: (containerNoInput.value || '').trim(),
  };
}

// The office's filing id for this container. It is written into the
// packing list's detail block; the download is named after the container.
function readFileNo() {
  return (fileNoInput.value || '').trim();
}

function downloadName(containerName) {
  const base = `PL(${containerName || 'container'})`;
  return base.replace(/[\\/:*?"<>|]/g, '_') + '.xlsx';
}

// ─────────────────────────────────────────────────────────────
// Shipment cells
// Drag a cell by its header to reorder it, so the row can be arranged to
// match whatever order the source sheet uses. The order is remembered per
// browser, and paste-splitting follows whatever order is on screen.
// ─────────────────────────────────────────────────────────────
const CELL_ORDER_KEY = 'plBuilder.shipmentCellOrder';
const cellsRow = document.getElementById('plCells');

function cellEls() { return [...cellsRow.querySelectorAll('.pl-cell')]; }

function saveCellOrder() {
  try {
    localStorage.setItem(CELL_ORDER_KEY, JSON.stringify(cellEls().map(el => el.dataset.cell)));
  } catch (_) { /* private browsing: the order simply will not persist */ }
}

function restoreCellOrder() {
  let saved;
  try { saved = JSON.parse(localStorage.getItem(CELL_ORDER_KEY)); } catch (_) { return; }
  if (!Array.isArray(saved)) return;
  for (const name of saved) {
    const el = cellsRow.querySelector(`.pl-cell[data-cell="${name}"]`);
    if (el) cellsRow.appendChild(el);          // appending in saved order sorts them
  }
}

let draggedCell = null;
cellsRow.addEventListener('dragstart', e => {
  const cell = e.target.closest('.pl-cell');
  if (!cell) return;
  draggedCell = cell;
  cell.classList.add('pl-cell-dragging');
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('text/plain', cell.dataset.cell);
});
cellsRow.addEventListener('dragend', () => {
  if (draggedCell) draggedCell.classList.remove('pl-cell-dragging');
  cellEls().forEach(c => c.classList.remove('pl-cell-over'));
  draggedCell = null;
});
cellsRow.addEventListener('dragover', e => {
  const over = e.target.closest('.pl-cell');
  if (!draggedCell || !over || over === draggedCell) return;
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
  cellEls().forEach(c => c.classList.toggle('pl-cell-over', c === over));
});
cellsRow.addEventListener('drop', e => {
  const over = e.target.closest('.pl-cell');
  if (!draggedCell || !over || over === draggedCell) return;
  e.preventDefault();
  const list = cellEls();
  // Insert before the target when moving left, after it when moving right.
  const movingRight = list.indexOf(draggedCell) < list.indexOf(over);
  over.parentNode.insertBefore(draggedCell, movingRight ? over.nextSibling : over);
  cellEls().forEach(c => c.classList.remove('pl-cell-over'));
  saveCellOrder();
});

restoreCellOrder();

// Two cells copied from Excel arrive as one tab-separated string. Whichever
// box receives that paste, split it across both.
[clientNameInput, containerNoInput].forEach(input => {
  input.addEventListener('paste', e => {
    const text = (e.clipboardData || window.clipboardData).getData('text') || '';
    const parts = text.split('\t').map(t => t.replace(/\r?\n/g, ' ').trim()).filter(Boolean);
    if (parts.length < 2) return;                  // a normal single-value paste
    e.preventDefault();
    // Fill in the order the cells are shown, skipping the file name.
    const targets = cellEls()
      .filter(c => c.dataset.cell !== 'fileNo')
      .map(c => c.querySelector('input'));
    targets.forEach((el, i) => {
      if (parts[i] === undefined) return;
      el.value = parts[i];
      el.dispatchEvent(new Event('input', { bubbles: true }));
    });
  });
});

// The container-detail block that sits at the top of the generated PL sheet.
// Mirrors writePackingList so the preview shows the real output, not a
// fragment of it.
let lastNumDests = 0;
function renderPreviewHead(numDests) {
  lastNumDests = numDests;
  const { cName, containerName } = readClientContainer();
  const blank = v => (v ? escapeHtml(v) : '<span class="pl-out-blank">not set</span>');
  document.getElementById('plPreviewHead').innerHTML = `
    <div class="pl-out-title">PACKING LIST AND DESTUFFING INSTRUCTION (FBA)</div>
    <div class="pl-out-grid">
      <span class="pl-out-k">Client Name</span><span class="pl-out-v">${blank(cName)}</span>
      <span class="pl-out-k">File #</span><span class="pl-out-v">${blank(readFileNo())}</span>
      <span class="pl-out-k">Container #</span><span class="pl-out-v">${blank(containerName)}</span>
      <span class="pl-out-k"># of Destinations</span><span class="pl-out-v">${numDests}</span>
      <span class="pl-out-k">Destuffing Time</span><span class="pl-out-v"><span class="pl-out-blank">filled in by hand</span></span>
    </div>`;
}

// Shows how many hand edits are in play, and lets the user drop them.
function renderEditBar() {
  const bar = document.getElementById('plEditBar');
  if (!bar) return;
  const n = previewEdits.loads.size + previewEdits.summary.size;
  if (!n) { bar.hidden = true; bar.innerHTML = ''; return; }
  bar.hidden = false;
  bar.innerHTML = `<span><strong>${n}</strong> hand edit${n === 1 ? '' : 's'} will be written into the file.</span>
    <button type="button" id="btnClearEdits" class="pl-range-reset">Undo all edits</button>`;
  document.getElementById('btnClearEdits').addEventListener('click', () => {
    clearPreviewEdits();
    refreshPreview().catch(err => console.error(err));
  });
}

// Lists destinations that match no warehouse code, with a short name on
// offer. Nothing changes until the user accepts or types one.
function renderDestPanel(longDests) {
  const panel = document.getElementById('plDestPanel');
  if (!panel) return;
  if (!longDests || !longDests.length) { panel.hidden = true; panel.innerHTML = ''; return; }

  const rows = longDests.map((d, i) => {
    const short = d.applied || '';
    const hint  = d.suggestion || '';
    return `
      <div class="pl-dest-row" data-i="${i}">
        <div class="pl-dest-orig" tabindex="0">${escapeHtml(d.text)}</div>
        <div class="pl-dest-actions">
          <span class="pl-dest-meta">${d.count} row${d.count === 1 ? '' : 's'}</span>
          <input type="text" class="pl-dest-input" value="${escapeHtml(short)}"
                 placeholder="${hint ? escapeHtml(hint) : 'type a short name'}"
                 aria-label="Short name for this destination">
          ${hint ? `<button type="button" class="pl-dest-use" data-name="${escapeHtml(hint)}">Use ${escapeHtml(hint)}</button>` : ''}
          ${d.applied ? '<button type="button" class="pl-dest-clear">Keep original</button>' : ''}
          <button type="button" class="pl-dest-copy">Copy original</button>
        </div>
      </div>`;
  }).join('');

  const n = longDests.length;
  panel.hidden = false;
  panel.innerHTML = `
    <div class="pl-dest-head">${n} destination${n === 1 ? '' : 's'} did not match a warehouse code.
      Give them a short name, or leave them as they are.</div>
    ${rows}`;

  const commit = (i, name) => {
    const d = longDests[i];
    if ((destRenames[d.text] || '') === (name || '')) return;   // nothing changed
    pushEditHistory();
    if (name) destRenames[d.text] = name; else delete destRenames[d.text];
    refreshPreview().catch(err => { console.error(err); showStatus('Error: ' + err.message, 'error'); });
  };

  panel.querySelectorAll('.pl-dest-row').forEach(row => {
    const i     = Number(row.dataset.i);
    const input = row.querySelector('.pl-dest-input');
    input.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
      if (e.key === 'Escape') { input.value = longDests[i].applied || ''; input.blur(); }
    });
    input.addEventListener('blur', () => {
      const v = input.value.trim();
      if (v !== (longDests[i].applied || '')) commit(i, v);
    });
    const use = row.querySelector('.pl-dest-use');
    if (use) use.addEventListener('click', () => commit(i, use.dataset.name));
    const clr = row.querySelector('.pl-dest-clear');
    if (clr) clr.addEventListener('click', () => commit(i, ''));

    // Copying the original text is how these get pasted back into a sheet or
    // looked up, so make it one click rather than a careful drag-select.
    const copy = row.querySelector('.pl-dest-copy');
    copy.addEventListener('click', async () => {
      const text = longDests[i].text;
      let ok = true;
      try {
        if (navigator.clipboard && window.isSecureContext) await navigator.clipboard.writeText(text);
        else throw new Error('no clipboard api');
      } catch (_) {
        // No clipboard permission, or an http:// origin. Select the text in
        // place so the Ctrl+C the message asks for actually does something.
        ok = false;
        const box = row.querySelector('.pl-dest-orig');
        const range = document.createRange();
        range.selectNodeContents(box);
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);
        box.focus({ preventScroll: true });
      }
      copy.textContent = ok ? 'Copied' : 'Selected — press Ctrl+C';
      copy.classList.toggle('pl-dest-copied', ok);
      setTimeout(() => {
        copy.textContent = 'Copy original';
        copy.classList.remove('pl-dest-copied');
      }, 1600);
    });
  });
}

function renderPreview(modH, modR, sumH, sumR, longDests) {
  const { headers: visH, rows: visRows } = getVisibleColumns(modH, modR);

  // The workbook writes nine summary columns; show those, not the
  // pipeline's internal names, so the preview is the document.
  const sumIdx  = ['Destination', 'Carton', 'pallet_round_count'].map(h => sumH.indexOf(h));
  const sumVisH = PL_SUMMARY_HEADERS;
  const sumVisR = sumR.map(r => {
    const line = sumIdx.map(i => (i >= 0 ? r[i] : null));
    while (line.length < PL_SUMMARY_HEADERS.length) line.push(null);
    return line;
  });

  // Rows where the Destination changes on the next row end a "part".
  const di = visH.indexOf('Destination');
  const splitAfter = new Set();
  if (di >= 0) {
    visRows.forEach((r, i) => {
      if (i < visRows.length - 1 && visRows[i + 1][di] !== r[di]) splitAfter.add(i);
    });
  }

  const numDests = di >= 0
    ? new Set(visRows.map(r => r[di]).filter(v => v != null)).size
    : 0;

  renderPreviewHead(numDests);

  // A colour per destination, assigned in the order they appear and shared
  // with the summary, so the same code is the same colour in both tables.
  const DEST_TINTS = 10;
  const tintOf = new Map();
  const tintFor = v => {
    const key = v == null ? '' : String(v);
    if (!tintOf.has(key)) tintOf.set(key, `pl-dest-c${tintOf.size % DEST_TINTS}`);
    return tintOf.get(key);
  };
  const loadTints = di >= 0 ? visRows.map(r => tintFor(r[di])) : null;

  const idIdx  = modH.indexOf('_Row');
  const rowIds = idIdx >= 0 ? modR.map(r => r[idIdx]) : null;
  document.getElementById('plPreviewData').innerHTML    =
    tableHtml(visH, visRows, { splitAfter, editKey: 'loads', rowIds, tints: loadTints });
  // The summary is read-only: every figure in it is derived from the loads
  // table, so it is corrected by editing the load rows, not the totals.
  document.getElementById('plPreviewSummary').innerHTML =
    tableHtml(sumVisH, sumVisR, { tints: sumVisR.map(r => tintFor(r[0])) });
  renderEditBar();
  syncBandTools();
  renderDestPanel(longDests);
  previewSec.style.display = 'block';
}

// Client name, container number and file number appear in the preview header,
// so keep it live as they are typed.
[clientNameInput, containerNoInput, fileNoInput].forEach(input => {
  input.addEventListener('input', () => {
    if (previewSec.style.display !== 'block') return;
    if (!document.getElementById('plPreviewHead')) return;
    renderPreviewHead(lastNumDests);
  });
});

// ─────────────────────────────────────────────────────────────
// Sheet grid: show / hide, and manual range entry
// ─────────────────────────────────────────────────────────────
btnViewSheet.addEventListener('click', async () => {
  const file = fileInput.files[0];
  if (!file) { showStatus('Please select an Excel file first.', 'error'); return; }

  if (sheetView.style.display === 'block') {
    sheetView.style.display   = 'none';
    sheetLegend.style.display = 'none';
    setSheetToggleLabel(false);
    return;
  }

  try {
    sheetView.style.display   = 'block';
    sheetLegend.style.display = 'block';
    setSheetToggleLabel(true);
    await refreshSheetView({ rebuild: true });
  } catch (err) {
    console.error(err);
    showStatus('Error: ' + err.message, 'error');
  }
});

// Typing a row number is the other way to set the range; it keeps whatever
// column span is currently selected.
[hdrRowInput, lastRowInput].forEach(input => {
  input.addEventListener('input', async () => {
    const file = fileInput.files[0];
    if (!file) return;
    const h = parseInt(hdrRowInput.value, 10);
    const l = parseInt(lastRowInput.value, 10);
    clearPreviewEdits();
    if (isNaN(h) && isNaN(l)) { manualRange = null; }
    else {
      const data = await loadSheetData(file);
      const cur  = manualRange || autoRange(data);
      const hdr  = isNaN(h) ? cur.hdrRow : h - 1;
      manualRange = {
        hdrRow:   hdr,
        lastRow:  isNaN(l) ? Math.max(cur.lastRow, hdr) : l - 1,
        firstCol: cur.firstCol,
        lastCol:  cur.lastCol,
      };
    }
    if (sheetView.style.display === 'block') await refreshSheetView();
    await refreshPreview();
  });
});

// ─────────────────────────────────────────────────────────────
// Preview
// ─────────────────────────────────────────────────────────────
document.getElementById('btnPreview').addEventListener('click', async () => {
  const file = fileInput.files[0];
  if (!file) { showStatus('Please select an Excel file.', 'error'); return; }

  try {
    showStatus('Building preview…', 'info', true);
    const { modH, modR, sumH, sumR, longDests } = await runPipeline(file);
    renderPreview(modH, modR, sumH, sumR, longDests);
    showStatus('✓ Preview generated.', 'success');
  } catch (err) {
    console.error(err);
    showStatus('Error: ' + err.message, 'error');
  }
});

// ─────────────────────────────────────────────────────────────
// Generate
// ─────────────────────────────────────────────────────────────
document.getElementById('plForm').addEventListener('submit', async e => {
  e.preventDefault();
  const file = fileInput.files[0];
  if (!file) { showStatus('Please select an Excel file.', 'error'); return; }

  const { cName, containerName } = readClientContainer();
  if (!cName || !containerName) {
    showStatus('Enter both the client name and the container number.', 'error');
    (cName ? containerNoInput : clientNameInput).focus();
    return;
  }

  const btn    = document.getElementById('submitBtn');
  btn.disabled = true;

  try {
    showStatus('Reading file…', 'info', true);
    const { wbIn, modH, modR, sumH, sumR } = await runPipeline(file);

    showStatus('Building Excel file…', 'info', true);
    await writePackingList(modH, modR, sumH, sumR, cName, containerName, wbIn,
                           downloadName(containerName), readFileNo());

    showStatus('✓ Packing list downloaded successfully.', 'success');

    // The file exists, so anything the user asked the app to learn is now
    // worth keeping. Nothing was written to the shared lists before this.
    try {
      const saved = await flushLearnQueue();
      if (saved && saved.length) {
        showStatus(`✓ Packing list downloaded. Saved ${saved.join(', ')} to the `
          + 'shared lists — remove under Settings if that was wrong.', 'success');
      }
    } catch (err) {
      console.error(err);
      showStatus('Packing list downloaded, but the shared lists could not be '
        + 'updated: ' + err.message, 'error');
    }

    // The mapping produced a real file, so now it is worth keeping.
    const data  = await loadSheetData(file);
    renderLearnBar(classifyColumns(data, getRange(data)));
  } catch (err) {
    console.error(err);
    showStatus('Error: ' + err.message, 'error');
  } finally {
    btn.disabled = false;
  }
});
