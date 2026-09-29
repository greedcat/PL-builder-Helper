// ── Packing List Builder UI · part 1: page wiring, workbook read, data range, table helpers ──
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
  hdrRowInput.value  = '';
  lastRowInput.value = '';
  sheetView.style.display   = 'none';
  sheetView.innerHTML       = '';
  sheetLegend.style.display = 'none';
  sheetLegend.innerHTML     = '';
  previewSec.style.display  = 'none';
  hideLearnBar();

  setSteps({ upload: ['done', file.name], read: 'active', table: 'todo', preview: 'todo', download: 'todo' });
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
      setSteps({ read: ['error', 'No readable rows'] });
      showStatus('That file has no readable rows. Check it opens in Excel and is not password protected.', 'error');
      return;
    }

    setSteps({ read: ['done', `Sheet “${data.sheetName}”`], table: 'active' });
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
    setSteps({ read: ['error', 'Could not read the file'] });
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

  const s        = nData === 1 ? '' : 's';
  const rowsText = `${nData} row${s}`;
  if (nData < 1) {
    setSteps({ table: ['error', 'No data rows found'] });
    showStatus('Sheet loaded, but no data rows were found below the header. Drag the box\'s corners on the grid, or highlight the table and right-click → Use as data range.', 'error');
  } else if (missing.length === PL_ROLES.length) {
    setSteps({ table: ['error', 'No columns recognised'] });
    showStatus('Sheet loaded, but no columns were recognised. The header row is probably wrong — highlight the table on the grid and right-click → Use as data range.', 'error');
  } else if (missing.length) {
    setSteps({ table: ['warn', `${rowsText} · pick ${missing.join(', ')}`] });
    showStatus(`${nData} data row${s} selected, but no column matched ${missing.join(', ')}. `
             + 'Click that chip to pick its column, or right-click the column letter in the grid.', 'error');
  } else {
    setSteps({ table: ['done', `${rowsText} · all columns`] });
    showStatus(`✓ ${nData} data row${s} selected, all columns matched.`, 'success');
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
  // value in every row of a merged block. Which cells that filled is kept,
  // because a merged quantity must still be counted only once.
  const mergeFilled = fillMergedRows(wbIn.Sheets[activeSheet], raw);

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
                      raw, cleaned, cleanedIdx, maxCols, mergeFilled };
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
// the wrong data. Drop them rather than silently misapply them — and drop
// the undo history with them, which holds the same row keys.
function clearPreviewEdits() {
  previewEdits = { loads: new Map(), summary: new Map() };
  clearEditHistory();
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
  const dataIdx = rangeDataRowIndices(range);
  const rows = dataIdx.map(i => sliceRow(data.raw[i], firstCol, lastCol));

  // Where a merged block wrote a copy of its value, in this slice's own
  // coordinates. A quantity is one figure for its whole block, so the copies
  // must not be added up again — but which columns are quantities is only
  // known once the columns have been recognised, so the mask travels with
  // the rows rather than being decided here.
  const mergeFilled = data.mergeFilled || new Set();
  const mergedMask = mergeFilled.size
    ? dataIdx.map(i => {
        const line = [];
        for (let c = firstCol; c <= lastCol; c++) line.push(mergeFilled.has(i + ':' + c));
        return line;
      })
    : null;

  const dropAbs   = droppedColumns(headers, firstCol);
  const dropLocal = new Set([...dropAbs].map(c => c - firstCol));
  const { headers: reH, rows: reR, matchDict } =
    readExcel(headers, rows, roleOverrides, dropLocal, mergedMask);
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
