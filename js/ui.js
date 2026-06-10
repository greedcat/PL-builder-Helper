// ─────────────────────────────────────────────────────────────
// UI
// ─────────────────────────────────────────────────────────────
const dropZone     = document.getElementById('dropZone');
const fileInput    = document.getElementById('excelFile');
const fileBadge    = document.getElementById('fileBadge');
const cfgDropZone  = document.getElementById('cfgDropZone');
const cfgFileInput = document.getElementById('cfgFile');
const cfgBadge     = document.getElementById('cfgBadge');
const statusEl     = document.getElementById('status');
const hdrRowInput  = document.getElementById('hdrRowInput');
const lastRowInput = document.getElementById('lastRowInput');
const sheetView    = document.getElementById('plOriginalSheet');
const previewSec   = document.getElementById('plPreview');

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

// Config file drop zone
cfgDropZone.addEventListener('click',     () => cfgFileInput.click());
cfgDropZone.addEventListener('dragover',  e  => { e.preventDefault(); cfgDropZone.classList.add('active'); });
cfgDropZone.addEventListener('dragleave', () => cfgDropZone.classList.remove('active'));
cfgDropZone.addEventListener('drop', e => {
  e.preventDefault();
  cfgDropZone.classList.remove('active');
  if (e.dataTransfer.files[0]) loadConfigFile(e.dataTransfer.files[0]);
});
cfgFileInput.addEventListener('change', () => { if (cfgFileInput.files[0]) loadConfigFile(cfgFileInput.files[0]); });

function setExcelFile(file, dt) {
  if (dt) { const dtp = new DataTransfer(); dtp.items.add(file); fileInput.files = dtp.files; }
  fileBadge.innerHTML     = `<span>📄</span> ${file.name}`;
  fileBadge.style.display = 'inline-flex';
  dropZone.querySelector('.drop-text').style.display = 'none';

  // A new file invalidates the cached sheet data, row range and previews.
  cachedFile      = null;
  cachedSheetData = null;
  hdrRowInput.value  = '';
  lastRowInput.value = '';
  sheetView.style.display  = 'none';
  sheetView.innerHTML       = '';
  previewSec.style.display = 'none';
}

function loadConfigFile(file) {
  const reader  = new FileReader();
  reader.onload = e => {
    applyConfig(e.target.result);
    cfgBadge.innerHTML     = `<span>⚙️</span> ${file.name} loaded`;
    cfgBadge.style.display = 'inline-flex';
    document.getElementById('cfgDropText').style.display = 'none';
  };
  reader.readAsText(file, 'utf-8');
}

function showStatus(msg, type = 'info', spinner = false) {
  statusEl.className     = type;
  statusEl.style.display = 'block';
  statusEl.innerHTML     = spinner ? `<span class="spinner"></span>${msg}` : msg;
}

// ─────────────────────────────────────────────────────────────
// Reads and caches the workbook + cleaned (non-blank) rows.
// ─────────────────────────────────────────────────────────────
async function loadSheetData(file) {
  if (cachedFile === file && cachedSheetData) return cachedSheetData;
  const buf     = await file.arrayBuffer();
  const wbIn    = XLSX.read(buf, { type: 'array', raw: true });
  const sheet   = wbIn.Sheets[wbIn.SheetNames[0]];
  const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
  const cleaned = rawRows.filter(r => r.some(v => !isEmpty(v)));
  cachedFile      = file;
  cachedSheetData = { wbIn, cleaned };
  return cachedSheetData;
}

// Resolves the header row and last data row, honoring manual overrides
// (1-based row numbers, as shown in the "View Original Sheet" table).
function getRowRange(cleaned) {
  const hdrOverride  = hdrRowInput.value.trim();
  const lastOverride = lastRowInput.value.trim();

  const hdrRow = hdrOverride
    ? Math.min(Math.max(parseInt(hdrOverride, 10) - 1, 0), cleaned.length - 1)
    : detectHeader(cleaned);

  const lastRow = lastOverride
    ? Math.min(Math.max(parseInt(lastOverride, 10) - 1, hdrRow), cleaned.length - 1)
    : detectTableEnd(cleaned, hdrRow)[0];

  return { hdrRow, lastRow };
}

// Runs detection + processing, returning everything needed to either
// preview or write the packing list.
async function runPipeline(file) {
  const { wbIn, cleaned }  = await loadSheetData(file);
  const { hdrRow, lastRow } = getRowRange(cleaned);
  console.log('Header row:', hdrRow, cleaned[hdrRow]);
  console.log('End row:',    lastRow, cleaned[lastRow]);

  const headers = cleaned[hdrRow].map(h => h != null ? String(h) : null);
  const rows    = cleaned.slice(hdrRow + 1, lastRow + 1);

  const { headers: reH, rows: reR, matchDict } = readExcel(headers, rows);
  console.log('Column map:', matchDict);

  const { headers: modH, rows: modR } = modifyDF(reH, reR, matchDict);
  const { headers: sumH, rows: sumR } = getSummaryDF(modH, modR);

  return { wbIn, cleaned, hdrRow, lastRow, modH, modR, sumH, sumR };
}

// ─────────────────────────────────────────────────────────────
// Table rendering helpers
// ─────────────────────────────────────────────────────────────
function tableHtml(headers, rows) {
  let html = '<table><thead><tr>';
  headers.forEach(h => { html += `<th>${escapeHtml(h ?? '')}</th>`; });
  html += '</tr></thead><tbody>';
  rows.forEach(row => {
    html += '<tr>';
    row.forEach(v => { html += `<td>${v != null ? escapeHtml(String(v)) : ''}</td>`; });
    html += '</tr>';
  });
  html += '</tbody></table>';
  return html;
}

function renderOriginalSheet(cleaned, hdrRow, lastRow) {
  const maxCols = cleaned.reduce((m, r) => Math.max(m, r.length), 1);
  let html = '<table><thead><tr><th>#</th>';
  for (let c = 0; c < maxCols; c++) html += `<th>Col ${c + 1}</th>`;
  html += '</tr></thead><tbody>';
  cleaned.forEach((row, i) => {
    let cls = '';
    if (i === hdrRow)                     cls = 'pl-row-header';
    else if (i > hdrRow && i <= lastRow)  cls = 'pl-row-data';
    html += `<tr class="${cls}"><td class="pl-rownum">${i + 1}</td>`;
    for (let c = 0; c < maxCols; c++) {
      const v = row[c];
      html += `<td>${v != null ? escapeHtml(String(v)) : ''}</td>`;
    }
    html += '</tr>';
  });
  html += '</tbody></table>';
  sheetView.innerHTML = html;
}

function renderPreview(modH, modR, sumH, sumR) {
  const { headers: visH, rows: visRows } = getVisibleColumns(modH, modR);

  const sumVis  = sumH.map((h, i) => ({ h, i })).filter(x => x.h !== 'pallet_number');
  const sumVisH = sumVis.map(x => x.h);
  const sumVisR = sumR.map(r => sumVis.map(x => r[x.i]));

  document.getElementById('plPreviewData').innerHTML    = tableHtml(visH, visRows);
  document.getElementById('plPreviewSummary').innerHTML = tableHtml(sumVisH, sumVisR);
  previewSec.style.display = 'block';
}

// ─────────────────────────────────────────────────────────────
// View Original Sheet
// ─────────────────────────────────────────────────────────────
document.getElementById('btnViewSheet').addEventListener('click', async () => {
  const file = fileInput.files[0];
  if (!file) { showStatus('Please select an Excel file first.', 'error'); return; }

  if (sheetView.style.display === 'block') {
    sheetView.style.display = 'none';
    return;
  }

  try {
    const { cleaned } = await loadSheetData(file);
    const { hdrRow, lastRow } = getRowRange(cleaned);
    renderOriginalSheet(cleaned, hdrRow, lastRow);
    sheetView.style.display = 'block';
  } catch (err) {
    console.error(err);
    showStatus('Error: ' + err.message, 'error');
  }
});

// Re-highlight the original sheet view when the manual row range changes.
[hdrRowInput, lastRowInput].forEach(input => {
  input.addEventListener('input', async () => {
    if (sheetView.style.display !== 'block') return;
    const file = fileInput.files[0];
    if (!file) return;
    const { cleaned } = await loadSheetData(file);
    const { hdrRow, lastRow } = getRowRange(cleaned);
    renderOriginalSheet(cleaned, hdrRow, lastRow);
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
    const { modH, modR, sumH, sumR } = await runPipeline(file);
    renderPreview(modH, modR, sumH, sumR);
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

  const clientContainer = document.getElementById('clientContainer').value;
  const [cName = '', containerName = ''] = clientContainer.split('\t').map(s => s.trim());
  if (!cName || !containerName) {
    showStatus('Please enter Client Name and Container # separated by a tab.', 'error');
    return;
  }

  const btn    = document.getElementById('submitBtn');
  btn.disabled = true;

  try {
    showStatus('Reading file…', 'info', true);
    const { wbIn, modH, modR, sumH, sumR } = await runPipeline(file);

    showStatus('Building Excel file…', 'info', true);
    await writePackingList(modH, modR, sumH, sumR, cName, containerName, wbIn);

    showStatus('✓ Packing list downloaded successfully.', 'success');
  } catch (err) {
    console.error(err);
    showStatus('Error: ' + err.message, 'error');
  } finally {
    btn.disabled = false;
  }
});
