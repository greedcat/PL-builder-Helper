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

renderConfigPanel();

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

document.getElementById('plForm').addEventListener('submit', async e => {
  e.preventDefault();
  const file          = fileInput.files[0];
  const cName         = document.getElementById('cName').value.trim();
  const containerName = document.getElementById('containerName').value.trim();
  if (!file) { showStatus('Please select an Excel file.', 'error'); return; }

  const btn    = document.getElementById('submitBtn');
  btn.disabled = true;

  try {
    showStatus('Reading file…', 'info', true);
    const buf     = await file.arrayBuffer();
    const wbIn    = XLSX.read(buf, { type: 'array', raw: true });
    const sheet   = wbIn.Sheets[wbIn.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
    const cleaned = rawRows.filter(r => r.some(v => !isEmpty(v)));

    showStatus('Detecting table structure…', 'info', true);
    const hdrRow    = detectHeader(cleaned);
    const [lastRow] = detectTableEnd(cleaned, hdrRow);
    console.log('Header row:', hdrRow, cleaned[hdrRow]);
    console.log('End row:',    lastRow, cleaned[lastRow]);

    const headers = cleaned[hdrRow].map(h => h != null ? String(h) : null);
    const rows    = cleaned.slice(hdrRow + 1, lastRow + 1);

    showStatus('Detecting columns…', 'info', true);
    const { headers: reH, rows: reR, matchDict } = readExcel(headers, rows);

    showStatus('Processing data…', 'info', true);
    const { headers: modH, rows: modR } = modifyDF(reH, reR, matchDict);
    const { headers: sumH, rows: sumR } = getSummaryDF(modH, modR);

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
