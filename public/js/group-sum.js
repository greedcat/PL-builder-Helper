  const PREVIEW_ROWS = 200;
  let workbook  = null;
  let sheetData = [];
  let allCols   = [];
  let groupCols = [];   // ordered list of selected group-by columns

  // ── file drop / click ──────────────────────────────────────────────────────
  const dropZone  = document.getElementById('drop-zone');
  const fileInput = document.getElementById('file-input');
  const status    = document.getElementById('status');

  dropZone.addEventListener('click', () => fileInput.click());
  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault(); dropZone.classList.remove('drag-over');
    handleFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', () => handleFile(fileInput.files[0]));

  function handleFile(file) {
    if (!file) return;
    status.className = '';
    status.textContent = `Loading "${file.name}" …`;
    const reader = new FileReader();
    reader.onload = e => {
      try {
        workbook = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
        populateSheets();
        status.textContent = `✔ Loaded "${file.name}"`;
      } catch(err) {
        status.className = 'error';
        status.textContent = '✖ Failed to read file: ' + err.message;
      }
    };
    reader.readAsBinaryString(file);
  }

  // ── sheet selector ─────────────────────────────────────────────────────────
  const sheetSelectWrap = document.getElementById('sheet-select-wrap');
  const sheetSelect     = document.getElementById('sheet-select');

  function populateSheets() {
    sheetSelect.innerHTML = '';
    workbook.SheetNames.forEach(name => {
      const opt = document.createElement('option');
      opt.value = opt.textContent = name;
      sheetSelect.appendChild(opt);
    });
    sheetSelectWrap.style.display = workbook.SheetNames.length > 1 ? 'flex' : 'none';
    document.getElementById('controls').style.display = 'flex';
    loadSheet(workbook.SheetNames[0]);
  }

  sheetSelect.addEventListener('change', () => loadSheet(sheetSelect.value));

  function loadSheet(name) {
    const ws = workbook.Sheets[name];
    sheetData = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (!sheetData.length) { status.className='error'; status.textContent='Sheet is empty.'; return; }
    allCols = Object.keys(sheetData[0]);
    groupCols = [];
    populateColumns(allCols);
    renderPreview();
    document.getElementById('result-card').style.display = 'none';
  }

  // ── multi-select group-by tags ─────────────────────────────────────────────
  const groupTagsBox = document.getElementById('group-tags');
  const groupPicker  = document.getElementById('group-picker');
  const runBtn       = document.getElementById('run-btn');
  const sumColSel    = document.getElementById('sum-col');

  function populateColumns(cols) {
    // group picker
    groupPicker.innerHTML = '<option value="">+ add column</option>';
    cols.forEach(c => groupPicker.appendChild(new Option(c, c)));

    // sum picker
    sumColSel.innerHTML = '<option value="">All numeric</option>';
    cols.forEach(c => sumColSel.appendChild(new Option(c, c)));

    // clear tags
    [...groupTagsBox.querySelectorAll('.tag')].forEach(t => t.remove());
    groupCols = [];
    runBtn.disabled = true;
  }

  groupPicker.addEventListener('change', () => {
    const val = groupPicker.value;
    if (!val || groupCols.includes(val)) { groupPicker.value = ''; return; }
    addGroupTag(val);
    groupPicker.value = '';
  });

  function addGroupTag(col) {
    groupCols.push(col);

    const tag = document.createElement('span');
    tag.className = 'tag';
    tag.dataset.col = col;
    tag.innerHTML = `${col} <button class="remove" title="Remove">×</button>`;
    tag.querySelector('.remove').addEventListener('click', () => {
      groupCols = groupCols.filter(c => c !== col);
      tag.remove();
      runBtn.disabled = groupCols.length === 0;
    });
    groupTagsBox.insertBefore(tag, groupPicker);
    runBtn.disabled = false;
  }

  // ── data preview ───────────────────────────────────────────────────────────
  function renderPreview() {
    const previewCard  = document.getElementById('preview-card');
    const previewHead  = document.getElementById('preview-head');
    const previewBody  = document.getElementById('preview-body');
    const previewCount = document.getElementById('preview-count');

    previewHead.innerHTML = '<tr>' + allCols.map(c => `<th>${c}</th>`).join('') + '</tr>';
    previewBody.innerHTML = '';

    const rows = sheetData.slice(0, PREVIEW_ROWS);
    rows.forEach(row => {
      const tr = document.createElement('tr');
      tr.innerHTML = allCols.map(c => {
        const v = row[c];
        const isNum = typeof v === 'number';
        return `<td class="${isNum ? 'num' : ''}">${isNum ? fmt(v) : v}</td>`;
      }).join('');
      previewBody.appendChild(tr);
    });

    previewCount.textContent =
      sheetData.length <= PREVIEW_ROWS
        ? `${sheetData.length} rows`
        : `showing ${PREVIEW_ROWS} of ${sheetData.length} rows`;

    previewCard.style.display = 'block';
  }

  // ── run group-by ───────────────────────────────────────────────────────────
  runBtn.addEventListener('click', runGroupBy);

  function runGroupBy() {
    if (!groupCols.length) return;
    const sumCol = sumColSel.value;

    const numericCols = sumCol
      ? [sumCol]
      : allCols.filter(k =>
          !groupCols.includes(k) &&
          sheetData.some(r => typeof r[k] === 'number' || (!isNaN(parseFloat(r[k])) && r[k] !== ''))
        );

    if (!numericCols.length) {
      status.className = 'error';
      status.textContent = '✖ No numeric columns found to sum.';
      return;
    }

    const map = new Map();
    sheetData.forEach(row => {
      const key = groupCols.map(c => String(row[c] ?? '')).join('\x00');
      if (!map.has(key)) {
        const init = {};
        groupCols.forEach(c => init[c] = row[c] ?? '');
        numericCols.forEach(c => init[c] = 0);
        map.set(key, init);
      }
      const acc = map.get(key);
      numericCols.forEach(c => {
        const v = parseFloat(row[c]);
        if (!isNaN(v)) acc[c] += v;
      });
    });

    const resultRows = [...map.values()].sort((a, b) => {
      for (const c of groupCols) {
        const cmp = String(a[c]).localeCompare(String(b[c]));
        if (cmp !== 0) return cmp;
      }
      return 0;
    });

    const totals = {};
    groupCols.forEach((c, i) => totals[c] = i === 0 ? 'TOTAL' : '');
    numericCols.forEach(c => totals[c] = resultRows.reduce((s, r) => s + r[c], 0));

    renderResult(groupCols, numericCols, resultRows, totals);
    status.className = '';
    status.textContent = '';
  }

  // ── render result table ────────────────────────────────────────────────────
  let exportData = [];

  function fmt(n) {
    if (typeof n !== 'number') return n;
    return n % 1 === 0
      ? n.toLocaleString()
      : n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function renderResult(gCols, nCols, rows, totals) {
    const cols = [...gCols, ...nCols];
    exportData = [cols, ...rows.map(r => cols.map(c => r[c])), cols.map(c => totals[c])];

    document.getElementById('result-head').innerHTML =
      '<tr>' + cols.map((c, i) => `<th${i < gCols.length ? '' : ' style="text-align:right"'}>${c}</th>`).join('') + '</tr>';

    const tbody = document.getElementById('result-body');
    tbody.innerHTML = '';

    rows.forEach(row => {
      const tr = document.createElement('tr');
      tr.innerHTML = cols.map((c, i) =>
        `<td class="${i >= gCols.length ? 'num' : ''}">${i >= gCols.length ? fmt(row[c]) : row[c]}</td>`
      ).join('');
      tbody.appendChild(tr);
    });

    const tr = document.createElement('tr');
    tr.className = 'total-row';
    tr.innerHTML = cols.map((c, i) =>
      `<td class="${i >= gCols.length ? 'num' : ''}">${i >= gCols.length ? fmt(totals[c]) : totals[c]}</td>`
    ).join('');
    tbody.appendChild(tr);

    document.getElementById('result-title').textContent =
      `Grouped by: ${gCols.join(' › ')}`;
    document.getElementById('row-count').textContent = `${rows.length} groups`;
    const card = document.getElementById('result-card');
    card.style.display = 'block';
    card.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // ── export CSV ─────────────────────────────────────────────────────────────
  document.getElementById('export-btn').addEventListener('click', () => {
    if (!exportData.length) return;
    const csv = exportData.map(row =>
      row.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')
    ).join('\r\n');
    const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'grouped_result.csv';
    a.click();
  });
