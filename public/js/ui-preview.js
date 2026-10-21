// ── Packing List Builder UI · part 4: shipment cells, sheet grid toggle, preview, generate ──
// Shipment cells
// The order is fixed — Container #, Client name, File number — the order
// the office copies them from its own sheet, so a copied line pastes
// straight in.
// ─────────────────────────────────────────────────────────────
const cellsRow = document.getElementById('plCells');

function cellEls() { return [...cellsRow.querySelectorAll('.pl-cell')]; }

// Cells copied from Excel arrive as one tab-separated string. Whichever box
// receives that paste, fill it and the boxes after it, one value each.
cellsRow.addEventListener('paste', e => {
  const input = e.target.closest('.pl-cell input');
  if (!input) return;
  const text  = (e.clipboardData || window.clipboardData).getData('text') || '';
  const parts = text.replace(/\r?\n$/, '').split('\t').map(t => t.replace(/\r?\n/g, ' ').trim());
  if (parts.length < 2) return;                    // a normal single-value paste
  e.preventDefault();
  const inputs = cellEls().map(c => c.querySelector('input'));
  inputs.slice(inputs.indexOf(input)).forEach((el, i) => {
    if (parts[i] === undefined) return;
    el.value = parts[i];
    el.dispatchEvent(new Event('input', { bubbles: true }));
  });
});

// ─────────────────────────────────────────────────────────────
// A PL already in SharePoint
// Once a container number and a file number are typed, the container's
// SharePoint folder is checked for a PL. One there is pointed out before
// a new one replaces it, and can be loaded back in to be changed: it reads
// like any other file, its loads table taken from the PL sheet.
// ─────────────────────────────────────────────────────────────
const existingBox = document.getElementById('plExisting');
let existingTimer = null;
let existingAsk   = 0;        // the latest lookup; older answers are dropped
let spSetUp       = null;     // null until the server has said
let existingPL    = null;     // the PL already in this shipment's folder

function shipmentNow() {
  return { container: containerNoInput.value.trim(), file_no: readFileNo(), customer: clientNameInput.value.trim() };
}

function hideExisting() { existingBox.hidden = true; existingBox.innerHTML = ''; }

/* The container's folder, opened from here. Until a lookup has found
   one, the button offers to make it (fba-app makes it where it would
   for the box: Containers Info / client / file-container). */
const spFolderBtn  = document.getElementById('btnSpFolder');
const spFolderHint = document.getElementById('plSpHint');
let knownFolder    = null;    // { url, name } of this shipment's folder

function paintFolderBtn(show) {
  spFolderBtn.hidden = !show;
  spFolderBtn.textContent = knownFolder ? '📁 Open SharePoint folder' : '📁 Create SharePoint folder';
  spFolderBtn.title = knownFolder
    ? `Open ${knownFolder.name} in SharePoint`
    : 'Make this container\u2019s folder in SharePoint (Containers Info / client / file-container) and open it';
  if (!show) { spFolderHint.textContent = ''; spFolderHint.classList.remove('bad'); }
}

spFolderBtn.addEventListener('click', async () => {
  // Opened now, while the click still counts, or the browser blocks it.
  const tab = window.open('', '_blank');
  if (knownFolder) { if (tab) tab.location = knownFolder.url; return; }
  spFolderBtn.disabled = true;
  spFolderHint.classList.remove('bad');
  spFolderHint.textContent = 'Making the folder…';
  try {
    const res = await fetch(`/api/sharepoint-pl/folder?${new URLSearchParams(shipmentNow())}`, { method: 'POST' });
    const out = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(out.error || `server answered ${res.status}`);
    knownFolder = { url: out.url, name: out.name };
    if (tab) tab.location = out.url;
    spFolderHint.textContent = `${out.created ? 'Made' : 'Found'} ${out.name}`;
    paintFolderBtn(true);
  } catch (err) {
    if (tab) tab.close();
    spFolderHint.textContent = `Could not open the folder: ${err.message}`;
    spFolderHint.classList.add('bad');
  } finally {
    spFolderBtn.disabled = false;
  }
});

async function checkExistingPL() {
  const ask = ++existingAsk;
  const sh  = shipmentNow();
  knownFolder = null;
  existingPL  = null;
  if (!/[A-Z]{4}\d{7}/i.test(sh.container) || !sh.file_no) { hideExisting(); paintFolderBtn(false); return; }
  try {
    if (spSetUp === null) spSetUp = !!(await (await fetch('/api/sharepoint-pl')).json()).enabled;
    if (!spSetUp) return;
    const res = await fetch(`/api/sharepoint-pl/find?${new URLSearchParams(sh)}`);
    const out = await res.json();
    if (ask !== existingAsk) return;
    knownFolder = res.ok && out.folder ? { url: out.folder, name: out.folder_name } : null;
    paintFolderBtn(true);
    if (!res.ok || !out.files || !out.files.length) { hideExisting(); return; }
    existingPL = out.files[0];
    showExisting(sh, out, out.files[0]);
  } catch (err) {
    console.error(err);       // the check is a courtesy; generating still works
  }
}

function showExisting(sh, out, f) {
  const when = f.at ? new Date(f.at).toLocaleString() : '';
  existingBox.innerHTML = `
    <span class="pl-existing-text">⚠ A PL for this container is already in SharePoint:
      <strong>${escapeHtml(f.name)}</strong> in ${escapeHtml(out.folder_name || 'its folder')}${
      when ? `, saved ${escapeHtml(when)}` : ''}${f.by ? ` by ${escapeHtml(f.by)}` : ''}.
      Generating a new one will replace it.</span>
    <span class="pl-existing-actions">
      <button type="button" class="pl-existing-load">Load it to edit</button>
      ${f.url ? `<a href="${escapeHtml(f.url)}" target="_blank" rel="noopener">Open in SharePoint</a>` : ''}
    </span>`;
  existingBox.hidden = false;
  const btn = existingBox.querySelector('.pl-existing-load');
  btn.addEventListener('click', async () => {
    btn.disabled = true;
    btn.textContent = 'Loading…';
    try {
      const res = await fetch(`/api/sharepoint-pl/content?${new URLSearchParams({ ...sh, item_id: f.item_id })}`);
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || `server answered ${res.status}`);
      const file = new File([await res.arrayBuffer()], f.name, { type: res.headers.get('content-type') || '' });
      setExcelFile(file, true);
      showStatus(`Loaded ${f.name} from SharePoint. Change what you need, then generate to replace it.`, 'success');
      btn.textContent = 'Loaded';
    } catch (err) {
      console.error(err);
      showStatus(`Could not load ${f.name} from SharePoint: ${err.message}`, 'error');
      btn.disabled = false;
      btn.textContent = 'Load it to edit';
    }
  });
}

[containerNoInput, fileNoInput, clientNameInput].forEach(input => {
  input.addEventListener('input', () => {
    clearTimeout(existingTimer);
    existingTimer = setTimeout(checkExistingPL, 600);
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
    clearCellEdits();
    showStatus('Cleared the hand edits. Ctrl+Z brings them back.', 'info');
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

function renderPreview(modH, modR, sumH, sumR, longDests, ccnSections = null) {
  const { headers: visH, rows: visRows, cols: visCols } = getVisibleColumns(modH, modR);
  lastLoadsHeaders = modH;

  // The workbook writes nine summary columns; show those, not the
  // pipeline's internal names, so the preview is the document.
  const sumIdx  = ['Destination', 'Carton', 'pallet_round_count'].map(h => sumH.indexOf(h));
  const sumVisH = PL_SUMMARY_HEADERS;
  const toSumVis = rows => rows.map(r => {
    const line = sumIdx.map(i => (i >= 0 ? r[i] : null));
    while (line.length < PL_SUMMARY_HEADERS.length) line.push(null);
    return line;
  });

  // Rows where the Destination changes on the next row end a "part".
  const di = visH.indexOf('Destination');
  const splitsOf = rows => {
    const out = new Set();
    if (di >= 0) {
      rows.forEach((r, i) => {
        if (i < rows.length - 1 && rows[i + 1][di] !== r[di]) out.add(i);
      });
    }
    return out;
  };

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

  const idIdx  = modH.indexOf('_Row');
  const loadsTable = rows => {
    const vis = rows.map(r => visCols.map(i => r[i]));
    return tableHtml(visH, vis, {
      splitAfter: splitsOf(vis), editKey: 'loads',
      rowIds: idIdx >= 0 ? rows.map(r => r[idIdx]) : null,
      tints:  di >= 0 ? vis.map(r => tintFor(r[di])) : null,
    });
  };
  // The summary is read-only: every figure in it is derived from the loads
  // table, so it is corrected by editing the load rows, not the totals.
  const summaryTable = rows => {
    const vis = toSumVis(rows);
    return tableHtml(sumVisH, vis, { tints: vis.map(r => tintFor(r[0])) });
  };

  const dataBox = document.getElementById('plPreviewData');
  const sumBox  = document.getElementById('plPreviewSummary');
  if (ccnSections) {
    // Bonded: each CCN# is its own loads table, as in the file.
    dataBox.innerHTML = ccnSections.map((sec, i) =>
      `<div class="pl-ccn-band${i ? ' pl-ccn-gap' : ''}">Loads — CCN# ${escapeHtml(sec.ccn || 'not set')}</div>`
      + loadsTable(sec.rows)).join('');
    // Each CCN# table carries its own "LOADS — CCN# …" title, as in the
    // file, so the band keeps only its Undo / Redo / Reset.
    document.getElementById('plLoadsLabel').textContent = '';
  } else {
    dataBox.innerHTML = loadsTable(modR);
    document.getElementById('plLoadsLabel').textContent = 'Loads';
  }
  // One summary for the whole container. Bonded, it is totalled per CCN#
  // with the CCN# first, as in the file; the CCN# shows once per group,
  // standing in for the merged cell.
  if (ccnSections) {
    const head = ['CCN', ...sumVisH];
    const rows = [], tints = [], ends = new Set();
    for (const sec of ccnSections) {
      const lines = toSumVis(getSummaryDF(modH, sec.rows).rows);
      lines.forEach((r, i) => {
        rows.push([i ? null : (sec.ccn || 'not set'), ...r]);
        tints.push(tintFor(r[0]));
      });
      if (lines.length) ends.add(rows.length - 1);
    }
    sumBox.innerHTML = tableHtml(head, rows, { splitAfter: ends, tints });
  } else {
    sumBox.innerHTML = summaryTable(sumR);
  }
  sumBox.classList.toggle('pl-sum-by-ccn', !!ccnSections);
  renderCcnList(ccnSections, modH);
  renderSheets();
  renderEditBar();
  syncBandTools();
  renderDestPanel(longDests);
  previewSec.style.display = 'block';
  // Laid out now, so the CCN# tables can be measured and lined up.
  if (ccnSections) alignColumns(dataBox.querySelectorAll(':scope > table'));
  // Then sized as the file: its column widths and row heights.
  applyOutputLayout(modH, modR, sumH, sumR, ccnSections);
  // The tables were just rebuilt, so a block dragged on the old ones is gone.
  hideQuickSum('Preview');
  hideCcnAssign();

  // Any change to the preview means the last download, if any, is out of date.
  setSteps({ preview: ['active', `${visRows.length} load line${visRows.length === 1 ? '' : 's'}`],
             download: 'todo' });
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
    const { modH, modR, sumH, sumR, longDests, ccnSections } = await runPipeline(file);
    renderPreview(modH, modR, sumH, sumR, longDests, ccnSections);
    showStatus('✓ Preview generated.', 'success');
  } catch (err) {
    console.error(err);
    showStatus('Error: ' + err.message, 'error');
  }
});

// ─────────────────────────────────────────────────────────────
// Generate
// ─────────────────────────────────────────────────────────────
// Builds the packing list from what is on screen. Generate downloads it and
// also saves it to SharePoint; Upload to SharePoint only saves it there.
async function buildPackingList({ download }) {
  const file = fileInput.files[0];
  if (!file) { showStatus('Please select an Excel file.', 'error'); return; }

  const { cName, containerName } = readClientContainer();
  if (!cName || !containerName) {
    setSteps({ download: ['warn', 'Enter client name and container #'] });
    showStatus('Enter both the client name and the container number.', 'error');
    (cName ? containerNoInput : clientNameInput).focus();
    return;
  }
  const outName = downloadName(containerName);
  if (!download && existingPL
      && !confirm(`${existingPL.name} is already in this container's SharePoint folder. Replace it?`)) return;

  const btns = [document.getElementById('submitBtn'), uploadSpBtn];
  btns.forEach(b => { b.disabled = true; });

  try {
    setSteps({ download: ['active', 'Building the file…'] });
    showStatus('Reading file…', 'info', true);
    const { wbIn, modH, modR, sumH, sumR, ccnSections } = await runPipeline(file);

    // The client's job number is required on every PL: a column of the
    // sheet picked as Job No, or one added in the preview named JOB NO.
    if (!modH.includes('JOB NO')) {
      setSteps({ download: ['warn', 'Pick the Job No column'] });
      showStatus('The Job No column (the client\'s job number) is required. Pick its column for <strong>Job No</strong> in '
        + '<strong>Column mapping</strong> above the grid — "Remember this header" then finds it on its own. '
        + 'If the sheet has none, add a column named <strong>JOB NO</strong> in the preview with + Column and type it in.', 'error');
      return;
    }

    showStatus('Building Excel file…', 'info', true);
    const out = await writePackingList(modH, modR, sumH, sumR, cName, containerName, wbIn,
                                       outName, readFileNo(),
                                       { sections: ccnSections, extras: extraSheets, download });

    setSteps({ preview: 'done', download: ['done', outName] });
    if (download) showStatus('✓ Packing list downloaded successfully.', 'success');

    // The file exists, so anything the user asked the app to learn is now
    // worth keeping. Nothing was written to the shared lists before this.
    let learnNote = '', learnFailed = false;
    try {
      const saved = await flushLearnQueue();
      if (saved && saved.length) {
        learnNote = ` Saved ${saved.join(', ')} to the shared lists — remove under Settings if that was wrong.`;
      }
    } catch (err) {
      console.error(err);
      learnNote = ' The shared lists could not be updated: ' + err.message;
      learnFailed = true;
    }
    await saveToSharePoint(out, { container: containerName, file_no: readFileNo(), customer: cName },
                           learnNote, learnFailed, download);

    // The mapping produced a real file, so now it is worth keeping.
    const data  = await loadSheetData(file);
    renderLearnBar(classifyColumns(data, getRange(data)));
  } catch (err) {
    console.error(err);
    setSteps({ download: ['error', 'Could not build the file'] });
    showStatus('Error: ' + err.message, 'error');
  } finally {
    btns.forEach(b => { b.disabled = false; });
  }
}

const uploadSpBtn = document.getElementById('btnUploadSp');
document.getElementById('plForm').addEventListener('submit', e => {
  e.preventDefault();
  buildPackingList({ download: true });
});
uploadSpBtn.addEventListener('click', () => buildPackingList({ download: false }));

// ─────────────────────────────────────────────────────────────
// SharePoint
// Each downloaded PL also goes to the container's SharePoint folder, by
// way of fba-app, so its Edit view can import it without the file being
// dragged across. A PL made again replaces the one there. The download
// has already happened, so a failure here only says so.
// ─────────────────────────────────────────────────────────────
// `note` is what the shared lists said, kept on the end of the message.
// `downloaded` says whether the file was also saved on this computer.
async function saveToSharePoint(out, shipment, note = '', noteFailed = false, downloaded = true) {
  if (!out || !out.buffer) return;
  const ok = noteFailed ? 'error' : 'success';
  const lead = downloaded ? '✓ Packing list downloaded' : '✓ Packing list';
  const q = new URLSearchParams({ ...shipment, name: out.name });
  try {
    showStatus(`${downloaded ? '✓ Packing list downloaded. ' : ''}Saving it to SharePoint…`, 'info', true);
    const res  = await fetch(`/api/sharepoint-pl?${q}`, { method: 'POST', body: out.buffer,
      headers: { 'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' } });
    const body = await res.json().catch(() => ({}));
    if (res.status === 503 && /not set up/.test(body.error || '')) {
      showStatus(downloaded ? '✓ Packing list downloaded. (SharePoint saving is not set up on this server.)' + note
        : 'SharePoint saving is not set up on this server.', downloaded ? ok : 'error');
      return;
    }
    if (!res.ok) throw new Error(body.error || `server answered ${res.status}`);
    showStatus(`${lead}${downloaded ? ' and' : ''} saved to SharePoint: ${body.folder_name || 'container folder'}`
      + `${body.created ? ' (folder made)' : ''}.` + note, ok);
    setSteps({ download: ['done', `${out.name} · in SharePoint`] });
    checkExistingPL();
  } catch (err) {
    console.error(err);
    showStatus(`${downloaded ? '✓ Packing list downloaded, but it' : 'The packing list'} could not be saved to SharePoint: ${err.message}.` + note, 'error');
  }
}

// ─────────────────────────────────────────────────────────────
// Assign a CCN# to selected lines (bonded PL)
// Drag down the loads, or click one line and Shift-click another; a bar
// offers to give every selected line one CCN#. The value goes in as a hand
// edit of each line's CCN cell, so Undo takes it back, and the lines move
// to that CCN#'s table on the next render.
// ─────────────────────────────────────────────────────────────
const ccnAssignBar   = document.getElementById('plCcnAssign');
const ccnAssignInput = document.getElementById('plCcnAssignInput');
const ccnAssignCount = document.getElementById('plCcnAssignCount');
const ccnAssignList  = document.getElementById('plCcnList');
const ccnAssignPicks = document.getElementById('plCcnPicks');

// Every CCN# in the order it was added — on lines or not yet (see ccnOrder).
// One typed in with "+ Add CCN#" stays until removed, so a whole
// container's CCN#s can be entered first and the lines handed out after.
function knownCcns() {
  return [...ccnOrder];
}
let ccnAssignTarget  = null;   // { col, ids } of the lines the bar will set

// The selected load lines, by row id, and the name of their CCN column.
function selectedCcnTarget() {
  const ids = new Set();
  let col = null;
  for (const line of rangeCells()) {
    const tr = line[0] && line[0].parentElement;
    if (!tr) continue;
    const td = [...tr.children].find(t => t.dataset.table === 'loads' && isCcnHeader(t.dataset.col));
    if (!td) continue;
    col = td.dataset.col;
    ids.add(td.dataset.row);
  }
  return col && ids.size ? { col, ids: [...ids] } : null;
}

function showCcnAssign() {
  ccnAssignTarget = bondedToggle.checked ? selectedCcnTarget() : null;
  if (!ccnAssignTarget) { hideCcnAssign(); return; }
  const n = ccnAssignTarget.ids.length;
  ccnAssignCount.textContent = `${n} line${n === 1 ? '' : 's'}`;
  // Every CCN# on the PL or added by hand is one click away, so a batch of
  // lines is given one without retyping it.
  const known = knownCcns();
  ccnAssignList.innerHTML = known.map(v => `<option value="${escapeHtml(v)}"></option>`).join('');
  ccnAssignPicks.innerHTML = known
    .map(v => `<button type="button" class="pl-ccn-pick" data-ccn="${escapeHtml(v)}">${escapeHtml(v)}</button>`).join('');
  ccnAssignPicks.querySelectorAll('.pl-ccn-pick')
    .forEach(btn => btn.addEventListener('click', () => applyCcnAssign(btn.dataset.ccn)));
  ccnAssignBar.hidden = false;
}

function hideCcnAssign() {
  if (!ccnAssignBar) return;
  ccnAssignBar.hidden = true;
  ccnAssignTarget = null;
}

function applyCcnAssign(picked) {
  if (!ccnAssignTarget) return;
  const value = (typeof picked === 'string' ? picked : ccnAssignInput.value).trim().toUpperCase();
  const { col, ids } = ccnAssignTarget;
  pushEditHistory();                       // the whole batch undoes at once
  for (const id of ids) previewEdits.loads.set(`${id}\u0000${col}`, value || null);
  ccnAssignInput.value = '';
  dropRange();
  showStatus(value
    ? `CCN# ${escapeHtml(value)} set on ${ids.length} line${ids.length === 1 ? '' : 's'}. Ctrl+Z / ⌘Z undoes it.`
    : `CCN# cleared on ${ids.length} line${ids.length === 1 ? '' : 's'}.`, 'success');
  refreshPreview().catch(err => { console.error(err); showStatus('Error: ' + err.message, 'error'); });
}

document.getElementById('plCcnAssignBtn').addEventListener('click', () => applyCcnAssign());
ccnAssignInput.addEventListener('keydown', e => {
  if (e.key === 'Enter')  { e.preventDefault(); applyCcnAssign(); }
  if (e.key === 'Escape') { e.preventDefault(); dropRange(); }
});
ccnAssignBar.querySelector('.pl-ccn-assign-close').addEventListener('click', () => dropRange());

// ─────────────────────────────────────────────────────────────
// Every CCN# on the bonded PL, listed under the Bonded switch
// with its lines and cartons. A CCN# can be renamed here in one go — the
// new value is written as a hand edit on each of its lines, so Undo takes
// it back — and "not set" shows what is still waiting for one.
// ─────────────────────────────────────────────────────────────
let lastCcnSections = null;
let lastCcnHeaders  = null;

function renderCcnList(sections = lastCcnSections, headers = lastCcnHeaders) {
  lastCcnSections = sections;
  lastCcnHeaders  = headers;
  const box = document.getElementById('ccnList');
  if (!box) return;
  if (!sections || !sections.length) { box.innerHTML = ''; return; }
  const ci    = headers.indexOf('Carton');
  const idIdx = headers.indexOf('_Row');
  const col   = headers.find(isCcnHeader);
  // In the order the CCN#s were added; one with no lines yet can be removed.
  const item = (sec, i) => {
    const lines = sec.rows.length;
    const ctns  = ci >= 0 ? sec.rows.reduce((n, r) => n + (parseFloat(r[ci]) || 0), 0) : null;
    const name  = sec.ccn ? escapeHtml(sec.ccn) : '<em>not set</em>';
    return `<div class="pl-ccn-item${sec.ccn ? '' : ' pl-ccn-unset'}" data-i="${i}">
        <span class="pl-ccn-name">${name}</span>
        <span class="pl-ccn-meta">${lines} line${lines === 1 ? '' : 's'}${ctns != null ? ` · ${tidyNumber(ctns)} ctns` : ''}</span>
        <button type="button" class="pl-ccn-rename">${sec.ccn ? 'Rename' : 'Set'}</button>
      </div>`;
  };
  const empty = v => `<div class="pl-ccn-item pl-ccn-added">
        <span class="pl-ccn-name">${escapeHtml(v)}</span>
        <span class="pl-ccn-meta">0 lines</span>
        <button type="button" class="pl-ccn-remove" data-ccn="${escapeHtml(v)}" aria-label="Remove ${escapeHtml(v)}">×</button>
      </div>`;
  const unset = sections.findIndex(sec => !sec.ccn);
  box.innerHTML = ccnOrder.map(v => {
    const i = sections.findIndex(sec => sec.ccn === v);
    return i >= 0 ? item(sections[i], i) : empty(v);
  }).join('')
  + (unset >= 0 ? item(sections[unset], unset) : '')
  + `<button type="button" class="pl-ccn-add">+ Add CCN#</button>`;

  box.querySelectorAll('.pl-ccn-remove').forEach(btn => btn.addEventListener('click', () => {
    ccnOrder = ccnOrder.filter(v => v !== btn.dataset.ccn);
    renderCcnList();
  }));

  const addBtn = box.querySelector('.pl-ccn-add');
  addBtn.addEventListener('click', () => {
    const item = document.createElement('div');
    item.className = 'pl-ccn-item';
    item.innerHTML = `<input type="text" class="pl-ccn-rename-input" spellcheck="false" autocomplete="off"
        placeholder="Type CCN#">
      <button type="button" class="pl-ccn-assign-btn pl-ccn-ok">Add</button>
      <button type="button" class="pl-ccn-assign-close" aria-label="Cancel">×</button>`;
    addBtn.replaceWith(item);
    const input = item.querySelector('input');
    input.focus();
    const add = () => {
      const value = input.value.trim().toUpperCase();
      noteCcn(value);
      renderCcnList();
      if (value) {
        showStatus(`CCN# ${escapeHtml(value)} added. Select its lines in the preview (drag, or click then `
          + 'Shift-click) and pick it in the bar that appears.', 'info');
        // Straight on to the next one: a container often has several.
        document.querySelector('#ccnList .pl-ccn-add').click();
      }
    };
    input.addEventListener('keydown', e => {
      if (e.key === 'Enter')  { e.preventDefault(); add(); }
      if (e.key === 'Escape') { e.preventDefault(); renderCcnList(); }
    });
    item.querySelector('.pl-ccn-ok').addEventListener('click', add);
    item.querySelector('.pl-ccn-assign-close').addEventListener('click', () => renderCcnList());
  });

  box.querySelectorAll('.pl-ccn-item[data-i]').forEach(item => {
    const sec = sections[Number(item.dataset.i)];
    item.querySelector('.pl-ccn-rename').addEventListener('click', () => {
      item.innerHTML = `<input type="text" class="pl-ccn-rename-input" spellcheck="false" autocomplete="off"
          placeholder="Type CCN#" value="${escapeHtml(sec.ccn)}">
        <button type="button" class="pl-ccn-assign-btn pl-ccn-ok">OK</button>
        <button type="button" class="pl-ccn-assign-close" aria-label="Cancel">×</button>`;
      const input = item.querySelector('input');
      input.focus();
      input.select();
      const done = () => {
        const value = input.value.trim().toUpperCase();
        if (value === sec.ccn || !col || idIdx < 0) { renderCcnList(sections, headers); return; }
        // A renamed CCN# keeps its place; renamed onto one already listed, it
        // merges into that one's place instead.
        const at = ccnOrder.indexOf(sec.ccn);
        if (at >= 0 && value && !ccnOrder.includes(value)) ccnOrder[at] = value;
        else if (at >= 0 && value) ccnOrder.splice(at, 1);
        pushEditHistory();                   // the whole CCN# renames at once
        for (const r of sec.rows) previewEdits.loads.set(`${r[idIdx]}\u0000${col}`, value || null);
        showStatus(`${sec.ccn ? `CCN# ${escapeHtml(sec.ccn)}` : 'Lines with no CCN#'} → `
          + `${value ? escapeHtml(value) : 'not set'} on ${sec.rows.length} line${sec.rows.length === 1 ? '' : 's'}. `
          + 'Ctrl+Z / ⌘Z undoes it.', 'success');
        refreshPreview().catch(err => { console.error(err); showStatus('Error: ' + err.message, 'error'); });
      };
      input.addEventListener('keydown', e => {
        if (e.key === 'Enter')  { e.preventDefault(); done(); }
        if (e.key === 'Escape') { e.preventDefault(); renderCcnList(sections, headers); }
      });
      item.querySelector('.pl-ccn-ok').addEventListener('click', done);
      item.querySelector('.pl-ccn-assign-close').addEventListener('click', () => renderCcnList(sections, headers));
    });
  });
}

// The CCN# loads tables are one table in the file — one set of columns down
// the sheet — so in the preview each column is given the width of its
// widest cell in any of them, and the tables line up under one another.
function alignColumns(tables) {
  tables = [...tables];
  if (tables.length < 2) return;
  const widths = [];
  for (const t of tables) {
    [...t.rows[0].cells].forEach((th, i) => {
      if (th.classList.contains('pl-fill')) return;
      widths[i] = Math.max(widths[i] || 0, th.getBoundingClientRect().width);
    });
  }
  for (const t of tables) {
    [...t.rows[0].cells].forEach((th, i) => {
      if (th.classList.contains('pl-fill') || !widths[i]) return;
      // border-box, so the measured outer width is the width it keeps
      th.style.boxSizing = 'border-box';
      th.style.width = th.style.minWidth = `${Math.ceil(widths[i])}px`;
    });
  }
}

// ─────────────────────────────────────────────────────────────
// Columns added by hand
// "+ Column" opens a name box in the Loads band; Enter adds an empty column
// to the loads, to be typed or pasted into like any other, and it is written
// into the file. The × on its header takes it out again, with what was typed
// in it. Both are ordinary changes, so Undo takes them back.
// ─────────────────────────────────────────────────────────────
const addColBtn  = document.getElementById('btnAddCol');
const addColName = document.getElementById('plNewColName');
let lastLoadsHeaders = [];

function closeAddCol() {
  addColName.value  = '';
  addColName.hidden = true;
}

function addHandColumn() {
  const name = addColName.value.replace(/\s+/g, ' ').trim();
  if (!name) { closeAddCol(); return; }
  const taken = [...lastLoadsHeaders, ...handCols, '_Pallet', '_Row']
    .some(h => h != null && String(h).toLowerCase() === name.toLowerCase());
  if (taken) {
    showStatus(`There is already a column called "${name}". Pick another name.`, 'error');
    addColName.focus();
    return;
  }
  pushEditHistory();
  handCols = [...handCols, name];
  closeAddCol();
  showStatus(`Added the column "${name}". Click a cell in it to type, or paste from Excel.`, 'info');
  refreshPreview().catch(err => { console.error(err); showStatus('Error: ' + err.message, 'error'); });
}

addColBtn.addEventListener('click', () => {
  if (addColName.hidden) { addColName.hidden = false; addColName.focus(); return; }
  addHandColumn();
});
addColName.addEventListener('keydown', e => {
  if (e.key === 'Enter')  { e.preventDefault(); addHandColumn(); }
  if (e.key === 'Escape') { e.preventDefault(); closeAddCol(); }
});
addColName.addEventListener('blur', () => {
  // Clicking "+ Column" itself adds; leaving the box empty just closes it.
  setTimeout(() => {
    if (document.activeElement !== addColBtn && !addColName.value.trim()) closeAddCol();
  }, 0);
});

document.getElementById('plPreviewData').addEventListener('click', e => {
  const btn = e.target.closest('.pl-col-remove');
  if (!btn) return;
  const name = btn.dataset.col;
  pushEditHistory();
  handCols = handCols.filter(h => h !== name);
  for (const key of [...previewEdits.loads.keys()]) {
    if (key.endsWith(`\u0000${name}`)) previewEdits.loads.delete(key);
  }
  showStatus(`Removed the column "${name}". Ctrl+Z brings it back.`, 'info');
  refreshPreview().catch(err => { console.error(err); showStatus('Error: ' + err.message, 'error'); });
});

// ─────────────────────────────────────────────────────────────
// Sheets
// Tabs along the foot of the preview, as in Excel: "PL" is the packing
// list above, and "+" adds a blank sheet for anything else the file should
// carry. A blank sheet is typed, pasted and copied like the loads; it goes
// into the file after the summary, under its tab's name. Double-click a
// tab to rename it; its × removes it. Undo covers all of it.
// ─────────────────────────────────────────────────────────────
const sheetTabs   = document.getElementById('plSheetTabs');
const sheetPLBox  = document.getElementById('plSheetPL');
const extraBox    = document.getElementById('plExtraSheet');
const EXTRA_ROWS  = 30;
const EXTRA_COLS  = 12;


// What Excel allows a sheet to be called, and not one the file has already.
function sheetNameProblem(name, exceptId = null) {
  if (!name) return 'A sheet needs a name.';
  if (name.length > 31) return 'A sheet name can be 31 characters at most.';
  if (/[\\/?*[\]:]/.test(name)) return 'A sheet name cannot have \\ / ? * [ ] or :';
  const low = name.toLowerCase();
  if (low === 'pl' || low === 'summary') return `"${name}" is the packing list's own sheet.`;
  if (extraSheets.some(x => x.id !== exceptId && x.name.toLowerCase() === low)) return `There is already a sheet called "${name}".`;
  return null;
}

function addExtraSheet() {
  let name = 'Extra info';
  for (let n = 2; sheetNameProblem(name); n++) name = `Extra info ${n}`;
  pushEditHistory();
  extraSheets = [...extraSheets, { id: ++extraSeq, name, cells: new Map() }];
  previewTab = extraSeq;
  renderSheets();
  syncBandTools();
  showStatus(`Added the sheet "${name}". Type or paste into it; it goes into the file after the summary.`, 'info');
}

function removeExtraSheet(id) {
  const sh = extraSheets.find(x => x.id === id);
  if (!sh) return;
  pushEditHistory();
  extraSheets = extraSheets.filter(x => x.id !== id);
  if (previewTab === id) previewTab = 'PL';
  renderSheets();
  syncBandTools();
  showStatus(`Removed the sheet "${sh.name}". Ctrl+Z brings it back.`, 'info');
}

function renameExtraSheet(id, tab) {
  const sh = extraSheets.find(x => x.id === id);
  if (!sh) return;
  const input = document.createElement('input');
  input.className = 'pl-xl-tab-input';
  input.value = sh.name;
  input.maxLength = 31;
  input.setAttribute('aria-label', 'Sheet name');
  tab.replaceChildren(input);
  input.focus();
  input.select();
  let done = false;
  const finish = (save) => {
    if (done) return;
    done = true;
    const name = input.value.trim();
    if (save && name !== sh.name) {
      const problem = sheetNameProblem(name, id);
      if (problem) showStatus(problem, 'error');
      else { pushEditHistory(); sh.name = name; syncBandTools(); }
    }
    renderSheets();
  };
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter')  { e.preventDefault(); finish(true); }
    if (e.key === 'Escape') { e.preventDefault(); finish(false); }
    e.stopPropagation();
  });
  input.addEventListener('blur', () => finish(true));
}

function renderSheetTabs() {
  const tab = (id, name, removable) => `
    <div class="pl-xl-tab${previewTab === id ? ' active' : ''}" role="tab" tabindex="0"
         aria-selected="${previewTab === id}" data-tab="${id}"
         title="${removable ? 'Double-click to rename' : 'The packing list'}">
      <span class="pl-xl-tab-name">${escapeHtml(name)}</span>${removable
        ? `<button type="button" class="pl-xl-tab-x" data-remove="${id}" aria-label="Remove sheet ${escapeHtml(name)}" title="Remove this sheet"></button>`
        : ''}
    </div>`;
  sheetTabs.innerHTML = tab('PL', 'PL', false)
    + extraSheets.map(x => tab(x.id, x.name, true)).join('')
    + '<button type="button" class="pl-xl-add" title="Add a blank sheet for extra information" aria-label="Add a sheet">+</button>';
}

function renderExtraSheet(sh) {
  let rows = EXTRA_ROWS, cols = EXTRA_COLS;
  for (const key of sh.cells.keys()) {
    const [r, c] = key.split('\u0000').map(Number);
    rows = Math.max(rows, r + 6);
    cols = Math.max(cols, c + 3);
  }
  let html = '<table><thead><tr><th class="pl-xl-corner" aria-hidden="true"></th>';
  for (let c = 0; c < cols; c++) html += `<th>${colLetter(c)}</th>`;
  html += '<th class="pl-fill" aria-hidden="true"></th></tr></thead><tbody>';
  for (let r = 0; r < rows; r++) {
    html += `<tr><th class="pl-xl-rownum" scope="row">${r + 1}</th>`;
    for (let c = 0; c < cols; c++) {
      const v = sh.cells.get(`${r}\u0000${c}`);
      const text = v != null ? String(tidyNumber(v)) : '';
      const num = typeof tidyNumber(v) === 'number' ? ' pl-num' : '';
      html += `<td tabindex="0" spellcheck="false" class="pl-edit${num}" data-table="x:${sh.id}"`
        + ` data-row="${r}" data-col="${c}">${escapeHtml(text)}</td>`;
    }
    html += '<td class="pl-fill"></td></tr>';
  }
  extraBox.innerHTML = html + '</tbody></table>';
}

function renderSheets() {
  if (previewTab !== 'PL' && !extraSheets.some(x => x.id === previewTab)) previewTab = 'PL';
  renderSheetTabs();
  const sh = extraSheets.find(x => x.id === previewTab);
  sheetPLBox.hidden = !!sh;
  extraBox.hidden   = !sh;
  if (sh) renderExtraSheet(sh); else extraBox.innerHTML = '';
}

sheetTabs.addEventListener('click', e => {
  const x = e.target.closest('[data-remove]');
  if (x) { e.stopPropagation(); removeExtraSheet(Number(x.dataset.remove)); return; }
  if (e.target.closest('.pl-xl-add')) { addExtraSheet(); return; }
  const tab = e.target.closest('.pl-xl-tab');
  if (!tab || e.target.closest('input')) return;
  const id = tab.dataset.tab === 'PL' ? 'PL' : Number(tab.dataset.tab);
  if (id === previewTab) return;
  dropRange();
  previewTab = id;
  renderSheets();
});
sheetTabs.addEventListener('dblclick', e => {
  const tab = e.target.closest('.pl-xl-tab');
  if (!tab || tab.dataset.tab === 'PL' || e.target.closest('input, button')) return;
  renameExtraSheet(Number(tab.dataset.tab), tab.querySelector('.pl-xl-tab-name'));
});
sheetTabs.addEventListener('keydown', e => {
  const tab = e.target.closest('.pl-xl-tab');
  if (tab && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); tab.click(); }
});

// The upload button shows only where SharePoint saving is set up.
fetch('/api/sharepoint-pl').then(r => r.json()).then(out => {
  spSetUp = !!out.enabled;
  uploadSpBtn.hidden = !spSetUp;
}).catch(() => {});

/* Opened from the FBA app's "Build PL" (?container=&customer=&file_no=):
   the shipment is filled in, and the folder lookup runs as if typed. A
   field already filled is left as it is. */
(function prefillFromLink() {
  const q = new URLSearchParams(location.search);
  const put = (el, v) => {
    v = String(v || '').trim();
    if (!el || !v || el.value.trim()) return;
    el.value = v;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  };
  put(containerNoInput, q.get('container'));
  put(clientNameInput, q.get('customer'));
  put(fileNoInput, q.get('file_no'));
})();

// ─────────────────────────────────────────────────────────────
// Shipment from the address
// Another app can open the builder ready filled in:
//   packing_list_builder.html?container=TRHU4693221&customer=米娅+Mia&file_no=2067-26253
// Each value goes into its box as if typed, so the SharePoint check for an
// existing PL runs on it as well. A box already filled in is left alone.
// ─────────────────────────────────────────────────────────────
(function shipmentFromUrl() {
  const q = new URLSearchParams(location.search);
  const pick = (...keys) => {
    for (const k of keys) { const v = (q.get(k) || '').trim(); if (v) return v; }
    return '';
  };
  const given = [
    [containerNoInput, pick('container', 'container_no', 'containerNo')],
    [clientNameInput,  pick('customer', 'client', 'client_name')],
    [fileNoInput,      pick('file_no', 'fileNo', 'file')],
  ];
  let filled = 0;
  for (const [input, value] of given) {
    if (!value || input.value.trim()) continue;
    input.value = value;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    filled++;
  }
  if (filled) showStatus(`Shipment filled in from the link: ${given.map(g => g[1]).filter(Boolean).map(escapeHtml).join(' · ')}.`, 'info');
})();

// ─────────────────────────────────────────────────────────────
// The preview sized as the file
// The writer works out the PL sheet's column widths and row heights (a
// column fitted to its longest value, up to 50; a load line 30pt, a
// summary line 85pt; a value too long for its column wrapping, its row
// grown to show it). The preview asks it for those numbers — no file is
// made — and draws its tables with them, scaled down by the same factor
// as its Times New Roman 16pt, so it wraps and grows where the file does.
// ─────────────────────────────────────────────────────────────
const XL_SCALE = 0.62;
const colPx = w => Math.round((w * 7 + 5) * XL_SCALE);     // Excel width units → px
const rowPx = pt => Math.round(pt * 96 / 72 * XL_SCALE);   // points → px
let layoutSeq = 0;

async function applyOutputLayout(modH, modR, sumH, sumR, ccnSections) {
  const ask = ++layoutSeq;
  let L;
  try {
    const { cName, containerName } = readClientContainer();
    L = await writePackingList(modH, modR, sumH, sumR, cName, containerName, null, '', readFileNo(),
                               { sections: ccnSections, layoutOnly: true });
  } catch (err) {
    console.error('preview layout:', err);      // the preview still shows, at its own sizes
    return;
  }
  if (ask !== layoutSeq || !L || !L.widths) return;
  const size = (table, hdrPt, rowPts) => {
    if (!table) return;
    table.classList.add('pl-as-output');
    table.querySelectorAll('thead th:not(.pl-fill)').forEach((th, i) => {
      const px = `${colPx(L.widths[i] ?? 12)}px`;
      th.style.width = th.style.minWidth = th.style.maxWidth = px;
    });
    const head = table.querySelector('thead tr');
    if (head) head.style.height = `${rowPx(hdrPt)}px`;
    [...table.querySelectorAll('tbody tr')].forEach((tr, i) => {
      if (rowPts[i]) tr.style.height = `${rowPx(rowPts[i])}px`;
    });
  };
  [...document.querySelectorAll('#plPreviewData > table, #plPreviewData table')]
    .filter((t, i, all) => all.indexOf(t) === i)
    .forEach((t, i) => { const s = L.loads[i]; if (s) size(t, s.hdr, s.rows); });
  size(document.querySelector('#plPreviewSummary table'), L.sumHdr, L.sumRows);
}
