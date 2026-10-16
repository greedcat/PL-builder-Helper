// ── Packing List Builder UI · part 4: shipment cells, sheet grid toggle, preview, generate ──
// Shipment cells
// The order is fixed — File number, Client name, Container # — the order
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
  renderEditBar();
  syncBandTools();
  renderDestPanel(longDests);
  previewSec.style.display = 'block';
  // Laid out now, so the CCN# tables can be measured and lined up.
  if (ccnSections) alignColumns(dataBox.querySelectorAll(':scope > table'));
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
document.getElementById('plForm').addEventListener('submit', async e => {
  e.preventDefault();
  const file = fileInput.files[0];
  if (!file) { showStatus('Please select an Excel file.', 'error'); return; }

  const { cName, containerName } = readClientContainer();
  if (!cName || !containerName) {
    setSteps({ download: ['warn', 'Enter client name and container #'] });
    showStatus('Enter both the client name and the container number.', 'error');
    (cName ? containerNoInput : clientNameInput).focus();
    return;
  }

  const btn    = document.getElementById('submitBtn');
  btn.disabled = true;

  try {
    setSteps({ download: ['active', 'Building the file…'] });
    showStatus('Reading file…', 'info', true);
    const { wbIn, modH, modR, sumH, sumR, ccnSections } = await runPipeline(file);

    showStatus('Building Excel file…', 'info', true);
    const outName = downloadName(containerName);
    await writePackingList(modH, modR, sumH, sumR, cName, containerName, wbIn,
                           outName, readFileNo(), { sections: ccnSections });

    setSteps({ preview: 'done', download: ['done', outName] });
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
    setSteps({ download: ['error', 'Could not build the file'] });
    showStatus('Error: ' + err.message, 'error');
  } finally {
    btn.disabled = false;
  }
});

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
