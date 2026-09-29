// ── Packing List Builder UI · part 4: shipment cells, sheet grid toggle, preview, generate ──
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
