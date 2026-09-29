// ── Packing List Builder UI · part 3: cell classification, pop-up menu, detector teaching ──
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
