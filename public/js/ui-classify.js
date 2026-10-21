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
  // The client's job number: required, see buildPackingList.
  { role: 'JOB NO',      cls: 'job',    label: 'Job No',      color: '#0f766e' },
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
    'JOB NO':      detectJobColumn(headers),
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
    if (bondedToggle.checked && local === ccnLocalIndex(headers, firstCol)) {
      cols.push({ kind: 'ccn', cls: 'ccn', label: 'CCN#', name, hasData });
    } else if (roleByLocal[local] !== undefined) {
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

  // One row per field of the packing list: the sheet column filling it
  // (pick another from the list), how it was found, and whether the
  // header is kept for next time. Destination, FBA ID and REF ID are
  // recognised from their values, so only the others have a saved list.
  const picked = role => Object.prototype.hasOwnProperty.call(roleOverrides, role);
  const mapRows = PL_ROLES.map(r => {
    const abs  = roleCol[r.role];
    const col  = abs != null ? cols[abs] : null;
    const name = col && col.name != null && col.name !== '' ? String(col.name) : null;
    const listKey = ROLE_KEYWORD_LIST[r.role];
    const options = ['<option value="">— none —</option>'].concat(info.choices
      .filter(c => c.name != null && c.name !== '')
      .map(c => `<option value="${c.col}"${c.col === abs ? ' selected' : ''}>${c.letter} — ${escapeHtml(String(c.name))}</option>`))
      .join('');
    let how, howCls = '';
    if (abs == null) { how = picked(r.role) ? 'Set to none' : 'Not found — pick its column'; howCls = ' pl-map-bad'; }
    else if (picked(r.role)) how = 'Picked by you';
    else how = listKey ? 'Found by its header' : 'Found by its values';
    let next = '<span class="pl-map-muted">—</span>';
    if (listKey && name) {
      next = alreadyKnown(listKey, name)
        ? '<span class="pl-map-ok">✓ Header saved</span>'
        : `<button type="button" class="pl-map-remember" data-remember="${escapeHtml(r.role)}"
             title="Save “${escapeHtml(name)}” as a ${escapeHtml(r.label)} header, so the next sheet headed this way is mapped on its own">Remember this header</button>`;
    } else if (!listKey) {
      next = '<span class="pl-map-muted">Found from the values each time</span>';
    }
    return `
      <tr${abs == null ? ' class="pl-map-missing"' : ''}>
        <th scope="row"><i class="pl-dot" style="background:${abs != null ? r.color : '#cbd5e1'}"></i>${escapeHtml(r.label)}</th>
        <td><select class="pl-map-select" data-role="${escapeHtml(r.role)}" aria-label="Sheet column for ${escapeHtml(r.label)}">${options}</select>
          ${picked(r.role) ? `<button type="button" class="pl-map-auto" data-auto="${escapeHtml(r.role)}" title="Let the app find this column again">↺ automatic</button>` : ''}</td>
        <td class="pl-map-how${howCls}">${how}</td>
        <td>${next}</td>
      </tr>`;
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

  return `
    <div class="pl-range-bar">
      <span class="pl-range-ref">${ref}</span>${mode}${picker}${reset}
    </div>
    <p class="pl-range-explain">${explain}</p>
    <p class="pl-range-hint">Drag the four blue corners to resize the box. The top row of the box is always the header.
      Dragging across the grid only highlights cells (Ctrl+C copies them) — right-click a highlight
      to <strong>use it as the data range</strong>, or, when the table is split by a gap,
      to <strong>add it as an extra block</strong>.</p>
    <div class="pl-map">
      <div class="pl-map-title">Column mapping
        <span>Which column of the sheet fills each field of the packing list${asides ? ` · ${asides}` : ''}</span></div>
      <table class="pl-map-table">
        <thead><tr><th>Field</th><th>Sheet column</th><th>How it was found</th><th>Next time</th></tr></thead>
        <tbody>${mapRows}</tbody>
      </table>
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
  if (plMenuEl && !e.target.closest('.pl-menu')) closePlMenu();
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
  // Bonded: the CCN# can be any column, whatever its header says.
  if (bondedToggle.checked) {
    const isCcn = cols[abs] && cols[abs].kind === 'ccn';
    items.push({
      label: 'Use as CCN#',
      checked: isCcn,
      onPick: () => {
        if (isCcn && ccnColOverride === abs) { ccnColOverride = null; reapply(); return; }
        ccnColOverride = abs;
        columnOverrides[abs] = 'keep';
        // A column cannot be both the CCN# and, say, the Destination.
        for (const r of PL_ROLES) if (roleCol[r.role] === abs) roleOverrides[r.role] = null;
        clearPreviewEdits();     // edits were keyed by the old CCN column name
        reapply();
      },
    });
  }
  items.push({ sep: true });
  items.push({
    label: 'Not used as a field',
    checked: !PL_ROLES.some(r => roleCol[r.role] === abs) && !(cols[abs] && cols[abs].kind === 'ccn'),
    onPick: () => {
      for (const r of PL_ROLES) if (roleCol[r.role] === abs) roleOverrides[r.role] = null;
      if (ccnColOverride === abs) ccnColOverride = null;
      reapply();
    },
  });
  items.push({ sep: true });
  items.push({
    label: 'Include in the packing list',
    checked: !isDrop,
    onPick: () => { columnOverrides[abs] = 'keep'; reapply(); },
  });
  items.push({
    label: 'Leave out of this packing list',
    checked: isDrop,
    onPick: () => { columnOverrides[abs] = 'drop'; reapply(); },
  });

  // Dropping the same wording on every file is a rule, not a decision to
  // repeat. This writes it to the shared ignore list so the next sheet
  // that uses the words drops the column without being asked.
  if (name != null && String(name).trim() !== '' && !alreadyKnown('NO_NEED_COL', name)) {
    items.push({
      label: `Always leave out columns named “${escapeHtml(String(name))}”`,
      onPick: async () => {
        columnOverrides[abs] = 'drop';
        try {
          queueLearn('NO_NEED_COL', String(name));
          await flushLearnQueue();
          delete columnOverrides[abs];       // the list drops it now
          await reapply();
          showStatus(`Columns named “${escapeHtml(String(name))}” are now left out of every packing list. `
            + 'Remove it from NO_NEED_COL under Settings to undo.', 'success');
        } catch (err) {
          console.error(err);
          await reapply();
          showStatus('Could not save to the shared ignore list: ' + err.message, 'error');
        }
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

// Re-runs classification and the preview after a menu choice.
function reapply() {
  hideLearnBar();
  return refreshSheetView({ keepSelection: true })
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
const ROLE_KEYWORD_LIST = { Carton: 'CARTON_KEYWORDS', Weight: 'WEIGHT_KEYWORDS', CMB: 'CBM_KEYWORDS',
  'JOB NO': 'JOB_KEYWORDS' };

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

// Saves one header wording to its list now, and maps by it from here on.
async function rememberWording(it) {
  queueLearn(it.listKey, String(it.name));
  try {
    await flushLearnQueue();
    // The wording is known now, so detection finds it without the override.
    if (it.role) delete roleOverrides[it.role];
    await reapply();
    const onlyHere = OWN_FIELDS.includes(it.listKey) && lastSaveWhere === 'browser';
    showStatus(onlyHere
      ? `“${escapeHtml(String(it.name))}” saved in this browser. The shared list needs the PL builder server's database `
        + '(MONGO_URI); until then other computers will not know it.'
      : `“${escapeHtml(String(it.name))}” saved — sheets headed this way are mapped as ${escapeHtml(it.label)} from now on. `
        + 'Remove it under Settings if that was wrong.', onlyHere ? 'info' : 'success');
  } catch (err) {
    console.error(err);
    showStatus('Could not save to the shared lists: ' + err.message, 'error');
    reapply();
  }
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
      await rememberWording(it);
    });
  });
}

// The grid is created once and reused for every file.
let sheetGrid = null;

function ensureGrid() {
  if (sheetGrid) return sheetGrid;
  sheetGrid = createSheetGrid({
    mount: sheetView,
    onCopy: (nRows, nCols) => showStatus(nRows === 1 && nCols === 1 ? 'Copied 1 cell.'
      : `Copied ${nRows} row${nRows === 1 ? '' : 's'} × ${nCols} column${nCols === 1 ? '' : 's'}.`, 'success'),
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
  sheetGrid.onHighlightChange(cells =>
    cells ? showQuickSum({ source: 'Sheet', ...cells }) : hideQuickSum('Sheet'));
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
  if (rebuild) {
    const ws = data.wbIn.Sheets[data.sheetName];
    grid.render(data.raw, sel, data.maxCols, (ws && ws['!merges']) || []);
    grid.setExtras(blocks);
  }
  else if (!keepSelection) { grid.setSelection(sel); grid.setExtras(blocks); }

  const info = classifyColumns(data, range);
  grid.setColumnInfo(info.cols);
  sheetLegend.innerHTML = renderSheetLegend(info, range, data);

  // The mapping table: pick a column, go back to automatic, or keep a header.
  sheetLegend.querySelectorAll('.pl-map-select').forEach(sel => {
    sel.addEventListener('change', () => {
      const role = sel.dataset.role;
      const c = info.choices.find(x => String(x.col) === sel.value);
      roleOverrides[role] = c ? c.name : null;
      if (c) grid.focusColumn(c.col, { toggle: false });
      reapply();
    });
  });
  sheetLegend.querySelectorAll('[data-auto]').forEach(btn => {
    btn.addEventListener('click', () => { delete roleOverrides[btn.dataset.auto]; reapply(); });
  });
  sheetLegend.querySelectorAll('[data-remember]').forEach(btn => {
    btn.addEventListener('click', async () => {
      const role = btn.dataset.remember;
      const abs  = info.roleCol[role];
      const name = abs != null && info.cols[abs] ? info.cols[abs].name : null;
      const meta = PL_ROLES.find(r => r.role === role);
      if (!name || !ROLE_KEYWORD_LIST[role]) return;
      btn.disabled = true;
      btn.textContent = 'Saving…';
      await rememberWording({ role, label: meta ? meta.label : role, listKey: ROLE_KEYWORD_LIST[role], name });
    });
  });

  // Two menus. A column letter opens the column's menu: which role it fills,
  // keep or drop. A cell or row number opens the menu for the highlighted
  // cells: use them as the range, add them as a block, copy them.
  grid.onColumnContextMenu((abs, x, y, where) => {
    const meta  = info.choices.find(c => c.col === abs);
    const name  = meta && meta.name != null && meta.name !== '' ? escapeHtml(String(meta.name)) : '(no header)';
    const colTitle = `Column ${colLetter(abs)} — ${name}`;
    const openColumnMenu = () => {
      grid.focusColumn(abs, { toggle: false });
      openPlMenu(x, y, colTitle, columnMenuItems(abs, info));
    };

    if (where && where.kind === 'col') { openColumnMenu(); return; }

    const items = [];
    const h = where && where.highlight;
    if (h) {
      const ref      = `${cellRef(h.r1, h.c1)}:${cellRef(h.r2, h.c2)}`;
      const rowsText = h.r1 === h.r2 ? `row ${h.r1 + 1}` : `rows ${h.r1 + 1}–${h.r2 + 1}`;
      items.push({
        label: 'Use as data range',
        note:  h.r2 <= h.r1 ? 'needs a header and a data row' : ref,
        disabled: h.r2 <= h.r1,
        onPick: () => grid.useHighlightAsRange(),
      }, {
        label: `Add ${rowsText} as extra block`,
        note:  where.highlightOverlaps ? 'already in the range' : '',
        disabled: where.highlightOverlaps,
        onPick: () => grid.addHighlightAsBlock(),
      }, {
        label: 'Copy',
        note:  'Ctrl+C',
        onPick: () => {
          const text = grid.highlightText();
          navigator.clipboard.writeText(text)
            .then(() => showStatus(`Copied ${ref}.`, 'success'))
            .catch(() => showStatus('Could not copy — press Ctrl+C instead.', 'error'));
        },
      });
    }

    // A block added on top of the main range can be taken back out here.
    if (where && where.extraIndex >= 0) {
      const b = grid.getExtras()[where.extraIndex];
      items.push({ sep: true }, {
        label: `Remove this block (rows ${b.r1 + 1}–${b.r2 + 1})`,
        onPick: () => grid.removeExtra(where.extraIndex),
      });
    }

    // The column's own menu stays one step away.
    items.push({ sep: true }, {
      label: `Column ${colLetter(abs)} options…`,
      note:  name,
      onPick: openColumnMenu,
    });

    const title = !h ? 'Cells'
      : (h.r1 === h.r2 && h.c1 === h.c2) ? `Cell ${cellRef(h.r1, h.c1)}`
      : `Cells ${cellRef(h.r1, h.c1)}:${cellRef(h.r2, h.c2)}`;
    openPlMenu(x, y, title, items);
  });

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
