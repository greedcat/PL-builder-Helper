// ── Packing List Builder UI · part 2: undo/reset, spreadsheet-style cell editing, drag-copy ──
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
