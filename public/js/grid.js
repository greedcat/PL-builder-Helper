// ─────────────────────────────────────────────────────────────
// EXCEL-LIKE SHEET GRID
// Renders the uploaded worksheet as a spreadsheet: column letters across the
// top, real row numbers down the side, and a rectangular selection you drag
// out with the mouse. The grid knows nothing about packing lists; the caller
// passes in per-column tint classes and receives selection changes.
// ─────────────────────────────────────────────────────────────

// Rendering every row of a very large sheet costs more than it is worth.
// Beyond this, the tail is omitted and the caller is told.
const GRID_MAX_ROWS = 3000;
const GRID_MAX_COLS = 200;

function colLetter(c) {
  let s = '';
  for (let n = c; n >= 0; n = Math.floor(n / 26) - 1) {
    s = String.fromCharCode(65 + (n % 26)) + s;
  }
  return s;
}

function cellRef(r, c) {
  return colLetter(c) + (r + 1);
}

function rangeRef(sel) {
  return sel ? `${cellRef(sel.r1, sel.c1)}:${cellRef(sel.r2, sel.c2)}` : '';
}

// Orders a raw anchor/focus pair into a normalized rectangle.
function normalizeSel(a, b) {
  return {
    r1: Math.min(a.r, b.r), r2: Math.max(a.r, b.r),
    c1: Math.min(a.c, b.c), c2: Math.max(a.c, b.c),
  };
}

function createSheetGrid({ mount, onSelect }) {
  let rows     = [];
  let nRows    = 0;
  let nCols    = 0;
  let shown    = 0;
  let sel      = null;
  // Data is not always in one run: a sheet can carry a block, a gap, then
  // more of the same table. Ctrl or Cmd held while dragging adds a block
  // instead of replacing the selection.
  let extras   = [];
  let colInfo  = [];
  let cellEls  = [];   // cellEls[r][c], body cells only
  let rowEls   = [];   // row-number header cells
  let colEls   = [];   // column-letter header cells

  // Drag state. mode: 'cell' | 'row' | 'col'
  let dragging = null;
  let anchor   = null;
  let addingExtra = null;
  // 'replace' starts a fresh selection on every drag; 'add' appends a block.
  // Ctrl or Cmd forces 'add' whatever the mode is.
  let selectMode = 'replace';
  function setSelectMode(m) { selectMode = m === 'add' ? 'add' : 'replace'; }
  function getSelectMode() { return selectMode; }

  let canvas    = null;
  let handleEls = {};

  // Set by the caller; receives (absoluteColumn, pageX, pageY).
  let contextHandler = null;
  function onColumnContextMenu(fn) { contextHandler = fn; }

  function rowsSelected(r) {
    if (sel && r >= sel.r1 && r <= sel.r2) return true;
    return extras.some(b => r >= b.r1 && r <= b.r2);
  }

  // Blocks that touch are one block. Leaving them separate drew a border
  // through the middle of a continuous run and stacked two handles on the
  // same corner, which is what made a rearranged selection unreadable.
  function mergeTouchingBlocks() {
    if (!sel) return false;
    const all = [{ ...sel, main: true }, ...extras.map(b => ({ ...b }))]
      .sort((a, b) => a.r1 - b.r1);

    const merged = [];
    for (const b of all) {
      const last = merged[merged.length - 1];
      if (last && b.r1 <= last.r2 + 1) {
        last.r2  = Math.max(last.r2, b.r2);
        last.main = last.main || b.main;
      } else {
        merged.push({ ...b });
      }
    }
    if (merged.length === all.length) return false;       // nothing touched

    // The block holding the header stays the main selection; it is the one
    // the header row falls in, which after a merge is the topmost.
    const mainIdx = Math.max(merged.findIndex(b => b.main), 0);
    const cols = { c1: sel.c1, c2: sel.c2 };
    sel = { r1: merged[mainIdx].r1, r2: merged[mainIdx].r2, ...cols };
    extras = merged.filter((_, i) => i !== mainIdx).map(b => ({ r1: b.r1, r2: b.r2, ...cols }));
    return true;
  }

  function clampSel(s) {
    if (!s) return null;
    return {
      r1: Math.max(0, Math.min(s.r1, shown - 1)),
      r2: Math.max(0, Math.min(s.r2, shown - 1)),
      c1: Math.max(0, Math.min(s.c1, nCols - 1)),
      c2: Math.max(0, Math.min(s.c2, nCols - 1)),
    };
  }

  // Paints selection classes. Called on every mouse move, so it touches only
  // the cells whose state can have changed rather than re-rendering.
  let paintedEls = [];
  let paintedRowHdrs = [];
  let paintedColHdrs = [];
  let paintQueued = false;

  // A drag fires many mousemoves; repainting a large selection on each one
  // stalls the page. Collapse them into one repaint per frame.
  function schedulePaint() {
    if (paintQueued) return;
    paintQueued = true;
    requestAnimationFrame(() => { paintQueued = false; paint(); });
  }

  function paint() {
    // Reset exactly what was painted last time. Relying on a cached
    // rectangle was fragile: several callers cleared that cache and the
    // repaint then left the previous selection on screen, so a shrunken
    // block still looked like its old self.
    for (const el of paintedEls) el.className = el.dataset.baseCls || '';
    for (const el of paintedRowHdrs) el.classList.remove('pl-grid-rownum-sel');
    for (const el of paintedColHdrs) el.classList.remove('pl-grid-colhdr-sel');
    paintedEls = [];
    paintedRowHdrs = [];
    paintedColHdrs = [];

    const paintBlock = (b, isMain) => {
      for (let i = b.r1; i <= b.r2; i++) {
        for (let j = b.c1; j <= b.c2; j++) {
          const el = cellEls[i] && cellEls[i][j];
          if (!el) continue;
          const edges = [
            i === b.r1 ? 'pl-sel-t' : '', i === b.r2 ? 'pl-sel-b' : '',
            j === b.c1 ? 'pl-sel-l' : '', j === b.c2 ? 'pl-sel-r' : '',
          ].filter(Boolean).join(' ');
          // Only the main block's first row is the header row.
          const band = (isMain && i === b.r1) ? 'pl-sel-hdr' : 'pl-sel-data';
          el.className = `${el.dataset.baseCls || ''} pl-sel ${band} ${edges}`.trim();
          paintedEls.push(el);
        }
        if (rowEls[i]) {
          rowEls[i].classList.add('pl-grid-rownum-sel');
          paintedRowHdrs.push(rowEls[i]);
        }
      }
      for (let j = b.c1; j <= b.c2; j++) {
        if (colEls[j]) {
          colEls[j].classList.add('pl-grid-colhdr-sel');
          paintedColHdrs.push(colEls[j]);
        }
      }
    };

    if (sel) paintBlock(sel, true);
    for (const b of extras) paintBlock(b, false);

    applyFocus(focusCol);
    placeHandles();
  }

  // Corners are positioned from the cells themselves, so they stay correct
  // whatever the column widths are.
  // Blocks share the selection's columns, so two blocks overlap only when
  // their rows do. A block may never eat into another.
  function occupiedRows(exclude) {
    const busy = new Set();
    if (exclude !== 'main' && sel) {
      for (let r = sel.r1; r <= sel.r2; r++) busy.add(r);
    }
    extras.forEach((b, i) => {
      if (i === exclude) return;
      for (let r = b.r1; r <= b.r2; r++) busy.add(r);
    });
    return busy;
  }

  // Grows out from the row the drag started on, stopping before anything
  // already taken. Dragging across a neighbour therefore stalls at its edge
  // instead of swallowing it.
  function clampSpan(r1, r2, anchorRow, exclude) {
    const busy = occupiedRows(exclude);
    if (busy.has(anchorRow)) return null;
    let lo = anchorRow, hi = anchorRow;
    while (lo - 1 >= r1 && !busy.has(lo - 1)) lo--;
    while (hi + 1 <= r2 && !busy.has(hi + 1)) hi++;
    return { r1: lo, r2: hi };
  }

  const CORNERS = ['tl', 'tr', 'bl', 'br'];

  // Corners are positioned from the cells themselves, so they stay correct
  // whatever the column widths are. One set per block.
  function placeHandles() {
    if (!canvas) return;
    const blocks = [];
    if (sel) blocks.push({ key: 'main', box: sel });
    extras.forEach((b, i) => blocks.push({ key: String(i), box: b }));

    const base = canvas.getBoundingClientRect();
    const wanted = new Set();

    for (const { key, box } of blocks) {
      for (const corner of CORNERS) {
        const id = `${key}:${corner}`;
        wanted.add(id);
        let el = handleEls[id];
        if (!el) {
          el = document.createElement('div');
          el.className = `pl-handle pl-handle-${corner}`;
          el.dataset.block = key;
          el.dataset.h = corner;
          canvas.appendChild(el);
          handleEls[id] = el;
        }
        const r = (corner === 'tl' || corner === 'tr') ? box.r1 : box.r2;
        const c = (corner === 'tl' || corner === 'bl') ? box.c1 : box.c2;
        const cell = cellEls[r] && cellEls[r][c];
        if (!cell) { el.hidden = true; continue; }
        const rect = cell.getBoundingClientRect();
        el.style.left = ((corner === 'tl' || corner === 'bl')
          ? rect.left - base.left : rect.right - base.left) + 'px';
        el.style.top = ((corner === 'tl' || corner === 'tr')
          ? rect.top - base.top : rect.bottom - base.top) + 'px';
        el.hidden = false;
      }
    }

    for (const id of Object.keys(handleEls)) {
      if (wanted.has(id)) continue;
      handleEls[id].remove();
      delete handleEls[id];
    }
  }

  // A column the caller asked to point at. Separate from `sel`, so pointing
  // at a column never alters what the pipeline reads.
  let focusCol = null;

  function applyFocus(prev) {
    const mark = (c, on) => {
      if (c == null) return;
      if (colEls[c]) colEls[c].classList.toggle('pl-col-focus-hdr', on);
      for (let r = 0; r < shown; r++) {
        const el = cellEls[r] && cellEls[r][c];
        if (el) el.classList.toggle('pl-col-focus', on);
      }
    };
    if (prev != null && prev !== focusCol) mark(prev, false);
    mark(focusCol, true);
  }

  // Returns the column now focused, or null when the same one was toggled off.
  function focusColumn(c, { toggle = true } = {}) {
    const prev = focusCol;
    focusCol = (toggle && focusCol === c) ? null : c;
    applyFocus(prev);
    if (focusCol != null && colEls[focusCol]) {
      colEls[focusCol].scrollIntoView({ block: 'nearest', inline: 'center' });
    }
    return focusCol;
  }

  function clearFocus() {
    if (focusCol == null) return;
    const prev = focusCol;
    focusCol = null;
    applyFocus(prev);
  }

  function setSelection(s, { fire = false } = {}) {
    sel = clampSel(s);
    paint();
    if (fire && onSelect) onSelect(sel);
  }

  function getSelection() { return sel; }
  function getExtras() { return extras.map(b => ({ ...b })); }

  // Which extra block covers this row, or -1 for the main selection.
  function extraAt(row) {
    return extras.findIndex(b => row >= b.r1 && row <= b.r2);
  }

  function removeExtra(i) {
    if (i < 0 || i >= extras.length) return null;
    const gone = extras.splice(i, 1)[0];
    paint();
    if (onSelect) onSelect(sel, extras.map(b => ({ ...b })));
    return gone;
  }
  function setExtras(list) {
    extras = (list || []).map(b => ({ ...b }));
    mergeTouchingBlocks();
    paint();
  }

  // Applies per-column tint classes without rebuilding the table.
  function setColumnInfo(info) {
    colInfo = info || [];
    for (let c = 0; c < nCols; c++) {
      const meta = colInfo[c] || {};
      const cls  = meta.cls ? `pl-k-${meta.cls}` : '';
      if (colEls[c]) {
        colEls[c].className = `pl-grid-colhdr ${cls}`.trim();
        const tag = colEls[c].querySelector('.pl-th-role');
        if (tag) tag.textContent = meta.label || '';
      }
      for (let r = 0; r < shown; r++) {
        const el = cellEls[r] && cellEls[r][c];
        if (!el) continue;
        // Only cells inside the selected block carry a role tint.
        const inSel = sel && r >= sel.r1 && r <= sel.r2 && c >= sel.c1 && c <= sel.c2;
        el.dataset.baseCls = inSel ? cls : '';
      }
    }
    paint();
  }

  // `colCount` is the caller's measure of how many columns actually hold
  // data. Excel's used range is not a safe substitute.
  function render(rawRows, initialSel, colCount) {
    rows  = rawRows || [];
    nRows = rows.length;
    const wide = colCount || rows.reduce((m, r) => Math.max(m, r ? r.length : 0), 1);
    nCols = Math.min(wide, GRID_MAX_COLS);
    shown = Math.min(nRows, GRID_MAX_ROWS);

    let html = '<table class="pl-grid"><thead><tr><th class="pl-grid-corner"></th>';
    for (let c = 0; c < nCols; c++) {
      html += `<th class="pl-grid-colhdr" data-c="${c}">${colLetter(c)}<span class="pl-th-role"></span></th>`;
    }
    html += '</tr></thead><tbody>';
    for (let r = 0; r < shown; r++) {
      const row = rows[r] || [];
      html += `<tr><th class="pl-grid-rownum" data-r="${r}">${r + 1}</th>`;
      for (let c = 0; c < nCols; c++) {
        const v = row[c];
        html += `<td data-r="${r}" data-c="${c}">${v != null ? escapeHtml(String(v)) : ''}</td>`;
      }
      html += '</tr>';
    }
    html += '</tbody></table>';
    // Corner handles live beside the table inside a positioned wrapper, so
    // they scroll with the content rather than floating over the viewport.
    html = `<div class="pl-grid-canvas">${html}</div>`;
    const notes = [];
    if (nRows > shown) notes.push(`the first ${shown} rows of ${nRows}`);
    if (wide > nCols) notes.push(`the first ${nCols} columns of ${wide}`);
    if (notes.length) {
      html += `<div class="pl-grid-trunc">This sheet is very large, so the grid shows
               ${notes.join(' and ')}. Use the row boxes above to reach the rest.</div>`;
    }
    mount.innerHTML = html;

    // Cache element references for fast selection painting.
    const table  = mount.querySelector('table');
    canvas       = mount.querySelector('.pl-grid-canvas');
    handleEls    = {};
    for (const h of canvas.querySelectorAll('.pl-handle')) h.remove();
    colEls = [...table.querySelectorAll('thead th[data-c]')];
    rowEls = [];
    cellEls = [];
    for (const tr of table.querySelectorAll('tbody tr')) {
      const rh = tr.querySelector('th[data-r]');
      rowEls[Number(rh.dataset.r)] = rh;
      const line = [];
      for (const td of tr.querySelectorAll('td[data-c]')) line[Number(td.dataset.c)] = td;
      cellEls[Number(rh.dataset.r)] = line;
    }

    // A fresh table means fresh elements; nothing from before is on screen.
    paintedEls = [];
    paintedRowHdrs = [];
    paintedColHdrs = [];
    handleEls = {};
    extras   = [];
    focusCol = null;
    sel      = clampSel(initialSel);
    paint();
    attach(table);
    attachHandles();
  }

  function pointFrom(target) {
    const td = target.closest && target.closest('td[data-r]');
    if (td) return { r: Number(td.dataset.r), c: Number(td.dataset.c), kind: 'cell' };
    const rh = target.closest && target.closest('th[data-r]');
    if (rh) return { r: Number(rh.dataset.r), c: 0, kind: 'row' };
    const ch = target.closest && target.closest('th[data-c]');
    if (ch) return { r: 0, c: Number(ch.dataset.c), kind: 'col' };
    return null;
  }

  function extendTo(p) {
    if (!anchor) return;

    if (dragging === 'extra' && addingExtra) {
      const span = clampSpan(Math.min(anchor.r, p.r), Math.max(anchor.r, p.r),
                             anchor.r, extras.indexOf(addingExtra));
      if (span) { addingExtra.r1 = span.r1; addingExtra.r2 = span.r2; }
      schedulePaint();
      return;
    }

    if (dragging === 'resize' && resizing) {
      const exclude = resizing.block;
      const box = exclude === 'main' ? sel : extras[exclude];
      if (!box) return;
      const span = clampSpan(Math.min(anchor.r, p.r), Math.max(anchor.r, p.r),
                             anchor.r, exclude);
      if (span) { box.r1 = span.r1; box.r2 = span.r2; }
      box.c1 = Math.min(anchor.c, p.c);
      box.c2 = Math.max(anchor.c, p.c);
      // Columns belong to the whole selection, not to one block.
      if (exclude === 'main') extras.forEach(b => { b.c1 = box.c1; b.c2 = box.c2; });
      else if (sel) { box.c1 = sel.c1; box.c2 = sel.c2; }
      schedulePaint();
      return;
    }
    if (dragging === 'row')      sel = normalizeSel({ r: anchor.r, c: 0 }, { r: p.r, c: nCols - 1 });
    else if (dragging === 'col') sel = normalizeSel({ r: 0, c: anchor.c }, { r: shown - 1, c: p.c });
    else                         sel = normalizeSel(anchor, p);
    if (dragging) schedulePaint(); else paint();
  }

  // The corner that stays put while its opposite is dragged.
  const OPPOSITE = { tl: b => ({ r: b.r2, c: b.c2 }), tr: b => ({ r: b.r2, c: b.c1 }),
                     bl: b => ({ r: b.r1, c: b.c2 }), br: b => ({ r: b.r1, c: b.c1 }) };

  let resizing = null;      // { block: 'main' | index }

  function attachHandles() {
    canvas.addEventListener('mousedown', e => {
      const h = e.target.closest('.pl-handle');
      if (!h || e.button !== 0) return;
      const key = h.dataset.block;
      const box = key === 'main' ? sel : extras[Number(key)];
      if (!box) return;
      e.preventDefault();
      e.stopPropagation();
      anchor   = OPPOSITE[h.dataset.h](box);
      resizing = { block: key === 'main' ? 'main' : Number(key) };
      dragging = 'resize';
      canvas.classList.add('pl-grid-resizing');
    });
  }

  function attach(table) {
    table.addEventListener('mousedown', e => {
      // Only the left button selects. Without this a right-click starts a
      // drag and collapses the range to the cell under the pointer before
      // the context menu even opens.
      if (e.button !== 0) return;
      const p = pointFrom(e.target);
      if (!p) return;
      e.preventDefault();                 // stop the browser's own text selection

      // Ctrl or Cmd starts an additional block rather than a new selection.
      if ((e.ctrlKey || e.metaKey || selectMode === 'add') && sel) {
        if (occupiedRows(-1).has(p.r)) return;   // that row is already taken
        addingExtra = { r1: p.r, r2: p.r, c1: sel.c1, c2: sel.c2 };
        extras.push(addingExtra);
        anchor   = p;
        dragging = 'extra';
        paint();
        return;
      }

      if (e.shiftKey && sel && p.kind === 'cell') {
        anchor   = anchor || { r: sel.r1, c: sel.c1 };
        dragging = 'cell';
        extendTo(p);
        return;
      }
      dragging = p.kind;
      anchor   = p;
      if (extras.length) extras = [];
      extendTo(p);
    });

    table.addEventListener('contextmenu', e => {
      const p = pointFrom(e.target);
      if (!p || !contextHandler) return;
      e.preventDefault();
      focusColumn(p.c, { toggle: false });
      contextHandler(p.c, e.pageX, e.pageY,
        { row: p.r, extraIndex: extraAt(p.r), inSelection: rowsSelected(p.r) });
    });

    const track = e => {
      if (!dragging) return;
      const p = pointFrom(e.target);
      if (p) extendTo(p);
    };
    table.addEventListener('mousemove', track);
    canvas.addEventListener('mousemove', track);
  }

  // Finish the drag wherever the mouse is released, including outside the grid.
  document.addEventListener('mouseup', e => {
    if (canvas) canvas.classList.remove('pl-grid-resizing');
    if (!dragging || (e && e.button !== 0)) return;
    dragging = null;
    addingExtra = null;
    resizing = null;
    // A block dragged to nothing is a mis-click, not an empty block.
    extras = extras.filter(b => b.r2 >= b.r1);
    extras.sort((a, b) => a.r1 - b.r1);
    if (mergeTouchingBlocks()) paint();
    if (onSelect) onSelect(sel, extras.map(b => ({ ...b })));
  });

  return { render, setSelection, getSelection, getExtras, setExtras, removeExtra, extraAt,
           setSelectMode, getSelectMode, setColumnInfo, focusColumn, clearFocus,
           onColumnContextMenu, rangeRef: () => rangeRef(sel) };
}
