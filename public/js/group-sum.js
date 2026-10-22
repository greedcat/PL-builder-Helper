// ─────────────────────────────────────────────────────────────
// Excel Group & Sum
// A sheet (a file, or cells pasted from Excel) is shown on the same grid as
// the packing list builder's. Its blue box is the range: the top row is the
// header, the rows under it are the data. The box is found on its own and
// can be dragged, typed in by row number, or set from a highlight. The data
// is grouped by the chosen columns and the chosen numbers added up; the
// result follows every change and can be copied, or saved as .xlsx / .csv.
// ─────────────────────────────────────────────────────────────
(() => {
  const $ = id => document.getElementById(id);
  const fileInput = $('gsFile'), drop = $('gsDrop'), paste = $('gsPaste');
  const statusEl = $('gsStatus');

  let wb = null;          // the workbook, from a file or the paste
  let sheetName = null;
  let raw = [];           // the sheet as rows of values, merged cells filled
  let maxCols = 0;
  let range = null;       // { r1, r2, c1, c2 }: header row r1, data r1+1..r2
  let groupCols = [];     // column indexes, in order
  let sumCols = null;     // Set of column indexes; null = the suggested ones
  let result = null;      // { head: [...], rows: [[...]], total: [...] }

  function say(msg, kind = 'info') {
    statusEl.hidden = !msg;
    statusEl.className = `gs-status gs-status-${kind}`;
    statusEl.innerHTML = msg || '';
  }

  // ── Loading ────────────────────────────────────────────────
  drop.addEventListener('click', e => { if (e.target !== fileInput) fileInput.click(); });
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('active'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('active'));
  drop.addEventListener('drop', e => {
    e.preventDefault(); drop.classList.remove('active');
    if (e.dataTransfer.files[0]) readFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', () => { if (fileInput.files[0]) readFile(fileInput.files[0]); });

  function readFile(file) {
    say(`Reading ${escapeHtml(file.name)}…`);
    const fr = new FileReader();
    fr.onload = () => {
      try {
        wb = XLSX.read(new Uint8Array(fr.result), { type: 'array', cellDates: true });
        loaded(file.name);
      } catch (err) {
        say(`Could not read ${escapeHtml(file.name)}: ${escapeHtml(err.message)}`, 'error');
      }
    };
    fr.readAsArrayBuffer(file);
  }

  // Cells copied in Excel arrive as tab-separated text: one sheet of it.
  paste.addEventListener('paste', e => {
    const text = (e.clipboardData || window.clipboardData).getData('text/plain');
    if (!text || !text.includes('\t') && !text.includes('\n')) return;
    e.preventDefault();
    const rows = parseTsv(text);
    if (!rows.length) return;
    wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Pasted');
    paste.value = '';
    loaded('Pasted from Excel');
  });

  function loaded(name) {
    $('gsBadge').innerHTML = `<span>📄</span> ${escapeHtml(name)}`;
    $('gsBadge').style.display = 'inline-flex';
    drop.querySelector('.drop-text').style.display = 'none';
    drop.classList.add('pl-drop-compact');
    const pick = $('gsSheet');
    pick.innerHTML = wb.SheetNames.map(n => `<option>${escapeHtml(n)}</option>`).join('');
    $('gsSheetWrap').hidden = wb.SheetNames.length < 2;
    openSheet(wb.SheetNames[0]);
  }
  $('gsSheet').addEventListener('change', e => openSheet(e.target.value));

  function openSheet(name) {
    sheetName = name;
    const ws = wb.Sheets[name];
    raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
    // A merged cell (a destination written once over its lines) counts on
    // every line it covers.
    fillMergedRows(ws, raw);
    maxCols = raw.reduce((m, r) => Math.max(m, lastFilled(r) + 1), 0);
    if (!raw.length || !maxCols) {
      say('That sheet is empty.', 'error');
      ['gsRangeSec', 'gsSetupSec', 'gsResultSec'].forEach(id => { $(id).hidden = true; });
      return;
    }
    say('');
    groupCols = [];
    sumCols = null;
    range = autoRange();
    $('gsRangeSec').hidden = false;
    grid.render(raw, range, maxCols, ws['!merges'] || []);
    rangeChanged();
  }

  const filled = v => v != null && String(v).trim() !== '';
  const lastFilled = r => { for (let i = (r || []).length - 1; i >= 0; i--) if (filled(r[i])) return i; return -1; };

  // The header is the first row about as wide as the widest row (a title
  // above it is one or two cells); the data runs to the last filled row.
  function autoRange() {
    const counts = raw.map(r => (r || []).filter(filled).length);
    const widest = Math.max(...counts);
    let r1 = counts.findIndex(n => n >= Math.max(2, Math.ceil(widest * 0.6)));
    if (r1 < 0) r1 = 0;
    let r2 = raw.length - 1;
    while (r2 > r1 && !counts[r2]) r2--;
    const hdr = raw[r1] || [];
    let c1 = hdr.findIndex(filled), c2 = lastFilled(hdr);
    if (c1 < 0) { c1 = 0; c2 = maxCols - 1; }
    return { r1, r2, c1, c2 };
  }

  // ── The grid and its range ─────────────────────────────────
  const grid = createSheetGrid({
    mount: $('gsGrid'),
    onSelect: sel => { if (sel) { range = { ...sel }; rangeChanged(); } },
  });
  grid.onHighlightChange(cells => {
    if (cells && cells.rows * cells.cols > 1) {
      showQuickSum({ source: 'Sheet', ref: cells.ref, rows: cells.rows, cols: cells.cols, values: cells.values });
    } else hideQuickSum('Sheet');
  });
  // A highlight right-clicked becomes the range.
  grid.onColumnContextMenu((abs, x, y, where) => {
    const h = where && where.highlight;
    if (!h) return;
    openMenu(x, y, [{ label: `Use ${rangeRefOf(h)} as the range`, onPick: () => {
      grid.useHighlightAsRange();
    } }]);
  });
  const rangeRefOf = s => `${cellRef(s.r1, s.c1)}:${cellRef(s.r2, s.c2)}`;

  $('gsAuto').addEventListener('click', () => {
    range = autoRange();
    grid.setSelection(range);
    rangeChanged();
  });
  const rowInput = (el, set) => el.addEventListener('change', () => {
    const n = parseInt(el.value, 10);
    if (!n || n < 1) { rangeChanged(); return; }
    set(Math.min(n - 1, raw.length - 1));
    if (range.r2 < range.r1) range.r2 = range.r1;
    grid.setSelection(range);
    rangeChanged();
  });
  rowInput($('gsHdr'), r => { range.r1 = r; });
  rowInput($('gsLast'), r => { range.r2 = r; });

  // ── Columns of the range ───────────────────────────────────
  function columns() {
    const hdr = raw[range.r1] || [];
    const out = [];
    for (let c = range.c1; c <= range.c2; c++) {
      const name = filled(hdr[c]) ? String(hdr[c]).replace(/\s+/g, ' ').trim() : `Column ${colLetter(c)}`;
      out.push({ c, name });
    }
    return out;
  }
  const num = v => {
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    if (!filled(v)) return null;
    const n = Number(String(v).replace(/[,\s]/g, ''));
    return Number.isFinite(n) ? n : null;
  };
  // A column is a number column when most of its filled cells are numbers.
  function isNumeric(c) {
    let n = 0, all = 0;
    for (let r = range.r1 + 1; r <= range.r2; r++) {
      const v = (raw[r] || [])[c];
      if (!filled(v)) continue;
      all++;
      if (num(v) !== null) n++;
    }
    return all > 0 && n / all >= 0.6;
  }

  // What was chosen stays chosen while its columns are still in the range.
  function rangeChanged() {
    const inRange = c => c >= range.c1 && c <= range.c2;
    groupCols = groupCols.filter(inRange);
    if (sumCols) sumCols = new Set([...sumCols].filter(inRange));
    $('gsRef').textContent = rangeRefOf(range);
    $('gsHdr').value = range.r1 + 1;
    $('gsLast').value = range.r2 + 1;
    const n = Math.max(0, range.r2 - range.r1);
    $('gsExplain').innerHTML = `Row <strong>${range.r1 + 1}</strong> is the header · rows <strong>${range.r1 + 2}–${range.r2 + 1}</strong> (${n}) are the data`;
    renderSetup();
    run();
  }

  // ── Setup: group by and add up ─────────────────────────────
  function renderSetup() {
    $('gsSetupSec').hidden = false;
    const cols = columns();
    if (!sumCols) sumCols = new Set(cols.filter(x => isNumeric(x.c) && !groupCols.includes(x.c)).map(x => x.c));
    const tags = $('gsGroupTags');
    tags.querySelectorAll('.gs-tag').forEach(t => t.remove());
    const pick = $('gsGroupPick');
    for (const c of groupCols) {
      const col = cols.find(x => x.c === c);
      if (!col) continue;
      const t = document.createElement('span');
      t.className = 'gs-tag';
      t.innerHTML = `${escapeHtml(col.name)} <button type="button" aria-label="Remove ${escapeHtml(col.name)}" data-c="${c}">×</button>`;
      tags.insertBefore(t, pick);
    }
    pick.innerHTML = '<option value="">+ add column</option>'
      + cols.filter(x => !groupCols.includes(x.c)).map(x => `<option value="${x.c}">${colLetter(x.c)} — ${escapeHtml(x.name)}</option>`).join('');
    $('gsSums').innerHTML = cols.filter(x => !groupCols.includes(x.c)).map(x => {
      const numeric = isNumeric(x.c);
      return `<label class="gs-sum${numeric ? '' : ' gs-sum-text'}" title="${numeric ? '' : 'Mostly not numbers'}">
        <input type="checkbox" data-c="${x.c}"${sumCols.has(x.c) ? ' checked' : ''}> ${escapeHtml(x.name)}</label>`;
    }).join('') || '<span class="gs-hint">No other columns.</span>';
  }
  $('gsGroupPick').addEventListener('change', e => {
    const c = Number(e.target.value);
    if (e.target.value === '' || groupCols.includes(c)) return;
    groupCols.push(c);
    sumCols.delete(c);
    renderSetup(); run();
  });
  $('gsGroupTags').addEventListener('click', e => {
    const b = e.target.closest('button[data-c]');
    if (!b) return;
    groupCols = groupCols.filter(c => c !== Number(b.dataset.c));
    renderSetup(); run();
  });
  $('gsSums').addEventListener('change', e => {
    const box = e.target.closest('input[data-c]');
    if (!box) return;
    const c = Number(box.dataset.c);
    if (box.checked) sumCols.add(c); else sumCols.delete(c);
    run();
  });
  ['gsCount', 'gsFill', 'gsSkipTotals'].forEach(id => $(id).addEventListener('change', run));

  // ── Grouping ───────────────────────────────────────────────
  // "Total", "Subtotal:", "合计" — followed by nothing that makes it a word.
  const TOTAL_RE = /^(total|totals|sub\s*total|grand\s*total|合计|合計|总计|總計|小计|小計)(?![a-z])/i;

  function run() {
    const cols = columns();
    const nameOf = c => (cols.find(x => x.c === c) || {}).name || colLetter(c);
    const sums = cols.map(x => x.c).filter(c => sumCols.has(c) && !groupCols.includes(c));
    const count = $('gsCount').checked, fill = $('gsFill').checked, skipTotals = $('gsSkipTotals').checked;
    const groups = new Map();
    const last = {};
    let used = 0, skipped = 0;
    for (let r = range.r1 + 1; r <= range.r2; r++) {
      const row = raw[r] || [];
      const cells = row.slice(range.c1, range.c2 + 1);
      if (!cells.some(filled)) continue;                          // a blank line
      if (skipTotals && cells.some(v => filled(v) && TOTAL_RE.test(String(v).trim()))) { skipped++; continue; }
      const key = groupCols.map(c => {
        let v = row[c];
        if (fill && !filled(v)) v = last[c];
        if (filled(v)) last[c] = v;
        return filled(v) ? String(v).trim() : '';
      });
      const k = key.join('\u0000');
      if (!groups.has(k)) groups.set(k, { key, n: 0, sums: sums.map(() => 0) });
      const g = groups.get(k);
      g.n++;
      sums.forEach((c, i) => { const v = num(row[c]); if (v !== null) g.sums[i] += v; });
      used++;
    }
    const list = [...groups.values()];
    // Sorted by the group columns, numbers as numbers.
    const cmp = (a, b) => {
      for (let i = 0; i < groupCols.length; i++) {
        const x = a.key[i], y = b.key[i];
        const nx = num(x), ny = num(y);
        const d = nx !== null && ny !== null ? nx - ny : x.localeCompare(y, undefined, { numeric: true });
        if (d) return d;
      }
      return 0;
    };
    list.sort(cmp);
    const round = v => Math.round(v * 1e6) / 1e6;
    const head = [...groupCols.map(nameOf), ...(count ? ['Count'] : []), ...sums.map(nameOf)];
    const rows = list.map(g => [...g.key, ...(count ? [g.n] : []), ...g.sums.map(round)]);
    const total = [...groupCols.map((c, i) => (i === 0 ? 'Total' : '')),
      ...(count ? [used] : []), ...sums.map((c, i) => round(list.reduce((s, g) => s + g.sums[i], 0)))];
    if (!groupCols.length) total[0] = 'Total';
    result = { head, rows, total };
    renderResult(skipped);
  }

  function renderResult(skipped) {
    $('gsResultSec').hidden = false;
    const { head, rows, total } = result;
    const g = groupCols.length;
    const cell = (v, i) => `<td class="${i >= g ? 'pl-num' : ''}">${escapeHtml(v == null ? '' : typeof v === 'number' ? v.toLocaleString(undefined, { maximumFractionDigits: 6 }) : v)}</td>`;
    $('gsResult').innerHTML = `<table><thead><tr>${head.map((h, i) => `<th class="${i >= g ? 'pl-num' : ''}">${escapeHtml(h)}</th>`).join('')}</tr></thead><tbody>`
      + (g ? rows.map(r => `<tr>${r.map(cell).join('')}</tr>`).join('') : '')
      + `<tr class="gs-total">${total.map(cell).join('')}</tr></tbody></table>`;
    $('gsResultTitle').textContent = g ? `Grouped by ${groupCols.map(c => columns().find(x => x.c === c)?.name).join(' › ')}` : 'Totals';
    $('gsResultCount').textContent = (g ? `${rows.length} group${rows.length === 1 ? '' : 's'}` : 'whole range')
      + (skipped ? ` · ${skipped} total row${skipped === 1 ? '' : 's'} skipped` : '');
  }

  // ── Out ────────────────────────────────────────────────────
  const table = () => (result ? [result.head, ...(groupCols.length ? result.rows : []), result.total] : []);
  $('gsCopy').addEventListener('click', async () => {
    const text = table().map(r => r.map(v => (v == null ? '' : String(v))).join('\t')).join('\n');
    try { await navigator.clipboard.writeText(text); say('Copied — paste it into Excel.', 'success'); }
    catch (_) { say('The browser would not copy. Download the .xlsx instead.', 'error'); }
  });
  const stem = () => `group-sum_${(sheetName || 'sheet').replace(/[^\w一-鿿-]+/g, '_')}`;
  $('gsXlsx').addEventListener('click', () => {
    const out = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(table());
    ws['!cols'] = result.head.map((h, i) => ({ wch: Math.min(50, Math.max(10, ...table().map(r => String(r[i] ?? '').length + 2))) }));
    XLSX.utils.book_append_sheet(out, ws, 'Group & Sum');
    XLSX.writeFile(out, `${stem()}.xlsx`);
  });
  $('gsCsv').addEventListener('click', () => {
    const csv = table().map(r => r.map(v => `"${String(v ?? '').replace(/"/g, '""')}"`).join(',')).join('\r\n');
    const a = Object.assign(document.createElement('a'), {
      href: URL.createObjectURL(new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' })),
      download: `${stem()}.csv`,
    });
    a.click();
    URL.revokeObjectURL(a.href);
  });

  // ── A small menu, for the grid's right-click ───────────────
  let menu = null;
  function openMenu(x, y, items) {
    if (menu) menu.remove();
    menu = document.createElement('div');
    menu.className = 'pl-menu';
    menu.innerHTML = items.map((it, i) => `<button type="button" class="pl-menu-item" data-i="${i}">
      <span class="pl-menu-tick"></span><span class="pl-menu-label">${escapeHtml(it.label)}</span></button>`).join('');
    document.body.appendChild(menu);
    menu.style.left = `${Math.min(x, window.scrollX + document.documentElement.clientWidth - menu.offsetWidth - 8)}px`;
    menu.style.top = `${y}px`;
    menu.addEventListener('click', e => {
      const b = e.target.closest('.pl-menu-item');
      if (!b) return;
      const it = items[Number(b.dataset.i)];
      menu.remove(); menu = null;
      it.onPick();
    });
  }
  document.addEventListener('mousedown', e => { if (menu && !menu.contains(e.target)) { menu.remove(); menu = null; } });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && menu) { menu.remove(); menu = null; } });
})();
