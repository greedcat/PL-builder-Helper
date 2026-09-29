// ── Packing List Builder: floating quick summary ──
// Excel's status-bar figures for whatever is selected, in a small box fixed
// to the bottom right: where the selection is, how big, and Count, Sum,
// Average, Min and Max of its numbers. Fed by the sheet grid's highlight and
// the preview's dragged block; whichever changed last is shown.
// ─────────────────────────────────────────────────────────────
let quickSumSource = null;     // 'Sheet' | 'Preview' | null

// A cell's number, if it is one. Thousands separators are allowed, since the
// preview shows its numbers tidied.
function quickSumNumber(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (v == null) return null;
  const t = String(v).trim().replace(/,/g, '');
  return /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : null;
}

function quickSumFormat(n) {
  return n.toLocaleString(undefined, { maximumFractionDigits: 4 });
}

// info: { source, ref, rows, cols, values } — values is every selected cell,
// null or '' for an empty one. A single cell shows nothing, as in Excel.
function showQuickSum(info) {
  const box = document.getElementById('plQuickSum');
  if (!box) return;
  if (!info || info.rows * info.cols < 2) { hideQuickSum(info && info.source); return; }

  const filled = info.values.filter(v => v != null && String(v).trim() !== '');
  const nums   = filled.map(quickSumNumber).filter(n => n !== null);
  // Loops rather than Math.min(...nums): a whole-sheet highlight can hold far
  // more values than a function call can take as arguments.
  let sum = 0, min = Infinity, max = -Infinity;
  for (const n of nums) { sum += n; if (n < min) min = n; if (n > max) max = n; }
  const stat   = (label, value) =>
    `<span class="pl-qs-stat"><span class="pl-qs-k">${label}</span><span class="pl-qs-v">${value}</span></span>`;

  const size = `${info.rows} row${info.rows === 1 ? '' : 's'} × ${info.cols} column${info.cols === 1 ? '' : 's'}`;
  box.innerHTML = `
    <div class="pl-qs-head">
      <span class="pl-qs-src">${escapeHtml(info.source)}</span>
      <span class="pl-qs-ref">${escapeHtml(info.ref)}</span>
      <span class="pl-qs-size">${size}</span>
      <button type="button" class="pl-qs-close" aria-label="Close">×</button>
    </div>
    <div class="pl-qs-stats">
      ${stat('Count', filled.length)}
      ${nums.length ? `
        ${stat('Sum', quickSumFormat(sum))}
        ${stat('Average', quickSumFormat(sum / nums.length))}
        ${stat('Min', quickSumFormat(min))}
        ${stat('Max', quickSumFormat(max))}
        ${nums.length !== filled.length ? stat('Numbers', nums.length) : ''}`
      : '<span class="pl-qs-none">No numbers selected</span>'}
    </div>`;
  box.querySelector('.pl-qs-close').addEventListener('click', () => hideQuickSum());
  box.hidden = false;
  quickSumSource = info.source;
}

// With a source given, hides only if that source is what is on show, so the
// grid clearing its highlight does not close the preview's figures.
function hideQuickSum(source) {
  const box = document.getElementById('plQuickSum');
  if (!box) return;
  if (source && quickSumSource && source !== quickSumSource) return;
  box.hidden = true;
  box.innerHTML = '';
  quickSumSource = null;
}
