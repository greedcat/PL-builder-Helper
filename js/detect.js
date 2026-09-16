// ─────────────────────────────────────────────────────────────
// TABLE STRUCTURE DETECTION
// ─────────────────────────────────────────────────────────────
function detectHeader(rawRows, lookAhead = 3) {
  let bestScore = -1, bestRow = 0;

  // Pre-build normalised keyword sets for fast lookup
  const ctnSet    = new Set(CARTON_KEYWORDS.map(k => k.toLowerCase().trim()));
  const cbmSet    = new Set(CBM_KEYWORDS.map(normCBM).filter(Boolean));
  const wgtSet    = new Set(WEIGHT_KEYWORDS.map(normWeight).filter(Boolean));
  const noNeedSet = new Set(NO_NEED_COL.map(normalizeText).filter(Boolean));

  for (let i = 0; i < rawRows.length - lookAhead; i++) {
    const row = rawRows[i];
    let score = 0;

    // All-string row
    if (row.every(v => isEmpty(v) || typeof v === 'string')) score += 5;

    // Non-empty cell count
    score += row.filter(v => !isEmpty(v)).length * 2;

    // Numeric values in following rows
    for (let j = i + 1; j <= i + lookAhead && j < rawRows.length; j++) {
      for (const v of rawRows[j]) {
        if (!isEmpty(v) && !isNaN(Number(v))) score += 2;
      }
    }

    // Bonus: cell matches a known column keyword
    for (const v of row) {
      if (isEmpty(v)) continue;
      if (ctnSet.has(String(v).toLowerCase().trim()) ||
          cbmSet.has(normCBM(v)) ||
          wgtSet.has(normWeight(v)) ||
          noNeedSet.has(normalizeText(v))) score += 3;
    }

    if (score > bestScore) { bestScore = score; bestRow = i; }
  }
  return bestRow;
}

function detectTableEnd(rawRows, headerRow, fillThreshold = 0.4, sumMultiplier = 3) {
  const dataStart   = headerRow + 1;

  // Hard stop: row containing this phrase marks the end of the table
  const END_MARKER = 'ensure deliveries meet the operational requirements as stipulated in the';
  for (let i = dataStart; i < rawRows.length; i++) {
    if (rawRows[i].some(v => !isEmpty(v) && String(v).toLowerCase().includes(END_MARKER))) {
      return [Math.max(i - 1, dataStart), []];
    }
  }

  const runningRows = [];
  const scores      = [];
  // Use the header row's non-empty cell count as effective table width to avoid
  // over-counting ghost columns from SheetJS's used-range (matches df.shape[1]).
  const numCols = rawRows[headerRow].filter(v => !isEmpty(v)).length || (rawRows[0] ? rawRows[0].length : 1);

  for (let i = dataStart; i < rawRows.length; i++) {
    const row          = rawRows[i];
    const nonEmpty     = row.filter(v => !isEmpty(v)).length;
    const numericRow   = row.map(v => { const n = Number(v); return (!isEmpty(v) && !isNaN(n)) ? n : null; });
    const numericCount = numericRow.filter(v => v !== null).length;
    let score = 0;

    // Signal 1: sparse row
    const fillRatio = nonEmpty / numCols;
    if      (fillRatio < fillThreshold) score += 70;
    else if (fillRatio < 0.6)           score += 10;

    // Signal 2: completely empty
    if (nonEmpty === 0) score += 100;

    // Signal 3: no numeric values at all
    if (numericCount === 0) score += 20;

    // Signal 4: numeric values look like column totals
    if (runningRows.length >= 3) {
      const colMeans = Array.from({ length: numCols }, (_, c) => {
        const vals = runningRows.map(r => r[c]).filter(v => v !== null);
        return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
      });
      let outlierCols = 0;
      for (let c = 0; c < numCols; c++) {
        if (numericRow[c] !== null && colMeans[c] !== null && colMeans[c] > 0 && numericRow[c] > colMeans[c] * sumMultiplier)
          outlierCols++;
      }
      score += outlierCols * 15;
    }

    scores.push([i, score]);
    if (score < 20) runningRows.push(numericRow);
  }

  if (!scores.length) return [dataStart, scores];

  // A row only marks the end of the table if it actually looks like one.
  // Below this, the signals are ordinary variation between data rows.
  const END_SCORE = 40;
  const maxScoreRow = scores.reduce((a, b) => b[1] > a[1] ? b : a);

  // Nothing resembles a footer, summary or spacer, so the table runs to the
  // last row. Without this a perfectly clean sheet collapses to one row,
  // because every score ties at zero and the reducer keeps the first.
  if (maxScoreRow[1] < END_SCORE) return [scores[scores.length - 1][0], scores];

  let lastDataRow = dataStart;
  for (const [idx] of scores) {
    if (idx < maxScoreRow[0]) lastDataRow = idx;
  }
  return [lastDataRow, scores];
}
