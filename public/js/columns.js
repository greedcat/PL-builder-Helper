// ─────────────────────────────────────────────────────────────
// COLUMN DETECTION
// ─────────────────────────────────────────────────────────────
function detectDestinationColumn(headers, rows) {
  const destSet = new Set(DEST_LIST.map(normalizeText));
  let bestCol = null, bestScore = 0;
  for (const col of headers) {
    if (!col) continue;
    const ci   = headers.indexOf(col);
    const vals = rows.map(r => r[ci]).filter(v => !isEmpty(v)).map(normalizeText);
    if (!vals.length) continue;
    const matchCount = vals.filter(v => [...destSet].some(d => v.includes(d))).length;
    const r     = matchCount / vals.length;
    const score = r > 0.3 ? 5 : r > 0.1 ? 2 : 0;
    if (score > bestScore) { bestScore = score; bestCol = col; }
  }
  return bestCol;
}

function isValidFBA(v) { return /^FBA[A-Z0-9]{8,12}$/.test(v); }

function detectFBAColumn(headers, rows) {
  let bestCol = null, bestScore = 0;
  for (const col of headers) {
    if (!col) continue;
    let score = 0;
    const ci   = headers.indexOf(col);
    const vals = rows.map(r => r[ci]).filter(v => !isEmpty(v)).map(normalizeText);
    if (vals.length > 0) {
      const ratio = vals.filter(isValidFBA).length / vals.length;
      if      (ratio > 0.5) score += 6;
      else if (ratio > 0.3) score += 3;
      else if (ratio > 0.1) score += 1;
    }
    if (score > bestScore) { bestScore = score; bestCol = col; }
  }
  return bestCol;
}

function isValidPO(v) {
  return v.length === 8 && !v.startsWith('FBA') && /^[A-Z0-9]{8}$/.test(v) && !/^\d+$/.test(v);
}

function detectAmazonPOColumn(headers, rows) {
  let bestCol = null, bestScore = 0;
  for (const col of headers) {
    if (!col) continue;
    let score = 0;
    const ci   = headers.indexOf(col);
    const vals = rows.map(r => r[ci]).filter(v => !isEmpty(v)).map(normalizeText);

    if (vals.length > 0) {
      const ratio = vals.filter(isValidPO).length / vals.length;
      if      (ratio > 0.5) score += 6;
      else if (ratio > 0.3) score += 3;
      else if (ratio > 0.1) score += 1;
    }

    // Header bonus: word "PO" in column name (word-boundary avoids "DEPOT" etc.)
    if (/\bPO\b/i.test(String(col))) score += 3;

    if (score > bestScore) { bestScore = score; bestCol = col; }
  }
  return bestScore >= 3 ? bestCol : null;
}

function detectCartonColumn(headers) {
  for (const col of headers) {
    if (!col) continue;
    const low = String(col).toLowerCase().trim();
    if (CARTON_KEYWORDS.some(k => k.toLowerCase() === low)) return col;
  }
  return null;
}

function detectWeightColumn(headers) {
  for (const col of headers) {
    if (!col) continue;
    if (WEIGHT_KEYWORDS.some(k => normWeight(k) === normWeight(col))) return col;
  }
  return null;
}

function detectCBMColumn(headers) {
  const kwSet   = new Set(CBM_KEYWORDS.map(normCBM).filter(Boolean));
  const subKeys = ['cbm', 'cmb', 'volume', '体积', '方数'];
  let bestCol = null, bestScore = 0;

  for (const col of headers) {
    if (!col) continue;
    const n = normCBM(col);
    let score = 0;
    if      (kwSet.has(n))                          score = 2; // exact keyword match
    else if (subKeys.some(t => n.includes(t)))      score = 1; // substring fallback

    if (score > bestScore) { bestScore = score; bestCol = col; }
  }
  return bestCol;
}