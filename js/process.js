// ─────────────────────────────────────────────────────────────
// DATA PROCESSING
// ─────────────────────────────────────────────────────────────
function readExcel(headers, rows) {
  const destCol   = detectDestinationColumn(headers, rows);
  const fbaCol    = detectFBAColumn(headers, rows);
  const poCol     = detectAmazonPOColumn(headers, rows);
  const ctnCol    = detectCartonColumn(headers);
  const weightCol = detectWeightColumn(headers);
  const cbmCol    = detectCBMColumn(headers);

  console.log('Column map:', { destCol, fbaCol, poCol, ctnCol, weightCol, cbmCol });

  const matchDict = {
    'Destination': destCol,
    'Carton':      ctnCol,
    'FBA_ID':      fbaCol,
    'REF ID':      poCol,
    'Weight':      weightCol,
    'CMB':         cbmCol
  };

  // Use indices to preserve null/empty-header columns
  const mainNames = [destCol, ctnCol, fbaCol, poCol, weightCol, cbmCol].filter(Boolean);
  const mainIdx   = mainNames.map(h => headers.indexOf(h));
  const otherIdx  = headers
    .map((h, i) => i)
    .filter(i => !mainIdx.includes(i) && !NO_NEED_COL.includes(headers[i]));
  const orderIdx  = [...mainIdx, ...otherIdx];

  return {
    headers:   orderIdx.map(i => headers[i]),
    rows:      rows.map(r => orderIdx.map(i => r[i])),
    matchDict
  };
}

function modifyDF(headers, rows, matchDict) {
  // Forward-fill destination
  const destIdx = headers.indexOf(matchDict['Destination']);
  let lastDest  = null;
  const filled  = rows.map(r => {
    const row = [...r];
    if (!isEmpty(row[destIdx])) lastDest = row[destIdx];
    else row[destIdx] = lastDest;
    return row;
  });

  // Reorder: matched columns first, then extras (index-based to preserve null headers)
  const matched    = Object.values(matchDict).filter(v => v && v !== 'null');
  const matchedIdx = matched.map(h => headers.indexOf(h));
  const extraIdx   = headers.map((h, i) => i).filter(i => !matchedIdx.includes(i));
  const orderIdx   = [...matchedIdx, ...extraIdx];

  const revMap = Object.fromEntries(
    Object.entries(matchDict).filter(([, v]) => v && v !== 'null').map(([k, v]) => [v, k])
  );
  const newHeaders = orderIdx.map(i => revMap[headers[i]] || headers[i]);

  let data = filled.map(r =>
    orderIdx.map(i => {
      const v = r[i];
      return (typeof v === 'string' && v.trim() === '') ? null : (v === undefined ? null : v);
    })
  );
  data = data.filter(r => r.some(v => v !== null));

  // Add _Pallet column
  let finalHeaders = [...newHeaders];
  const cbmIdx     = newHeaders.indexOf('CMB');
  if (cbmIdx >= 0) {
    finalHeaders.push('_Pallet');
    data = data.map(r => { const cbm = parseFloat(r[cbmIdx]); return [...r, isNaN(cbm) ? null : cbm / 1.7]; });
  }

  // Sort by Destination (stable)
  const destNewIdx = finalHeaders.indexOf('Destination');
  data = data
    .map((r, i) => [r, i])
    .sort(([a, ai], [b, bi]) => {
      const da = String(a[destNewIdx] ?? ''), db = String(b[destNewIdx] ?? '');
      return da < db ? -1 : da > db ? 1 : ai - bi;
    })
    .map(([r]) => r);

  // Normalize destination codes to canonical form
  function matchDest(val) {
    if (val == null) return val;
    const s = String(val).trim();
    for (const code of DEST_LIST) { if (s.includes(code)) return code; }
    return s;
  }
  data = data.map(r => {
    const row = [...r];
    if (destNewIdx >= 0) row[destNewIdx] = matchDest(row[destNewIdx]);
    return row;
  });

  return { headers: finalHeaders, rows: data };
}

function getSummaryDF(headers, rows) {
  const di     = headers.indexOf('Destination');
  const ci     = headers.indexOf('Carton');
  const pi     = headers.indexOf('_Pallet');
  const groups = new Map();

  for (const row of rows) {
    const dest = row[di];
    const key  = dest != null ? String(dest) : '';
    if (!groups.has(key)) groups.set(key, { Destination: dest, Carton: 0, pallet: 0 });
    const g = groups.get(key);
    const ctn = parseFloat(row[ci]); if (!isNaN(ctn)) g.Carton += ctn;
    const plt = parseFloat(row[pi]); if (!isNaN(plt)) g.pallet += plt;
  }

  return {
    headers: ['Destination', 'Carton', 'pallet_round_count', 'pallet_number'],
    rows:    [...groups.values()].map(g => [g.Destination, g.Carton, Math.ceil(g.pallet - 0.05), g.pallet])
  };
}
