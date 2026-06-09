// ─────────────────────────────────────────────────────────────
// EXCEL OUTPUT
// ─────────────────────────────────────────────────────────────
async function writePackingList(dataHeaders, dataRows, sumHeaders, sumRows, cName, containerName, wbIn) {
  const wb  = new ExcelJS.Workbook();
  const TNR = 'Times New Roman';
  const bdr = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

  function styledCell(ws, r, c, opts = {}) {
    const cl = ws.getCell(r, c);
    cl.alignment = { horizontal: 'center', vertical: 'middle', wrapText: opts.wrap || false };
    cl.font      = { name: TNR, size: opts.size || 16, bold: opts.bold || false };
    if (opts.border !== false) cl.border = bdr;
    return cl;
  }

  // Visible columns: drop _Pallet and all-null columns
  const visCols = dataHeaders
    .map((h, ci) => ({ h, ci }))
    .filter(({ h, ci }) => h !== '_Pallet' && dataRows.some(r => r[ci] != null));
  const visH    = visCols.map(x => x.h);
  const visRows = dataRows.map(r => visCols.map(x => r[x.ci]));

  // Summary: drop pallet_number column
  const sumVis  = sumHeaders.map((h, i) => ({ h, i })).filter(x => x.h !== 'pallet_number');
  const sumVisH = sumVis.map(x => x.h);
  const sumVisR = sumRows.map(r => sumVis.map(x => r[x.i]));

  // ── PL sheet ──────────────────────────────────────────────
  const wsPL = wb.addWorksheet('PL');

  wsPL.mergeCells('A1:I1');
  styledCell(wsPL, 1, 1, { size: 28, bold: true }).value = 'PACKING LIST AND DESTUFFING INSTRUCTION (FBA)';
  wsPL.getRow(1).height = 65;

  wsPL.mergeCells('A2:B2'); styledCell(wsPL, 2, 1, { size: 20, bold: true }).value = 'Client_Name';
  wsPL.mergeCells('C2:D2'); styledCell(wsPL, 2, 3, { size: 20, bold: true }).value = cName;
  wsPL.mergeCells('E2:F2'); styledCell(wsPL, 2, 5, { size: 20, bold: true }).value = 'Container #';
  wsPL.mergeCells('G2:I2'); styledCell(wsPL, 2, 7, { size: 20, bold: true }).value = containerName;
  wsPL.getRow(2).height = 65;

  const numDests = new Set(dataRows.map(r => r[dataHeaders.indexOf('Destination')]).filter(v => v != null)).size;
  wsPL.mergeCells('A3:B3'); styledCell(wsPL, 3, 1, { size: 20, bold: true }).value = '# of Destination';
  wsPL.mergeCells('C3:D3'); styledCell(wsPL, 3, 3, { size: 20, bold: true }).value = numDests;
  wsPL.mergeCells('E3:F3'); styledCell(wsPL, 3, 5, { size: 20, bold: true }).value = 'Destuffing Time';
  wsPL.mergeCells('G3:I3'); styledCell(wsPL, 3, 7, { size: 20, bold: true }).value = '';
  wsPL.getRow(3).height = 65;

  wsPL.getRow(4).height = 20; // blank separator

  const DATA_HDR_ROW = 5;
  visH.forEach((h, ci) => { styledCell(wsPL, DATA_HDR_ROW, ci + 1).value = h; });
  wsPL.getRow(DATA_HDR_ROW).height = 40;

  const DATA_START = 6;
  visRows.forEach((row, ri) => {
    row.forEach((v, ci) => { styledCell(wsPL, DATA_START + ri, ci + 1).value = v ?? ''; });
    wsPL.getRow(DATA_START + ri).height = 30;
  });
  const lastDataRow = DATA_START + visRows.length - 1;

  // Merge consecutive identical Destination cells
  const destVisIdx = visH.indexOf('Destination') + 1;
  if (destVisIdx > 0) {
    let s = DATA_START;
    while (s <= lastDataRow) {
      let e = s;
      const val = wsPL.getCell(s, destVisIdx).value;
      while (e + 1 <= lastDataRow && wsPL.getCell(e + 1, destVisIdx).value === val) e++;
      if (e > s) wsPL.mergeCells(s, destVisIdx, e, destVisIdx);
      s = e + 1;
    }
  }

  // Embedded summary section
  const SUM_TITLE_ROW  = lastDataRow + 2;
  const SUM_HDR_ROW    = SUM_TITLE_ROW + 1;
  const SUM_DATA_START = SUM_HDR_ROW + 1;

  wsPL.mergeCells(SUM_TITLE_ROW, 1, SUM_TITLE_ROW, 9);
  styledCell(wsPL, SUM_TITLE_ROW, 1, { size: 36, bold: true }).value = 'SUMMARY';
  wsPL.getRow(SUM_TITLE_ROW).height = 85;

  const SUM_HEADERS = [
    'Destination\n地址',
    'Total Carton\n箱数',
    'Estimated Skid #\n预计托盘数量',
    'Actual Skid # (Standard)\n实际打托数量（标准）',
    'Actual Skid # (60 Inches)\n实际打托数量（长板60寸）',
    'Actual Skid # (75 Inches)\n实际打托数量（长板75寸）',
    'Actual Skid # (96 Inches)\n实际打托数量（长板96寸）',
    'Standard plt but overhanging',
    'Self Palletized\n自托'
  ];
  SUM_HEADERS.forEach((lbl, ci) => {
    styledCell(wsPL, SUM_HDR_ROW, ci + 1, { size: 18, bold: true, wrap: true }).value = lbl;
  });
  wsPL.getRow(SUM_HDR_ROW).height = 125;

  sumVisR.forEach((row, ri) => {
    const rn = SUM_DATA_START + ri;
    for (let ci = 0; ci < 9; ci++) {
      styledCell(wsPL, rn, ci + 1, { size: 18 }).value = ci < row.length ? (row[ci] ?? '') : '';
    }
    wsPL.getRow(rn).height = 85;
  });

  // Auto-fit column widths
  const plLastRow = SUM_DATA_START + sumVisR.length - 1;
  for (let c = 1; c <= Math.max(visH.length, 9); c++) {
    let maxLen = 0;
    for (let r = 1; r <= plLastRow; r++) {
      const v = wsPL.getCell(r, c).value;
      if (v != null) maxLen = Math.max(maxLen, ...String(v).split('\n').map(l => l.length));
    }
    wsPL.getColumn(c).width = Math.min(maxLen * 1.2 + 4, 90);
  }

  // ── Summary sheet ─────────────────────────────────────────
  const wsSUM = wb.addWorksheet('summary');
  sumVisH.forEach((h, ci) => { styledCell(wsSUM, 1, ci + 1).value = h; });
  wsSUM.getRow(1).height = 25;
  sumVisR.forEach((row, ri) => {
    row.forEach((v, ci) => { styledCell(wsSUM, 2 + ri, ci + 1).value = v ?? ''; });
    wsSUM.getRow(2 + ri).height = 25;
  });
  for (let c = 1; c <= sumVisH.length; c++) {
    let maxLen = 0;
    for (let r = 1; r <= 1 + sumVisR.length; r++) {
      const v = wsSUM.getCell(r, c).value;
      if (v != null) maxLen = Math.max(maxLen, String(v).length);
    }
    wsSUM.getColumn(c).width = Math.min(maxLen * 1.2 + 4, 50);
  }

  // ── Original sheets ───────────────────────────────────────
  if (wbIn) {
    for (const sheetName of wbIn.SheetNames) {
      const sjs = wbIn.Sheets[sheetName];
      let name  = sheetName.replace(/[\\/?*[\]:]/g, '_').slice(0, 31);
      if (name === 'PL' || name === 'summary') name += '_orig';
      const ws = wb.addWorksheet(name);

      if (sjs['!ref']) {
        const range     = XLSX.utils.decode_range(sjs['!ref']);
        const colMaxLen = {};

        for (let r = range.s.r; r <= range.e.r; r++) {
          for (let c = range.s.c; c <= range.e.c; c++) {
            const cell = sjs[XLSX.utils.encode_cell({ r, c })];
            if (!cell) continue;
            ws.getCell(r + 1, c + 1).value = cell.v ?? null;
            const len = cell.v != null ? String(cell.v).length : 0;
            colMaxLen[c] = Math.max(colMaxLen[c] || 0, len);
          }
        }

        // Auto-fit columns: content-based width, floored at 8, capped at 60
        for (let c = range.s.c; c <= range.e.c; c++) {
          const contentW = (colMaxLen[c] || 0) * 1.15 + 2;
          const origW    = sjs['!cols']?.[c]?.wch || 0;
          ws.getColumn(c + 1).width = Math.min(Math.max(contentW, origW, 8), 60);
        }
      }

      if (sjs['!merges']) {
        for (const m of sjs['!merges']) {
          try { ws.mergeCells(m.s.r + 1, m.s.c + 1, m.e.r + 1, m.e.c + 1); } catch (_) {}
        }
      }

      // Restore original row heights
      if (sjs['!rows']) {
        sjs['!rows'].forEach((row, i) => {
          if (row?.hpt) ws.getRow(i + 1).height = row.hpt;
        });
      }
    }
  }

  // ── Download ──────────────────────────────────────────────
  const outBuf = await wb.xlsx.writeBuffer();
  const url    = URL.createObjectURL(new Blob([outBuf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  const a      = Object.assign(document.createElement('a'), { href: url, download: `PL(${containerName}).xlsx` });
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
