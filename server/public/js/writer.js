// ─────────────────────────────────────────────────────────────
// EXCEL OUTPUT
// ─────────────────────────────────────────────────────────────
async function writePackingList(dataHeaders, dataRows, sumHeaders, sumRows, cName, containerName, wbIn, fileName, fileNo) {
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
  const { headers: visH, rows: visRows } = getVisibleColumns(dataHeaders, dataRows);

  // Summary: drop pallet_number column
  const sumVis  = sumHeaders.map((h, i) => ({ h, i })).filter(x => x.h !== 'pallet_number');
  const sumVisH = sumVis.map(x => x.h);
  const sumVisR = sumRows.map(r => sumVis.map(x => r[x.i]));

  // ── PL sheet ──────────────────────────────────────────────
  const wsPL = wb.addWorksheet('PL');

  const sectionCols = Math.max(visH.length, 9);
  const numDests    = new Set(dataRows.map(r => r[dataHeaders.indexOf('Destination')]).filter(v => v != null)).size;

  wsPL.mergeCells(1, 1, 1, sectionCols);
  styledCell(wsPL, 1, 1, { size: 28, bold: true, wrap: true }).value = 'PACKING LIST AND DESTUFFING INSTRUCTION (FBA)';
  wsPL.getRow(1).height = 60;

  // Two label/value pairs per row. Labels span 2 columns, values span the
  // rest of the half so the block lines up with the tables below.
  const half  = Math.ceil(sectionCols / 2);
  const pairs = [
    ['Client Name',     cName,     'Container #',       containerName],
    ['File #',          fileNo || '', '# of Destinations', numDests],
  ];
  pairs.forEach(([l1, v1, l2, v2], i) => {
    const r = 2 + i;
    wsPL.mergeCells(r, 1, r, 2);
    wsPL.mergeCells(r, 3, r, half);
    wsPL.mergeCells(r, half + 1, r, half + 2);
    wsPL.mergeCells(r, half + 3, r, sectionCols);
    styledCell(wsPL, r, 1,        { size: 18, bold: true }).value = l1;
    styledCell(wsPL, r, 3,        { size: 18, bold: true }).value = v1;
    styledCell(wsPL, r, half + 1, { size: 18, bold: true }).value = l2;
    styledCell(wsPL, r, half + 3, { size: 18, bold: true }).value = v2;
    wsPL.getRow(r).height = 55;
  });

  const BLANK_ROW = 2 + pairs.length;
  wsPL.getRow(BLANK_ROW).height = 20; // blank separator

  // Section title row. The word "LOADS" in column A is the marker code uses
  // to find the loads table (see parsePLSections in utils.js), so a hand-made
  // PL only needs this word typed above its header row to be readable.
  const LOADS_TITLE_ROW = BLANK_ROW + 1;
  wsPL.mergeCells(LOADS_TITLE_ROW, 1, LOADS_TITLE_ROW, sectionCols);
  styledCell(wsPL, LOADS_TITLE_ROW, 1, { size: 24, bold: true }).value = PL_SECTION_LOADS;
  wsPL.getRow(LOADS_TITLE_ROW).height = 50;

  const DATA_HDR_ROW = LOADS_TITLE_ROW + 1;
  visH.forEach((h, ci) => { styledCell(wsPL, DATA_HDR_ROW, ci + 1, { bold: true, wrap: true }).value = h; });
  wsPL.getRow(DATA_HDR_ROW).height = 45;

  const DATA_START = DATA_HDR_ROW + 1;
  visRows.forEach((row, ri) => {
    row.forEach((v, ci) => { styledCell(wsPL, DATA_START + ri, ci + 1).value = tidyNumber(v) ?? ''; });
    wsPL.getRow(DATA_START + ri).height = 30;
  });
  const lastDataRow = DATA_START + visRows.length - 1;

  // Destination groups: consecutive rows sharing the same Destination.
  // Each group is one "part" of the packing list.
  const destVisIdx = visH.indexOf('Destination') + 1;
  const groups = [];   // [{ s, e }] as 1-based sheet row numbers
  if (destVisIdx > 0 && visRows.length) {
    let s = DATA_START;
    while (s <= lastDataRow) {
      let e = s;
      const val = visRows[s - DATA_START][destVisIdx - 1];
      while (e + 1 <= lastDataRow && visRows[e + 1 - DATA_START][destVisIdx - 1] === val) e++;
      groups.push({ s, e });
      s = e + 1;
    }
  }

  // Merge consecutive identical Destination cells
  for (const { s, e } of groups) {
    if (e > s) wsPL.mergeCells(s, destVisIdx, e, destVisIdx);
  }

  // Split line between parts: a thick line under the last row of every
  // destination group. Excel draws the shared edge from either cell, so set
  // the bottom edge of the group end and the top edge of the next group.
  // A merged Destination cell shares one style across its range, so setting
  // its top/bottom applies to the top/bottom of the whole merged block.
  const thick = { style: 'medium' };
  for (let gi = 0; gi < groups.length - 1; gi++) {
    const end  = groups[gi].e;
    const next = groups[gi + 1].s;
    for (let c = 1; c <= visH.length; c++) {
      const a = wsPL.getCell(end,  c); a.border = { ...a.border, bottom: thick };
      const b = wsPL.getCell(next, c); b.border = { ...b.border, top:    thick };
    }
  }

  // Embedded summary section
  const SUM_TITLE_ROW  = lastDataRow + 2;
  const SUM_HDR_ROW    = SUM_TITLE_ROW + 1;
  const SUM_DATA_START = SUM_HDR_ROW + 1;

  // The summary is its own table of nine fixed columns, A to I. The loads
  // above it can be wider — every column carried through from the client's
  // sheet adds one — so the summary is measured against itself, not against
  // them, or its heading overhangs the table it belongs to.
  const SUM_HEADERS = PL_SUMMARY_HEADERS;
  const sumCols     = SUM_HEADERS.length;

  wsPL.mergeCells(SUM_TITLE_ROW, 1, SUM_TITLE_ROW, sumCols);
  styledCell(wsPL, SUM_TITLE_ROW, 1, { size: 24, bold: true }).value = PL_SECTION_SUMMARY;
  wsPL.getRow(SUM_TITLE_ROW).height = 50;
  SUM_HEADERS.forEach((lbl, ci) => {
    styledCell(wsPL, SUM_HDR_ROW, ci + 1, { size: 14, bold: true, wrap: true }).value = lbl;
  });
  wsPL.getRow(SUM_HDR_ROW).height = 95;

  sumVisR.forEach((row, ri) => {
    const rn = SUM_DATA_START + ri;
    for (let ci = 0; ci < sumCols; ci++) {
      styledCell(wsPL, rn, ci + 1, { size: 16 }).value = ci < row.length ? (tidyNumber(row[ci]) ?? '') : '';
    }
    wsPL.getRow(rn).height = 85; // roomy: actual skid counts are written in by hand
  });

  // Section split lines: a thick full-width line between the container
  // detail block, the loads table and the summary. Each line is drawn on the
  // shared edge (bottom of the blank row / top of the next section header)
  // so Excel shows it regardless of which side it picks.
  const heavy = { style: 'thick' };
  function sectionLine(blankRow, nextRow, cols = sectionCols) {
    for (let c = 1; c <= cols; c++) {
      const a = wsPL.getCell(blankRow, c); a.border = { ...a.border, bottom: heavy };
      const b = wsPL.getCell(nextRow,  c); b.border = { ...b.border, top:    heavy };
    }
  }
  sectionLine(LOADS_TITLE_ROW - 1, LOADS_TITLE_ROW); // container detail | loads
  sectionLine(SUM_TITLE_ROW - 1,   SUM_TITLE_ROW, Math.max(sectionCols, sumCols)); // loads | summary

  // Auto-fit column widths from the table cells only (titles and the
  // container-detail block are merged across columns and must not count).
  // Width units are ~one character of the 11pt default font, so scale by
  // font size and count CJK / full-width characters as two.
  const plLastRow = SUM_DATA_START + sumVisR.length - 1;
  const CJK       = /[\u1100-\u115F\u2E80-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE30-\uFE4F\uFF00-\uFF60\uFFE0-\uFFE6]/;
  const dispLen   = str => [...str].reduce((n, ch) => n + (CJK.test(ch) ? 2 : 1), 0);
  // Longest run Excel cannot break: whitespace and CJK characters are break points.
  const longestRun = str => Math.max(0, ...str.split(new RegExp(`[\\s${CJK.source.slice(1, -1)}]+`)).map(dispLen));
  const measure   = (r, c) => {
    const cl = wsPL.getCell(r, c);
    if (cl.value == null || cl.value === '') return 0;
    const lines = String(cl.value).split('\n');
    const size  = (cl.font && cl.font.size) || 11;
    const wrap  = !!(cl.alignment && cl.alignment.wrapText);
    // Wrapped cells may break onto ~2 lines, but never inside a word.
    const need  = wrap
      ? Math.max(...lines.map(l => Math.max(dispLen(l) / 2, longestRun(l))))
      : Math.max(...lines.map(dispLen));
    // The trailing margin is what keeps a value off the cell border. At 2 a
    // long FBA id ran edge to edge and the sheet read as a solid block.
    return need * (size / 11) * 1.1 + 4.5;
  };
  const measureRows = [];
  for (let r = DATA_HDR_ROW; r <= lastDataRow; r++) measureRows.push(r);
  for (let r = SUM_HDR_ROW;  r <= plLastRow;   r++) measureRows.push(r);
  for (let c = 1; c <= sectionCols; c++) {
    let w = 12;
    for (const r of measureRows) w = Math.max(w, measure(r, c));
    wsPL.getColumn(c).width = Math.min(w, 50);
  }

  // Print setup: landscape, one page wide, repeat nothing.
  wsPL.pageSetup = {
    orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0,
    margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
    horizontalCentered: true,
  };

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
  const name   = fileName || `PL(${containerName}).xlsx`;
  const a      = Object.assign(document.createElement('a'), { href: url, download: name });
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
