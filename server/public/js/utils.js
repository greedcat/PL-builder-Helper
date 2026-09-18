// ─────────────────────────────────────────────────────────────
// TEXT UTILITIES
// ─────────────────────────────────────────────────────────────
function isEmpty(v) {
  return v === null || v === undefined || v === '';
}

function normalizeText(text) {
  if (text == null) return '';
  return String(text).toUpperCase().trim()
    .replace(/\s+/g, '').replace(/-/g, '/').replace(/[^\w一-鿿/]/g, '');
}

function normCBM(s) {
  if (s == null) return '';
  return String(s).normalize('NFKC').toLowerCase().trim()
    .replace(/\s+/g, '').replace(/（/g, '(').replace(/）/g, ')').replace(/[^a-z0-9一-鿿()]+/g, '');
}

function normWeight(s) {
  if (s == null) return '';
  return String(s).normalize('NFKC').toLowerCase().replace(/\s+/g, '').trim();
}

// Drops the internal _Pallet column and any column that is null in every row.
// Excel stores a merged block as one value in its top-left cell and nothing
// in the rest, so a destination merged down five rows reads as one value and
// four blanks. Copy the value down the block once, when the sheet is read.
// Only merges that span rows are filled: a merged banner across the top of a
// sheet is a title, not data, and filling it would make that row look like a
// full header row.
function fillMergedRows(ws, rows) {
  const merges = ws && ws['!merges'];
  if (!merges || !merges.length) return rows;
  for (const m of merges) {
    if (m.e.r <= m.s.r) continue;                       // single row: a banner
    const src = rows[m.s.r] ? rows[m.s.r][m.s.c] : undefined;
    if (isEmpty(src)) continue;
    for (let r = m.s.r; r <= m.e.r; r++) {
      if (!rows[r]) rows[r] = [];
      for (let c = m.s.c; c <= m.e.c; c++) {
        if (r === m.s.r && c === m.s.c) continue;
        if (isEmpty(rows[r][c])) rows[r][c] = src;
      }
    }
  }
  return rows;
}

function getVisibleColumns(headers, rows) {
  const cols = headers
    .map((h, ci) => ({ h, ci }))
    .filter(({ h, ci }) => h !== '_Pallet' && h !== '_Row' && rows.some(r => r[ci] != null));
  return {
    headers: cols.map(x => x.h),
    rows:    rows.map(r => cols.map(x => r[x.ci])),
  };
}

// ─────────────────────────────────────────────────────────────
// PL SECTIONS
// ─────────────────────────────────────────────────────────────
// A PL sheet has three parts, separated by visible title rows that a person
// can also type by hand:
//   rows 1..n      container detail (title, client, container #, ...)
//   "LOADS"        title row, then the loads header row, then load rows
//   "SUMMARY"      title row, then the summary header row, then summary rows
// A section ends at the first fully blank row or at the next title row.
//
// parsePLSections(rows) reads a 2D row array such as
//   XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null })
// and returns 0-based row indices (null where a part is missing):
//   {
//     containerDetail: { start, end },
//     loads:           { title, header, start, end },
//     summary:         { title, header, start, end }
//   }
// If no "LOADS" title exists (older or hand-made sheets), the loads header is
// taken as the first row that contains a "Destination" cell. Returns null when
// no loads table can be found at all.
const PL_SECTION_LOADS   = 'LOADS';
const PL_SECTION_SUMMARY = 'SUMMARY';

function parsePLSections(rows) {
  const norm    = v => (v == null ? '' : String(v).trim().toUpperCase());
  const isBlank = row => !row || row.every(v => norm(v) === '');
  const rowHas  = (row, t) => !!row && row.some(v => norm(v) === t);
  const findRow = (t, from = 0, to = rows.length) => {
    for (let i = from; i < to; i++) if (rowHas(rows[i], t)) return i;
    return -1;
  };
  // Data rows run from `from` until a blank row or a row index in `stops`.
  const dataEnd = (from, stops) => {
    let end = from - 1;
    for (let i = from; i < rows.length; i++) {
      if (isBlank(rows[i]) || stops.includes(i)) break;
      end = i;
    }
    return end;
  };

  const loadsTitle   = findRow(PL_SECTION_LOADS);
  const summaryTitle = findRow(PL_SECTION_SUMMARY, loadsTitle + 1);

  let loadsHeader = loadsTitle >= 0 ? loadsTitle + 1 : -1;
  if (loadsHeader < 0) {
    // Fallback: first row with a "Destination" cell, before any SUMMARY title.
    loadsHeader = findRow('DESTINATION', 0, summaryTitle >= 0 ? summaryTitle : rows.length);
  }
  if (loadsHeader < 0 || loadsHeader >= rows.length) return null;

  const loadsStart = loadsHeader + 1;
  const loadsEnd   = dataEnd(loadsStart, [summaryTitle]);

  // Container detail: everything above the loads title (or header), minus
  // trailing blank rows.
  let cdEnd = (loadsTitle >= 0 ? loadsTitle : loadsHeader) - 1;
  while (cdEnd >= 0 && isBlank(rows[cdEnd])) cdEnd--;

  let summary = { title: null, header: null, start: null, end: null };
  if (summaryTitle >= 0 && summaryTitle + 1 < rows.length) {
    const h = summaryTitle + 1;
    summary = { title: summaryTitle, header: h, start: h + 1, end: dataEnd(h + 1, []) };
  }

  return {
    containerDetail: cdEnd >= 0 ? { start: 0, end: cdEnd } : null,
    loads:           { title: loadsTitle >= 0 ? loadsTitle : null, header: loadsHeader, start: loadsStart, end: loadsEnd },
    summary
  };
}

// ─────────────────────────────────────────────────────────────
// PRIVATE ADDRESSES
// A delivery to a private address arrives as the whole consignee block:
// name, phone, province, postcode and the street address all in one cell.
// Grouping and reading a packing list by that is hopeless, so reduce it to
// the street: "PA - 3600 Ridgeway Dr".
// ─────────────────────────────────────────────────────────────
const PL_PRIVATE_PREFIX = 'PA - ';

const STREET_TYPES = {
  rd: 'Rd', road: 'Rd', st: 'St', street: 'St', ave: 'Ave', avenue: 'Ave',
  dr: 'Dr', drive: 'Dr', blvd: 'Blvd', boulevard: 'Blvd', way: 'Way',
  cres: 'Cres', crescent: 'Cres', crt: 'Crt', ct: 'Crt', court: 'Crt',
  pl: 'Pl', place: 'Pl', ln: 'Lane', lane: 'Lane', hwy: 'Hwy', highway: 'Hwy',
  pkwy: 'Pkwy', parkway: 'Pkwy', terr: 'Terr', terrace: 'Terr',
  trl: 'Trail', trail: 'Trail', cir: 'Cir', circle: 'Cir', sq: 'Sq', square: 'Sq',
  gate: 'Gate', gdns: 'Gdns', gardens: 'Gdns', close: 'Close', row: 'Row',
};

// Title-cases a word but leaves things like "3600" and "McLeod" alone.
function titleWord(w) {
  if (/\d/.test(w)) return w;
  if (/[a-z][A-Z]/.test(w)) return w;                  // already mixed case
  return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
}

// Returns "PA - 3600 Ridgeway Dr", or null when no street address is present.
function simplifyPrivateAddress(text) {
  if (text == null) return null;
  const s = String(text).replace(/\s+/g, ' ').trim();
  if (!s) return null;

  const types = Object.keys(STREET_TYPES).sort((a, b) => b.length - a.length).join('|');
  // A street number, up to three name words, then a street type.
  const re = new RegExp(
    `(\\d+[A-Za-z]?)\\s+([A-Za-z][A-Za-z'\\-\\.]*(?:\\s+[A-Za-z][A-Za-z'\\-\\.]*){0,2})\\s+(${types})\\b\\.?`,
    'i');
  const m = s.match(re);
  if (!m) return null;

  const name = m[2].split(/\s+/).map(titleWord).join(' ');
  const type = STREET_TYPES[m[3].toLowerCase()] || titleWord(m[3]);

  // "St Clair Ave West" and "St Clair Ave East" are different streets, so a
  // direction immediately after the street type belongs in the short name.
  const DIRS = { n: 'N', north: 'N', s: 'S', south: 'S', e: 'E', east: 'E',
                 w: 'W', west: 'W', ne: 'NE', nw: 'NW', se: 'SE', sw: 'SW' };
  // Only when the direction ends the street name. "North York" is a city.
  const tail = s.slice(m.index + m[0].length).match(/^[\s,]*([A-Za-z]{1,5})\s*(?=,|$)/);
  const dir  = tail && DIRS[tail[1].toLowerCase()] ? ' ' + DIRS[tail[1].toLowerCase()] : '';

  return `${PL_PRIVATE_PREFIX}${m[1]} ${name} ${type}${dir}`;
}

// ISO container numbers: four letters then seven digits, e.g. MSMU7284163.
function findContainerNumber(rows) {
  for (const row of rows || []) {
    for (const v of row || []) {
      if (v == null) continue;
      const m = String(v).toUpperCase().match(/\b([A-Z]{4}\d{7})\b/);
      if (m) return m[1];
    }
  }
  return null;
}

// ─────────────────────────────────────────────────────────────
// The summary block's column headings. Shared so the preview shows the
// same nine columns the workbook is written with, rather than the
// internal names the pipeline uses.
// ─────────────────────────────────────────────────────────────
const PL_SUMMARY_HEADERS = [
  'Destination\n地址',
  'Total Carton\n箱数',
  'Estimated Skid #\n预计托盘数量',
  'Actual Skid # (Standard)\n实际打托数量（标准）',
  'Actual Skid # (60 Inches)\n实际打托数量（长板60寸）',
  'Actual Skid # (75 Inches)\n实际打托数量（长板75寸）',
  'Actual Skid # (96 Inches)\n实际打托数量（长板96寸）',
  'Standard plt but overhanging',
  'Self Palletized\n自托',
];

// Excel stores 2.27 as 2.2699999809265137. Reading it back at full
// precision put that noise on the page and into the file, so trim the
// binary artefact without touching a genuine long decimal.
function tidyNumber(v) {
  if (typeof v !== 'number' || !Number.isFinite(v)) return v;
  const r = Number(v.toFixed(6));
  return Object.is(r, -0) ? 0 : r;
}
