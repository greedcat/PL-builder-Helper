// ─────────────────────────────────────────────────────────────
// SHARED NORMALIZERS
// Pure functions used by the match tools. One copy, used by
// every page; also loadable in Node for the unit tests.
// ─────────────────────────────────────────────────────────────

// The fallback destination order, shared by Match Helper and Container
// Match Helper. The live list comes from /api/match-helper/dest-order.
const DEFAULT_DEST_ORDER = [
  "YYZ3", "YOO1", "XYY1", "YGK1", "YYZ4", "YYZ9", "YYZ7", "YXU1",
  "YHM1", "YOW1", "YOW3", "私人地址", "HOLD", "YYZ1", "YYC4/6",
  "YEG1/2", "YVR3/4", "自提", "换标", "快递"
];

// Container cells sometimes carry a prefix before the real container number,
// e.g. "DS-TIIU8097334" — pull out the standard ISO 6346 code (4 letters + 7 digits).
function normContainer(s) {
  if (s == null) return "";
  const str = String(s).trim().toUpperCase();
  const match = str.match(/[A-Z]{4}\d{7}/);
  return match ? match[0] : str.replace(/[^A-Z0-9]/g, "");
}

// File numbers like "2067-26033-01" belong to container file "2067-26033" —
// keep the first two dash-separated segments so per-shipment suffixes don't
// count as different files.
function normFileNo(s) {
  if (s == null) return "";
  const str = String(s).trim();
  if (!str) return "";
  const m = str.match(/^([^-\s]+-[^-\s]+)/);
  return m ? m[1] : str;
}

function normDest(s) {
  if (!s) return "";
  const u = String(s).toUpperCase();
  if (u.includes("私仓"))  return "私人地址";
  if (u.includes("HOLD"))  return "HOLD";
  if (u.includes("UPS"))   return "快递";
  const clean = u.replace(/[\s\-]/g, "");
  if (clean.includes("SELFPICKUP")) return "自提";
  return String(s).trim();
}

function normOrderItem(s) {
  if (!s) return "";
  const low = s.toLowerCase().replace(/[\s\-]/g, "");
  if (low.includes("hold"))        return "HOLD";
  if (low.includes("ups"))         return "快递";
  if (low.includes("selfpickup"))  return "自提";
  return s.trim();
}

// Destinations like "YYC6" / "YEG1" / "YVR4" belong to the combined order
// column that shares the prefix ("YYC4/6", "YEG1/2", "YVR3/4"). Returns the
// index into `orderUpper` (already upper-cased), or -1.
function matchDestIndex(dest, orderUpper) {
  const destUpper = dest.toUpperCase();
  let idx = orderUpper.indexOf(destUpper);
  if (idx === -1) {
    const prefixMatch = destUpper.match(/^(YYC|YEG|YVR)/);
    if (prefixMatch) idx = orderUpper.findIndex(o => o.startsWith(prefixMatch[1]));
  }
  return idx;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;" }[c]));
}

function guessColumn(cols, hints) {
  const lower = cols.map(c => String(c).toLowerCase().trim());
  for (const hint of hints) {
    const idx = lower.findIndex(c => c === hint);
    if (idx !== -1) return cols[idx];
  }
  for (const hint of hints) {
    const idx = lower.findIndex(c => c.includes(hint));
    if (idx !== -1) return cols[idx];
  }
  return "";
}

// Node (unit tests) — the browser just uses the globals.
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    DEFAULT_DEST_ORDER, normContainer, normFileNo, normDest,
    normOrderItem, matchDestIndex, escapeHtml, guessColumn
  };
}
