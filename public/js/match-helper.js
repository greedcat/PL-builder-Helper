// Match Helper — groups pasted Dest+Value rows and lays them out in the
// shared destination order. Uses js/shared/normalize.js and clipboard.js.

// ── order list: same storage key and API as the Container Match Helper ─────
const ORDER_STORAGE_KEY = "containerMatchHelper.destOrderList";
const orderInput = document.getElementById("order-input");
const savedOrder = localStorage.getItem(ORDER_STORAGE_KEY);
orderInput.value = (savedOrder && savedOrder.trim()) ? savedOrder : DEFAULT_DEST_ORDER.join("\n");

// the shared list from the server overrides the local copy when available
fetch("/api/match-helper/dest-order")
  .then(r => r.ok ? r.json() : {})
  .then(cfg => {
    if (Array.isArray(cfg.destOrder) && cfg.destOrder.length) {
      orderInput.value = cfg.destOrder.join("\n");
    }
  })
  .catch(() => {}); // offline / server unavailable — keep local copy

orderInput.addEventListener("input", () => {
  localStorage.setItem(ORDER_STORAGE_KEY, orderInput.value);
});

// ── parse TSV pasted from Excel ───────────────────────────────────────────
function parseTSV(text) {
  const lines = text.trim().split(/\r?\n/).filter(l => l.trim());
  if (lines.length < 2) return null;

  const header = lines[0].split("\t").map(h => h.trim().toLowerCase());
  const destIdx  = header.findIndex(h => h === "dest" || h === "destination" || h === "目的地");
  const valueIdx = header.findIndex(h => h === "value" || h === "值" || h === "金额" || h === "amount");

  if (destIdx === -1 || valueIdx === -1) {
    // no recognised header — treat every line as data
    return lines.map(l => {
      const cols = l.split("\t");
      return { dest: normDest(cols[0]), value: parseFloat(cols[1]) || 0 };
    }).filter(r => r.dest);
  }

  return lines.slice(1).map(l => {
    const cols = l.split("\t");
    return { dest: normDest(cols[destIdx]), value: parseFloat(cols[valueIdx]) || 0 };
  }).filter(r => r.dest);
}

// ── group by dest, folding YYC/YEG/YVR into their combined columns ─────────
function groupByOrder(rows, order) {
  const orderUpper = order.map(o => o.toUpperCase());
  const grouped = {};   // order item -> total
  const extra = new Set();
  for (const { dest, value } of rows) {
    const idx = matchDestIndex(dest, orderUpper);
    if (idx === -1) { extra.add(dest); continue; }
    const key = order[idx];
    grouped[key] = (grouped[key] || 0) + value;
  }
  return { grouped, extra: [...extra] };
}

// ── render ────────────────────────────────────────────────────────────────
let _horizData = []; // store for copy

function run() {
  const excelRaw = document.getElementById("excel-input").value.trim();
  const orderRaw = orderInput.value.trim();

  const alertsEl  = document.getElementById("alerts");
  const summaryEl = document.getElementById("summary");
  const section   = document.getElementById("results-section");
  alertsEl.innerHTML = "";
  summaryEl.innerHTML = "";

  if (!excelRaw) { alert("Please paste your Excel data first."); return; }

  const rows = parseTSV(excelRaw);
  if (!rows || rows.length === 0) {
    alertsEl.innerHTML = `<div class="alert alert-error">Could not parse data. Make sure you copy at least two columns (Dest + Value) from Excel, including the header row.</div>`;
    section.classList.add("visible");
    return;
  }

  const order = orderRaw.split(/\r?\n/)
    .map(normOrderItem)
    .filter(Boolean);

  const { grouped, extra } = groupByOrder(rows, order);
  const noData = order.filter(d => !(d in grouped));

  // build result array
  const result = order.map(d => ({ dest: d, value: grouped[d] ?? null }));
  _horizData = result;

  // horizontal table
  const headers = result.map(r => `<th>${escapeHtml(r.dest)}</th>`).join("");
  const cells   = result.map(r => {
    if (r.value !== null && r.value !== 0)
      return `<td class="val">${r.value.toLocaleString()}</td>`;
    return `<td class="empty">—</td>`;
  }).join("");
  const htHtml = `<table><thead><tr>${headers}</tr></thead><tbody><tr>${cells}</tr></tbody></table>`;
  document.getElementById("horiz-table").innerHTML = htHtml;

  // alerts
  let alertHtml = "";
  if (extra.length)
    alertHtml += `<div class="alert alert-error">⚠️ <strong>${extra.length} destination(s) in Excel but NOT in order list:</strong> ${extra.map(escapeHtml).join(", ")}</div>`;
  if (noData.length)
    alertHtml += `<div class="alert alert-warn">ℹ️ <strong>${noData.length} destination(s) in order list with no data:</strong> ${noData.map(escapeHtml).join(", ")}</div>`;
  if (!extra.length && !noData.length)
    alertHtml += `<div class="alert alert-success">✅ All destinations matched perfectly.</div>`;
  alertsEl.innerHTML = alertHtml;

  // summary chips
  const matched = result.filter(r => r.value !== null && r.value !== 0).length;
  const total   = result.reduce((s, r) => s + (r.value || 0), 0);
  summaryEl.innerHTML = `
    <span class="chip chip-blue">${result.length} destinations</span>
    <span class="chip chip-green">${matched} with data</span>
    <span class="chip chip-yellow">${noData.length} empty</span>
    ${extra.length ? `<span class="chip chip-red">${extra.length} unmatched</span>` : ""}
    <span class="chip chip-blue">Total: ${total.toLocaleString()}</span>
  `;

  section.classList.add("visible");
  section.scrollIntoView({ behavior: "smooth", block: "start" });

  copyHoriz();
}

function copyHoriz() {
  if (!_horizData.length) return;
  const values = _horizData.map(r => (r.value !== null && r.value !== 0) ? r.value : "").join("\t");
  copyToClipboard(values).then(() => {
    const btn = document.getElementById("copy-horiz-btn");
    btn.textContent = "Copied!";
    btn.classList.add("copied");
    setTimeout(() => { btn.textContent = "Copy"; btn.classList.remove("copied"); }, 2000);
  });
}

function clearAll() {
  document.getElementById("excel-input").value = "";
  document.getElementById("results-section").classList.remove("visible");
}

// allow Ctrl+Enter to run
document.addEventListener("keydown", e => {
  if ((e.ctrlKey || e.metaKey) && e.key === "Enter") run();
});
