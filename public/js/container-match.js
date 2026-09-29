// ── pre-fill destination order list ─────────────────────────────────────────
const ORDER_STORAGE_KEY = "containerMatchHelper.destOrderList";
const orderListEl = document.getElementById("order-list");

let destOrder = (() => {
  const saved = localStorage.getItem(ORDER_STORAGE_KEY);
  const list = saved ? saved.split("\n").map(s => s.trim()).filter(Boolean) : [];
  return list.length ? list : DEFAULT_DEST_ORDER.slice();
})();

function persistOrderLocal() {
  localStorage.setItem(ORDER_STORAGE_KEY, destOrder.join("\n"));
  document.getElementById("save-order-status").textContent = "Saved on this device — 💾 to save online";
}

function renderOrderList() {
  orderListEl.innerHTML = destOrder.map((d, i) => `
    <li class="order-item" draggable="true" data-index="${i}">
      <span class="drag-handle">⠿</span>
      <span class="order-num">${i + 1}</span>
      <span class="dest-name" data-edit="${i}" title="Click to edit">${escapeHtml(d)}</span>
      <button type="button" class="remove-btn" title="Remove" data-remove="${i}">✕</button>
    </li>
  `).join("");
}

function addDest() {
  const input = document.getElementById("new-dest-input");
  const name = input.value.trim();
  if (!name) return;
  if (destOrder.some(d => d.toLowerCase() === name.toLowerCase())) {
    document.getElementById("save-order-status").textContent = `"${name}" is already in the list`;
    return;
  }
  destOrder.push(name);
  input.value = "";
  renderOrderList();
  persistOrderLocal();
  orderListEl.scrollTop = orderListEl.scrollHeight;
}

function resetOrderList() {
  destOrder = DEFAULT_DEST_ORDER.slice();
  renderOrderList();
  persistOrderLocal();
}

// remove item / click-to-edit
orderListEl.addEventListener("click", e => {
  const btn = e.target.closest("[data-remove]");
  if (btn) {
    destOrder.splice(Number(btn.dataset.remove), 1);
    renderOrderList();
    persistOrderLocal();
    return;
  }
  const nameEl = e.target.closest("[data-edit]");
  if (nameEl) startEditDest(nameEl);
});

function startEditDest(nameEl) {
  const idx = Number(nameEl.dataset.edit);
  const item = nameEl.closest(".order-item");
  item.draggable = false; // don't let a text-select drag reorder the row
  const input = document.createElement("input");
  input.type = "text";
  input.className = "dest-edit";
  input.value = destOrder[idx];
  nameEl.replaceWith(input);
  input.focus();
  input.select();

  let done = false;
  const commit = () => {
    if (done) return;
    done = true;
    const name = input.value.trim();
    const isDup = name && destOrder.some((d, i) => i !== idx && d.toLowerCase() === name.toLowerCase());
    if (name && !isDup) {
      destOrder[idx] = name;
      persistOrderLocal();
    } else if (isDup) {
      document.getElementById("save-order-status").textContent = `"${name}" is already in the list`;
    }
    renderOrderList();
  };
  input.addEventListener("blur", commit);
  input.addEventListener("keydown", ev => {
    if (ev.key === "Enter") commit();
    if (ev.key === "Escape") { done = true; renderOrderList(); }
  });
}

// drag & drop reorder
let _dragIndex = null;
orderListEl.addEventListener("dragstart", e => {
  const item = e.target.closest(".order-item");
  if (!item) return;
  _dragIndex = Number(item.dataset.index);
  item.classList.add("dragging");
  e.dataTransfer.effectAllowed = "move";
});
orderListEl.addEventListener("dragover", e => {
  e.preventDefault();
  const item = e.target.closest(".order-item");
  orderListEl.querySelectorAll(".drag-target").forEach(el => el.classList.remove("drag-target"));
  if (item && Number(item.dataset.index) !== _dragIndex) item.classList.add("drag-target");
});
orderListEl.addEventListener("drop", e => {
  e.preventDefault();
  const item = e.target.closest(".order-item");
  if (item && _dragIndex !== null) {
    const to = Number(item.dataset.index);
    if (to !== _dragIndex) {
      const [moved] = destOrder.splice(_dragIndex, 1);
      destOrder.splice(to, 0, moved);
      renderOrderList();
      persistOrderLocal();
    }
  }
});
orderListEl.addEventListener("dragend", () => {
  _dragIndex = null;
  orderListEl.querySelectorAll(".dragging, .drag-target").forEach(el => el.classList.remove("dragging", "drag-target"));
});

// load the shared list from the server (overrides the local copy when available)
fetch("/api/match-helper/dest-order")
  .then(r => r.ok ? r.json() : {})
  .then(cfg => {
    if (Array.isArray(cfg.destOrder) && cfg.destOrder.length) {
      destOrder = cfg.destOrder.slice();
      renderOrderList();
      localStorage.setItem(ORDER_STORAGE_KEY, destOrder.join("\n"));
    }
  })
  .catch(() => {}); // offline / server unavailable — keep local copy

function saveOrderList() {
  const btn    = document.getElementById("save-order-btn");
  const status = document.getElementById("save-order-status");
  btn.disabled = true;
  status.textContent = "Saving…";
  fetch("/api/match-helper/dest-order", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ destOrder })
  })
    .then(r => { if (!r.ok) throw new Error(); return r.json(); })
    .then(() => { status.textContent = "✅ Saved online"; })
    .catch(() => { status.textContent = "⚠️ Save failed — saved on this device only"; })
    .finally(() => { btn.disabled = false; });
}

renderOrderList();

// ── state ──────────────────────────────────────────────────────────────────
let workbook  = null;
let sheetData = [];
let allCols   = [];
let _pivot    = null; // { order, rows } — store for copy

let detectedContainers = [];      // [{ container, client }] in first-seen order
let selectedContainers  = new Set(); // containers to include in output

// ── column auto-detection helpers ───────────────────────────────────────────
const CLIENT_HINTS    = ["client", "client name", "客户", "客户名称"];
const CONTAINER_HINTS = ["container", "container#", "container #", "container no", "container no.", "柜号", "箱号", "集装箱号"];
const DEST_HINTS      = ["destination", "dest", "目的地"];
const VALUE_HINTS     = ["value", "值", "金额", "amount", "qty", "quantity", "数量", "pallet", "板数", "skid"];
const FILE_HINTS      = ["file no", "file#", "file #", "file number", "fileno", "file", "文件号", "档案号", "so no", "job no"];


// ── input mode tabs (upload vs paste) ───────────────────────────────────────
function setInputMode(mode) {
  document.getElementById("tab-upload").classList.toggle("active", mode === "upload");
  document.getElementById("tab-paste").classList.toggle("active", mode === "paste");
  document.getElementById("drop-zone").style.display = mode === "upload" ? "" : "none";
  document.getElementById("paste-zone").style.display = mode === "paste" ? "" : "none";
}

function loadPastedData() {
  const fileStatus = document.getElementById("file-status");
  const raw = document.getElementById("paste-input").value.replace(/\r/g, "");
  const lines = raw.split("\n").filter(l => l.trim() !== "");
  if (lines.length < 2) {
    fileStatus.className = "file-status error";
    fileStatus.textContent = "✖ Paste at least a header row and one data row.";
    return;
  }
  // Excel copies are tab-separated; fall back to commas for hand-typed CSV.
  const delim = lines[0].includes("\t") ? "\t" : ",";
  const headers = lines[0].split(delim).map((h, i) => h.trim() || `Column ${i + 1}`);
  sheetData = lines.slice(1).map(line => {
    const cells = line.split(delim);
    const row = {};
    headers.forEach((h, i) => { row[h] = (cells[i] ?? "").trim(); });
    return row;
  });
  workbook = null;
  allCols = headers;
  document.getElementById("sheet-field").style.display = "none";
  document.getElementById("column-controls").style.display = "flex";
  populateColumnPickers();
  populateContainerList();
  matchBtn.disabled = false;
  fileStatus.className = "file-status ok";
  fileStatus.textContent = `✔ Loaded ${sheetData.length} pasted row(s), ${headers.length} column(s)`;
}

// ── file drop / click ────────────────────────────────────────────────────────
const dropZone   = document.getElementById("drop-zone");
const fileInput  = document.getElementById("file-input");
const fileStatus = document.getElementById("file-status");
const matchBtn   = document.getElementById("match-btn");

dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("dragover", e => { e.preventDefault(); dropZone.classList.add("drag-over"); });
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("drag-over"));
dropZone.addEventListener("drop", e => {
  e.preventDefault(); dropZone.classList.remove("drag-over");
  handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener("change", () => handleFile(fileInput.files[0]));

function handleFile(file) {
  if (!file) return;
  matchBtn.disabled = true;
  fileStatus.className = "file-status";
  fileStatus.textContent = `Loading "${file.name}" …`;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      workbook = XLSX.read(e.target.result, { type: "binary", cellDates: true });
      populateSheets();
      fileStatus.className = "file-status ok";
      fileStatus.textContent = `✔ Loaded "${file.name}"`;
    } catch (err) {
      fileStatus.className = "file-status error";
      fileStatus.textContent = "✖ Failed to read file: " + err.message;
    }
  };
  reader.readAsBinaryString(file);
}

// ── sheet selector ───────────────────────────────────────────────────────────
const sheetField  = document.getElementById("sheet-field");
const sheetSelect = document.getElementById("sheet-select");

function populateSheets() {
  sheetSelect.innerHTML = "";
  workbook.SheetNames.forEach(name => {
    const opt = document.createElement("option");
    opt.value = opt.textContent = name;
    sheetSelect.appendChild(opt);
  });
  sheetField.style.display = workbook.SheetNames.length > 1 ? "flex" : "none";
  document.getElementById("column-controls").style.display = "flex";
  loadSheet(workbook.SheetNames[0]);
}

sheetSelect.addEventListener("change", () => loadSheet(sheetSelect.value));

function loadSheet(name) {
  const ws = workbook.Sheets[name];
  sheetData = XLSX.utils.sheet_to_json(ws, { defval: "" });
  if (!sheetData.length) {
    fileStatus.className = "file-status error";
    fileStatus.textContent = "✖ Sheet is empty.";
    matchBtn.disabled = true;
    return;
  }
  allCols = Object.keys(sheetData[0]);
  populateColumnPickers();
  populateContainerList();
  matchBtn.disabled = false;
}

// ── column pickers ───────────────────────────────────────────────────────────
const clientColSel    = document.getElementById("client-col");
const containerColSel = document.getElementById("container-col");
const destColSel      = document.getElementById("dest-col");
const valueColSel     = document.getElementById("value-col");
const fileColSel      = document.getElementById("file-col");

function populateColumnPickers() {
  [clientColSel, containerColSel, destColSel, valueColSel, fileColSel].forEach(sel => sel.innerHTML = "");
  fileColSel.appendChild(new Option("— none —", ""));
  allCols.forEach(c => {
    clientColSel.appendChild(new Option(c, c));
    containerColSel.appendChild(new Option(c, c));
    destColSel.appendChild(new Option(c, c));
    valueColSel.appendChild(new Option(c, c));
    fileColSel.appendChild(new Option(c, c));
  });

  clientColSel.value    = guessColumn(allCols, CLIENT_HINTS) || allCols[0];
  containerColSel.value = guessColumn(allCols, CONTAINER_HINTS) || allCols[0];
  destColSel.value      = guessColumn(allCols, DEST_HINTS) || allCols[0];
  valueColSel.value      = guessColumn(allCols, VALUE_HINTS) || allCols[allCols.length > 1 ? 1 : 0];
  fileColSel.value      = guessColumn(allCols, FILE_HINTS) || "";
}

containerColSel.addEventListener("change", populateContainerList);
clientColSel.addEventListener("change", populateContainerList);

// ── container filter card ────────────────────────────────────────────────────
const containerSelectCard = document.getElementById("container-select-card");

function populateContainerList() {
  const containerCol = containerColSel.value;
  const clientCol     = clientColSel.value;
  // one entry per client + container pair, so a reused container number shows
  // up once per client instead of being collapsed into a single entry
  const pairs = new Map(); // "client container" -> { container, client }

  for (const raw of sheetData) {
    const container = normContainer(raw[containerCol]);
    if (!container) continue;
    const client = String(raw[clientCol] ?? "").trim();
    const key = client.toUpperCase() + " " + container;
    if (!pairs.has(key)) pairs.set(key, { container, client });
  }

  detectedContainers = [...pairs.values()];
  selectedContainers = new Set(detectedContainers.map(c => c.container));
  containerSelectCard.style.display = detectedContainers.length ? "block" : "none";
  applyContainerFilter(); // re-apply any pasted filter to the fresh data
}

function applyContainerFilter() {
  const statusEl = document.getElementById("container-filter-status");
  const raw = document.getElementById("container-filter-input").value.trim();
  if (!raw) {
    selectedContainers = new Set(detectedContainers.map(c => c.container));
    statusEl.textContent = selectedContainers.size ? `Including all ${selectedContainers.size} containers` : "";
    return;
  }

  const wanted      = raw.split(/\r?\n/).map(normContainer).filter(Boolean);
  const detectedSet = new Set(detectedContainers.map(c => c.container));
  const found        = wanted.filter(c => detectedSet.has(c));
  const notFound      = wanted.filter(c => !detectedSet.has(c));

  selectedContainers = new Set(found);

  statusEl.textContent = notFound.length
    ? `✔ ${found.length} matched · ✖ ${notFound.length} not found: ${notFound.join(", ")}`
    : `✔ ${found.length} matched`;
}

function clearContainerFilter() {
  document.getElementById("container-filter-input").value = "";
  applyContainerFilter();
}

// ── build the client/container × destination pivot ──────────────────────────
function buildPivot(order) {
  const clientCol    = clientColSel.value;
  const containerCol = containerColSel.value;
  const destCol       = destColSel.value;
  const valueCol       = valueColSel.value;

  const orderUpper = order.map(o => o.toUpperCase());
  // A container number can be reused across shipments/clients, so a row is
  // identified by client + container, not container alone.
  const rowsMap   = new Map(); // "client container" -> row
  const rowOrder  = [];
  const unmatchedDest = new Set();
  const fileCol = fileColSel.value;
  // container number -> set of file numbers it appears under (across all rows)
  const containerFiles = new Map();

  for (const raw of sheetData) {
    const container = normContainer(raw[containerCol]);
    if (!container || !selectedContainers.has(container)) continue;

    const client = String(raw[clientCol] ?? "").trim();
    const rowKey = client.toUpperCase() + " " + container;

    if (!rowsMap.has(rowKey)) {
      rowsMap.set(rowKey, { client, container, values: {} });
      rowOrder.push(rowKey);
    }
    const row = rowsMap.get(rowKey);
    if (!row.client && client) row.client = client;

    if (fileCol) {
      const fileNo = normFileNo(raw[fileCol]);
      if (fileNo) {
        if (!containerFiles.has(container)) containerFiles.set(container, new Set());
        containerFiles.get(container).add(fileNo);
      }
    }

    const dest  = normDest(raw[destCol]);
    const value = parseFloat(raw[valueCol]) || 0;
    if (!dest) continue;

    const idx = matchDestIndex(dest, orderUpper);
    if (idx === -1) { unmatchedDest.add(dest); continue; }

    const key = order[idx];
    row.values[key] = (row.values[key] || 0) + value;
  }

  const rows = rowOrder.map(k => rowsMap.get(k));
  // attach the file set so rows whose container spans several files can warn
  for (const r of rows) r.files = containerFiles.get(r.container) || new Set();

  return {
    rows,
    unmatchedDest: [...unmatchedDest]
  };
}

// ── render ────────────────────────────────────────────────────────────────
function run() {
  const alertsEl  = document.getElementById("alerts");
  const summaryEl = document.getElementById("summary");
  const section   = document.getElementById("results-section");
  alertsEl.innerHTML = "";
  summaryEl.innerHTML = "";

  if (!sheetData.length) { alert("Please upload an Excel file first."); return; }

  const order = destOrder.map(normOrderItem).filter(Boolean);
  const { rows, unmatchedDest } = buildPivot(order);
  _pivot = { order, rows };

  // pivot table
  const headerCells = ["CLIENT", "CNTR", ...order].map(h => `<th>${escapeHtml(h)}</th>`).join("");
  const bodyRows = rows.map(r => {
    const multiFile = r.files && r.files.size > 1;
    const cells = [
      `<td>${escapeHtml(r.client)}</td>`,
      `<td>${escapeHtml(r.container)}${multiFile ? ` ⚠️ <small>(${[...r.files].map(escapeHtml).join(", ")})</small>` : ""}</td>`,
      ...order.map(d => {
        const v = r.values[d];
        return v ? `<td class="val">${v.toLocaleString()}</td>` : `<td class="empty"></td>`;
      })
    ].join("");
    return `<tr${multiFile ? ' class="warn-row"' : ""}>${cells}</tr>`;
  }).join("");
  document.getElementById("pivot-table").innerHTML =
    `<table><thead><tr>${headerCells}</tr></thead><tbody>${bodyRows}</tbody></table>`;

  // alerts
  let alertHtml = "";
  const dupFiles = [...new Map(rows.filter(r => r.files && r.files.size > 1).map(r => [r.container, r])).values()];
  if (dupFiles.length)
    alertHtml += `<div class="alert alert-error">⚠️ <strong>${dupFiles.length} container(s) appear under more than one file number:</strong> ${dupFiles.map(r => `${escapeHtml(r.container)} (${[...r.files].map(escapeHtml).join(", ")})`).join(" · ")}</div>`;
  if (unmatchedDest.length)
    alertHtml += `<div class="alert alert-warn">ℹ️ <strong>${unmatchedDest.length} destination value(s) in Excel not in the order list (excluded from table):</strong> ${unmatchedDest.map(escapeHtml).join(", ")}</div>`;
  if (!unmatchedDest.length)
    alertHtml += `<div class="alert alert-success">✅ All destination values matched the order list.</div>`;
  alertsEl.innerHTML = alertHtml;

  // summary chips
  const total = rows.reduce((s, r) => s + Object.values(r.values).reduce((a, b) => a + b, 0), 0);
  const distinctContainers = new Set(rows.map(r => r.container)).size;
  const reused = rows.length - distinctContainers;
  summaryEl.innerHTML = `
    <span class="chip chip-blue">${rows.length} rows · ${distinctContainers} containers</span>
    ${reused ? `<span class="chip chip-blue">${reused} reused container row(s)</span>` : ""}
    <span class="chip chip-blue">${order.length} destination columns</span>
    <span class="chip chip-blue">Total: ${total.toLocaleString()}</span>
  `;

  section.classList.add("visible");
  section.scrollIntoView({ behavior: "smooth", block: "start" });

  copyPivot();
}

function copyPivot() {
  if (!_pivot || !_pivot.rows.length) return;
  const { order, rows } = _pivot;
  const lines = rows.map(r => [r.client, r.container, ...order.map(d => r.values[d] || "")].join("\t"));
  copyToClipboard(lines.join("\r\n")).then(() => {
    const btn = document.getElementById("copy-pivot-btn");
    btn.textContent = "Copied!";
    btn.classList.add("copied");
    setTimeout(() => { btn.textContent = "Copy"; btn.classList.remove("copied"); }, 2000);
  });
}

function clearAll() {
  workbook = null;
  sheetData = [];
  allCols = [];
  detectedContainers = [];
  selectedContainers = new Set();
  fileInput.value = "";
  document.getElementById("paste-input").value = "";
  fileStatus.className = "file-status";
  fileStatus.textContent = "";
  document.getElementById("column-controls").style.display = "none";
  containerSelectCard.style.display = "none";
  document.getElementById("container-filter-input").value = "";
  document.getElementById("container-filter-status").textContent = "";
  matchBtn.disabled = true;
  document.getElementById("results-section").classList.remove("visible");
}

// allow Ctrl+Enter to run
document.addEventListener("keydown", e => {
  if ((e.ctrlKey || e.metaKey) && e.key === "Enter") run();
});
