// ─────────────────────────────────────────────────────────────
// CONFIGURATION
// These are fallbacks only. On load, initConfigFromDB() replaces them with
// the shared lists from MongoDB; edit those in the settings drawer.
// ─────────────────────────────────────────────────────────────
let DEST_LIST = [
  'YYZ3','YYZ4','YOO1','XYY1','YHM1','YYZ9','YYZ7','YXU1',
  'YOW1','YOW3','YGK1','私人地址','YYZ1','YYC4/6','YEG1/2','YVR3/4'
];
let CARTON_KEYWORDS = [
  'ctn','ctns','carton','cartons','no of carton','total ctn',
  '箱数','箱','数量','件数','pkgs','发货数量','箱数(ctns)','箱数(paks)'
];
let CBM_KEYWORDS = [
  'cbm','cmb','volume','体积','总体积','totalvolume','方数',
  '入库体积','体积(cbm)','total volumn(cbm)','totalvolumecbm','totalvolume(cbm)','VOL'
];
let WEIGHT_KEYWORDS = [
  'weight','kgs','kg','重量','重量kg','g.w.(毛重)','实重','实际重量','毛重（kgs)'
];
let NO_NEED_COL = ['序号','战略客户标记','派送邮编','柜号','是否FBA','派送方式','渠道名'];

const CFG_SECTIONS = [
  { key: 'DEST_LIST',       label: 'DEST_LIST',       color: '#dbeafe', text: '#1e40af' },
  { key: 'CARTON_KEYWORDS', label: 'CARTON_KEYWORDS', color: '#dcfce7', text: '#166534' },
  { key: 'CBM_KEYWORDS',    label: 'CBM_KEYWORDS',    color: '#fef9c3', text: '#854d0e' },
  { key: 'WEIGHT_KEYWORDS', label: 'WEIGHT_KEYWORDS', color: '#fce7f3', text: '#9d174d' },
  { key: 'NO_NEED_COL',     label: 'NO_NEED_COL',     color: '#f1f5f9', text: '#475569' },
];

function getCfgList(key) {
  switch (key) {
    case 'DEST_LIST':       return DEST_LIST;
    case 'CARTON_KEYWORDS': return CARTON_KEYWORDS;
    case 'CBM_KEYWORDS':    return CBM_KEYWORDS;
    case 'WEIGHT_KEYWORDS': return WEIGHT_KEYWORDS;
    case 'NO_NEED_COL':     return NO_NEED_COL;
    default: return null;
  }
}

function setCfgList(key, values) {
  switch (key) {
    case 'DEST_LIST':       DEST_LIST       = values; break;
    case 'CARTON_KEYWORDS': CARTON_KEYWORDS = values; break;
    case 'CBM_KEYWORDS':    CBM_KEYWORDS    = values; break;
    case 'WEIGHT_KEYWORDS': WEIGHT_KEYWORDS = values; break;
    case 'NO_NEED_COL':     NO_NEED_COL     = values; break;
  }
}

// The keyword lists and the backend URL live in a drawer that starts closed;
// this survives the re-render that follows every edit.
let cfgSettingsOpen = false;

// Holds the most recently removed item so it can be restored via "Undo".
let cfgLastRemoved = null;
let cfgUndoTimer   = null;

function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]));
}

// ─────────────────────────────────────────────────────────────
// CONFIG PANEL  (view + edit + database sync)
// ─────────────────────────────────────────────────────────────
function renderConfigPanel() {
  const panel = document.getElementById('cfgPanel');
  if (!panel) return;

  panel.innerHTML = `
    <div class="cfg-db-bar">
      <button type="button" id="cfgDbToggle" class="cfg-link-btn" aria-expanded="${cfgSettingsOpen}"
              aria-controls="cfgSettingsBody">
        <span class="cfg-caret">${cfgSettingsOpen ? '▾' : '▸'}</span> ⚙ Settings &amp; keyword lists
      </button>
      <span id="cfgDbStatus" class="cfg-db-status">⚪ not checked</span>
      <button type="button" id="cfgDbReload" class="cfg-link-btn">⟳ Reload from DB</button>
    </div>
    <div id="cfgSettingsBody" class="cfg-settings-body" style="display:${cfgSettingsOpen ? 'block' : 'none'};">
      ${CFG_SECTIONS.map(renderSection).join('')}
    </div>
    ${cfgLastRemoved ? `
      <div id="cfgUndoToast" class="cfg-undo-toast">
        Removed "${escapeHtml(cfgLastRemoved.value)}" from ${cfgLastRemoved.key}.
        <button type="button" id="cfgUndoBtn" class="cfg-link-btn">Undo</button>
      </div>
    ` : ''}
  `;

  attachConfigPanelHandlers();
  panel.style.display = 'block';
}

function renderSection(s) {
  const values = getCfgList(s.key) || [];
  return `
    <div class="cfg-section" data-key="${s.key}">
      <div class="cfg-label">${s.label}</div>
      <div class="cfg-chips">
        ${values.map((v, i) => `
          <span class="cfg-chip" style="background:${s.color};color:${s.text};">
            ${escapeHtml(v)}
            <button type="button" class="cfg-chip-x" data-key="${s.key}" data-idx="${i}" title="Remove">×</button>
          </span>
        `).join('')}
      </div>
      <div class="cfg-add-row">
        <input type="text" class="cfg-add-input" data-key="${s.key}" placeholder="Add value…">
        <button type="button" class="cfg-add-btn" data-key="${s.key}">+ Add</button>
      </div>
    </div>
  `;
}

function attachConfigPanelHandlers() {
  const panel = document.getElementById('cfgPanel');

  const toggleBtn = document.getElementById('cfgDbToggle');
  const body      = document.getElementById('cfgSettingsBody');
  toggleBtn.addEventListener('click', () => {
    cfgSettingsOpen     = !cfgSettingsOpen;
    body.style.display  = cfgSettingsOpen ? 'block' : 'none';
    toggleBtn.setAttribute('aria-expanded', String(cfgSettingsOpen));
    toggleBtn.querySelector('.cfg-caret').textContent = cfgSettingsOpen ? '▾' : '▸';
  });

  // It now hangs off the corner as a dropdown, so dismiss it like one.
  if (!attachConfigPanelHandlers.dismissBound) {
    attachConfigPanelHandlers.dismissBound = true;
    document.addEventListener('mousedown', e => {
      if (!cfgSettingsOpen) return;
      if (e.target.closest('#cfgPanel')) return;
      cfgSettingsOpen = false;
      renderConfigPanel();
    });
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape' && cfgSettingsOpen) { cfgSettingsOpen = false; renderConfigPanel(); }
    });
  }

  const reloadBtn = document.getElementById('cfgDbReload');
  reloadBtn.addEventListener('click', () => initConfigFromDB());

  // Remove a value from a list (with confirmation + undo)
  panel.querySelectorAll('.cfg-chip-x').forEach(btn => {
    btn.addEventListener('click', () => {
      const key    = btn.dataset.key;
      const idx    = Number(btn.dataset.idx);
      const values = getCfgList(key).slice();
      const value  = values[idx];
      if (!confirm(`Remove "${value}" from ${key}?`)) return;
      values.splice(idx, 1);
      setCfgList(key, values);

      clearTimeout(cfgUndoTimer);
      cfgLastRemoved = { key, value, idx };
      cfgUndoTimer = setTimeout(() => {
        cfgLastRemoved = null;
        renderConfigPanel();
      }, 8000);

      renderConfigPanel();
      persistConfigList(key, values);
    });
  });

  // Undo the most recent removal
  const undoBtn = document.getElementById('cfgUndoBtn');
  if (undoBtn) {
    undoBtn.addEventListener('click', () => {
      const { key, value, idx } = cfgLastRemoved;
      clearTimeout(cfgUndoTimer);
      cfgLastRemoved = null;

      const values = getCfgList(key).slice();
      values.splice(Math.min(idx, values.length), 0, value);
      setCfgList(key, values);

      renderConfigPanel();
      persistConfigList(key, values);
    });
  }

  // Add a value to a list
  panel.querySelectorAll('.cfg-add-btn').forEach(btn => {
    const addValue = () => {
      const key   = btn.dataset.key;
      const input = panel.querySelector(`.cfg-add-input[data-key="${key}"]`);
      const val   = input.value.trim();
      if (!val) return;
      const values = getCfgList(key).slice();
      if (values.includes(val)) return;
      values.push(val);
      setCfgList(key, values);
      renderConfigPanel();
      persistConfigList(key, values);
    };
    btn.addEventListener('click', addValue);
  });
  panel.querySelectorAll('.cfg-add-input').forEach(input => {
    input.addEventListener('keydown', e => {
      if (e.key === 'Enter') {
        e.preventDefault();
        panel.querySelector(`.cfg-add-btn[data-key="${input.dataset.key}"]`).click();
      }
    });
  });
}

async function persistConfigList(key, values) {
  if (!isDbConfigured()) return;
  const status = document.getElementById('cfgDbStatus');
  try {
    if (status) status.textContent = '🟡 saving…';
    await saveConfigToDB({ [key]: values });
    if (status) status.textContent = '🟢 connected';
  } catch (err) {
    console.error(err);
    if (status) status.textContent = '🔴 save failed';
  }
}

// Loads the shared config document from MongoDB (if configured) and
// overrides the in-memory lists, then re-renders the panel.
async function initConfigFromDB() {
  if (!isDbConfigured()) {
    renderConfigPanel();
    return;
  }
  const status = document.getElementById('cfgDbStatus');
  try {
    if (status) status.textContent = '🟡 loading…';
    const doc = await loadConfigFromDB();
    if (doc) {
      for (const { key } of CFG_SECTIONS) {
        if (Array.isArray(doc[key])) setCfgList(key, doc[key]);
      }
    }
    renderConfigPanel();
    const refreshed = document.getElementById('cfgDbStatus');
    if (refreshed) refreshed.textContent = '🟢 connected';
  } catch (err) {
    console.error(err);
    renderConfigPanel();
    const refreshed = document.getElementById('cfgDbStatus');
    if (refreshed) refreshed.textContent = '🔴 connection failed';
  }
}
