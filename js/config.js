// ─────────────────────────────────────────────────────────────
// CONFIGURATION  (overridden at runtime by config.properties)
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

function applyConfig(text) {
  const props = {};
  for (const raw of text.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || line.startsWith('#')) continue;
    const eq = line.indexOf('=');
    if (eq < 0) continue;
    props[line.slice(0, eq).trim()] = line.slice(eq + 1).trim();
  }
  const split = key => (props[key] || '').split(',').map(s => s.trim()).filter(Boolean);
  if (props['DEST_LIST'])        DEST_LIST       = split('DEST_LIST');
  if (props['CARTON_KEYWORDS'])  CARTON_KEYWORDS = split('CARTON_KEYWORDS');
  if (props['CBM_KEYWORDS'])     CBM_KEYWORDS    = split('CBM_KEYWORDS');
  if (props['WEIGHT_KEYWORDS'])  WEIGHT_KEYWORDS = split('WEIGHT_KEYWORDS');
  if (props['NO_NEED_KEYWORDS']) NO_NEED_COL     = split('NO_NEED_KEYWORDS');
  renderConfigPanel();
}

function renderConfigPanel() {
  const sections = [
    { label: 'DEST_LIST',       values: DEST_LIST,       color: '#dbeafe', text: '#1e40af' },
    { label: 'CARTON_KEYWORDS', values: CARTON_KEYWORDS, color: '#dcfce7', text: '#166534' },
    { label: 'CBM_KEYWORDS',    values: CBM_KEYWORDS,    color: '#fef9c3', text: '#854d0e' },
    { label: 'WEIGHT_KEYWORDS', values: WEIGHT_KEYWORDS, color: '#fce7f3', text: '#9d174d' },
    { label: 'NO_NEED_COL',     values: NO_NEED_COL,     color: '#f1f5f9', text: '#475569' },
  ];
  const panel = document.getElementById('cfgPanel');
  panel.innerHTML = sections.map(s => `
    <div style="margin-bottom:12px;">
      <div style="font-size:11px;font-weight:700;color:#888;letter-spacing:.05em;margin-bottom:5px;">${s.label}</div>
      <div style="display:flex;flex-wrap:wrap;gap:4px;">
        ${s.values.map(v => `<span style="background:${s.color};color:${s.text};border-radius:4px;padding:2px 8px;font-size:12px;font-weight:500;">${v}</span>`).join('')}
      </div>
    </div>
  `).join('');
  panel.style.display = 'block';
}
