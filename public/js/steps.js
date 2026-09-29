// ── Packing List Builder: the step bar at the top of the page ──
// Five named stages, each shown as waiting, in progress, done, needs a look
// or failed. The rest of the page reports progress with setSteps(); nothing
// here decides anything about the data.
// ─────────────────────────────────────────────────────────────
const PL_STEPS = [
  { id: 'upload',   name: 'Upload',     hint: 'Choose the Excel file',     target: 'dropZone' },
  { id: 'read',     name: 'Read',       hint: 'Open the sheet',            target: 'dropZone' },
  { id: 'table',    name: 'Find table', hint: 'Header, rows and columns',  target: 'plSheetLegend' },
  { id: 'preview',  name: 'Preview',    hint: 'Check and edit the result', target: 'plPreview' },
  { id: 'download', name: 'Download',   hint: 'Generate the .xlsx',        target: 'submitBtn' },
];

// State per step: 'todo' | 'active' | 'done' | 'warn' | 'error'.
// A note, when set, replaces the step's hint with what actually happened.
const stepState = {};
const stepNote  = {};
PL_STEPS.forEach(s => { stepState[s.id] = 'todo'; stepNote[s.id] = ''; });
stepState.upload = 'active';

const STEP_MARK  = { todo: '', active: '', done: '✓', warn: '!', error: '✕' };
const STEP_LABEL = { todo: 'waiting', active: 'in progress', done: 'done', warn: 'needs a look', error: 'failed' };

// setSteps({ read: 'done', table: ['warn', 'Pick the Carton column'] })
// A plain state clears the note; a [state, note] pair sets both.
function setSteps(changes) {
  for (const [id, value] of Object.entries(changes)) {
    const [state, note] = Array.isArray(value) ? value : [value, ''];
    stepState[id] = state;
    stepNote[id]  = note;
  }
  renderSteps();
}

function renderSteps() {
  const bar = document.getElementById('plSteps');
  if (!bar) return;
  bar.innerHTML = PL_STEPS.map((s, i) => {
    const state = stepState[s.id];
    const mark  = STEP_MARK[state] || String(i + 1);
    const text  = stepNote[s.id] || s.hint;
    return `
      <li class="pl-step pl-step-${state}" ${state === 'active' ? 'aria-current="step"' : ''}>
        <button type="button" class="pl-step-btn" data-target="${s.target}"
                title="${escapeHtml(s.name)}: ${STEP_LABEL[state]}">
          <span class="pl-step-dot">${mark}</span>
          <span class="pl-step-text">
            <span class="pl-step-name">${escapeHtml(s.name)}</span>
            <span class="pl-step-hint">${escapeHtml(text)}</span>
          </span>
        </button>
      </li>`;
  }).join('');
}

// Clicking a step scrolls to the part of the page it is about, if that part
// is showing yet.
document.getElementById('plSteps').addEventListener('click', e => {
  const btn = e.target.closest('.pl-step-btn');
  if (!btn) return;
  const el = document.getElementById(btn.dataset.target);
  if (el && el.offsetParent !== null) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
});

renderSteps();
