/* ════════════════════════════════════════════
   State
═══════════════════════════════════════════ */
let originalRows  = [];   // raw objects from XLSX
let headers       = [];   // all columns
let selectedCols  = new Set();
let currentRows   = [];   // displayed rows (with __origOrder)
let isRandomized  = false;

/* Condition config — built in step 3 */
// conditionCols: Array of { col, delimiter, selectedTokens: Set }
let conditionCols = [];

/* Which step (0-4) the user is currently on */
let currentStep = 0;

/* ════════════════════════════════════════════
   DOM refs
═══════════════════════════════════════════ */
const fileInput          = document.getElementById('fileInput');
const dropZone           = document.getElementById('dropZone');
const fileNameEl         = document.getElementById('fileName');
const columnChips        = document.getElementById('columnChips');
const btnSelectAll       = document.getElementById('btnSelectAll');
const btnDeselectAll     = document.getElementById('btnDeselectAll');
const btnApplyColumns    = document.getElementById('btnApplyColumns');
const conditionBuilder   = document.getElementById('conditionBuilder');
const btnBackToColumns   = document.getElementById('btnBackToColumns');
const btnSkipConditions  = document.getElementById('btnSkipConditions');
const btnApplyConditions = document.getElementById('btnApplyConditions');
const tableHead          = document.getElementById('tableHead');
const tableBody          = document.getElementById('tableBody');
const btnRandomize       = document.getElementById('btnRandomize');
const btnReset           = document.getElementById('btnReset');
const rowCountEl         = document.getElementById('rowCount');
const btnBackToConditions= document.getElementById('btnBackToConditions');
const btnConfirm         = document.getElementById('btnConfirm');
const confirmHint        = document.getElementById('confirmHint');
const conditionSummary   = document.getElementById('conditionSummary');
const btnBackToPreview   = document.getElementById('btnBackToPreview');
const btnExport          = document.getElementById('btnExport');
const btnStartOver       = document.getElementById('btnStartOver');

const sections = [
  document.getElementById('section-upload'),
  document.getElementById('section-columns'),
  document.getElementById('section-conditions'),
  document.getElementById('section-preview'),
  document.getElementById('section-export'),
];

/* ════════════════════════════════════════════
   Navigation
═══════════════════════════════════════════ */
function goTo(step) {
  currentStep = step;
  sections.forEach((s, i) => s.classList.toggle('hidden', i !== step));
  sections[step].scrollIntoView({ behavior: 'smooth', block: 'start' });
  updateStepper();
}

function updateStepper() {
  document.querySelectorAll('.step-btn').forEach(btn => {
    const s = +btn.dataset.step;
    btn.classList.remove('active', 'done');
    if (s === currentStep) btn.classList.add('active');
    else if (s < currentStep) btn.classList.add('done');
    btn.disabled = (s > currentStep);
  });
}

document.querySelectorAll('.step-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    const s = +btn.dataset.step;
    if (s < currentStep) goTo(s);
  });
});

/* Back buttons */
btnBackToColumns.addEventListener('click',    () => goTo(1));
btnBackToConditions.addEventListener('click', () => goTo(2));
btnBackToPreview.addEventListener('click',    () => goTo(3));

/* ════════════════════════════════════════════
   File upload
═══════════════════════════════════════════ */
fileInput.addEventListener('change', e => handleFile(e.target.files[0]));

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
});

function handleFile(file) {
  if (!file) return;
  fileNameEl.textContent = '📄 ' + file.name;
  const reader = new FileReader();
  reader.onload = evt => {
    const data = new Uint8Array(evt.target.result);
    const wb   = XLSX.read(data, { type: 'array' });
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (!json.length) { alert('The sheet appears to be empty.'); return; }

    headers      = Object.keys(json[0]);
    originalRows = json;
    selectedCols = new Set(headers);
    conditionCols = [];

    buildChips();
    goTo(1);
  };
  reader.readAsArrayBuffer(file);
}

/* ════════════════════════════════════════════
   Step 2 — Column chips
═══════════════════════════════════════════ */
function buildChips() {
  columnChips.innerHTML = '';
  headers.forEach(col => {
    const btn = document.createElement('button');
    btn.className   = 'chip active';
    btn.textContent = col;
    btn.dataset.col = col;
    btn.addEventListener('click', () => {
      if (selectedCols.has(col)) { selectedCols.delete(col); btn.classList.remove('active'); }
      else                       { selectedCols.add(col);    btn.classList.add('active'); }
    });
    columnChips.appendChild(btn);
  });
}

btnSelectAll.addEventListener('click', () => {
  headers.forEach(c => selectedCols.add(c));
  columnChips.querySelectorAll('.chip').forEach(c => c.classList.add('active'));
});
btnDeselectAll.addEventListener('click', () => {
  selectedCols.clear();
  columnChips.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
});

btnApplyColumns.addEventListener('click', () => {
  if (!selectedCols.size) { alert('Please select at least one column.'); return; }
  buildConditionUI();
  goTo(2);
});

/* ════════════════════════════════════════════
   Step 3 — Condition configuration UI
═══════════════════════════════════════════ */
const DELIMITERS = [
  { label: 'Space',    value: ' '  },
  { label: 'Comma',   value: ','   },
  { label: 'Semicolon', value: ';' },
  { label: 'Pipe',    value: '|'   },
  { label: 'Tab',     value: '\t'  },
  { label: 'Custom',  value: '__custom__' },
];

function buildConditionUI() {
  conditionBuilder.innerHTML = '';
  const cols = headers.filter(h => selectedCols.has(h));

  // Restore or initialise conditionCols entries
  const prevMap = {};
  conditionCols.forEach(cc => { prevMap[cc.col] = cc; });
  conditionCols = cols.map(col => prevMap[col] || { col, enabled: false, delimiter: ' ', customDelim: '', selectedTokens: new Set() });

  cols.forEach((col, idx) => {
    const state = conditionCols[idx];
    const card  = document.createElement('div');
    card.className = 'condition-col-card';

    /* ── Header row: enable toggle ── */
    card.innerHTML = `
      <h3>${escHtml(col)}</h3>
      <div class="col-toggle-row">
        <input type="checkbox" id="cond-enable-${idx}" ${state.enabled ? 'checked' : ''}>
        <label class="toggle-label" for="cond-enable-${idx}">Use this column as a condition source</label>
      </div>
      <div class="delimiter-row" id="delim-row-${idx}" style="${state.enabled ? '' : 'display:none'}">
        <span>Split cell values by:</span>
        <select id="delim-sel-${idx}">
          ${DELIMITERS.map(d => `<option value="${d.value}" ${d.value === state.delimiter ? 'selected' : ''}>${d.label}</option>`).join('')}
        </select>
        <input type="text" id="delim-custom-${idx}" placeholder="e.g. /" style="display:${state.delimiter === '__custom__' ? 'inline-block' : 'none'}" value="${escHtml(state.customDelim)}">
      </div>
      <div id="token-area-${idx}" style="${state.enabled ? '' : 'display:none'}">
        <p class="token-hint">Click the tokens that represent conditions (highlighted in red):</p>
        <div class="token-preview" id="token-preview-${idx}"></div>
      </div>`;
    conditionBuilder.appendChild(card);

    /* Wire up enable toggle */
    const chk     = card.querySelector(`#cond-enable-${idx}`);
    const delimRow = card.querySelector(`#delim-row-${idx}`);
    const tokenArea = card.querySelector(`#token-area-${idx}`);
    chk.addEventListener('change', () => {
      state.enabled = chk.checked;
      delimRow.style.display  = chk.checked ? '' : 'none';
      tokenArea.style.display = chk.checked ? '' : 'none';
      if (chk.checked) refreshTokenPreview(idx);
    });

    /* Delimiter selector */
    const delimSel    = card.querySelector(`#delim-sel-${idx}`);
    const customInput = card.querySelector(`#delim-custom-${idx}`);
    delimSel.addEventListener('change', () => {
      state.delimiter = delimSel.value;
      customInput.style.display = delimSel.value === '__custom__' ? 'inline-block' : 'none';
      state.selectedTokens.clear();
      refreshTokenPreview(idx);
    });
    customInput.addEventListener('input', () => {
      state.customDelim = customInput.value;
      state.selectedTokens.clear();
      refreshTokenPreview(idx);
    });

    if (state.enabled) refreshTokenPreview(idx);
  });
}

function getDelimiter(state) {
  return state.delimiter === '__custom__' ? (state.customDelim || ' ') : state.delimiter;
}

function getAllTokens(state) {
  const delim  = getDelimiter(state);
  const tokens = new Set();
  originalRows.forEach(row => {
    const cell = String(row[state.col] ?? '');
    cell.split(delim).map(t => t.trim()).filter(Boolean).forEach(t => tokens.add(t));
  });
  return [...tokens].sort();
}

function refreshTokenPreview(idx) {
  const state   = conditionCols[idx];
  const preview = document.getElementById(`token-preview-${idx}`);
  if (!preview) return;
  preview.innerHTML = '';
  getAllTokens(state).forEach(tok => {
    const chip = document.createElement('button');
    chip.className   = 'token-chip' + (state.selectedTokens.has(tok) ? ' selected' : '');
    chip.textContent = tok;
    chip.addEventListener('click', () => {
      if (state.selectedTokens.has(tok)) state.selectedTokens.delete(tok);
      else                               state.selectedTokens.add(tok);
      chip.classList.toggle('selected');
    });
    preview.appendChild(chip);
  });
}

btnApplyConditions.addEventListener('click', () => {
  // Validate: any enabled column must have ≥1 token selected
  const active = conditionCols.filter(cc => cc.enabled);
  for (const cc of active) {
    if (!cc.selectedTokens.size) {
      alert(`Please select at least one condition token for column "${cc.col}", or disable it.`);
      return;
    }
  }
  isRandomized = false;
  currentRows  = originalRows.map((row, i) => ({ __origOrder: i + 1, ...row }));
  renderTable(currentRows);
  updateConditionSummary();
  btnConfirm.disabled = true;
  confirmHint.textContent = 'Randomize first to enable export.';
  goTo(3);
});

btnSkipConditions.addEventListener('click', () => {
  conditionCols.forEach(cc => { cc.enabled = false; cc.selectedTokens.clear(); });
  isRandomized = false;
  currentRows  = originalRows.map((row, i) => ({ __origOrder: i + 1, ...row }));
  renderTable(currentRows);
  updateConditionSummary();
  btnConfirm.disabled = true;
  confirmHint.textContent = 'Randomize first to enable export.';
  goTo(3);
});

function updateConditionSummary() {
  const active = conditionCols.filter(cc => cc.enabled && cc.selectedTokens.size);
  if (!active.length) {
    conditionSummary.classList.add('hidden');
    return;
  }
  conditionSummary.classList.remove('hidden');
  const parts = active.map(cc => `<strong>${escHtml(cc.col)}</strong>: ${[...cc.selectedTokens].map(escHtml).join(', ')}`);
  conditionSummary.innerHTML = '🎯 Spacing conditions — ' + parts.join(' | ');
}

/* ════════════════════════════════════════════
   Step 4 — Table rendering
═══════════════════════════════════════════ */
function conditionColSet() {
  return new Set(conditionCols.filter(cc => cc.enabled && cc.selectedTokens.size).map(cc => cc.col));
}

function renderTable(rows) {
  const cols   = headers.filter(h => selectedCols.has(h));
  const condSet = conditionColSet();

  /* Header */
  tableHead.innerHTML = '';
  const tr = document.createElement('tr');
  addTH(tr, 'New Order', 'col-new-order');
  addTH(tr, 'Original Order', 'col-orig-order');
  cols.forEach(c => addTH(tr, c, condSet.has(c) ? 'col-condition' : ''));
  tableHead.appendChild(tr);

  /* Rows */
  tableBody.innerHTML = '';
  rows.forEach((row, idx) => {
    const tr = document.createElement('tr');
    addTD(tr, idx + 1, 'col-new-order');
    addTD(tr, row.__origOrder, 'col-orig-order');
    cols.forEach(c => addTD(tr, row[c] ?? '', condSet.has(c) ? 'col-condition' : ''));
    tableBody.appendChild(tr);
  });

  rowCountEl.textContent = `${rows.length} row${rows.length !== 1 ? 's' : ''}`;
}

function addTH(tr, text, cls) {
  const th = document.createElement('th');
  th.textContent = text;
  if (cls) th.className = cls;
  tr.appendChild(th);
}
function addTD(tr, text, cls) {
  const td = document.createElement('td');
  td.textContent = text;
  if (cls) td.className = cls;
  tr.appendChild(td);
}

/* ════════════════════════════════════════════
   Condition-aware shuffle
═══════════════════════════════════════════ */

/**
 * Extract condition key for a row: concatenation of all selected condition tokens
 * found in the row (across all enabled condition columns).
 */
function getConditionKey(row) {
  const parts = [];
  conditionCols.forEach(cc => {
    if (!cc.enabled || !cc.selectedTokens.size) return;
    const delim  = getDelimiter(cc);
    const cell   = String(row[cc.col] ?? '');
    const tokens = cell.split(delim).map(t => t.trim()).filter(t => cc.selectedTokens.has(t));
    if (tokens.length) parts.push(tokens.sort().join('+'));
  });
  return parts.length ? parts.join('|') : '__none__';
}

/**
 * Condition-aware shuffle: tries to maximise the minimum gap between rows
 * that share the same condition key.
 * Algorithm: greedy insertion — always pick the candidate row whose condition
 * key was seen least recently (furthest back in the output so far).
 */
function conditionAwareShuffle(rows) {
  const active = conditionCols.some(cc => cc.enabled && cc.selectedTokens.size);
  if (!active) return shuffle([...rows]);

  // Group rows by condition key
  const byKey = {};
  rows.forEach(row => {
    const k = getConditionKey(row);
    (byKey[k] = byKey[k] || []).push(row);
  });

  // Shuffle within each group first (so ties are random)
  Object.values(byKey).forEach(arr => shuffle(arr));

  // Build result greedily
  const result     = [];
  const lastSeen   = {};   // key → last index in result
  const remaining  = Object.values(byKey).flat();

  // We keep a "pool" and pick greedily
  const pool = [...remaining];
  shuffle(pool); // initial random order as base

  // Sort pool: key that was seen longest ago goes first
  const pickBest = (pool) => {
    let bestIdx = 0;
    let bestScore = -Infinity;
    pool.forEach((row, i) => {
      const k    = getConditionKey(row);
      const seen = lastSeen[k] !== undefined ? lastSeen[k] : -Infinity;
      const gap  = result.length - seen;
      if (gap > bestScore) { bestScore = gap; bestIdx = i; }
    });
    return bestIdx;
  };

  while (pool.length) {
    const idx  = pickBest(pool);
    const row  = pool.splice(idx, 1)[0];
    lastSeen[getConditionKey(row)] = result.length;
    result.push(row);
  }
  return result;
}

function shuffle(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

/* ════════════════════════════════════════════
   Randomize / Reset buttons
═══════════════════════════════════════════ */
btnRandomize.addEventListener('click', () => {
  const base   = originalRows.map((row, i) => ({ __origOrder: i + 1, ...row }));
  currentRows  = conditionAwareShuffle(base);
  isRandomized = true;
  renderTable(currentRows);
  btnConfirm.disabled = false;
  confirmHint.textContent = 'Looks good? Confirm to proceed to export.';
});

btnReset.addEventListener('click', () => {
  isRandomized = false;
  currentRows  = originalRows.map((row, i) => ({ __origOrder: i + 1, ...row }));
  renderTable(currentRows);
  btnConfirm.disabled = true;
  confirmHint.textContent = 'Randomize first to enable export.';
});

btnConfirm.addEventListener('click', () => goTo(4));

/* ════════════════════════════════════════════
   Export
═══════════════════════════════════════════ */
btnExport.addEventListener('click', () => {
  const cols = headers.filter(h => selectedCols.has(h));
  const exportData = currentRows.map((row, idx) => {
    const obj = { 'New Order': idx + 1, 'Original Order': row.__origOrder };
    cols.forEach(c => { obj[c] = row[c] ?? ''; });
    return obj;
  });

  const ws = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Randomized');
  ws['!cols'] = Object.keys(exportData[0]).map(k => ({ wch: Math.max(k.length + 2, 10) }));
  XLSX.writeFile(wb, 'randomized_output.xlsx');
});

/* ════════════════════════════════════════════
   Start over
═══════════════════════════════════════════ */
btnStartOver.addEventListener('click', () => {
  originalRows  = [];
  headers       = [];
  selectedCols  = new Set();
  currentRows   = [];
  conditionCols = [];
  isRandomized  = false;
  fileInput.value = '';
  fileNameEl.textContent = '';
  tableHead.innerHTML = '';
  tableBody.innerHTML = '';
  columnChips.innerHTML = '';
  conditionBuilder.innerHTML = '';
  conditionSummary.classList.add('hidden');
  goTo(0);
});

/* ════════════════════════════════════════════
   Helpers
═══════════════════════════════════════════ */
function escHtml(str) {
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

/* Initial stepper state */
updateStepper();
