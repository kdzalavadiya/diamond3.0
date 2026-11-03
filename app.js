/* =========================================================================
   Diamond Ledger — UI controller
   Loaded after config.js, gas-shim.js and Code.js.
   All handlers stay global because the markup wires them with inline events.
   ========================================================================= */

/* ------------------------------------------------------------------ utils */
function el(id) { return document.getElementById(id); }

function localISO(d) {
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function esc(v) {
  return String(v === undefined || v === null ? '' : v).replace(/[&<>"']/g,
    c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function safeRow(o) {
  const c = {};
  for (const k in o) c[k] = typeof o[k] === 'string' ? esc(o[k]) : o[k];
  return c;
}

function num(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }

/* Money always uses a leading sign: -$12.34, never $-12.34 */
function fmtUsd(v, dp) {
  dp = dp === undefined ? 2 : dp;
  const raw = parseFloat(v);
  if (isNaN(raw)) return '$0.' + (dp ? '0'.repeat(dp) : '');
  const body = Math.abs(raw).toLocaleString(undefined, { minimumFractionDigits: dp, maximumFractionDigits: dp });
  return (raw < 0 ? '-$' : '$') + body;
}

function fmtInr(v) {
  const raw = parseFloat(v);
  if (isNaN(raw)) return '₹0';
  return (raw < 0 ? '-₹' : '₹') + Math.abs(Math.round(raw)).toLocaleString();
}

function fmtPct(v, dp) {
  const raw = parseFloat(v);
  if (isNaN(raw)) return '0.00%';
  return (raw < 0 ? '-' : '') + Math.abs(raw).toFixed(dp === undefined ? 2 : dp) + '%';
}

/* Class for signed money: losses red, gains green. */
function moneyClass(v) {
  const raw = parseFloat(v);
  if (isNaN(raw) || raw === 0) return '';
  return raw < 0 ? 'money-neg' : 'money-pos';
}

function dash(v) {
  const s = String(v === undefined || v === null ? '' : v).trim();
  return s === '' ? '<span class="muted">–</span>' : esc(s);
}

function debounce(fn, ms) {
  let t = null;
  return function () {
    const args = arguments;
    clearTimeout(t);
    t = setTimeout(() => fn.apply(null, args), ms);
  };
}

/* -------------------------------------------------------------- app state */
let globalStore = {
  purchases: [],
  sales: [],
  currentUser: 'Keyur Zalavadiya'
};
let currentSelectedPurchase = null;
let currentReportData = [];
let currentPurchaseFilter = 'ALL';
let currentPurchaseTypeFilter = 'ALL';
let currentSheetUrl = '';
let readOnly = false;
let advancedFilters = { purchase: null, sales: null };

const TABLES = {
  purchases: { id: 'tbl-purchases', cols: 20, rerender: () => filterPurchasesTable() },
  stock: { id: 'tbl-stock', cols: 11, rerender: () => filterStockTable() },
  sales: { id: 'tbl-sales', cols: 13, rerender: () => filterSalesTable() },
  report: { id: 'tbl-report', cols: 34, rerender: () => filterReportTable() }
};

const SORT_DEFAULTS = {
  purchases: { key: 'PurchaseDate', dir: 'desc' },
  stock: { key: 'Size', dir: 'desc' },
  sales: { key: 'Date', dir: 'desc' },
  report: { key: 'SaleDate', dir: 'desc' }
};
let sortState = loadSortState();

function loadSortState() {
  const out = JSON.parse(JSON.stringify(SORT_DEFAULTS));
  try {
    const saved = JSON.parse(localStorage.getItem('gem_sort') || '{}');
    Object.keys(out).forEach(k => { if (saved[k] && saved[k].key) out[k] = saved[k]; });
  } catch (e) { /* ignore corrupt state */ }
  return out;
}
function saveSortState() {
  try { localStorage.setItem('gem_sort', JSON.stringify(sortState)); } catch (e) { /* ignore */ }
}

/* ------------------------------------------------------ toast + confirm UI */
const TOAST_ICON = { success: '✓', error: '!', info: 'i' };
let toastQueue = [];
let toastBusy = false;
let toastTimer = null;
let toastCurrentType = null;

function showToast(msg, type = 'info') {
  toastQueue.push({ msg: String(msg || ''), type: TOAST_ICON[type] ? type : 'info' });
  pumpToasts();
}

/* Same severity waits its turn; a different severity replaces what is on screen,
   so the result of a save never appears after an earlier validation error. */
function pumpToasts() {
  const box = el('toast');
  if (!box || !toastQueue.length) return;

  if (toastBusy) {
    if (toastQueue[0].type === toastCurrentType) return;
    clearTimeout(toastTimer);
    box.className = '';
    toastBusy = false;
  }

  const item = toastQueue.shift();
  toastBusy = true;
  toastCurrentType = item.type;
  box.className = '';
  box.innerHTML = `<span class="toast-ico" aria-hidden="true">${TOAST_ICON[item.type]}</span>`
    + `<span class="toast-msg">${esc(item.msg)}</span>`
    + `<button type="button" class="toast-x" aria-label="Dismiss notification">×</button>`;
  const close = () => {
    clearTimeout(toastTimer);
    box.className = '';
    toastBusy = false;
    toastCurrentType = null;
    setTimeout(pumpToasts, 260);
  };
  box.querySelector('.toast-x').addEventListener('click', close);
  void box.offsetWidth;               /* restart the transition */
  box.className = `show ${item.type}`;
  toastTimer = setTimeout(close, item.type === 'error' ? 5200 : 3200);
}

let confirmResolve = null;
function askConfirm(opts) {
  return new Promise(resolve => {
    confirmResolve = resolve;
    el('confirm-title').innerText = opts.title || 'Please confirm';
    el('confirm-body').innerHTML = opts.body || '';
    const ok = el('confirm-ok');
    ok.innerText = opts.okText || 'Confirm';
    ok.className = 'btn ' + (opts.primary ? 'btn-primary' : 'btn-danger');
    const modal = el('confirmModal');
    modal.classList.add('active');
    setTimeout(() => ok.focus(), 40);
  });
}

function closeConfirm(value) {
  const modal = el('confirmModal');
  if (modal) modal.classList.remove('active');
  if (confirmResolve) {
    const r = confirmResolve;
    confirmResolve = null;
    r(!!value);
  }
}

/* -------------------------------------------------------- connection state */
function setConnState(state, text) {
  const dot = el('conn-dot');
  const label = el('conn-state');
  if (dot) dot.className = 'status-dot ' + state;
  if (label) label.innerText = text;
  const refresh = el('btn-refresh');
  if (refresh) refresh.classList.toggle('spin', state === 'syncing');
}

function setReadOnly(on) {
  readOnly = !!on;
  const badge = el('ro-badge');
  if (badge) badge.style.display = readOnly ? 'inline-flex' : 'none';
  ['purchaseSaveBtn', 'saleSaveBtn'].forEach(id => {
    const b = el(id);
    if (b) { b.disabled = readOnly; b.title = readOnly ? 'This sheet is read-only for your Google account' : ''; }
  });
  if (readOnly) {
    ['purchaseDeleteBtn', 'saleDeleteBtn'].forEach(id => { const b = el(id); if (b) b.disabled = true; });
    setConnState('readonly', 'Read-only');
  }
}

function friendlyError(err) {
  const msg = String((err && err.message) || err || 'Unknown error');
  if (/View-only|permission to edit/i.test(msg)) return 'This sheet is read-only for your account — ask the owner for Edit access.';
  if (/Cannot open this Google Sheet/i.test(msg)) return 'That sheet is not shared with your Google account, or the link is wrong.';
  if (/No Google Sheet linked/i.test(msg)) return 'No Google Sheet is connected yet — click “Connect Sheet”.';
  if (/401|invalid_grant|token/i.test(msg)) return 'Your Google session expired. Sign in again, then retry.';
  if (/quota|rate limit|429/i.test(msg)) return 'Google is rate-limiting this account. Wait a minute and try again.';
  if (/not found in the sheet/i.test(msg)) return msg;
  if (/already exists|already sold|already recorded|is required/i.test(msg)) return msg;
  return msg;
}

/* Single place that talks to the backend, so every call reports its state. */
function callBackend(name, args) {
  return new Promise((resolve, reject) => {
    setConnState('syncing', 'Syncing…');
    google.script.run
      .withSuccessHandler(res => {
        setConnState('connected', 'Connected');
        resolve(res);
      })
      .withFailureHandler(err => {
        const msg = String((err && err.message) || '');
        if (/View-only|permission to edit/i.test(msg)) setReadOnly(true);
        setConnState('error', 'Sync failed');
        reject(err);
      })
      [name].apply(null, args || []);
  });
}

/* ------------------------------------------------------- boot / init wiring */
window.addEventListener('DOMContentLoaded', () => {
  initTheme();
  bootUI();

  applyDatePreset('THIS_MONTH');
  const todayStr = localISO(new Date());
  if (el('p-PurchaseDate')) el('p-PurchaseDate').value = todayStr;
  if (el('s-Date')) el('s-Date').value = todayStr;

  const savedUser = localStorage.getItem('gem_user') || '';
  setOperatorState(savedUser);

  // Add search listeners
  if(el('pur-search')) el('pur-search').addEventListener('input', filterPurchasesTable);
  if(el('sales-search')) el('sales-search').addEventListener('input', filterSalesTable);
  if(el('stock-search')) el('stock-search').addEventListener('input', filterStockTable);
  if(el('rep-search')) el('rep-search').addEventListener('input', filterReportTable);

  setConnState('syncing', 'Connecting…');
  renderSkeletons();
  loadInitial();
});

async function loadInitial() {
  try {
    const data = await callBackend('getInitialData');
    applyPayload(data);
    showToast('Connected · ' + globalStore.purchases.length + ' purchases · ' + globalStore.sales.length + ' sales', 'success');
  } catch (err) {
    showToast(friendlyError(err), 'error');
    if (/No Google Sheet linked/i.test(String(err.message))) openSheetModal();
    renderAllTables();
  }
}

function applyPayload(data) {
  globalStore.purchases = data.purchases || [];
  globalStore.sales = data.sales || [];
  if (data.sheetUrl) updateSheetLink(data.sheetUrl);
  if (data.currentUser && !localStorage.getItem('gem_user')) setOperatorState(data.currentUser);
  renderAllTables();
}

async function refreshFromSheet(silent) {
  if (!silent) showToast('Re-reading the Google Sheet…', 'info');
  try {
    const data = await callBackend('getInitialData');
    applyPayload(data);
    if (!silent) showToast('Sheet refreshed', 'success');
  } catch (err) {
    showToast(friendlyError(err), 'error');
  }
}

/* Everything the markup cannot express declaratively: label association,
   column scope, table ids, sortable headers, collapsible sections, search. */
function bootUI() {
  /* 1. programmatic labels + table semantics */
  document.querySelectorAll('.form-group').forEach(group => {
    const label = group.querySelector('label');
    const control = group.querySelector('input, select, textarea');
    if (label && control && control.id) label.setAttribute('for', control.id);
  });

  document.querySelectorAll('.table-search-input').forEach(input => {
    if (!input.getAttribute('aria-label')) input.setAttribute('aria-label', input.placeholder || 'Search');
    const wrap = input.parentNode;
    const clear = document.createElement('button');
    clear.type = 'button';
    clear.className = 'search-clear';
    clear.setAttribute('aria-label', 'Clear search');
    clear.innerText = '×';
    clear.addEventListener('click', () => {
      input.value = '';
      wrap.classList.remove('has-value');
      input.dispatchEvent(new Event('input'));
      input.focus();
    });
    wrap.appendChild(clear);
    input.addEventListener('input', () => wrap.classList.toggle('has-value', !!input.value));
  });

  /* 2. stable ids for the four data tables + header scope */
  const tableIds = [
    ['#purchase-tab table', 'tbl-purchases'],
    ['#sale-tab .upper-panel table', 'tbl-stock'],
    ['#sale-tab .lower-panel table', 'tbl-sales'],
    ['#report-tab table', 'tbl-report']
  ];
  tableIds.forEach(([sel, id]) => {
    const t = document.querySelector(sel);
    if (!t) return;
    t.id = id;
    t.classList.add('sticky-cols');
    t.querySelectorAll('thead th').forEach(th => th.setAttribute('scope', 'col'));
  });

  /* 3. modal semantics */
  [['sheetModal', 'Connect Google Sheet'], ['userModal', 'Switch active operator'], ['confirmModal', 'Confirm']]
    .forEach(([id, label]) => {
      const m = el(id);
      if (!m) return;
      m.setAttribute('role', 'dialog');
      m.setAttribute('aria-modal', 'true');
      m.setAttribute('aria-label', label);
      m.addEventListener('mousedown', e => { if (e.target === m) closeOverlay(id); });
    });

  /* 4. collapsible purchase sections (remembered per section) */
  initCollapsibleSections();

  /* 5. sortable headers */
  initSorting();

  /* 6. search inputs (debounced) */
  const wire = (id, fn) => { const i = el(id); if (i) i.addEventListener('input', debounce(fn, 160)); };
  wire('pur-search', filterPurchasesTable);
  wire('stock-search', filterStockTable);
  wire('sales-search', filterSalesTable);
  wire('rep-search', filterReportTable);

  /* 7. keyboard */
  window.addEventListener('keydown', onGlobalKey);
}

function initCollapsibleSections() {
  const cards = Array.from(document.querySelectorAll('#purchaseForm .form-section-card'));
  cards.forEach((card, index) => {
    const label = card.querySelector('.form-section-label');
    if (!label || card.dataset.collapsible === 'off') return;

    const body = document.createElement('div');
    body.className = 'section-body';
    Array.from(card.children).forEach(child => { if (child !== label) body.appendChild(child); });
    card.appendChild(body);
    card.classList.add('collapsible');

    const key = 'gem_sec_' + index;
    const stored = localStorage.getItem(key);
    const shouldCollapse = stored === 'closed' || (stored === null && index === 1);
    if (shouldCollapse) card.classList.add('collapsed');

    label.setAttribute('role', 'button');
    label.setAttribute('tabindex', '0');
    label.setAttribute('aria-expanded', String(!card.classList.contains('collapsed')));
    const chev = document.createElement('span');
    chev.className = 'chev';
    chev.setAttribute('aria-hidden', 'true');
    chev.textContent = '▾';
    label.appendChild(chev);

    const toggle = () => {
      card.classList.toggle('collapsed');
      const open = !card.classList.contains('collapsed');
      label.setAttribute('aria-expanded', String(open));
      localStorage.setItem(key, open ? 'open' : 'closed');
    };
    label.addEventListener('click', toggle);
    label.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(); }
    });
  });
}

/* Header text -> record field, per table. Keys are lowercased with spaces and
   dots removed but currency symbols kept, so the two "Total Amount" columns stay distinct. */
const SORT_MAP = {
  'purid': 'PurId', 'status': 'Status', 'gia/lotno': 'GIANo', 'date': 'PurchaseDate',
  'purchasedate': 'PurchaseDate', 'pur.date': 'PurchaseDate', 'supplier(fromname)': 'FromName',
  'fromname': 'FromName', 'shape': 'Shape', 'weight': 'Size', 'size': 'Size', 'color': 'Color',
  'clarity': 'Quality', 'quality': 'Quality', 'percarat': 'PerCarat', 'cost($)': 'DAmount',
  'totalamount($)': 'DAmount', 'totalamount(₹)': 'IndAmount', 'damount($)': 'DAmount',
  'indamount(₹)': 'IndAmount', 'rate(₹)': 'Rate', 'fxrate': 'Rate', 'inbound(₹)': 'IndAmount',
  'btype': 'BType', 'bill%': 'BillPct', 'days': 'PaymentDay', 'broker': 'Broker', 'operator': 'User', 'type': 'Type',
  'saleid': 'SaleId', 'client(buyer)': 'ClientName', 'clientname': 'ClientName', 'country': 'Country',
  'salepercarat': 'SalePerCarat', 'grossinvoice($)': 'SaleDFinal', 'saledfinal($)': 'SaleDFinal',
  'advance($)': 'AdvDAmt', 'advdamt': 'AdvDAmt', 'duebalance($)': 'DueAmt', 'dueamt': 'DueAmt',
  'netprofit($)': 'SaleDProfit', 'saledprofit': 'SaleDProfit', 'margin%': 'ProfitPct',
  'profit%': 'ProfitPct', 'remark': 'SaleRemark', 'saleremark': 'SaleRemark', 'saledate': 'SaleDate',
  'pur.user': 'PurUser', 'saleuser': 'SaleUser', 'saleless%': 'SaleLess', 'less%': 'Less'
};

function normHeader(s) {
  return String(s || '').toLowerCase().replace(/\s+/g, '').replace(/\./g, '');
}

function initSorting() {
  Object.keys(TABLES).forEach(key => {
    const table = el(TABLES[key].id);
    if (!table) return;
    const rows = table.querySelectorAll('thead tr');
    const lastRow = rows[rows.length - 1];
    if (!lastRow) return;
    lastRow.querySelectorAll('th').forEach(th => {
      const field = SORT_MAP[normHeader(th.textContent)];
      if (!field) return;
      th.dataset.sortKey = field;
      th.classList.add('sortable');
      th.setAttribute('tabindex', '0');
      th.setAttribute('title', 'Click to sort by ' + th.textContent.trim());
      const ind = document.createElement('span');
      ind.className = 'sort-ind';
      ind.setAttribute('aria-hidden', 'true');
      th.appendChild(ind);
      const apply = () => {
        const st = sortState[key];
        if (st.key === field) st.dir = st.dir === 'asc' ? 'desc' : 'asc';
        else { st.key = field; st.dir = 'desc'; }
        saveSortState();
        TABLES[key].rerender();
      };
      th.addEventListener('click', apply);
      th.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); apply(); } });
    });
  });
}

function markSort(tableKey) {
  const table = el(TABLES[tableKey].id);
  if (!table) return;
  const st = sortState[tableKey];
  table.querySelectorAll('th.sortable').forEach(th => {
    const active = th.dataset.sortKey === st.key;
    th.classList.toggle('sorted', active);
    const ind = th.querySelector('.sort-ind');
    if (ind) ind.textContent = active ? (st.dir === 'asc' ? '▲' : '▼') : '⇅';
    th.setAttribute('aria-sort', active ? (st.dir === 'asc' ? 'ascending' : 'descending') : 'none');
  });
}

function applySort(list, tableKey) {
  const st = sortState[tableKey];
  if (!st || !st.key) return list;
  const out = list.slice();
  out.sort((a, b) => {
    const av = a[st.key];
    const bv = b[st.key];
    /* Only compare numerically when the whole value is a number; parseFloat
       would turn "2026-07-01" into 2026 and sort every date in a year as equal. */
    const isNumeric = v => /^-?\d+(\.\d+)?$/.test(String(v === undefined || v === null ? '' : v).trim());
    const aNum = isNumeric(av);
    const bNum = isNumeric(bv);
    let cmp;
    if (aNum && bNum) cmp = parseFloat(av) - parseFloat(bv);
    else cmp = String(av === undefined || av === null ? '' : av)
      .localeCompare(String(bv === undefined || bv === null ? '' : bv), undefined, { numeric: true, sensitivity: 'base' });
    return st.dir === 'asc' ? cmp : -cmp;
  });
  return out;
}

/* --------------------------------------------------------- keyboard support */
function isTyping() {
  const a = document.activeElement;
  if (!a) return false;
  const tag = a.tagName;
  return tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA' || a.isContentEditable;
}

function activeTabId() {
  const tab = document.querySelector('.tab-content.active');
  return tab ? tab.id : 'purchase-tab';
}

function onGlobalKey(e) {
  const mod = e.ctrlKey || e.metaKey;

  if (mod && (e.key === 's' || e.key === 'S')) {
    e.preventDefault();
    const formId = activeTabId() === 'sale-tab' ? 'saleForm' : 'purchaseForm';
    const form = el(formId);
    if (form && !readOnly) form.dispatchEvent(new Event('submit', { cancelable: true, bubbles: true }));
    else if (readOnly) showToast('This sheet is read-only for your account.', 'error');
    return;
  }

  if (e.key === '/' && !isTyping()) {
    const map = { 'purchase-tab': 'pur-search', 'sale-tab': 'stock-search', 'report-tab': 'rep-search' };
    const input = el(map[activeTabId()]);
    if (input) { e.preventDefault(); input.focus(); input.select(); }
    return;
  }

  if (e.key === 'Escape') {
    if (el('confirmModal') && el('confirmModal').classList.contains('active')) { closeConfirm(false); return; }
    if (el('sheetModal') && el('sheetModal').classList.contains('active')) { closeSheetModal(); return; }
    if (el('userModal') && el('userModal').classList.contains('active')) { closeUserModal(); return; }
    if (el('header-actions') && el('header-actions').classList.contains('open')) { toggleHeaderMenu(false); return; }
    clearRowHighlights();
    return;
  }

  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    if (isTyping()) return;
    const moved = moveRowHighlight(e.key === 'ArrowDown' ? 1 : -1);
    if (moved) e.preventDefault();
    return;
  }

  if (e.key === 'Enter' && !isTyping()) {
    const row = document.querySelector('tr.kb-row');
    if (row) { e.preventDefault(); row.click(); }
  }
}

function rowTablesForActiveTab() {
  const t = activeTabId();
  if (t === 'purchase-tab') return [el('purchaseTableBody')];
  if (t === 'sale-tab') return [el('saleStockLookupBody'), el('saleRegisterTableBody')];
  return [el('reportTableBody')];
}

function moveRowHighlight(step) {
  const bodies = rowTablesForActiveTab().filter(Boolean);
  const rows = [];
  bodies.forEach(b => b.querySelectorAll('tr[data-id], tr').forEach(r => {
    if (r.querySelector('td[colspan]')) return;
    rows.push(r);
  }));
  if (!rows.length) return false;
  let index = rows.findIndex(r => r.classList.contains('kb-row'));
  rows.forEach(r => r.classList.remove('kb-row'));
  index = index === -1 ? (step > 0 ? 0 : rows.length - 1) : Math.min(rows.length - 1, Math.max(0, index + step));
  const target = rows[index];
  target.classList.add('kb-row');
  target.scrollIntoView({ block: 'nearest' });
  return true;
}

function clearRowHighlights() {
  document.querySelectorAll('tr.kb-row, tr.selected-row').forEach(r => {
    r.classList.remove('kb-row');
    r.classList.remove('selected-row');
  });
}

/* ------------------------------------------------------ sheet + operators */
function updateSheetLink(url) {
  currentSheetUrl = url;
  const link = el('sheet-link');
  if (!link) return;
  if (url) {
    link.href = url;
    link.innerText = 'Open sheet ↗';
    link.style.display = 'inline';
  } else {
    link.href = '#';
    link.innerText = 'Connect a sheet';
  }
}

function openSheetModal() {
  const modal = el('sheetModal');
  const input = el('modal-sheet-input');
  if (!modal || !input) return;
  input.value = currentSheetUrl || '';
  modal.classList.add('active');
  setTimeout(() => input.focus(), 50);
}

function closeSheetModal() {
  const modal = el('sheetModal');
  if (modal) modal.classList.remove('active');
}

// Theme Toggle
function toggleTheme(){
  const html = document.documentElement;
  const current = html.getAttribute('data-theme');
  const newTheme = current === 'dark' ? 'light' : 'dark';
  html.setAttribute('data-theme', newTheme);
  try { localStorage.setItem('gem_theme', newTheme); } catch(e) {}
  const btn = el('btn-theme-toggle');
  if(btn) btn.innerHTML = newTheme === 'dark' ? '☀️' : '🌙';
}

function initTheme(){
  let saved = 'light';
  try { saved = localStorage.getItem('gem_theme') || 'light'; } catch(e) {}
  document.documentElement.setAttribute('data-theme', saved);
  const btn = el('btn-theme-toggle');
  if(btn) btn.innerHTML = saved === 'dark' ? '☀️' : '🌙';
}

// Help Modal
function openHelpModal(){ const m=el('helpModal'); if(m) m.style.display='flex'; }
function closeHelpModal(){ const m=el('helpModal'); if(m) m.style.display='none'; }

function closeOverlay(id) {
  if (id === 'confirmModal') closeConfirm(false);
  else if (id === 'sheetModal') closeSheetModal();
  else if (id === 'userModal') closeUserModal();
  else if (id === 'helpModal') closeHelpModal();
}

async function saveSheetConnection() {
  const input = el('modal-sheet-input');
  if (!input || !input.value.trim()) {
    showToast('Paste a Google Sheet link or ID first.', 'error');
    input && input.focus();
    return;
  }
  const btn = el('btn-save-sheet');
  btn.disabled = true;
  btn.innerText = 'Connecting…';
  try {
    const res = await callBackend('setConnectedSpreadsheet', [input.value.trim()]);
    updateSheetLink(res.sheetUrl);
    globalStore.purchases = res.purchases || [];
    globalStore.sales = res.sales || [];
    renderAllTables();
    closeSheetModal();
    showToast('Connected to the Google Sheet', 'success');
  } catch (err) {
    showToast(friendlyError(err), 'error');
  } finally {
    btn.disabled = false;
    btn.innerText = 'Connect & Sync';
  }
}

function setOperatorState(userName) {
  if (!userName) return;
  userName = String(userName).trim();
  globalStore.currentUser = userName;
  localStorage.setItem('gem_user', userName);

  const headerName = el('header-user-name');
  const headerAvatar = el('header-user-avatar');
  if (headerName) headerName.innerText = userName;
  if (headerAvatar) headerAvatar.innerText = userName.charAt(0).toUpperCase();

  const pUser = el('p-User');
  const sUser = el('s-User');
  if (pUser && !pUser.value) pUser.value = userName;
  if (sUser && !sUser.value) sUser.value = userName;
}

function toggleHeaderMenu(force) {
  const actions = el('header-actions');
  if (!actions) return;
  const open = force === undefined ? !actions.classList.contains('open') : !!force;
  actions.classList.toggle('open', open);
  const btn = el('btn-menu');
  if (btn) btn.setAttribute('aria-expanded', String(open));
}

function switchTab(tabId) {
  document.querySelectorAll('.tab-content').forEach(node => node.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.classList.remove('active');
    btn.setAttribute('aria-selected', 'false');
  });

  const tab = el(tabId);
  const btn = el('btn-' + tabId);
  if (tab) tab.classList.add('active');
  if (btn) { btn.classList.add('active'); btn.setAttribute('aria-selected', 'true'); }
  clearRowHighlights();
  toggleHeaderMenu(false);

  if (tabId === 'report-tab' && currentReportData.length === 0) fetchDateReport();
}

/* --------------------------------------------------------- operator modal */
function promptChangeUser() {
  const modal = el('userModal');
  const input = el('modal-user-input');
  if (!modal || !input) return;
  input.value = globalStore.currentUser;
  modal.classList.add('active');
  setTimeout(() => input.focus(), 50);
}

function closeUserModal() {
  const modal = el('userModal');
  if (modal) modal.classList.remove('active');
}

function saveNewUserFromModal() {
  const input = el('modal-user-input');
  if (input && input.value.trim()) {
    setOperatorState(input.value.trim());
    closeUserModal();
    showToast('Active operator set to ' + globalStore.currentUser, 'success');
  } else if (input) {
    input.classList.add('field-error');
    input.focus();
  }
}

/* ------------------------------------------------------------ date presets */
function applyDatePreset(preset) {
  const now = new Date();
  let fromStr = '';
  let toStr = localISO(now);

  if (preset === 'THIS_MONTH') {
    fromStr = localISO(new Date(now.getFullYear(), now.getMonth(), 1));
  } else if (preset === 'LAST_MONTH') {
    fromStr = localISO(new Date(now.getFullYear(), now.getMonth() - 1, 1));
    toStr = localISO(new Date(now.getFullYear(), now.getMonth(), 0));
  } else if (preset === 'YTD') {
    fromStr = localISO(new Date(now.getFullYear(), 0, 1));
  } else if (preset === 'ALL') {
    fromStr = '1900-01-01';
    toStr = '2099-12-31';
  }

  const fromInput = el('rep-FromDate');
  const toInput = el('rep-ToDate');
  if (fromInput) fromInput.value = fromStr;
  if (toInput) toInput.value = toStr;

  /* reflect the active preset */
  document.querySelectorAll('#report-tab .filter-pills .pill-btn').forEach(b => {
    b.classList.toggle('active', b.getAttribute('onclick') === `applyDatePreset('${preset}')`);
  });
}

/* ------------------------------------------------ math: inbound & outbound */
function onBTypeChange() {
  const type = el('p-BType') ? el('p-BType').value : 'CASH';
  const isBill = type === 'BILL';
  const group = el('p-BillPct-group');
  if (group) group.style.display = isBill ? 'block' : 'none';
  if (!isBill && el('p-BillPct')) el('p-BillPct').value = '';
  calculatePurchaseMath();
}

function calculatePurchaseMath() {
  const size = num(el('p-Size') && el('p-Size').value);
  const perCarat = num(el('p-PerCarat') && el('p-PerCarat').value);
  const less = num(el('p-Less') && el('p-Less').value);
  const rate = num(el('p-Rate') && el('p-Rate').value);
  const isBill = el('p-BType') && el('p-BType').value === 'BILL';
  const billPct = isBill ? num(el('p-BillPct') && el('p-BillPct').value) : 0;

  const baseAfterLess = size * perCarat * (1 + less / 100);
  const billAmt = baseAfterLess * billPct / 100;
  const dAmount = baseAfterLess + billAmt;
  const indAmount = dAmount * rate;

  if (el('p-DAmount')) el('p-DAmount').value = dAmount > 0 ? dAmount.toFixed(3) : '';
  if (el('p-IndAmount')) el('p-IndAmount').value = indAmount > 0 ? indAmount.toFixed(3) : '';

  setText('p-BaseAmt-display', fmtUsd(baseAfterLess, 2));
  setText('p-DAmount-display', fmtUsd(dAmount, 2));
  setText('p-IndAmount-display', fmtInr(indAmount));
  setText('p-BillAmt-display', billAmt > 0 ? '+' + fmtUsd(billAmt, 2) : fmtUsd(0, 2));
}

function calculateSaleMath() {
  if (!currentSelectedPurchase) return;
  const size = num(currentSelectedPurchase.Size);
  const purchaseCost = num(currentSelectedPurchase.DAmount);

  const salePerCarat = num(el('s-SalePerCarat') && el('s-SalePerCarat').value);
  const saleLess = num(el('s-SaleLess') && el('s-SaleLess').value);
  const advAmt = num(el('s-AdvDAmt') && el('s-AdvDAmt').value);

  const saleDFinal = size * salePerCarat * (1 - saleLess / 100);
  const profit = saleDFinal - purchaseCost;
  const dueAmt = saleDFinal - advAmt;
  const profitPct = saleDFinal > 0 ? ((profit / saleDFinal) * 100) : 0;

  if (el('s-SaleDFinal')) el('s-SaleDFinal').value = saleDFinal > 0 ? saleDFinal.toFixed(2) : '';
  if (el('s-SaleDProfit')) el('s-SaleDProfit').value = profit.toFixed(2);
  if (el('s-ProfitPct')) el('s-ProfitPct').value = profitPct.toFixed(2);
  if (el('s-DueAmt')) el('s-DueAmt').value = dueAmt.toFixed(2);

  setText('s-SaleDFinal-display', fmtUsd(saleDFinal, 2));
  setText('s-SaleDProfit-display', fmtUsd(profit, 2));
  setText('s-ProfitPct-display', fmtPct(profitPct, 2));
  setText('s-DueAmt-display', fmtUsd(dueAmt, 2));

  const profitDisplay = el('s-SaleDProfit-display');
  if (profitDisplay) profitDisplay.className = 'terminal-metric-val tabular-nums ' + (profit < 0 ? 'money-neg' : 'money-pos');
  const marginDisplay = el('s-ProfitPct-display');
  if (marginDisplay) marginDisplay.className = 'terminal-metric-val tabular-nums ' + (profitPct < 0 ? 'money-neg' : 'money-pos');
}

/* ------------------------------------------- Total Amount <-> Per Carat
   Type either value and the other fills in: total = carats × per carat.
   The per-carat rate stays the figure written to the sheet (SalePerCarat),
   so on blur the pair settles back onto it and the screen matches the row. */
function roundTo2(v) {
  return Math.round((v + Number.EPSILON) * 100) / 100;
}

function saleStoneSize() {
  return currentSelectedPurchase ? num(currentSelectedPurchase.Size) : 0;
}

function syncSaleTotalField() {
  const field = el('s-TotalAmount');
  if (!field) return;
  const size = saleStoneSize();
  if (!size) return;                     /* no linked stone = nothing to convert */
  const rate = num(el('s-SalePerCarat') && el('s-SalePerCarat').value);
  const gross = size * rate;
  field.value = gross > 0 ? roundTo2(gross).toFixed(2) : '';
}

function onSalePerCaratInput() {
  syncSaleTotalField();
  calculateSaleMath();
}

function onSaleTotalInput() {
  const totalField = el('s-TotalAmount');
  if (!totalField) return;
  const size = saleStoneSize();
  const total = num(totalField.value);

  if (!size) {
    if (total > 0) showToast('Link a stock diamond first — its carat weight turns a total into a per-carat rate.', 'info');
    calculateSaleMath();
    return;
  }

  const rateField = el('s-SalePerCarat');
  if (rateField) rateField.value = total > 0 ? roundTo2(total / size).toFixed(2) : '';
  calculateSaleMath();
}

function syncSaleTotals() {
  if (!saleStoneSize()) { calculateSaleMath(); return; }
  syncSaleTotalField();
  calculateSaleMath();
}

function setText(id, value) {
  const node = el(id);
  if (node) node.innerText = value;
}

/* ------------------------------------------------------------- suggestions */
const OPTION_DEFAULTS = {
  'color-options': { field: 'Color', list: ['D', 'E', 'F', 'G', 'H'] },
  'clarity-options': { field: 'Quality', list: ['FL', 'IF', 'VVS1', 'VVS2', 'VS1', 'VS2'] },
  'broker-options': { field: 'Broker', list: ['Self', 'Broker'] },
  'type-options': { field: 'Type', list: ['IGI', 'GIA', 'LOOSE', 'OTHER'] }
};

function fillDatalist(id, values) {
  const dl = el(id);
  if (!dl) return;
  dl.innerHTML = values.map(v => '<option value="' + esc(v) + '">').join('');
}

function refreshPurchaseOptions() {
  Object.keys(OPTION_DEFAULTS).forEach(id => {
    const cfg = OPTION_DEFAULTS[id];
    const seen = new Set(cfg.list.map(v => v.toLowerCase()));
    const all = cfg.list.slice();
    globalStore.purchases.forEach(p => {
      const v = String(p[cfg.field] || '').trim();
      if (v && !seen.has(v.toLowerCase())) { seen.add(v.toLowerCase()); all.push(v); }
    });
    fillDatalist(id, all);
  });

  /* operators and countries come from real data so the columns stay consistent */
  const operators = new Map();
  globalStore.purchases.concat(globalStore.sales).forEach(r => {
    const v = String((r && r.User) || '').trim();
    if (v) operators.set(v.toLowerCase(), v);
  });
  if (globalStore.currentUser) operators.set(globalStore.currentUser.toLowerCase(), globalStore.currentUser);
  fillDatalist('operator-options', Array.from(operators.values()).sort());

  const countries = new Map([['usa', 'USA'], ['uk', 'UK'], ['thailand', 'Thailand'], ['dubai', 'Dubai'], ['hongkong', 'Hongkong']]);
  globalStore.sales.forEach(s => {
    const v = String(s.Country || '').trim();
    if (v) countries.set(v.toLowerCase(), v);
  });
  fillDatalist('country-options', Array.from(countries.values()).sort());
}

/* --------------------------------------------------------- table rendering */
function renderAllTables() {
  refreshPurchaseOptions();
  renderPurchaseTable();
  renderStockTable();
  renderSalesTable();

  if (currentReportData.length) {
    const hint = el('report-stale');
    if (hint) hint.classList.add('show');
  }
  currentReportData = [];
}

function renderSkeletons() {
  const skeleton = (cols, rows) => Array.from({ length: rows }, () =>
    `<tr class="skeleton-row"><td colspan="${cols}"><span class="skel-bar"></span></td></tr>`).join('');
  if (el('purchaseTableBody')) el('purchaseTableBody').innerHTML = skeleton(20, 5);
  if (el('saleStockLookupBody')) el('saleStockLookupBody').innerHTML = skeleton(11, 4);
  if (el('saleRegisterTableBody')) el('saleRegisterTableBody').innerHTML = skeleton(13, 3);
}

function emptyState(cols, icon, title, sub, ctaLabel, ctaCall) {
  const cta = ctaLabel
    ? `<button type="button" class="btn btn-primary es-cta" onclick="${ctaCall}">${esc(ctaLabel)}</button>`
    : '';
  return `<tr><td colspan="${cols}"><div class="empty-state">
      <div class="es-icon" aria-hidden="true">${icon}</div>
      <div class="es-title">${esc(title)}</div>
      <div class="es-sub">${esc(sub)}</div>
      ${cta}
    </div></td></tr>`;
}

function markSelectedRow(bodyId, id) {
  clearRowHighlights();
  const body = el(bodyId);
  if (!body) return;
  const row = Array.from(body.querySelectorAll('tr')).find(r => r.dataset.id === String(id));
  if (row) row.classList.add('selected-row');
}

function setPurchaseFilter(filterType) {
  currentPurchaseFilter = filterType;
  ['all', 'stock', 'sold'].forEach(k => {
    const b = el('pur-filter-' + k);
    if (b) b.classList.toggle('active', k === filterType.toLowerCase());
  });
  filterPurchasesTable();
}

function setPurchaseTypeFilter(filterType) {
  currentPurchaseTypeFilter = filterType;
  ['all', 'igi', 'gia', 'loose', 'other'].forEach(k => {
    const b = el('pur-type-' + k);
    if (b) b.classList.toggle('active', k === filterType.toLowerCase());
  });
  filterPurchasesTable();
}

function toggleAdvancedFilter(type){
  const elId = type === 'purchase' ? 'adv-filter-purchase' : 'adv-filter-sales';
  const panel = el(elId);
  if(panel){
    panel.style.display = panel.style.display === 'flex' ? 'none' : 'flex';
  }
}

function applyAdvancedFilters(type){
  if(type === 'purchase'){
    const from = el('pur-adv-from')?.value;
    const to = el('pur-adv-to')?.value;
    advancedFilters.purchase = (from || to) ? {from, to} : null;
    filterPurchasesTable();
  }else if(type === 'sales'){
    const from = el('sales-adv-from')?.value;
    const to = el('sales-adv-to')?.value;
    advancedFilters.sales = (from || to) ? {from, to} : null;
    filterSalesTable();
  }
  showToast('Advanced filters applied');
}

function clearAdvancedFilters(type){
  if(type === 'purchase'){
    if(el('pur-adv-from')) el('pur-adv-from').value = '';
    if(el('pur-adv-to')) el('pur-adv-to').value = '';
    advancedFilters.purchase = null;
    filterPurchasesTable();
  }else if(type === 'sales'){
    if(el('sales-adv-from')) el('sales-adv-from').value = '';
    if(el('sales-adv-to')) el('sales-adv-to').value = '';
    advancedFilters.sales = null;
    filterSalesTable();
  }
  showToast('Filters cleared');
}

function renderPurchaseTable(customList) {
  const list = customList !== undefined ? customList : globalStore.purchases;
  const pBody = el('purchaseTableBody');
  const badge = el('pur-count-badge');
  if (badge) badge.innerText = `${list.length} records`;
  updateInventorySummary(list);
  if (!pBody) return;

  if (list.length === 0) {
    pBody.innerHTML = emptyState(20, '◆', 'No stones in this view',
      'Record an inbound purchase, or clear the filter and search to see everything.',
      'Record a purchase', 'focusPurchaseEntry()');
    markSort('purchases');
    return;
  }

  const rows = applySort(list, 'purchases').map(safeRow).map(row => {
    const isSold = String(row.Status).toUpperCase() === 'SOLD';
    const statusBadge = isSold
      ? `<span class="badge stock-gray">● SOLD</span>`
      : `<span class="badge stock-green">● IN STOCK</span>`;

    const billPct = parseFloat(row.BillPct) || 0;
    const totalUsd = parseFloat(row.DAmount) || 0;
    const baseUsd = totalUsd / (1 + billPct / 100);
    const totalInr = parseFloat(row.IndAmount) || 0;

    return `
      <tr onclick="selectPurchaseRowForEdit(this.dataset.id)" data-id="${esc(row.PurId)}" class="hover-row">
        <td style="font-weight:700; color:var(--brand-600);">${esc(row.PurId)}</td>
        <td>${statusBadge}</td>
        <td style="font-weight:600;" class="font-mono">${esc(row.GIANo)}</td>
        <td>${esc(row.PurchaseDate)}</td>
        <td style="font-weight:600;">${esc(row.FromName)}</td>
        <td><span class="chip">${esc(row.Shape)}</span></td>
        <td style="text-align:right; font-weight:700;" class="tabular-nums">${num(row.Size).toFixed(2)} ct</td>
        <td>${dash(row.Color)}</td>
        <td>${dash(row.Quality)}</td>
        <td style="text-align:right;" class="tabular-nums">${fmtUsd(row.PerCarat)}</td>
        <td style="text-align:right;" class="tabular-nums">${fmtUsd(baseUsd)}</td>
        <td style="text-align:right; font-weight:700; color:var(--brand-700);" class="tabular-nums">${fmtUsd(totalUsd)}</td>
        <td style="text-align:right;" class="tabular-nums">${row.Rate ? '₹' + esc(row.Rate) : '<span class="muted">–</span>'}</td>
        <td style="text-align:right; font-weight:700;" class="tabular-nums">${fmtInr(totalInr)}</td>
        <td>${dash(row.BType)}</td>
        <td class="tabular-nums">${billPct > 0 ? billPct + '%' : '<span class="muted">–</span>'}</td>
        <td class="tabular-nums">${dash(row.PaymentDay)}</td>
        <td>${dash(row.Broker)}</td>
        <td>${dash(row.User)}</td>
        <td>${row.Type ? `<span class="type-badge type-${String(row.Type).toLowerCase().replace(/[^a-z]/g,'')}">${esc(row.Type)}</span>` : '<span class="muted">–</span>'}</td>
      </tr>`;
  }).join('');

  pBody.innerHTML = rows;
  markSort('purchases');
}

function updateInventorySummary(list) {
  const box = el('inv-summary');
  if (!box) return;
  let stones = list.length;
  let carats = 0;
  let cost = 0;
  let inr = 0;
  let inStock = 0;
  list.forEach(p => {
    carats += num(p.Size);
    cost += num(p.DAmount);
    inr += num(p.IndAmount);
    if (String(p.Status).toUpperCase() !== 'SOLD') inStock++;
  });
  box.innerHTML = `
    <span><span class="ts-label">Shown</span> <b>${stones}</b></span>
    <span><span class="ts-label">In stock</span> <b>${inStock}</b></span>
    <span><span class="ts-label">Carats</span> <b>${carats.toFixed(2)}</b></span>
    <span><span class="ts-label">Cost</span> <b>${fmtUsd(cost)}</b></span>
    <span><span class="ts-label">Inbound</span> <b>${fmtInr(inr)}</b></span>`;
}

function getFilteredPurchases() {
  const q = (el('pur-search') && el('pur-search').value || '').toLowerCase().trim();
  const soldPurIds = globalStore.sales.map(s => String(s.PurId).trim());

  let filtered = globalStore.purchases.filter(p => {
    const isSold = soldPurIds.includes(String(p.PurId).trim()) || String(p.Status).toUpperCase() === 'SOLD';
    if (currentPurchaseFilter === 'STOCK' && isSold) return false;
    if (currentPurchaseFilter === 'SOLD' && !isSold) return false;
    if (currentPurchaseTypeFilter && currentPurchaseTypeFilter !== 'ALL') {
      const pType = String(p.Type || 'IGI').toLowerCase();
      const fType = currentPurchaseTypeFilter.toLowerCase();
      if (fType === 'other') {
        if (['igi', 'gia', 'loose'].includes(pType)) return false;
      } else if (pType !== fType) {
        return false;
      }
    }
    if (advancedFilters.purchase) {
      try {
        const pDate = p.PurchaseDate ? new Date(p.PurchaseDate) : null;
        if (pDate) {
          if (advancedFilters.purchase.from) {
            const fromD = new Date(advancedFilters.purchase.from);
            if (pDate < fromD) return false;
          }
          if (advancedFilters.purchase.to) {
            const toD = new Date(advancedFilters.purchase.to);
            toD.setHours(23,59,59,999);
            if (pDate > toD) return false;
          }
        }
      } catch(e) {}
    }
    return true;
  });

  if (q) {
    filtered = filtered.filter(p =>
      String(p.PurId).toLowerCase().includes(q) ||
      String(p.GIANo).toLowerCase().includes(q) ||
      String(p.FromName).toLowerCase().includes(q) ||
      String(p.Shape).toLowerCase().includes(q) ||
      String(p.Color).toLowerCase().includes(q) ||
      String(p.Quality).toLowerCase().includes(q) ||
      String(p.Broker || '').toLowerCase().includes(q) ||
      String(p.User).toLowerCase().includes(q) ||
      String(p.Type || '').toLowerCase().includes(q)
    );
  }
  return filtered;
}

function filterPurchasesTable() {
  const filtered = getFilteredPurchases();
  renderPurchaseTable(filtered);
}

function exportPurchaseCSV(){
  try {
    const data = getFilteredPurchases();
    const headers = ['PurId','Status','GIANo','PurchaseDate','FromName','Shape','Weight_Carat','Color','Clarity','Flor','Cut','Pol','Sym','Measur','LessPct','PerCarat_USD','Cost_USD','FX_Rate','Total_INR','BType','BillPct','PaymentDays','Broker','Operator','Type'];
    const rows = [headers.join(',')];
    data.forEach(p => {
      const isSold = (globalStore.sales||[]).some(s=>String(s.PurId).trim()===String(p.PurId).trim()) || String(p.Status).toUpperCase()==='SOLD';
      const row = [
        p.PurId, isSold ? 'SOLD' : 'IN STOCK', p.GIANo, p.PurchaseDate, p.FromName, p.Shape, p.Size, p.Color, p.Quality, p.Flor, p.Cut, p.Pol, p.Sym, p.Measur, p.Less !== undefined ? p.Less : 0, p.PerCarat, p.DAmount, p.Rate, p.IndAmount, p.BType, p.BillPct || 0, p.PaymentDay, p.Broker, p.User, p.Type || 'IGI'
      ].map(v=>`"${String(v||'').replace(/"/g,'""')}"`).join(',');
      rows.push(row);
    });
    const blob = new Blob([rows.join('\n')], {type:'text/csv;charset=utf-8;'});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'purchases_' + new Date().toISOString().slice(0,10) + '.csv';
    link.click();
    showToast('Purchases exported (' + data.length + ' rows)', 'success');
  } catch(e){ showToast('Export failed'); }
}

function exportPurchaseExcel(){
  try {
    const data = getFilteredPurchases();
    let html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="utf-8"></head><body><table border="1">';
    const headers = ['PurId','Status','GIA / Lot No','Date','Supplier','Shape','Weight (ct)','Color','Clarity','Flor.','Cut','Pol.','Sym.','Measur','Less %','PerCarat ($)','Cost ($)','Rate (₹)','Total (₹)','BType','Bill %','Days','Broker','Operator','Type'];
    html += '<tr>' + headers.map(h => `<th>${h}</th>`).join('') + '</tr>';
    data.forEach(p => {
      const isSold = (globalStore.sales||[]).some(s=>String(s.PurId).trim()===String(p.PurId).trim()) || String(p.Status).toUpperCase()==='SOLD';
      html += '<tr>';
      html += `<td>${p.PurId||''}</td>`;
      html += `<td>${isSold ? 'SOLD' : 'IN STOCK'}</td>`;
      html += `<td>${p.GIANo||''}</td>`;
      html += `<td>${p.PurchaseDate||''}</td>`;
      html += `<td>${p.FromName||''}</td>`;
      html += `<td>${p.Shape||''}</td>`;
      html += `<td>${p.Size||''}</td>`;
      html += `<td>${p.Color||''}</td>`;
      html += `<td>${p.Quality||''}</td>`;
      html += `<td>${p.Flor||''}</td>`;
      html += `<td>${p.Cut||''}</td>`;
      html += `<td>${p.Pol||''}</td>`;
      html += `<td>${p.Sym||''}</td>`;
      html += `<td>${p.Measur||''}</td>`;
      html += `<td>${p.Less !== undefined ? p.Less : 0}</td>`;
      html += `<td>${p.PerCarat||''}</td>`;
      html += `<td>${p.DAmount||''}</td>`;
      html += `<td>${p.Rate||''}</td>`;
      html += `<td>${p.IndAmount||''}</td>`;
      html += `<td>${p.BType||''}</td>`;
      html += `<td>${p.BillPct||0}</td>`;
      html += `<td>${p.PaymentDay||''}</td>`;
      html += `<td>${p.Broker||''}</td>`;
      html += `<td>${p.User||''}</td>`;
      html += `<td>${p.Type||'IGI'}</td>`;
      html += '</tr>';
    });
    html += '</table></body></html>';
    const blob = new Blob([html], {type:'application/vnd.ms-excel'});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'purchases_' + new Date().toISOString().slice(0,10) + '.xls';
    link.click();
    showToast('Purchases Excel exported (' + data.length + ' rows)', 'success');
  } catch(e){ showToast('Export failed'); }
}

function getFilteredSales() {
  const q = (el('sales-search') && el('sales-search').value || '').toLowerCase().trim();
  let filtered = globalStore.sales.filter(s => {
    if (advancedFilters.sales) {
      try {
        const sDate = s.Date ? new Date(s.Date) : null;
        if (sDate) {
          if (advancedFilters.sales.from) {
            const fromD = new Date(advancedFilters.sales.from);
            if (sDate < fromD) return false;
          }
          if (advancedFilters.sales.to) {
            const toD = new Date(advancedFilters.sales.to);
            toD.setHours(23,59,59,999);
            if (sDate > toD) return false;
          }
        }
      } catch(e) {}
    }
    return true;
  });

  if (q) {
    filtered = filtered.filter(s =>
      String(s.SaleId).toLowerCase().includes(q) ||
      String(s.PurId).toLowerCase().includes(q) ||
      String(s.ClientName).toLowerCase().includes(q) ||
      String(s.Country || '').toLowerCase().includes(q) ||
      String(s.SaleRemark || '').toLowerCase().includes(q) ||
      String(s.User).toLowerCase().includes(q)
    );
  }
  return filtered;
}

function exportSalesCSV(){
  try {
    const data = getFilteredSales();
    const headers = ['SaleId','PurId','Date','ClientName','Country','SalePerCarat','SaleDFinal','AdvDAmt','DueAmt','SaleDProfit','ProfitPct','SaleRemark','User'];
    const rows = [headers.join(',')];
    data.forEach(s => {
      const row = headers.map(h=>`"${String(s[h]||'').replace(/"/g,'""')}"`).join(',');
      rows.push(row);
    });
    const blob = new Blob([rows.join('\n')], {type:'text/csv;charset=utf-8;'});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'sales_' + new Date().toISOString().slice(0,10) + '.csv';
    link.click();
    showToast('Sales exported (' + data.length + ' rows)', 'success');
  } catch(e){ showToast('Export failed'); }
}

function exportSalesExcel(){
  try {
    const data = getFilteredSales();
    let html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="utf-8"></head><body><table border="1">';
    html += '<tr><th>SaleId</th><th>PurId</th><th>Date</th><th>Client</th><th>Country</th><th>Sale PerCarat ($)</th><th>Gross Invoice ($)</th><th>Advance ($)</th><th>Due Balance ($)</th><th>Remark</th><th>Operator</th><th>Net Profit ($)</th><th>Margin %</th></tr>';
    data.forEach(s => {
      html += `<tr><td>${s.SaleId}</td><td>${s.PurId}</td><td>${s.Date}</td><td>${s.ClientName}</td><td>${s.Country||''}</td><td>${s.SalePerCarat}</td><td>${s.SaleDFinal}</td><td>${s.AdvDAmt||0}</td><td>${s.DueAmt||0}</td><td>${s.SaleRemark||''}</td><td>${s.User||''}</td><td>${s.SaleDProfit||0}</td><td>${s.ProfitPct||0}%</td></tr>`;
    });
    html += '</table></body></html>';
    const blob = new Blob([html], {type:'application/vnd.ms-excel'});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'sales_' + new Date().toISOString().slice(0,10) + '.xls';
    link.click();
    showToast('Sales Excel exported (' + data.length + ' rows)', 'success');
  } catch(e){ showToast('Export failed'); }
}

function exportReportToExcel(){
  try {
    if(!currentReportData || currentReportData.length === 0){
      showToast('No report data to export — generate the report first.', 'error');
      return;
    }
    let html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="utf-8">';
    html += '<style>table{border-collapse:collapse;font-family:Arial,sans-serif;font-size:11px;} th{background:#1e3a8a;color:#fff;font-weight:bold;padding:6px;border:1px solid #ccc;} .th-sale{background:#065f46;} td{padding:5px;border:1px solid #ddd;}</style>';
    html += '</head><body><table>';
    
    // Group Header
    html += '<tr><th colspan="22" style="background:#2563eb;color:#fff;text-align:center;">INBOUND PURCHASE SPECIFICATIONS</th>';
    html += '<th colspan="12" style="background:#059669;color:#fff;text-align:center;">OUTBOUND SALES DISPATCH & PERFORMANCE</th></tr>';

    const headers = [
      'GIA No','Pur Date','From Name','Size (ct)','Shape','Color','Quality','Flor.','Cut','Pol.','Sym.','Measur','Less %','PerCarat ($)','Cost ($)','Rate (₹)','IndAmount (₹)','BType','Payment Days','Broker','Pur Operator','Type',
      'Sale Date','Client Name','Country','Sale PerCarat ($)','Sale Less %','Sale Final ($)','Advance Paid ($)','Due Amount ($)','Remark','Sale Operator','Net Profit ($)','Profit %'
    ];
    html += '<tr>' + headers.map((h, i) => `<th class="${i >= 22 ? 'th-sale' : ''}">${h}</th>`).join('') + '</tr>';

    applySort(currentReportData, 'report').forEach(r => {
      html += '<tr>';
      html += `<td>${r.GIANo||''}</td>`;
      html += `<td>${r.PurchaseDate||''}</td>`;
      html += `<td>${r.FromName||''}</td>`;
      html += `<td>${r.Size||''}</td>`;
      html += `<td>${r.Shape||''}</td>`;
      html += `<td>${r.Color||''}</td>`;
      html += `<td>${r.Quality||''}</td>`;
      html += `<td>${r.Flor||''}</td>`;
      html += `<td>${r.Cut||''}</td>`;
      html += `<td>${r.Pol||''}</td>`;
      html += `<td>${r.Sym||''}</td>`;
      html += `<td>${r.Measur||''}</td>`;
      html += `<td>${r.Less !== undefined ? r.Less : 0}</td>`;
      html += `<td>${r.PerCarat||''}</td>`;
      html += `<td>${r.DAmount||''}</td>`;
      html += `<td>${r.Rate||''}</td>`;
      html += `<td>${r.IndAmount||''}</td>`;
      html += `<td>${r.BType||''}</td>`;
      html += `<td>${r.PaymentDay||''}</td>`;
      html += `<td>${r.Broker||''}</td>`;
      html += `<td>${r.PurUser||''}</td>`;
      html += `<td>${r.Type||'IGI'}</td>`;
      html += `<td>${r.SaleDate||''}</td>`;
      html += `<td>${r.ClientName||''}</td>`;
      html += `<td>${r.Country||''}</td>`;
      html += `<td>${r.SalePerCarat||''}</td>`;
      html += `<td>${r.SaleLess !== undefined ? r.SaleLess : 0}</td>`;
      html += `<td>${r.SaleDFinal||''}</td>`;
      html += `<td>${r.AdvDAmt||0}</td>`;
      html += `<td>${r.DueAmt||0}</td>`;
      html += `<td>${r.SaleRemark||''}</td>`;
      html += `<td>${r.SaleUser||''}</td>`;
      html += `<td>${r.SaleDProfit||0}</td>`;
      html += `<td>${r.ProfitPct||0}%</td>`;
      html += '</tr>';
    });
    html += '</table></body></html>';

    const blob = new Blob([html], {type:'application/vnd.ms-excel'});
    const link = document.createElement('a');
    const fromDate = el('rep-FromDate').value || 'start';
    const toDate = el('rep-ToDate').value || 'end';
    link.href = URL.createObjectURL(blob);
    link.download = `Diamond_Ledger_Report_${fromDate}_to_${toDate}.xls`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);

    showToast('Report Excel exported (' + currentReportData.length + ' rows)', 'success');
  } catch(e){
    console.error(e);
    showToast('Export failed', 'error');
  }
}

function availableStock(list) {
  const soldPurIds = globalStore.sales.map(s => String(s.PurId).trim());
  return list.filter(p => !(soldPurIds.includes(String(p.PurId).trim()) || String(p.Status).toUpperCase() === 'SOLD'));
}

function renderStockTable(customList) {
  const available = customList !== undefined ? customList : availableStock(globalStore.purchases);
  const sBody = el('saleStockLookupBody');
  const badge = el('stock-count-badge');
  if (badge) badge.innerText = `${available.length} available`;
  if (!sBody) return;

  if (available.length === 0) {
    sBody.innerHTML = emptyState(11, '◇', 'Nothing in stock',
      'Every recorded stone has been sold, or no purchase exists yet.',
      'Record a purchase', "switchTab('purchase-tab')");
    markSort('stock');
    return;
  }

  sBody.innerHTML = applySort(available, 'stock').map(safeRow).map(row => `
      <tr onclick="linkPurchaseToSale(this.dataset.id)" data-id="${esc(row.PurId)}" class="hover-row">
        <td style="font-weight:700; color:var(--emerald-600);">${esc(row.PurId)}</td>
        <td style="font-weight:600;" class="font-mono">${esc(row.GIANo)}</td>
        <td>${dash(row.FromName)}</td>
        <td><span class="chip">${esc(row.Shape)}</span></td>
        <td style="text-align:right; font-weight:700;" class="tabular-nums">${num(row.Size).toFixed(2)} ct</td>
        <td>${dash(row.Color)}</td>
        <td>${dash(row.Quality)}</td>
        <td style="text-align:right;" class="tabular-nums">${fmtUsd(row.PerCarat)}</td>
        <td style="text-align:right; font-weight:700; color:var(--brand-700);" class="tabular-nums">${fmtUsd(row.DAmount)}</td>
        <td style="text-align:right;" class="tabular-nums">${row.Rate ? '₹' + esc(row.Rate) : '<span class="muted">–</span>'}</td>
        <td style="text-align:right; font-weight:700;" class="tabular-nums">${fmtInr(row.IndAmount)}</td>
      </tr>`).join('');
  markSort('stock');
}

function filterStockTable() {
  const q = (el('stock-search') && el('stock-search').value || '').toLowerCase().trim();
  let available = availableStock(globalStore.purchases);
  if (q) {
    available = available.filter(p =>
      String(p.PurId).toLowerCase().includes(q) ||
      String(p.GIANo).toLowerCase().includes(q) ||
      String(p.FromName).toLowerCase().includes(q) ||
      String(p.Shape).toLowerCase().includes(q) ||
      String(p.Color).toLowerCase().includes(q) ||
      String(p.Type || '').toLowerCase().includes(q)
    );
  }
  renderStockTable(available);
}

function renderSalesTable(customList) {
  const list = customList !== undefined ? customList : globalStore.sales;
  const sBody = el('saleRegisterTableBody');
  const badge = el('sales-count-badge');
  if (badge) badge.innerText = `${list.length} sales`;
  if (!sBody) return;

  if (list.length === 0) {
    sBody.innerHTML = emptyState(13, '⇪', 'No dispatches recorded yet',
      'Pick a stone from the stock vault, then record the sale details to see margins here.',
      'Record a sale', "switchTab('sale-tab')");
    markSort('sales');
    return;
  }

  sBody.innerHTML = applySort(list, 'sales').map(safeRow).map(row => {
    const due = parseFloat(row.DueAmt || 0);
    const dueBadge = due > 0
      ? `<span class="badge danger tabular-nums">${fmtUsd(due)}</span>`
      : `<span class="badge stock-green tabular-nums">PAID</span>`;
    const profit = parseFloat(row.SaleDProfit || 0);

    return `
      <tr onclick="selectSaleRowForEdit(this.dataset.id)" data-id="${esc(row.SaleId)}" class="hover-row">
        <td style="font-weight:700; color:var(--emerald-600);">${esc(row.SaleId)}</td>
        <td><a href="#" class="jump-link" onclick="event.stopPropagation();openPurchaseFromSale('${esc(row.PurId)}')" title="Open this purchase">${esc(row.PurId)}</a></td>
        <td>${esc(row.Date)}</td>
        <td style="font-weight:700;">${esc(row.ClientName)}</td>
        <td><span class="chip">${dash(row.Country)}</span></td>
        <td style="text-align:right;" class="tabular-nums">${fmtUsd(row.SalePerCarat)}</td>
        <td style="text-align:right; font-weight:700; color:var(--emerald-700);" class="tabular-nums">${fmtUsd(row.SaleDFinal)}</td>
        <td style="text-align:right; color:var(--emerald-600);" class="tabular-nums">${fmtUsd(row.AdvDAmt)}</td>
        <td style="text-align:right;">${dueBadge}</td>
        <td style="text-align:right;" class="tabular-nums ${moneyClass(profit)}">${fmtUsd(profit)}</td>
        <td style="text-align:right;" class="tabular-nums ${moneyClass(row.ProfitPct)}">${fmtPct(row.ProfitPct)}</td>
        <td style="font-size:11.5px; color:var(--slate-500);">${dash(row.SaleRemark)}</td>
        <td>${dash(row.User)}</td>
      </tr>`;
  }).join('');
  markSort('sales');
}

function filterSalesTable() {
  const filtered = getFilteredSales();
  renderSalesTable(filtered);
}

function openPurchaseFromSale(purId) {
  const match = globalStore.purchases.find(p => String(p.PurId).trim() === String(purId).trim());
  switchTab('purchase-tab');
  setPurchaseFilter('ALL');
  if (el('pur-search')) { el('pur-search').value = ''; el('pur-search').parentNode.classList.remove('has-value'); }
  filterPurchasesTable();
  if (!match) { showToast('Purchase ' + purId + ' is no longer in the sheet.', 'error'); return; }
  selectPurchaseRowForEdit(match.PurId);
  const form = el('purchaseForm');
  if (form) form.scrollIntoView({ behavior: 'smooth', block: 'start' });
  showToast('Showing purchase ' + match.PurId, 'info');
}

/* ------------------------------------------------------- form: validation */
const REQUIRED_FIELDS = {
  purchase: [
    ['p-GIANo', 'GIA / Lot No'], ['p-PurchaseDate', 'Purchase date'], ['p-FromName', 'Supplier'],
    ['p-Size', 'Carat weight'], ['p-Shape', 'Shape'], ['p-Type', 'Type'], ['p-PerCarat', 'Per carat rate'], ['p-User', 'Operator']
  ],
  sale: [
    ['s-Date', 'Sale date'], ['s-ClientName', 'Client name'], ['s-Country', 'Country'],
    ['s-SalePerCarat', 'Sale per carat'], ['s-User', 'Operator']
  ]
};

function clearFieldErrors(scope) {
  (scope || document).querySelectorAll('.field-msg').forEach(n => n.remove());
  (scope || document).querySelectorAll('.field-error').forEach(n => {
    n.classList.remove('field-error');
    if (n.tagName === 'INPUT' || n.tagName === 'SELECT' || n.tagName === 'TEXTAREA') n.removeAttribute('aria-invalid');
  });
  (scope || document).querySelectorAll('.field-error-label').forEach(n => n.classList.remove('field-error-label'));
}

function markFieldError(control, message) {
  control.classList.add('field-error');
  control.setAttribute('aria-invalid', 'true');
  const group = control.closest('.form-group') || control.parentNode;
  const label = group.querySelector('label');
  if (label) label.classList.add('field-error-label');
  if (!group.querySelector('.field-msg')) {
    const msg = document.createElement('div');
    msg.className = 'field-msg';
    msg.innerText = message;
    group.appendChild(msg);
  }
}

function validate(kind) {
  const specs = REQUIRED_FIELDS[kind] || [];
  const missing = [];
  let first = null;
  specs.forEach(([id, label]) => {
    const control = el(id);
    if (!control) return;
    if (!String(control.value || '').trim()) {
      markFieldError(control, label + ' is required');
      missing.push(label);
      if (!first) first = control;
    }
  });
  if (first) { first.focus(); }
  return missing;
}

/* ----------------------------------------------------- remembered defaults */
function rememberLastPurchase(formData) {
  const keep = ['FromName', 'Broker', 'BType', 'BillPct', 'Rate', 'PaymentDay', 'Shape', 'Type'];
  const stored = {};
  keep.forEach(k => { if (formData[k] !== undefined && formData[k] !== '') stored[k] = formData[k]; });
  try { localStorage.setItem('gem_last_purchase', JSON.stringify(stored)); } catch (e) { /* ignore */ }
  if (formData.PurId) {
    try { localStorage.setItem('gem_last_purid', formData.PurId); } catch (e) { /* ignore */ }
  }
}

function applyPurchaseDefaults() {
  let stored = {};
  try { stored = JSON.parse(localStorage.getItem('gem_last_purchase') || '{}'); } catch (e) { stored = {}; }
  const setIfEmpty = (id, value) => {
    const node = el(id);
    if (node && !node.value && value !== undefined && value !== null && value !== '') node.value = value;
  };
  setIfEmpty('p-FromName', stored.FromName);
  setIfEmpty('p-Broker', stored.Broker);
  setIfEmpty('p-Rate', stored.Rate);
  setIfEmpty('p-PaymentDay', stored.PaymentDay);
  setIfEmpty('p-Shape', stored.Shape);
  if (el('p-Type') && !el('p-Type').value) {
    el('p-Type').value = stored.Type || 'IGI';
  } else if (el('p-Type') && stored.Type && !localStorage.getItem('gem_last_purchase_just_set')) {
    // keep existing behavior but ensure default
  }
  if (stored.BType && el('p-BType')) el('p-BType').value = stored.BType;
  if (stored.BType === 'BILL' && stored.BillPct && el('p-BillPct')) el('p-BillPct').value = stored.BillPct;
  onBTypeChange();
}

function focusPurchaseEntry() {
  switchTab('purchase-tab');
  const target = el('p-GIANo');
  if (target) target.focus();
}

/* ---------------------------------------------------------- purchase CRUD */
function selectPurchaseRowForEdit(purId) {
  const match = globalStore.purchases.find(p => String(p.PurId).trim() === String(purId).trim());
  if (!match) return;
  clearFieldErrors(el('purchaseForm'));

  el('p-PurId').value = match.PurId;
  el('p-Status').value = match.Status || 'IN STOCK';
  el('p-GIANo').value = match.GIANo || '';
  el('p-PurchaseDate').value = match.PurchaseDate || '';
  el('p-FromName').value = match.FromName || '';
  el('p-Type').value = match.Type || '';
  el('p-Size').value = match.Size || '';
  el('p-Shape').value = match.Shape || 'ROUND';
  el('p-Color').value = match.Color || '';
  el('p-Quality').value = match.Quality || '';
  el('p-Flor').value = match.Flor || 'NONE';
  el('p-Cut').value = match.Cut || '';
  el('p-Pol').value = match.Pol || '';
  el('p-Sym').value = match.Sym || '';
  el('p-Measur').value = match.Measur || '';
  el('p-Less').value = match.Less !== undefined ? match.Less : '0';
  el('p-PerCarat').value = match.PerCarat || '';
  el('p-Rate').value = match.Rate || '';
  el('p-BType').value = match.BType || 'CASH';
  el('p-PaymentDay').value = match.PaymentDay || '';
  el('p-Broker').value = match.Broker || '';
  el('p-User').value = match.User || globalStore.currentUser;
  el('p-BillPct').value = match.BillPct > 0 ? match.BillPct : '';
  onBTypeChange();

  const badge = el('pur-mode-badge');
  if (badge) { badge.innerText = '✎ Editing ' + match.PurId; badge.className = 'badge modified'; }

  const delBtn = el('purchaseDeleteBtn');
  if (delBtn) { delBtn.style.display = 'inline-flex'; delBtn.disabled = readOnly; }

  const saveText = el('purchaseSaveBtnText');
  if (saveText) saveText.innerText = 'Update Record';

  markSelectedRow('purchaseTableBody', match.PurId);
}

function clearPurchaseForm() {
  el('purchaseForm').reset();
  el('p-PurId').value = '';
  if (el('p-Type')) el('p-Type').value = '';
  el('p-Status').value = 'IN STOCK';
  el('p-PurchaseDate').value = localISO(new Date());
  el('p-Broker').value = '';
  el('p-User').value = globalStore.currentUser;
  clearFieldErrors(el('purchaseForm'));

  setText('p-DAmount-display', fmtUsd(0));
  setText('p-IndAmount-display', fmtInr(0));
  setText('p-BillAmt-display', fmtUsd(0));
  setText('p-BaseAmt-display', fmtUsd(0));

  const badge = el('pur-mode-badge');
  if (badge) { badge.innerText = '✦ New Record'; badge.className = 'badge primary'; }
  const delBtn = el('purchaseDeleteBtn');
  if (delBtn) { delBtn.style.display = 'none'; delBtn.disabled = false; }
  const saveText = el('purchaseSaveBtnText');
  if (saveText) saveText.innerText = 'Save Inbound Record';

  applyPurchaseDefaults();
  if (el('p-Type') && !el('p-Type').value) el('p-Type').value = 'IGI';
  clearRowHighlights();
}

function duplicateLastPurchase() {
  let lastId = '';
  try { lastId = localStorage.getItem('gem_last_purid') || ''; } catch (e) { lastId = ''; }
  const src = globalStore.purchases.find(p => String(p.PurId) === lastId)
    || globalStore.purchases[globalStore.purchases.length - 1];
  if (!src) { showToast('No previous purchase to copy yet.', 'info'); return; }

  clearPurchaseForm();
  el('p-FromName').value = src.FromName || '';
  if (el('p-Type')) el('p-Type').value = src.Type || 'IGI';
  el('p-Shape').value = src.Shape || 'ROUND';
  el('p-Color').value = src.Color || '';
  el('p-Quality').value = src.Quality || '';
  el('p-Flor').value = src.Flor || 'NONE';
  el('p-Cut').value = src.Cut || '';
  el('p-Pol').value = src.Pol || '';
  el('p-Sym').value = src.Sym || '';
  el('p-Less').value = src.Less !== undefined ? src.Less : '0';
  el('p-PerCarat').value = src.PerCarat || '';
  el('p-Rate').value = src.Rate || '';
  el('p-BType').value = src.BType || 'CASH';
  el('p-PaymentDay').value = src.PaymentDay || '';
  el('p-Broker').value = src.Broker || '';
  el('p-BillPct').value = src.BillPct > 0 ? src.BillPct : '';
  onBTypeChange();
  calculatePurchaseMath();
  applyPurchaseDefaults();
  el('p-GIANo').focus();
  showToast('Copied terms from ' + src.PurId + ' — enter the new GIA / Lot No.', 'info');
}

async function handlePurchaseSubmit(e) {
  e.preventDefault();
  if (readOnly) { showToast('This sheet is read-only for your account.', 'error'); return; }

  const form = el('purchaseForm');
  clearFieldErrors(form);
  const missing = validate('purchase');
  if (missing.length) {
    showToast(missing.length === 1 ? missing[0] + ' is required.' : missing.join(', ') + ' are required.', 'error');
    return;
  }

  const formData = Object.fromEntries(new FormData(form));
  if (formData.User) setOperatorState(formData.User);
  const isEdit = !!el('p-PurId').value;

  const btn = el('purchaseSaveBtn');
  const txt = el('purchaseSaveBtnText');
  btn.disabled = true;
  txt.innerText = 'Saving to Sheets…';

  try {
    const res = await callBackend('savePurchaseRecord', [formData]);
    globalStore.purchases = res.purchases || [];
    rememberLastPurchase(formData);
    clearPurchaseForm();
    renderAllTables();
    showToast(`Purchase ${res.purId} ${isEdit ? 'updated' : 'saved'}`, 'success');
    if (globalStore.sales.some(s => String(s.PurId) === String(res.purId))) refreshFromSheet(true);
  } catch (err) {
    showToast(friendlyError(err), 'error');
  } finally {
    btn.disabled = readOnly;
    txt.innerText = isEdit ? 'Update Record' : 'Save Inbound Record';
  }
}

async function deleteCurrentPurchase() {
  const purId = el('p-PurId').value;
  if (!purId) return;
  if (readOnly) { showToast('This sheet is read-only for your account.', 'error'); return; }

  const linked = globalStore.sales.find(s => String(s.PurId).trim() === String(purId).trim());
  if (linked) { showToast('Cannot delete ' + purId + ': it is recorded in sale ' + linked.SaleId + '.', 'error'); return; }

  const ok = await askConfirm({
    title: 'Delete purchase ' + purId,
    body: `This permanently removes <b>${esc(purId)}</b> from the <b>Purchases</b> tab of your Google Sheet. This cannot be undone.`,
    okText: 'Delete permanently'
  });
  if (!ok) return;

  const delBtn = el('purchaseDeleteBtn');
  const originalHTML = delBtn.innerHTML;
  delBtn.disabled = true;
  delBtn.textContent = 'Deleting…';

  try {
    const res = await callBackend('deletePurchaseRecord', [purId]);
    globalStore.purchases = res.purchases || [];
    clearPurchaseForm();
    renderAllTables();
    showToast('Purchase ' + purId + ' deleted', 'success');
  } catch (err) {
    showToast(friendlyError(err), 'error');
  } finally {
    delBtn.disabled = readOnly;
    delBtn.innerHTML = originalHTML;
  }
}

/* -------------------------------------------------------------- sale CRUD */
function linkPurchaseToSale(purId) {
  const match = globalStore.purchases.find(p => String(p.PurId).trim() === String(purId).trim());
  if (!match) return;

  currentSelectedPurchase = match;
  el('s-PurId').value = match.PurId;

  const sm = safeRow(match);
  const displayCard = el('linked-purchase-card');
  if (displayCard) {
    displayCard.className = 'linked-card locked';
    displayCard.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:4px; gap:6px; flex-wrap:wrap;">
        <span style="font-weight:800; font-size:12px; color:var(--brand-700);">&#9670; LOCKED: ${esc(sm.PurId)}</span>
        <div style="display:flex; align-items:center; gap:6px;">
          ${sm.Type ? `<span class="type-badge type-${String(sm.Type).toLowerCase().replace(/[^a-z]/g,'')}">${esc(sm.Type)}</span>` : ''}
          <span class="badge primary font-mono">GIA: ${esc(sm.GIANo) || 'N/A'}</span>
        </div>
      </div>
      <div style="font-size:12px; color:var(--slate-700);">
        <b>${esc(sm.FromName)}</b> &bull; Shape: <b>${esc(sm.Shape) || '-'}</b> &bull; Weight: <b>${num(match.Size).toFixed(2)} ct</b><br>
        Cost: <b>${fmtUsd(match.DAmount)}</b> (${fmtInr(match.IndAmount)}) &bull; FX: <b>${match.Rate ? '₹' + esc(match.Rate) : '–'}</b>
      </div>`;
  }
  markSelectedRow('saleStockLookupBody', match.PurId);
  syncSaleTotalField();
  calculateSaleMath();
}

function selectSaleRowForEdit(saleId) {
  const match = globalStore.sales.find(s => String(s.SaleId).trim() === String(saleId).trim());
  if (!match) return;
  clearFieldErrors(el('saleForm'));

  const linked = globalStore.purchases.find(p => String(p.PurId).trim() === String(match.PurId).trim());
  if (linked) {
    linkPurchaseToSale(linked.PurId);
  } else {
    currentSelectedPurchase = null;
    el('s-PurId').value = match.PurId;
  }

  el('s-SaleId').value = match.SaleId;
  el('s-Date').value = match.Date || '';
  el('s-ClientName').value = match.ClientName || '';
  el('s-Country').value = match.Country || '';
  el('s-SalePerCarat').value = match.SalePerCarat || '';
  el('s-SaleLess').value = match.SaleLess !== undefined ? match.SaleLess : '0';
  el('s-AdvDAmt').value = match.AdvDAmt || '0';
  el('s-SaleRemark').value = match.SaleRemark || '';
  el('s-User').value = match.User || globalStore.currentUser;

  if (linked) { syncSaleTotalField(); calculateSaleMath(); }

  const badge = el('sale-mode-badge');
  if (badge) { badge.innerText = '✎ Editing ' + match.SaleId; badge.className = 'badge modified'; }
  const delBtn = el('saleDeleteBtn');
  if (delBtn) { delBtn.style.display = 'inline-flex'; delBtn.disabled = readOnly; }
  const saveText = el('saleSaveBtnText');
  if (saveText) saveText.innerText = 'Update Sale Dispatch';

  markSelectedRow('saleRegisterTableBody', match.SaleId);
}

function clearSaleForm() {
  el('saleForm').reset();
  el('s-SaleId').value = '';
  el('s-PurId').value = '';
  el('s-Country').value = '';
  el('s-Date').value = localISO(new Date());
  el('s-User').value = globalStore.currentUser;
  currentSelectedPurchase = null;
  clearFieldErrors(el('saleForm'));

  if (el('s-TotalAmount')) el('s-TotalAmount').value = '';
  setText('s-SaleDFinal-display', fmtUsd(0));
  setText('s-SaleDProfit-display', fmtUsd(0));
  setText('s-ProfitPct-display', fmtPct(0));
  setText('s-DueAmt-display', fmtUsd(0));

  const displayCard = el('linked-purchase-card');
  if (displayCard) {
    displayCard.className = 'linked-card';
    displayCard.innerHTML = `
      <div style="display:flex; align-items:center; gap:8px; color:var(--slate-500);">
        <svg fill="none" stroke="currentColor" viewBox="0 0 24 24" style="width:20px; height:20px; flex-shrink:0;"><path stroke-linecap="round" stroke-linejoin="round" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
        <div>
          <div style="font-weight:700; font-size:11px; text-transform:uppercase; letter-spacing:0.04em;">No Asset Linked</div>
          <div style="font-size:11.5px;">Click any open inventory stone from the matrix on the right.</div>
        </div>
      </div>`;
  }

  const badge = el('sale-mode-badge');
  if (badge) { badge.innerText = '✦ New Dispatch'; badge.className = 'badge stock-green'; }
  const delBtn = el('saleDeleteBtn');
  if (delBtn) { delBtn.style.display = 'none'; delBtn.disabled = false; }
  const saveText = el('saleSaveBtnText');
  if (saveText) saveText.innerText = 'Finalize Sale';

  clearRowHighlights();
}

async function handleSaleSubmit(e) {
  e.preventDefault();
  if (readOnly) { showToast('This sheet is read-only for your account.', 'error'); return; }
  if (!el('s-PurId').value) {
    showToast('Choose an open stock diamond from the vault first.', 'error');
    el('stock-search') && el('stock-search').focus();
    return;
  }

  const form = el('saleForm');
  clearFieldErrors(form);
  const missing = validate('sale');
  if (missing.length) {
    showToast(missing.length === 1 ? missing[0] + ' is required.' : missing.join(', ') + ' are required.', 'error');
    return;
  }

  const formData = Object.fromEntries(new FormData(form));
  if (formData.User) setOperatorState(formData.User);
  const isEdit = !!el('s-SaleId').value;

  const btn = el('saleSaveBtn');
  const txt = el('saleSaveBtnText');
  btn.disabled = true;
  txt.innerText = isEdit ? 'Updating…' : 'Dispatching…';

  try {
    const res = await callBackend('saveSaleRecord', [formData]);
    globalStore.sales = res.sales || [];
    globalStore.purchases = res.purchases || [];
    clearSaleForm();
    renderAllTables();
    showToast(`Sale ${res.saleId} ${isEdit ? 'updated' : 'logged'}`, 'success');
  } catch (err) {
    showToast(friendlyError(err), 'error');
  } finally {
    btn.disabled = readOnly;
    txt.innerText = isEdit ? 'Update Sale Dispatch' : 'Finalize Sale';
  }
}

async function deleteCurrentSale() {
  const saleId = el('s-SaleId').value;
  if (!saleId) return;
  if (readOnly) { showToast('This sheet is read-only for your account.', 'error'); return; }

  const ok = await askConfirm({
    title: 'Void sale ' + saleId,
    body: `Sale <b>${esc(saleId)}</b> will be removed from the sheet and its diamond returned to <b>IN STOCK</b>.`,
    okText: 'Void sale',
    primary: false
  });
  if (!ok) return;

  const delBtn = el('saleDeleteBtn');
  const originalHTML = delBtn.innerHTML;
  delBtn.disabled = true;
  delBtn.textContent = 'Cancelling…';

  try {
    const res = await callBackend('deleteSaleRecord', [saleId]);
    globalStore.sales = res.sales || [];
    globalStore.purchases = res.purchases || [];
    clearSaleForm();
    renderAllTables();
    showToast('Sale ' + saleId + ' voided · stone back in stock', 'success');
  } catch (err) {
    showToast(friendlyError(err), 'error');
  } finally {
    delBtn.disabled = readOnly;
    delBtn.innerHTML = originalHTML;
  }
}

/* --------------------------------------------------------------- reporting */
async function fetchDateReport() {
  const fromDate = el('rep-FromDate').value;
  const toDate = el('rep-ToDate').value;
  const rBody = el('reportTableBody');
  if (rBody) rBody.innerHTML = `<tr><td colspan="34"><div class="empty-state"><div class="es-title">Loading the ledger…</div><div class="es-sub">Joining purchases with dispatches for this date range.</div></div></td></tr>`;

  try {
    const data = await callBackend('getSalesReportByDate', [fromDate, toDate]);
    currentReportData = data || [];
    const hint = el('report-stale');
    if (hint) hint.classList.remove('show');
    updateReportKPIs(currentReportData);
    renderReportTable(currentReportData);
  } catch (err) {
    showToast('Report error: ' + friendlyError(err), 'error');
    if (rBody) rBody.innerHTML = `<tr><td colspan="34"><div class="empty-state"><div class="es-title" style="color:var(--rose-600)">Could not load the report</div><div class="es-sub">${esc(friendlyError(err))}</div></div></td></tr>`;
  }
}

function updateReportKPIs(data) {
  let totalStones = data.length;
  let totalCarats = 0;
  let totalCost = 0;
  let totalInr = 0;
  let totalSales = 0;
  let totalProfit = 0;
  let totalDue = 0;

  data.forEach(r => {
    totalCarats += num(r.Size);
    totalCost += num(r.DAmount);
    totalInr += num(r.IndAmount);
    totalSales += num(r.SaleDFinal);
    totalProfit += num(r.SaleDProfit);
    totalDue += num(r.DueAmt);
  });

  const overallMargin = totalSales > 0 ? ((totalProfit / totalSales) * 100) : 0;

  setText('kpi-stones', totalStones);
  setText('kpi-carats', totalCarats.toFixed(2) + ' ct');
  setText('kpi-cost', fmtUsd(totalCost));
  setText('kpi-inr', fmtInr(totalInr));
  setText('kpi-sales', fmtUsd(totalSales));
  setText('kpi-profit', fmtUsd(totalProfit));
  setText('kpi-margin', fmtPct(overallMargin));
  setText('kpi-due', fmtUsd(totalDue));

  const profitNode = el('kpi-profit');
  if (profitNode) profitNode.className = 'kpi-value tabular-nums ' + (totalProfit < 0 ? 'money-neg' : 'money-pos');
  const marginNode = el('kpi-margin');
  if (marginNode) marginNode.className = 'kpi-value tabular-nums ' + (overallMargin < 0 ? 'money-neg' : 'money-pos');

  const badge = el('report-summary-badge');
  if (badge) badge.innerText = `${totalStones} records`;
}

function renderReportTable(data) {
  const rBody = el('reportTableBody');
  if (!rBody) return;

  if (data.length === 0) {
    rBody.innerHTML = emptyState(34, '⌕', 'No transactions in this range',
      'Widen the dates or choose “All Time” to see the whole ledger.',
      'Generate report', 'fetchDateReport()');
    markSort('report');
    return;
  }

  rBody.innerHTML = applySort(data, 'report').map(safeRow).map(r => `
      <tr class="hover-row">
        <td style="font-weight:600;" class="font-mono">${r.PurchaseMissing ? '⚠ missing ' + esc(r.PurId) : esc(r.GIANo)}</td>
        <td>${esc(r.PurchaseDate)}</td>
        <td style="font-weight:600;">${esc(r.FromName)}</td>
        <td style="text-align:right; font-weight:700;" class="tabular-nums">${num(r.Size).toFixed(2)}</td>
        <td><span class="chip">${esc(r.Shape)}</span></td>
        <td>${dash(r.Color)}</td>
        <td>${dash(r.Quality)}</td>
        <td>${dash(r.Flor)}</td>
        <td>${dash(r.Cut)}</td>
        <td>${dash(r.Pol)}</td>
        <td>${dash(r.Sym)}</td>
        <td style="font-size:11px; color:var(--slate-600);">${dash(r.Measur)}</td>
        <td style="text-align:right;" class="tabular-nums">${r.Less !== '' ? esc(r.Less) : '0'}</td>
        <td style="text-align:right;" class="tabular-nums">${fmtUsd(r.PerCarat)}</td>
        <td style="text-align:right; font-weight:700; color:var(--brand-700);" class="tabular-nums">${fmtUsd(r.DAmount)}</td>
        <td style="text-align:right;" class="tabular-nums">${r.Rate ? '₹' + esc(r.Rate) : '<span class="muted">–</span>'}</td>
        <td style="text-align:right; font-weight:600;" class="tabular-nums">${fmtInr(r.IndAmount)}</td>
        <td>${dash(r.BType)}</td>
        <td class="tabular-nums">${dash(r.PaymentDay)}</td>
        <td>${dash(r.Broker)}</td>
        <td>${dash(r.PurUser)}</td>
        <td>${r.Type ? `<span class="type-badge type-${String(r.Type).toLowerCase().replace(/[^a-z]/g,'')}">${esc(r.Type)}</span>` : '<span class="muted">–</span>'}</td>
        <td style="font-weight:600; color:var(--emerald-700);">${esc(r.SaleDate)}</td>
        <td style="font-weight:700;">${esc(r.ClientName)}</td>
        <td><span class="chip">${dash(r.Country)}</span></td>
        <td style="text-align:right;" class="tabular-nums">${fmtUsd(r.SalePerCarat)}</td>
        <td style="text-align:right;" class="tabular-nums">${r.SaleLess !== '' ? esc(r.SaleLess) : '0'}</td>
        <td style="text-align:right; font-weight:700; color:var(--emerald-700);" class="tabular-nums">${fmtUsd(r.SaleDFinal)}</td>
        <td style="text-align:right; color:var(--emerald-600);" class="tabular-nums">${fmtUsd(r.AdvDAmt)}</td>
        <td style="text-align:right;" class="tabular-nums ${num(r.DueAmt) > 0 ? 'money-neg' : ''}">${fmtUsd(r.DueAmt)}</td>
        <td style="font-style:italic; font-size:11px; color:var(--slate-500);">${esc(r.SaleRemark)}</td>
        <td>${dash(r.SaleUser)}</td>
        <td style="text-align:right;" class="tabular-nums ${moneyClass(r.SaleDProfit)}">${fmtUsd(r.SaleDProfit)}</td>
        <td style="text-align:right;" class="tabular-nums ${moneyClass(r.ProfitPct)}">${fmtPct(r.ProfitPct)}</td>
      </tr>`).join('');
  markSort('report');
  syncReportHeaderOffset();
}

function syncReportHeaderOffset() {
  const table = el('tbl-report');
  if (!table) return;
  const groupRow = table.querySelector('thead tr');
  if (!groupRow) return;
  const height = groupRow.getBoundingClientRect().height;
  if (height > 0) document.documentElement.style.setProperty('--report-gh', Math.round(height) + 'px');
}

/* KPIs follow the filtered rows so the numbers never contradict the table. */
function filterReportTable() {
  const q = (el('rep-search') && el('rep-search').value || '').toLowerCase().trim();
  if (!q) {
    updateReportKPIs(currentReportData);
    renderReportTable(currentReportData);
    return;
  }
  const filtered = currentReportData.filter(r =>
    String(r.GIANo).toLowerCase().includes(q) ||
    String(r.FromName).toLowerCase().includes(q) ||
    String(r.ClientName).toLowerCase().includes(q) ||
    String(r.Country || '').toLowerCase().includes(q) ||
    String(r.Broker || '').toLowerCase().includes(q) ||
    String(r.Shape).toLowerCase().includes(q) ||
    String(r.Color).toLowerCase().includes(q) ||
    String(r.Type || '').toLowerCase().includes(q) ||
    String(r.SaleRemark).toLowerCase().includes(q)
  );
  updateReportKPIs(filtered);
  renderReportTable(filtered);
}

function printReport() {
  if (!currentReportData.length) {
    showToast('Generate the report first, then print.', 'info');
    return;
  }
  switchTab('report-tab');
  setTimeout(() => window.print(), 250);
}

/* CSV export keeps raw, unformatted numbers for spreadsheet round-trips. */
function exportReportToCSV() {
  if (currentReportData.length === 0) {
    showToast('No report data to export — generate the report first.', 'error');
    return;
  }

  const line1 = ',,,,,,,,,,,,,,,,,,,,,PURCHASE,,,,,,,,,,,,SALE,,,,,,,,,,,';
  const line2 = 'GIA No,Purchase Date,From Name,Size,Shape,Color,Quality,Flor.,Cut,Pol.,Sym.,Measur,Less,PerCarat,DAmount,Rate,IndAmount,BType,PaymentDay,Broker,Pur User,Type,Date,ClientName,Country,SalePerCarat,SaleLess,SaleDFinal,ADV D AMT,DUE AMT,SaleRemark,Sale User,SaleDProfit,Profit %';

  const escapeCSV = val => {
    if (val === undefined || val === null) return '';
    let str = String(val);
    if (str.includes(',') || str.includes('"') || str.includes('\n')) {
      str = '"' + str.replace(/"/g, '""') + '"';
    }
    return str;
  };

  const rows = applySort(currentReportData, 'report').map(r => [
    r.GIANo, r.PurchaseDate, r.FromName, r.Size, r.Shape, r.Color, r.Quality, r.Flor, r.Cut,
    r.Pol, r.Sym, r.Measur, r.Less, r.PerCarat, r.DAmount, r.Rate, r.IndAmount, r.BType,
    r.PaymentDay, r.Broker, r.PurUser, r.Type || 'IGI', r.SaleDate, r.ClientName, r.Country, r.SalePerCarat,
    r.SaleLess, r.SaleDFinal, r.AdvDAmt, r.DueAmt, r.SaleRemark, r.SaleUser, r.SaleDProfit, r.ProfitPct
  ].map(escapeCSV).join(','));

  const csvContent = [line1, line2, ...rows].join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const fromDate = el('rep-FromDate').value || 'start';
  const toDate = el('rep-ToDate').value || 'end';

  a.href = url;
  a.download = `Diamond_Sales_Report_${fromDate}_to_${toDate}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  showToast('Report downloaded', 'success');
}
