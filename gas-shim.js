/* gas-shim.js - lets the unchanged Code.js backend logic run in the browser,
   reading/writing the user's own Google Sheet through the Google Sheets API. */
(function () {
  const API = 'https://sheets.googleapis.com/v4/spreadsheets/';
  const SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/userinfo.email';
  let token = null, tokenExp = 0, userEmail = '', tokenClient = null;
  let ssId = localStorage.getItem('gem_sheet_id') || '';
  let SS = null, ops = [];
  let loginResolve; const loginPromise = new Promise(r => loginResolve = r);

  /* ---------- Google sign-in ---------- */
  function initClient() {
    if (tokenClient) return true;
    if (!(window.google && google.accounts && google.accounts.oauth2)) return false;
    tokenClient = google.accounts.oauth2.initTokenClient({ client_id: window.APP_CONFIG.CLIENT_ID, scope: SCOPES, callback: () => {} });
    return true;
  }
  function ensureToken() {
    if (token && Date.now() < tokenExp - 60000) return Promise.resolve();
    return new Promise((res, rej) => {
      if (!initClient()) return rej(new Error('Google sign-in is still loading. Please try again.'));
      tokenClient.callback = r => {
        if (r.error) { showLogin(); return rej(new Error('Please sign in with Google.')); }
        token = r.access_token; tokenExp = Date.now() + r.expires_in * 1000; res();
      };
      tokenClient.requestAccessToken({ prompt: token ? '' : 'select_account' });
    });
  }
  function showLogin() { const o = document.getElementById('login-overlay'); if (o) o.style.display = 'flex'; }
  window.gemSignIn = async function (btn) {
    const msg = document.getElementById('login-msg');
    const label = document.getElementById('login-btn-text');
    if (msg) msg.innerText = '';
    if (btn) btn.disabled = true;
    if (label) label.innerText = 'Signing in\u2026';
    try {
      await ensureToken();
      const u = await (await fetch('https://www.googleapis.com/oauth2/v3/userinfo', { headers: { Authorization: 'Bearer ' + token } })).json();
      userEmail = u.email || '';
      document.getElementById('login-overlay').style.display = 'none';
      loginResolve();
    } catch (e) { if (msg) msg.innerText = e.message; }
    if (btn) btn.disabled = false;
    if (label) label.innerText = 'Continue with Google';
  };

  /* ---------- Sheets API helper ---------- */
  async function api(path, opts) {
    await ensureToken();
    const r = await fetch(API + path, Object.assign({ headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' } }, opts || {}));
    if (!r.ok) {
      let m = ''; try { m = (await r.json()).error.message; } catch (e) {}
      const err = new Error(m || ('Google error ' + r.status)); err.status = r.status; throw err;
    }
    return r.json();
  }

  /* ---------- in-memory spreadsheet (mirrors the Apps Script API used by Code.js) ---------- */
  const colName = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const norm = v => (typeof v === 'string' && /^-?(0|[1-9]\d{0,14})(\.\d+)?$/.test(v)) ? Number(v) : v;

  function makeSheet(name, id, rows) {
    const sh = { name, id, rows };
    const W = () => sh.rows.reduce((m, r) => Math.max(m, r.length), 0);
    const A1 = (r, c, nr, nc) => `'${name}'!${colName(c)}${r}:${colName(c + nc - 1)}${r + nr - 1}`;
    const put = (r, c, vals) => {
      vals = vals.map(row => row.map(norm));
      vals.forEach((row, i) => {
        while (sh.rows.length < r + i) sh.rows.push([]);
        row.forEach((v, j) => { const t = sh.rows[r - 1 + i]; while (t.length < c + j) t.push(''); t[c - 1 + j] = v; });
      });
      const range = A1(r, c, vals.length, vals[0].length);
      ops.push(() => api(`${ssId}/values/${encodeURIComponent(range)}?valueInputOption=RAW`, { method: 'PUT', body: JSON.stringify({ values: vals }) }));
    };
    const batch = req => ops.push(() => api(`${ssId}:batchUpdate`, { method: 'POST', body: JSON.stringify({ requests: [req(sh.id)] }) }));
    sh.getDataRange = () => ({ getValues: () => { const w = W(); return sh.rows.length ? sh.rows.map(r => { const x = r.slice(); while (x.length < w) x.push(''); return x; }) : [['']]; } });
    sh.getLastRow = () => sh.rows.length;
    sh.getLastColumn = () => W();
    sh.getRange = (r, c, nr, nc) => {
      nr = nr || 1; nc = nc || 1;
      return {
        getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => { const v = (sh.rows[r - 1 + i] || [])[c - 1 + j]; return v === undefined ? '' : v; })),
        setValue: v => put(r, c, Array.from({ length: nr }, () => Array(nc).fill(v))),
        setValues: v => put(r, c, v)
      };
    };
    sh.appendRow = vals => put(sh.rows.length + 1, 1, [vals]);
    sh.deleteRow = i => { sh.rows.splice(i - 1, 1); batch(id => ({ deleteDimension: { range: { sheetId: id, dimension: 'ROWS', startIndex: i - 1, endIndex: i } } })); };
    sh.clear = () => { sh.rows.length = 0; ops.push(() => api(`${ssId}/values/${encodeURIComponent("'" + name + "'")}:clear`, { method: 'POST', body: '{}' })); };
    sh.setFrozenRows = n => batch(id => ({ updateSheetProperties: { properties: { sheetId: id, gridProperties: { frozenRowCount: n } }, fields: 'gridProperties.frozenRowCount' } }));
    sh.insertColumnBefore = c => { sh.rows.forEach(r => { if (r.length >= c - 1) r.splice(c - 1, 0, ''); }); batch(id => ({ insertDimension: { range: { sheetId: id, dimension: 'COLUMNS', startIndex: c - 1, endIndex: c }, inheritFromBefore: false } })); };
    sh.insertColumnAfter = c => { sh.rows.forEach(r => { if (r.length >= c) r.splice(c, 0, ''); }); batch(id => ({ insertDimension: { range: { sheetId: id, dimension: 'COLUMNS', startIndex: c, endIndex: c + 1 }, inheritFromBefore: true } })); };
    return sh;
  }

  async function loadSS() {
    const meta = await api(ssId + '?fields=spreadsheetUrl,sheets.properties(sheetId,title)');
    const props = meta.sheets.map(s => s.properties);
    const byName = {};
    if (props.length) {
      const q = props.map(p => 'ranges=' + encodeURIComponent("'" + p.title + "'")).join('&');
      const v = await api(ssId + '/values:batchGet?' + q + '&valueRenderOption=UNFORMATTED_VALUE');
      v.valueRanges.forEach((vr, i) => {
        const rows = (vr.values || []).map(r => r.slice());
        // turn date serial numbers into Date objects (like Apps Script does)
        if (rows.length) rows[0].forEach((h, c) => {
          if (/date$/i.test(String(h).replace(/[^a-zA-Z]/g, ''))) for (let r = 1; r < rows.length; r++) {
            const x = rows[r][c];
            if (typeof x === 'number') { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(x) * 86400000); rows[r][c] = new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()); }
          }
        });
        byName[props[i].title] = makeSheet(props[i].title, props[i].sheetId, rows);
      });
    }
    SS = {
      getUrl: () => meta.spreadsheetUrl || ('https://docs.google.com/spreadsheets/d/' + ssId + '/edit'),
      getSheetByName: n => byName[n] || null,
      insertSheet: n => {
        const sh = makeSheet(n, null, []); byName[n] = sh;
        ops.push(async () => { const r = await api(`${ssId}:batchUpdate`, { method: 'POST', body: JSON.stringify({ requests: [{ addSheet: { properties: { title: n } } }] }) }); sh.id = r.replies[0].addSheet.properties.sheetId; });
        return sh;
      }
    };
  }

  /* ---------- Apps Script globals used by Code.js ---------- */
  const pad = n => String(n).padStart(2, '0');
  window.SpreadsheetApp = {
    getActiveSpreadsheet: () => SS,
    openById: id => { if (id === ssId && SS) return SS; throw new Error('Cannot open that sheet.'); },
    create: () => { throw new Error('No Google Sheet linked. Click "Connect Sheet" and paste your sheet link.'); }
  };
  window.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
  window.PropertiesService = { getUserProperties: () => ({ getProperty: k => localStorage.getItem('gem_prop_' + k), setProperty: (k, v) => localStorage.setItem('gem_prop_' + k, v) }) };
  window.Session = { getActiveUser: () => ({ getEmail: () => userEmail }), getScriptTimeZone: () => Intl.DateTimeFormat().resolvedOptions().timeZone };
  window.Utilities = { formatDate: d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` };
  window.Logger = { log: m => console.log(m) };

  /* ---------- google.script.run replacement ---------- */
  const READONLY = ['getInitialData', 'getSalesReportByDate'];
  const parseId = t => { t = String(t || '').trim(); const m = t.match(/\/d\/([a-zA-Z0-9-_]+)/); return m ? m[1] : t; };
  let chain = Promise.resolve();

  async function invoke(name, args) {
    if (!token) { showLogin(); await loginPromise; }
    const prev = ssId;
    if (name === 'setConnectedSpreadsheet') ssId = parseId(args[0]);
    try {
      if (!ssId) throw new Error('No Google Sheet linked. Click "Connect Sheet" and paste your sheet link.');
      try { await loadSS(); }
      catch (e) { throw new Error(e.status === 403 || e.status === 404 ? 'Cannot open this Google Sheet. Check the link and make sure the owner shared it with your Google account.' : e.message); }
      ops = [];
      let result;
      try { result = window[name].apply(null, args); } catch (e) { ops = []; throw e; }
      const todo = ops; ops = [];
      try { for (const op of todo) await op(); }
      catch (e) {
        if (!(READONLY.includes(name) && e.status === 403)) throw (e.status === 403 ? new Error('View-only access: you do not have permission to edit this Google Sheet.') : e);
      }
      localStorage.setItem('gem_sheet_id', ssId);
      return result;
    } catch (e) { ssId = prev; throw e; }
  }

  function makeRun() {
    const h = {};
    const proxy = new Proxy({}, { get(_, name) {
      if (name === 'withSuccessHandler') return fn => { h.s = fn; return proxy; };
      if (name === 'withFailureHandler') return fn => { h.f = fn; return proxy; };
      return (...args) => { chain = chain.then(() => invoke(name, args)).then(r => h.s && h.s(r), e => { console.error(e); h.f ? h.f(e) : alert(e.message); }); };
    } });
    return proxy;
  }
  window.google = window.google || {};
  window.google.script = {};
  Object.defineProperty(window.google.script, 'run', { get: makeRun });
})();
