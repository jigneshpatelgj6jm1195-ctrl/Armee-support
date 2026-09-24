const assert = require('assert/strict');
const crypto = require('crypto');
const fs = require('fs');
const vm = require('vm');
const zlib = require('zlib');

const backendSource = fs.readFileSync('google_apps_script_backend.js', 'utf8');
const legacySource = fs.readFileSync(process.env.LEGACY_BACKEND || 'server_backups/pre-perf-fix-20260924-1048/google_apps_script_backend.js', 'utf8');
const adminSource = fs.readFileSync('admin.html', 'utf8');
const indexSource = fs.readFileSync('index.html', 'utf8');

function range(sheet, row, col, nr = 1, nc = 1) {
  return new Proxy({
    getValues() {
      return Array.from({ length: nr }, (_, r) =>
        Array.from({ length: nc }, (_, c) => sheet.rows[row - 1 + r]?.[col - 1 + c] ?? ''));
    },
    getValue() { return this.getValues()[0][0]; },
    getFormulas() {
      return this.getValues().map(r => r.map(v => typeof v === 'string' && v.startsWith('=') ? v : ''));
    },
    setValues(values) {
      for (let r = 0; r < nr; r++) for (let c = 0; c < nc; c++) {
        sheet.rows[row - 1 + r] ??= [];
        sheet.rows[row - 1 + r][col - 1 + c] = values[r][c];
      }
      return this;
    },
    setValue(value) { return this.setValues([[value]]); },
    createTextFinder(text) {
      const values = this.getValues();
      return {
        matchEntireCell() { return this; },
        matchCase() { return this; },
        findNext() {
          for (let r = 0; r < nr; r++) for (let c = 0; c < nc; c++) {
            if (String(values[r][c]) === String(text)) return { getRow: () => row + r, getColumn: () => col + c };
          }
          return null;
        },
      };
    },
  }, { get(target, prop) { return prop in target ? target[prop] : function () { return this; }; } });
}

function makeSheet(name, rows) {
  return {
    name,
    rows: rows.map(r => r.slice()),
    getName() { return name; },
    getLastRow() { return this.rows.length; },
    getLastColumn() { return Math.max(0, ...this.rows.map(r => r.length)); },
    getRange(...args) { return range(this, ...args); },
    getDataRange() { return range(this, 1, 1, this.getLastRow(), this.getLastColumn()); },
    appendRow(row) { this.rows.push(row.slice()); },
    deleteRow(n) { this.rows.splice(n - 1, 1); },
    deleteRows(n, count) { this.rows.splice(n - 1, count); },
    clear() { this.rows = []; },
    setFrozenRows() {},
    setColumnWidth() {},
    getMaxRows() { return Math.max(1000, this.rows.length); },
    insertColumnsAfter() {},
  };
}

function makeSpreadsheet(sheets) {
  return {
    getSheetByName(name) { return sheets.find(s => s.name === name) || null; },
    getSheets() { return sheets; },
    insertSheet(name) { const s = makeSheet(name, []); sheets.push(s); return s; },
  };
}

function bytes(value) {
  if (typeof value === 'string') return Buffer.from(value, 'utf8');
  return Buffer.from(Array.from(value || [], b => (Number(b) + 256) % 256));
}

function makeContext(legacy) {
  const properties = new Map([['SESSION_SECRET', 'synthetic-test-secret'], ['IMPORT_KEY', 'synthetic-import-key']]);
  let uuidCounter = 0;
  const cacheStore = new Map();
  const context = vm.createContext({
    __legacy: !!legacy,
    console: { log() {}, warn() {}, error() {} },
    Logger: { log() {} },
    ContentService: {
      MimeType: { JSON: 'json' },
      createTextOutput(text) { return { text, setMimeType() { return this; } }; },
    },
    LockService: { getScriptLock() { return { waitLock() {}, tryLock() { return true; }, releaseLock() {} }; } },
    SpreadsheetApp: { flush() {} },
    CacheService: { getScriptCache() { const m = cacheStore; return {
      get: k => (m.has(k) ? m.get(k) : null), put: (k, v) => m.set(k, v),
      getAll: ks => Object.fromEntries(ks.map(k => [k, m.has(k) ? m.get(k) : null])),
      removeAll: ks => ks.forEach(k => m.delete(k)) }; } },
    PropertiesService: {
      getScriptProperties() {
        return { getProperty: k => properties.get(k) || null, setProperty: (k, v) => properties.set(k, v) };
      },
    },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' },
      Charset: { UTF_8: 'utf8' },
      computeDigest(_alg, value) { return Array.from(crypto.createHash('sha256').update(String(value)).digest()); },
      computeHmacSha256Signature(value, secret) {
        return Array.from(crypto.createHmac('sha256', String(secret)).update(String(value)).digest());
      },
      base64EncodeWebSafe(value) { return bytes(value).toString('base64url'); },
      base64DecodeWebSafe(value) { return Array.from(Buffer.from(String(value), 'base64url')); },
      newBlob(value) { return { getDataAsString: () => bytes(value).toString('utf8'), getBytes: () => Array.from(bytes(value)) }; },
      gzip(blob) { const out = Array.from(zlib.gzipSync(Buffer.from(blob.getBytes()))); return { getBytes: () => out }; },
      ungzip(blob) { const text = zlib.gunzipSync(Buffer.from(blob.getBytes())).toString('utf8'); return { getDataAsString: () => text }; },
      base64Encode(value) { return bytes(value).toString('base64'); },
      base64Decode(value) { return Array.from(Buffer.from(String(value), 'base64')); },
      getUuid() { uuidCounter += 1; return `00000000-0000-4000-8000-${String(uuidCounter).padStart(12, '0')}`; },
      formatDate(value, timeZone) {
        const parts = new Intl.DateTimeFormat('en-CA', {
          timeZone, year: 'numeric', month: '2-digit', day: '2-digit'
        }).formatToParts(new Date(value));
        const part = type => parts.find(item => item.type === type).value;
        return `${part('year')}-${part('month')}-${part('day')}`;
      },
    },
  });
  vm.runInContext(context.__legacy ? legacySource : backendSource, context);
  return context;
}

function extractBlock(source, marker) {
  const start = source.indexOf(marker);
  assert.notEqual(start, -1, `Could not find ${marker}`);
  const brace = source.indexOf('{', start);
  let depth = 0;
  let quote = '';
  let escaped = false;
  for (let i = brace; i < source.length; i++) {
    const ch = source[i];
    if (quote) {
      if (escaped) escaped = false;
      else if (ch === '\\') escaped = true;
      else if (ch === quote) quote = '';
      continue;
    }
    if (ch === '"' || ch === "'" || ch === '`') { quote = ch; continue; }
    if (ch === '{') depth++;
    if (ch === '}' && --depth === 0) return source.slice(start, i + 1);
  }
  throw new Error(`Unclosed block for ${marker}`);
}

const results = [];
async function test(name, fn) {
  try { await fn(); results.push({ name, status: 'pass' }); }
  catch (error) { results.push({ name, status: 'fail', error: error.stack || error.message }); }
}
function clone(x) { return JSON.parse(JSON.stringify(x)); }
function withLockSpy(c) {
  const state = { held: false, events: [] };
  c.LockService = { getScriptLock() { return {
    waitLock() { state.held = true; state.events.push('lock'); },
    tryLock() { state.held = true; state.events.push('lock'); return true; },
    releaseLock() { state.held = false; state.events.push('unlock'); } }; } };
  return state;
}
function complaintRow(headers, o) {
  const r = Array(headers.length).fill('');
  Object.entries(o).forEach(([k, v]) => { r[headers.indexOf(k)] = v; });
  return r;
}
function masterRow(o) { const r = Array(23).fill(''); Object.entries(o).forEach(([i, v]) => { r[Number(i)] = v; }); return r; }

(async () => {
  await test('SR No. never re-uses numbers after archived rows are deleted', () => {
    const c = makeContext();
    const H = c.HEADERS || vm.runInContext('HEADERS', c);
    const active = makeSheet('Complaints', [H,
      complaintRow(H, { 'SR No.': 101, 'Case ID': 'CASE-A', 'Serial Number': 'S1' }),
      complaintRow(H, { 'SR No.': 102, 'Case ID': 'CASE-B', 'Serial Number': 'S2' })]);
    const archive = makeSheet('Archive_2026_06', [H,
      ...Array.from({ length: 100 }, (_, i) => complaintRow(H, { 'SR No.': i + 1, 'Case ID': 'CASE-OLD-' + i }))]);
    const log = makeSheet('SubmissionLog', [vm.runInContext('SUBMISSION_LOG_HEADERS', c)]);
    const db = makeSpreadsheet([active, archive, log]);
    c.getPhotoFolder = () => ({}); c.formatLastRow = () => {};
    const a = JSON.parse(c.handleSubmitComplaint(db, { serialNumber: 'NEW-1', submissionId: 'SUB-A', photos: [] }).text);
    assert.equal(a.srNo, 103); // legacy formula would have produced 3 (last row)
    active.rows.splice(1, 2); // admin archives + deletes the two older active rows
    const b = JSON.parse(c.handleSubmitComplaint(db, { serialNumber: 'NEW-2', submissionId: 'SUB-B', photos: [] }).text);
    assert.equal(b.srNo, 104);
    const legacy = makeContext(true);
    const la = makeSheet('Complaints', [H, complaintRow(H, { 'SR No.': 1, 'Case ID': 'X' })]);
    const ldb = makeSpreadsheet([la, makeSheet('SubmissionLog', [vm.runInContext('SUBMISSION_LOG_HEADERS', legacy)])]);
    legacy.getPhotoFolder = () => ({}); legacy.formatLastRow = () => {}; legacy.syncSchoolComplaintMasterStatus = () => {};
    la.rows.splice(1, 1);
    const lo = JSON.parse(legacy.handleSubmitComplaint(ldb, { serialNumber: 'Q', submissionId: 'L1', photos: [] }).text);
    assert.equal(lo.srNo, 1, 'legacy behaviour re-used SR 1 after the delete (documents the bug)');
  });

  await test('photo uploads run outside the script lock; append runs inside it', () => {
    const c = makeContext();
    const spy = withLockSpy(c);
    const H = vm.runInContext('HEADERS', c);
    const active = makeSheet('Complaints', [H]);
    const db = makeSpreadsheet([active, makeSheet('SubmissionLog', [vm.runInContext('SUBMISSION_LOG_HEADERS', c)])]);
    const uploadsWhileLocked = [];
    c.getPhotoFolder = () => ({});
    c.uploadPhotoToDrive = () => { uploadsWhileLocked.push(spy.held); return { viewUrl: 'https://lh3.googleusercontent.com/d/x', openUrl: 'https://drive.google.com/file/d/x/view', name: 'x' }; };
    const append = active.appendRow.bind(active); let appendLocked = null;
    active.appendRow = r => { appendLocked = spy.held; append(r); };
    c.formatLastRow = () => {};
    const out = JSON.parse(c.handleSubmitComplaint(db, { serialNumber: 'SN1', submissionId: 'S-1', photos: ['a', 'b'], serialPhoto: 'c' }).text);
    assert.equal(out.status, 'ok');
    assert.deepEqual(uploadsWhileLocked, [false, false, false]);
    assert.equal(appendLocked, true);
    assert.equal(spy.held, false, 'lock released');
    assert.match(String(active.rows[1][27]), /^=IMAGE\("https:\/\/lh3/);
  });

  await test('retry of the same submission is ignored without re-appending', () => {
    const c = makeContext();
    const H = vm.runInContext('HEADERS', c);
    const active = makeSheet('Complaints', [H]);
    const db = makeSpreadsheet([active, makeSheet('SubmissionLog', [vm.runInContext('SUBMISSION_LOG_HEADERS', c)])]);
    c.getPhotoFolder = () => ({}); c.formatLastRow = () => {};
    const p = { serialNumber: 'SN9', submissionId: 'DUP-1', photos: [] };
    const first = JSON.parse(c.handleSubmitComplaint(db, p).text);
    const second = JSON.parse(c.handleSubmitComplaint(db, p).text);
    assert.equal(second.duplicateIgnored, true);
    assert.equal(second.caseId, first.caseId);
    assert.equal(active.rows.length, 2);
  });

  await test('a busy lock returns a retryable error and writes nothing', () => {
    const c = makeContext();
    c.LockService = { getScriptLock() { return { tryLock() { return false; }, waitLock() { throw new Error('busy'); }, releaseLock() {} }; } };
    const H = vm.runInContext('HEADERS', c);
    const active = makeSheet('Complaints', [H]);
    const db = makeSpreadsheet([active, makeSheet('SubmissionLog', [vm.runInContext('SUBMISSION_LOG_HEADERS', c)])]);
    c.getPhotoFolder = () => ({}); c.formatLastRow = () => {};
    const out = JSON.parse(c.handleSubmitComplaint(db, { serialNumber: 'SN2', submissionId: 'B-1', photos: [] }).text);
    assert.equal(out.retryable, true);
    assert.equal(active.rows.length, 1);
  });

  await test('targeted school-master sync equals the full sync for the submitted serial', () => {
    const scenarios = [
      { status: 'Open', part: '' }, { status: 'Open', part: 'HDD' },
      { status: 'Part Request', part: 'Motherboard' }, { status: 'Closed', part: '' },
    ];
    for (const sc of scenarios) {
      for (const acer of ['', 'Awaiting Spares']) {
        const base = [Array(23).fill('H'),
          masterRow({ 0: 1, 1: 'ICT', 2: 'D1', 6: 'School A', 11: 'IFP', 13: 'SER-1', 16: '', 17: '', 22: acer }),
          masterRow({ 0: 2, 1: 'GK', 2: 'D2', 6: 'School B', 11: 'CPU', 13: 'other', 16: 'Open' }),
          masterRow({ 0: 3, 1: 'GK', 2: 'D3', 6: 'School C', 11: 'IFP', 13: ' ser-1 ', 16: 'Open', 17: 'x' })];
        const c1 = makeContext(); const H = vm.runInContext('HEADERS', c1);
        const comp = makeSheet('Complaints', [H, complaintRow(H, { 'Serial Number': 'SER-1', 'Submitted At': '2026-09-24T05:00:00Z', 'Status': sc.status, 'Suspected Part': sc.part, 'Case ID': 'C1' })]);
        const mFull = makeSheet('SchoolComplaintMaster', clone(base));
        c1.syncSchoolComplaintMasterStatus(makeSpreadsheet([comp, mFull]));
        const c2 = makeContext();
        const mT = makeSheet('SchoolComplaintMaster', clone(base));
        c2.syncSchoolComplaintMasterForSerial_(makeSpreadsheet([mT]), 'SER-1', sc.status, sc.part);
        const legacy = makeContext(true);
        const mL = makeSheet('SchoolComplaintMaster', clone(base));
        legacy.syncSchoolComplaintMasterStatus(makeSpreadsheet([makeSheet('Complaints', clone(comp.rows)), mL]));
        // rows carrying the submitted serial must match the full (new and legacy) sync
        for (const r of [1, 3]) {
          assert.deepEqual(mT.rows[r].slice(0, 19), mFull.rows[r].slice(0, 19), `row ${r} ${JSON.stringify(sc)} acer=${acer}`);
          assert.deepEqual(mFull.rows[r].slice(0, 19), mL.rows[r].slice(0, 19), 'refactored full sync == legacy');
        }
        assert.deepEqual(mFull.rows[2], mL.rows[2]);
      }
    }
  });

  function deptFixture(c) {
    const D = vm.runInContext('DEPT_HEADERS', c), R = vm.runInContext('RES_HEADERS', c);
    const row = (school, ticket) => { const r = Array(D.length).fill(''); r[0] = 'Dist'; r[7] = school; r[8] = 'Sch ' + school; r[9] = ticket; r[21] = '2026-09-01'; return r; };
    const dept = makeSheet('DepartmentComplaints', [D, row('111', 'T1'), row('222', 'T2'), row(' 111 ', 'T3'), row('111', 'T4'), row('333', 'T5')]);
    const res = makeSheet('DepartmentResolutions', [R,
      ['T4', 'Closed', ...Array(R.length - 2).fill('')], ['T3', 'Part Request', 'x', ...Array(R.length - 3).fill('')]]);
    return makeSpreadsheet([dept, res]);
  }
  await test('department lookup for a school returns the same tickets as before', () => {
    const c = makeContext(), legacy = makeContext(true);
    for (const dise of ['111', '222', '999', '']) {
      const a = c.getDepartmentComplaintsForSchool(deptFixture(c), dise).map(t => { const x = { ...t }; delete x.portalToken; return x; });
      const b = legacy.getDepartmentComplaintsForSchool(deptFixture(legacy), dise).map(t => { const x = { ...t }; delete x.portalToken; return x; });
      assert.deepEqual(clone(a), clone(b), 'dise ' + dise);
    }
    assert.equal(c.getDepartmentComplaintsForSchool(deptFixture(c), '111').length, 2);
  });

  await test('school complaint lookup returns the same rows as before', () => {
    const base = [Array(21).fill('H'),
      masterRow({ 0: 1, 2: 'D1', 6: 'Alpha School', 13: 'S1', 16: '' }),
      masterRow({ 0: 2, 2: 'D1', 6: 'Alpha School', 13: 'S2', 16: 'Open' }),
      masterRow({ 0: 3, 2: 'D1', 6: 'Alpha School', 13: 'S3', 16: 'Closed' }),
      masterRow({ 0: 4, 2: '', 6: 'Beta School', 13: 'S4', 16: '' }),
      masterRow({ 0: 5, 2: '', 6: 'Beta School No 2', 13: 'S5', 16: 'open' }),
      masterRow({ 0: 6, 2: 'D9', 6: 'Beta School', 13: 'S6', 16: '' })];
    const c = makeContext(), legacy = makeContext(true);
    for (const [d, n] of [['D1', 'Alpha School'], ['', 'Beta School'], ['D9', 'Beta School'], ['D404', ''], ['', '']]) {
      const strip = l => clone(l).map(x => { delete x.portalToken; return x; });
      const a = strip(c.getSchoolComplaints(makeSpreadsheet([makeSheet('SchoolComplaintMaster', clone(base))]), d, n));
      const b = strip(legacy.getSchoolComplaints(makeSpreadsheet([makeSheet('SchoolComplaintMaster', clone(base))]), d, n));
      assert.deepEqual(a, b, d + '/' + n);
    }
  });

  await test('duplicate-serial check gives the same answer as before', () => {
    const c = makeContext(), legacy = makeContext(true);
    const H = vm.runInContext('HEADERS', c);
    const now = new Date();
    const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 15).toISOString();
    const rows = [H,
      complaintRow(H, { 'Serial Number': ' abc123 ', 'Submitted At': now.toISOString() }),
      complaintRow(H, { 'Serial Number': 'OLD1', 'Submitted At': lastMonth })];
    for (const s of ['ABC123', 'OLD1', 'NONE']) {
      const a = c.checkDuplicateSerial(makeSpreadsheet([makeSheet('Complaints', clone(rows))]), s);
      const b = legacy.checkDuplicateSerial(makeSpreadsheet([makeSheet('Complaints', clone(rows))]), s);
      assert.deepEqual(clone(a), clone(b), s);
    }
  });

  await test('cached GETs do not open the workbook; health and misses still do', () => {
    const c = makeContext();
    let opens = 0;
    const master = makeSheet('MasterData', [[JSON.stringify({ _v: '42', equipment: [{ name: 'CPU', status: 'active' }], users: [], accessUsers: [] })]]);
    c.SpreadsheetApp = { flush() {}, openById() { opens++; return makeSpreadsheet([master]); } };
    const get = a => JSON.parse(c.doGet({ parameter: { action: a } }).text);
    assert.equal(get('get_master_version').v, '42'); assert.equal(opens, 1, 'first call reads the sheet');
    assert.equal(get('get_master_version').v, '42'); assert.equal(opens, 1, 'second call served from cache');
    assert.equal(get('get_master').equipment.length, 1); assert.equal(opens, 1);
    get('get_complaints'); assert.equal(opens, 1, 'auth failure never opens the workbook');
    assert.equal(get('health').status, 'ok'); assert.equal(opens, 2, 'health still probes the workbook');
  });

  await test('a portal write retires the cached master data', () => {
    const c = makeContext();
    const gen1 = c.cacheGeneration_(c.MASTER_CACHE_GEN_KEY || vm.runInContext('MASTER_CACHE_GEN_KEY', c));
    c.SpreadsheetApp = { flush() {}, openById() { return makeSpreadsheet([]); } };
    c.doPost({ postData: { contents: JSON.stringify({ action: 'nonexistent_action' }) } });
    const gen2 = c.cacheGeneration_(vm.runInContext('MASTER_CACHE_GEN_KEY', c));
    assert.notEqual(gen1, gen2);
  });

  await test('submission invalidates complaint caches but keeps the department cache', () => {
    const c = makeContext();
    const H = vm.runInContext('HEADERS', c);
    const db = makeSpreadsheet([makeSheet('Complaints', [H]), makeSheet('SubmissionLog', [vm.runInContext('SUBMISSION_LOG_HEADERS', c)])]);
    c.SpreadsheetApp = { flush() {}, openById() { return db; } };
    c.getPhotoFolder = () => ({}); c.formatLastRow = () => {};
    const key = vm.runInContext('DEPT_LIST_CACHE_KEY', c);
    c.cachePutLarge(key, JSON.stringify([{ ticketId: 'T' }]));
    const g1 = c.cacheGeneration_(vm.runInContext('COMPLAINTS_CACHE_GEN_KEY', c));
    const out = JSON.parse(c.doPost({ postData: { contents: JSON.stringify({ serialNumber: 'SNX', submissionId: 'P-1', photos: [] }) } }).text);
    assert.equal(out.status, 'ok');
    assert.notEqual(c.cacheGeneration_(vm.runInContext('COMPLAINTS_CACHE_GEN_KEY', c)), g1);
    assert.ok(c.cacheGetLarge(key), 'department cache kept');
  });

  const failed = results.filter(r => r.status !== 'pass');
  results.forEach(r => console.log((r.status === 'pass' ? 'PASS  ' : 'FAIL  ') + r.name + (r.error ? '\n      ' + r.error : '')));
  console.log(`\n${results.length - failed.length}/${results.length} performance regression checks passed.`);
  process.exitCode = failed.length ? 1 : 0;
})();
