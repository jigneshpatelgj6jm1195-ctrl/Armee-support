const assert = require('assert/strict');
const crypto = require('crypto');
const fs = require('fs');
const vm = require('vm');
const zlib = require('zlib');

const backendSource = fs.readFileSync('google_apps_script_backend.js', 'utf8');
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

function makeContext() {
  const properties = new Map([['SESSION_SECRET', 'synthetic-test-secret'], ['IMPORT_KEY', 'synthetic-import-key']]);
  let uuidCounter = 0;
  const cacheStore = new Map();
  const context = vm.createContext({
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
  vm.runInContext(backendSource, context);
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
  try {
    await fn();
    results.push({ name, status: 'pass' });
  } catch (error) {
    results.push({ name, status: 'fail', error: error.message });
  }
}

(async () => {
  await test('slow admin refreshes never overlap and release the guard after failure', async () => {
    let finish;
    let loads = 0;
    let deptLoads = 0;
    const context = vm.createContext({
      document: { querySelector: () => null }, currentTab: 'dept',
      loadData: () => { loads++; return new Promise(resolve => { finish = resolve; }); },
      loadDeptDashboard: async () => { deptLoads++; throw new Error('offline'); },
      loadComplaints: async () => { throw new Error('Unexpected full complaint refresh'); },
      showToast() {}
    });
    vm.runInContext('let refreshInProgress = false;\n' + extractBlock(adminSource, 'async function silentRefreshData()'), context);
    const first = context.silentRefreshData();
    await context.silentRefreshData();
    assert.equal(loads, 1);
    finish();
    await first;
    assert.equal(deptLoads, 1);
    assert.equal(vm.runInContext('refreshInProgress', context), false);
  });

  await test('archive uses Archived column and preserves serial photo formula', () => {
    const c = makeContext();
    const headers = vm.runInContext('HEADERS', c);
    const row = Array(headers.length).fill('');
    row[0] = 1; row[1] = '2026-09-01T12:00:00Z'; row[24] = 'CASE-TEST'; row[31] = '=IMAGE("serial-photo")';
    const active = makeSheet('Complaints', [headers, row]);
    const archive = makeSheet('Archive_2026_09', [headers]);
    const db = makeSpreadsheet([active, archive]);
    c.archiveComplaints(db, '2026-09-01', '2026-09-02');
    assert.equal(archive.rows[1][34], 'YES');
    assert.equal(archive.rows[1][31], '=IMAGE("serial-photo")');
    c.restoreComplaints(db, ['CASE-TEST']);
    assert.equal(active.rows[1][34], '');
    assert.equal(active.rows[1][31], '=IMAGE("serial-photo")');
  });

  await test('archive splits a cross-month range into the correct monthly tabs', () => {
    const c = makeContext();
    const headers = vm.runInContext('HEADERS', c);
    const august = Array(headers.length).fill(''); august[1] = '2026-08-31T10:00:00'; august[24] = 'CASE-AUG';
    const september = Array(headers.length).fill(''); september[1] = '2026-09-01T10:00:00'; september[24] = 'CASE-SEP';
    const active = makeSheet('Complaints', [headers, august, september]);
    const db = makeSpreadsheet([active]);
    const out = c.archiveComplaints(db, '2026-08-31', '2026-09-01');
    assert.equal(out.archived, 2);
    assert.equal(db.getSheetByName('Archive_2026_08').rows[1][24], 'CASE-AUG');
    assert.equal(db.getSheetByName('Archive_2026_09').rows[1][24], 'CASE-SEP');
  });

  await test('department batch skips duplicate Ticket IDs', () => {
    const c = makeContext();
    const dept = makeSheet('DepartmentComplaints', [Array(33).fill('')]);
    const resolutions = makeSheet('DepartmentResolutions', [Array(19).fill('')]);
    c.requireImportOrAdmin_ = () => null;
    c.getOrCreateDeptSheet = () => dept;
    c.getOrCreateResSheet = () => resolutions;
    c.loadBranchStructure = () => ({});
    c.getOwningDistrictAdmin = () => null;
    c.getBranchForDeptComplaint = () => null;
    c.syncAllDepartmentToComplaints = () => {};
    const out = c.importDepartmentComplaints(makeSpreadsheet([dept, resolutions]), {
      rows: [{ TicketId: 'T-1', Ticket_Status: 'Pending' }, { TicketId: 'T-1', Ticket_Status: 'Pending' }],
      sendEmails: false,
    });
    assert.equal(out.inserted, 1);
    assert.equal(out.duplicateInputRowsSkipped, 1);
    assert.equal(resolutions.rows.length, 2);
  });

  await test('school complaint import writes an upload batch without row-by-row sheet calls', () => {
    const c = makeContext();
    c.requireImportOrAdmin_ = () => null;
    const complaints = makeSheet('Complaints', [Array(40).fill('')]);
    const master = makeSheet('SchoolComplaintMaster', [Array(23).fill('')]);
    const upload = makeSheet('SchoolComplaintUpload', [Array(13).fill('')]);
    const matchLog = makeSheet('SchoolComplaintMatchLog', [Array(9).fill('')]);
    const exceptions = makeSheet('SchoolComplaintExceptions', [Array(5).fill('')]);
    const out = c.importSchoolComplaints(makeSpreadsheet([complaints, master, upload, matchLog, exceptions]), {
      uploadId: 'UPLOAD-TEST', importDate: '2026-09-19', finalChunk: false,
      records: [
        { serialNumber: 'SERIAL-1', customerName: 'School One', equipment: 'CPU' },
        { serialNumber: '', customerName: 'School Missing Serial' },
      ],
    });
    assert.equal(out.status, 'ok');
    assert.equal(out.masterCount, 1);
    assert.equal(out.exceptionCount, 1);
    assert.equal(master.rows.length, 2);
    assert.equal(upload.rows.length, 3);
    assert.equal(upload.rows[1][0], 'UPLOAD-TEST');
    assert.equal(exceptions.rows.length, 2);
  });

  await test('DISE healer never overwrites an existing value', () => {
    const c = makeContext();
    const row = Array(19).fill(''); row[0] = 1; row[2] = 'ORIGINAL';
    const s = makeSheet('SchoolComplaintMaster', [Array(19).fill(''), row]);
    const count = c.updateSchoolComplaintDiseBulkInSheet(makeSpreadsheet([s]), [{ srNo: 1, dise: 'REPLACEMENT' }]);
    assert.equal(count, 0);
    assert.equal(s.rows[1][2], 'ORIGINAL');
  });

  await test('Pending OTP is excluded from active branch aging', () => {
    const c = makeContext();
    const row = Array(33).fill('');
    row[0] = 'District-X'; row[9] = 'T-OTP'; row[21] = '2026-09-01';
    const dept = makeSheet('DepartmentComplaints', [Array(33).fill(''), row]);
    c.loadBranchStructure = () => ({
      branches: [{ id: 'B1', name: 'Branch-X', districtOfficeId: 'O1' }],
      offices: [{ id: 'O1', name: 'Office-X' }],
      officeById: { O1: { name: 'Office-X' } },
    });
    c.getResolutionsMap = () => ({ 'T-OTP': { internalStatus: 'PendingOTP', owningDistrictAdmin: 'Admin-X' } });
    c.resolveBranchId = () => ({ branchId: 'B1' });
    const out = c.getDepartmentDashboard(makeSpreadsheet([dept]));
    const agingTotal = out.branches.reduce((sum, branch) =>
      sum + Object.values(branch.department).reduce((a, b) => a + b, 0), 0);
    assert.equal(out.pendency.pendingOtp, 1);
    assert.equal(agingTotal, 0);
  });

  await test('school master validates input before clearing existing rows', () => {
    const c = makeContext();
    c.requireImportOrAdmin_ = () => null;
    const s = makeSheet('SchoolMaster', [Array(10).fill(''), Array(10).fill('existing')]);
    const out = c.importSchoolMaster(makeSpreadsheet([s]), {});
    assert.equal(out.status, 'error');
    assert.equal(s.rows.length, 2);
  });

  await test('school identity replacement records delete and add together after validation', () => {
    const c = makeContext();
    const ss = makeSpreadsheet([]);
    const out = c.replaceSchoolRecord(ss, {
      oldDise: 'OLD-DISE',
      oldProject: 'ICT',
      newRecord: { dise: 'NEW-DISE', project: 'GK', school: 'Replacement School' },
    });
    const updates = ss.getSheetByName('SchoolUpdates');
    assert.equal(out.status, 'ok');
    assert.equal(updates.rows.length, 3);
    assert.deepEqual(JSON.parse(JSON.stringify(updates.rows[1].slice(0, 2))), ['OLD-DISE', 'deleted']);
    assert.deepEqual(JSON.parse(JSON.stringify(updates.rows[2].slice(0, 2))), ['NEW-DISE', 'added']);
    assert.equal(c.replaceSchoolRecord(ss, { oldDise: 'OLD-DISE', oldProject: 'ICT', newRecord: { dise: '', project: 'GK' } }).status, 'error');
  });

  await test('master save migrates plaintext passwords to non-exported hashes', () => {
    const c = makeContext();
    const existing = {
      equipment: [], users: [],
      accessUsers: [{ id: 'U1', email: 'admin@example.test', password: 'legacy-password', role: 'super_admin', status: 'active' }],
    };
    const masterSheet = makeSheet('MasterData', [[JSON.stringify(existing)]]);
    const db = makeSpreadsheet([masterSheet]);
    c.saveMasterData(db, {
      equipment: [], users: [],
      accessUsers: [{ id: 'U1', email: 'admin@example.test', role: 'super_admin', status: 'active' }],
    });
    const stored = JSON.parse(masterSheet.rows[0][0]);
    assert.equal('password' in stored.accessUsers[0], false);
    assert.ok(stored.accessUsers[0].passwordHash);
    assert.ok(stored.accessUsers[0].passwordSalt);
    const client = c.sanitizeMasterForClient(stored);
    assert.equal('passwordHash' in client.accessUsers[0], false);
    assert.equal('passwordSalt' in client.accessUsers[0], false);
    const login = c.handleLogin(db, { email: 'admin@example.test', password: 'legacy-password' });
    assert.equal(login.status, 'ok');
  });

  await test('master save prevents removing the last active super admin', () => {
    const c = makeContext();
    const existing = {
      equipment: [], users: [],
      accessUsers: [{ id: 'U1', email: 'admin@example.test', password: 'legacy-password', role: 'super_admin', status: 'active' }],
    };
    const sheet = makeSheet('MasterData', [[JSON.stringify(existing)]]);
    assert.throws(() => c.saveMasterData(makeSpreadsheet([sheet]), {
      equipment: [], users: [],
      accessUsers: [{ id: 'U1', email: 'admin@example.test', role: 'viewer', status: 'active' }],
    }), /At least one active super admin/);
    assert.equal(JSON.parse(sheet.rows[0][0]).accessUsers[0].role, 'super_admin');
  });

  await test('bootstrap Super Admin is protected by Script Properties, not source code', () => {
    const loginBlock = extractBlock(backendSource, 'function handleLogin(ss, data)');
    const bootstrapBlock = extractBlock(backendSource, 'function getBootstrapSuperAdmin_()');
    assert.match(loginBlock, /getBootstrapSuperAdmin_\(\)/);
    assert.match(bootstrapBlock, /PropertiesService\.getScriptProperties\(\)/);
    assert.match(bootstrapBlock, /BOOTSTRAP_SUPER_ADMIN_PASSWORD_HASH/);
    assert.doesNotMatch(backendSource, /fdJr-nJq5-QJJX/);
  });

  await test('stable submission ID prevents duplicate after log failure', () => {
    const c = makeContext();
    const headers = vm.runInContext('HEADERS', c);
    const active = makeSheet('Complaints', [headers]);
    const log = makeSheet('SubmissionLog', [vm.runInContext('SUBMISSION_LOG_HEADERS', c)]);
    log.appendRow = () => { throw new Error('synthetic log outage'); };
    const db = makeSpreadsheet([active, log]);
    c.getPhotoFolder = () => ({});
    c.formatLastRow = () => {};
    c.syncSchoolComplaintMasterStatus = () => {};
    const payload = { serialNumber: 'SERIAL-1', submissionId: 'SUBMISSION-1', photos: [] };
    const first = JSON.parse(c.handleSubmitComplaint(db, payload).text);
    const second = JSON.parse(c.handleSubmitComplaint(db, payload).text);
    assert.equal(first.status, 'ok');
    assert.equal(second.duplicateIgnored, true);
    assert.equal(active.rows.length, 2);
    assert.equal(first.caseId, second.caseId);
  });

  await test('bulk complaint update preserves formulas while applying changes', () => {
    const c = makeContext();
    const headers = vm.runInContext('HEADERS', c);
    const row = Array(headers.length).fill('');
    row[24] = 'CASE-UPDATE'; row[8] = 'District-X'; row[23] = 'Open';
    row[27] = '=IMAGE("https://example.test/photo.jpg")';
    const active = makeSheet('Complaints', [headers, row]);
    c.syncSchoolComplaintMasterStatus = () => {};
    const count = c.updateComplaintsStatus(makeSpreadsheet([active]), [{
      caseId: 'CASE-UPDATE', district: 'District-X', status: 'Closed'
    }]);
    assert.equal(count, 1);
    assert.equal(active.rows[1][23], 'Closed');
    assert.equal(active.rows[1][27], '=IMAGE("https://example.test/photo.jpg")');
  });

  await test('school resolution rejects a missing portal token', () => {
    const c = makeContext();
    const row = Array(21).fill(''); row[0] = 1; row[2] = 'DISE-1'; row[13] = 'SERIAL-1';
    const s = makeSheet('SchoolComplaintMaster', [Array(21).fill(''), row]);
    const out = c.resolveSchoolComplaintFromPortal(makeSpreadsheet([s]), { srNo: 1, dise: 'DISE-1', status: 'Closed' });
    assert.equal(out.status, 'error');
    assert.equal(s.rows[1][16], '');
  });

  await test('school resolution accepts a valid token bound to the complaint', () => {
    const c = makeContext();
    const row = Array(21).fill(''); row[0] = 1; row[2] = 'DISE-1'; row[13] = 'SERIAL-1';
    const s = makeSheet('SchoolComplaintMaster', [Array(21).fill(''), row]);
    const token = c.issuePortalResolutionToken_(1, 'DISE-1', 'SERIAL-1');
    const out = c.resolveSchoolComplaintFromPortal(makeSpreadsheet([s]), {
      srNo: 1, dise: 'DISE-1', serialNumber: 'SERIAL-1', status: 'Closed', portalToken: token,
    });
    assert.equal(out, 1);
    assert.equal(s.rows[1][16], 'Closed');
  });

  await test('school resolution rejects invalid state changes and token replay', () => {
    const c = makeContext();
    const row = Array(21).fill(''); row[0] = 1; row[2] = 'DISE-1'; row[13] = 'SERIAL-1';
    const s = makeSheet('SchoolComplaintMaster', [Array(21).fill(''), row]);
    const db = makeSpreadsheet([s]);
    const token = c.issuePortalResolutionToken_(1, 'DISE-1', 'SERIAL-1');
    let out = c.resolveSchoolComplaintFromPortal(db, {
      srNo: 1, dise: 'DISE-1', serialNumber: 'SERIAL-1', status: 'Deleted', portalToken: token,
    });
    assert.equal(out.code, 'validation');
    out = c.resolveSchoolComplaintFromPortal(db, {
      srNo: 1, dise: 'DISE-1', serialNumber: 'SERIAL-1', status: 'Closed', portalToken: token,
    });
    assert.equal(out, 1);
    out = c.resolveSchoolComplaintFromPortal(db, {
      srNo: 1, dise: 'DISE-1', serialNumber: 'SERIAL-1', status: 'Open', portalToken: token,
    });
    assert.equal(out.code, 'conflict');
    out = c.resolveSchoolComplaintFromPortal(db, {
      srNo: 1, dise: 'DISE-1', serialNumber: 'SERIAL-1', status: 'Closed', portalToken: token,
    });
    assert.equal(out.duplicateIgnored, true);
  });

  await test('department resolution rejects missing portal token and unauthenticated access', () => {
    const c = makeContext();
    const row = ['SURAT', '', 'SURAT CORPO.', '', '', '', '', '24221503186', 'GURUKUL', '2026/016942'];
    const s = makeSheet('DepartmentComplaints', [Array(10).fill(''), row]);
    const db = makeSpreadsheet([s]);
    const out = c.requireDepartmentTicketAuth_(db, { ticketId: '2026/016942' });
    assert.equal(out.status, 'error');
    assert.equal(out.code, 'auth');
  });

  await test('department resolution accepts a valid token bound to ticket and school', () => {
    const c = makeContext();
    const row = ['SURAT', '', 'SURAT CORPO.', '', '', '', '', '24221503186', 'GURUKUL', '2026/016942'];
    const s = makeSheet('DepartmentComplaints', [Array(10).fill(''), row]);
    const db = makeSpreadsheet([s]);
    const token = c.issueDepartmentResolutionToken_('2026/016942', '24221503186');
    const out = c.requireDepartmentTicketAuth_(db, {
      ticketId: '2026/016942',
      portalToken: token,
      resolutionAction: 'closed_without_otp'
    });
    assert.equal(out, null);
  });

  await test('department resolution rejects finalize_otp without admin credentials', () => {
    const c = makeContext();
    const row = ['SURAT', '', 'SURAT CORPO.', '', '', '', '', '24221503186', 'GURUKUL', '2026/016942'];
    const s = makeSheet('DepartmentComplaints', [Array(10).fill(''), row]);
    const db = makeSpreadsheet([s]);
    const token = c.issueDepartmentResolutionToken_('2026/016942', '24221503186');
    const out = c.requireDepartmentTicketAuth_(db, {
      ticketId: '2026/016942',
      portalToken: token,
      resolutionAction: 'finalize_otp'
    });
    assert.equal(out.status, 'error');
    assert.equal(out.code, 'auth');
    assert.match(out.message, /Only administrators/);
  });

  await test('department resolution rejects mismatched or expired portal token', () => {
    const c = makeContext();
    const row = ['SURAT', '', 'SURAT CORPO.', '', '', '', '', '24221503186', 'GURUKUL', '2026/016942'];
    const s = makeSheet('DepartmentComplaints', [Array(10).fill(''), row]);
    const db = makeSpreadsheet([s]);
    const wrongTicketToken = c.issueDepartmentResolutionToken_('2026/999999', '24221503186');
    const wrongSchoolToken = c.issueDepartmentResolutionToken_('2026/016942', '99999999999');
    const out1 = c.requireDepartmentTicketAuth_(db, { ticketId: '2026/016942', portalToken: wrongTicketToken });
    assert.equal(out1.code, 'auth');
    const out2 = c.requireDepartmentTicketAuth_(db, { ticketId: '2026/016942', portalToken: wrongSchoolToken });
    assert.equal(out2.code, 'auth');
  });

  await test('department resolution rejects modifying an already closed ticket', () => {
    const c = makeContext();
    const resHeaders = Array(19).fill('');
    const closedRow = ['2026/016942', 'Closed', 'ClosedWithOTP', '1234', 'Tech', '', '', '', '', '', '', '', '', 'UD123', '', '', '', '', ''];
    const resSheet = makeSheet('DepartmentResolutions', [resHeaders, closedRow]);
    const compSheet = makeSheet('DepartmentComplaints', [Array(10).fill('')]);
    const db = makeSpreadsheet([resSheet, compSheet]);
    const outConflict = c.resolveDepartmentComplaint(db, {
      ticketId: '2026/016942',
      resolutionAction: 'closed_without_otp',
      serialNumber: 'UD123'
    });
    assert.equal(outConflict.code, 'conflict');
    const outDuplicate = c.resolveDepartmentComplaint(db, {
      ticketId: '2026/016942',
      resolutionAction: 'closed_with_otp',
      serialNumber: 'UD123'
    });
    assert.equal(outDuplicate.duplicateIgnored, true);
    assert.equal(outDuplicate.internalStatus, 'Closed');
  });

  await test('sensitive department read rejects missing admin authentication', () => {
    const c = makeContext();
    c.SpreadsheetApp = { openById: () => makeSpreadsheet([]) };
    const out = JSON.parse(c.doGet({ parameter: { action: 'get_department_complaints_list' } }).text);
    assert.equal(out.status, 'error');
    assert.equal(out.code, 'auth');
  });

  await test('department page bounds payload and applies search, status, sorting, and pagination server-side', () => {
    const c = makeContext();
    const rows = [
      { ticketId: 'T-3', school: 'Gamma School', district: 'North', internalStatus: 'Pending', branchId: 'B-1', businessDays: 3 },
      { ticketId: 'T-1', school: 'Alpha School', district: 'North', internalStatus: 'Closed', branchId: 'B-1', businessDays: 5 },
      { ticketId: 'T-2', school: 'Beta School', district: 'North', internalStatus: 'Pending', branchId: 'B-2', businessDays: 7 },
      { ticketId: 'T-4', school: 'Delta School', district: 'North', internalStatus: 'PendingOTP', branchId: 'B-2', businessDays: 9 },
    ];
    const out = c.getDepartmentComplaintsPage_(rows, {
      activeOnly: 'true', page: '1', pageSize: '1', sortKey: 'businessDays', sortOrder: 'desc'
    });
    assert.equal(out.total, 2);
    assert.equal(out.items.length, 1);
    assert.equal(out.items[0].ticketId, 'T-2');
    assert.equal(out.hasMore, true);
    const search = c.getDepartmentComplaintsPage_(rows, {
      search: 'gamma', status: 'Pending', pageSize: '5000', sortKey: 'ticketId', sortOrder: 'asc'
    });
    assert.equal(search.items.length, 1);
    assert.equal(search.items[0].ticketId, 'T-3');
    assert.equal(search.pageSize, 200);
  });

  await test('department page applies district authorization before returning any page', () => {
    const c = makeContext();
    c.SpreadsheetApp = { openById: () => makeSpreadsheet([]) };
    c.verifyAuthToken_ = token => token === 'district-token' ? { email: 'north@example.test', role: 'district_admin' } : null;
    c.getAuthorizedDistrictMap_ = () => ({ NORTH: true });
    c.getCachedDepartmentComplaints_ = () => [
      { ticketId: 'N-1', district: 'North', internalStatus: 'Pending', businessDays: 1 },
      { ticketId: 'S-1', district: 'South', internalStatus: 'Pending', businessDays: 9 },
    ];
    const out = JSON.parse(c.doGet({ parameter: {
      action: 'get_department_complaints_page', authToken: 'district-token', pageSize: '50'
    } }).text);
    assert.equal(out.total, 1);
    assert.deepEqual(out.items.map(row => row.ticketId), ['N-1']);
  });

  await test('department summary preserves pendency, aging, first-attended flow, and compact trend counts', () => {
    const c = makeContext();
    const out = c.getDepartmentDashboardSummary_([
      { ticketId: 'P', district: 'North', branchId: 'B1', branchName: 'One', internalStatus: 'Pending', businessDays: 1, createdDate: '2026-09-18' },
      { ticketId: 'I', district: 'North', branchId: 'B1', branchName: 'One', internalStatus: 'InProgress', businessDays: 4, createdDate: '2026-09-18', resolvedAt: '2026-09-18' },
      { ticketId: 'O', district: 'South', branchId: '', branchName: 'Unmapped', internalStatus: 'PendingOTP', businessDays: 9, createdDate: '2026-09-17', resolutionDate: '2026-09-18' },
      { ticketId: 'C', district: 'South', branchId: 'B2', branchName: 'Two', internalStatus: 'Closed', closureType: 'ClosedWithOTP', businessDays: 8, createdDate: '2026-09-17', resolvedAt: '2026-09-18', resolutionDate: '2026-09-19' },
    ], { flowDate: '2026-09-18' });
    assert.deepEqual(JSON.parse(JSON.stringify(out.dashboard.pendency)), {
      total: 4, pending: 1, inProgress: 1, partRequest: 0, pendingOtp: 1, closedWithOTP: 1, closedWithoutOTP: 0
    });
    assert.deepEqual(JSON.parse(JSON.stringify(out.dashboard.branches.find(row => row.branchId === 'B1').department)), { '0-2': 1, '3-5': 1, '6+': 0 });
    assert.equal(out.dashboard.branches.some(row => row.branchId === 'B2'), false);
    assert.deepEqual(JSON.parse(JSON.stringify(out.dashboard.unmapped.department)), { '0-2': 0, '3-5': 0, '6+': 0 });
    assert.deepEqual(JSON.parse(JSON.stringify(out.todayFlow.find(row => row.branchName === 'One'))), {
      branchName: 'One', inflow: 2, InProgress: 1, PartRequest: 0, PendingOTP: 0, Closed: 0
    });
    assert.equal(out.todayFlow.find(row => row.branchName === 'Two').Closed, 1);
    assert.deepEqual(JSON.parse(JSON.stringify(out.trend)), [{ createdDate: '2026-09-17', count: 2 }, { createdDate: '2026-09-18', count: 2 }]);
  });

  await test('department inflow uses the India business date for logged timestamps', () => {
    const c = makeContext();
    const justAfterMidnightInIndia = vm.runInContext("new Date('2026-09-18T18:30:00.000Z')", c);
    assert.equal(c.departmentDateKey_(justAfterMidnightInIndia), '2026-09-19');
    const out = c.getDepartmentDashboardSummary_([
      { ticketId: 'IST-1', district: 'North', branchId: 'B1', branchName: 'One', internalStatus: 'Pending', businessDays: 0, createdDate: justAfterMidnightInIndia }
    ], { flowDate: '2026-09-19' });
    assert.equal(out.todayFlow[0].inflow, 1);
  });

  await test('department summary applies district authorization before calculating aggregates', () => {
    const c = makeContext();
    c.SpreadsheetApp = { openById: () => makeSpreadsheet([]) };
    c.verifyAuthToken_ = token => token === 'district-token' ? { email: 'north@example.test', role: 'district_admin' } : null;
    c.getAuthorizedDistrictMap_ = () => ({ NORTH: true });
    c.getCachedDepartmentComplaints_ = () => [
      { ticketId: 'N-1', district: 'North', internalStatus: 'Pending', businessDays: 1, createdDate: '2026-09-18' },
      { ticketId: 'S-1', district: 'South', internalStatus: 'Closed', businessDays: 1, createdDate: '2026-09-18' },
    ];
    const out = JSON.parse(c.doGet({ parameter: {
      action: 'get_department_dashboard_summary', authToken: 'district-token', flowDate: '2026-09-18'
    } }).text);
    assert.equal(out.dashboard.pendency.total, 1);
    assert.equal(out.dashboard.pendency.pending, 1);
    assert.deepEqual(out.trend, [{ createdDate: '2026-09-18', count: 1 }]);
  });

  await test('complaint list rejects anonymous reads and filters district access', () => {
    const c = makeContext();
    const headers = vm.runInContext('HEADERS', c);
    const north = Array(headers.length).fill(''); north[8] = 'North'; north[24] = 'CASE-N';
    const south = Array(headers.length).fill(''); south[8] = 'South'; south[24] = 'CASE-S';
    const active = makeSheet('Complaints', [headers, north, south]);
    const db = makeSpreadsheet([active]);
    c.SpreadsheetApp = { openById: () => db };
    let out = JSON.parse(c.doGet({ parameter: { action: 'get_complaints' } }).text);
    assert.equal(out.code, 'auth');
    c.verifyAuthToken_ = token => token === 'district-token' ? { email: 'north@example.test', role: 'district_admin' } : null;
    c.getAuthorizedDistrictMap_ = () => ({ NORTH: true });
    out = JSON.parse(c.doGet({ parameter: { action: 'get_complaints', authToken: 'district-token' } }).text);
    assert.deepEqual(out.map(row => row.caseId), ['CASE-N']);
  });

  await test('complaint updates authorize against the server-side district', () => {
    const c = makeContext();
    const headers = vm.runInContext('HEADERS', c);
    const north = Array(headers.length).fill(''); north[8] = 'North'; north[24] = 'CASE-N';
    const south = Array(headers.length).fill(''); south[8] = 'South'; south[24] = 'CASE-S';
    const db = makeSpreadsheet([makeSheet('Complaints', [headers, north, south])]);
    c.verifyAuthToken_ = () => ({ email: 'north@example.test', role: 'district_admin' });
    c.getAuthorizedDistrictMap_ = () => ({ NORTH: true });
    const scoped = c.filterComplaintUpdatesForAdmin_(db, [
      { caseId: 'CASE-N', district: 'South' },
      { caseId: 'CASE-S', district: 'North' },
    ], 'district-token');
    assert.deepEqual(scoped.map(row => row.caseId), ['CASE-N']);
  });

  await test('bare GET no longer exposes the full complaint list', () => {
    const c = makeContext();
    c.SpreadsheetApp = { openById: () => makeSpreadsheet([]) };
    const out = JSON.parse(c.doGet({ parameter: {} }).text);
    assert.equal(out.code, 'action_required');
  });

  await test('public duplicate check returns no complaint identity fields', () => {
    const c = makeContext();
    const headers = vm.runInContext('HEADERS', c);
    const row = Array(headers.length).fill('');
    row[1] = new Date(); row[10] = 'Private School'; row[17] = 'SERIAL-PROBE'; row[24] = 'CASE-PRIVATE';
    const out = c.checkDuplicateSerial(makeSpreadsheet([makeSheet('Complaints', [headers, row])]), 'SERIAL-PROBE');
    assert.equal(out.isDuplicate, true);
    assert.equal('existingSchool' in out, false);
    assert.equal('existingCaseId' in out, false);
    assert.equal('existingDate' in out, false);
  });

  await test('unknown GET action returns an explicit route error', () => {
    const c = makeContext();
    c.SpreadsheetApp = { openById: () => makeSpreadsheet([]) };
    const out = JSON.parse(c.doGet({ parameter: { action: 'not-a-route' } }).text);
    assert.equal(out.code, 'unknown_action');
  });

  await test('cache readers see one complete generation during refresh', () => {
    const c = makeContext();
    const map = new Map();
    let observeNextGeneration = false;
    let observed;
    const cache = {
      get: key => map.has(key) ? map.get(key) : null,
      getAll: keys => Object.fromEntries(keys.map(k => [k, map.has(k) ? map.get(k) : null])),
      put(key, value) {
        map.set(key, value);
        if (observeNextGeneration && key !== 'test__meta' && key.endsWith('__0')) {
          observeNextGeneration = false;
          observed = c.cacheGetLarge('test');
        }
      },
      removeAll: keys => keys.forEach(k => map.delete(k)),
    };
    c.CacheService = { getScriptCache: () => cache };
    vm.runInContext('CACHE_CHUNK_SIZE=5', c);
    c.cachePutLarge('test', 'OOOOOOOOOO', 60);
    observeNextGeneration = true;
    c.cachePutLarge('test', 'NNNNNNNNNN', 60);
    assert.equal(observed, 'OOOOOOOOOO');
    assert.equal(c.cacheGetLarge('test'), 'NNNNNNNNNN');
  });

  await test('API timeout remains active while response body is read', async () => {
    let aborted = false;
    const context = vm.createContext({
      Map, Date, Math, performance, setTimeout, clearTimeout, AbortController,
      logToDebug() {},
      fetch: async (_url, options) => ({
        ok: true,
        text: () => new Promise((resolve, reject) => {
          options.signal.addEventListener('abort', () => { aborted = true; reject(new Error('aborted')); });
        }),
      }),
    });
    vm.runInContext(extractBlock(adminSource, 'class ArmeeApiClient'), context);
    const outcome = await Promise.race([
      vm.runInContext('new ArmeeApiClient().request("https://example.invalid", { timeoutMs: 5, attempts: 1 })', context)
        .then(() => 'resolved', () => 'rejected'),
      new Promise(resolve => setTimeout(() => resolve('pending'), 50)),
    ]);
    assert.equal(outcome, 'rejected');
    assert.equal(aborted, true);
  });

  await test('failed online save keeps dirty state and never reports success', async () => {
    const button = { disabled: false, textContent: '' };
    const messages = [];
    const context = vm.createContext({
      Date, JSON,
      navigator: { onLine: true },
      GOOGLE_SCRIPT_URL: 'https://example.invalid',
      LS_MASTER_KEY: 'master', LS_COMPLAINTS_KEY: 'complaints',
      masterData: {}, complaintsList: [], isDirty: true,
      document: { getElementById: () => button },
      localStorage: { setItem() {} },
      isLocal: () => false,
      adminAuthToken: () => 'synthetic-token',
      updateSaveBar() {},
      showToast: (message, type) => messages.push({ message, type }),
      fetch: async () => { throw new Error('synthetic network failure'); },
      postJsonWithDeadline: async () => { throw new Error('synthetic network failure'); },
    });
    vm.runInContext(extractBlock(adminSource, 'async function saveAll'), context);
    await vm.runInContext('saveAll()', context);
    assert.equal(context.isDirty, true);
    assert.equal(messages.some(m => m.type === 'success'), false);
    assert.equal(messages.some(m => m.type === 'warn'), true);
  });

  await test('save data uses bounded confirmed requests for local and live writes', () => {
    const block = extractBlock(adminSource, 'async function saveAll()');
    assert.match(block, /postJsonWithDeadline\('\/update_master'/);
    assert.match(block, /postJsonWithDeadline\('\/update_complaints'/);
    assert.match(block, /postJsonWithDeadline\(GOOGLE_SCRIPT_URL, payload, \{ timeoutMs: 45000 \}\)/);
    assert.doesNotMatch(block, /const res = await fetch\(/);
  });

  await test('downloadable source contains no embedded import or fallback admin secret', () => {
    assert.equal(/ArmeeICT-[A-Za-z0-9_-]+/.test(adminSource), false);
    assert.equal(/var\s+SUPER_ADMIN\s*=/.test(backendSource), false);
    assert.equal(adminSource.includes('armee@2026'), false);
  });

  await test('browser cache strips every stored credential field', () => {
    const context = vm.createContext({ JSON });
    vm.runInContext(extractBlock(adminSource, 'function masterDataForBrowserCache'), context);
    const clean = vm.runInContext(`masterDataForBrowserCache({
      users: [{ password: 'one', passwordHash: 'hash', passwordSalt: 'salt', name: 'A' }],
      accessUsers: [{ password: 'two', passwordHash: 'hash', passwordSalt: 'salt', name: 'B' }]
    })`, context);
    for (const group of [clean.users, clean.accessUsers]) {
      assert.equal('password' in group[0], false);
      assert.equal('passwordHash' in group[0], false);
      assert.equal('passwordSalt' in group[0], false);
    }
  });

  await test('large dashboard charts do not force one tick per record', () => {
    assert.equal(/stepSize\s*:\s*1/.test(adminSource), false);
    assert.ok((adminSource.match(/precision\s*:\s*0/g) || []).length >= 8);
    assert.match(adminSource, /\+ \(Number\(r\.count\) \|\| 1\)/);
  });

  await test('local branch summary resolves without contacting production', () => {
    const block = extractBlock(adminSource, 'async function ensureBranchDataLoaded');
    assert.match(block, /if \(isLocal\(\)\)/);
    assert.match(block, /bmBranches\s*=\s*\[\]/);
  });

  await test('consolidated dashboard uses authenticated complaint route and local fixtures', () => {
    const block = extractBlock(adminSource, 'async function loadUnifiedDashboard');
    // fetchEnvelope always sends the admin authToken (v2, server-cached route)
    assert.match(block, /fetchEnvelope\('get_complaints'/);
    assert.match(extractBlock(adminSource, 'async function fetchEnvelope'), /&authToken=' \+ encodeURIComponent\(adminAuthToken\(\)\)/);
    assert.match(block, /\/school_complaint_data\.json/);
    assert.match(block, /\/deptlist\.json/);
    assert.doesNotMatch(block, /fetchJsonWithRetry\(GOOGLE_SCRIPT_URL\s*,/);
  });

  await test('local department dashboard never contacts production', () => {
    const block = extractBlock(adminSource, 'async fetchDepartmentComplaints(forceNetwork = false)');
    assert.match(block, /if \(isLocal\(\)\)/);
    assert.match(block, /fetchRequiredArray\('\/deptlist\.json'/);
    assert.ok(block.indexOf('if (isLocal())') < block.indexOf('const summaryUrl = GOOGLE_SCRIPT_URL'));
  });

  await test('production department dashboard loads a compact summary before a bounded ticket page', () => {
    const block = extractBlock(adminSource, 'async fetchDepartmentComplaints(forceNetwork = false)');
    assert.match(block, /get_department_dashboard_summary/);
    assert.match(block, /get_department_complaints_page/);
    assert.ok(block.indexOf('const summary = await') < block.indexOf('const page = await'));
    assert.doesNotMatch(block, /get_department_complaints_list&authToken/);
  });

  await test('production department client requests a scoped summary then one bounded page', async () => {
    const block = extractBlock(adminSource, 'const departmentApi = {');
    const requests = [];
    const c = vm.createContext({
      URLSearchParams,
      Date,
      currentUser: { email: 'admin@example.test' },
      deptList: [], deptDash: null, deptDrillFilter: null, deptPageSize: 50, deptCurrentPage: 1,
      deptSortKey: 'createdDate', deptSortOrder: 'desc',
      deptPageMeta: null, deptTodayFlow: [], deptTrendData: [], deptUsesServerPaging: false,
      unifiedDepartmentRows: null,
      DEPT_CACHE_STORAGE_PREFIX: 'test:', DEPT_CACHE_MAX_AGE_MS: 900000,
      GOOGLE_SCRIPT_URL: 'https://backend.example.test/exec',
      document: { getElementById: id => ({ value: id === 'todayFlowDatePicker' ? '2026-09-19' : '' }) },
      localStorage: { getItem: () => null, setItem() {}, removeItem() {} },
      isLocal: () => false,
      adminAuthToken: () => 'test-token',
      adminApi: { request: async url => {
        requests.push(url);
        if (url.includes('get_department_dashboard_summary')) {
          return { status: 'ok', dashboard: { pendency: {} }, todayFlow: [], trend: [] };
        }
        return { status: 'ok', items: [{ ticketId: 'T-1' }], total: 1, page: 1, pageSize: 50, totalPages: 1 };
      } },
      filterByUserDistricts: rows => rows,
      logToDebug() {}, console,
    });
    vm.runInContext("function deptCacheStorageKey(){return 'test:admin';} function deptLocalDateKey(){return '2026-09-19';}\n" + block, c);
    await vm.runInContext('departmentApi.fetchDepartmentComplaints()', c);
    assert.equal(requests.length, 2);
    assert.match(requests[0], /action=get_department_dashboard_summary/);
    assert.match(requests[0], /flowDate=2026-09-19/);
    assert.match(requests[1], /action=get_department_complaints_page/);
    assert.match(requests[1], /pageSize=50/);
    assert.match(requests[1], /activeOnly=true/);
    assert.equal(vm.runInContext('deptList[0].ticketId', c), 'T-1');
    assert.equal(vm.runInContext('deptUsesServerPaging', c), true);
  });

  await test('school master refresh uses bounded API requests and retains prior data on failure', () => {
    const block = extractBlock(adminSource, 'async function loadSchoolList()');
    assert.match(block, /const previousSchoolDataRaw = schoolDataRaw/);
    assert.match(block, /adminApi\.request\('\/school_data\.json', \{[\s\S]*?timeoutMs: 5000/);
    assert.match(block, /cacheKey: 'get_school_master'/);
    assert.match(block, /cacheKey: 'get_school_updates'/);
    assert.match(block, /Retained the previous school list after a failed refresh/);
  });

  await test('master data refresh uses bounded API requests and retains prior data on failure', () => {
    const block = extractBlock(adminSource, 'async function loadData()');
    assert.match(block, /const previousMasterData = masterData/);
    assert.match(block, /cacheKey: 'get_master'/);
    assert.match(block, /cacheKey: 'local_master_data'/);
    assert.match(block, /Retained the previous master data after a failed refresh/);
    assert.match(block, /Empty setup values are not confirmed data/);
  });

  await test('department photo preview validates stored URLs before assigning image sources', () => {
    const marker = adminSource.indexOf('// Photos');
    const block = adminSource.slice(marker, marker + 1100);
    assert.match(block, /const serialPhotoSrc = safePhotoUrl\(r\.serialPhotoUrl\)/);
    assert.match(block, /const suspectedPhotoSrc = safePhotoUrl\(r\.suspectedPartPhotoUrl\)/);
    assert.match(block, /deptDetSerialImg'\)\.src = serialPhotoSrc/);
    assert.match(block, /deptDetSuspectedImg'\)\.src = suspectedPhotoSrc/);
  });

  await test('complaint photo editor validates preview and saved image URLs', () => {
    const previewBlock = extractBlock(adminSource, 'function updateEditPhotoPreview()');
    const serialPreviewBlock = extractBlock(adminSource, 'function updateEditSerialPhotoPreview()');
    const saveBlock = extractBlock(adminSource, 'function saveComplaintEdit()');
    assert.match(previewBlock, /const safeUrl = safePhotoUrl\(url\)/);
    assert.match(serialPreviewBlock, /const safeUrl = safePhotoUrl\(url\)/);
    assert.match(saveBlock, /newPhotoUrl && !safePhotoUrl\(newPhotoUrl\)/);
    assert.match(saveBlock, /newSerialPhotoUrl && !safePhotoUrl\(newSerialPhotoUrl\)/);
  });

  await test('branch summary labels a failed mapping refresh as unavailable instead of zero data', () => {
    const loadBlock = extractBlock(adminSource, 'async function ensureBranchDataLoaded()');
    const renderBlock = extractBlock(adminSource, 'async function renderBranchSummary(complaints)');
    assert.match(loadBlock, /branchLoadError = 'Branch mapping data is temporarily unavailable/);
    assert.match(loadBlock, /Branch totals are not shown as zero/);
    assert.match(renderBlock, /if \(branchLoadError\)/);
  });

  await test('pending-OTP finalization uses the authenticated department resolution POST', () => {
    const block = extractBlock(adminSource, 'async function finalizePendingOtp(ticketId)');
    assert.match(block, /postDepartmentResolution\(/);
    assert.match(block, /action: 'resolve_department_complaint'/);
    assert.match(block, /resolutionAction: 'finalize_otp'/);
    assert.match(block, /authToken: adminAuthToken\(\)/);
    assert.doesNotMatch(block, /action=finalize_pending_otp/);
  });

  await test('live login and pending-OTP finalization use a bounded confirmed JSON POST', () => {
    const loginBlock = extractBlock(adminSource, 'async function submitAdminLogin()');
    const otpBlock = extractBlock(adminSource, 'async function finalizePendingOtp(ticketId)');
    const start = adminSource.indexOf('async function postJsonWithDeadline(url, payload, options = {})');
    const postBlock = adminSource.slice(start, adminSource.indexOf('/* ─────────────── DEPARTMENT DASHBOARD', start));
    assert.ok(start >= 0, 'postJsonWithDeadline helper must exist');
    assert.match(loginBlock, /postJsonWithDeadline\([\s\S]*?timeoutMs: 20000/);
    assert.match(loginBlock, /postJsonWithDeadline\('\/local_login'[\s\S]*?timeoutMs: 15000/);
    assert.match(otpBlock, /postDepartmentResolution\(/);
    assert.match(postBlock, /const text = await response\.text\(\)/);
    assert.match(postBlock, /Request timed out after/);
  });

  await test('Admin content stays hidden until a valid session is available', () => {
    const initBlock = extractBlock(adminSource, "window.addEventListener('DOMContentLoaded', async () =>");
    const loginBlock = extractBlock(adminSource, 'async function submitAdminLogin()');
    assert.match(adminSource, /<body class="auth-pending">/);
    assert.match(adminSource, /body\.auth-pending > :not\(#adminLoginModal\):not\(\.toast\):not\(script\)/);
    assert.ok(initBlock.indexOf('const session = getSession()') < initBlock.indexOf('await loadData()'), 'session must be checked before loading application data');
    assert.match(initBlock, /document\.body\.classList\.remove\('auth-pending'\);/);
    assert.match(loginBlock, /localStorage\.setItem\(SESSION_KEY, JSON\.stringify\(session\)\);\s*document\.body\.classList\.remove\('auth-pending'\);/);
  });

  await test('School Master uploads use bounded requests for both local files and live records', () => {
    const block = extractBlock(adminSource, 'async function uploadSchoolExcel(event)');
    const bodyHelper = extractBlock(adminSource, 'async function postBodyWithDeadline(url, body, options)');
    assert.match(block, /postBodyWithDeadline\('\/upload_school_excel', file, \{ timeoutMs: 60000 \}\)/);
    assert.match(block, /postJsonWithDeadline\(GOOGLE_SCRIPT_URL, \{[\s\S]*?action: 'import_school_master'/);
    assert.match(block, /timeoutMs: 60000/);
    assert.doesNotMatch(block, /await fetch\(/);
    assert.match(bodyHelper, /const text = await response\.text\(\)/);
    assert.match(bodyHelper, /Request timed out after/);
  });

  await test('admin delete unlock uses configured authentication with a bounded request', () => {
    const block = extractBlock(adminSource, 'async function submitAdminDeleteLogin()');
    assert.match(block, /postJsonWithDeadline\('\/local_login'/);
    assert.match(block, /postJsonWithDeadline\([\s\S]*?timeoutMs: 20000/);
    assert.doesNotMatch(block, /armee@2026/);
  });

  await test('branch management and notification email actions use bounded confirmed requests', () => {
    const branchPost = extractBlock(adminSource, 'async function branchMgmtPost(payload)');
    const emailLoad = extractBlock(adminSource, 'async function loadBranchEmails()');
    const emailAdd = extractBlock(adminSource, 'async function addBranchEmail()');
    const emailDelete = extractBlock(adminSource, 'async function deleteBranchEmail(rowNumber)');
    assert.match(branchPost, /postJsonWithDeadline\(/);
    assert.match(emailLoad, /cacheKey: 'get_branch_emails'/);
    assert.match(emailAdd, /postJsonWithDeadline\([\s\S]*?timeoutMs: 30000/);
    assert.match(emailDelete, /postJsonWithDeadline\([\s\S]*?timeoutMs: 30000/);
  });

  await test('school complaint status updates use bounded confirmed requests', () => {
    const bulkBlock = extractBlock(adminSource, 'async function bulkUpdateSchoolStatus(status)');
    const singleBlock = extractBlock(adminSource, 'async function updateSingleSchoolStatus(srNo, status)');
    assert.match(bulkBlock, /postJsonWithDeadline\([\s\S]*?action: 'update_school_complaint_status'/);
    assert.match(singleBlock, /postJsonWithDeadline\([\s\S]*?action: 'update_school_complaint_status'/);
    assert.match(bulkBlock, /timeoutMs: 30000/);
    assert.match(singleBlock, /timeoutMs: 30000/);
    assert.doesNotMatch(bulkBlock, /await fetch\(/);
    assert.doesNotMatch(singleBlock, /await fetch\(/);
  });

  await test('school DISE and project maintenance updates use bounded confirmed requests', () => {
    const diseBlock = extractBlock(adminSource, 'async function promptUpdateDiseCode(srNo, currentDise, schoolName)');
    const healBlock = extractBlock(adminSource, 'async function autoHealSchoolDiseCodes()');
    const projectBlock = extractBlock(adminSource, 'async function runFixProjectCodes()');
    for (const block of [diseBlock, healBlock, projectBlock]) {
      assert.match(block, /postJsonWithDeadline\(/);
      assert.match(block, /timeoutMs: 30000/);
      assert.doesNotMatch(block, /await fetch\(/);
    }
  });

  await test('department and school bulk uploads use bounded confirmed requests', () => {
    const deptBlock = extractBlock(adminSource, 'async function runDeptUpload()');
    const schoolBlock = extractBlock(adminSource, 'async function runSchoolUpload()');
    for (const block of [deptBlock, schoolBlock]) {
      assert.match(block, /postJsonWithDeadline\([\s\S]*?timeoutMs: 60000/);
      assert.doesNotMatch(block, /await fetch\(/);
    }
    assert.match(schoolBlock, /const batchSize = 50/);
    assert.match(schoolBlock, /schoolParsedRows\.slice\(/);
    assert.match(schoolBlock, /out\.retryable/);
    assert.match(schoolBlock, /finalChunk:/);
  });

  await test('department ticket status updates use bounded confirmed requests', () => {
    const bulkBlock = extractBlock(adminSource, 'async function bulkUpdateDeptStatus(action)');
    const singleBlock = extractBlock(adminSource, 'async function submitDeptStatusUpdate(action)');
    const helper = extractBlock(adminSource, 'async function postDepartmentResolution(payload, onSlow)');
    assert.match(helper, /postJsonWithDeadline\(GOOGLE_SCRIPT_URL, payload, \{ timeoutMs: 90000 \}\)/);
    assert.match(helper, /departmentUpdateLanded\(payload\)/, 'a timeout must be verified before reporting failure');
    assert.match(helper, /out\.retryable && attempt < 2/, 'busy responses are retried once');
    for (const block of [bulkBlock, singleBlock]) {
      assert.match(block, /postDepartmentResolution\([\s\S]*?action: 'resolve_department_complaint'/);
      assert.doesNotMatch(block, /await fetch\(/);
    }
  });

  await test('Acer mapping batches retain retries while using bounded confirmed requests', () => {
    const block = extractBlock(adminSource, 'async function runAcerMapping()');
    assert.match(block, /const maxAttempts = 3/);
    assert.match(block, /postJsonWithDeadline\([\s\S]*?action: 'bulk_acer_mapping'/);
    assert.match(block, /timeoutMs: 60000/);
    assert.doesNotMatch(block, /const res = await fetch\(GOOGLE_SCRIPT_URL/);
  });

  await test('complaint deletion changes local data only after a confirmed server response', () => {
    const block = extractBlock(adminSource, 'async function deleteComplaint(origIndex)');
    assert.match(block, /const updatedComplaints = complaintsList\.filter/);
    assert.match(block, /postJsonWithDeadline\('\/update_complaints', updatedComplaints/);
    assert.match(block, /action: 'delete_complaint'/);
    assert.match(block, /result\.status !== 'ok'/);
    assert.match(block, /complaintsList = updatedComplaints/);
    assert.doesNotMatch(block, /mode: 'no-cors'/);
  });

  await test('school master deletion uses a bounded confirmed request', () => {
    const block = extractBlock(adminSource, 'async function deleteSchoolMaster(dise, project)');
    assert.match(block, /postJsonWithDeadline\('\/update_school'/);
    assert.match(block, /postJsonWithDeadline\([\s\S]*?action: 'update_school'/);
    assert.match(block, /timeoutMs: 15000/);
    assert.match(block, /timeoutMs: 30000/);
    assert.doesNotMatch(block, /await fetch\(/);
  });

  await test('school identity edits use one replacement request instead of delete then add', () => {
    const block = extractBlock(adminSource, 'async function submitSchoolModal()');
    assert.match(block, /action: 'replace_school_record'/);
    assert.match(block, /field: 'replaced'/);
    assert.match(block, /oldDise: origDise/);
    assert.match(block, /oldProject: origProj/);
    assert.equal((block.match(/postJsonWithDeadline/g) || []).length, 4);
    assert.doesNotMatch(block, /await fetch\(/);
  });

  await test('archive operations use bounded requests and do not label a failed list load as empty data', () => {
    const listBlock = extractBlock(adminSource, 'async function renderArchiveManager()');
    const archiveBlock = extractBlock(adminSource, 'async function archiveDataRange()');
    const restoreBlock = extractBlock(adminSource, 'async function restoreArchiveRecord(caseId)');
    assert.match(listBlock, /cacheKey: 'get_archive_list'/);
    assert.match(listBlock, /const previousArchiveList = archiveList/);
    assert.match(listBlock, /Archive data is temporarily unavailable/);
    assert.match(archiveBlock, /postJsonWithDeadline\([\s\S]*?timeoutMs: 180000/);
    assert.match(archiveBlock, /result.remaining/, 'bounded archive runs tell the user to continue');
    assert.match(restoreBlock, /postJsonWithDeadline\([\s\S]*?timeoutMs: 60000/);
  });

  await test('consolidated dashboard does not treat one department page as complete data', () => {
    const block = extractBlock(adminSource, 'async function loadUnifiedDashboard()');
    assert.match(block, /deptUsesServerPaging \? unifiedDepartmentRows : deptList/);
    assert.match(block, /get_department_complaints_list/);
    assert.match(block, /unifiedDeptData = d;\s*unifiedDepartmentRows = d;/);
    assert.doesNotMatch(block, /departmentApi\.fetchDepartmentComplaints\(\)/);
  });

  await test('large complaint tables render a bounded page', () => {
    const complaintsBlock = extractBlock(adminSource, 'function renderComplaintsTable');
    const unifiedBlock = extractBlock(adminSource, 'function renderUnifiedList');
    assert.match(complaintsBlock, /const pageSize = 50/);
    assert.match(complaintsBlock, /filtered\.slice\(pageStart, pageStart \+ pageSize\)/);
    assert.match(unifiedBlock, /const pageSize = 50/);
    assert.match(unifiedBlock, /rows\.slice\(pageStart, pageStart \+ pageSize\)/);
    assert.doesNotMatch(unifiedBlock, /slice\(0,\s*500\)/);
    assert.match(adminSource, /id="complaintsPagination"/);
    assert.match(adminSource, /id="unifiedPagination"/);
  });

  await test('admin URL routing covers every operational dashboard', () => {
    assert.match(adminSource, /'unified':\s*'unified'/);
    assert.match(adminSource, /'branchemails':\s*'branchemails'/);
    assert.doesNotMatch(adminSource, /module=undefined/);
  });

  // ── Department update speed + lock scope ──
  function deptFixture(c, opts = {}) {
    const ticketId = '2026/016942', dise = '24221503186';
    const deptRow = Array(30).fill('');
    Object.assign(deptRow, { 0: 'ANAND', 2: 'UMRETH', 7: dise, 8: 'BAJIPURA PRIMARY SCHOOL', 9: ticketId, 12: 'IFP', 13: 'NO DISPLAY', 21: '2026-09-20' });
    const dept = makeSheet('DepartmentComplaints', [Array(30).fill('H'), deptRow]);
    const res = makeSheet('DepartmentResolutions', [vm.runInContext('RES_HEADERS', c).slice()]);
    const mainRow = Array(37).fill('');
    Object.assign(mainRow, { 0: 1, 17: 'OLDSERIAL', 23: 'Pending', 24: ticketId,
      31: '=IMAGE("https://drive.google.com/old-serial")', 32: '=HYPERLINK("https://drive.google.com/old-serial", "🔗 View Serial Photo")', 33: 'https://drive.google.com/old-serial' });
    const main = makeSheet('Complaints', [vm.runInContext('HEADERS', c).slice(), mainRow]);
    const reads = [];
    for (const sh of [dept, res, main]) {
      const orig = sh.getRange.bind(sh);
      sh.getRange = (...a) => { reads.push([sh.name, a[2] || 1, a[3] || 1]); return orig(...a); };
    }
    const db = makeSpreadsheet([dept, res, main]);
    let locked = 0;
    c.SpreadsheetApp = { openById: () => db, flush() {} };
    c.LockService = { getScriptLock: () => ({
      tryLock: () => { if (opts.busy) return false; locked++; return true; },
      waitLock() { throw new Error('global lock must not be used for this action'); },
      releaseLock() {} }) };
    c.CacheService = { getScriptCache: () => ({ get: () => null, put() {}, removeAll() {} }) };
    const token = c.issueDepartmentResolutionToken_(ticketId, dise);
    return { ticketId, dise, dept, res, main, db, reads, token, lockedCount: () => locked };
  }

  await test('department status update bypasses the global lock and reads single rows only', () => {
    const c = makeContext();
    const f = deptFixture(c);
    const out = JSON.parse(c.doPost({ postData: { contents: JSON.stringify({
      action: 'resolve_department_complaint', ticketId: f.ticketId, portalToken: f.token,
      resolutionAction: 'part_request', serialNumber: 'UM0979100431900CED0700',
      suspectedPart: 'IFP PANEL', technicianName: 'Ayan', diagnosisNotes: 'panel glass damaged' }) } }).text);
    assert.equal(out.status, 'ok');
    assert.equal(out.internalStatus, 'PartRequest');
    assert.equal(f.lockedCount(), 1);
    assert.equal(f.res.rows.length, 2);
    assert.equal(f.res.rows[1][15], 'IFP PANEL');
    const m = f.main.rows[1];
    assert.equal(m[23], 'PartRequest');
    assert.equal(m[17], 'UM0979100431900CED0700');
    assert.equal(m[20], 'IFP PANEL');
    assert.match(String(m[31]), /^=IMAGE\("https:\/\/drive.google.com\/old-serial"\)$/, 'existing serial photo formula must survive');
    const wide = f.reads.filter(([, nr, nc]) => nr > 1 && nc > 1);
    assert.deepEqual(wide, [], 'no multi-row, multi-column sheet reads');
  });

  await test('department status update returns a retryable busy error instead of hanging', () => {
    const c = makeContext();
    const f = deptFixture(c, { busy: true });
    const out = JSON.parse(c.doPost({ postData: { contents: JSON.stringify({
      action: 'resolve_department_complaint', ticketId: f.ticketId, portalToken: f.token,
      resolutionAction: 'in_progress' }) } }).text);
    assert.equal(out.status, 'error');
    assert.equal(out.retryable, true);
    assert.equal(out.code, 'busy');
    assert.equal(f.res.rows.length, 1);
  });

  // ── Field principal / contact corrections ──
  function fieldPost(c, db, payload) {
    c.SpreadsheetApp = { openById: () => db, flush() {} };
    c.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) };
    c.CacheService = { getScriptCache: () => ({ get: () => null, put() {} }) };
    return JSON.parse(c.doPost({ postData: { contents: JSON.stringify(Object.assign({
      action: 'field_update_school_contact', dise: '24150800501', modifiedBy: 'Ronak', modifiedPhone: '9876543210'
    }, payload)) } }).text);
  }

  await test('field form can correct principal name and contact number with an audit trail', () => {
    const c = makeContext();
    const db = makeSpreadsheet([]);
    const a = fieldPost(c, db, { field: 'principal', newValue: '  Ashaben  k damor ', oldValue: 'OLD NAME' });
    assert.equal(a.status, 'ok');
    assert.equal(a.newValue, 'ASHABEN K DAMOR');
    const b = fieldPost(c, db, { field: 'mobile', newValue: '+91 95378 53597' });
    assert.equal(b.status, 'ok');
    assert.equal(b.newValue, '9537853597');
    const sheet = db.getSheetByName('SchoolUpdates');
    assert.deepEqual(sheet.rows[1].slice(0, 4), ['24150800501', 'principal', 'ASHABEN K DAMOR', 'OLD NAME']);
    assert.equal(sheet.rows[1][5], '', 'blank project applies to every project row');
    assert.deepEqual(sheet.rows[1].slice(6, 9), ['Ronak', '9876543210', 'field_form']);
    assert.equal(sheet.rows[0][6], 'Modified By');
    const updates = c.getSchoolUpdates(db);
    assert.equal(updates.length, 2);
    assert.equal(updates[1].field, 'mobile');
  });

  await test('field form cannot change school identity or write invalid values', () => {
    const c = makeContext();
    const db = makeSpreadsheet([]);
    for (const [payload, re] of [
      [{ field: 'school', newValue: 'X SCHOOL' }, /Only Principal/],
      [{ field: 'added', newValue: '{}' }, /Only Principal/],
      [{ field: 'mobile', newValue: '12345' }, /10-digit/],
      [{ field: 'principal', newValue: '=HYPERLINK("x")' }, /valid principal/],
      [{ field: 'principal', newValue: 'ASHABEN', modifiedPhone: '' }, /profile/],
      [{ field: 'principal', newValue: 'ASHABEN', dise: '../x' }, /DISE/],
    ]) {
      const out = fieldPost(c, db, payload);
      assert.equal(out.status, 'error');
      assert.match(out.message, re);
    }
    assert.equal(db.getSheetByName('SchoolUpdates'), null, 'nothing written for rejected edits');
  });

  await test('field form Edit buttons save to the live server', () => {
    const block = extractBlock(indexSource, 'function openEditField(field)');
    assert.doesNotMatch(block, /require an administrator/);
    const sync = extractBlock(indexSource, 'async function postSchoolEdit(');
    assert.match(sync, /field_update_school_contact/);
  });

  // ── Resilient complaint loading (behavioural: runs the real admin code) ──
  function makeAdminLoader(responder) {
    const start = adminSource.indexOf('/* ─────────────── COMPLAINT DATA SYNC (resilient)');
    const end = adminSource.indexOf('async function loadSchoolList() {');
    assert.ok(start > 0 && end > start, 'loader section must exist');
    const els = {};
    const el = id => (els[id] ??= { id, value: '', textContent: '', innerHTML: '', style: {}, classList: { toggle() {}, add() {}, remove() {} } });
    const calls = [];
    const ctx = vm.createContext({
      console: { log() {}, warn() {}, error() {} },
      document: { getElementById: el },
      navigator: { onLine: true },
      performance: { now: () => Date.now() },
      setTimeout: (fn) => { fn(); return 0; },
      GOOGLE_SCRIPT_URL: 'https://script.example/exec',
      currentUser: { email: 'admin@example.com' },
      isLocal: () => false,
      logToDebug() {},
      esc: v => String(v),
      normalizeComplaintStatuses: list => list,
      parseDateString: v => (v ? new Date(v) : null),
      adminAuthToken: () => 'tok',
      renderDashboard() { ctx.renders = (ctx.renders || 0) + 1; },
      renderComplaintsTable() {},
      openModal() {},
      adminApi: { async request(url, opts) { calls.push({ url, opts }); return responder(url, calls.length); } },
    });
    ctx.window = ctx;
    vm.runInContext('var complaintsList = [];\n' + adminSource.slice(start, end), ctx);
    return { ctx, els, calls, state: () => vm.runInContext('complaintsLoadState', ctx),
             list: () => vm.runInContext('complaintsList', ctx) };
  }
  const envelope = data => ({ success: true, data, count: data.length, warnings: [], warningCount: 0, meta: { activeCount: data.length, cache: 'miss' } });

  await test('dashboard loads complaints through the v2 envelope with a 90s deadline', async () => {
    const t = makeAdminLoader(url => envelope(url.includes('get_complaints') ? [{ caseId: 'C1', submittedAt: '2026-09-01' }] : []));
    await t.ctx.loadComplaints();
    assert.equal(t.state(), 'live');
    assert.equal(t.list().length, 1);
    assert.ok(t.calls.every(c => /[?&]v=2/.test(c.url) && c.opts.timeoutMs === 90000));
    assert.equal(t.els.complaintsLoadNotice.style.display, 'none');
  });

  await test('the dashboard still works against the previous backend (bare array response)', async () => {
    const t = makeAdminLoader(url => (url.includes('get_complaints') ? [{ caseId: 'C1' }] : []));
    await t.ctx.loadComplaints();
    assert.equal(t.state(), 'live');
    assert.equal(t.list().length, 1);
  });

  await test('a legitimately empty active sheet is reported as loaded, not as a failure', async () => {
    const t = makeAdminLoader(() => envelope([]));
    await t.ctx.loadComplaints();
    assert.equal(t.state(), 'live');
    assert.match(t.els.complaintsLoadNotice.innerHTML, /Data loaded successfully\. 0 complaints/);
  });

  await test('an API failure with no earlier data is shown as unavailable, never as zero', async () => {
    const t = makeAdminLoader(() => ({ success: false, error: { code: 'DATA_SOURCE_ERROR', message: 'Unable to read' } }));
    await t.ctx.loadComplaints();
    assert.equal(t.state(), 'failed');
    assert.match(t.els.complaintsLoadNotice.innerHTML, /Unable to load complaint data[\s\S]*unavailable, not zero[\s\S]*Retry/);
  });

  await test('a failed refresh keeps the last successful data and shows its sync time', async () => {
    let fail = false;
    const t = makeAdminLoader(url => (fail ? Promise.reject(new Error('Request timed out after 90s'))
                                           : envelope(url.includes('get_complaints') ? [{ caseId: 'C1' }, { caseId: 'C2' }] : [])));
    await t.ctx.loadComplaints();
    fail = true;
    await t.ctx.loadComplaints();
    assert.equal(t.state(), 'stale');
    assert.equal(t.list().length, 2);
    assert.match(t.els.complaintsLoadNotice.innerHTML, /Live data unavailable[\s\S]*Showing last successful sync from \d{2}-\d{2}-\d{4} \d{2}:\d{2}/);
  });

  await test('timeouts are not retried but quick server errors are retried once', async () => {
    const timeouts = makeAdminLoader(() => Promise.reject(new Error('Request timed out after 90s')));
    await timeouts.ctx.loadComplaints();
    assert.equal(timeouts.calls.filter(c => c.url.includes('action=get_complaints')).length, 1);
    const busy = makeAdminLoader(() => Promise.reject(new Error('server busy (returned an error page)')));
    await busy.ctx.loadComplaints();
    assert.equal(busy.calls.filter(c => c.url.includes('action=get_complaints')).length, 2);
  });

  await test('incomplete API data is rejected instead of rendered as zero', async () => {
    const t = makeAdminLoader(() => ({ success: true, count: 3 }));
    await t.ctx.loadComplaints();
    assert.equal(t.state(), 'failed');
    assert.match(t.els.complaintsLoadNotice.innerHTML, /incomplete data/);
  });

  await test('dates before the active data pull archived records for that range only', async () => {
    const t = makeAdminLoader(url => url.includes('includeArchive=1')
      ? envelope([{ caseId: 'OLD', submittedAt: '2026-06-02' }, { caseId: 'C1', submittedAt: '2026-09-01' }])
      : envelope(url.includes('get_complaints') ? [{ caseId: 'C1', submittedAt: '2026-09-01' }] : []));
    await t.ctx.loadComplaints();
    t.els.dashFromDate.value = '2026-06-01';
    await t.ctx.onDashDateRangeChange();
    assert.ok(t.calls.some(c => /includeArchive=1&from=2026-06-01/.test(c.url)));
    assert.equal(t.list().length, 2);
    t.els.dashFromDate.value = '';
    await t.ctx.onDashDateRangeChange();
    assert.equal(t.list().length, 1, 'clearing the range returns to active-only data');
  });

  // ── Archive-safe data layer: the 18 operational scenarios ──
  function complaintRow(headers, i, overrides = {}) {
    const row = Array(headers.length).fill('');
    const set = (h, v) => { const k = headers.indexOf(h); if (k !== -1) row[k] = v; };
    set('SR No.', i); set('Submitted At', new Date(Date.UTC(2026, 6, 1) + i * 3600000));
    set('Project', i % 2 ? 'ICT' : 'GK'); set('DISE Code', 24000000000 + (i % 50)); set('District', 'ANAND');
    set('School Name', 'SCHOOL ' + (i % 50)); set('Equipment', 'IFP'); set('Serial Number', 'SN' + i);
    set('Status', i % 3 ? 'Closed' : 'Open'); set('Case ID', 'CASE-' + i);
    for (const [h, v] of Object.entries(overrides)) set(h, v);
    return row;
  }
  function complaintsDb(c, n, opts = {}) {
    const headers = (opts.headers || vm.runInContext('HEADERS', c)).slice();
    const rows = [headers];
    for (let i = 1; i <= n; i++) rows.push(complaintRow(headers, i));
    const active = makeSheet('Complaints', rows);
    const sheets = [active].concat(opts.extra || []);
    return { headers, active, db: makeSpreadsheet(sheets) };
  }
  function v2(c, db, params = {}) {
    c.SpreadsheetApp = { openById: () => db, flush() {} };
    c.requireAdminAuth_ = () => null;
    c.filterRowsForAdminDistricts_ = (ss, rows) => rows;
    return JSON.parse(c.doGet({ parameter: Object.assign({ action: 'get_complaints', v: '2', authToken: 't' }, params) }).text);
  }
  const ids = out => out.data.map(r => r.caseId).sort();

  await test('scenario 1-5: active data loads and any number of old rows can be removed', () => {
    const c = makeContext();
    const { active, db } = complaintsDb(c, 1500);
    let out = v2(c, db);
    assert.equal(out.success, true); assert.equal(out.count, 1500);
    for (const remove of [100, 1000, 399]) {
      active.rows.splice(1, remove); // delete the oldest rows directly under the header
      c.invalidateComplaintCaches_();
      out = v2(c, db);
      assert.equal(out.success, true);
      assert.equal(out.count, active.rows.length - 1);
    }
    assert.equal(out.count, 1, 'header + one open case still works');
    assert.deepEqual(ids(out), ['CASE-1500']);
  });

  await test('scenario 6-8: blank rows, sorting and moved rows do not change the result', () => {
    const c = makeContext();
    const { active, db, headers } = complaintsDb(c, 40);
    const before = ids(v2(c, db));
    active.rows.splice(10, 0, Array(headers.length).fill(''), Array(headers.length).fill(''));
    active.rows.push(Array(headers.length).fill(''));
    const body = active.rows.slice(1).reverse();                 // sort descending
    active.rows = [active.rows[0]].concat(body);
    c.invalidateComplaintCaches_();
    const out = v2(c, db);
    assert.deepEqual(ids(out), before);
    assert.equal(out.meta.blankRowsIgnored, 3);
  });

  await test('scenario 9-10: new columns and re-ordered non-critical columns are mapped by header', () => {
    const c = makeContext();
    const base = vm.runInContext('HEADERS', c).slice();
    const headers = base.slice();
    headers.splice(5, 0, 'Remarks (new column)');                     // inserted column
    const [status] = headers.splice(headers.indexOf('Status'), 1); headers.push(status); // moved column
    const a = headers.indexOf('Address'), p = headers.indexOf('Pin Code');
    [headers[a], headers[p]] = [headers[p], headers[a]];            // swapped columns
    const { db } = complaintsDb(c, 12, { headers });
    const out = v2(c, db);
    assert.equal(out.count, 12);
    const first = out.data.find(r => r.caseId === 'CASE-3');
    assert.equal(first.status, 'Open');
    assert.equal(first.serialNumber, 'SN3');
    assert.deepEqual(out.meta.missingColumns, []);
  });

  await test('scenario 11-13: invalid dates, missing optional data and malformed rows never fail the dataset', () => {
    const c = makeContext();
    const { active, db, headers } = complaintsDb(c, 10);
    active.rows[2][headers.indexOf('Submitted At')] = '31-02-2026';          // impossible date
    active.rows[3][headers.indexOf('Principal Name')] = '';                  // optional data missing
    active.rows[4][headers.indexOf('Case ID')] = { toString() { throw new Error('corrupt cell'); } }; // malformed
    active.rows[5][headers.indexOf('Submitted At')] = '15/07/2026';          // dd/MM/yyyy
    active.rows[6][headers.indexOf('Submitted At')] = '16-07-2026 10:30';    // dd-MM-yyyy HH:mm
    const out = v2(c, db);
    assert.equal(out.success, true);
    assert.equal(out.count, 9, 'only the malformed row is skipped');
    assert.equal(out.meta.invalidRowsSkipped, 1);
    assert.ok(out.warnings.some(w => /date/i.test(w.reason)));
    assert.match(out.data.find(r => r.caseId === 'CASE-5').submittedAt, /^2026-07-1[45]T/);
  });

  await test('scenario 14: normal loads never read the archive; history merges it by range without duplicates', () => {
    const c = makeContext();
    const headers = vm.runInContext('HEADERS', c).slice();
    const archiveRows = [headers];
    for (let i = 1; i <= 3000; i++) archiveRows.push(complaintRow(headers, 100000 + i, { 'Submitted At': new Date(Date.UTC(2026, 4, 1) + i * 60000) }));
    archiveRows.push(complaintRow(headers, 100001, { 'Submitted At': new Date(Date.UTC(2026, 4, 1)) })); // duplicate archive copy
    archiveRows.push(complaintRow(headers, 5)); // a case that is also still active
    const archive = makeSheet('Archive_2026_05', archiveRows);
    let archiveReads = 0;
    const orig = archive.getRange.bind(archive);
    archive.getRange = (...a) => { archiveReads++; return orig(...a); };
    const { db } = complaintsDb(c, 20, { extra: [archive] });
    const normal = v2(c, db);
    assert.equal(normal.count, 20);
    assert.equal(archiveReads, 0, 'daily dashboard must not touch the archive');
    const hist = v2(c, db, { includeArchive: '1', from: '2026-05-01', to: '2026-05-01' });
    assert.equal(hist.meta.source, 'active+archive');
    const archived = hist.data.filter(r => r.source === 'archive');
    assert.ok(archived.length > 0 && archived.length < 3000, 'archive rows are limited to the requested range');
    assert.equal(new Set(hist.data.map(r => r.caseId)).size, hist.data.length, 'no duplicate Case IDs');
  });

  await test('scenario 15: a header-only active sheet is a successful empty result', () => {
    const c = makeContext();
    const { db } = complaintsDb(c, 0);
    const out = v2(c, db);
    assert.equal(out.success, true);
    assert.equal(out.count, 0);
    assert.deepEqual(out.data, []);
  });

  await test('scenario 16: an unavailable sheet returns a DATA_SOURCE_ERROR, not an empty list', () => {
    const c = makeContext();
    const out = v2(c, makeSpreadsheet([]));
    assert.equal(out.success, false);
    assert.equal(out.error.code, 'DATA_SOURCE_ERROR');
    c.SpreadsheetApp = { openById() { throw new Error('Service Spreadsheets timed out'); } };
    c.requireAdminAuth_ = () => null;
    const down = JSON.parse(c.doGet({ parameter: { action: 'get_complaints', v: '2' } }).text);
    assert.equal(down.success, false);
    assert.doesNotMatch(JSON.stringify(down), /Service Spreadsheets/, 'no internal details to users');
  });

  await test('complaint reads are cached and any write invalidates the cache', () => {
    const c = makeContext();
    const { active, db, headers } = complaintsDb(c, 5);
    assert.equal(v2(c, db).meta.cache, 'miss');
    assert.equal(v2(c, db).meta.cache, 'hit');
    active.rows.push(complaintRow(headers, 6));
    c.invalidateComplaintCaches_();
    const out = v2(c, db);
    assert.equal(out.meta.cache, 'miss');
    assert.equal(out.count, 6);
  });

  await test('reading complaints never deletes rows (no side effects in GET)', () => {
    const c = makeContext();
    const { active, db, headers } = complaintsDb(c, 3);
    active.rows.splice(2, 0, headers.slice()); // a pasted duplicate header row
    const before = active.rows.length;
    const out = v2(c, db);
    assert.equal(active.rows.length, before);
    assert.equal(out.count, 3);
  });

  await test('archiving is idempotent and deletes in blocks', () => {
    const c = makeContext();
    const { active, db } = complaintsDb(c, 30);
    let deletes = 0;
    const origDel = active.deleteRows.bind(active);
    active.deleteRows = (a, n) => { deletes++; return origDel(a, n); };
    const out = c.archiveComplaints(db, '2026-07-01', '2026-07-01');
    assert.ok(out.archived > 0);
    assert.equal(deletes, 1, 'contiguous rows are removed in one call');
    // Simulate an interrupted earlier run: the same rows appear in Active again.
    const archive = db.getSheets().find(s => s.name.startsWith('Archive_'));
    const archivedCount = archive.rows.length - 1;
    active.rows.splice(1, 0, ...archive.rows.slice(1).map(r => r.slice()));
    const again = c.archiveComplaints(db, '2026-07-01', '2026-07-01');
    assert.equal(again.skippedDuplicates, archivedCount);
    assert.equal(archive.rows.length - 1, archivedCount, 'no duplicate archive rows');
  });

  await test('status updates rewrite only the changed row', () => {
    const c = makeContext();
    const { active, db, headers } = complaintsDb(c, 50);
    const writes = [];
    const orig = active.getRange.bind(active);
    active.getRange = (...a) => { const r = orig(...a); const sv = r.setValues.bind(r); r.setValues = v => { writes.push(a); return sv(v); }; return r; };
    c.syncSchoolComplaintMasterStatus = () => {};
    const n = c.updateComplaintsStatus(db, [{ caseId: 'CASE-7', status: 'Part Request' }]);
    assert.equal(n, 1);
    assert.deepEqual(writes, [[8, 1, 1, headers.length]]);
    assert.equal(active.rows[7][headers.indexOf('Status')], 'Part Request');
  });

  await test('data health reports reachability, duplicates and overlaps', () => {
    const c = makeContext();
    const headers = vm.runInContext('HEADERS', c).slice();
    const archive = makeSheet('Archive_2026_06', [headers, complaintRow(headers, 1), complaintRow(headers, 1), complaintRow(headers, 900)]);
    const { active, db } = complaintsDb(c, 5, { extra: [archive] });
    active.rows.push(complaintRow(headers, 2)); // duplicate Case ID in active
    active.rows.push(Array(headers.length).fill(''));
    const h = c.apiGetDataHealth_(db);
    assert.equal(h.active.reachable, true);
    assert.equal(h.active.records, 6);
    assert.equal(h.active.duplicateCaseIds, 1);
    assert.equal(h.active.blankRowsIgnored, 1);
    assert.equal(h.archive.duplicateCaseIdsAcrossArchive, 1);
    assert.equal(h.archive.alsoPresentInActive, 1);
  });

  await test('all inline browser scripts compile', () => {
    for (const [name, source] of [['admin.html', adminSource], ['index.html', indexSource]]) {
      const scripts = [...source.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi)];
      assert.ok(scripts.length > 0, `No inline scripts found in ${name}`);
      scripts.forEach((match, i) => new vm.Script(match[1], { filename: `${name}:inline-${i + 1}` }));
    }
  });

  for (const result of results) console.log(`${result.status.toUpperCase()}  ${result.name}${result.error ? ` — ${result.error}` : ''}`);
  const failures = results.filter(r => r.status === 'fail');
  console.log(`\n${results.length - failures.length}/${results.length} regression checks passed.`);
  if (failures.length) process.exitCode = 1;
})();
