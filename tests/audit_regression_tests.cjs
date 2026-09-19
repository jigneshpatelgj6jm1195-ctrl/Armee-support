const assert = require('assert/strict');
const crypto = require('crypto');
const fs = require('fs');
const vm = require('vm');

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
  const context = vm.createContext({
    console: { log() {}, warn() {}, error() {} },
    Logger: { log() {} },
    ContentService: {
      MimeType: { JSON: 'json' },
      createTextOutput(text) { return { text, setMimeType() { return this; } }; },
    },
    LockService: { getScriptLock() { return { waitLock() {}, releaseLock() {} }; } },
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
      newBlob(value) { return { getDataAsString: () => bytes(value).toString('utf8') }; },
      getUuid() { uuidCounter += 1; return `00000000-0000-4000-8000-${String(uuidCounter).padStart(12, '0')}`; },
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
    });
    vm.runInContext(extractBlock(adminSource, 'async function saveAll'), context);
    await vm.runInContext('saveAll()', context);
    assert.equal(context.isDirty, true);
    assert.equal(messages.some(m => m.type === 'success'), false);
    assert.equal(messages.some(m => m.type === 'warn'), true);
  });

  await test('downloadable source contains no embedded import or fallback admin secret', () => {
    assert.equal(/ArmeeICT-[A-Za-z0-9_-]+/.test(adminSource), false);
    assert.equal(/var\s+SUPER_ADMIN\s*=/.test(backendSource), false);
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
    assert.match(block, /action=get_complaints&authToken=/);
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
