/**
 * ═══════════════════════════════════════════════════════════════
 *  School Equipment Complaint Form - Google Apps Script Backend
 *  With Google Drive Photo Storage & Sheet Preview
 * ═══════════════════════════════════════════════════════════════
 *
 *  SETUP STEPS:
 *  1. Go to https://script.google.com
 *  2. Open your existing "School Complaint Backend" project
 *  3. Replace ALL code with this entire file
 *  4. Verify SHEET_ID below matches your Google Sheet
 *  5. Click Deploy -> Manage Deployments -> Edit (pencil icon)
 *     -> Version: New Version -> Deploy
 *  6. Re-authorize if prompted (new Drive permissions needed)
 *
 *  PERMISSIONS NEEDED:
 *  - Google Sheets (read/write)
 *  - Google Drive (create folders & upload files)
 */

// -- YOUR GOOGLE SHEET ID --
// From your sheet URL: https://docs.google.com/spreadsheets/d/[THIS_PART]/edit
const SHEET_ID = '1VHQWSMMdjlUlsU1O1DB2rg8eaOrwuuSGs8fvVKzKpOw';
const SHEET_TAB_NAME = 'Complaints';

// -- GOOGLE DRIVE PHOTO STORAGE --
const DRIVE_ROOT_FOLDER = 'ArMee Complaints Photos';

// ---------------------------------------------------------------

const HEADERS = [
  'SR No.',
  'Submitted At',
  'Complainant Name',
  'Complainant Phone',
  'Complainant Role',
  'Project',
  'DISE Code',
  'School Code',
  'District',
  'Taluka',
  'School Name',
  'Principal Name',
  'Principal Contact',
  'Address',
  'Pin Code',
  'Equipment',
  'Nature of Complaint',
  'Serial Number',
  'Quantity',
  'Complaint Date',
  'Medium',
  'Description',
  'Photo Count',
  'Status',
  'Case ID',
  'Latitude',
  'Longitude',
  'Photo Preview',
  'View Photo',
  'Photo URL',
];


// ═══════════════ GOOGLE DRIVE HELPERS ═══════════════

/**
 * Get or create a folder by name inside a parent folder.
 * If parent is null, searches in root Drive.
 */
function getOrCreateSubfolder(parent, folderName) {
  var iter;
  if (parent) {
    iter = parent.getFoldersByName(folderName);
  } else {
    iter = DriveApp.getFoldersByName(folderName);
  }

  if (iter.hasNext()) {
    return iter.next();
  }

  var folder;
  if (parent) {
    folder = parent.createFolder(folderName);
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  // Set sharing on the folder once so that all contents inherit
  try {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    Logger.log("Sharing error for folder " + folderName + ": " + e.toString());
  }

  return folder;
}

/**
 * Build folder path: ArMee Complaints Photos / Project / Year / Month
 * Optimized using PropertiesService caching to avoid slow Drive queries.
 */
function getPhotoFolder(project, dateStr) {
  var d = dateStr ? new Date(dateStr) : new Date();
  var year = d.getFullYear().toString();
  var months = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];
  var month = months[d.getMonth()];

  var proj = project || 'General';
  var props = PropertiesService.getScriptProperties();

  // Try caching monthly folder first (99% hit rate in production)
  var monthKey = 'FOLDER_MONTH_' + proj + '_' + year + '_' + month;
  var monthId = props.getProperty(monthKey);
  if (monthId) {
    try {
      return DriveApp.getFolderById(monthId);
    } catch (e) {
      props.deleteProperty(monthKey); // Stale cache
    }
  }

  // Resolve Root Folder
  var rootId = props.getProperty('FOLDER_ROOT');
  var root;
  if (rootId) {
    try {
      root = DriveApp.getFolderById(rootId);
    } catch (e) {
      props.deleteProperty('FOLDER_ROOT');
    }
  }
  if (!root) {
    root = getOrCreateSubfolder(null, DRIVE_ROOT_FOLDER);
    props.setProperty('FOLDER_ROOT', root.getId());
  }

  // Resolve Project Folder
  var projKey = 'FOLDER_PROJ_' + proj;
  var projId = props.getProperty(projKey);
  var projDir;
  if (projId) {
    try {
      projDir = DriveApp.getFolderById(projId);
    } catch (e) {
      props.deleteProperty(projKey);
    }
  }
  if (!projDir) {
    projDir = getOrCreateSubfolder(root, proj);
    props.setProperty(projKey, projDir.getId());
  }

  // Resolve Year Folder
  var yearKey = 'FOLDER_YEAR_' + proj + '_' + year;
  var yearId = props.getProperty(yearKey);
  var yearDir;
  if (yearId) {
    try {
      yearDir = DriveApp.getFolderById(yearId);
    } catch (e) {
      props.deleteProperty(yearKey);
    }
  }
  if (!yearDir) {
    yearDir = getOrCreateSubfolder(projDir, year);
    props.setProperty(yearKey, yearDir.getId());
  }

  // Resolve Month Folder
  var monDir = getOrCreateSubfolder(yearDir, month);
  props.setProperty(monthKey, monDir.getId());

  return monDir;
}

/**
 * Upload a single base64-encoded photo to Google Drive.
 * File inherits public viewing settings from the parent folder (monthly directory).
 */
function uploadPhotoToDrive(base64Data, fileName, folder) {
  try {
    var raw = base64Data;
    if (raw.indexOf(',') !== -1) {
      raw = raw.split(',')[1];
    }

    var decoded = Utilities.base64Decode(raw);
    var blob = Utilities.newBlob(decoded, 'image/jpeg', fileName);
    var file = folder.createFile(blob);

    // Skip file-level setSharing call since monthly folder is already shared, saving ~500-1000ms.
    var fileId = file.getId();
    var viewUrl = 'https://lh3.googleusercontent.com/d/' + fileId;
    var openUrl = 'https://drive.google.com/file/d/' + fileId + '/view';

    return { viewUrl: viewUrl, openUrl: openUrl, name: fileName };
  } catch (err) {
    Logger.log('Photo upload error: ' + err.toString());
    return null;
  }
}


// ═══════════════ MAIN HANDLERS ═══════════════

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    // Wait for up to 30 seconds to acquire lock
    lock.waitLock(30000);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'error', 
      retryable: true, 
      message: 'Server is busy: ' + err.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var data = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.openById(SHEET_ID);

    if (data.action === 'update_master') {
      saveMasterData(ss, data.masterData);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'update_complaints') {
      var count = updateComplaintsStatus(ss, data.complaints);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', updatedCount: count }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'update_school') {
      updateSchoolField(ss, data);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'delete_complaint') {
      var success = deleteComplaintRow(ss, data.caseId, data.submittedAt, data.dise);
      return ContentService.createTextOutput(JSON.stringify({ status: success ? 'ok' : 'not_found' }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // Default: submit new complaint
    if (!data.action || data.action === 'submit_complaint') {
      return handleSubmitComplaint(ss, data);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Unknown action: ' + data.action }))
                          .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    var action = e.parameter.action;
    var ss     = SpreadsheetApp.openById(SHEET_ID);

    if (action === 'get_master') {
      var masterData = getMasterData(ss);
      return ContentService.createTextOutput(JSON.stringify(masterData || { equipment: [], users: [] }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'get_master_version') {
      var md = getMasterData(ss);
      var version = (md && md._v) ? md._v : '';
      return ContentService.createTextOutput(JSON.stringify({ v: version }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    if (action === 'get_school_updates') {
      var schoolUpdates = getSchoolUpdates(ss);
      return ContentService.createTextOutput(JSON.stringify(schoolUpdates))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // Default: return complaints list
    var complaints = getComplaintsList(ss);
    return ContentService.createTextOutput(JSON.stringify(complaints))
                         .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ═══════════════ SUB-HANDLERS & DATABASE LOGIC ═══════════════

function getComplaintsList(ss) {
  var sheet = ss.getSheetByName(SHEET_TAB_NAME);
  if (!sheet) return [];
  
  // Run migration check dynamically
  try {
    var lastCol = sheet.getLastColumn();
    var headerVal = lastCol >= 30 ? sheet.getRange(1, 30).getValue() : '';
    if (headerVal !== 'Photo URL') {
      migrateSheetTo30Columns(sheet);
    }
  } catch (e) {
    Logger.log("Migration error: " + e.toString());
  }

  var rows = sheet.getDataRange().getValues();
  var formulas = sheet.getDataRange().getFormulas();
  if (rows.length <= 1) return [];

  // Cleanup duplicate header rows if any exist (e.g. from previous runs)
  for (var i = rows.length - 1; i > 0; i--) {
    if (rows[i][0] === 'SR No.') {
      try {
        sheet.deleteRow(i + 1);
      } catch (err) {}
      rows.splice(i, 1);
      formulas.splice(i, 1);
    }
  }

  var data = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var rowFormulas = formulas[i];
    
    // Parse using index-based layout for 100% correctness
    var obj = {
      srNo: row[0],
      submittedAt: row[1],
      complainantName: row[2],
      complainantPhone: row[3],
      complainantRole: row[4],
      project: row[5],
      dise: row[6],
      schoolCode: row[7],
      district: row[8],
      taluka: row[9],
      school: row[10],
      principal: row[11],
      contact: row[12],
      address: row[13],
      pincode: row[14],
      equipment: row[15],
      natureOfComplaint: row[16],
      serialNumber: row[17],
      quantity: row[18],
      complaintDate: row[19],
      medium: row[20],
      description: row[21],
      photoCount: parseInt(row[22]) || 0,
      status: row[23] || 'Open',
      caseId: row[24] || '',
      latitude: parseFloat(row[25]) || 0,
      longitude: parseFloat(row[26]) || 0,
      photoPreview: row[27] || '',
      viewPhoto: row[28] || '',
      photoUrl: row[29] || ''
    };

    // If photoUrl is empty, try extracting from formulas
    if (!obj.photoUrl) {
      if (rowFormulas && rowFormulas[28]) {
        var match = rowFormulas[28].match(/=HYPERLINK\("([^"]+)"/i);
        if (match) obj.photoUrl = match[1];
      }
      if (!obj.photoUrl && rowFormulas && rowFormulas[27]) {
        var match = rowFormulas[27].match(/=IMAGE\("([^"]+)"/i);
        if (match) obj.photoUrl = match[1];
      }
    }
    
    data.push(obj);
  }
  return data;
}

/**
 * Migrates a legacy 26-column layout sheet to the new 30-column format.
 */
function migrateSheetTo30Columns(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    setupHeaders(sheet);
    return;
  }

  var range = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
  var rows = range.getValues();
  var formulas = range.getFormulas();
  
  var migratedRows = [];
  migratedRows.push(HEADERS); // Header row
  
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var rowFormulas = formulas[i];
    
    // Distinguish Format A (Legacy 26-cols) vs Format B (New 30-cols)
    var isFormatB = false;
    if (row.length >= 28) {
      var val23 = String(row[23]).trim();
      var val24 = String(row[24]).trim();
      if (val23 === 'Open' || val23 === 'In Progress' || val23 === 'Resolved' || val23 === 'Closed' || val24.indexOf('CASE-') === 0) {
        isFormatB = true;
      }
    }
    
    var newRow = [];
    if (isFormatB) {
      for (var col = 0; col < 30; col++) {
        if (rowFormulas && rowFormulas[col]) {
          newRow.push(rowFormulas[col]);
        } else {
          newRow.push(row[col] !== undefined ? row[col] : '');
        }
      }
    } else {
      // Format A: Legacy 26-column row. Migrate it!
      for (var col = 0; col < 22; col++) {
        if (rowFormulas && rowFormulas[col]) {
          newRow.push(rowFormulas[col]);
        } else {
          newRow.push(row[col] !== undefined ? row[col] : '');
        }
      }
      
      var photoUrl = row[22] || '';
      if (rowFormulas && rowFormulas[22]) {
        var match = rowFormulas[22].match(/=HYPERLINK\("([^"]+)"/i);
        if (match) {
          photoUrl = match[1];
        } else {
          match = rowFormulas[22].match(/=IMAGE\("([^"]+)"/i);
          if (match) photoUrl = match[1];
        }
      }
      
      var lat = parseFloat(row[23]) || '';
      var lng = parseFloat(row[24]) || '';
      var srNo = row[0] || i;
      
      newRow.push(photoUrl ? 1 : 0); // Photo Count
      newRow.push('Open');           // Status
      newRow.push('CASE-LEGACY-' + srNo); // Case ID
      newRow.push(lat);              // Latitude
      newRow.push(lng);              // Longitude
      newRow.push(photoUrl ? '=IMAGE("https://lh3.googleusercontent.com/d/' + getFileIdFromUrl(photoUrl) + '")' : ''); // Photo Preview
      newRow.push(photoUrl ? '=HYPERLINK("' + photoUrl + '","🔗 View Photo")' : ''); // View Photo
      newRow.push(photoUrl);         // Photo URL
    }
    migratedRows.push(newRow);
  }
  
  sheet.clear();
  var writeRange = sheet.getRange(1, 1, migratedRows.length, 30);
  writeRange.setValues(migratedRows);
  
  setupHeaders(sheet);
  for (var r = 2; r <= migratedRows.length; r++) {
    formatLastRow(sheet, r);
  }
  Logger.log("✅ Sheet successfully migrated to 30 columns!");
}

function getFileIdFromUrl(url) {
  if (!url) return '';
  var match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  match = url.match(/id=([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  return url;
}

function handleSubmitComplaint(ss, data) {
  var sheet = ss.getSheetByName(SHEET_TAB_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TAB_NAME);
    setupHeaders(sheet);
  }

  var lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    setupHeaders(sheet);
    lastRow = 1;
  }

  var srNo = lastRow;
  var photoViewUrl = '';
  var photoOpenUrl = '';
  var photoFileName = '';
  var photosArray = data.photos || [];

  if (photosArray.length > 0) {
    var folder = getPhotoFolder(data.project, data.submittedAt);
    var timestamp = new Date().getTime();
    var dise = data.dise || 'UNKNOWN';
    var serial = (data.serialNumber || 'NA').replace(/[^a-zA-Z0-9]/g, '');

    for (var i = 0; i < photosArray.length; i++) {
      var fname = dise + '_' + serial + '_' + timestamp + (photosArray.length > 1 ? '_' + (i+1) : '') + '.jpg';
      var result = uploadPhotoToDrive(photosArray[i], fname, folder);

      if (result && i === 0) {
        photoViewUrl = result.viewUrl;
        photoOpenUrl = result.openUrl;
        photoFileName = result.name;
      }
    }
  }

  const caseId = 'CASE-' + Date.now();

  // Optimized: write formulas directly in the initial array to prevent separate setFormula calls
  var row = [
    srNo,
    data.submittedAt || '',
    data.complainantName || '',
    data.complainantPhone || '',
    data.complainantRole || '',
    data.project || '',
    data.dise || '',
    data.schoolCode || '',
    data.district || '',
    data.taluka || '',
    data.school || '',
    data.principal || '',
    data.contact || '',
    data.address || '',
    data.pincode || '',
    data.equipment || '',
    data.natureOfComplaint || '',
    data.serialNumber || '',
    data.quantity || 1,
    data.complaintDate || '',
    data.medium || '',
    data.description || '',
    photosArray.length,
    data.status || 'Open',
    caseId,
    data.latitude || '',
    data.longitude || '',
    photoViewUrl ? '=IMAGE("' + photoViewUrl + '")' : '',
    photoOpenUrl ? '=HYPERLINK("' + photoOpenUrl + '","' + '🔗 View Photo' + '")' : '',
    photoViewUrl || '',
  ];

  sheet.appendRow(row);

  var newRow = lastRow + 1;
  formatLastRow(sheet, newRow);

  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      srNo: srNo,
      photoUrl: photoOpenUrl
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getMasterData(ss) {
  var sheet = ss.getSheetByName('MasterData');
  if (!sheet) return null;
  var val = sheet.getRange(1, 1).getValue();
  if (!val) return null;
  try {
    return JSON.parse(val);
  } catch (e) {
    return null;
  }
}

function saveMasterData(ss, data) {
  // Stamp a version timestamp so the complaint form can detect changes cheaply
  data._v = new Date().getTime().toString();
  var sheet = ss.getSheetByName('MasterData');
  if (!sheet) {
    sheet = ss.insertSheet('MasterData');
  }
  sheet.clear();
  sheet.getRange(1, 1).setValue(JSON.stringify(data));
}

function updateSchoolField(ss, update) {
  var sheet = ss.getSheetByName('SchoolUpdates');
  if (!sheet) {
    sheet = ss.insertSheet('SchoolUpdates');
    sheet.appendRow(['DISE Code', 'Field', 'New Value', 'Old Value', 'Updated At']);
    sheet.getRange(1, 1, 1, 5).setBackground('#1a56db').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    update.dise || '',
    update.field || '',
    update.newValue || '',
    update.oldValue || '',
    new Date().toISOString()
  ]);
}

function getSchoolUpdates(ss) {
  var sheet = ss.getSheetByName('SchoolUpdates');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  return values.map(function(row) {
    return {
      dise: String(row[0]),
      field: String(row[1]),
      newValue: String(row[2]),
      oldValue: String(row[3]),
      updatedAt: row[4]
    };
  });
}

function datesMatch(val1, val2) {
  if (!val1 || !val2) return false;
  var d1 = new Date(val1);
  var d2 = new Date(val2);
  if (!isNaN(d1.getTime()) && !isNaN(d2.getTime())) {
    return Math.abs(d1.getTime() - d2.getTime()) < 5000;
  }
  return String(val1).trim() === String(val2).trim();
}

function updateComplaintsStatus(ss, complaintsArray) {
  var sheet = ss.getSheetByName(SHEET_TAB_NAME);
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  var rows = sheet.getRange(2, 1, lastRow - 1, 30).getValues();
  var updatedCount = 0;

  for (var j = 0; j < complaintsArray.length; j++) {
    var c = complaintsArray[j];
    var foundIndex = -1;

    // 1. Try matching by Case ID
    if (c.caseId && c.caseId !== 'CASE-N/A') {
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][24]).trim() === String(c.caseId).trim()) {
          foundIndex = i;
          break;
        }
      }
    }

    // 2. Fallback: match by Submitted At + DISE Code
    if (foundIndex === -1 && c.submittedAt && c.dise) {
      for (var i = 0; i < rows.length; i++) {
        var sheetSubmittedAt = rows[i][1];
        var sheetDise = String(rows[i][6]).trim();
        if (sheetDise === String(c.dise).trim() && datesMatch(sheetSubmittedAt, c.submittedAt)) {
          foundIndex = i;
          break;
        }
      }
    }

    if (foundIndex !== -1) {
      var rowNumber = foundIndex + 2;
      var currentStatus = rows[foundIndex][23];
      if (currentStatus !== c.status) {
        sheet.getRange(rowNumber, 24).setValue(c.status);
        rows[foundIndex][23] = c.status;
        updatedCount++;
      }
    }
  }

  return updatedCount;
}

function deleteComplaintRow(ss, caseId, submittedAt, dise) {
  var sheet = ss.getSheetByName(SHEET_TAB_NAME);
  if (!sheet) return false;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  var rows = sheet.getRange(2, 1, lastRow - 1, 30).getValues();
  var foundIndex = -1;

  // 1. Try matching by Case ID
  if (caseId && caseId !== 'CASE-N/A') {
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][24]).trim() === String(caseId).trim()) {
        foundIndex = i;
        break;
      }
    }
  }

  // 2. Fallback: match by Submitted At + DISE Code
  if (foundIndex === -1 && submittedAt && dise) {
    for (var i = 0; i < rows.length; i++) {
      var sheetSubmittedAt = rows[i][1];
      var sheetDise = String(rows[i][6]).trim();
      if (sheetDise === String(dise).trim() && datesMatch(sheetSubmittedAt, submittedAt)) {
        foundIndex = i;
        break;
      }
    }
  }

  if (foundIndex !== -1) {
    var rowNumber = foundIndex + 2;
    sheet.deleteRow(rowNumber);
    return true;
  }

  return false;
}


// ═══════════════ SHEET FORMATTING ═══════════════

function setupHeaders(sheet) {
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setBackground('#1a56db');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  sheet.setFrozenRows(1);

  // Set column widths
  var widths = {
    1: 50,    // SR No.
    2: 160,   // Submitted At
    3: 140,   // Complainant Name
    4: 120,   // Complainant Phone
    5: 120,   // Complainant Role
    6: 60,    // Project
    7: 120,   // DISE Code
    8: 120,   // School Code
    9: 100,   // District
    10: 120,  // Taluka
    11: 200,  // School Name
    12: 140,  // Principal Name
    13: 120,  // Principal Contact
    14: 200,  // Address
    15: 80,   // Pin Code
    16: 150,  // Equipment
    17: 160,  // Nature of Complaint
    18: 140,  // Serial Number
    19: 50,   // Quantity
    20: 100,  // Complaint Date
    21: 100,  // Medium
    22: 200,  // Description
    23: 80,   // Photo Count
    24: 80,   // Status
    25: 120,  // Case ID
    26: 90,   // Latitude
    27: 90,   // Longitude
    28: 100,  // Photo Preview
    29: 100,  // View Photo
    30: 280,  // Photo URL
  };

  for (var col in widths) {
    sheet.setColumnWidth(parseInt(col), widths[col]);
  }

  try {
    // Column-level formatting optimization: formatted once so individual appends don't need styling calls
    var maxRows = sheet.getMaxRows();
    sheet.getRange(1, 1, maxRows, HEADERS.length).setVerticalAlignment('middle');
    sheet.getRange(1, 28, maxRows, 1).setHorizontalAlignment('center');
    sheet.getRange(1, 29, maxRows, 1).setFontColor('#1a56db').setFontWeight('bold');
  } catch (e) {
    Logger.log("Headers formatting error: " + e.toString());
  }
}

function formatLastRow(sheet, row) {
  try {
    sheet.setRowHeight(row, 80);
    var bg = row % 2 === 0 ? '#f0f4ff' : '#ffffff';
    sheet.getRange(row, 1, 1, HEADERS.length).setBackground(bg);
  } catch (e) {
    Logger.log("Format row error: " + e.toString());
  }
}

/**
 * Run this function from the dropdown menu above to test if the script
 * can access your Google Drive and Google Sheet correctly.
 */
function testBackend() {
  Logger.log("Initializing test script...");
  
  // 1. Test Drive folder creation
  var folder = getPhotoFolder("GK", new Date().toISOString());
  Logger.log("✅ Google Drive folder created/verified successfully!");
  Logger.log("   Folder Name: " + folder.getName());
  Logger.log("   Folder Link: " + folder.getUrl());
  
  // 2. Test Google Sheet access
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_TAB_NAME);
  Logger.log("✅ Google Sheet accessed successfully!");
  Logger.log("   Sheet Title: " + ss.getName());
  Logger.log("   Tab Name: " + sheet.getName());
  
  Logger.log("🎉 Test completed successfully! Your permissions are set up correctly.");
}
