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

  if (parent) {
    return parent.createFolder(folderName);
  } else {
    return DriveApp.createFolder(folderName);
  }
}

/**
 * Build folder path: ArMee Complaints Photos / Project / Year / Month
 * Example: ArMee Complaints Photos / GK / 2026 / June
 */
function getPhotoFolder(project, dateStr) {
  var d = dateStr ? new Date(dateStr) : new Date();
  var year = d.getFullYear().toString();
  var months = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];
  var month = months[d.getMonth()];

  var root    = getOrCreateSubfolder(null, DRIVE_ROOT_FOLDER);
  var projDir = getOrCreateSubfolder(root, project || 'General');
  var yearDir = getOrCreateSubfolder(projDir, year);
  var monDir  = getOrCreateSubfolder(yearDir, month);

  return monDir;
}

/**
 * Upload a single base64-encoded photo to Google Drive.
 * Returns { url, name } or null on failure.
 */
function uploadPhotoToDrive(base64Data, fileName, folder) {
  try {
    // Strip data URL prefix if present: "data:image/jpeg;base64,..."
    var raw = base64Data;
    if (raw.indexOf(',') !== -1) {
      raw = raw.split(',')[1];
    }

    var decoded = Utilities.base64Decode(raw);
    var blob = Utilities.newBlob(decoded, 'image/jpeg', fileName);
    var file = folder.createFile(blob);

    // Make file viewable by anyone with the link
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Direct view URL (works with =IMAGE() in Sheets)
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
      var success = deleteComplaintRow(ss, data.caseId);
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
      // Lightweight endpoint — returns only the version timestamp.
      // The complaint form polls this every 15 s to detect changes cheaply.
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
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  var headers = rows[0];
  var data = [];
  var keyMap = {
    'SR No.': 'srNo',
    'Submitted At': 'submittedAt',
    'Complainant Name': 'complainantName',
    'Complainant Phone': 'complainantPhone',
    'Complainant Role': 'complainantRole',
    'Project': 'project',
    'DISE Code': 'dise',
    'School Code': 'schoolCode',
    'District': 'district',
    'Taluka': 'taluka',
    'School Name': 'school',
    'Principal Name': 'principal',
    'Principal Contact': 'contact',
    'Address': 'address',
    'Pin Code': 'pincode',
    'Equipment': 'equipment',
    'Nature of Complaint': 'natureOfComplaint',
    'Serial Number': 'serialNumber',
    'Quantity': 'quantity',
    'Complaint Date': 'complaintDate',
    'Medium': 'medium',
    'Description': 'description',
    'Photo Count': 'photoCount',
    'Status': 'status',
    'Case ID': 'caseId',
    'Latitude': 'latitude',
    'Longitude': 'longitude',
    'Photo Preview': 'photoPreview',
    'View Photo': 'viewPhoto',
    'Photo URL': 'photoUrl',
  };

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    var obj = {};
    headers.forEach(function(header, index) {
      var key = keyMap[header] || header;
      var val = row[index];
      if (header === 'Latitude' || header === 'Longitude') {
        val = parseFloat(val) || 0;
      }
      obj[key] = val;
    });
    data.push(obj);
  }
  return data;
}

function handleSubmitComplaint(ss, data) {
  var sheet = ss.getSheetByName(SHEET_TAB_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TAB_NAME);
    setupHeaders(sheet);
  }

  if (sheet.getLastRow() === 0) {
    setupHeaders(sheet);
  }

  var srNo = sheet.getLastRow();
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
    '',
    '',
    photoViewUrl,
  ];

  sheet.appendRow(row);

  var newRow = sheet.getLastRow();

  if (photoViewUrl) {
    var previewCell = sheet.getRange(newRow, 28);
    previewCell.setFormula('=IMAGE("' + photoViewUrl + '")');

    var linkCell = sheet.getRange(newRow, 29);
    linkCell.setFormula('=HYPERLINK("' + photoOpenUrl + '","' + '\uD83D\uDD17 View Photo' + '")');
  }

  formatLastRow(sheet);

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

function updateComplaintsStatus(ss, complaintsArray) {
  var sheet = ss.getSheetByName(SHEET_TAB_NAME);
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  var rangeCaseIds = sheet.getRange(2, 25, lastRow - 1, 1);
  var caseIds = rangeCaseIds.getValues();

  var rangeStatus = sheet.getRange(2, 24, lastRow - 1, 1);
  var statuses = rangeStatus.getValues();

  var caseMap = {};
  for (var i = 0; i < caseIds.length; i++) {
    var cId = caseIds[i][0];
    if (cId) {
      caseMap[cId] = i;
    }
  }

  var updatedCount = 0;
  for (var j = 0; j < complaintsArray.length; j++) {
    var c = complaintsArray[j];
    var cId = c.caseId;
    if (cId && cId in caseMap) {
      var idx = caseMap[cId];
      if (statuses[idx][0] !== c.status) {
        statuses[idx][0] = c.status;
        updatedCount++;
      }
    }
  }

  if (updatedCount > 0) {
    rangeStatus.setValues(statuses);
  }
  return updatedCount;
}

function deleteComplaintRow(ss, caseId) {
  var sheet = ss.getSheetByName(SHEET_TAB_NAME);
  if (!sheet) return false;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  var rangeCaseIds = sheet.getRange(2, 25, lastRow - 1, 1);
  var caseIds = rangeCaseIds.getValues();

  for (var i = 0; i < caseIds.length; i++) {
    if (caseIds[i][0] === caseId) {
      // Row index in sheet is 1-indexed, and we started at row 2 (index 0 corresponds to row 2)
      var rowIndex = i + 2;
      sheet.deleteRow(rowIndex);
      return true;
    }
  }
  return false;
}


// ═══════════════ SHEET FORMATTING ═══════════════

function setupHeaders(sheet) {
  sheet.appendRow(HEADERS);
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

  // Set default row height for photo preview
  sheet.setRowHeightsForced(1, 1, 30);
}

function formatLastRow(sheet) {
  var row   = sheet.getLastRow();
  var range = sheet.getRange(row, 1, 1, HEADERS.length);
  // Alternate row shading
  var bg = row % 2 === 0 ? '#f0f4ff' : '#ffffff';
  range.setBackground(bg);
  range.setVerticalAlignment('middle');

  // Set row height to 80px for photo thumbnail visibility
  sheet.setRowHeightsForced(row, 1, 80);

  // Center-align Photo Preview column (Col 28)
  sheet.getRange(row, 28).setHorizontalAlignment('center');
  // Style View Photo link (Col 29)
  sheet.getRange(row, 29).setFontColor('#1a56db').setFontWeight('bold');
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
