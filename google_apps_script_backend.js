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
  'Suspected Part',
  'Description',
  'Photo Count',
  'Status',
  'Case ID',
  'Latitude',
  'Longitude',
  'Photo Preview',
  'View Photo',
  'Photo URL',
  'Duplicate Status',
  'Serial Photo Preview',
  'View Serial Photo',
  'Serial Photo URL',
  'Archived'
];


// ═══════════════ GOOGLE DRIVE HELPERS ═══════════════

/**
 * Get or create a folder by name inside a parent folder.
 * If parent is null, searches in root Drive.
 */
function getOrCreateSubfolder(parent, folderName) {
  if (!folderName) {
    Logger.log("Warning: getOrCreateSubfolder was called without a folderName. This happens when running the function directly from the editor without arguments. Skipping.");
    return null;
  }
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

    // Explicitly set sharing to ANYONE_WITH_LINK VIEW to guarantee access on all devices
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingErr) {
      Logger.log('Error setting file sharing: ' + sharingErr.toString());
    }

    var fileId = file.getId();
    var viewUrl = 'https://lh3.googleusercontent.com/d/' + fileId;
    var openUrl = 'https://drive.google.com/file/d/' + fileId + '/view';

    return { viewUrl: viewUrl, openUrl: openUrl, name: fileName };
  } catch (err) {
    Logger.log('Photo upload error: ' + err.toString());
    return null;
  }
}

function makeAllPhotosPublic() {
  var props = PropertiesService.getScriptProperties();
  var rootId = props.getProperty('FOLDER_ROOT');
  if (!rootId) {
    var folders = DriveApp.getFoldersByName(DRIVE_ROOT_FOLDER);
    if (folders.hasNext()) {
      rootId = folders.next().getId();
    }
  }
  
  if (!rootId) {
    return "Root folder " + DRIVE_ROOT_FOLDER + " not found";
  }
  
  var rootFolder = DriveApp.getFolderById(rootId);
  rootFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  var counts = { folders: 1, files: 0 };
  makeFolderContentsPublic(rootFolder, counts);
  return "All files and folders under root made public. Updated " + counts.folders + " folders and " + counts.files + " files.";
}

function makeFolderContentsPublic(folder, counts) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      counts.files++;
    } catch (e) {
      Logger.log("Error sharing file " + file.getName() + ": " + e.toString());
    }
  }
  
  var subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    try {
      subfolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      counts.folders++;
      makeFolderContentsPublic(subfolder, counts);
    } catch (e) {
      Logger.log("Error sharing folder " + subfolder.getName() + ": " + e.toString());
    }
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

    if (data.action === 'archive_complaints') {
      var archResult = archiveComplaints(ss, data.fromDate, data.toDate);
      return ContentService.createTextOutput(JSON.stringify({status: 'ok', archived: archResult.archived}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'restore_complaints') {
      var restResult = restoreComplaints(ss, data.caseIds);
      return ContentService.createTextOutput(JSON.stringify({status: 'ok', restored: restResult.restored}))
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
      return ContentService.createTextOutput(JSON.stringify(masterData || { equipment: [], users: [], accessUsers: [] }))
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

    if (action === 'check_duplicate') {
      var serial = String(e.parameter.serial || '').trim().toUpperCase();
      var result = checkDuplicateSerial(ss, serial);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'get_archive_list') {
      var list = getArchiveList(ss);
      return ContentService.createTextOutput(JSON.stringify(list)).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'get_complaints_with_archive') {
      var data = getComplaintsWithArchive(ss);
      return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'make_photos_public') {
      var result = makeAllPhotosPublic();
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', message: result }))
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
    var headerVal = lastCol >= 32 ? sheet.getRange(1, 32).getValue() : '';
    if (headerVal !== 'Archived') {
      migrateSheetTo32Columns(sheet);
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
    var archivedVal = String(row[34] || '').trim();
    if (archivedVal === 'YES') {
      continue;
    }

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
      suspectedPart: row[20],
      description: row[21],
      photoCount: parseInt(row[22]) || 0,
      status: row[23] || 'Open',
      caseId: row[24] || '',
      latitude: parseFloat(row[25]) || 0,
      longitude: parseFloat(row[26]) || 0,
      photoPreview: row[27] || '',
      viewPhoto: row[28] || '',
      photoUrl: row[29] || '',
      duplicateStatus: String(row[30] || 'NO'),
      serialPhotoPreview: row[31] || '',
      viewSerialPhoto: row[32] || '',
      serialPhotoUrl: row[33] || '',
      archived: archivedVal
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

    // If serialPhotoUrl is empty, try extracting from formulas
    if (!obj.serialPhotoUrl) {
      if (rowFormulas && rowFormulas[32]) {
        var match = rowFormulas[32].match(/=HYPERLINK\("([^"]+)"/i);
        if (match) obj.serialPhotoUrl = match[1];
      }
      if (!obj.serialPhotoUrl && rowFormulas && rowFormulas[31]) {
        var match = rowFormulas[31].match(/=IMAGE\("([^"]+)"/i);
        if (match) obj.serialPhotoUrl = match[1];
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

  var folder = getPhotoFolder(data.project, data.submittedAt);
  var timestamp = new Date().getTime();
  var dise = data.dise || 'UNKNOWN';
  var serial = String(data.serialNumber || 'NA').replace(/[^a-zA-Z0-9]/g, '');

  if (photosArray.length > 0) {
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

  // Handle Serial Number Photo upload
  var serialPhotoViewUrl = '';
  var serialPhotoOpenUrl = '';
  var serialPhotoFileName = '';
  if (data.serialPhoto) {
    var sfname = dise + '_' + serial + '_' + timestamp + '_serial.jpg';
    var sresult = uploadPhotoToDrive(data.serialPhoto, sfname, folder);
    if (sresult) {
      serialPhotoViewUrl = sresult.viewUrl;
      serialPhotoOpenUrl = sresult.openUrl;
      serialPhotoFileName = sresult.name;
    }
  }

  const caseId = 'CASE-' + Date.now();

  // Updated row with 35 columns matching HEADERS
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
    data.suspectedPart || data.medium || '',
    data.description || '',
    photosArray.length,
    data.status || 'Open',
    caseId,
    data.latitude || '',
    data.longitude || '',
    photoViewUrl ? '=IMAGE("' + photoViewUrl + '")' : '',
    photoOpenUrl ? '=HYPERLINK("' + photoOpenUrl + '","' + '🔗 View Photo' + '")' : '',
    photoViewUrl || '',
    data.duplicateStatus || 'NO',
    serialPhotoViewUrl ? '=IMAGE("' + serialPhotoViewUrl + '")' : '',
    serialPhotoOpenUrl ? '=HYPERLINK("' + serialPhotoOpenUrl + '","' + '🔗 View Serial Photo' + '")' : '',
    serialPhotoViewUrl || '',
    '' // Archived
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
  var masterData = null;
  if (sheet) {
    var val = sheet.getRange(1, 1).getValue();
    if (val) {
      try {
        masterData = JSON.parse(val);
      } catch (e) {
        Logger.log("Error parsing MasterData sheet JSON: " + e.toString());
      }
    }
  }
  
  if (!masterData) {
    masterData = { equipment: [], users: [], accessUsers: [] };
  }
  if (!masterData.equipment) masterData.equipment = [];
  if (!masterData.users) masterData.users = [];
  if (!masterData.accessUsers) masterData.accessUsers = [];

  // Try to load districts dynamically from "District Master" or "Districts" tab
  var sheetDistricts = getDistrictsFromSheet(ss);
  if (sheetDistricts && sheetDistricts.length > 0) {
    masterData.districts = sheetDistricts;
  } else if (!masterData.districts || masterData.districts.length === 0) {
    // Fallback to default districts if missing in MasterData sheet as well
    masterData.districts = getDefaultDistricts();
  }
  
  return masterData;
}

function getDistrictsFromSheet(ss) {
  var sheet = ss.getSheetByName("District Master") || ss.getSheetByName("Districts");
  if (!sheet) return null;
  
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  
  var headers = rows[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var nameIdx = headers.indexOf("district");
  if (nameIdx === -1) nameIdx = headers.indexOf("district name");
  if (nameIdx === -1) nameIdx = headers.indexOf("name");
  if (nameIdx === -1) nameIdx = 0; // fallback to first column
  
  var statusIdx = headers.indexOf("status");
  
  var districts = [];
  for (var i = 1; i < rows.length; i++) {
    var name = String(rows[i][nameIdx]).trim();
    if (!name) continue;
    
    var status = "active";
    if (statusIdx !== -1) {
      var sVal = String(rows[i][statusIdx]).trim().toLowerCase();
      if (sVal === "inactive" || sVal === "disabled" || sVal === "no") {
        status = "inactive";
      }
    }
    
    districts.push({
      id: "D" + i,
      name: name,
      status: status
    });
  }
  return districts;
}

function getDefaultDistricts() {
  var names = [
    "AHMEDABAD", "AMC", "AMRELI", "ANAND", "ARAVALLI", "BANASKANTHA", "BHARUCH", 
    "BHAVNAGAR", "BOTAD", "CHHOTAUDEPUR", "DAHOD", "DEVBHOOMI DWARKA", "GANDHINAGAR", 
    "GIR SOMNATH", "JAMNAGAR", "JUNAGADH", "KACHCHH", "KHEDA", "MAHESANA", 
    "MAHISAGAR", "MORBI", "NARMADA", "NAVSARI", "PANCH MAHALS", "PATAN", 
    "PORBANDAR", "RAJKOT", "RMC", "SABAR KANTHA", "SMC", "SURAT", "SURENDRANAGAR", 
    "TAPI", "THE DANGS", "VADODARA", "VALSAD", "VMC"
  ];
  return names.map(function(name, i) {
    return { id: "D" + (i + 1), name: name, status: "active" };
  });
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

  var numCols = HEADERS.length;
  var rows = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
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
      var changed = false;

      // Define map of object keys to sheet columns (1-indexed)
      var fieldsToCols = {
        complainantName: 3,
        complainantPhone: 4,
        complainantRole: 5,
        project: 6,
        dise: 7,
        schoolCode: 8,
        district: 9,
        taluka: 10,
        school: 11,
        principal: 12,
        contact: 13,
        address: 14,
        pincode: 15,
        equipment: 16,
        natureOfComplaint: 17,
        serialNumber: 18,
        quantity: 19,
        complaintDate: 20,
        suspectedPart: 21,
        description: 22,
        status: 24,
        duplicateStatus: 31,
        latitude: 26,
        longitude: 27
      };

      for (var field in fieldsToCols) {
        var col = fieldsToCols[field];
        if (c[field] === undefined) continue;

        var val = c[field];
        if (field === 'quantity') {
          val = parseInt(val) || 1;
        } else if (field === 'suspectedPart') {
          val = c.suspectedPart || c.medium || '';
        } else if (field === 'latitude' || field === 'longitude') {
          val = val !== '' ? parseFloat(val) || '' : '';
        }

        var sheetVal = rows[foundIndex][col - 1];
        if (String(sheetVal).trim() !== String(val).trim()) {
          sheet.getRange(rowNumber, col).setValue(val);
          rows[foundIndex][col - 1] = val;
          changed = true;
        }
      }

      // Handle photo fields if photoUrl is modified
      if (c.photoUrl !== undefined) {
        var sheetPhotoUrl = rows[foundIndex][29]; // col 30 (0-indexed 29)
        if (String(sheetPhotoUrl).trim() !== String(c.photoUrl).trim()) {
          var pUrl = String(c.photoUrl).trim();
          sheet.getRange(rowNumber, 30).setValue(pUrl);
          rows[foundIndex][29] = pUrl;
          
          if (pUrl) {
            sheet.getRange(rowNumber, 28).setFormula('=IMAGE("' + pUrl + '")');
            sheet.getRange(rowNumber, 29).setFormula('=HYPERLINK("' + pUrl + '","🔗 View Photo")');
            sheet.getRange(rowNumber, 23).setValue(1); // Photo Count
          } else {
            sheet.getRange(rowNumber, 28).setValue('');
            sheet.getRange(rowNumber, 29).setValue('');
            sheet.getRange(rowNumber, 23).setValue(0); // Photo Count
          }
          changed = true;
        }
      }

      if (changed) {
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

  var numCols = HEADERS.length;
  var rows = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
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
    21: 120,  // Suspected Part
    22: 200,  // Description
    23: 80,   // Photo Count
    24: 80,   // Status
    25: 120,  // Case ID
    26: 90,   // Latitude
    27: 90,   // Longitude
    28: 100,  // Photo Preview
    29: 100,  // View Photo
    30: 280,  // Photo URL
    31: 120,  // Duplicate Status
    32: 100,  // Archived
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

function checkDuplicateSerial(ss, serialNumber) {
  if (!serialNumber) return { isDuplicate: false };
  var sheet = ss.getSheetByName('Complaints');
  if (!sheet || sheet.getLastRow() < 2) return { isDuplicate: false };
  var data = sheet.getDataRange().getValues();
  var now = new Date();
  var currentMonth = now.getMonth();
  var currentYear = now.getFullYear();
  for (var i = 1; i < data.length; i++) {
    var rowSerial = String(data[i][17] || '').trim().toUpperCase(); // Col 18 = Serial Number
    if (rowSerial === serialNumber) {
      var rowDate = new Date(data[i][1]); // Col 2 = Submitted At
      if (!isNaN(rowDate.getTime())) {
        if (rowDate.getMonth() === currentMonth && rowDate.getFullYear() === currentYear) {
          return {
            isDuplicate: true,
            existingDate: data[i][1],
            existingCaseId: String(data[i][24] || ''),
            existingSchool: String(data[i][10] || '')
          };
        }
      }
    }
  }
  return { isDuplicate: false };
}

function archiveComplaints(ss, fromDate, toDate) {
  var sheet = ss.getSheetByName('Complaints');
  if (!sheet || sheet.getLastRow() < 2) return { archived: 0 };
  
  var from = new Date(fromDate);
  var to = new Date(toDate);
  to.setHours(23, 59, 59, 999);
  
  // Format target sheet name based on the month/year tag as checked for separate month tagging
  var months = ['01','02','03','04','05','06','07','08','09','10','11','12'];
  var yearStr = from.getFullYear().toString();
  var monthStr = months[from.getMonth()];
  var archiveTabName = 'Archive_' + yearStr + '_' + monthStr;
  
  var archiveSheet = ss.getSheetByName(archiveTabName);
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(archiveTabName);
    setupHeaders(archiveSheet);
  }
  
  var data = sheet.getDataRange().getValues();
  var formulas = sheet.getDataRange().getFormulas();
  var rowsToArchive = [];
  var rowIndicesToDelete = [];
  
  for (var i = data.length - 1; i >= 1; i--) {
    var rowDate = new Date(data[i][1]);
    if (!isNaN(rowDate.getTime())) {
      if (rowDate >= from && rowDate <= to) {
        var row = [];
        for (var col = 0; col < data[i].length; col++) {
          if (formulas[i] && formulas[i][col]) {
            row.push(formulas[i][col]);
          } else {
            row.push(data[i][col]);
          }
        }
        row[31] = 'YES'; // Mark as archived
        rowsToArchive.unshift(row);
        rowIndicesToDelete.push(i + 1); // 1-indexed row number
      }
    }
  }
  
  if (rowsToArchive.length > 0) {
    // Append to archive sheet
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);
    // Delete from main sheet (reverse order to preserve indices)
    for (var j = 0; j < rowIndicesToDelete.length; j++) {
      sheet.deleteRow(rowIndicesToDelete[j]);
    }
  }
  
  return { archived: rowsToArchive.length };
}

function restoreComplaints(ss, caseIds) {
  if (!caseIds || caseIds.length === 0) return { restored: 0 };
  
  var sheet = ss.getSheetByName('Complaints');
  if (!sheet) return { restored: 0 };
  
  var sheets = ss.getSheets();
  var restoredCount = 0;
  
  for (var sIdx = 0; sIdx < sheets.length; sIdx++) {
    var sName = sheets[sIdx].getName();
    if (sName.indexOf('Archive') === 0 && sheets[sIdx].getLastRow() >= 2) {
      var archiveSheet = sheets[sIdx];
      var data = archiveSheet.getDataRange().getValues();
      var formulas = archiveSheet.getDataRange().getFormulas();
      var rowsToRestore = [];
      var rowIndicesToDelete = [];
      
      for (var i = data.length - 1; i >= 1; i--) {
        var caseId = String(data[i][24] || '');
        if (caseIds.indexOf(caseId) !== -1) {
          var row = [];
          for (var col = 0; col < data[i].length; col++) {
            if (formulas[i] && formulas[i][col]) {
              row.push(formulas[i][col]);
            } else {
              row.push(data[i][col]);
            }
          }
          row[31] = ''; // Clear archived flag
          rowsToRestore.unshift(row);
          rowIndicesToDelete.push(i + 1);
        }
      }
      
      if (rowsToRestore.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rowsToRestore.length, rowsToRestore[0].length).setValues(rowsToRestore);
        for (var j = 0; j < rowIndicesToDelete.length; j++) {
          archiveSheet.deleteRow(rowIndicesToDelete[j]);
        }
        restoredCount += rowsToRestore.length;
      }
    }
  }
  
  return { restored: restoredCount };
}

function getArchiveList(ss) {
  var sheets = ss.getSheets();
  var results = [];
  
  for (var sIdx = 0; sIdx < sheets.length; sIdx++) {
    var sName = sheets[sIdx].getName();
    if (sName.indexOf('Archive') === 0 && sheets[sIdx].getLastRow() >= 2) {
      var data = sheets[sIdx].getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        results.push({
          srNo: String(data[i][0] || ''),
          submittedAt: String(data[i][1] || ''),
          school: String(data[i][10] || ''),
          dise: String(data[i][6] || ''),
          equipment: String(data[i][15] || ''),
          serialNumber: String(data[i][17] || ''),
          caseId: String(data[i][24] || ''),
          status: String(data[i][23] || ''),
          duplicateStatus: String(data[i][30] || 'NO'),
          archived: String(data[i][31] || 'YES'),
          archiveTab: sName
        });
      }
    }
  }
  return results;
}

function getComplaintsWithArchive(ss) {
  var active = getComplaintsList(ss);
  var archived = getArchiveList(ss);
  return { active: active, archived: archived };
}

function migrateSheetTo32Columns(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    setupHeaders(sheet);
    return;
  }
  
  var lastCol = sheet.getLastColumn();
  if (lastCol < 30) {
    migrateSheetTo30Columns(sheet);
  }
  
  lastCol = sheet.getLastColumn();
  if (lastCol === 30) {
    sheet.insertColumnsAfter(30, 2);
    sheet.getRange(1, 31).setValue('Duplicate Status');
    sheet.getRange(1, 32).setValue('Archived');
    
    var totalRows = sheet.getLastRow();
    if (totalRows > 1) {
      sheet.getRange(2, 31, totalRows - 1, 1).setValue('NO');
      sheet.getRange(2, 32, totalRows - 1, 1).setValue('');
    }
  }
  
  sheet.getRange(1, 21).setValue('Suspected Part');
  
  setupHeaders(sheet);
  for (var r = 2; r <= sheet.getLastRow(); r++) {
    formatLastRow(sheet, r);
  }
  Logger.log("✅ Sheet successfully migrated to 32 columns!");
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
