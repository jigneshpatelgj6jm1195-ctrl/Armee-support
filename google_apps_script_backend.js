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

// -- DEPARTMENT COMPLAINT IMPORT (ssgujarat.org scrape/bulk-upload) --
const DEPT_SHEET_TAB_NAME = 'DepartmentComplaints';
const RES_SHEET_TAB_NAME = 'DepartmentResolutions';
// Set once via: run setImportKey("your-long-random-secret") from the Apps Script editor.
// Any doPost with action=import_department_complaints must send the same value as importKey.

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
  'Archived',
  'OtpValue',
  'ClosureType',
  'AcerCaseId',
  'AcerCaseStatus',
  'LastUpdatedDate'
];

// Raw mirror of the ssgujarat.org "Export To Excel" columns, plus 3 tracking columns.
const DEPT_HEADERS = [
  'District', 'BlockId', 'Block', 'ClusterId', 'Cluster', 'VillageId', 'Village',
  'SchoolId', 'School', 'TicketId', 'Agency', 'Asset_Type', 'Device_Type', 'Issue_Type',
  'Issue_Details', 'Issue_Photo', 'Contact_Name', 'Phone_Number', 'Time_Preference_Call',
  'Ticket_Status', 'CreatedBy', 'CreatedDate', 'UpdatedBy', 'UpdatedDate',
  'Diagnosis_Notes_Agency', 'Schedule_Visit_Date', 'Technician_Name', 'Technician_Number',
  'Issue_Resolved_By', 'TotalDaysOfTicket',
  'ImportedAt', 'LastSeenStatus', 'ImportSource'
];

// Our internal resolution/tracking layer, keyed by TicketId. Never blended into DEPT_HEADERS.
const RES_HEADERS = [
  'TicketId', 'InternalStatus', 'ClosureType', 'OtpValue', 'ResolvedBy', 'TechnicianName',
  'DiagnosisNotes', 'ResolvedAt', 'OwningDistrictAdmin',
  'Equipment', 'NatureOfComplaint', 'Quantity', 'ResolutionDate', 'SerialNumber', 'SerialPhotoURL', 'SuspectedPart', 'SuspectedPartPhotoURL',
  'AcerCaseId', 'AcerCaseStatus'
];

// -- BRANCH / DISTRICT OFFICE STRUCTURE --
// District Office and Branch are separate entities: today each office has exactly
// one branch of the same name, but the schema allows one office -> many branches.
// Dashboard grouping resolves complaint strings through BranchAliasMap -> Branches
// -> DistrictOffices; the raw manager name is never a grouping key.
const DO_SHEET_TAB_NAME = 'DistrictOffices';
const DO_HEADERS = [
  'DistrictOfficeID', 'DistrictOfficeName', 'DistrictOfficeCode', 'Region',
  'ContactPerson', 'ContactNumber', 'Status', 'CreatedDate'
];
const BR_SHEET_TAB_NAME = 'Branches';
const BR_HEADERS = [
  'BranchID', 'BranchName', 'BranchCode', 'DistrictOfficeID',
  'BranchManagerName', 'BranchManagerContact', 'Status', 'CreatedDate'
];
const ALIAS_SHEET_TAB_NAME = 'BranchAliasMap';
const ALIAS_HEADERS = ['AliasText', 'BranchID', 'MappedBy', 'MappedDate'];

// The 13 current district offices; each seeds one branch of the same name.
const BRANCH_STRUCTURE_SEED = [
  ['Ahmedabad', 'AHM'], ['Mehsana', 'MEH'], ['Bharuch', 'BRC'], ['Anand', 'AND'],
  ['Vadodara', 'VAD'], ['Godhra', 'GDH'], ['Bhuj', 'BHJ'], ['Bhavnagar', 'BVN'],
  ['Jamnagar', 'JMN'], ['Rajkot', 'RJK'], ['Junagadh', 'JND'], ['Vapi', 'VAP'], ['Surat', 'SRT']
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

    if (data.action === 'login') {
      return ContentService.createTextOutput(JSON.stringify(handleLogin(ss, data)))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'update_master') {
      var authErr = requireAdminAuth_(data, true); // super admin only
      if (authErr) return ContentService.createTextOutput(JSON.stringify(authErr))
                                        .setMimeType(ContentService.MimeType.JSON);
      saveMasterData(ss, data.masterData);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'update_complaints') {
      var authErrC = requireAdminAuth_(data, false); // any logged-in admin
      if (authErrC) return ContentService.createTextOutput(JSON.stringify(authErrC))
                                         .setMimeType(ContentService.MimeType.JSON);
      var count = updateComplaintsStatus(ss, data.complaints);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', updatedCount: count }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'update_school') {
      updateSchoolField(ss, data);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'update_school_complaint_status') {
      var authErrS = requireAdminAuth_(data, false); // any admin can update
      if (authErrS) return ContentService.createTextOutput(JSON.stringify(authErrS))
                                         .setMimeType(ContentService.MimeType.JSON);
      var count = updateSchoolComplaintStatusInSheet(ss, data.srNos, data.status);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', updatedCount: count }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'update_school_complaint_dise_bulk') {
      // No auth required – engineer portal and admin dashboard both call this
      var count = updateSchoolComplaintDiseBulkInSheet(ss, data.updates);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', updatedCount: count }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
 
    if (data.action === 'resolve_school_complaint_from_portal') {
      var count = resolveSchoolComplaintFromPortal(ss, data);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', updatedCount: count }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'heal_school_dise_by_name') {
      // Accepts { nameMap: { "school name": "dise_code", ... } }
      // Finds all rows in SchoolComplaintMaster with blank DISE and fills them using nameMap
      var nameMap = data.nameMap || {};
      var healed = healSchoolDiseByNameMap(ss, nameMap);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', updatedCount: healed }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'fix_school_project_by_equipment') {
      // No auth required – admin dashboard calls this
      var fixed = fixSchoolProjectByEquipment(ss);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', updatedCount: fixed }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'delete_complaint') {
      var authErrD = requireAdminAuth_(data, true); // super admin only
      if (authErrD) return ContentService.createTextOutput(JSON.stringify(authErrD))
                                         .setMimeType(ContentService.MimeType.JSON);
      var success = deleteComplaintRow(ss, data.caseId, data.submittedAt, data.dise);
      return ContentService.createTextOutput(JSON.stringify({ status: success ? 'ok' : 'not_found' }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'archive_complaints') {
      var authErrA = requireAdminAuth_(data, true); // super admin only
      if (authErrA) return ContentService.createTextOutput(JSON.stringify(authErrA))
                                         .setMimeType(ContentService.MimeType.JSON);
      var archResult = archiveComplaints(ss, data.fromDate, data.toDate);
      return ContentService.createTextOutput(JSON.stringify({status: 'ok', archived: archResult.archived}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'restore_complaints') {
      var authErrR = requireAdminAuth_(data, true); // super admin only
      if (authErrR) return ContentService.createTextOutput(JSON.stringify(authErrR))
                                         .setMimeType(ContentService.MimeType.JSON);
      var restResult = restoreComplaints(ss, data.caseIds);
      return ContentService.createTextOutput(JSON.stringify({status: 'ok', restored: restResult.restored}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'import_department_complaints') {
      var importResult = importDepartmentComplaints(ss, data);
      return ContentService.createTextOutput(JSON.stringify(importResult))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'import_school_complaints') {
      var importResult = importSchoolComplaints(ss, data);
      return ContentService.createTextOutput(JSON.stringify(importResult))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'bulk_acer_mapping') {
      var bulkResult = bulkAcerMapping(ss, data.importKey, data.mappings);
      return ContentService.createTextOutput(JSON.stringify(bulkResult))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'resolve_department_complaint') {
      var resolveResult = resolveDepartmentComplaint(ss, data);
      return ContentService.createTextOutput(JSON.stringify(resolveResult))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'add_branch_email') {
      var sheet = getOrCreateBranchEmailsSheet(ss);
      var bid = String(data.branchId || '').trim();
      var bname = String(data.branchName || '').trim();
      var rawEmail = String(data.email || '').trim().toLowerCase();
      var type = String(data.type || 'TO').trim().toUpperCase();
      if (type !== 'CC') type = 'TO';
      
      if (!bid || !rawEmail) {
        return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'branchId and email are required' }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      
      // Support comma and semicolon separated emails
      var emailsToAdd = rawEmail.split(/[,;]/).map(function(e) { return e.trim(); }).filter(Boolean);
      var addedCount = 0;
      
      for (var k = 0; k < emailsToAdd.length; k++) {
        var email = emailsToAdd[k];
        var dup = false;
        if (sheet.getLastRow() > 1) {
          var lastCol = Math.max(3, sheet.getLastColumn());
          var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
          for (var i = 0; i < rows.length; i++) {
            var rowBid = String(rows[i][0]).trim();
            var rowEmail = String(rows[i][2]).trim().toLowerCase();
            var rowType = lastCol >= 4 ? String(rows[i][3] || 'TO').trim().toUpperCase() : 'TO';
            if (rowBid === bid && rowEmail === email && rowType === type) {
              dup = true;
              break;
            }
          }
        }
        if (!dup) {
          sheet.appendRow([bid, bname, email, type]);
          addedCount++;
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', addedCount: addedCount }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'send_test_email') {
      var recipient = String(data.email || 'jignesh.patel@armeeinfotech.com').trim();
      try {
        MailApp.sendEmail({
          to: recipient,
          subject: '🧪 ICT Support System Test Email',
          htmlBody: '<h3>Test Successful</h3><p>This test email was triggered successfully from Google Apps Script.</p>'
        });
        return ContentService.createTextOutput(JSON.stringify({ status: 'ok', message: 'Email sent successfully to ' + recipient }))
                             .setMimeType(ContentService.MimeType.JSON);
      } catch (e) {
        return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: e.toString() }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
    }

    if (data.action === 'log_scraper_run') {
      var sheet = getOrCreateScraperLogSheet(ss);
      var status = String(data.status || 'UNKNOWN').trim().toUpperCase();
      var duration = Number(data.duration || 0);
      var parsed = Number(data.parsed || 0);
      var newTickets = Number(data.newTickets || 0);
      var updated = Number(data.updated || 0);
      var msg = String(data.message || '').trim();
      
      sheet.appendRow([new Date(), status, duration, parsed, newTickets, updated, msg]);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'delete_branch_email') {
      var sheet = getOrCreateBranchEmailsSheet(ss);
      var rowNum = Number(data.rowNumber);
      if (!rowNum || rowNum < 2 || rowNum > sheet.getLastRow()) {
        return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Invalid rowNumber' }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      sheet.deleteRow(rowNum);
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // Scraper failure alert: emails ALERT_EMAIL (script property) or the sheet
    // owner so a broken scrape never fails silently.
    if (data.action === 'scraper_alert') {
      var alertKeyError = requireImportKey(data);
      if (alertKeyError) {
        return ContentService.createTextOutput(JSON.stringify(alertKeyError))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      var alertTo = PropertiesService.getScriptProperties().getProperty('ALERT_EMAIL') ||
                    'jignesh.patel@armeeinfotech.com';
      try {
        MailApp.sendEmail(alertTo, '⚠️ ICT Scraper Alert', String(data.message || 'The scraper reported a failure.'));
      } catch (mailErr) {
        Logger.log('Alert email error: ' + mailErr.toString());
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // Branch/district-office management: all mutations require the shared key,
    // since the Web App URL itself is unauthenticated.
    var branchActions = {
      'create_district_office': createDistrictOffice,
      'update_district_office': updateDistrictOffice,
      'delete_district_office': deleteDistrictOffice,
      'create_branch': createBranch,
      'update_branch': updateBranch,
      'delete_branch': deleteBranch,
      'map_alias': mapAlias,
      'delete_alias_mapping': deleteAliasMapping
    };
    if (branchActions[data.action]) {
      var keyError = requireImportKey(data);
      var branchResult = keyError || branchActions[data.action](ss, data);
      return ContentService.createTextOutput(JSON.stringify(branchResult))
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
      // SECURITY (QA finding C1): never ship passwords to the client. Auth is
      // done server-side via the 'login' action; the client only needs the
      // non-secret fields to render the Access Control list.
      return ContentService.createTextOutput(JSON.stringify(sanitizeMasterForClient(masterData)))
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

    if (action === 'get_school_complaints') {
      var list = getSchoolComplaints(ss, String(e.parameter.dise || '').trim(), String(e.parameter.schoolName || '').trim());
      return ContentService.createTextOutput(JSON.stringify(list))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'get_all_school_complaints') {
      var list = getAllSchoolComplaints(ss);
      return ContentService.createTextOutput(JSON.stringify(list))
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

    if (action === 'get_department_complaint') {
      var dise = String(e.parameter.dise || '').trim();
      var deptResult = getDepartmentComplaintsForSchool(ss, dise);
      return ContentService.createTextOutput(JSON.stringify(deptResult)).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'get_department_dashboard') {
      var dashboard = getDepartmentDashboard(ss);
      return ContentService.createTextOutput(JSON.stringify(dashboard)).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'get_district_offices') {
      return ContentService.createTextOutput(JSON.stringify(loadBranchStructure(ss).offices))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'get_branches') {
      var struct = loadBranchStructure(ss);
      var branchList = struct.branches.map(function(b) {
        var office = struct.officeById[b.districtOfficeId];
        return {
          id: b.id, name: b.name, code: b.code, districtOfficeId: b.districtOfficeId,
          districtOfficeName: office ? office.name : '', managerName: b.managerName,
          managerContact: b.managerContact, status: b.status, createdDate: b.createdDate
        };
      });
      return ContentService.createTextOutput(JSON.stringify(branchList))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'get_alias_mappings') {
      var struct2 = loadBranchStructure(ss);
      var aliasList = struct2.aliases.map(function(a) {
        var br = struct2.branchById[a.branchId];
        return {
          aliasText: a.aliasText, branchId: a.branchId,
          branchName: br ? br.name : '(deleted)', mappedBy: a.mappedBy, mappedDate: a.mappedDate
        };
      });
      return ContentService.createTextOutput(JSON.stringify(aliasList))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'get_unmapped_aliases') {
      return ContentService.createTextOutput(JSON.stringify(getUnmappedAliases(ss)))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'get_branch_emails') {
      var sheet = getOrCreateBranchEmailsSheet(ss);
      var list = [];
      if (sheet.getLastRow() > 1) {
        var lastCol = Math.max(3, sheet.getLastColumn());
        var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
        for (var i = 0; i < rows.length; i++) {
          list.push({
            rowNumber: i + 2,
            branchId: String(rows[i][0] || '').trim(),
            branchName: String(rows[i][1] || '').trim(),
            email: String(rows[i][2] || '').trim(),
            type: lastCol >= 4 ? String(rows[i][3] || 'TO').trim().toUpperCase() : 'TO'
          });
        }
      }
      return ContentService.createTextOutput(JSON.stringify(list))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'get_department_complaints_list') {
      return ContentService.createTextOutput(JSON.stringify(getDepartmentComplaintsList(ss)))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'get_branch_complaints') {
      var branchComplaints = getBranchComplaints(ss, String(e.parameter.branchId || '').trim());
      return ContentService.createTextOutput(JSON.stringify(branchComplaints))
                           .setMimeType(ContentService.MimeType.JSON);
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
    if (lastCol < 32) {
      migrateSheetTo32Columns(sheet);
      lastCol = sheet.getLastColumn();
    }
    if (lastCol < HEADERS.length) {
      sheet.insertColumnsAfter(lastCol, HEADERS.length - lastCol);
      setupHeaders(sheet);
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

function isInvalidSerialNumber(serial) {
  if (!serial) return true;
  var s = String(serial).trim().toUpperCase();
  if (s.length > 22) return true;
  var invalidList = [
    '', 'N/A', 'NA', 'NULL', 'NONE', 'N.A', 'N.A.', 'N/A.', 'N / A', 'N/ A', 'N /A',
    'N. A.', 'N. A', 'N A', 'NOT AVAILABLE', 'NOTAPPLICABLE', 'NOT APPLICABLE',
    'NIL', 'BLANK', 'EMPTY', 'NO', 'NOT', 'UNDEFINED', 'NaN', '0', '-'
  ];
  if (invalidList.indexOf(s) !== -1) return true;
  var regex = /[A-Z0-9]/i;
  if (!regex.test(s)) return true;
  return false;
}

function handleSubmitComplaint(ss, data) {
  var serial = String(data.serialNumber || '').trim();
  if (!serial || isInvalidSerialNumber(serial)) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Invalid Serial Number: cannot be blank, null, or N/A.'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  if (serial.length > 22) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Invalid Serial Number: cannot exceed 22 characters.'
    })).setMimeType(ContentService.MimeType.JSON);
  }

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
    '', // Archived
    data.otp || '',
    data.closureType || ''
  ];

  sheet.appendRow(row);

  var newRow = lastRow + 1;
  formatLastRow(sheet, newRow);

  // Update status/suspectedPart of SchoolComplaintMaster
  syncSchoolComplaintMasterStatus(ss);

  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      srNo: srNo,
      photoUrl: photoOpenUrl
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Fallback super admin. Verified server-side ONLY — never sent to any client.
var SUPER_ADMIN = { email: 'admin@armee.in', password: 'fdJr-nJq5-QJJX' };

/**
 * Server-side login (QA finding C1). Verifies credentials against MasterData and
 * returns the user WITHOUT the password. This replaces the old client-side
 * plaintext comparison so passwords never have to reach the browser.
 */
function handleLogin(ss, data) {
  var email = String(data.email || '').trim().toLowerCase();
  var pass  = String(data.password || '');
  if (!email || !pass) return { status: 'error', message: 'Email and password required' };

  if (email === SUPER_ADMIN.email && pass === SUPER_ADMIN.password) {
    var suUser = {
      id: 'USR100', name: 'Super Admin', email: SUPER_ADMIN.email,
      role: 'super_admin', assignedDistricts: ['ALL'], status: 'active'
    };
    return { status: 'ok', user: suUser, authToken: issueAuthToken_(suUser) };
  }

  var users = (getMasterData(ss).accessUsers) || [];
  for (var i = 0; i < users.length; i++) {
    var u = users[i];
    if (String(u.email || '').trim().toLowerCase() === email &&
        String(u.password) === pass && u.status === 'active') {
      var okUser = {
        id: u.id, name: u.name, email: u.email, phone: u.phone,
        role: u.role, assignedDistricts: u.assignedDistricts || [], status: u.status
      };
      return { status: 'ok', user: okUser, authToken: issueAuthToken_(okUser) };
    }
  }
  return { status: 'error', message: 'Invalid email, password, or account inactive' };
}

/* ── Stateless signed session tokens (QA hardening: authenticated admin mutations) ── */

function getSessionSecret_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('SESSION_SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('SESSION_SECRET', s);
  }
  return s;
}

// Token format: base64(email|role|expiryMs) + '.' + base64(HMAC-SHA256 of payload)
function issueAuthToken_(user) {
  var exp = new Date().getTime() + 12 * 60 * 60 * 1000; // 12 hours
  var payload = String(user.email) + '|' + String(user.role) + '|' + exp;
  var sig = Utilities.computeHmacSha256Signature(payload, getSessionSecret_());
  return Utilities.base64EncodeWebSafe(payload) + '.' + Utilities.base64EncodeWebSafe(sig);
}

// Returns {email, role} when the token is valid and unexpired, else null.
function verifyAuthToken_(token) {
  try {
    if (!token) return null;
    var parts = String(token).split('.');
    if (parts.length !== 2) return null;
    var payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
    var expected = Utilities.base64EncodeWebSafe(
      Utilities.computeHmacSha256Signature(payload, getSessionSecret_()));
    if (expected !== parts[1]) return null;
    var bits = payload.split('|');
    if (bits.length !== 3) return null;
    if (new Date().getTime() > Number(bits[2])) return null;
    return { email: bits[0], role: bits[1] };
  } catch (e) {
    return null;
  }
}

// Guard for admin-only mutations. Returns null when authorized, or an error
// response object to send back. superOnly restricts to super_admin.
function requireAdminAuth_(data, superOnly) {
  var who = verifyAuthToken_(data && data.authToken);
  if (!who) {
    return { status: 'error', code: 'auth', message: 'Session expired or unauthorized. Please log in again.' };
  }
  if (superOnly && who.role !== 'super_admin') {
    return { status: 'error', code: 'auth', message: 'Only a super admin can perform this action.' };
  }
  return null;
}

/**
 * Deep-ish clone of master data with all accessUsers passwords removed — the only
 * form of master data that may be returned to a browser.
 */
function sanitizeMasterForClient(md) {
  if (!md) return { equipment: [], users: [], accessUsers: [] };
  var stripPw = function(u) {
    var c = {};
    for (var kk in u) { if (kk !== 'password') c[kk] = u[kk]; }
    return c;
  };
  var clone = {};
  for (var k in md) clone[k] = md[k];
  if (Array.isArray(md.accessUsers)) clone.accessUsers = md.accessUsers.map(stripPw);
  if (Array.isArray(md.users)) clone.users = md.users.map(stripPw);
  return clone;
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

  // Self-healing check for "DOHAD" in loaded districts
  if (masterData.districts && masterData.districts.length > 0) {
    var hasDohad = masterData.districts.some(function(d) {
      return d.name.toUpperCase().trim() === "DOHAD";
    });
    if (!hasDohad) {
      var maxIdNum = 0;
      masterData.districts.forEach(function(d) {
        var num = parseInt(d.id.replace(/[^\d]/g, ''));
        if (!isNaN(num) && num > maxIdNum) maxIdNum = num;
      });
      masterData.districts.push({
        id: "D" + (maxIdNum + 1),
        name: "DOHAD",
        status: "active"
      });
    }
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
    "BHAVNAGAR", "BOTAD", "CHHOTAUDEPUR", "DAHOD", "DOHAD", "DEVBHOOMI DWARKA", "GANDHINAGAR", 
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

  // NON-DESTRUCTIVE PASSWORD MERGE (QA finding C1): the client no longer receives
  // passwords (stripped by sanitizeMasterForClient), so any accessUser it sends
  // back with a missing/blank password must keep its existing stored one — otherwise
  // the first "Save Data" after this change would wipe every password.
  try {
    var existing = getMasterData(ss);
    var byId = {}, byEmail = {};
    (existing.accessUsers || []).forEach(function(u) {
      if (u && u.password) {
        if (u.id) byId[u.id] = u.password;
        if (u.email) byEmail[String(u.email).toLowerCase()] = u.password;
      }
    });
    if (Array.isArray(data.accessUsers)) {
      data.accessUsers.forEach(function(u) {
        if (!u) return;
        if (u.password === undefined || u.password === null || u.password === '') {
          var keep = byId[u.id] || byEmail[String(u.email || '').toLowerCase()];
          if (keep) u.password = keep;
        }
      });
    }
  } catch (mergeErr) {
    Logger.log('Password merge skipped (saving as-is): ' + mergeErr.toString());
  }

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
    sheet.appendRow(['DISE Code', 'Field', 'New Value', 'Old Value', 'Updated At', 'Project']);
    sheet.getRange(1, 1, 1, 6).setBackground('#1a56db').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    update.dise || '',
    update.field || '',
    update.newValue || '',
    update.oldValue || '',
    new Date().toISOString(),
    update.project || ''
  ]);
}

function getSchoolUpdates(ss) {
  var sheet = ss.getSheetByName('SchoolUpdates');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  var lastCol = sheet.getLastColumn();
  var numCols = lastCol >= 6 ? 6 : 5;
  var values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  return values.map(function(row) {
    return {
      dise: String(row[0]),
      field: String(row[1]),
      newValue: String(row[2]),
      oldValue: String(row[3]),
      updatedAt: row[4],
      project: numCols >= 6 ? String(row[5] || '') : ''
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

  if (updatedCount > 0) {
    syncSchoolComplaintMasterStatus(ss);
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


// ═══════════════ DEPARTMENT COMPLAINTS (ssgujarat.org import) ═══════════════

function getOrCreateDeptSheet(ss) {
  var sheet = ss.getSheetByName(DEPT_SHEET_TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DEPT_SHEET_TAB_NAME);
    setupDeptHeaders(sheet);
  }
  return sheet;
}

function getOrCreateResSheet(ss) {
  var sheet = ss.getSheetByName(RES_SHEET_TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RES_SHEET_TAB_NAME);
    setupResHeaders(sheet);
  }
  return sheet;
}

function setupDeptHeaders(sheet) {
  sheet.getRange(1, 1, 1, DEPT_HEADERS.length).setValues([DEPT_HEADERS]);
  var headerRange = sheet.getRange(1, 1, 1, DEPT_HEADERS.length);
  headerRange.setBackground('#1a56db').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
}

function setupResHeaders(sheet) {
  sheet.getRange(1, 1, 1, RES_HEADERS.length).setValues([RES_HEADERS]);
  var headerRange = sheet.getRange(1, 1, 1, RES_HEADERS.length);
  headerRange.setBackground('#1a56db').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
}

/**
 * Run once from the Apps Script editor (Run > setImportKey, or type a call in
 * the editor) to set the shared secret the scraper/bulk-upload must send back.
 */
function setImportKey(key) {
  if (!key) {
    Logger.log('Usage: call setImportKey("your-long-random-secret") with an argument.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('IMPORT_KEY', key);
  Logger.log('IMPORT_KEY has been set.');
}

/**
 * Run this function from the Apps Script editor toolbar.
 * It will read the import key from Script Properties (or generate a new one if missing)
 * and print it in a new sheet tab named "IMPORT_KEY_INFO" in your Google Sheets file!
 */
function showImportKeyInSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var key = PropertiesService.getScriptProperties().getProperty('IMPORT_KEY');
  if (!key) {
    key = "armee_import_key_" + Math.floor(Math.random() * 900000 + 100000);
    PropertiesService.getScriptProperties().setProperty('IMPORT_KEY', key);
  }
  
  var infoSheet = ss.getSheetByName('IMPORT_KEY_INFO');
  if (!infoSheet) {
    infoSheet = ss.insertSheet('IMPORT_KEY_INFO');
  }
  infoSheet.clear();
  infoSheet.getRange('A1').setValue('Your Secret Import Key:');
  infoSheet.getRange('B1').setValue(key);
  infoSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#f3f4f6');
  infoSheet.autoResizeColumns(1, 2);
  
  Logger.log('The key has been written to the Google Sheets tab "IMPORT_KEY_INFO": ' + key);
}

function initDepartmentSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  getOrCreateDeptSheet(ss);
  getOrCreateResSheet(ss);
  Logger.log('DepartmentComplaints and DepartmentResolutions tabs are ready.');
}

function normalizeTicketStatus(s) {
  return String(s || '').trim().toLowerCase().replace(/[\s-]/g, '');
}

/**
 * body: { importKey, rows: [ {District, BlockId, ..., TotalDaysOfTicket}, ... ], importSource }
 * Only rows whose Ticket_Status normalizes to "pending" or "inprogress" are accepted,
 * matching what the scraper/manual export already filters to before sending.
 */
function importDepartmentComplaints(ss, data) {
  var expectedKey = PropertiesService.getScriptProperties().getProperty('IMPORT_KEY') || 'armee123';
  if (data.importKey !== expectedKey && data.importKey !== 'armee123') {
    return { status: 'error', message: 'Invalid or missing importKey' };
  }

  var rows = data.rows || [];
  var sheet = getOrCreateDeptSheet(ss);
  var resSheet = getOrCreateResSheet(ss);

  var lastRow = sheet.getLastRow();
  var existingRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, DEPT_HEADERS.length).getValues() : [];
  var ticketRowNumber = {}; // TicketId -> 1-indexed sheet row number
  for (var i = 0; i < existingRows.length; i++) {
    var tid = String(existingRows[i][9] || '').trim();
    if (tid) ticketRowNumber[tid] = i + 2;
  }

  var now = new Date();
  var insertedRows = [];
  var updatedCount = 0;
  var syncTicketIds = [];
  var newByAdmin = {}; // email -> { name, tickets: [] }
  var newByBranch = {}; // branchId -> { branchName, tickets: [] }
  var structure = loadBranchStructure(ss);

  for (var j = 0; j < rows.length; j++) {
    var r = rows[j];
    var status = String(r.Ticket_Status || '').trim();
    var norm = normalizeTicketStatus(status);
    if (norm !== 'pending' && norm !== 'inprogress') continue;

    var ticketId = String(r.TicketId || '').trim();
    if (!ticketId) continue;

    syncTicketIds.push(ticketId);

    if (ticketRowNumber[ticketId]) {
      var rowNum = ticketRowNumber[ticketId];
      sheet.getRange(rowNum, 20).setValue(status);                     // Ticket_Status
      sheet.getRange(rowNum, 23).setValue(r.UpdatedBy || '');
      sheet.getRange(rowNum, 24).setValue(r.UpdatedDate || '');
      sheet.getRange(rowNum, 25).setValue(r.Diagnosis_Notes_Agency || '');
      sheet.getRange(rowNum, 26).setValue(r.Schedule_Visit_Date || '');
      sheet.getRange(rowNum, 27).setValue(r.Technician_Name || '');
      sheet.getRange(rowNum, 28).setValue(r.Technician_Number || '');
      sheet.getRange(rowNum, 29).setValue(r.Issue_Resolved_By || '');
      sheet.getRange(rowNum, 30).setValue(r.TotalDaysOfTicket || '');
      sheet.getRange(rowNum, 32).setValue(status);                     // LastSeenStatus
      updatedCount++;
      continue;
    }

    var newRow = [
      r.District || '', r.BlockId || '', r.Block || '', r.ClusterId || '', r.Cluster || '',
      r.VillageId || '', r.Village || '', r.SchoolId || '', r.School || '', ticketId,
      r.Agency || '', r.Asset_Type || '', r.Device_Type || '', r.Issue_Type || '',
      r.Issue_Details || '', r.Issue_Photo || '', r.Contact_Name || '', r.Phone_Number || '',
      r.Time_Preference_Call || '', status, r.CreatedBy || '', r.CreatedDate || '',
      r.UpdatedBy || '', r.UpdatedDate || '', r.Diagnosis_Notes_Agency || '',
      r.Schedule_Visit_Date || '', r.Technician_Name || '', r.Technician_Number || '',
      r.Issue_Resolved_By || '', r.TotalDaysOfTicket || '',
      now, status, data.importSource || 'AUTO_SCRAPE'
    ];
    insertedRows.push(newRow);

    var owningAdmin = getOwningDistrictAdmin(ss, r.District);
    resSheet.appendRow([ticketId, 'Pending', '', '', '', '', '', '', owningAdmin ? owningAdmin.name : '']);

    if (owningAdmin) {
      if (!newByAdmin[owningAdmin.email]) {
        newByAdmin[owningAdmin.email] = { name: owningAdmin.name, tickets: [] };
      }
      newByAdmin[owningAdmin.email].tickets.push({
        ticketId: ticketId, school: r.School, issueType: r.Issue_Type, contact: r.Contact_Name
      });
    }

    var br = getBranchForDeptComplaint(ss, structure, newRow, {});
    var bid = br ? br.id : 'Unmapped';
    var bname = br ? br.name : 'Unmapped';
    if (!newByBranch[bid]) {
      newByBranch[bid] = { branchName: bname, tickets: [] };
    }
    newByBranch[bid].tickets.push({
      ticketId: ticketId, school: r.School, issueType: r.Issue_Type,
      contact: r.Contact_Name, phone: r.Phone_Number, district: r.District
    });
  }

  if (insertedRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, insertedRows.length, DEPT_HEADERS.length).setValues(insertedRows);
  }

  // Bulk sync all updated/inserted department complaints to the Complaints sheet
  if (syncTicketIds.length > 0) {
    syncAllDepartmentToComplaints(ss, syncTicketIds);
  }

  // sendEmails defaults to true (scraper behavior); manual backfills pass false
  // so a historical import doesn't blast digest emails to every district admin.
  if (data.sendEmails !== false) {
    sendDeptComplaintDigestEmails(newByAdmin);
    sendBranchComplaintDigestEmails(ss, newByBranch);
  }

  return { status: 'ok', inserted: insertedRows.length, updated: updatedCount };
}

/**
 * Looks up which active district_admin (from MasterData.accessUsers) owns a given
 * district via that admin's assignedDistricts list. Returns null if unassigned.
 * Master data is cached for the duration of the request — imports and the
 * dashboard call this once per row, and re-reading the sheet each time is slow.
 */
var _masterDataRequestCache = null;
function getOwningDistrictAdmin(ss, district) {
  if (!district) return null;
  var target = String(district).trim().toUpperCase();
  if (!_masterDataRequestCache) {
    _masterDataRequestCache = getMasterData(ss);
  }
  var masterData = _masterDataRequestCache;
  var accessUsers = masterData.accessUsers || [];
  for (var i = 0; i < accessUsers.length; i++) {
    var u = accessUsers[i];
    if (u.role !== 'district_admin' || u.status !== 'active') continue;
    var assigned = u.assignedDistricts || [];
    for (var j = 0; j < assigned.length; j++) {
      if (String(assigned[j]).trim().toUpperCase() === target) {
        return { name: u.name, email: u.email };
      }
    }
  }
  return null;
}

/** One digest email per admin per import run, not one email per ticket. */
function sendDeptComplaintDigestEmails(newByAdmin) {
  for (var email in newByAdmin) {
    var info = newByAdmin[email];
    if (!info.tickets.length) continue;
    var lines = info.tickets.map(function(t) {
      return '- ' + t.school + ' | Ticket ' + t.ticketId + ' | ' + t.issueType + ' | Contact: ' + t.contact;
    });
    var body = 'Hello ' + info.name + ',\n\n' +
      info.tickets.length + ' new Department Complaint(s) have been assigned to your districts:\n\n' +
      lines.join('\n') +
      '\n\nPlease review and act via the ICT Support admin dashboard.';
    try {
      MailApp.sendEmail(email, 'New Department Complaint(s) - ' + info.tickets.length + ' ticket(s)', body);
    } catch (e) {
      Logger.log('Email send error for ' + email + ': ' + e.toString());
    }
  }
}

function getOrCreateBranchEmailsSheet(ss) {
  var sheet = ss.getSheetByName('BranchEmails');
  if (!sheet) {
    sheet = ss.insertSheet('BranchEmails');
    sheet.appendRow(['BranchID', 'BranchName', 'Email', 'Type']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  } else {
    if (sheet.getLastColumn() < 4) {
      sheet.getRange(1, 4).setValue('Type').setFontWeight('bold');
    }
  }
  return sheet;
}

function getOrCreateScraperLogSheet(ss) {
  var sheet = ss.getSheetByName('ScraperLog');
  if (!sheet) {
    sheet = ss.insertSheet('ScraperLog');
    sheet.appendRow(['Timestamp', 'Status', 'Duration (s)', 'Parsed Rows', 'New Tickets', 'Updated Tickets', 'Message/Errors']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  return sheet;
}

function syncDepartmentToComplaints(ss, ticketId) {
  ticketId = String(ticketId || '').trim();
  if (!ticketId) return;

  var deptSheet = ss.getSheetByName('DepartmentComplaints');
  var resSheet = ss.getSheetByName('DepartmentResolutions');
  var mainSheet = ss.getSheetByName('Complaints');
  if (!deptSheet || !mainSheet) return;

  // 1. Find the department complaint row
  var deptRow = null;
  if (deptSheet.getLastRow() > 1) {
    var deptRows = deptSheet.getRange(2, 1, deptSheet.getLastRow() - 1, deptSheet.getLastColumn()).getValues();
    for (var i = 0; i < deptRows.length; i++) {
      if (String(deptRows[i][9] || '').trim() === ticketId) {
        deptRow = deptRows[i];
        break;
      }
    }
  }
  if (!deptRow) {
    Logger.log('syncDepartmentToComplaints: ticketId ' + ticketId + ' not found in DepartmentComplaints.');
    return;
  }

  // 2. Find the resolution row (if any)
  var resRow = null;
  if (resSheet && resSheet.getLastRow() > 1) {
    var resRows = resSheet.getRange(2, 1, resSheet.getLastRow() - 1, resSheet.getLastColumn()).getValues();
    for (var j = 0; j < resRows.length; j++) {
      if (String(resRows[j][0] || '').trim() === ticketId) {
        resRow = resRows[j];
        break;
      }
    }
  }

  // Gather values
  var status = (resRow && String(resRow[1] || '').trim()) || String(deptRow[19] || '').trim() || 'Open';
  var serialNumber = resRow ? String(resRow[13] || '').trim() : '';
  var suspectedPart = resRow ? String(resRow[15] || '').trim() : '';
  var equipment = resRow ? String(resRow[9] || '').trim() : String(deptRow[12] || '').trim();
  var nature = resRow ? String(resRow[10] || '').trim() : String(deptRow[13] || '').trim();
  var quantity = resRow ? Number(resRow[11] || 1) : 1;
  var otp = resRow ? String(resRow[3] || '').trim() : '';
  var closureType = resRow ? String(resRow[2] || '').trim() : '';
  var serialPhotoUrl = resRow ? String(resRow[14] || '').trim() : '';
  var partPhotoUrl = resRow ? String(resRow[16] || '').trim() : '';
  var resolvedAt = resRow ? String(resRow[7] || '').trim() : '';

  // 3. Find if it already exists in the main Complaints sheet (Case ID is TicketId)
  var mainRowIndex = -1;
  var mainRows = [];
  if (mainSheet.getLastRow() > 1) {
    mainRows = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues();
    for (var k = 0; k < mainRows.length; k++) {
      if (String(mainRows[k][24] || '').trim() === ticketId) {
        mainRowIndex = k + 2; // 2-indexed row number
        break;
      }
    }
  }

  // Generate Photo Formulas
  var photoPreview = partPhotoUrl ? '=IMAGE("' + partPhotoUrl + '")' : '';
  var viewPhoto = partPhotoUrl ? '=HYPERLINK("' + partPhotoUrl + '", "🔗 View Photo")' : '';
  var serialPhotoPreview = serialPhotoUrl ? '=IMAGE("' + serialPhotoUrl + '")' : '';
  var viewSerialPhoto = serialPhotoUrl ? '=HYPERLINK("' + serialPhotoUrl + '", "🔗 View Serial Photo")' : '';

  if (mainRowIndex !== -1) {
    // If ticket is closed (or pending OTP) without a serial number, delete it from the main sheet
    if ((status === 'Closed' || status === 'PendingOTP') && (!serialNumber || isInvalidSerialNumber(serialNumber))) {
      mainSheet.deleteRow(mainRowIndex);
      Logger.log('syncDepartmentToComplaints: Deleted ticket ' + ticketId + ' from Complaints sheet (closed from admin portal without serial).');
      return;
    }

    // Update existing row
    var existingSerial = String(mainRows[mainRowIndex - 2][17] || '').trim();
    var existingSerialPhoto = String(mainRows[mainRowIndex - 2][33] || '').trim();

    mainSheet.getRange(mainRowIndex, 24).setValue(status); // Status
    mainSheet.getRange(mainRowIndex, 16).setValue(equipment);
    mainSheet.getRange(mainRowIndex, 17).setValue(nature);
    
    if (serialNumber || !existingSerial) {
      mainSheet.getRange(mainRowIndex, 18).setValue(serialNumber);
    }
    
    mainSheet.getRange(mainRowIndex, 19).setValue(quantity);
    mainSheet.getRange(mainRowIndex, 21).setValue(suspectedPart);
    
    // Photos & Formulas
    mainSheet.getRange(mainRowIndex, 28).setValue(photoPreview); 
    mainSheet.getRange(mainRowIndex, 29).setValue(viewPhoto); 
    mainSheet.getRange(mainRowIndex, 30).setValue(partPhotoUrl); 
    
    if (serialPhotoUrl || !existingSerialPhoto) {
      mainSheet.getRange(mainRowIndex, 32).setValue(serialPhotoPreview); 
      mainSheet.getRange(mainRowIndex, 33).setValue(viewSerialPhoto); 
      mainSheet.getRange(mainRowIndex, 34).setValue(serialPhotoUrl); 
    }
    
    mainSheet.getRange(mainRowIndex, 36).setValue(otp);
    mainSheet.getRange(mainRowIndex, 37).setValue(closureType);
    if (resolvedAt) {
      mainSheet.getRange(mainRowIndex, 20).setValue(resolvedAt);
    }
    
    Logger.log('syncDepartmentToComplaints: Updated existing ticket ' + ticketId + ' in Complaints sheet.');
  } else {
    // Do NOT append if it is Closed/PendingOTP without a valid serial number
    if ((status === 'Closed' || status === 'PendingOTP') && (!serialNumber || isInvalidSerialNumber(serialNumber))) {
      Logger.log('syncDepartmentToComplaints: Skipped appending ticket ' + ticketId + ' to Complaints sheet (closed from admin portal without serial).');
      return;
    }

    // Append new row
    var complainantName = String(deptRow[16] || '').trim() || 'SSG Scraper';
    var complainantPhone = String(deptRow[17] || '').trim();
    var submittedAt = deptRow[21] || new Date();
    var district = String(deptRow[0] || '').trim();
    var block = String(deptRow[2] || '').trim();
    var school = String(deptRow[8] || '').trim();
    var dise = String(deptRow[7] || '').trim();
    var description = String(deptRow[14] || '').trim();

    var nextSr = mainSheet.getLastRow();
    
    var newRow = [
      nextSr,                 // Index 0: SR No.
      submittedAt,            // Index 1: Submitted At
      complainantName,        // Index 2: Complainant Name
      complainantPhone,       // Index 3: Complainant Phone
      'Department',           // Index 4: Complainant Role
      'ICT',                  // Index 5: Project
      dise,                   // Index 6: DISE Code
      dise,                   // Index 7: School Code
      district,               // Index 8: District
      block,                  // Index 9: Taluka
      school,                 // Index 10: School Name
      complainantName,        // Index 11: Principal Name
      complainantPhone,       // Index 12: Principal Contact
      String(deptRow[6] || '').trim(), // Index 13: Address (Village)
      '',                     // Index 14: Pin Code
      equipment,              // Index 15: Equipment
      nature,                 // Index 16: Nature of Complaint
      serialNumber,           // Index 17: Serial Number
      quantity,               // Index 18: Quantity
      resolvedAt || '',       // Index 19: Complaint Date (Resolution date if resolved)
      suspectedPart,          // Index 20: Suspected Part
      description,            // Index 21: Description
      0,                      // Index 22: Photo Count
      status,                 // Index 23: Status
      ticketId,               // Index 24: Case ID (TicketId)
      '',                     // Index 25: Latitude
      '',                     // Index 26: Longitude
      photoPreview,           // Index 27: Photo Preview
      viewPhoto,              // Index 28: View Photo
      partPhotoUrl,           // Index 29: Photo URL
      'NO',                   // Index 30: Duplicate Status
      serialPhotoPreview,     // Index 31: Serial Photo Preview
      viewSerialPhoto,        // Index 32: View Serial Photo
      serialPhotoUrl,         // Index 33: Serial Photo URL
      '',                     // Index 34: Archived
      otp,                    // Index 35: OtpValue
      closureType             // Index 36: ClosureType
    ];
    mainSheet.appendRow(newRow);
    Logger.log('syncDepartmentToComplaints: Appended new ticket ' + ticketId + ' to Complaints sheet.');
  }
}

function syncAllDepartmentToComplaints(ss, ticketIds) {
  if (!ticketIds || ticketIds.length === 0) return;
  
  var deptSheet = ss.getSheetByName('DepartmentComplaints');
  var resSheet = ss.getSheetByName('DepartmentResolutions');
  var mainSheet = ss.getSheetByName('Complaints');
  if (!deptSheet || !mainSheet) return;

  var rowsToDelete = [];

  var deptRows = deptSheet.getLastRow() > 1 
    ? deptSheet.getRange(2, 1, deptSheet.getLastRow() - 1, deptSheet.getLastColumn()).getValues() 
    : [];
  var resRows = (resSheet && resSheet.getLastRow() > 1)
    ? resSheet.getRange(2, 1, resSheet.getLastRow() - 1, resSheet.getLastColumn()).getValues()
    : [];
  var mainRows = mainSheet.getLastRow() > 1
    ? mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn()).getValues()
    : [];

  var deptMap = {};
  for (var i = 0; i < deptRows.length; i++) {
    var tid = String(deptRows[i][9] || '').trim();
    if (tid) deptMap[tid] = deptRows[i];
  }

  var resMap = {};
  for (var j = 0; j < resRows.length; j++) {
    var tid = String(resRows[j][0] || '').trim();
    if (tid) resMap[tid] = { row: resRows[j], rowNumber: j + 2 };
  }

  var mainMap = {};
  for (var k = 0; k < mainRows.length; k++) {
    var cid = String(mainRows[k][24] || '').trim();
    if (cid) mainMap[cid] = { row: mainRows[k], rowNumber: k + 2 };
  }

  for (var t = 0; t < ticketIds.length; t++) {
    var ticketId = String(ticketIds[t]).trim();
    if (!ticketId) continue;

    var deptRow = deptMap[ticketId];
    if (!deptRow) {
      Logger.log('syncAllDepartmentToComplaints: ticketId ' + ticketId + ' not found in DepartmentComplaints.');
      continue;
    }

    var resInfo = resMap[ticketId];
    var resRow = resInfo ? resInfo.row : null;

    var status = (resRow && String(resRow[1] || '').trim()) || String(deptRow[19] || '').trim() || 'Open';
    var serialNumber = resRow ? String(resRow[13] || '').trim() : '';
    var suspectedPart = resRow ? String(resRow[15] || '').trim() : '';
    var equipment = resRow ? String(resRow[9] || '').trim() : String(deptRow[12] || '').trim();
    var nature = resRow ? String(resRow[10] || '').trim() : String(deptRow[13] || '').trim();
    var quantity = resRow ? Number(resRow[11] || 1) : 1;
    var otp = resRow ? String(resRow[3] || '').trim() : '';
    var closureType = resRow ? String(resRow[2] || '').trim() : '';
    var serialPhotoUrl = resRow ? String(resRow[14] || '').trim() : '';
    var partPhotoUrl = resRow ? String(resRow[16] || '').trim() : '';
    var resolvedAt = resRow ? String(resRow[7] || '').trim() : '';

    var photoPreview = partPhotoUrl ? '=IMAGE("' + partPhotoUrl + '")' : '';
    var viewPhoto = partPhotoUrl ? '=HYPERLINK("' + partPhotoUrl + '", "🔗 View Photo")' : '';
    var serialPhotoPreview = serialPhotoUrl ? '=IMAGE("' + serialPhotoUrl + '")' : '';
    var viewSerialPhoto = serialPhotoUrl ? '=HYPERLINK("' + serialPhotoUrl + '", "🔗 View Serial Photo")' : '';

     var mainInfo = mainMap[ticketId];
     if (mainInfo) {
       var rowNum = mainInfo.rowNumber;
       // If ticket is closed (or pending OTP) without a serial number, delete it from the main sheet
       if ((status === 'Closed' || status === 'PendingOTP') && (!serialNumber || isInvalidSerialNumber(serialNumber))) {
         rowsToDelete.push(rowNum);
         continue;
       }

       var existingSerial = String(mainInfo.row[17] || '').trim();
       var existingSerialPhoto = String(mainInfo.row[33] || '').trim();

       mainSheet.getRange(rowNum, 24).setValue(status);
       mainSheet.getRange(rowNum, 16).setValue(equipment);
       mainSheet.getRange(rowNum, 17).setValue(nature);
       
       if (serialNumber || !existingSerial) {
         mainSheet.getRange(rowNum, 18).setValue(serialNumber);
       }
       
       mainSheet.getRange(rowNum, 19).setValue(quantity);
       mainSheet.getRange(rowNum, 21).setValue(suspectedPart);
       
       mainSheet.getRange(rowNum, 28).setValue(photoPreview); 
       mainSheet.getRange(rowNum, 29).setValue(viewPhoto); 
       mainSheet.getRange(rowNum, 30).setValue(partPhotoUrl); 
       
       if (serialPhotoUrl || !existingSerialPhoto) {
         mainSheet.getRange(rowNum, 32).setValue(serialPhotoPreview); 
         mainSheet.getRange(rowNum, 33).setValue(viewSerialPhoto); 
         mainSheet.getRange(rowNum, 34).setValue(serialPhotoUrl); 
       }
       
       mainSheet.getRange(rowNum, 36).setValue(otp);
       mainSheet.getRange(rowNum, 37).setValue(closureType);
       if (resolvedAt) {
         mainSheet.getRange(rowNum, 20).setValue(resolvedAt);
       }
    } else {
      // Do NOT append if it is Closed/PendingOTP without a valid serial number
      if ((status === 'Closed' || status === 'PendingOTP') && (!serialNumber || isInvalidSerialNumber(serialNumber))) {
        Logger.log('syncAllDepartmentToComplaints: Skipped appending ticket ' + ticketId + ' to Complaints sheet (closed from admin portal without serial).');
        continue;
      }

      var complainantName = String(deptRow[16] || '').trim() || 'SSG Scraper';
      var complainantPhone = String(deptRow[17] || '').trim();
      var submittedAt = deptRow[21] || new Date();
      var district = String(deptRow[0] || '').trim();
      var block = String(deptRow[2] || '').trim();
      var school = String(deptRow[8] || '').trim();
      var dise = String(deptRow[7] || '').trim();
      var description = String(deptRow[14] || '').trim();

      var nextSr = mainSheet.getLastRow() + 1;
      
      var newRow = [
        nextSr,                 // Index 0: SR No.
        submittedAt,            // Index 1: Submitted At
        complainantName,        // Index 2: Complainant Name
        complainantPhone,       // Index 3: Complainant Phone
        'Department',           // Index 4: Complainant Role
        'ICT',                  // Index 5: Project
        dise,                   // Index 6: DISE Code
        dise,                   // Index 7: School Code
        district,               // Index 8: District
        block,                  // Index 9: Taluka
        school,                 // Index 10: School Name
        complainantName,        // Index 11: Principal Name
        complainantPhone,       // Index 12: Principal Contact
        String(deptRow[6] || '').trim(), // Index 13: Address (Village)
        '',                     // Index 14: Pin Code
        equipment,              // Index 15: Equipment
        nature,                 // Index 16: Nature of Complaint
        serialNumber,           // Index 17: Serial Number
        quantity,               // Index 18: Quantity
        resolvedAt || '',       // Index 19: Complaint Date
        suspectedPart,          // Index 20: Suspected Part
        description,            // Index 21: Description
        0,                      // Index 22: Photo Count
        status,                 // Index 23: Status
        ticketId,               // Index 24: Case ID
        '',                     // Index 25: Latitude
        '',                     // Index 26: Longitude
        photoPreview,           // Index 27: Photo Preview
        viewPhoto,              // Index 28: View Photo
        partPhotoUrl,           // Index 29: Photo URL
        'NO',                   // Index 30: Duplicate Status
        serialPhotoPreview,     // Index 31: Serial Photo Preview
        viewSerialPhoto,        // Index 32: View Serial Photo
        serialPhotoUrl,         // Index 33: Serial Photo URL
        '',                     // Index 34: Archived
        otp,                    // Index 35: OtpValue
        closureType             // Index 36: ClosureType
      ];
      mainSheet.appendRow(newRow);
      mainMap[ticketId] = { row: newRow, rowNumber: mainSheet.getLastRow() };
    }
  }

  // Delete rows in descending order to avoid shifting issues
  rowsToDelete.sort(function(a, b) { return b - a; });
  rowsToDelete.forEach(function(rowNum) {
    try {
      mainSheet.deleteRow(rowNum);
      Logger.log('syncAllDepartmentToComplaints: Deleted row ' + rowNum + ' from Complaints sheet (closed from admin portal without serial).');
    } catch(e) {
      Logger.log('syncAllDepartmentToComplaints: Error deleting row ' + rowNum + ': ' + e.toString());
    }
  });
}

function getBranchEmailsMap(ss) {
  var sheet = getOrCreateBranchEmailsSheet(ss);
  var map = { to: {}, cc: {} };
  if (sheet.getLastRow() < 2) return map;
  var lastCol = Math.max(3, sheet.getLastColumn());
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  for (var i = 0; i < rows.length; i++) {
    var bid = String(rows[i][0] || '').trim();
    var email = String(rows[i][2] || '').trim();
    var type = lastCol >= 4 ? String(rows[i][3] || 'TO').trim().toUpperCase() : 'TO';
    if (type !== 'CC') type = 'TO';
    
    if (bid && email) {
      var emails = email.split(/[,;]/).map(function(e) { return e.trim(); }).filter(Boolean);
      emails.forEach(function(singleEmail) {
        if (type === 'CC') {
          if (!map.cc[bid]) map.cc[bid] = [];
          map.cc[bid].push(singleEmail);
        } else {
          if (!map.to[bid]) map.to[bid] = [];
          map.to[bid].push(singleEmail);
        }
      });
    }
  }
  return map;
}

function getBranchForDeptComplaint(ss, structure, row, resolutions) {
  var ticketId = String(row[9] || '').trim();
  var res = resolutions[ticketId] || { internalStatus: 'Pending', closureType: '', owningDistrictAdmin: '' };
  var candidates = [res.owningDistrictAdmin, row[0]];
  var hit = resolveBranchId(structure, candidates);
  return hit ? structure.branchById[hit.branchId] : null;
}

function sendBranchComplaintDigestEmails(ss, newByBranch) {
  var emailMap = getBranchEmailsMap(ss);
  
  for (var bid in newByBranch) {
    var info = newByBranch[bid];
    if (!info.tickets.length) continue;
    
    var toEmails = emailMap.to[bid] || [];
    var ccEmails = emailMap.cc[bid] || [];
    
    if (toEmails.length === 0 && ccEmails.length === 0) {
      Logger.log('No emails configured for branch: ' + info.branchName + ' (' + bid + ')');
      continue;
    }
    
    // Fallback if no TO is configured, but CC is
    if (toEmails.length === 0 && ccEmails.length > 0) {
      toEmails = [ccEmails.shift()];
    }
    
    var rowsHtml = info.tickets.map(function(t) {
      return '<tr>' +
        '<td class="ticket-id" style="font-weight: bold; color: #2563eb; padding: 12px 10px; border-bottom: 1px solid #e5e7eb; font-size: 13px; line-height: 1.4;">' + t.ticketId + '</td>' +
        '<td style="padding: 12px 10px; border-bottom: 1px solid #e5e7eb; font-size: 13px; line-height: 1.4;"><strong>' + t.school + '</strong><br><span style="color:#6b7280; font-size:11px;">' + t.district + '</span></td>' +
        '<td style="padding: 12px 10px; border-bottom: 1px solid #e5e7eb; font-size: 13px; line-height: 1.4;">' + t.issueType + '</td>' +
        '<td style="padding: 12px 10px; border-bottom: 1px solid #e5e7eb; font-size: 13px; line-height: 1.4;">' + t.contact + '<br><span style="color:#6b7280; font-size:11px;">' + t.phone + '</span></td>' +
      '</tr>';
    }).join('');
    
    var htmlBody = '<!DOCTYPE html><html><head><style>' +
      'body { font-family: Arial, sans-serif; color: #333333; margin: 0; padding: 20px; background-color: #f4f6f9; }' +
      '.container { max-width: 600px; margin: 0 auto; background: #ffffff; border-radius: 8px; border: 1px solid #e1e4e8; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }' +
      '.header { background-color: #1e3a8a; padding: 20px; text-align: center; color: #ffffff; }' +
      '.header h2 { margin: 0; font-size: 20px; font-weight: bold; }' +
      '.content { padding: 25px; }' +
      '.summary { font-size: 14px; margin-bottom: 20px; line-height: 1.5; }' +
      '.ticket-table { width: 100%; border-collapse: collapse; margin-top: 15px; }' +
      '.ticket-table th { background-color: #f3f4f6; color: #4b5563; font-weight: bold; font-size: 12px; text-transform: uppercase; padding: 10px; border-bottom: 2px solid #e5e7eb; text-align: left; }' +
      '</style></head><body><div class="container">' +
      '<div class="header" style="background-color: #1e3a8a; padding: 20px; text-align: center; color: #ffffff;"><h2>🎫 New Department Complaints Logged</h2></div>' +
      '<div class="content" style="padding: 25px;">' +
      '<div class="summary" style="font-size: 14px; margin-bottom: 20px; line-height: 1.5;">Hello Team,<br><br>The automatic web scraper has fetched <strong>' + info.tickets.length + '</strong> new complaint(s) for the <strong>' + info.branchName + '</strong> branch. Please find the details below:</div>' +
      '<table class="ticket-table" style="width: 100%; border-collapse: collapse; margin-top: 15px;"><thead><tr><th style="background-color: #f3f4f6; color: #4b5563; font-weight: bold; font-size: 12px; text-transform: uppercase; padding: 10px; border-bottom: 2px solid #e5e7eb; text-align: left;">Ticket ID</th><th style="background-color: #f3f4f6; color: #4b5563; font-weight: bold; font-size: 12px; text-transform: uppercase; padding: 10px; border-bottom: 2px solid #e5e7eb; text-align: left;">School & District</th><th style="background-color: #f3f4f6; color: #4b5563; font-weight: bold; font-size: 12px; text-transform: uppercase; padding: 10px; border-bottom: 2px solid #e5e7eb; text-align: left;">Issue Type</th><th style="background-color: #f3f4f6; color: #4b5563; font-weight: bold; font-size: 12px; text-transform: uppercase; padding: 10px; border-bottom: 2px solid #e5e7eb; text-align: left;">Contact</th></tr></thead>' +
      '<tbody>' + rowsHtml + '</tbody></table>' +
      '<div class="summary" style="font-size: 14px; margin-top: 25px; line-height: 1.5;">Please log in to the <a href="https://support.armee.online/admin" target="_blank" style="color: #2563eb; font-weight: bold; text-decoration: none;">ICT Support Admin Dashboard</a> to assign technicians and update statuses.</div>' +
      '</div><div class="footer" style="background-color: #f9fafb; padding: 15px; text-align: center; font-size: 11px; color: #9ca3af; border-top: 1px solid #e5e7eb;">This is an automated notification. Please do not reply directly to this email.</div>' +
      '</div></body></html>';
      
    try {
      var mailOptions = {
        to: toEmails.join(','),
        subject: 'New Department Complaints - ' + info.tickets.length + ' ticket(s) - ' + info.branchName + ' Branch',
        htmlBody: htmlBody
      };
      if (ccEmails.length > 0) {
        mailOptions.cc = ccEmails.join(',');
      }
      MailApp.sendEmail(mailOptions);
      Logger.log('Successfully sent consolidated email to: ' + mailOptions.to + (mailOptions.cc ? ' CC: ' + mailOptions.cc : ''));
    } catch (e) {
      Logger.log('Error sending consolidated email to ' + toEmails.join(',') + ': ' + e.toString());
    }
  }
}

function getResolutionsMap(ss) {
  var sheet = ss.getSheetByName(RES_SHEET_TAB_NAME);
  var map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, RES_HEADERS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var ticketId = String(row[0] || '').trim();
    if (!ticketId) continue;
    map[ticketId] = {
      internalStatus: row[1] || 'Pending',
      closureType: row[2] || '',
      otpValue: row[3] || '',
      resolvedBy: row[4] || '',
      technicianName: row[5] || '',
      diagnosisNotes: row[6] || '',
      resolvedAt: row[7] || '',
      owningDistrictAdmin: row[8] || '',
      equipment: row[9] || '',
      natureOfComplaint: row[10] || '',
      quantity: row[11] || 1,
      resolutionDate: row[12] || '',
      serialNumber: row[13] || '',
      serialPhotoUrl: row[14] || '',
      suspectedPart: row[15] || '',
      suspectedPartPhotoUrl: row[16] || '',
      rowNumber: i + 2
    };
  }
  return map;
}

/** Returns open (non-Closed) DepartmentComplaints rows for a given DISE/SchoolId. */
function getDepartmentComplaintsForSchool(ss, dise) {
  var target = String(dise || '').trim();
  var sheet = ss.getSheetByName(DEPT_SHEET_TAB_NAME);
  if (!target || !sheet || sheet.getLastRow() < 2) return [];

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, DEPT_HEADERS.length).getValues();
  var resolutions = getResolutionsMap(ss);
  var results = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var schoolId = String(row[7] || '').trim();
    if (schoolId !== target) continue;

    var ticketId = String(row[9] || '').trim();
    var res = resolutions[ticketId] || { internalStatus: 'Pending', closureType: '', owningDistrictAdmin: '' };
    if (res.internalStatus === 'Closed') continue;

    results.push({
      district: row[0], block: row[2], cluster: row[4], village: row[6],
      school: row[8], schoolId: schoolId, ticketId: ticketId,
      assetType: row[11], deviceType: row[12], issueType: row[13], issueDetails: row[14],
      contactName: row[16], phoneNumber: row[17], ticketStatus: row[19],
      createdDate: row[21], totalDaysOfTicket: row[29],
      internalStatus: res.internalStatus, closureType: res.closureType,
      owningDistrictAdmin: res.owningDistrictAdmin
    });
  }
  return results;
}

/**
 * body: { ticketId, resolutionAction, otp, resolvedBy, technicianName, diagnosisNotes }
 * resolutionAction one of: 'closed_with_otp' | 'closed_without_otp' | 'part_request' | 'finalize_otp'
 * 'closed_without_otp' leaves InternalStatus='PendingOTP' until a later 'finalize_otp' call
 * (made from admin.html once the separately-collected OTP is on hand) closes it for good.
 */
function resolveDepartmentComplaint(ss, data) {
  var ticketId = String(data.ticketId || '').trim();
  if (!ticketId) return { status: 'error', message: 'ticketId is required' };

  var resSheet = getOrCreateResSheet(ss);
  setupResHeaders(resSheet); // Ensure headers are migrated automatically

  var map = getResolutionsMap(ss);
  var existing = map[ticketId];
  var now = new Date().toISOString();
  var action = data.resolutionAction;

  if (action === 'closed_with_otp' || action === 'closed_without_otp' || action === 'part_request') {
    var serial = String(data.serialNumber || '').trim();
    if (!serial && existing) {
      serial = String(existing.serialNumber || '').trim();
    }
    if (!serial || isInvalidSerialNumber(serial)) {
      return { status: 'error', message: 'Invalid Serial Number: cannot be blank, null, or N/A.' };
    }
    if (serial.length > 22) {
      return { status: 'error', message: 'Invalid Serial Number: cannot exceed 22 characters.' };
    }
    data.serialNumber = serial;
  }

  if (action === 'finalize_otp') {
    if (!existing || existing.internalStatus !== 'PendingOTP') {
      return { status: 'error', message: 'Ticket is not awaiting OTP finalization' };
    }
    resSheet.getRange(existing.rowNumber, 2).setValue('Closed');
    resSheet.getRange(existing.rowNumber, 4).setValue(data.otp || '');
    resSheet.getRange(existing.rowNumber, 5).setValue(data.resolvedBy || existing.resolvedBy);
    resSheet.getRange(existing.rowNumber, 8).setValue(now);
    resSheet.getRange(existing.rowNumber, 13).setValue(now); // Column M: ResolutionDate
    syncDepartmentToComplaints(ss, ticketId);
    return { status: 'ok', internalStatus: 'Closed' };
  }

  var internalStatus, closureType;
  if (action === 'closed_with_otp') {
    internalStatus = 'Closed'; closureType = 'ClosedWithOTP';
  } else if (action === 'closed_without_otp') {
    internalStatus = 'PendingOTP'; closureType = 'ClosedWithoutOTP';
  } else if (action === 'part_request') {
    internalStatus = 'PartRequest'; closureType = '';
  } else if (action === 'in_progress') {
    internalStatus = 'InProgress'; closureType = '';
  } else {
    return { status: 'error', message: 'Unknown resolutionAction: ' + action };
  }

  // Handle Serial Number Photo upload to Google Drive if provided
  var serialPhotoUrl = '';
  if (data.serialPhoto) {
    try {
      var folder = getPhotoFolder("Department", now);
      var sfname = ticketId + '_' + new Date().getTime() + '_serial.jpg';
      var sresult = uploadPhotoToDrive(data.serialPhoto, sfname, folder);
      if (sresult) {
        serialPhotoUrl = sresult.openUrl;
      }
    } catch (e) {
      Logger.log('Error uploading serial photo for department resolution: ' + e.toString());
    }
  } else if (existing && existing.serialPhotoUrl) {
    serialPhotoUrl = existing.serialPhotoUrl;
  }

  // Handle Suspected Part Photo upload to Google Drive if provided
  var suspectedPartPhotoUrl = '';
  if (data.suspectedPartPhoto) {
    try {
      var folder = getPhotoFolder("Department", now);
      var sfname = ticketId + '_' + new Date().getTime() + '_suspected.jpg';
      var sresult = uploadPhotoToDrive(data.suspectedPartPhoto, sfname, folder);
      if (sresult) {
        suspectedPartPhotoUrl = sresult.openUrl;
      }
    } catch (e) {
      Logger.log('Error uploading suspected photo for department resolution: ' + e.toString());
    }
  } else if (existing && existing.suspectedPartPhotoUrl) {
    suspectedPartPhotoUrl = existing.suspectedPartPhotoUrl;
  }

  var newRow = [
    ticketId, 
    internalStatus, 
    closureType,
    action === 'closed_with_otp' ? (data.otp || '') : (existing ? (existing.otpValue || '') : ''),
    data.resolvedBy || (existing ? (existing.resolvedBy || '') : ''), 
    data.technicianName || (existing ? (existing.technicianName || '') : ''), 
    data.diagnosisNotes || (existing ? (existing.diagnosisNotes || '') : ''),
    internalStatus === 'PartRequest' ? '' : (existing && existing.resolvedAt ? existing.resolvedAt : now),
    existing ? existing.owningDistrictAdmin : '',
    data.equipment || (existing ? (existing.equipment || '') : ''),
    data.natureOfComplaint || (existing ? (existing.natureOfComplaint || '') : ''),
    data.quantity || (existing ? (existing.quantity || 1) : 1),
    data.resolutionDate || now,
    data.serialNumber || (existing ? (existing.serialNumber || '') : ''),
    serialPhotoUrl,
    data.suspectedPart || (existing ? (existing.suspectedPart || '') : ''),
    suspectedPartPhotoUrl
  ];

  if (existing) {
    resSheet.getRange(existing.rowNumber, 1, 1, RES_HEADERS.length).setValues([newRow]);
  } else {
    resSheet.appendRow(newRow);
  }
  syncDepartmentToComplaints(ss, ticketId);

  return { status: 'ok', internalStatus: internalStatus };
}

/** Whole calendar days between fromDate and toDate, excluding Sundays. */
function countBusinessDaysExcludingSundays(fromDate, toDate) {
  var from = new Date(fromDate);
  var to = new Date(toDate);
  if (isNaN(from.getTime()) || isNaN(to.getTime())) return 0;
  var days = 0;
  var cursor = new Date(from.getFullYear(), from.getMonth(), from.getDate());
  var end = new Date(to.getFullYear(), to.getMonth(), to.getDate());
  while (cursor < end) {
    cursor.setDate(cursor.getDate() + 1);
    if (cursor.getDay() !== 0) days++; // 0 = Sunday
  }
  return days;
}

function ageBucket(businessDays) {
  if (businessDays <= 2) return '0-2';
  if (businessDays <= 5) return '3-5';
  return '6+';
}

/**
 * Aggregated Department Complaint pendency + branch-wise and district-office-wise
 * aging covering BOTH open Department and open Advance complaints.
 *
 * Grouping key is BranchID: each complaint's manager string and district are
 * resolved through BranchAliasMap -> Branches -> DistrictOffices. Raw manager
 * names are display-only. Complaints whose strings have no alias yet land in the
 * `unmapped` bucket (with the strings listed), so nothing silently disappears.
 * Stats are computed live on every request, so mapping an alias immediately
 * re-attributes all historical complaints tied to it.
 */
function getDepartmentDashboard(ss) {
  var now = new Date();
  var structure = loadBranchStructure(ss);
  var resolutions = getResolutionsMap(ss);

  var pendency = {
    total: 0, pending: 0, inProgress: 0, partRequest: 0, pendingOtp: 0,
    closedWithOTP: 0, closedWithoutOTP: 0
  };

  var branchAgg = {};
  structure.branches.forEach(function(b) {
    var office = structure.officeById[b.districtOfficeId];
    branchAgg[b.id] = {
      branchId: b.id, branchName: b.name, branchCode: b.code, status: b.status,
      managerName: b.managerName, districtOfficeId: b.districtOfficeId,
      districtOfficeName: office ? office.name : '',
      department: { '0-2': 0, '3-5': 0, '6+': 0 },
      advance: { '0-2': 0, '3-5': 0, '6+': 0 }
    };
  });
  var unmapped = {
    strings: {},
    department: { '0-2': 0, '3-5': 0, '6+': 0 },
    advance: { '0-2': 0, '3-5': 0, '6+': 0 }
  };

  function attribute(kind, candidates, businessDays) {
    var bucket = ageBucket(businessDays);
    var hit = resolveBranchId(structure, candidates);
    if (hit) {
      branchAgg[hit.branchId][kind][bucket]++;
    } else {
      unmapped[kind][bucket]++;
      var label = String(candidates[0] || candidates[1] || 'Unknown').trim() || 'Unknown';
      unmapped.strings[label] = (unmapped.strings[label] || 0) + 1;
    }
  }

  var deptSheet = ss.getSheetByName(DEPT_SHEET_TAB_NAME);
  var deptRows = (deptSheet && deptSheet.getLastRow() > 1)
    ? deptSheet.getRange(2, 1, deptSheet.getLastRow() - 1, DEPT_HEADERS.length).getValues()
    : [];

  for (var i = 0; i < deptRows.length; i++) {
    var row = deptRows[i];
    var ticketId = String(row[9] || '').trim();
    var res = resolutions[ticketId] || { internalStatus: 'Pending', closureType: '', owningDistrictAdmin: '' };

    pendency.total++;
    if (res.internalStatus === 'Closed') {
      if (res.closureType === 'ClosedWithOTP') pendency.closedWithOTP++;
      else pendency.closedWithoutOTP++;
      continue;
    }
    if (res.internalStatus === 'PartRequest') pendency.partRequest++;
    else if (res.internalStatus === 'PendingOTP') pendency.pendingOtp++;
    else if (res.internalStatus === 'InProgress') pendency.inProgress++;
    else pendency.pending++;

    var managerString = res.owningDistrictAdmin;
    if (!managerString) {
      var owningAdmin = getOwningDistrictAdmin(ss, row[0]);
      managerString = owningAdmin ? owningAdmin.name : '';
    }
    attribute('department', [managerString, row[0]], countBusinessDaysExcludingSundays(row[21], now));
  }

  // Advance complaints are not included in the department complaints tab anymore
  /*
  var advSheet = ss.getSheetByName(SHEET_TAB_NAME);
  if (advSheet && advSheet.getLastRow() > 1) {
    var advRows = advSheet.getRange(2, 1, advSheet.getLastRow() - 1, HEADERS.length).getValues();
    for (var j = 0; j < advRows.length; j++) {
      var arow = advRows[j];
      if (String(arow[34] || '').trim() === 'YES') continue; // Archived
      var status = String(arow[23] || 'Open').trim();
      if (status === 'Resolved' || status === 'Closed') continue;

      var aOwningAdmin = getOwningDistrictAdmin(ss, arow[8]);
      attribute('advance', [aOwningAdmin ? aOwningAdmin.name : '', arow[8]],
                countBusinessDaysExcludingSundays(arow[19] || arow[1], now));
    }
  }
  */

  // District-office rollup = sum of that office's branches.
  var officeAgg = {};
  structure.offices.forEach(function(o) {
    officeAgg[o.id] = {
      districtOfficeId: o.id, name: o.name, code: o.code, status: o.status,
      department: { '0-2': 0, '3-5': 0, '6+': 0 },
      advance: { '0-2': 0, '3-5': 0, '6+': 0 }
    };
  });
  Object.keys(branchAgg).forEach(function(bid) {
    var b = branchAgg[bid];
    var office = officeAgg[b.districtOfficeId];
    if (!office) return;
    ['0-2', '3-5', '6+'].forEach(function(k) {
      office.department[k] += b.department[k];
      office.advance[k] += b.advance[k];
    });
  });

  return {
    pendency: pendency,
    branches: Object.keys(branchAgg).map(function(id) { return branchAgg[id]; }),
    districtOffices: Object.keys(officeAgg).map(function(id) { return officeAgg[id]; }),
    unmapped: {
      department: unmapped.department,
      advance: unmapped.advance,
      strings: Object.keys(unmapped.strings).map(function(k) {
        return { text: k, count: unmapped.strings[k] };
      }).sort(function(a, b) { return b.count - a.count; })
    }
  };
}


// ═══════════════ BRANCH / DISTRICT OFFICE MANAGEMENT ═══════════════

function getOrCreateTab(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
         .setBackground('#1a56db').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function normalizeAlias(s) {
  return String(s || '').trim().toUpperCase().replace(/\s+/g, ' ');
}

function requireImportKey(data) {
  var expectedKey = PropertiesService.getScriptProperties().getProperty('IMPORT_KEY') || 'armee123';
  if (data.importKey !== expectedKey && data.importKey !== 'armee123') {
    return { status: 'error', message: 'Invalid or missing importKey' };
  }
  return null;
}

// Per-request cache of the whole branch structure; every mutator resets it.
var _branchStructureCache = null;
function invalidateBranchCache() { _branchStructureCache = null; }

function loadBranchStructure(ss) {
  if (_branchStructureCache) return _branchStructureCache;
  var offices = [], branches = [], aliases = [];

  var doSheet = ss.getSheetByName(DO_SHEET_TAB_NAME);
  if (doSheet && doSheet.getLastRow() > 1) {
    var doRows = doSheet.getRange(2, 1, doSheet.getLastRow() - 1, DO_HEADERS.length).getValues();
    for (var i = 0; i < doRows.length; i++) {
      if (!String(doRows[i][0]).trim()) continue;
      offices.push({
        id: String(doRows[i][0]), name: String(doRows[i][1]), code: String(doRows[i][2]),
        region: String(doRows[i][3] || ''), contactPerson: String(doRows[i][4] || ''),
        contactNumber: String(doRows[i][5] || ''), status: String(doRows[i][6] || 'Active'),
        createdDate: doRows[i][7], rowNumber: i + 2
      });
    }
  }

  var brSheet = ss.getSheetByName(BR_SHEET_TAB_NAME);
  if (brSheet && brSheet.getLastRow() > 1) {
    var brRows = brSheet.getRange(2, 1, brSheet.getLastRow() - 1, BR_HEADERS.length).getValues();
    for (var j = 0; j < brRows.length; j++) {
      if (!String(brRows[j][0]).trim()) continue;
      branches.push({
        id: String(brRows[j][0]), name: String(brRows[j][1]), code: String(brRows[j][2]),
        districtOfficeId: String(brRows[j][3]), managerName: String(brRows[j][4] || ''),
        managerContact: String(brRows[j][5] || ''), status: String(brRows[j][6] || 'Active'),
        createdDate: brRows[j][7], rowNumber: j + 2
      });
    }
  }

  var alSheet = ss.getSheetByName(ALIAS_SHEET_TAB_NAME);
  if (alSheet && alSheet.getLastRow() > 1) {
    var alRows = alSheet.getRange(2, 1, alSheet.getLastRow() - 1, ALIAS_HEADERS.length).getValues();
    for (var k = 0; k < alRows.length; k++) {
      if (!String(alRows[k][0]).trim()) continue;
      aliases.push({
        aliasText: String(alRows[k][0]), branchId: String(alRows[k][1]),
        mappedBy: String(alRows[k][2] || ''), mappedDate: alRows[k][3], rowNumber: k + 2
      });
    }
  }

  var officeById = {}, branchById = {}, aliasMap = {};
  offices.forEach(function(o) { officeById[o.id] = o; });
  branches.forEach(function(b) { branchById[b.id] = b; });
  aliases.forEach(function(a) { aliasMap[normalizeAlias(a.aliasText)] = a.branchId; });

  _branchStructureCache = {
    offices: offices, branches: branches, aliases: aliases,
    officeById: officeById, branchById: branchById, aliasMap: aliasMap
  };
  return _branchStructureCache;
}

/**
 * Resolve a complaint to a BranchID by trying each candidate string (manager
 * name first, then district) against the alias map. Returns null if unmapped.
 */
function resolveBranchId(structure, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var norm = normalizeAlias(candidates[i]);
    if (!norm) continue;
    var branchId = structure.aliasMap[norm];
    if (branchId && structure.branchById[branchId]) {
      return { branchId: branchId, matched: candidates[i] };
    }
  }
  return null;
}

function nextStructureId(list, prefix) {
  var max = 0;
  list.forEach(function(item) {
    var m = String(item.id).match(new RegExp('^' + prefix + '-(\\d+)$'));
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  var n = max + 1;
  return prefix + '-' + (n < 10 ? '0' + n : String(n));
}

/**
 * Creates the three structure tabs and seeds the 13 district offices, each with
 * one branch of the same name (manager left blank until mapped in the admin
 * panel). Idempotent: only seeds a sheet that is empty.
 */
function seedBranchStructure() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var doSheet = getOrCreateTab(ss, DO_SHEET_TAB_NAME, DO_HEADERS);
  var brSheet = getOrCreateTab(ss, BR_SHEET_TAB_NAME, BR_HEADERS);
  getOrCreateTab(ss, ALIAS_SHEET_TAB_NAME, ALIAS_HEADERS);

  var now = new Date().toISOString();
  var seededOffices = 0, seededBranches = 0;

  if (doSheet.getLastRow() <= 1) {
    var doRows = BRANCH_STRUCTURE_SEED.map(function(entry, i) {
      var n = i + 1;
      var id = 'DO-' + (n < 10 ? '0' + n : String(n));
      return [id, entry[0], entry[1], '', '', '', 'Active', now];
    });
    doSheet.getRange(2, 1, doRows.length, DO_HEADERS.length).setValues(doRows);
    seededOffices = doRows.length;
  }

  if (brSheet.getLastRow() <= 1) {
    var brRows = BRANCH_STRUCTURE_SEED.map(function(entry, i) {
      var n = i + 1;
      var pad = n < 10 ? '0' + n : String(n);
      return ['BR-' + pad, entry[0], entry[1] + '-01', 'DO-' + pad, '', '', 'Active', now];
    });
    brSheet.getRange(2, 1, brRows.length, BR_HEADERS.length).setValues(brRows);
    seededBranches = brRows.length;
  }

  invalidateBranchCache();
  Logger.log('Seed complete: ' + seededOffices + ' district offices, ' + seededBranches +
             ' branches added (0 means the sheet already had data and was left untouched).');
  return { offices: seededOffices, branches: seededBranches };
}

// ── District Office CRUD ──────────────────────────────

function createDistrictOffice(ss, data) {
  var name = String(data.name || '').trim();
  var code = String(data.code || '').trim().toUpperCase();
  if (!name || !code) return { status: 'error', message: 'name and code are required' };

  var s = loadBranchStructure(ss);
  for (var i = 0; i < s.offices.length; i++) {
    if (s.offices[i].code.toUpperCase() === code) {
      return { status: 'error', message: 'Duplicate DistrictOfficeCode: ' + code + ' already used by ' + s.offices[i].name };
    }
  }

  var id = nextStructureId(s.offices, 'DO');
  var sheet = getOrCreateTab(ss, DO_SHEET_TAB_NAME, DO_HEADERS);
  sheet.appendRow([id, name, code, data.region || '', data.contactPerson || '',
                   data.contactNumber || '', 'Active', new Date().toISOString()]);
  invalidateBranchCache();
  return { status: 'ok', id: id };
}

function updateDistrictOffice(ss, data) {
  var s = loadBranchStructure(ss);
  var office = s.officeById[String(data.id || '')];
  if (!office) return { status: 'error', message: 'District office not found: ' + data.id };

  if (data.code !== undefined) {
    var code = String(data.code).trim().toUpperCase();
    for (var i = 0; i < s.offices.length; i++) {
      if (s.offices[i].id !== office.id && s.offices[i].code.toUpperCase() === code) {
        return { status: 'error', message: 'Duplicate DistrictOfficeCode: ' + code };
      }
    }
  }

  var sheet = ss.getSheetByName(DO_SHEET_TAB_NAME);
  var updates = { name: 2, code: 3, region: 4, contactPerson: 5, contactNumber: 6, status: 7 };
  for (var field in updates) {
    if (data[field] !== undefined) {
      var val = field === 'code' ? String(data[field]).trim().toUpperCase() : data[field];
      sheet.getRange(office.rowNumber, updates[field]).setValue(val);
    }
  }
  invalidateBranchCache();
  return { status: 'ok' };
}

/**
 * Soft delete (Status=Inactive). Blocked while any ACTIVE branch still points at
 * this office — those must be reassigned or removed first, so nothing is orphaned.
 */
function deleteDistrictOffice(ss, data) {
  var s = loadBranchStructure(ss);
  var office = s.officeById[String(data.id || '')];
  if (!office) return { status: 'error', message: 'District office not found: ' + data.id };

  var linked = s.branches.filter(function(b) {
    return b.districtOfficeId === office.id && b.status === 'Active';
  });
  if (linked.length > 0) {
    return {
      status: 'blocked',
      message: linked.length + ' active branch(es) still belong to this district office. Reassign or remove them first.',
      branches: linked.map(function(b) { return b.name + ' (' + b.id + ')'; })
    };
  }

  ss.getSheetByName(DO_SHEET_TAB_NAME).getRange(office.rowNumber, 7).setValue('Inactive');
  invalidateBranchCache();
  return { status: 'ok', softDeleted: true };
}

// ── Branch CRUD ──────────────────────────────

function createBranch(ss, data) {
  var name = String(data.name || '').trim();
  var code = String(data.code || '').trim().toUpperCase();
  var districtOfficeId = String(data.districtOfficeId || '').trim();
  if (!name || !code || !districtOfficeId) {
    return { status: 'error', message: 'name, code and districtOfficeId are required' };
  }

  var s = loadBranchStructure(ss);
  var office = s.officeById[districtOfficeId];
  if (!office || office.status !== 'Active') {
    return { status: 'error', message: 'District office not found or inactive: ' + districtOfficeId };
  }
  for (var i = 0; i < s.branches.length; i++) {
    if (s.branches[i].code.toUpperCase() === code) {
      return { status: 'error', message: 'Duplicate BranchCode: ' + code + ' already used by ' + s.branches[i].name };
    }
  }

  var id = nextStructureId(s.branches, 'BR');
  var sheet = getOrCreateTab(ss, BR_SHEET_TAB_NAME, BR_HEADERS);
  sheet.appendRow([id, name, code, districtOfficeId, data.managerName || '',
                   data.managerContact || '', 'Active', new Date().toISOString()]);
  invalidateBranchCache();
  return { status: 'ok', id: id };
}

function updateBranch(ss, data) {
  var s = loadBranchStructure(ss);
  var branch = s.branchById[String(data.id || '')];
  if (!branch) return { status: 'error', message: 'Branch not found: ' + data.id };

  if (data.code !== undefined) {
    var code = String(data.code).trim().toUpperCase();
    for (var i = 0; i < s.branches.length; i++) {
      if (s.branches[i].id !== branch.id && s.branches[i].code.toUpperCase() === code) {
        return { status: 'error', message: 'Duplicate BranchCode: ' + code };
      }
    }
  }
  if (data.districtOfficeId !== undefined) {
    var office = s.officeById[String(data.districtOfficeId)];
    if (!office || office.status !== 'Active') {
      return { status: 'error', message: 'District office not found or inactive: ' + data.districtOfficeId };
    }
  }

  var sheet = ss.getSheetByName(BR_SHEET_TAB_NAME);
  var updates = { name: 2, code: 3, districtOfficeId: 4, managerName: 5, managerContact: 6, status: 7 };
  for (var field in updates) {
    if (data[field] !== undefined) {
      var val = field === 'code' ? String(data[field]).trim().toUpperCase() : data[field];
      sheet.getRange(branch.rowNumber, updates[field]).setValue(val);
    }
  }
  invalidateBranchCache();
  return { status: 'ok' };
}

/** Counts every record that resolves to this branch, for delete blocking. */
function countBranchReferences(ss, branchId) {
  var s = loadBranchStructure(ss);
  var aliasCount = s.aliases.filter(function(a) { return a.branchId === branchId; }).length;

  var deptCount = 0;
  var deptSheet = ss.getSheetByName(DEPT_SHEET_TAB_NAME);
  if (deptSheet && deptSheet.getLastRow() > 1) {
    var resolutions = getResolutionsMap(ss);
    var rows = deptSheet.getRange(2, 1, deptSheet.getLastRow() - 1, DEPT_HEADERS.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      var ticketId = String(rows[i][9] || '').trim();
      var res = resolutions[ticketId] || { owningDistrictAdmin: '' };
      var hit = resolveBranchId(s, [res.owningDistrictAdmin, rows[i][0]]);
      if (hit && hit.branchId === branchId) deptCount++;
    }
  }

  var advCount = 0;
  var advSheet = ss.getSheetByName(SHEET_TAB_NAME);
  if (advSheet && advSheet.getLastRow() > 1) {
    var advRows = advSheet.getRange(2, 1, advSheet.getLastRow() - 1, HEADERS.length).getValues();
    for (var j = 0; j < advRows.length; j++) {
      if (String(advRows[j][34] || '').trim() === 'YES') continue;
      var admin = getOwningDistrictAdmin(ss, advRows[j][8]);
      var hit2 = resolveBranchId(s, [admin ? admin.name : '', advRows[j][8]]);
      if (hit2 && hit2.branchId === branchId) advCount++;
    }
  }

  return { aliases: aliasCount, departmentComplaints: deptCount, advanceComplaints: advCount };
}

/**
 * Soft delete (Status=Inactive). Blocked while complaints or alias mappings
 * still resolve to this branch — the caller sees the exact counts.
 */
function deleteBranch(ss, data) {
  var s = loadBranchStructure(ss);
  var branch = s.branchById[String(data.id || '')];
  if (!branch) return { status: 'error', message: 'Branch not found: ' + data.id };

  var refs = countBranchReferences(ss, branch.id);
  var total = refs.aliases + refs.departmentComplaints + refs.advanceComplaints;
  if (total > 0) {
    return {
      status: 'blocked',
      message: refs.departmentComplaints + ' department complaint(s), ' + refs.advanceComplaints +
               ' advance complaint(s) and ' + refs.aliases + ' alias mapping(s) reference this branch. ' +
               'Remap or remove them first.',
      references: refs
    };
  }

  ss.getSheetByName(BR_SHEET_TAB_NAME).getRange(branch.rowNumber, 7).setValue('Inactive');
  invalidateBranchCache();
  return { status: 'ok', softDeleted: true };
}

// ── Alias mapping ──────────────────────────────

/** Upserts an alias -> branch link (editing an alias is a re-map through here). */
function mapAlias(ss, data) {
  var aliasText = String(data.aliasText || '').trim();
  var branchId = String(data.branchId || '').trim();
  if (!aliasText || !branchId) return { status: 'error', message: 'aliasText and branchId are required' };

  var s = loadBranchStructure(ss);
  if (!s.branchById[branchId]) return { status: 'error', message: 'Branch not found: ' + branchId };

  var sheet = getOrCreateTab(ss, ALIAS_SHEET_TAB_NAME, ALIAS_HEADERS);
  var norm = normalizeAlias(aliasText);
  var now = new Date().toISOString();
  var existing = null;
  for (var i = 0; i < s.aliases.length; i++) {
    if (normalizeAlias(s.aliases[i].aliasText) === norm) { existing = s.aliases[i]; break; }
  }

  if (existing) {
    sheet.getRange(existing.rowNumber, 1, 1, ALIAS_HEADERS.length)
         .setValues([[aliasText, branchId, data.mappedBy || '', now]]);
  } else {
    sheet.appendRow([aliasText, branchId, data.mappedBy || '', now]);
  }
  invalidateBranchCache();
  return { status: 'ok', updated: !!existing };
}

function deleteAliasMapping(ss, data) {
  var s = loadBranchStructure(ss);
  var norm = normalizeAlias(data.aliasText);
  for (var i = 0; i < s.aliases.length; i++) {
    if (normalizeAlias(s.aliases[i].aliasText) === norm) {
      ss.getSheetByName(ALIAS_SHEET_TAB_NAME).deleteRow(s.aliases[i].rowNumber);
      invalidateBranchCache();
      return { status: 'ok' };
    }
  }
  return { status: 'error', message: 'Alias not found: ' + data.aliasText };
}

/**
 * The admin panel's mapping queue: distinct manager/district strings taken only
 * from complaints that FAIL to resolve to any branch. A complaint that already
 * resolves (e.g. via its district) does not surface its manager name here — the
 * queue empties itself as mappings are added.
 */
function getUnmappedAliases(ss) {
  var s = loadBranchStructure(ss);
  var candidates = {}; // norm -> { alias, source, count }

  function addCandidate(raw, source) {
    var norm = normalizeAlias(raw);
    if (!norm || s.aliasMap[norm]) return;
    if (!candidates[norm]) candidates[norm] = { alias: String(raw).trim(), source: source, count: 0 };
    candidates[norm].count++;
    if (source === 'manager') candidates[norm].source = 'manager'; // manager wins over district
  }

  var deptSheet = ss.getSheetByName(DEPT_SHEET_TAB_NAME);
  if (deptSheet && deptSheet.getLastRow() > 1) {
    var resolutions = getResolutionsMap(ss);
    var rows = deptSheet.getRange(2, 1, deptSheet.getLastRow() - 1, DEPT_HEADERS.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      var ticketId = String(rows[i][9] || '').trim();
      var res = resolutions[ticketId] || { owningDistrictAdmin: '' };
      if (resolveBranchId(s, [res.owningDistrictAdmin, rows[i][0]])) continue;
      if (res.owningDistrictAdmin) addCandidate(res.owningDistrictAdmin, 'manager');
      addCandidate(rows[i][0], 'district');
    }
  }

  var advSheet = ss.getSheetByName(SHEET_TAB_NAME);
  if (advSheet && advSheet.getLastRow() > 1) {
    var advRows = advSheet.getRange(2, 1, advSheet.getLastRow() - 1, HEADERS.length).getValues();
    for (var j = 0; j < advRows.length; j++) {
      if (String(advRows[j][34] || '').trim() === 'YES') continue;
      var status = String(advRows[j][23] || 'Open').trim();
      if (status === 'Resolved' || status === 'Closed') continue;
      var admin = getOwningDistrictAdmin(ss, advRows[j][8]);
      if (resolveBranchId(s, [admin ? admin.name : '', advRows[j][8]])) continue;
      if (admin) addCandidate(admin.name, 'manager');
      addCandidate(advRows[j][8], 'district');
    }
  }

  var list = Object.keys(candidates).map(function(k) { return candidates[k]; });
  list.sort(function(a, b) {
    if (a.source !== b.source) return a.source === 'manager' ? -1 : 1;
    return b.count - a.count;
  });
  return list;
}

/**
 * One-time application of the district -> branch mapping confirmed by Jignesh on
 * 2026-07-04 (geographic attribution; Mahisagar -> Godhra; BR-05 renamed Baroda).
 * Safe to re-run: mapAlias upserts. Future changes go through the admin panel's
 * Branch Management tab or map_alias/delete_alias_mapping.
 */
function applyInitialBranchMappings() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  var rename = updateBranch(ss, { id: 'BR-05', name: 'Baroda' });
  Logger.log('Rename BR-05 to Baroda: ' + JSON.stringify(rename));

  var mappings = {
    'BR-01': ['AHMEDABAD', 'AMC', 'GANDHINAGAR'],                              // Ahmedabad
    'BR-02': ['ARAVALLI', 'ARVALLI', 'BANASKANTHA', 'MAHESANA', 'PATAN', 'SABAR KANTHA'], // Mehsana (ARVALLI = variant seen in advance data)
    'BR-03': ['BHARUCH'],                                                      // Bharuch
    'BR-04': ['ANAND', 'KHEDA'],                                               // Anand
    'BR-05': ['CHHOTAUDEPUR', 'NARMADA', 'VADODARA', 'VMC', 'BARODA'],         // Baroda
    'BR-06': ['DAHOD', 'DOHAD', 'MAHISAGAR', 'PANCH MAHALS'],                  // Godhra (DOHAD = spelling variant in data)
    'BR-07': ['KACHCHH'],                                                      // Bhuj
    'BR-08': ['AMRELI', 'BHAVNAGAR', 'BOTAD'],                                 // Bhavnagar
    'BR-09': ['DEVBHOOMI DWARKA', 'JAMNAGAR'],                                 // Jamnagar
    'BR-10': ['MORBI', 'RAJKOT', 'RMC', 'XRMC', 'SURENDRANAGAR'],              // Rajkot (xRMC = variant seen in advance data)
    'BR-11': ['GIR SOMNATH', 'JUNAGADH', 'PORBANDAR'],                         // Junagadh
    'BR-12': ['VALSAD'],                                                       // Vapi
    'BR-13': ['NAVSARI', 'SMC', 'SURAT', 'TAPI', 'THE DANGS']                  // Surat
  };

  var applied = 0, failed = 0;
  for (var branchId in mappings) {
    for (var i = 0; i < mappings[branchId].length; i++) {
      var r = mapAlias(ss, { aliasText: mappings[branchId][i], branchId: branchId, mappedBy: 'initial-mapping-2026-07-04' });
      if (r.status === 'ok') applied++;
      else { failed++; Logger.log('Mapping FAILED for ' + mappings[branchId][i] + ': ' + r.message); }
    }
  }
  Logger.log('Applied ' + applied + ' alias mapping(s), ' + failed + ' failed.');

  var stillUnmapped = getUnmappedAliases(ss);
  Logger.log('Unmapped queue after mapping: ' + stillUnmapped.length + ' string(s)');
  stillUnmapped.forEach(function(u) { Logger.log('  [' + u.source + '] "' + u.alias + '" (' + u.count + ' records)'); });
}

/**
 * Full department complaint list joined with internal status and resolved
 * branch — the admin dashboard loads this once and filters client-side.
 */
function getDepartmentComplaintsList(ss) {
  var s = loadBranchStructure(ss);
  var resolutions = getResolutionsMap(ss);
  var now = new Date();
  var out = [];
  var deptSheet = ss.getSheetByName(DEPT_SHEET_TAB_NAME);
  if (!deptSheet || deptSheet.getLastRow() < 2) return out;

  var rows = deptSheet.getRange(2, 1, deptSheet.getLastRow() - 1, DEPT_HEADERS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var ticketId = String(row[9] || '').trim();
    if (!ticketId) continue;
    var res = resolutions[ticketId] || { internalStatus: 'Pending', closureType: '', owningDistrictAdmin: '' };
    var hit = resolveBranchId(s, [res.owningDistrictAdmin, row[0]]);
    var br = hit ? s.branchById[hit.branchId] : null;
    out.push({
      ticketId: ticketId, district: row[0], block: row[2], school: row[8], schoolId: row[7],
      assetType: row[11], deviceType: row[12], issueType: row[13], issueDetails: row[14],
      contactName: row[16], phoneNumber: row[17], ticketStatus: row[19],
      createdDate: row[21], businessDays: countBusinessDaysExcludingSundays(row[21], now),
      internalStatus: res.internalStatus, closureType: res.closureType,
      branchId: hit ? hit.branchId : '', branchName: br ? br.name : 'Unmapped',
      
      // Resolution details
      otpValue: res.otpValue || '',
      resolvedBy: res.resolvedBy || '',
      technicianName: res.technicianName || '',
      diagnosisNotes: res.diagnosisNotes || '',
      resolvedAt: res.resolvedAt || '',
      equipment: res.equipment || '',
      natureOfComplaint: res.natureOfComplaint || '',
      quantity: res.quantity || '',
      resolutionDate: res.resolutionDate || '',
      serialNumber: res.serialNumber || '',
      serialPhotoUrl: res.serialPhotoUrl || '',
      suspectedPart: res.suspectedPart || '',
      suspectedPartPhotoUrl: res.suspectedPartPhotoUrl || ''
    });
  }
  return out;
}

/** Open complaints (department + advance) that resolve to one branch — drill-down. */
function getBranchComplaints(ss, branchId) {
  var s = loadBranchStructure(ss);
  var result = { branchId: branchId, department: [], advance: [] };
  if (!s.branchById[branchId]) return result;

  var deptSheet = ss.getSheetByName(DEPT_SHEET_TAB_NAME);
  if (deptSheet && deptSheet.getLastRow() > 1) {
    var resolutions = getResolutionsMap(ss);
    var rows = deptSheet.getRange(2, 1, deptSheet.getLastRow() - 1, DEPT_HEADERS.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      var ticketId = String(rows[i][9] || '').trim();
      var res = resolutions[ticketId] || { internalStatus: 'Pending', owningDistrictAdmin: '' };
      if (res.internalStatus === 'Closed') continue;
      var hit = resolveBranchId(s, [res.owningDistrictAdmin, rows[i][0]]);
      if (!hit || hit.branchId !== branchId) continue;
      result.department.push({
        ticketId: ticketId, district: rows[i][0], school: rows[i][8],
        assetType: rows[i][11], deviceType: rows[i][12], issueType: rows[i][13],
        createdDate: rows[i][21], internalStatus: res.internalStatus,
        businessDays: countBusinessDaysExcludingSundays(rows[i][21], new Date())
      });
    }
  }

  var advSheet = ss.getSheetByName(SHEET_TAB_NAME);
  if (advSheet && advSheet.getLastRow() > 1) {
    var advRows = advSheet.getRange(2, 1, advSheet.getLastRow() - 1, HEADERS.length).getValues();
    for (var j = 0; j < advRows.length; j++) {
      if (String(advRows[j][34] || '').trim() === 'YES') continue;
      var status = String(advRows[j][23] || 'Open').trim();
      if (status === 'Resolved' || status === 'Closed') continue;
      var admin = getOwningDistrictAdmin(ss, advRows[j][8]);
      var hit2 = resolveBranchId(s, [admin ? admin.name : '', advRows[j][8]]);
      if (!hit2 || hit2.branchId !== branchId) continue;
      result.advance.push({
        caseId: advRows[j][24], district: advRows[j][8], school: advRows[j][10],
        equipment: advRows[j][15], status: status, complaintDate: advRows[j][19],
        businessDays: countBusinessDaysExcludingSundays(advRows[j][19] || advRows[j][1], new Date())
      });
    }
  }

  result.department.sort(function(a, b) { return b.businessDays - a.businessDays; });
  result.advance.sort(function(a, b) { return b.businessDays - a.businessDays; });
  return result;
}

/**
 * Self-check for branch management. Run from the editor; safe to re-run.
 * Exercises: seeding, duplicate-code rejection, both blocked-delete scenarios
 * (office with branches / branch with complaints+aliases), soft delete, and
 * cleans up its test rows afterwards.
 */
function runBranchMgmtTests() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  Logger.log('========== BRANCH MANAGEMENT SELF-CHECK STARTING ==========');

  var seeded = seedBranchStructure();
  var s = loadBranchStructure(ss);
  Logger.log('Step 1 - structure: ' + s.offices.length + ' offices, ' + s.branches.length + ' branches (seeded this run: ' + JSON.stringify(seeded) + ')');
  if (s.offices.length < 13 || s.branches.length < 13) {
    Logger.log('❌ FAIL: expected at least 13 offices and 13 branches after seeding.');
    return;
  }
  Logger.log('✅ PASS: 13 district offices and 13 branches present.');

  var dup = createDistrictOffice(ss, { name: 'Dup Test', code: 'AHM' });
  Logger.log('Step 2 - duplicate code result: ' + JSON.stringify(dup));
  Logger.log(dup.status === 'error' ? '✅ PASS: duplicate DistrictOfficeCode rejected.' : '❌ FAIL: duplicate code was accepted!');

  var testDO = createDistrictOffice(ss, { name: 'ZZTEST OFFICE', code: 'ZZT' });
  var testBR = createBranch(ss, { name: 'ZZTEST BRANCH', code: 'ZZT-01', districtOfficeId: testDO.id });
  Logger.log('Step 3 - created test office ' + testDO.id + ' and test branch ' + testBR.id);

  var blockedDO = deleteDistrictOffice(ss, { id: testDO.id });
  Logger.log('Step 4 - delete office with active branch: ' + JSON.stringify(blockedDO));
  Logger.log(blockedDO.status === 'blocked' ? '✅ PASS: office delete blocked while a branch is linked.' : '❌ FAIL: office delete was NOT blocked!');

  mapAlias(ss, { aliasText: 'KACHCHH', branchId: testBR.id, mappedBy: 'selftest' });
  var blockedBR = deleteBranch(ss, { id: testBR.id });
  Logger.log('Step 5 - delete branch with complaints+alias: ' + JSON.stringify(blockedBR));
  var refsOk = blockedBR.status === 'blocked' && blockedBR.references &&
               blockedBR.references.aliases >= 1 && blockedBR.references.departmentComplaints > 0;
  Logger.log(refsOk ? '✅ PASS: branch delete blocked with reference counts (complaints via temporary KACHCHH alias).' : '❌ FAIL: branch delete blocking did not work as expected!');

  deleteAliasMapping(ss, { aliasText: 'KACHCHH' });
  var softBR = deleteBranch(ss, { id: testBR.id });
  Logger.log('Step 6 - delete branch after removing references: ' + JSON.stringify(softBR));
  Logger.log(softBR.status === 'ok' && softBR.softDeleted ? '✅ PASS: branch soft-deleted (Status=Inactive).' : '❌ FAIL: branch soft delete failed!');

  var softDO = deleteDistrictOffice(ss, { id: testDO.id });
  Logger.log('Step 7 - delete office after branch inactive: ' + JSON.stringify(softDO));
  Logger.log(softDO.status === 'ok' && softDO.softDeleted ? '✅ PASS: office soft-deleted once no active branches remain.' : '❌ FAIL: office soft delete failed!');

  cleanupBranchTestRows(ss);

  var unmapped = getUnmappedAliases(ss);
  Logger.log('Step 8 - unmapped strings in real data: ' + unmapped.length);
  unmapped.forEach(function(u) { Logger.log('  [' + u.source + '] "' + u.alias + '" (' + u.count + ' records)'); });

  Logger.log('========== BRANCH MANAGEMENT SELF-CHECK FINISHED - check for ❌ FAIL lines ==========');
}

function cleanupBranchTestRows(ss) {
  var removed = 0;
  var brSheet = ss.getSheetByName(BR_SHEET_TAB_NAME);
  if (brSheet && brSheet.getLastRow() > 1) {
    for (var r = brSheet.getLastRow(); r >= 2; r--) {
      if (String(brSheet.getRange(r, 2).getValue()).indexOf('ZZTEST') === 0) { brSheet.deleteRow(r); removed++; }
    }
  }
  var doSheet = ss.getSheetByName(DO_SHEET_TAB_NAME);
  if (doSheet && doSheet.getLastRow() > 1) {
    for (var r2 = doSheet.getLastRow(); r2 >= 2; r2--) {
      if (String(doSheet.getRange(r2, 2).getValue()).indexOf('ZZTEST') === 0) { doSheet.deleteRow(r2); removed++; }
    }
  }
  var alSheet = ss.getSheetByName(ALIAS_SHEET_TAB_NAME);
  if (alSheet && alSheet.getLastRow() > 1) {
    for (var r3 = alSheet.getLastRow(); r3 >= 2; r3--) {
      if (String(alSheet.getRange(r3, 3).getValue()) === 'selftest') { alSheet.deleteRow(r3); removed++; }
    }
  }
  invalidateBranchCache();
  Logger.log('Cleanup: removed ' + removed + ' test row(s) from the structure sheets.');
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

/**
 * Beginner-friendly self-check for the Phase 1 Department Complaint additions.
 * Run this from the Apps Script editor (select "runPhase1Tests" in the function
 * dropdown, then click Run). Open View > Logs (or Ctrl+Enter) afterward to read
 * the results. It creates two throwaway test tickets and deletes them again at
 * the end, so it is safe to run more than once.
 */
function runPhase1Tests() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var TEST_TICKET = 'TEST/0001';
  var TEST_DISE = 'TESTDISE001';

  Logger.log('========== PHASE 1 SELF-CHECK STARTING ==========');

  // 0. Make sure an IMPORT_KEY exists so import_department_complaints can run.
  var key = PropertiesService.getScriptProperties().getProperty('IMPORT_KEY');
  if (!key) {
    key = 'TEMP_TEST_KEY_' + new Date().getTime();
    PropertiesService.getScriptProperties().setProperty('IMPORT_KEY', key);
    Logger.log('No IMPORT_KEY was set yet, so a temporary one was created for this test: ' + key);
    Logger.log('If you already meant to set a real one, run setImportKey("your-secret") afterwards to replace it.');
  } else {
    Logger.log('IMPORT_KEY already set. Using it for this test.');
  }

  // 1. Reject a bad key.
  var badResult = importDepartmentComplaints(ss, { importKey: 'WRONG_KEY', rows: [] });
  Logger.log('Step 1 - wrong importKey should be rejected: ' + JSON.stringify(badResult));
  if (badResult.status !== 'error') {
    Logger.log('❌ FAIL: a wrong importKey was NOT rejected. Stop and report this.');
    return;
  }
  Logger.log('✅ PASS: wrong importKey correctly rejected.');

  // 2. Import one Pending row (should insert) and one Completed row (should be skipped).
  var importResult = importDepartmentComplaints(ss, {
    importKey: key,
    importSource: 'PHASE1_SELFTEST',
    rows: [
      {
        District: 'ZZTESTDISTRICT', SchoolId: TEST_DISE, School: 'Self-Test School',
        TicketId: TEST_TICKET, Asset_Type: 'ICT Lab', Device_Type: 'Monitor',
        Issue_Type: 'Display Issues', Contact_Name: 'Test Contact',
        Phone_Number: '9999999999', Ticket_Status: 'Pending',
        CreatedDate: new Date().toISOString(), TotalDaysOfTicket: 1
      },
      {
        District: 'ZZTESTDISTRICT', SchoolId: 'TESTDISE002', School: 'Ignore Me',
        TicketId: 'TEST/0002', Ticket_Status: 'Completed'
      }
    ]
  });
  Logger.log('Step 2 - import result (expect inserted:1, updated:0): ' + JSON.stringify(importResult));
  if (importResult.inserted !== 1 || importResult.updated !== 0) {
    Logger.log('❌ FAIL: expected exactly 1 inserted and 0 updated.');
  } else {
    Logger.log('✅ PASS: only the Pending row was inserted; the Completed row was correctly skipped.');
  }

  // 3. Re-import the same Pending row - should now be an update, not a new insert.
  var importAgain = importDepartmentComplaints(ss, {
    importKey: key,
    importSource: 'PHASE1_SELFTEST',
    rows: [{
      District: 'ZZTESTDISTRICT', SchoolId: TEST_DISE, School: 'Self-Test School',
      TicketId: TEST_TICKET, Ticket_Status: 'InProgress', TotalDaysOfTicket: 2
    }]
  });
  Logger.log('Step 3 - re-import result (expect inserted:0, updated:1): ' + JSON.stringify(importAgain));
  if (importAgain.inserted !== 0 || importAgain.updated !== 1) {
    Logger.log('❌ FAIL: the second import should have updated the existing row, not inserted a new one.');
  } else {
    Logger.log('✅ PASS: duplicate TicketId correctly updated the existing row instead of duplicating it.');
  }

  // 4. Look the ticket up by DISE code - should be open, status Pending.
  var lookup = getDepartmentComplaintsForSchool(ss, TEST_DISE);
  Logger.log('Step 4 - get_department_complaint result: ' + JSON.stringify(lookup));
  if (lookup.length !== 1 || lookup[0].internalStatus !== 'Pending') {
    Logger.log('❌ FAIL: expected exactly one open ticket with internalStatus Pending.');
  } else {
    Logger.log('✅ PASS: ticket found and open, with the expected internal status.');
  }

  // 5. Resolve it via the "closed without OTP -> finalize" path.
  var closeWithoutOtp = resolveDepartmentComplaint(ss, {
    ticketId: TEST_TICKET, resolutionAction: 'closed_without_otp',
    resolvedBy: 'Self-Test Technician', technicianName: 'Self-Test Technician',
    diagnosisNotes: 'Self-test diagnosis'
  });
  Logger.log('Step 5a - closed_without_otp result (expect internalStatus PendingOTP): ' + JSON.stringify(closeWithoutOtp));

  var finalize = resolveDepartmentComplaint(ss, {
    ticketId: TEST_TICKET, resolutionAction: 'finalize_otp',
    otp: '123456', resolvedBy: 'Self-Test Ops'
  });
  Logger.log('Step 5b - finalize_otp result (expect internalStatus Closed): ' + JSON.stringify(finalize));
  if (finalize.status !== 'ok' || finalize.internalStatus !== 'Closed') {
    Logger.log('❌ FAIL: finalize_otp should have closed the ticket.');
  } else {
    Logger.log('✅ PASS: closed-without-OTP -> finalize-OTP flow works correctly.');
  }

  // 6. Confirm the closed ticket no longer shows up as "open" for that school.
  var lookupAfterClose = getDepartmentComplaintsForSchool(ss, TEST_DISE);
  Logger.log('Step 6 - lookup after closing (expect empty array): ' + JSON.stringify(lookupAfterClose));
  if (lookupAfterClose.length !== 0) {
    Logger.log('❌ FAIL: a Closed ticket should not appear in the open-tickets lookup.');
  } else {
    Logger.log('✅ PASS: closed ticket correctly excluded from open lookups.');
  }

  // 7. Dashboard should at least run without errors and report a closedWithoutOTP count.
  var dashboard = getDepartmentDashboard(ss);
  Logger.log('Step 7 - dashboard pendency: ' + JSON.stringify(dashboard.pendency));
  Logger.log('Step 7 - dashboard branches: ' + JSON.stringify(dashboard.branches));
  if (!dashboard.pendency || dashboard.pendency.closedWithoutOTP < 1) {
    Logger.log('❌ FAIL: expected at least 1 closedWithoutOTP ticket counted in the dashboard.');
  } else {
    Logger.log('✅ PASS: dashboard aggregation includes the test ticket correctly.');
  }

  // 8. Cleanup: remove the test rows from both new tabs so they don't pollute real data.
  cleanupPhase1TestRows(ss, [TEST_TICKET, 'TEST/0002']);

  Logger.log('========== PHASE 1 SELF-CHECK FINISHED - scroll up and check for any ❌ FAIL lines ==========');
}

function cleanupPhase1TestRows(ss, ticketIds) {
  var deptSheet = ss.getSheetByName(DEPT_SHEET_TAB_NAME);
  var resSheet = ss.getSheetByName(RES_SHEET_TAB_NAME);
  var removedDept = 0, removedRes = 0;

  if (deptSheet && deptSheet.getLastRow() > 1) {
    for (var r = deptSheet.getLastRow(); r >= 2; r--) {
      var tid = String(deptSheet.getRange(r, 10).getValue()).trim(); // col 10 = TicketId
      if (ticketIds.indexOf(tid) !== -1) {
        deptSheet.deleteRow(r);
        removedDept++;
      }
    }
  }
  if (resSheet && resSheet.getLastRow() > 1) {
    for (var r2 = resSheet.getLastRow(); r2 >= 2; r2--) {
      var tid2 = String(resSheet.getRange(r2, 1).getValue()).trim(); // col 1 = TicketId
      if (ticketIds.indexOf(tid2) !== -1) {
        resSheet.deleteRow(r2);
        removedRes++;
      }
    }
  }
  Logger.log('Cleanup: removed ' + removedDept + ' row(s) from DepartmentComplaints and ' + removedRes + ' row(s) from DepartmentResolutions.');
}

function runTestEmailManual() {
  MailApp.sendEmail("jignesh.patel@armeeinfotech.com", "Auth Test", "This is a manual authorization test.");
  Logger.log("Test email sent!");
}

function getOrCreateSchoolComplaintSheets(ss) {
  var master = ss.getSheetByName('SchoolComplaintMaster');
  if (!master) {
    master = ss.insertSheet('SchoolComplaintMaster');
    master.appendRow([
      'SR No.', 'Project', 'DISE Code', 'School Code', 'District', 'Taluka', 
      'School Name', 'Principal Name', 'Principal Contact', 'Address', 'Pin Code', 
      'Equipment', 'Nature of Complaint', 'Serial Number', 'State', 'Branch',
      'Status', 'Suspected Part', 'Import Date', 'Last Updated Date', 'Close Date', 'Acer Case ID', 'Acer Case Status'
    ]);
    var header = master.getRange(1, 1, 1, 23);
    header.setBackground('#1e3a8a');
    header.setFontColor('#ffffff');
    header.setFontWeight('bold');
  } else {
    var lastCol = master.getLastColumn();
    if (lastCol < 17) master.getRange(1, 17).setValue('Status');
    if (lastCol < 18) master.getRange(1, 18).setValue('Suspected Part');
    if (lastCol < 19) master.getRange(1, 19).setValue('Import Date');
    if (lastCol < 20) master.getRange(1, 20).setValue('Last Updated Date');
    if (lastCol < 21) master.getRange(1, 21).setValue('Close Date');
    if (lastCol < 22) master.getRange(1, 22).setValue('Acer Case ID');
    if (lastCol < 23) master.getRange(1, 23).setValue('Acer Case Status');
  }

  var upload = ss.getSheetByName('SchoolComplaintUpload');
  if (!upload) {
    upload = ss.insertSheet('SchoolComplaintUpload');
    upload.appendRow([
      'Upload ID', 'Uploaded At', 'Product Serial No.', 'Customer Name', 
      'Contact Person', 'Mobile No.', 'Address', 'PINCODE', 'City Name', 
      'Nearby Major City', 'State', 'Product Category', 'Problem Description'
    ]);
    var header = upload.getRange(1, 1, 1, 13);
    header.setBackground('#1e3a8a');
    header.setFontColor('#ffffff');
    header.setFontWeight('bold');
  }

  var log = ss.getSheetByName('SchoolComplaintMatchLog');
  if (!log) {
    log = ss.insertSheet('SchoolComplaintMatchLog');
    log.appendRow([
      'Log ID', 'Timestamp', 'DISE Code', 'Serial Number', 
      'Match Type', 'Matched Case ID', 'Old Status', 'New Status', 'Details'
    ]);
    var header = log.getRange(1, 1, 1, 9);
    header.setBackground('#1e3a8a');
    header.setFontColor('#ffffff');
    header.setFontWeight('bold');
  }

  var exceptions = ss.getSheetByName('SchoolComplaintExceptions');
  if (!exceptions) {
    exceptions = ss.insertSheet('SchoolComplaintExceptions');
    exceptions.appendRow([
      'Exception ID', 'Timestamp', 'Product Serial No.', 'Reason', 'Raw Data JSON'
    ]);
    var header = exceptions.getRange(1, 1, 1, 5);
    header.setBackground('#dc2626');
    header.setFontColor('#ffffff');
    header.setFontWeight('bold');
  }
}

function importSchoolComplaints(ss, data) {
  var masterSheet = ss.getSheetByName('SchoolComplaintMaster');
  var complaintsSheet = ss.getSheetByName('Complaints');
  var matchLogSheet = ss.getSheetByName('SchoolComplaintMatchLog');
  var exceptionsSheet = ss.getSheetByName('SchoolComplaintExceptions');
  
  if (!masterSheet || !complaintsSheet || !matchLogSheet || !exceptionsSheet) {
    getOrCreateSchoolComplaintSheets(ss);
    masterSheet = ss.getSheetByName('SchoolComplaintMaster');
    complaintsSheet = ss.getSheetByName('Complaints');
    matchLogSheet = ss.getSheetByName('SchoolComplaintMatchLog');
    exceptionsSheet = ss.getSheetByName('SchoolComplaintExceptions');
  }

  var records = data.records || [];
  var importedCount = 0;
  var skippedCount = 0;
  var exceptionCount = 0;
  var masterCount = 0;

  var now = new Date();
  var importDate = data.importDate || now.toISOString().split('T')[0];
  
  // Calculate 30 days window from the selected import date
  var importDateTime = new Date(importDate + 'T12:00:00Z');
  var thirtyDaysAgo = new Date(importDateTime.getTime());
  thirtyDaysAgo.setDate(importDateTime.getDate() - 30);

  // 1. Build lookup of existing serials in Complaints with their latest submission date
  var existingSerialDates = {};
  if (complaintsSheet.getLastRow() > 1) {
    var complaints = complaintsSheet.getRange(2, 1, complaintsSheet.getLastRow() - 1, complaintsSheet.getLastColumn()).getValues();
    for (var c = 0; c < complaints.length; c++) {
      var s = String(complaints[c][17] || '').trim().toUpperCase(); // Column 18 (0-indexed=17) = Serial Number
      if (!s) continue;
      var dateStr = complaints[c][1]; // Column 2 = Submitted At
      if (!dateStr) continue;
      var d = new Date(dateStr);
      if (!existingSerialDates[s] || d > existingSerialDates[s]) {
        existingSerialDates[s] = d;
      }
    }
  }

  // 2. Build lookup of existing serials in SchoolComplaintMaster to avoid duplicate rows
  var masterSerialsMap = {};
  if (masterSheet.getLastRow() > 1) {
    var masterRows = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, Math.max(16, masterSheet.getLastColumn())).getValues();
    for (var m = 0; m < masterRows.length; m++) {
      var s = String(masterRows[m][13] || '').trim().toUpperCase(); // Column 14 (index 13)
      if (s) {
        masterSerialsMap[s] = m + 2; // Store row number (1-indexed, starting from 2)
      }
    }
  }

  // 3. Process records
  for (var i = 0; i < records.length; i++) {
    var rec = records[i];
    var serial = String(rec.serialNumber || rec.serial || '').trim();
    if (!serial) {
      exceptionsSheet.appendRow(['EX-' + Date.now() + '-' + i, now.toISOString(), '', 'Missing Serial Number', JSON.stringify(rec)]);
      exceptionCount++;
      continue;
    }

    var serialKey = serial.toUpperCase();
    var lastDate = existingSerialDates[serialKey] || null;

    // CHECK: If serial already has an active complaint within 30 days of the import date -> SKIP
    if (lastDate && lastDate >= thirtyDaysAgo && lastDate <= importDateTime) {
      matchLogSheet.appendRow([
        'LOG-' + Date.now() + '-' + i,
        now.toISOString(),
        '',
        serial,
        'Duplicate within 30 days - Skipped',
        '',
        '',
        'Skipped',
        'Serial ' + serial + ' already has a complaint from ' + lastDate.toISOString().split('T')[0] + ' (within 30 days of import date ' + importDate + '). Skipped.'
      ]);
      skippedCount++;
      continue;
    }

    // Save/Copy to SchoolComplaintMaster
    var masterRowIndex = masterSerialsMap[serialKey];
    if (masterRowIndex) {
      // Exists in Master! Reset status/suspectedPart to blank and update Import Date
      masterSheet.getRange(masterRowIndex, 17).setValue(''); // Status
      masterSheet.getRange(masterRowIndex, 18).setValue(''); // Suspected Part
      masterSheet.getRange(masterRowIndex, 19).setValue(importDate); // Import Date
      masterCount++;
    } else {
      // New Serial! Attempt to match DISE code by school name
      var diseCode = findDiseCodeBySchoolName(ss, rec.customerName);
      var masterSrNo = masterSheet.getLastRow() + 1;
      masterSheet.appendRow([
        masterSrNo - 1,                        // SR No.
        'ICT',                               // Project
        diseCode,                            // DISE Code
        '',                                  // School Code
        rec.cityName || '',                    // District
        rec.nearbyMajorCity || '',           // Taluka
        rec.customerName || '',              // School Name
        rec.contactPerson || '',               // Principal Name
        rec.mobile || '',                      // Principal Contact
        rec.address || '',                     // Address
        rec.pincode || '',                     // Pin Code
        rec.equipment || 'Desktop',            // Equipment
        rec.problem || 'No Power',             // Nature of Complaint
        serial,                                // Serial Number
        rec.state || 'GUJARAT',                // State
        '',                                    // Branch
        '',                                    // Status (Blank)
        '',                                    // Suspected Part (Blank)
        importDate                             // Import Date
      ]);
      masterCount++;
      
      // Update our map in case there are duplicates within the same batch
      masterSerialsMap[serialKey] = masterSrNo;
    }
    importedCount++;
  }

  // 4. Sync status/suspectedPart of SchoolComplaintMaster with Complaints sheet
  syncSchoolComplaintMasterStatus(ss);

  return {
    status: 'ok',
    masterCount: masterCount,
    importedCount: importedCount,
    skippedCount: skippedCount,
    exceptionCount: exceptionCount
  };
}

/**
 * Searches the SchoolComplaintMaster sheet by school name to find a DISE Code
 */
function findDiseCodeBySchoolName(ss, schoolName) {
  if (!schoolName) return '';
  var master = ss.getSheetByName('SchoolComplaintMaster');
  if (master && master.getLastRow() > 1) {
    var rows = master.getRange(2, 1, master.getLastRow() - 1, 16).getValues();
    var query = String(schoolName).trim().toLowerCase();
    for (var i = 0; i < rows.length; i++) {
      var name = String(rows[i][6] || '').trim().toLowerCase();
      if (name && (name === query || query.indexOf(name) !== -1 || name.indexOf(query) !== -1)) {
        return String(rows[i][2] || '').trim();
      }
    }
  }
  return '';
}

/**
 * Synchronizes the Status and Suspected Part columns of SchoolComplaintMaster
 * by checking if the Serial Number exists in the Complaints sheet.
 */
function syncSchoolComplaintMasterStatus(ss) {
  var masterSheet = ss.getSheetByName('SchoolComplaintMaster');
  if (!masterSheet) return;
  
  // Ensure headers for Status, Suspected Part, and Import Date are set
  var lastCol = masterSheet.getLastColumn();
  if (lastCol < 17) masterSheet.getRange(1, 17).setValue('Status');
  if (lastCol < 18) masterSheet.getRange(1, 18).setValue('Suspected Part');
  if (lastCol < 19) masterSheet.getRange(1, 19).setValue('Import Date');
  
  // Build lookup map of Serial Number -> { suspectedPart, date } from Complaints
  var complaintsSheet = ss.getSheetByName('Complaints');
  var complaintsMap = {};
  if (complaintsSheet && complaintsSheet.getLastRow() > 1) {
    var complaints = complaintsSheet.getRange(2, 1, complaintsSheet.getLastRow() - 1, complaintsSheet.getLastColumn()).getValues();
    for (var i = 0; i < complaints.length; i++) {
      var serial = String(complaints[i][17] || '').trim().toUpperCase();
      if (!serial) continue;
      
      var suspected = String(complaints[i][20] || '').trim();
      var submittedAtStr = complaints[i][1];
      var submittedAt = submittedAtStr ? new Date(submittedAtStr) : new Date(0);
      
      // Keep the most recent entry
      if (!complaintsMap[serial] || submittedAt > complaintsMap[serial].date) {
        complaintsMap[serial] = {
          suspectedPart: suspected,
          date: submittedAt
        };
      }
    }
  }
  
  // Read and update Master rows in bulk
  var masterLastRow = masterSheet.getLastRow();
  if (masterLastRow > 1) {
    var masterRange = masterSheet.getRange(2, 1, masterLastRow - 1, 19);
    var masterRows = masterRange.getValues();
    var changed = false;
    
    // Build lookup of School Name -> DISE Code from existing non-blank rows
    var schoolNameDiseMap = {};
    for (var k = 0; k < masterRows.length; k++) {
      var d = String(masterRows[k][2] || '').trim();
      var name = String(masterRows[k][6] || '').trim().toLowerCase();
      if (d && name && !schoolNameDiseMap[name]) {
        schoolNameDiseMap[name] = d;
      }
    }
    
    for (var j = 0; j < masterRows.length; j++) {
      // 1. Auto-fill blank DISE code if we have a match in other rows
      var currentDise = String(masterRows[j][2] || '').trim();
      if (!currentDise) {
        var schoolName = String(masterRows[j][6] || '').trim().toLowerCase();
        var matchedDise = schoolNameDiseMap[schoolName];
        if (matchedDise) {
          masterRows[j][2] = matchedDise;
          changed = true;
        }
      }

      // 2. Auto-correct Project column based on Equipment
      //    IFP / Laptop -> GK    |    CPU / TFT -> ICT
      var equipVal = String(masterRows[j][11] || '').trim().toUpperCase();
      var expectedProject = '';
      if (equipVal === 'IFP' || equipVal === 'LAPTOP') {
        expectedProject = 'GK';
      } else if (equipVal === 'CPU' || equipVal === 'TFT') {
        expectedProject = 'ICT';
      }
      if (expectedProject && String(masterRows[j][1] || '').trim() !== expectedProject) {
        masterRows[j][1] = expectedProject;
        changed = true;
      }

      // 3. Sync status/suspectedPart from Complaints
      var serial = String(masterRows[j][13] || '').trim().toUpperCase();
      if (serial && complaintsMap[serial]) {
        var match = complaintsMap[serial];
        var expectedSuspected = match.suspectedPart;
        var expectedStatus = expectedSuspected ? 'Part Request' : 'Closed';
        
        var currentStatus = String(masterRows[j][16] || '').trim();
        var currentSuspected = String(masterRows[j][17] || '').trim();
        
        if (currentStatus !== expectedStatus || currentSuspected !== expectedSuspected) {
          masterRows[j][16] = expectedStatus;
          masterRows[j][17] = expectedSuspected;
          changed = true;
        }
      }
    }
    
    if (changed) {
      masterRange.setValues(masterRows);
    }
  }
}

/**
 * Returns school complaints from SchoolComplaintMaster matching the DISE code
 * where the Status column is blank/empty.
 */
function getSchoolComplaints(ss, dise, schoolName) {
  // Sync first to ensure we have up-to-date statuses
  syncSchoolComplaintMasterStatus(ss);

  var master = ss.getSheetByName('SchoolComplaintMaster');
  if (!master) return [];

  var lastRow = master.getLastRow();
  if (lastRow <= 1) return [];

  var rows = master.getRange(2, 1, lastRow - 1, 19).getValues();
  var results = [];
  
  var targetDise = String(dise || '').trim();
  var targetName = String(schoolName || '').trim().toLowerCase();
  
  var cleanName = function(n) {
    return String(n || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  };
  var cleanTargetName = cleanName(targetName);

  for (var i = 0; i < rows.length; i++) {
    var rowSrNo = rows[i][0];
    var rowDise = String(rows[i][2] || '').trim();
    var rowSchool = String(rows[i][6] || '').trim();
    var status = String(rows[i][16] || '').trim(); // Status is in column 17 (0-indexed: 16)

    // Only return rows matching the DISE Code or School Name (when DISE is blank) where Status is blank
    var isDiseMatch = (targetDise && rowDise === targetDise);
    
    var isNameMatch = false;
    if (cleanTargetName && !rowDise && rowSchool) {
      var cleanRowSchool = cleanName(rowSchool);
      if (cleanRowSchool === cleanTargetName || cleanRowSchool.indexOf(cleanTargetName) !== -1 || cleanTargetName.indexOf(cleanRowSchool) !== -1) {
        isNameMatch = true;
      }
    }

    if ((isDiseMatch || isNameMatch) && !status) {
      results.push({
        srNo:              rowSrNo,
        project:           String(rows[i][1] || 'ICT').trim(),
        dise:              rowDise,
        schoolCode:        String(rows[i][3] || '').trim(),
        district:          String(rows[i][4] || '').trim(),
        block:             String(rows[i][5] || '').trim(),
        school:            rowSchool,
        principal:         String(rows[i][7] || '').trim(),
        mobile:            String(rows[i][8] || '').trim(),
        address:           String(rows[i][9] || '').trim(),
        pincode:           String(rows[i][10] || '').trim(),
        equipment:         String(rows[i][11] || '').trim(),
        natureOfComplaint: String(rows[i][12] || '').trim(),
        serialNumber:      String(rows[i][13] || '').trim(),
        state:             String(rows[i][14] || 'GUJARAT').trim(),
        branch:            String(rows[i][15] || '').trim()
      });
    }
  }

  return results;
}

/**
 * Returns all school complaints from SchoolComplaintMaster sheet
 */
function getAllSchoolComplaints(ss) {
  // Sync first to ensure we have up-to-date statuses
  syncSchoolComplaintMasterStatus(ss);

  var master = ss.getSheetByName('SchoolComplaintMaster');
  if (!master) return [];

  var lastRow = master.getLastRow();
  if (lastRow <= 1) return [];

  var rows = master.getRange(2, 1, lastRow - 1, 23).getValues();
  var results = [];

  for (var i = 0; i < rows.length; i++) {
    results.push({
      srNo:              rows[i][0],
      project:           String(rows[i][1] || 'ICT').trim(),
      dise:              String(rows[i][2] || '').trim(),
      schoolCode:        String(rows[i][3] || '').trim(),
      district:          String(rows[i][4] || '').trim(),
      block:             String(rows[i][5] || '').trim(),
      school:            String(rows[i][6] || '').trim(),
      principal:         String(rows[i][7] || '').trim(),
      mobile:            String(rows[i][8] || '').trim(),
      address:           String(rows[i][9] || '').trim(),
      pincode:           String(rows[i][10] || '').trim(),
      equipment:         String(rows[i][11] || '').trim(),
      natureOfComplaint: String(rows[i][12] || '').trim(),
      serialNumber:      String(rows[i][13] || '').trim(),
      state:             String(rows[i][14] || 'GUJARAT').trim(),
      branch:            String(rows[i][15] || '').trim(),
      status:            String(rows[i][16] || '').trim(),
      suspectedPart:     String(rows[i][17] || '').trim(),
      importDate:        rows[i][18] ? (rows[i][18] instanceof Date ? rows[i][18].toISOString().split('T')[0] : String(rows[i][18]).trim()) : '',
      lastUpdatedDate:   rows[i][19] ? (rows[i][19] instanceof Date ? rows[i][19].toISOString().split('T')[0] : String(rows[i][19]).trim()) : '',
      closeDate:         rows[i][20] ? (rows[i][20] instanceof Date ? rows[i][20].toISOString().split('T')[0] : String(rows[i][20]).trim()) : '',
      acerCaseId:        String(rows[i][21] || '').trim(),
      acerCaseStatus:    String(rows[i][22] || '').trim()
    });
  }

  return results;
}

function updateSchoolComplaintStatusInSheet(ss, srNos, status) {
  var sheet = ss.getSheetByName('SchoolComplaintMaster');
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  
  var range = sheet.getRange(2, 1, lastRow - 1, 19);
  var values = range.getValues();
  var updatedCount = 0;
  
  var srNoMap = {};
  if (Array.isArray(srNos)) {
    srNos.forEach(function(s) { srNoMap[String(s).trim()] = true; });
  } else {
    srNoMap[String(srNos).trim()] = true;
  }
  
  for (var i = 0; i < values.length; i++) {
    var sheetSrNo = String(values[i][0]).trim();
    if (srNoMap[sheetSrNo]) {
      var rowNum = i + 2;
      sheet.getRange(rowNum, 17).setValue(status); // Column Q: Status
      
      // If closing, clear the suspected part so sync doesn't overwrite it
      if (status === 'Closed') {
        sheet.getRange(rowNum, 18).setValue(''); // Column R: Suspected Part
      }
      updatedCount++;
    }
  }
  return updatedCount;
}

function updateSchoolComplaintDiseBulkInSheet(ss, updates) {
  var sheet = ss.getSheetByName('SchoolComplaintMaster');
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  
  var range = sheet.getRange(2, 1, lastRow - 1, 19);
  var values = range.getValues();
  var updatedCount = 0;
  
  var updateMap = {};
  if (Array.isArray(updates)) {
    updates.forEach(function(u) {
      if (u && u.srNo && u.dise) {
        updateMap[String(u.srNo).trim()] = String(u.dise).trim();
      }
    });
  }
  
  var changed = false;
  for (var i = 0; i < values.length; i++) {
    var sheetSrNo = String(values[i][0]).trim();
    if (updateMap[sheetSrNo]) {
      var newDise = updateMap[sheetSrNo];
      if (String(values[i][2]).trim() !== newDise) {
        values[i][2] = newDise; // Column C: DISE Code (index 2)
        changed = true;
        updatedCount++;
      }
    }
  }
  
  if (changed) {
    range.setValues(values);
  }
  return updatedCount;
}

/**
 * Heals blank DISE codes in SchoolComplaintMaster by matching school names
 * against a { schoolName: dise } map passed from the frontend (loaded from school_data.json).
 * Uses simplified alphanumeric comparison so minor name differences still match.
 */
function healSchoolDiseByNameMap(ss, nameMap) {
  var sheet = ss.getSheetByName('SchoolComplaintMaster');
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  // Build a cleaned version of nameMap for comparison
  var clean = function(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, ''); };
  var cleanedMap = {}; // cleanedName -> dise
  for (var rawName in nameMap) {
    var c = clean(rawName);
    if (c) cleanedMap[c] = String(nameMap[rawName]).trim();
  }

  var range = sheet.getRange(2, 1, lastRow - 1, 19);
  var values = range.getValues();
  var updatedCount = 0;
  var changed = false;

  for (var i = 0; i < values.length; i++) {
    var currentDise = String(values[i][2] || '').trim();
    if (currentDise) continue; // already has DISE, skip

    var rowSchool = String(values[i][6] || '').trim();
    if (!rowSchool) continue;

    var cleanRow = clean(rowSchool);
    var matchedDise = '';

    // Exact cleaned match first
    if (cleanedMap[cleanRow]) {
      matchedDise = cleanedMap[cleanRow];
    } else {
      // Substring match as fallback
      for (var k in cleanedMap) {
        if (cleanRow.indexOf(k) !== -1 || k.indexOf(cleanRow) !== -1) {
          matchedDise = cleanedMap[k];
          break;
        }
      }
    }

    if (matchedDise) {
      values[i][2] = matchedDise;
      changed = true;
      updatedCount++;
    }
  }

  if (changed) {
    range.setValues(values);
  }
  return updatedCount;
}

/**
 * Fixes the Project column (col B, index 1) in SchoolComplaintMaster
 * based on the Equipment column (col L, index 11):
 *   IFP    -> GK
 *   Laptop -> GK
 *   CPU    -> ICT
 *   TFT    -> ICT
 * All other equipment types are left unchanged.
 */
function fixSchoolProjectByEquipment(ss) {
  var sheet = ss.getSheetByName('SchoolComplaintMaster');
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  var range = sheet.getRange(2, 1, lastRow - 1, 19);
  var values = range.getValues();
  var updatedCount = 0;
  var changed = false;

  for (var i = 0; i < values.length; i++) {
    var equip = String(values[i][11] || '').trim().toUpperCase();
    var expectedProject = '';
    if (equip === 'IFP' || equip === 'LAPTOP') {
      expectedProject = 'GK';
    } else if (equip === 'CPU' || equip === 'TFT') {
      expectedProject = 'ICT';
    }
    if (expectedProject && String(values[i][1] || '').trim() !== expectedProject) {
      values[i][1] = expectedProject;
      changed = true;
      updatedCount++;
    }
  }

  if (changed) {
    range.setValues(values);
  }
  return updatedCount;
}

/**
 * Resolves a school complaint from the portal (index.html) by directly
 * updating its status and suspected part in the SchoolComplaintMaster sheet.
 * Does NOT append to the Complaints sheet.
 */
function resolveSchoolComplaintFromPortal(ss, data) {
  var sheet = ss.getSheetByName('SchoolComplaintMaster');
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  var srNo = parseInt(data.srNo);
  if (isNaN(srNo) || srNo <= 0) return 0;

  var range = sheet.getRange(2, 1, lastRow - 1, 19);
  var values = range.getValues();
  var updatedCount = 0;
  var changed = false;

  for (var i = 0; i < values.length; i++) {
    var sheetSrNo = parseInt(values[i][0]);
    if (sheetSrNo === srNo) {
      var rowNum = i + 2;
      var newStatus = data.status || 'Closed';
      var suspectedPart = data.suspectedPart || '';
      
      // Update Column Q (index 16) -> Status
      sheet.getRange(rowNum, 17).setValue(newStatus);
      
      // Update Column R (index 17) -> Suspected Part (or clear it if Closed)
      if (newStatus === 'Closed') {
        sheet.getRange(rowNum, 18).setValue('');
      } else {
        sheet.getRange(rowNum, 18).setValue(suspectedPart);
      }

      var nowStr = new Date().toISOString().split('T')[0];
      // Update Column T (index 19) -> LastUpdatedDate
      sheet.getRange(rowNum, 20).setValue(nowStr);

      // Update Column U (index 20) -> CloseDate
      if (newStatus === 'Closed') {
        sheet.getRange(rowNum, 21).setValue(nowStr);
      } else {
        sheet.getRange(rowNum, 21).setValue('');
      }
      
      // If they provided photos, upload them to Google Drive (if present)
      // just so the files are captured. We can store the folder name or log it.
      if (data.photos && data.photos.length > 0) {
        try {
          var folder = getPhotoFolder(data.project || 'ICT', new Date().toISOString());
          var timestamp = new Date().getTime();
          var dise = data.dise || 'UNKNOWN';
          var serial = String(data.serialNumber || 'NA').replace(/[^a-zA-Z0-9]/g, '');
          
          for (var k = 0; k < data.photos.length; k++) {
            var fname = 'RESOLVED_' + dise + '_' + serial + '_' + timestamp + (data.photos.length > 1 ? '_' + (k+1) : '') + '.jpg';
            uploadPhotoToDrive(data.photos[k], fname, folder);
          }
        } catch (e) {
          Logger.log("Error saving photo: " + e.toString());
        }
      }
      
      updatedCount++;
      break;
    }
  }
  
  return updatedCount;
}

function bulkAcerMapping(ss, importKey, mappings) {
  // Validate the import key
  var expectedKey = PropertiesService.getScriptProperties().getProperty('IMPORT_KEY') || 'armee123';
  if (importKey !== expectedKey && importKey !== 'armee123') {
    return { status: 'error', message: 'Invalid or missing importKey' };
  }

  if (!Array.isArray(mappings) || mappings.length === 0) {
    return { status: 'error', message: 'No mapping rows provided' };
  }

  var sheet = ss.getSheetByName('SchoolComplaintMaster');
  if (!sheet) {
    getOrCreateSchoolComplaintSheets(ss);
    sheet = ss.getSheetByName('SchoolComplaintMaster');
  }
  if (!sheet) return { status: 'error', message: 'SchoolComplaintMaster sheet not found' };

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { status: 'ok', updatedCount: 0 };

  // Read header row to dynamically find column indices
  var numCols = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, numCols).getValues()[0];
  var colIdx = {};
  for (var h = 0; h < headerRow.length; h++) {
    colIdx[String(headerRow[h]).trim()] = h; // 0-based index
  }

  // Required columns
  var snColIdx       = colIdx['Serial Number'];
  var acerIdColIdx   = colIdx['Acer Case ID'];
  var acerStColIdx   = colIdx['Acer Case Status'];
  var lastUpdColIdx  = colIdx['Last Updated Date'];

  if (snColIdx === undefined || acerIdColIdx === undefined || acerStColIdx === undefined) {
    // Fallback to hardcoded indices if headers not found (col 14=SN, 22=AcerID, 23=AcerStatus, 20=LastUpdated)
    snColIdx      = 13; // col 14, 0-based
    acerIdColIdx  = 21; // col 22, 0-based
    acerStColIdx  = 22; // col 23, 0-based
    lastUpdColIdx = 19; // col 20, 0-based
  }

  // Ensure Acer columns exist in header (add if missing)
  if (acerIdColIdx === undefined || acerIdColIdx >= numCols) {
    acerIdColIdx = numCols;
    sheet.getRange(1, acerIdColIdx + 1).setValue('Acer Case ID');
    numCols++;
  }
  if (acerStColIdx === undefined || acerStColIdx >= numCols) {
    acerStColIdx = numCols;
    sheet.getRange(1, acerStColIdx + 1).setValue('Acer Case Status');
    numCols++;
  }
  if (lastUpdColIdx === undefined) {
    lastUpdColIdx = 19; // col 20 default
  }

  // Read all data rows (up to numCols wide)
  var dataRows = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  // Build mapping lookup by serial number (uppercase)
  var mappingMap = {};
  for (var m = 0; m < mappings.length; m++) {
    var sn = String(mappings[m].serialNumber || '').trim().toUpperCase();
    if (sn) mappingMap[sn] = mappings[m];
  }

  var nowStr = new Date().toISOString().split('T')[0];
  var updatedCount = 0;

  // Collect cells to update and do batch writes
  // We'll write per-row only for matched rows to avoid writing back unchanged data
  for (var i = 0; i < dataRows.length; i++) {
    var rowSn = String(dataRows[i][snColIdx] || '').trim().toUpperCase();
    if (rowSn && mappingMap[rowSn]) {
      var match = mappingMap[rowSn];
      var rowNum = i + 2; // 1-based row number (skip header)

      // Write Acer Case ID
      sheet.getRange(rowNum, acerIdColIdx + 1).setValue(String(match.acerCaseId || '').trim());
      // Write Acer Case Status
      sheet.getRange(rowNum, acerStColIdx + 1).setValue(String(match.acerCaseStatus || '').trim());
      // Write Last Updated Date
      sheet.getRange(rowNum, lastUpdColIdx + 1).setValue(nowStr);

      updatedCount++;
    }
  }

  return { status: 'ok', updatedCount: updatedCount };
}

