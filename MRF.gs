const SHEET_NAME_DATA      = 'Details';
const SHEET_NAME_RESPONSES = 'MRF_Responses';
const CV_FOLDER_ID         = '15VBhbILGlA4v2Qrd_1wkuYhyQphaiHhe';

function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('mrf')
    .setTitle('MRF Form — Sphinx Worldbiz')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var action = data.action;
    var result;

    if (action === 'addOption') {
      result = addOptionToSheet(data.header, data.value);
    } else if (action === 'submitForm') {
      result = saveFormResponse(data);
    } else if (action === 'getLogo') {
      result = getLogoBase64();
    } else if (action === 'uploadCV') {
      result = uploadCVToDrive(data.fileName, data.fileData, data.mimeType);
    } else {
      result = { success: false, message: 'Unknown action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── One-time Drive authorization helper ───────────────────────────────────────
// IMPORTANT: Run this function MANUALLY once in the editor (▶ Run).
// A Google Drive permission consent screen will appear — click Allow.
function authorizeDrive() {
  var folder = DriveApp.getFolderById(CV_FOLDER_ID);
  Logger.log('Drive access OK. Folder found: ' + folder.getName());
}

// ── Upload CV to Drive Folder ─────────────────────────────────────────────────
function uploadCVToDrive(fileName, fileDataBase64, mimeType) {
  try {
    if (!fileDataBase64) {
      return { success: false, message: 'File data missing (empty base64).' };
    }

    var folder = DriveApp.getFolderById(CV_FOLDER_ID);

    // Sanitize mimeType — fallback if blank/invalid
    var mt = (mimeType && (mimeType + '').indexOf('/') !== -1)
               ? mimeType
               : 'application/octet-stream';

    // Sanitize fileName — use timestamp name if blank
    var nm = (fileName && (fileName + '').trim())
               ? (fileName + '').trim()
               : ('CV_' + new Date().getTime());

    // Clean base64 string (remove data: prefix or whitespace if present)
    var clean = (fileDataBase64 + '').replace(/^data:[^,]*,/, '').replace(/\s/g, '');

    var decoded = Utilities.base64Decode(clean);
    if (!decoded || !decoded.length) {
      return { success: false, message: 'Decoded file is empty.' };
    }

    var blob = Utilities.newBlob(decoded, mt, nm);
    var file = folder.createFile(blob);

    // Try sharing separately — so the file still gets created even if Workspace policy blocks it
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      // Sharing failed, but the file was still created — return the link anyway
    }

    var fileUrl = 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=sharing';
    return { success: true, fileUrl: fileUrl, fileId: file.getId() };

  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ── Serve Logo from Drive as Base64 ──────────────────────────────────────────
function getLogoBase64() {
  try {
    var fileId = '1SxPs_gf83gCWWZxF_9gQ9QtCaymmcOLT';
    var file   = DriveApp.getFileById(fileId);
    var blob   = file.getBlob();
    var b64    = Utilities.base64Encode(blob.getBytes());
    var mime   = blob.getContentType();
    return { success: true, dataUrl: 'data:' + mime + ';base64,' + b64 };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ── Details Column Mapping (FIXED — no skipped columns now) ─────────────
//   Col A (0) = Client Name
//   Col B (1) = Account Manager Name
//   Col C (2) = Job Title
//   Col D (3) = Location
//   Col E (4) = Recruiter Assign Person Name
//   Col F (5) = Skills
//   Col G (6) = Department

function getSheetDropdowns() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return {};

  var allData = sheet.getRange(1, 1, lastRow, 7).getValues();

  var firstCell = (allData[0][0] + '').trim().toLowerCase();
  var HEADER_HINTS = [
    'client name','client','name','sl no','sl.no','s.no','sr no',
    'account','account manager','job title','location','skills','department'
  ];
  var hasHeader = HEADER_HINTS.indexOf(firstCell) !== -1;
  var startRow  = hasHeader ? 1 : 0;

  function colVals(colIdx) {
    var vals = [], seen = {};
    for (var r = startRow; r < allData.length; r++) {
      var raw = allData[r][colIdx];
      var v   = (raw + '').trim();
      if (!v || v === 'undefined' || v === 'null') continue;
      if (v.slice(-1) === ':') continue;
      var vl = v.toLowerCase();
      if (!seen[vl]) { seen[vl] = 1; vals.push(v); }
    }
    return vals;
  }

  return {
    'Client Name'                  : colVals(0),
    'Account Manager Name'         : colVals(1),
    'Job Title'                    : colVals(2),
    'Location'                     : colVals(3),
    'Recruiter Assign Person Name' : colVals(4),
    'Skills'                       : colVals(5),
    'Department'                   : colVals(6)
  };
}

// ── Add new option to Details (FIXED mapping) ─────────────────────────────────
var HEADER_COL_MAP = {
  'Client Name'                 : 0,
  'Account Manager Name'        : 1,
  'Job Title'                   : 2,
  'Location'                    : 3,
  'Recruiter Assign Person Name': 4,
  'Skills'                      : 5,
  'Department'                  : 6
};

function addOptionToSheet(header, value) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_DATA);
  if (!sheet)            return { success: false, message: 'Details not found.' };
  if (!header || !value) return { success: false, message: 'Header or value missing.' };

  var trimmedVal = value.trim();
  var colIndex   = HEADER_COL_MAP[header];
  if (colIndex === undefined)
    return { success: false, message: 'Unknown header: ' + header };

  var lastRow = sheet.getLastRow();
  var colData = lastRow > 0
    ? sheet.getRange(1, colIndex + 1, lastRow, 1).getValues().flat()
    : [];

  var lower = colData.map(function(v){ return (v+'').trim().toLowerCase(); });
  if (lower.indexOf(trimmedVal.toLowerCase()) !== -1)
    return { success: true, message: 'Already exists.' };

  var writeRow = lastRow + 1;
  for (var i = 0; i < colData.length; i++) {
    if ((colData[i] + '').trim() === '') { writeRow = i + 1; break; }
  }
  sheet.getRange(writeRow, colIndex + 1).setValue(trimmedVal);
  return { success: true, message: '"' + trimmedVal + '" added to ' + header + '.' };
}

// ── Generate next Unique ID (JD1001, JD1002, ...) from Column B ─────────────
function getNextUniqueId(sheet) {
  var lastRow = sheet.getLastRow();
  var maxNum  = 1000; // First ID will be JD1001

  if (lastRow >= 2) {
    var ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Column B, starting row 2
    for (var i = 0; i < ids.length; i++) {
      var idStr = (ids[i][0] + '').trim();
      var match = idStr.match(/^JD(\d+)$/i);
      if (match) {
        var num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    }
  }
  return 'JD' + (maxNum + 1);
}

// ── Save form response ────────────────────────────────────────────────────────
function saveFormResponse(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_RESPONSES);

  var hdrs = [
    'Timestamp','Unique ID','Business Unit','Client Name / Department','Account Manager Name',
    'Account Type','Job Title','Skills','Recruitment Type','Position Type',
    'Open Positions','Location','Job Opening Date','JD',
    'CV Link',
    'Recruiter Assign Person Name','TAT (Days)',
    'Requirement Status','No Of CV Planned Submit'
  ];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_RESPONSES);
    sheet.getRange(1,1,1,hdrs.length).setValues([hdrs]);
    sheet.getRange(1,1,1,hdrs.length)
         .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, hdrs.length, 160);
  } else {
    // If it's an old sheet and Column B doesn't have the "Unique ID" header,
    // insert a new column B and add the header (existing data shifts to column B)
    var existingHeader = (sheet.getRange(1, 2).getValue() + '').trim();
    if (existingHeader !== 'Unique ID') {
      sheet.insertColumnAfter(1);
      sheet.getRange(1, 2).setValue('Unique ID')
           .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
    }

    // If the old sheet is missing the 2 new columns (Requirement Status, No Of CV Planned Submit),
    // add them at the end
    var lastCol = sheet.getLastColumn();
    var existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){
      return (h + '').trim();
    });
    if (existingHeaders.indexOf('Requirement Status') === -1) {
      sheet.getRange(1, lastCol + 1).setValue('Requirement Status')
           .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
      lastCol = lastCol + 1;
    }
    if (existingHeaders.indexOf('No Of CV Planned Submit') === -1) {
      sheet.getRange(1, lastCol + 1).setValue('No Of CV Planned Submit')
           .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
    }
  }

  var dateVal = data.jobOpeningDate || '';
  if (dateVal) {
    var p = dateVal.split('-');
    if (p.length === 3) dateVal = p[2]+'/'+p[1]+'/'+p[0];
  }

  var uniqueId = getNextUniqueId(sheet);

  sheet.appendRow([
    new Date(),
    uniqueId,
    data.businessUnit        || '',
    data.clientOrDept        || '',
    data.accountManager      || '',
    data.accountType         || '',
    data.jobTitle            || '',
    data.skills              || '',
    data.recruitmentType     || '',
    data.positionType        || '',
    Number(data.openPositions)||0,
    data.location            || '',
    dateVal,
    data.jd                  || '',
    data.cvLink              || '',
    data.recruiterPersonName || '',
    Number(data.deadlineDays)||0,
    data.requirementStatus   || 'In Progress',
    Number(data.cvPlanned)   || 0
  ]);

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');

  // Make CV Link clickable if present (CV Link is now Column O)
  if (data.cvLink) {
    var cvColIndex = 15; // Column O (1-based)
    var cvCell = sheet.getRange(lastRow, cvColIndex);
    cvCell.setFormula('=HYPERLINK("' + data.cvLink + '","View CV")');
    cvCell.setFontColor('#1155CC');
  }

  return { success: true, message: 'Saved!', uniqueId: uniqueId };
}
