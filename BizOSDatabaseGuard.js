// ============================================================
// BizOSDatabaseGuard.js
// Canonical database schema + safe heading repair.
//
// IMPORTANT:
// - Canonical headings are always A1, B1, C1...
// - Never append a missing heading after unrelated columns.
// - Existing row data is preserved.
// - Existing columns are only shifted when required to restore
//   the canonical schema order.
// - Client_Info is a Property/Value sheet and is handled separately.
// ============================================================

var BIZOS_MASTER_SCHEMAS = {
  Users: [
    'Email','Password_Hash','Name','Phone','Role','Business_ID','Business_Name',
    'Is_Verified','Invited_By','Invitation_Status','Invitation_Code','Created_At',
    'Last_Login','Is_Active','Is_Demo','Is_Client'
  ],
  Businesses: [
    'Business_ID','Business_Name','CAC_Number','Tax_ID','Verification_Status',
    'Verification_Date','Verified_By','Owner_Email','Owner_Name','Owner_Phone',
    'Address','City','Country','Website','Subscription_Tier','Subscription_Status',
    'Workspace_ID','Created_At','Last_Active','Is_Active','Assigned_Modules'
  ],
  Clients: [
    'Client_ID','Email','Client_Name','Domain','Primary_Color','Logo_Url',
    'Custom_Domain','Tier','Status','Sheet_ID','Web_App_URL','API_Key','Created_At','Script_ID'
  ],
  Upgrade_Requests: [
    'Request_ID','Email','Name','Business_Name','Country','Business_Details','Tier',
    'Payment_ID','Amount','Currency','Status','Certificate_URL','Certificate_Name',
    'Created_At','Updated_At'
  ],
  Payments: [
    'Payment_ID','Request_ID','Email','Amount','Currency','Status','Transaction_Ref',
    'Created_At','Completed_At'
  ],
  Email_Verifications: [
    'Email','Token','Business_Name','Business_Id','Created_At','Expires_At','Is_Used'
  ],
  Password_Resets: [
    'Email','Token','Created_At','Used'
  ]
};

// Exact schemas used when a new client workspace is provisioned.
var BIZOS_CLIENT_SCHEMAS = {
  Financial_Data: ['Date','Type','Category','Subcategory','Amount','Description','Status','Created_By','Created_At'],
  Ecommerce_Data: ['Order_ID','Date','Customer_Email','Product','Quantity','Price','Status','Created_At'],
  Sales_Data: ['Deal_ID','Date','Customer_Name','Product','Amount','Stage','Created_At'],
  CRM_Data: ['Contact_ID','Name','Email','Phone','Company','Status','Created_At'],
  HR_Data: ['Employee_ID','Name','Position','Department','Salary','Status','Created_At'],
  Logistics_Data: ['Shipment_ID','Order_ID','Status','Location','Carrier','Created_At'],
  Tax_Data: ['Tax_ID','Date','Type','Amount','Status','Created_At'],
  Agro_Data: ['Record_ID','Date','Type','Activity','Quantity','Cost','Revenue','Created_At'],
  Productivity_Data: ['Record_ID','Staff','Date','Task','Hours','Created_At'],
  POS_Data: ['Transaction_ID','Date','Customer','Product','Quantity','Total','Created_At'],
  Attendance_Data: ['Record_ID','Staff','Date','Clock_In','Clock_Out','Status','Created_At'],
  Warehouse_Data: ['Item_ID','Item_Name','SKU','Quantity','Location','Status','Created_At']
};

var BIZOS_CLIENT_INFO_HEADERS = ['Property','Value'];

function normalizeHeader_(value) {
  return String(value == null ? '' : value).trim().toLowerCase();
}

function headerIndexMap_(headers) {
  var map = {};
  headers.forEach(function(header, index) {
    var key = normalizeHeader_(header);
    if (key && map[key] === undefined) map[key] = index;
  });
  return map;
}

function rowHasAnyValue_(row) {
  for (var i = 0; i < row.length; i++) {
    if (String(row[i] == null ? '' : row[i]).trim() !== '') return true;
  }
  return false;
}

function ensureNewSheetHeaders_(sheet, requiredHeaders) {
  sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
  return {created:true, repaired:false, added:requiredHeaders.slice(), shifted:false};
}

/**
 * Repairs a schema without blindly appending missing headings.
 *
 * Case 1: existing canonical headers already begin at A1 -> preserve them.
 * Case 2: our previous guard appended the complete canonical header block
 *         after an existing data block -> move the header labels to A1 and
 *         preserve the existing data by moving it down one row.
 * Case 3: some canonical headers exist at the wrong positions -> insert
 *         missing columns at their canonical positions so data stays aligned.
 *
 * The function intentionally does NOT guess when it cannot identify the
 * existing structure safely.
 */
function ensureSheetSchema_(ss, sheetName, requiredHeaders) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    return ensureNewSheetHeaders_(sheet, requiredHeaders);
  }

  if (!requiredHeaders || !requiredHeaders.length) {
    return {created:false, repaired:false, added:[], shifted:false};
  }

  var lastCol = Math.max(sheet.getLastColumn(), requiredHeaders.length);
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var normalized = headers.map(normalizeHeader_);
  var requiredNormalized = requiredHeaders.map(normalizeHeader_);
  var map = headerIndexMap_(headers);

  // Already canonical at A1 in the correct order.
  var canonicalAtA = true;
  for (var i = 0; i < requiredHeaders.length; i++) {
    if (normalized[i] !== requiredNormalized[i]) {
      canonicalAtA = false;
      break;
    }
  }

  if (canonicalAtA) {
    return {created:false, repaired:false, added:[], shifted:false};
  }

  // Detect the exact mistake made by the old guard: the complete canonical
  // header block exists as one contiguous block to the right of the data.
  for (var offset = 1; offset <= lastCol - requiredHeaders.length + 1; offset++) {
    var blockMatches = true;
    for (var j = 0; j < requiredHeaders.length; j++) {
      if (normalized[offset + j] !== requiredNormalized[j]) {
        blockMatches = false;
        break;
      }
    }
    if (!blockMatches) continue;

    // Only migrate this pattern when the cells before the block are not
    // themselves canonical headers. This prevents us from moving a legitimate
    // existing schema just because a duplicate header block exists later.
    var prefixHasCanonicalHeader = false;
    for (var p = 0; p < offset; p++) {
      if (requiredNormalized.indexOf(normalized[p]) !== -1) {
        prefixHasCanonicalHeader = true;
        break;
      }
    }
    if (prefixHasCanonicalHeader) break;

    var oldDataCols = offset;
    var oldDataRows = lastRow;
    var existing = sheet.getRange(1, 1, oldDataRows, oldDataCols).getValues();

    // Preserve the old A:prefix data by moving it down one row. This is the
    // safe repair for a table whose first data row occupied row 1.
    if (oldDataRows > 0 && rowHasAnyValue_(existing[0])) {
      sheet.insertRowsBefore(1, 1);
      sheet.getRange(2, 1, oldDataRows, oldDataCols).setValues(existing);
    }

    // Put the canonical schema at A1 and remove the duplicate appended block.
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sheet.getRange(1, offset + 1, 1, requiredHeaders.length).clearContent();

    return {
      created:false,
      repaired:true,
      added:requiredHeaders.slice(),
      shifted:true,
      mode:'moved_appended_header_block_to_A1'
    };
  }

  // Existing recognizable schema: add any missing canonical columns at their
  // correct positions. insertColumnBefore shifts the data with its heading,
  // preventing values from becoming associated with the wrong field.
  var added = [];
  for (var r = 0; r < requiredHeaders.length; r++) {
    var wanted = requiredNormalized[r];
    var currentMap = headerIndexMap_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
    if (currentMap[wanted] !== undefined) continue;

    sheet.insertColumnBefore(r + 1);
    sheet.getRange(1, r + 1).setValue(requiredHeaders[r]);
    added.push(requiredHeaders[r]);
  }

  // If there are still duplicate/wrong header labels, do not overwrite them.
  // The canonical columns are now guaranteed to exist at their A1 positions.
  return {created:false, repaired:added.length > 0, added:added, shifted:added.length > 0};
}

function ensureClientInfoSchema_(ss) {
  var sheet = ss.getSheetByName('Client_Info');
  if (!sheet) {
    sheet = ss.insertSheet('Client_Info');
    sheet.getRange(1, 1, 1, 2).setValues([BIZOS_CLIENT_INFO_HEADERS]);
    return {created:true, repaired:false};
  }

  var headers = sheet.getRange(1, 1, 1, Math.max(2, sheet.getLastColumn())).getValues()[0];
  if (normalizeHeader_(headers[0]) !== 'property' || normalizeHeader_(headers[1]) !== 'value') {
    sheet.insertRowsBefore(1, 1);
    sheet.getRange(1, 1, 1, 2).setValues([BIZOS_CLIENT_INFO_HEADERS]);
    return {created:false, repaired:true};
  }
  return {created:false, repaired:false};
}

function ensureBizOSMasterDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var report = {success:true, created:[], repaired:[], addedHeaders:[], sheets:{}};

  Object.keys(BIZOS_MASTER_SCHEMAS).forEach(function(name) {
    var result = ensureSheetSchema_(ss, name, BIZOS_MASTER_SCHEMAS[name]);
    if (result.created) report.created.push(name);
    if (result.repaired) report.repaired.push(name + (result.mode ? ': ' + result.mode : ''));
    if (result.added && result.added.length) report.addedHeaders.push(name + ': ' + result.added.join(', '));
    report.sheets[name] = result;
  });

  return report;
}

function ensureBizOSClientWorkspace(sheetId, clientId, businessId, businessName, ownerEmail) {
  var ss = SpreadsheetApp.openById(String(sheetId));

  Object.keys(BIZOS_CLIENT_SCHEMAS).forEach(function(name) {
    ensureSheetSchema_(ss, name, BIZOS_CLIENT_SCHEMAS[name]);
  });

  ensureClientInfoSchema_(ss);

  var info = ss.getSheetByName('Client_Info');
  var data = info.getDataRange().getValues();
  var headers = data[0] || [];
  var propertyCol = headers.indexOf('Property');
  var valueCol = headers.indexOf('Value');
  if (propertyCol === -1 || valueCol === -1) {
    return {success:false, sheetId:String(sheetId), message:'Client_Info schema could not be established.'};
  }

  var properties = {};
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][propertyCol] || '').trim();
    if (key) properties[key] = i + 1;
  }

  var values = {
    Client_ID: clientId || '',
    Business_ID: businessId || clientId || '',
    Business_Name: businessName || 'My Business',
    Owner_Email: String(ownerEmail || '').toLowerCase(),
    Status: 'active',
    Created_At: new Date().toISOString()
  };

  Object.keys(values).forEach(function(key) {
    if (properties[key]) {
      info.getRange(properties[key], valueCol + 1).setValue(values[key]);
    } else {
      info.appendRow([key, values[key]]);
    }
  });

  return {success:true, sheetId:String(sheetId)};
}

function ensureUpgradeDatabase_() {
  return ensureBizOSMasterDatabase();
}
