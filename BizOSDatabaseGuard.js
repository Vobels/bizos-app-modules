// ============================================================
// BizOSDatabaseGuard.js
// Keeps required master sheets and headings intact.
// Safe to run repeatedly: creates missing sheets/columns only.
// Never deletes existing data or columns.
// ============================================================

var BIZOS_MASTER_SCHEMAS = {
  Users: ['Email','Password','Name','Role','Business_ID','Is_Verified','Is_Client','Created_At'],
  Businesses: ['Business_ID','Business_Name','Owner_Email','Subscription_Tier','Status','Created_At'],
  Clients: ['Client_ID','Email','Client_Name','Domain','Primary_Color','Logo_Url','Custom_Domain','Tier','Status','Sheet_ID','Web_App_URL','API_Key','Created_At','Script_ID'],
  Upgrade_Requests: ['Request_ID','Email','Business_Name','Country','Has_Certificate','Certificate_Name','Business_Details','Tier','Status','Created_At','Updated_At'],
  Payments: ['Payment_ID','Request_ID','Email','Business_Name','Amount','Currency','Reference','Status','Created_At','Updated_At']
};

var BIZOS_CLIENT_SCHEMAS = {
  Client_Info: ['Client_ID','Business_ID','Business_Name','Owner_Email','Status','Created_At'],
  Financial_Data: ['Date','Type','Amount','Description','Created_At'],
  Ecommerce_Data: ['Date','Product','Quantity','Amount','Status','Created_At'],
  Sales_Data: ['Date','Customer','Product','Amount','Status','Created_At'],
  CRM_Data: ['Date','Customer','Contact','Status','Notes','Created_At'],
  HR_Data: ['Date','Employee','Role','Status','Notes','Created_At'],
  Logistics_Data: ['Date','Item','Quantity','Status','Notes','Created_At'],
  Tax_Data: ['Date','Tax_Type','Amount','Status','Notes','Created_At'],
  Agro_Data: ['Date','Product','Quantity','Amount','Status','Notes','Created_At'],
  Productivity_Data: ['Date','Task','Owner','Status','Notes','Created_At'],
  POS_Data: ['Date','Product','Quantity','Amount','Payment_Status','Created_At'],
  Attendance_Data: ['Date','Employee','Check_In','Check_Out','Status','Created_At'],
  Warehouse_Data: ['Date','Item','Quantity','Location','Status','Created_At']
};

function ensureSheetSchema_(ss, sheetName, requiredHeaders) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return {sheet: sheet, created: true, added: requiredHeaders.slice()};
  }

  if (requiredHeaders.length === 0) return {sheet: sheet, created: false, added: []};

  var lastCol = sheet.getLastColumn();
  var headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var added = [];
  requiredHeaders.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      headers.push(header);
      added.push(header);
    }
  });
  return {sheet: sheet, created: false, added: added};
}

function ensureBizOSMasterDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var report = {success:true, created:[], addedHeaders:[], sheets:{}};
  Object.keys(BIZOS_MASTER_SCHEMAS).forEach(function(name) {
    var r = ensureSheetSchema_(ss, name, BIZOS_MASTER_SCHEMAS[name]);
    if (r.created) report.created.push(name);
    if (r.added.length) report.addedHeaders.push(name + ': ' + r.added.join(', '));
    report.sheets[name] = {created:r.created, added:r.added};
  });
  return report;
}

function ensureBizOSClientWorkspace(sheetId, clientId, businessId, businessName, ownerEmail) {
  var ss = SpreadsheetApp.openById(String(sheetId));
  Object.keys(BIZOS_CLIENT_SCHEMAS).forEach(function(name) {
    ensureSheetSchema_(ss, name, BIZOS_CLIENT_SCHEMAS[name]);
  });
  var info = ss.getSheetByName('Client_Info');
  var data = info.getDataRange().getValues();
  var headers = data[0] || [];
  var idCol = headers.indexOf('Client_ID');
  var found = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol] || '') === String(clientId)) { found = i + 1; break; }
  }
  if (found === -1) {
    var row = headers.map(function(h) {
      if (h === 'Client_ID') return clientId || '';
      if (h === 'Business_ID') return businessId || clientId || '';
      if (h === 'Business_Name') return businessName || 'My Business';
      if (h === 'Owner_Email') return ownerEmail || '';
      if (h === 'Status') return 'active';
      if (h === 'Created_At') return new Date().toISOString();
      return '';
    });
    info.appendRow(row);
  }
  return {success:true, sheetId:String(sheetId)};
}

function ensureUpgradeDatabase_() {
  return ensureBizOSMasterDatabase();
}
