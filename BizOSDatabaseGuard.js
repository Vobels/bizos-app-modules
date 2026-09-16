// ============================================================
// BizOSDatabaseGuard.js
// Canonical database schema + safe heading repair.
// ============================================================

var BIZOS_MASTER_SCHEMAS = {
  Users: ['Email','Password_Hash','Name','Phone','Role','Business_ID','Business_Name','Is_Verified','Invited_By','Invitation_Status','Invitation_Code','Created_At','Last_Login','Is_Active','Is_Demo','Is_Client'],
  Businesses: ['Business_ID','Business_Name','CAC_Number','Tax_ID','Verification_Status','Verification_Date','Verified_By','Owner_Email','Owner_Name','Owner_Phone','Address','City','Country','Website','Subscription_Tier','Subscription_Status','Workspace_ID','Created_At','Last_Active','Is_Active','Assigned_Modules'],
  Clients: ['Client_ID','Email','Client_Name','Domain','Primary_Color','Logo_Url','Custom_Domain','Tier','Status','Sheet_ID','Web_App_URL','API_Key','Created_At','Script_ID'],
  Upgrade_Requests: ['Request_ID','Email','Workspace_Email','Name','Business_Name','Country','Business_Details','Tier','Payment_ID','Amount','Currency','Status','Certificate_URL','Certificate_Name','Created_At','Updated_At'],
  Payments: ['Payment_ID','Request_ID','Email','Workspace_Email','Amount','Currency','Status','Transaction_Ref','Created_At','Completed_At'],
  Email_Verifications: ['Email','Token','Business_Name','Business_Id','Created_At','Expires_At','Is_Used'],
  Password_Resets: ['Email','Token','Created_At','Used']
};

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

function normalizeHeader_(value) { return String(value == null ? '' : value).trim().toLowerCase(); }
function headerIndexMap_(headers) { var map={}; headers.forEach(function(h,i){var k=normalizeHeader_(h);if(k&&map[k]===undefined)map[k]=i;}); return map; }
function rowHasAnyValue_(row) { for(var i=0;i<row.length;i++){if(String(row[i]==null?'':row[i]).trim()!=='')return true;} return false; }
function ensureNewSheetHeaders_(sheet,headers){sheet.getRange(1,1,1,headers.length).setValues([headers]);sheet.getRange(1,1,1,headers.length).setFontWeight('bold');return{created:true,repaired:false,added:headers.slice(),shifted:false};}
function findContiguousHeaderBlock_(normalized,requiredNormalized){if(normalized.length<requiredNormalized.length)return-1;for(var o=1;o<=normalized.length-requiredNormalized.length;o++){var ok=true;for(var i=0;i<requiredNormalized.length;i++){if(normalized[o+i]!==requiredNormalized[i]){ok=false;break;}}if(ok)return o;}return-1;}

function ensureSheetSchema_(ss,sheetName,requiredHeaders){
  var sheet=ss.getSheetByName(sheetName);
  if(!sheet){sheet=ss.insertSheet(sheetName);return ensureNewSheetHeaders_(sheet,requiredHeaders);}
  if(!requiredHeaders||!requiredHeaders.length)return{created:false,repaired:false,added:[],shifted:false};

  var lastCol=Math.max(sheet.getLastColumn(),requiredHeaders.length,1),lastRow=Math.max(sheet.getLastRow(),1);
  var headers=sheet.getRange(1,1,1,lastCol).getValues()[0],normalized=headers.map(normalizeHeader_),requiredNormalized=requiredHeaders.map(normalizeHeader_);
  var canonicalAtA=true;
  for(var i=0;i<requiredHeaders.length;i++){if(normalized[i]!==requiredNormalized[i]){canonicalAtA=false;break;}}
  if(canonicalAtA)return{created:false,repaired:false,added:[],shifted:false};

  // Repair the known legacy error where the complete canonical block was
  // appended to the right of existing row-1 data.
  var blockOffset=findContiguousHeaderBlock_(normalized,requiredNormalized);
  if(blockOffset>0){
    var prefixHasCanonicalHeader=false;
    for(var p=0;p<blockOffset;p++){if(requiredNormalized.indexOf(normalized[p])!==-1){prefixHasCanonicalHeader=true;break;}}
    if(!prefixHasCanonicalHeader){
      var oldPrefix=sheet.getRange(1,1,lastRow,blockOffset).getValues();
      if(lastRow>0&&rowHasAnyValue_(oldPrefix[0])){sheet.insertRowsBefore(1,1);sheet.getRange(2,1,lastRow,blockOffset).setValues(oldPrefix);}
      sheet.getRange(1,1,1,requiredHeaders.length).setValues([requiredHeaders]);
      sheet.getRange(1,1,1,requiredHeaders.length).setFontWeight('bold');
      sheet.getRange(1,blockOffset+1,1,requiredHeaders.length).clearContent();
      return{created:false,repaired:true,added:requiredHeaders.slice(),shifted:true,mode:'moved_appended_header_block_to_A1'};
    }
  }

  // Recognizable table: move complete columns, not cell contents, so each
  // existing value remains attached to its original heading.
  var currentHeaders=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0],currentMap=headerIndexMap_(currentHeaders),recognizedCount=0;
  requiredNormalized.forEach(function(k){if(currentMap[k]!==undefined)recognizedCount++;});
  if(recognizedCount>=2){
    var moved=[];
    for(var r=0;r<requiredHeaders.length;r++){
      currentHeaders=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      currentMap=headerIndexMap_(currentHeaders);
      var sourceIndex=currentMap[requiredNormalized[r]];
      if(sourceIndex===undefined){sheet.insertColumnBefore(r+1);sheet.getRange(1,r+1).setValue(requiredHeaders[r]);moved.push(requiredHeaders[r]);continue;}
      if(sourceIndex!==r){sheet.moveColumns(sheet.getRange(1,sourceIndex+1,sheet.getMaxRows(),1),r+1);moved.push(requiredHeaders[r]);}
    }
    sheet.getRange(1,1,1,requiredHeaders.length).setValues([requiredHeaders]);
    sheet.getRange(1,1,1,requiredHeaders.length).setFontWeight('bold');
    return{created:false,repaired:moved.length>0,added:moved,shifted:moved.length>0,mode:'reordered_existing_schema'};
  }

  if(sheet.getLastRow()>1||rowHasAnyValue_(headers))return{created:false,repaired:false,added:[],shifted:false,mode:'unrecognized_schema_left_untouched'};
  return ensureNewSheetHeaders_(sheet,requiredHeaders);
}

function ensureClientInfoSchema_(ss){
  var sheet=ss.getSheetByName('Client_Info');
  if(!sheet){sheet=ss.insertSheet('Client_Info');sheet.getRange(1,1,1,2).setValues([BIZOS_CLIENT_INFO_HEADERS]);sheet.getRange(1,1,1,2).setFontWeight('bold');return{created:true,repaired:false};}
  var lastCol=Math.max(sheet.getLastColumn(),2),headers=sheet.getRange(1,1,1,lastCol).getValues()[0].map(normalizeHeader_);
  if(headers[0]==='property'&&headers[1]==='value')return{created:false,repaired:false};
  for(var offset=1;offset<=headers.length-2;offset++){
    if(headers[offset]==='property'&&headers[offset+1]==='value'){
      var rows=Math.max(sheet.getLastRow(),1),prefix=sheet.getRange(1,1,rows,offset).getValues();
      if(rowHasAnyValue_(prefix[0])){sheet.insertRowsBefore(1,1);sheet.getRange(2,1,rows,offset).setValues(prefix);}
      sheet.getRange(1,1,1,2).setValues([BIZOS_CLIENT_INFO_HEADERS]);sheet.getRange(1,1,1,2).setFontWeight('bold');sheet.getRange(1,offset+1,1,2).clearContent();return{created:false,repaired:true,mode:'moved_appended_header_block_to_A1'};
    }
  }
  var propertyIndex=headers.indexOf('property'),valueIndex=headers.indexOf('value');
  if(propertyIndex!==-1&&valueIndex!==-1){if(propertyIndex!==0)sheet.moveColumns(sheet.getRange(1,propertyIndex+1,sheet.getMaxRows(),1),1);headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(normalizeHeader_);valueIndex=headers.indexOf('value');if(valueIndex!==1)sheet.moveColumns(sheet.getRange(1,valueIndex+1,sheet.getMaxRows(),1),2);sheet.getRange(1,1,1,2).setValues([BIZOS_CLIENT_INFO_HEADERS]);sheet.getRange(1,1,1,2).setFontWeight('bold');return{created:false,repaired:true,mode:'reordered_property_value'};}
  if(sheet.getLastRow()<=1&&!rowHasAnyValue_(sheet.getRange(1,1,1,lastCol).getValues()[0])){sheet.getRange(1,1,1,2).setValues([BIZOS_CLIENT_INFO_HEADERS]);return{created:false,repaired:true,mode:'initialized_empty_sheet'};}
  return{created:false,repaired:false,mode:'unrecognized_schema_left_untouched'};
}

function ensureBizOSMasterDatabase(){var ss=SpreadsheetApp.getActiveSpreadsheet(),report={success:true,created:[],repaired:[],addedHeaders:[],sheets:{}};Object.keys(BIZOS_MASTER_SCHEMAS).forEach(function(name){var result=ensureSheetSchema_(ss,name,BIZOS_MASTER_SCHEMAS[name]);if(result.created)report.created.push(name);if(result.repaired)report.repaired.push(name+(result.mode?': '+result.mode:''));if(result.added&&result.added.length)report.addedHeaders.push(name+': '+result.added.join(', '));report.sheets[name]=result;});return report;}

function ensureBizOSClientWorkspace(sheetId,clientId,businessId,businessName,ownerEmail){
  var ss=SpreadsheetApp.openById(String(sheetId));
  Object.keys(BIZOS_CLIENT_SCHEMAS).forEach(function(name){ensureSheetSchema_(ss,name,BIZOS_CLIENT_SCHEMAS[name]);});
  ensureClientInfoSchema_(ss);
  var info=ss.getSheetByName('Client_Info'),data=info.getDataRange().getValues(),headers=data[0]||[],propertyCol=headers.indexOf('Property'),valueCol=headers.indexOf('Value');
  if(propertyCol===-1||valueCol===-1)return{success:false,sheetId:String(sheetId),message:'Client_Info schema could not be established.'};
  var properties={};for(var i=1;i<data.length;i++){var key=String(data[i][propertyCol]||'').trim();if(key)properties[key]=i+1;}
  var values={Client_ID:clientId||'',Business_ID:businessId||clientId||'',Business_Name:businessName||'My Business',Owner_Email:String(ownerEmail||'').toLowerCase(),Status:'active',Created_At:new Date().toISOString()};
  Object.keys(values).forEach(function(key){if(properties[key])info.getRange(properties[key],valueCol+1).setValue(values[key]);else info.appendRow([key,values[key]]);});
  return{success:true,sheetId:String(sheetId)};
}

function ensureUpgradeDatabase_(){return ensureBizOSMasterDatabase();}
