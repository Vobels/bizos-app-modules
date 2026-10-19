// ============================================================
// ClientAuth.js - CLIENT DEPLOYMENT AUTH + DATA ACCESS
// ============================================================

function getClientDeploymentConfig() {
  try {
    if (typeof CLIENT_CONFIG !== 'undefined') return {success:true,clientId:String(CLIENT_CONFIG.clientId||''),clientName:CLIENT_CONFIG.clientName||'',sheetId:String(CLIENT_CONFIG.sheetId||''),businessId:String(CLIENT_CONFIG.businessId||CLIENT_CONFIG.clientId||''),masterApiUrl:CLIENT_CONFIG.masterApiUrl||''};
  } catch(e) {}
  return {success:false,message:'Client deployment configuration is unavailable.'};
}

function clientMasterModuleNames_(){return ['Finance','Sales','Ecommerce','CRM','HR','Logistics','Tax','Agro','Productivity','POS','Attendance','Warehouse'];}

// The request's clientId/sheetId/businessId are treated only as deployment
// binding hints. The authoritative paid-client identity and workspace always
// come from the master Clients registry after the account is authenticated.
function clientLogin(requestClientId,email,password,requestSheetId,requestBusinessId,businessName) {
  try {
    if(!email||!password)return{success:false,message:'Email and password are required',code:'INVALID_INPUT'};

    var authResult=null;
    try{if(typeof authenticateUser==='function')authResult=authenticateUser(email,password,null);}catch(e){console.log('Local authentication unavailable:',e.message);}
    if(!authResult||!authResult.success)authResult=authenticateAgainstMaster(email,password);
    if(!authResult||!authResult.success)return{success:false,message:authResult&&authResult.message?authResult.message:'Invalid email or password',code:'AUTH_FAILED'};

    // Re-read the authoritative client record by the authenticated account
    // email. Never select a workspace from browser/request-supplied IDs.
    var registryResult = getActiveClientByEmail_(email);
    if(!registryResult||!registryResult.success){
      return {success:false,message:registryResult&&registryResult.message?registryResult.message:'No active paid client is registered for this account.',code:registryResult&&registryResult.code?registryResult.code:'CLIENT_NOT_FOUND'};
    }

    var client=registryResult.client;
    var normalizedRequestClientId=String(requestClientId||'').trim();
    if(normalizedRequestClientId&&normalizedRequestClientId!==String(client.clientId)){
      return {success:false,message:'This deployment is not authorized for the authenticated client.',code:'CLIENT_BINDING_MISMATCH'};
    }

    if(client.status&&String(client.status).toLowerCase()!=='active')return{success:false,message:'This client workspace is inactive.',code:'CLIENT_INACTIVE'};
    if(!client.sheetId||!client.businessId)return{success:false,message:'Client workspace is not fully configured.',code:'INVALID_WORKSPACE'};

    var user=authResult.user||{email:String(email).toLowerCase(),name:String(email).split('@')[0],role:'owner',isClient:true};
    var role=String(user.role||'owner').toLowerCase();
    var tier=String(client.tier||user.subscriptionTier||'sovereign').toLowerCase();
    var assignedModules=[];
    try{
      if(typeof clientAssignedModulesV11_==='function'){
        var teamSheet=ensureClientTeamSheetV11_();
        var teamRows=clientTeamRowsV11_(teamSheet);
        var team=teamRows.find(function(r){return String(r.Email||'').toLowerCase()===String(email||'').toLowerCase();});
        if(team)assignedModules=clientAssignedModulesV11_(team);
      }
    }catch(ignore){}
    var accessibleModules=[];
    try{
      if(typeof clientModulesForEntitlementV12_==='function'){
        accessibleModules=clientModulesForEntitlementV12_(role,tier,assignedModules);
      }else if(role==='owner'||role==='admin'){
        accessibleModules=clientMasterModuleNames_();
      }else{
        accessibleModules=assignedModules;
      }
    }catch(ignore2){
      accessibleModules=(role==='owner'||role==='admin')?clientMasterModuleNames_():assignedModules;
    }
    user={
      email:String(user.email||email).toLowerCase(),
      name:String(user.name||email.split('@')[0]),
      role:role,
      subscriptionTier:tier,
      workspaceId:String(client.sheetId),
      businessId:String(client.businessId),
      businessName:client.clientName||'My Business',
      isDemo:false,
      isClient:true,
      accessibleModules:accessibleModules
    };
    var now=Date.now(),sessionId='CLIENT_SESS_'+Utilities.getUuid(),session={
      email:String(client.email||email).toLowerCase(),
      clientId:String(client.clientId),
      sheetId:String(client.sheetId),
      businessId:String(client.businessId),
      businessName:client.clientName||'My Business',
      created:now,
      expiresAt:now+21600000,
      isClient:true,
      user:user,
      apiKey:''
    };
    CacheService.getScriptCache().put(sessionId,JSON.stringify(session),21600);

    return{success:true,sessionId:sessionId,message:'Login successful',clientName:client.clientName||'My Business',primaryColor:client.primaryColor||'#5D2A86',logoUrl:client.logoUrl||'',user:{email:user.email,name:user.name,role:user.role,subscriptionTier:user.subscriptionTier,clientId:String(client.clientId),businessId:String(client.businessId),accessibleModules:user.accessibleModules},code:'SUCCESS'};
  }catch(error){console.error('clientLogin error:',error);return{success:false,message:'Login error: '+error.message,code:'LOGIN_ERROR'};}
}

function authenticateAgainstMaster(email,password){
  try{
    var masterUrl='';
    if(typeof getMasterApiUrl==='function')masterUrl=getMasterApiUrl();
    if(!masterUrl&&typeof CLIENT_CONFIG!=='undefined')masterUrl=CLIENT_CONFIG.masterApiUrl||'';
    if(!masterUrl)return{success:false,message:'Master authentication service is not configured.'};

    var response=UrlFetchApp.fetch(masterUrl,{method:'post',contentType:'application/json',payload:JSON.stringify({action:'LOGIN',payload:{email:email,password:password}}),muteHttpExceptions:true});
    var code=response.getResponseCode();
    var text=response.getContentText();
    if(code!==200)return{success:false,message:'Authentication service returned HTTP '+code};
    return JSON.parse(text)||{success:false,message:'Authentication service returned no result.'};
  }catch(error){return{success:false,message:'Unable to contact authentication service: '+error.message};}
}

function getClientFromWorkspace(sheetId,expectedClientId){try{var ss=SpreadsheetApp.openById(sheetId),info=ss.getSheetByName('Client_Info');if(!info)return null;var rows=info.getDataRange().getValues(),values={};for(var i=1;i<rows.length;i++)if(rows[i][0])values[String(rows[i][0])]=rows[i][1];if(expectedClientId&&String(values.Client_ID||'')!==String(expectedClientId))return null;return{clientId:values.Client_ID||expectedClientId,clientName:values.Business_Name||'My Business',businessName:values.Business_Name||'My Business',tier:values.Tier||'sovereign',status:values.Status||'active',sheetId:String(sheetId),primaryColor:values.Primary_Color||'#5D2A86',logoUrl:values.Logo_Url||'',businessId:values.Business_ID||expectedClientId,apiKey:values.API_Key||''};}catch(error){return null;}}

function validateClientSession(sessionId){if(!sessionId)return null;var cache=CacheService.getScriptCache();try{var raw=cache.get(sessionId);if(!raw)return null;var session=JSON.parse(raw),now=Date.now();if(!session.created||!session.clientId||!session.sheetId||!session.businessId||(session.expiresAt&&now>=Number(session.expiresAt))||now-Number(session.created)>21600000){cache.remove(sessionId);return null;}var expected=getClientDeploymentConfig();if(!expected.success||String(session.clientId)!==String(expected.clientId)){cache.remove(sessionId);return null;}var client=getClientById(session.clientId);if(!client)client=getClientFromWorkspace(session.sheetId,session.clientId);if(!client||String(client.clientId)!==String(session.clientId)||String(client.sheetId)!==String(session.sheetId)||String(client.businessId)!==String(session.businessId)||(client.status&&String(client.status).toLowerCase()!=='active')){cache.remove(sessionId);return null;}return session;}catch(error){try{cache.remove(sessionId);}catch(ignore){}return null;}}
function getClientUserFromSession(sessionId){var session=validateClientSession(sessionId);if(!session)return null;var client=getClientById(session.clientId);if(!client)client=getClientFromWorkspace(session.sheetId,session.clientId);if(!client)return null;return{email:session.email,name:session.user&&session.user.name?session.user.name:session.email,role:session.user&&session.user.role?session.user.role:'owner',clientId:session.clientId,businessId:session.businessId,clientName:client.clientName||client.businessName||session.businessName,primaryColor:client.primaryColor||'#5D2A86',logoUrl:client.logoUrl||'',sheetId:session.sheetId,isClient:true,apiKey:client.apiKey||''};}
function clientLogout(sessionId){try{if(sessionId)CacheService.getScriptCache().remove(sessionId);return{success:true,message:'Logged out successfully',code:'LOGOUT_SUCCESS'};}catch(error){return{success:false,message:error.message,code:'LOGOUT_ERROR'};}}
function getClientDashboardData(sessionId){try{var session=validateClientSession(sessionId);if(!session)return{success:false,message:'Session expired. Please login again.',code:'SESSION_EXPIRED'};var client=getClientById(session.clientId);if(!client)client=getClientFromWorkspace(session.sheetId,session.clientId);if(!client)return{success:false,message:'Client workspace could not be verified.',code:'CLIENT_NOT_FOUND'};var ss=SpreadsheetApp.openById(session.sheetId),financeSheet=ss.getSheetByName('Financial_Data'),revenue=0,expenses=0,recentTransactions=[];if(financeSheet){var data=financeSheet.getDataRange().getValues();if(data.length>1){var headers=data[0],typeCol=headers.indexOf('Type'),amountCol=headers.indexOf('Amount'),dateCol=headers.indexOf('Date'),descCol=headers.indexOf('Description');for(var i=1;i<data.length;i++){var amount=parseFloat(data[i][amountCol])||0,type=String(data[i][typeCol]||'').toLowerCase();if(type==='revenue'||type==='income')revenue+=Math.abs(amount);else if(type==='expense')expenses+=Math.abs(amount);}if(dateCol!==-1)recentTransactions=data.slice(1).filter(function(r){return r[dateCol];}).sort(function(a,b){return new Date(b[dateCol])-new Date(a[dateCol]);}).slice(0,5).map(function(r){return{date:r[dateCol],description:descCol!==-1?r[descCol]:'',amount:amountCol!==-1?r[amountCol]:0,type:typeCol!==-1?r[typeCol]:''};});}}var profit=revenue-expenses,margin=revenue>0?(profit/revenue)*100:0,names=['Finance','Ecommerce','Sales','CRM','HR','Logistics','Tax','Agro','Productivity','POS','Attendance','Warehouse'],modules=names.map(function(name){var sheet=ss.getSheetByName(name+'_Data')||ss.getSheetByName(name);return{name:name,label:getModuleLabel(name),icon:getModuleIcon(name),count:sheet&&sheet.getLastRow()>1?sheet.getLastRow()-1:0,isAccessible:true};});return{success:true,dashboard:{kpis:{revenue:Math.round(revenue*100)/100,expenses:Math.round(expenses*100)/100,profit:Math.round(profit*100)/100,margin:Math.round(margin*10)/10},modules:modules,recentTransactions:recentTransactions,clientName:client.clientName||client.businessName||session.businessName,primaryColor:client.primaryColor||'#5D2A86',logoUrl:client.logoUrl||'',sheetId:session.sheetId},code:'SUCCESS'};}catch(error){return{success:false,message:error.message,code:'DASHBOARD_ERROR'};}}
function getClientModuleData(moduleName,sessionId){try{var session=validateClientSession(sessionId);if(!session)return{success:false,message:'Session expired',code:'SESSION_EXPIRED'};var ss=SpreadsheetApp.openById(session.sheetId),sheet=ss.getSheetByName(moduleName+'_Data')||ss.getSheetByName(moduleName);if(!sheet)return{success:true,data:[],code:'SHEET_NOT_FOUND'};var values=sheet.getDataRange().getValues();if(values.length<2)return{success:true,data:[],code:'NO_DATA'};var headers=values[0],records=values.slice(1).map(function(row){var obj={};headers.forEach(function(h,i){if(h&&String(h).trim())obj[String(h).trim()]=row[i]!==undefined?row[i]:'';});return obj;});return{success:true,data:records,code:'SUCCESS'};}catch(error){return{success:false,message:error.message,code:'FETCH_ERROR'};}}
function saveClientModuleRecord(moduleName, record, sessionId) {
  try {
    var session = validateClientSession(sessionId);
    if (!session) return {success:false,message:'Session expired. Please sign in again.',code:'SESSION_EXPIRED'};
    if (!getClientModuleSheetNameV12_(moduleName)) return {success:false,message:'Module not found.',code:'MODULE_NOT_FOUND'};
    if (getClientAccessibleModulesV12_(session).indexOf(String(moduleName)) === -1) return {success:false,message:'You do not have access to this module.',code:'ACCESS_DENIED'};
    var role = String(session.user && session.user.role || 'staff').toLowerCase();
    if (role !== 'owner' && role !== 'admin' && role !== 'staff') return {success:false,message:'This module is read-only for your account.',code:'READ_ONLY'};
    var isDemo = session.user && (session.user.isDemo === true || String(session.user.isDemo || '').toUpperCase() === 'YES');
    if (isDemo) return {success:false,message:'Demo accounts are read-only.',code:'READ_ONLY'};
    var ss = SpreadsheetApp.openById(session.sheetId), sheet = getClientModuleSheetV12_(ss, moduleName);
    if (!sheet) return {success:false,message:'The existing sheet for this app could not be found. No data was changed.',code:'SHEET_NOT_FOUND'};
    var width = sheet.getLastColumn();
    if (!width || sheet.getLastRow() < 1) return {success:false,message:'The existing app sheet has no headers. No data was changed.',code:'INVALID_HEADERS'};
    var headers = sheet.getRange(1,1,1,width).getValues()[0];
    if (!headers.some(function(h){return String(h || '').trim() !== '';})) return {success:false,message:'The existing app sheet has no headers. No data was changed.',code:'INVALID_HEADERS'};
    record = record && typeof record === 'object' ? record : {};
    var idColumn = getClientModuleIdColumnV12_(moduleName), now = new Date().toISOString();
    var idValue = record[idColumn] || (idColumn ? getClientModuleRecordIdV12_(moduleName) : '');
    var row = headers.map(function(header) {
      var key = String(header || '').trim();
      if (!key) return '';
      if (key === idColumn) return idValue;
      if (key === 'Created_By') return String(session.email || '');
      if (key === 'Created_At') return now;
      return Object.prototype.hasOwnProperty.call(record,key) ? record[key] : '';
    });
    sheet.appendRow(row);
    var savedRecord = {};
    headers.forEach(function(header, index) { var key = String(header || '').trim(); if (key) savedRecord[key] = row[index]; });
    auditRecordChange_(ss, sessionId, moduleName, 'CREATE', idValue, {}, savedRecord, session);
    return {success:true,message:'Record saved successfully.',code:'SAVE_SUCCESS',recordId:idValue};
  } catch (error) {
    console.error('saveClientModuleRecord error:', error);
    return {success:false,message:'We could not save this record. No existing records were deleted.',code:'SAVE_ERROR'};
  }
}

function getDefaultDashboardData(clientName){return{kpis:{revenue:0,expenses:0,profit:0,margin:0},modules:[{name:'Finance',label:getModuleLabel('Finance'),icon:getModuleIcon('Finance'),count:0},{name:'Ecommerce',label:getModuleLabel('Ecommerce'),icon:getModuleIcon('Ecommerce'),count:0}],recentTransactions:[],clientName:clientName||'My Business',primaryColor:'#5D2A86',logoUrl:''};}
function getModuleLabel(n){var x={Finance:'Financial Management',Ecommerce:'E-commerce',Sales:'Sales Pipeline',CRM:'Customer Relations',HR:'Human Resources',Logistics:'Logistics',Tax:'Tax Compliance',Agro:'Agriculture',Productivity:'Productivity Suite',POS:'Point of Sale',Attendance:'Staff Attendance',Warehouse:'Warehouse Management'};return x[n]||n;}
function getModuleIcon(n){var x={Finance:'fa-calculator',Ecommerce:'fa-cart-shopping',Sales:'fa-chart-line',CRM:'fa-users',HR:'fa-briefcase',Logistics:'fa-truck',Tax:'fa-file-invoice',Agro:'fa-seedling',Productivity:'fa-clock',POS:'fa-cash-register',Attendance:'fa-fingerprint',Warehouse:'fa-warehouse'};return x[n]||'fa-folder';}

// V12 client-scoped module/records readers. These must use the client session
// and client workspace, never the Master getUserFromSession()/getWorkspaceFile()
// path, because CLIENT_SESS_* tokens are not Master sessions.
function getClientModuleSheetNameV12_(moduleName) {
  var names = {Finance:'Financial_Data',Sales:'Sales_Data',Ecommerce:'Ecommerce_Data',CRM:'CRM_Data',HR:'HR_Data',Logistics:'Logistics_Data',Tax:'Tax_Data',Agro:'Agro_Data',Productivity:'Productivity_Data',POS:'POS_Data',Attendance:'Attendance_Data',Warehouse:'Warehouse_Data'};
  return names[String(moduleName || '')] || '';
}

function getClientModuleSheetV12_(spreadsheet, moduleName) {
  if (!spreadsheet) throw new Error('Client workspace is unavailable.');
  var canonical = getClientModuleSheetNameV12_(moduleName);
  if (!canonical) return null;
  var candidates = [canonical, String(moduleName || '') + '_Data', String(moduleName || '')];
  if (String(moduleName) === 'Finance') candidates.push('Finance_Data');
  var unique = [];
  candidates.forEach(function(name) { if (name && unique.indexOf(name) === -1) unique.push(name); });
  var existing = unique.map(function(name) { return spreadsheet.getSheetByName(name); }).filter(Boolean);
  if (!existing.length) return null;
  if (existing.length === 1) return existing[0];
  var populated = existing.filter(function(sheet) { return sheet.getLastRow() > 1; });
  if (populated.length === 1) return populated[0];
  if (populated.length > 1) throw new Error('Multiple populated sheets exist for ' + moduleName + '. No data was changed; resolve the sheet-name conflict first.');
  for (var i = 0; i < existing.length; i++) if (existing[i].getName() === canonical) return existing[i];
  return existing[0];
}

function getClientModuleIdColumnV12_(moduleName) {
  var ids = {Finance:'Transaction_ID',Sales:'Deal_ID',Ecommerce:'Order_ID',CRM:'Contact_ID',HR:'Employee_ID',Logistics:'Shipment_ID',Tax:'Tax_ID',Agro:'Record_ID',Productivity:'Record_ID',POS:'Transaction_ID',Attendance:'Record_ID',Warehouse:'Item_ID'};
  return ids[String(moduleName || '')] || '';
}

function getClientModuleRecordIdV12_(moduleName) {
  var prefixes = {Finance:'TXN',Sales:'DEAL',Ecommerce:'ORD',CRM:'CRM',HR:'EMP',Logistics:'SHIP',Tax:'TAX',Agro:'AGRO',Productivity:'TASK',POS:'POS',Attendance:'ATT',Warehouse:'ITEM'};
  return (prefixes[String(moduleName || '')] || 'REC') + '-' + Utilities.getUuid().substring(0,8).toUpperCase();
}

function getClientModuleDataWithSync(moduleName, sessionId) {
  try {
    var session = validateClientSession(sessionId);
    if (!session) return {success:false,message:'Session expired. Please sign in again.',code:'SESSION_EXPIRED'};
    if (!getClientModuleSheetNameV12_(moduleName)) return {success:false,message:'Module not found.',code:'MODULE_NOT_FOUND'};
    if (getClientAccessibleModulesV12_(session).indexOf(String(moduleName)) === -1) return {success:false,message:'You do not have access to this module.',code:'ACCESS_DENIED'};
    var ss = SpreadsheetApp.openById(session.sheetId), sheet = getClientModuleSheetV12_(ss,moduleName);
    if (!sheet) return {success:false,headers:[],records:[],recordCount:0,message:'The existing sheet for this app could not be found. No data was changed.',code:'SHEET_NOT_FOUND'};
    var lastRow = sheet.getLastRow(), width = sheet.getLastColumn();
    if (!lastRow || !width) return {success:false,headers:[],records:[],recordCount:0,message:'The existing app sheet has no readable headers.',code:'INVALID_HEADERS'};
    var headers = sheet.getRange(1,1,1,width).getValues()[0];
    if (lastRow < 2) return {success:true,headers:headers,records:[],recordCount:0,financialSummary:{totalRevenue:0,totalExpenses:0},code:'NO_DATA',sheetName:sheet.getName()};
    var startRow = Math.max(2,lastRow-999), values = sheet.getRange(startRow,1,lastRow-startRow+1,width).getValues();
    var records = values.map(function(row,offset) { var item={}; headers.forEach(function(h,i){if(h!==''&&h!=null)item[String(h)]=row[i];}); item._row=startRow+offset; return item; });
    return {success:true,headers:headers,records:records,recordCount:records.length,financialSummary:{totalRevenue:0,totalExpenses:0},code:'SUCCESS',sheetName:sheet.getName()};
  } catch (error) {
    console.error('getClientModuleDataWithSync error:',error);
    return {success:false,message:String(error&&error.message||'We could not load this app’s records. Please refresh and try again.'),code:'FETCH_ERROR'};
  }
}

function updateClientModuleRecord(moduleName, record, recordId, sessionId) {
  try {
    var session=validateClientSession(sessionId);
    if(!session)return{success:false,message:'Session expired. Please sign in again.',code:'SESSION_EXPIRED'};
    if(!getClientModuleSheetNameV12_(moduleName))return{success:false,message:'Module not found.',code:'MODULE_NOT_FOUND'};
    if(getClientAccessibleModulesV12_(session).indexOf(String(moduleName))===-1)return{success:false,message:'You do not have access to this module.',code:'ACCESS_DENIED'};
    var role=String(session.user&&session.user.role||'staff').toLowerCase();
    if(role!=='owner'&&role!=='admin'&&role!=='staff')return{success:false,message:'This module is read-only for your account.',code:'READ_ONLY'};
    var isDemo=session.user&&(session.user.isDemo===true||String(session.user.isDemo||'').toUpperCase()==='YES');
    if(isDemo)return{success:false,message:'Demo accounts are read-only.',code:'READ_ONLY'};
    var idColumn=getClientModuleIdColumnV12_(moduleName);
    if(!idColumn||recordId==null||String(recordId)==='')return{success:false,message:'Record ID is missing.',code:'RECORD_ID_MISSING'};
    var ss=SpreadsheetApp.openById(session.sheetId),sheet=getClientModuleSheetV12_(ss,moduleName);
    if(!sheet)return{success:false,message:'The existing sheet for this app could not be found. No data was changed.',code:'SHEET_NOT_FOUND'};
    var lastRow=sheet.getLastRow(),width=sheet.getLastColumn();
    if(lastRow<2||!width)return{success:false,message:'The record could not be found.',code:'RECORD_NOT_FOUND'};
    var headers=sheet.getRange(1,1,1,width).getValues()[0],idIndex=headers.indexOf(idColumn);
    if(idIndex<0)return{success:false,message:'The existing sheet is missing its record ID column.',code:'ID_COLUMN_MISSING'};
    var values=sheet.getRange(2,1,lastRow-1,width).getValues(),targetRow=-1;
    // Prefer the physical row captured by the editor. This avoids editing a
    // different record when legacy sheets contain duplicate IDs.
    var requestedRow=recordId&&typeof recordId==='object'?Number(recordId.rowNumber||recordId._row||0):0;
    var expectedId=recordId&&typeof recordId==='object'?String(recordId.recordId||recordId.id||''):String(recordId==null?'':recordId);
    if(Number.isInteger(requestedRow)&&requestedRow>=2&&requestedRow<=lastRow){
      var selected=sheet.getRange(requestedRow,1,1,width).getValues()[0];
      if(expectedId&&String(selected[idIndex])!==expectedId)return{success:false,message:'This record changed while you were editing. Refresh the list and try again.',code:'RECORD_CHANGED'};
      targetRow=requestedRow;
    } else {
      for(var i=0;i<values.length;i++)if(String(values[i][idIndex])===expectedId){targetRow=i+2;break;}
    }
    if(targetRow<0)return{success:false,message:'The record could not be found.',code:'RECORD_NOT_FOUND'};
    var current=sheet.getRange(targetRow,1,1,width).getValues()[0],incoming=record&&typeof record==='object'?record:{};
    var beforeRecord = {};
    headers.forEach(function(header,i){var key=String(header||'').trim();if(key)beforeRecord[key]=current[i];});
    headers.forEach(function(header,i){var key=String(header||'').trim();if(key&&key!==idColumn&&key!=='Created_At'&&key!=='Created_By'&&key!=='Updated_At'&&key!=='Updated_By'&&key!=='_row'&&Object.prototype.hasOwnProperty.call(incoming,key))current[i]=incoming[key];});
    sheet.getRange(targetRow,1,1,width).setValues([current]);
    var afterRecord = {};
    headers.forEach(function(header,i){var key=String(header||'').trim();if(key)afterRecord[key]=current[i];});
    auditRecordChange_(ss, sessionId, moduleName, 'UPDATE', beforeRecord[idColumn] || expectedId || '', beforeRecord, afterRecord, session);
    return{success:true,message:'Record updated successfully.',code:'UPDATE_SUCCESS'};
  }catch(error){console.error('updateClientModuleRecord error:',error);return{success:false,message:'We could not update this record. No existing records were deleted.',code:'UPDATE_ERROR'};}
}

function deleteClientModuleRecord(moduleName, dataIndex, sessionId) {
  try {
    var session=validateClientSession(sessionId);
    if(!session)return{success:false,message:'Session expired. Please sign in again.',code:'SESSION_EXPIRED'};
    if(!getClientModuleSheetNameV12_(moduleName))return{success:false,message:'Module not found.',code:'MODULE_NOT_FOUND'};
    if(getClientAccessibleModulesV12_(session).indexOf(String(moduleName))===-1)return{success:false,message:'You do not have access to this module.',code:'ACCESS_DENIED'};
    var role=String(session.user&&session.user.role||'staff').toLowerCase();
    var isDemo=session.user&&(session.user.isDemo===true||String(session.user.isDemo||'').toUpperCase()==='YES');
    if(isDemo)return{success:false,message:'Demo accounts are read-only.',code:'READ_ONLY'};
    if(role!=='owner'&&role!=='admin'&&role!=='staff')return{success:false,message:'You do not have permission to delete records.',code:'FORBIDDEN'};
    var idColumn=getClientModuleIdColumnV12_(moduleName);
    if(!idColumn)return{success:false,message:'This app has no supported record ID field.',code:'ID_COLUMN_MISSING'};
    var index=Number(dataIndex);
    if(!Number.isInteger(index)||index<0)return{success:false,message:'Invalid record position.',code:'INVALID_ROW'};
    var ss=SpreadsheetApp.openById(session.sheetId),sheet=getClientModuleSheetV12_(ss,moduleName);
    if(!sheet)return{success:false,message:'The existing sheet for this app could not be found. No data was changed.',code:'SHEET_NOT_FOUND'};
    var rowNumber=index+2;
    if(rowNumber>sheet.getLastRow())return{success:false,message:'The record could not be found.',code:'RECORD_NOT_FOUND'};
    var width=sheet.getLastColumn(),headers=sheet.getRange(1,1,1,width).getValues()[0],row=sheet.getRange(rowNumber,1,1,width).getValues()[0],beforeRecord={};
    var idIndex=headers.indexOf(idColumn);
    if(idIndex<0)return{success:false,message:'The existing sheet is missing its record ID column.',code:'ID_COLUMN_MISSING'};
    headers.forEach(function(header,i){var key=String(header||'').trim();if(key)beforeRecord[key]=row[i];});
    var recordId=beforeRecord[idColumn]||'';
    auditRecordChange_(ss, sessionId, moduleName, 'DELETE', recordId, beforeRecord, {}, session);
    sheet.deleteRow(rowNumber);
    return{success:true,message:'Record deleted successfully.',code:'DELETE_SUCCESS'};
  }catch(error){console.error('deleteClientModuleRecord error:',error);return{success:false,message:'We could not delete this record.',code:'DELETE_ERROR'};}
}

function updateClientModuleRecordField(moduleName, dataIndex, columnName, newValue, sessionId) {
  try {
    var session=validateClientSession(sessionId);
    if(!session)return{success:false,message:'Session expired. Please sign in again.',code:'SESSION_EXPIRED'};
    if(!getClientModuleSheetNameV12_(moduleName))return{success:false,message:'Module not found.',code:'MODULE_NOT_FOUND'};
    if(getClientAccessibleModulesV12_(session).indexOf(String(moduleName))===-1)return{success:false,message:'You do not have access to this module.',code:'ACCESS_DENIED'};
    var role=String(session.user&&session.user.role||'owner').toLowerCase();
    if(role!=='owner'&&role!=='admin')return{success:false,message:'This module is read-only for staff accounts.',code:'READ_ONLY'};
    var index=Number(dataIndex);
    if(!Number.isInteger(index)||index<0)return{success:false,message:'Invalid record position.',code:'INVALID_ROW'};
    var ss=SpreadsheetApp.openById(session.sheetId),sheet=getClientModuleSheetV12_(ss,moduleName);
    if(!sheet)return{success:false,message:'The existing sheet for this app could not be found. No data was changed.',code:'SHEET_NOT_FOUND'};
    var lastRow=sheet.getLastRow(),width=sheet.getLastColumn();
    if(lastRow<2||!width)return{success:false,message:'The record could not be found.',code:'RECORD_NOT_FOUND'};
    var headers=sheet.getRange(1,1,1,width).getValues()[0],columnIndex=headers.indexOf(String(columnName||''));
    if(columnIndex<0)return{success:false,message:'The requested field does not exist in this sheet.',code:'FIELD_NOT_FOUND'};
    var startRow=Math.max(2,lastRow-999),rowNumber=startRow+index;
    if(rowNumber>lastRow)return{success:false,message:'The record could not be found.',code:'RECORD_NOT_FOUND'};
    var rowValues=sheet.getRange(rowNumber,1,1,width).getValues()[0],beforeRecord={};
    headers.forEach(function(header,i){var key=String(header||'').trim();if(key)beforeRecord[key]=rowValues[i];});
    var recordIdColumn=getClientModuleIdColumnV12_(moduleName),recordId=recordIdColumn&&headers.indexOf(recordIdColumn)>=0?rowValues[headers.indexOf(recordIdColumn)]:'';
    sheet.getRange(rowNumber,columnIndex+1).setValue(newValue);
    var afterValues=sheet.getRange(rowNumber,1,1,width).getValues()[0],afterRecord={};
    headers.forEach(function(header,i){var key=String(header||'').trim();if(key)afterRecord[key]=afterValues[i];});
    auditRecordChange_(ss, sessionId, moduleName, 'UPDATE', recordId, beforeRecord, afterRecord, session);
    return{success:true,message:'Record updated successfully.',code:'UPDATE_SUCCESS'};
  }catch(error){console.error('updateClientModuleRecordField error:',error);return{success:false,message:'We could not update this field.',code:'UPDATE_ERROR'};}
}

function getClientRecordsDirectory(sessionId) {
  try {
    var session = validateClientSession(sessionId);
    if (!session) return {success:false,message:'Session expired. Please sign in again.',code:'SESSION_EXPIRED'};
    var modules = getClientAccessibleModulesV12_(session);
    var ss = SpreadsheetApp.openById(session.sheetId), result = [];
    modules.forEach(function(moduleName) {
      var sheetName = getClientModuleSheetNameV12_(moduleName);
      if (!sheetName) return;
      var sheet = ss.getSheetByName(sheetName) || ss.getSheetByName(moduleName);
      result.push({key:moduleName,name:getModuleLabel(moduleName),module:moduleName,count:sheet && sheet.getLastRow()>1 ? sheet.getLastRow()-1 : 0,source:'module'});
    });
    if (modules.indexOf('Ecommerce') !== -1) {
      [['Products','BusinessCenter_Products'],['Customers','BusinessCenter_Customers'],['Sales','BusinessCenter_Sales'],['Sale Items','BusinessCenter_Sale_Items'],['Stock Movements','BusinessCenter_Stock_Movements'],['Business Center Finance','BusinessCenter_Finance_Transactions']].forEach(function(item) {
        var sheet = ss.getSheetByName(item[1]);
        result.push({key:'BC:'+item[1],name:item[0],module:'Business Center',count:sheet && sheet.getLastRow()>1 ? sheet.getLastRow()-1 : 0,source:'business_center'});
      });
    }
    return {success:true,records:result};
  } catch (error) {
    console.error('getClientRecordsDirectory error:', error);
    return {success:false,message:'We could not load your records. Please try again.',code:'RECORDS_DIRECTORY_ERROR'};
  }
}

function getClientRecordRows(recordKey, sessionId) {
  try {
    var session = validateClientSession(sessionId);
    if (!session) return {success:false,message:'Session expired. Please sign in again.',code:'SESSION_EXPIRED'};
    var modules = getClientAccessibleModulesV12_(session);
    var key = String(recordKey || ''), sheetName = '', moduleName = '';
    if (key.indexOf('BC:') === 0) {
      if (modules.indexOf('Ecommerce') === -1) return {success:false,message:'You do not have access to Business Center records.',code:'ACCESS_DENIED'};
      sheetName = key.substring(3); moduleName = 'Business Center';
    } else {
      moduleName = key;
      sheetName = getClientModuleSheetNameV12_(moduleName);
      if (!sheetName || modules.indexOf(moduleName) === -1) return {success:false,message:'You do not have access to this record type.',code:'ACCESS_DENIED'};
    }
    var ss = SpreadsheetApp.openById(session.sheetId), sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2 || !sheet.getLastColumn()) return {success:true,recordType:moduleName,records:[],headers:[]};
    var lastRow = sheet.getLastRow(), width = sheet.getLastColumn(), startRow = Math.max(2,lastRow-499);
    var headers = sheet.getRange(1,1,1,width).getValues()[0];
    var records = sheet.getRange(startRow,1,lastRow-startRow+1,width).getValues().reverse().map(function(row) {
      var item = {}; headers.forEach(function(h,i) { item[h] = row[i]; }); return item;
    });
    return {success:true,recordType:moduleName,headers:headers,records:records};
  } catch (error) {
    console.error('getClientRecordRows error:', error);
    return {success:false,message:'We could not load these records. Please try again.',code:'RECORD_ROWS_ERROR'};
  }
}
