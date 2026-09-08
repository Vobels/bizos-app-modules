// ============================================================
// ClientAuth.js - CLIENT DEPLOYMENT AUTH + DATA ACCESS
// ============================================================

function getClientDeploymentConfig() {
  try {
    if (typeof CLIENT_CONFIG !== 'undefined') {
      return {
        success: true,
        clientId: CLIENT_CONFIG.clientId || '',
        clientName: CLIENT_CONFIG.clientName || '',
        sheetId: CLIENT_CONFIG.sheetId || '',
        businessId: CLIENT_CONFIG.businessId || CLIENT_CONFIG.clientId || '',
        masterApiUrl: CLIENT_CONFIG.masterApiUrl || ''
      };
    }
  } catch (e) {}
  return { success: false, message: 'Client deployment configuration is unavailable.' };
}

function clientLogin(clientId, email, password, sheetId, businessId, businessName) {
  try {
    if (!clientId || !email || !password) return { success:false, message:'Email, password and client ID are required', code:'INVALID_INPUT' };

    var client = getClientById(clientId);
    if (!client && sheetId) client = getClientFromWorkspace(sheetId, clientId);
    if (!client) return { success:false, message:'Client workspace could not be identified. Please contact support.', code:'CLIENT_NOT_FOUND' };

    var authResult = null;
    try {
      if (typeof authenticateUser === 'function') authResult = authenticateUser(email, password, null);
    } catch (localAuthError) {
      console.log('Local authentication unavailable:', localAuthError.message);
    }
    if (!authResult || !authResult.success) authResult = authenticateAgainstMaster(email, password, businessName || client.clientName || '', client);
    if (!authResult || !authResult.success) return { success:false, message:authResult && authResult.message ? authResult.message : 'Invalid email or password', code:'AUTH_FAILED' };

    var user = authResult.user || { email:email.toLowerCase(), name:email.split('@')[0], role:'owner', isClient:true };
    var resolvedSheetId = sheetId || client.sheetId || '';
    if (!resolvedSheetId) return { success:false, message:'Client workspace is not configured.', code:'NO_WORKSPACE' };

    var sessionId = 'CLIENT_SESS_' + Utilities.getUuid();
    var session = { email:email.toLowerCase(), clientId:clientId, sheetId:resolvedSheetId, businessId:businessId || client.businessId || clientId, businessName:businessName || client.clientName || 'My Business', created:Date.now(), isClient:true, user:user, apiKey:client.apiKey || '' };
    CacheService.getScriptCache().put(sessionId, JSON.stringify(session), 86400);

    return { success:true, sessionId:sessionId, message:'Login successful', clientName:client.clientName || businessName || 'My Business', primaryColor:client.primaryColor || '#5D2A86', logoUrl:client.logoUrl || '', user:{ email:user.email || email.toLowerCase(), name:user.name || email.split('@')[0], role:user.role || 'owner', clientId:clientId, businessId:businessId || client.businessId || clientId }, code:'SUCCESS' };
  } catch (error) {
    console.error('clientLogin error:', error);
    return { success:false, message:'Login error: ' + error.message, code:'LOGIN_ERROR' };
  }
}

function authenticateAgainstMaster(email, password, businessName, client) {
  try {
    var masterUrl = '';
    if (typeof getMasterApiUrl === 'function') masterUrl = getMasterApiUrl();
    if (!masterUrl && typeof CLIENT_CONFIG !== 'undefined') masterUrl = CLIENT_CONFIG.masterApiUrl || '';
    if (!masterUrl) return { success:false, message:'Master authentication service is not configured.' };

    var response = UrlFetchApp.fetch(masterUrl, { method:'post', contentType:'application/json', payload:JSON.stringify({ action:'LOGIN', payload:{ email:email, password:password, businessName:businessName || '' } }), muteHttpExceptions:true });
    return JSON.parse(response.getContentText()) || { success:false, message:'Authentication service returned no result.' };
  } catch (error) {
    console.error('Master authentication error:', error);
    return { success:false, message:'Unable to contact authentication service: ' + error.message };
  }
}

function getClientFromWorkspace(sheetId, expectedClientId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var info = ss.getSheetByName('Client_Info');
    if (!info) return null;
    var rows = info.getDataRange().getValues(), values = {};
    for (var i=1; i<rows.length; i++) if (rows[i][0]) values[String(rows[i][0])] = rows[i][1];
    if (expectedClientId && String(values.Client_ID || '') !== String(expectedClientId)) return null;
    return { clientId:values.Client_ID || expectedClientId, clientName:values.Business_Name || 'My Business', businessName:values.Business_Name || 'My Business', tier:values.Tier || 'sovereign', status:values.Status || 'active', sheetId:sheetId, primaryColor:values.Primary_Color || '#5D2A86', logoUrl:values.Logo_Url || '', businessId:values.Business_ID || expectedClientId, apiKey:values.API_Key || '' };
  } catch (error) { console.error('getClientFromWorkspace error:', error); return null; }
}

function validateClientSession(sessionId) {
  if (!sessionId) return null;
  try {
    var cache=CacheService.getScriptCache(), raw=cache.get(sessionId);
    if (!raw) return null;
    var session=JSON.parse(raw);
    if (!session.created || Date.now()-session.created>86400000) { cache.remove(sessionId); return null; }
    var client=getClientById(session.clientId);
    if (!client && session.sheetId) client=getClientFromWorkspace(session.sheetId,session.clientId);
    if (!client) return null;
    if (client.status && String(client.status).toLowerCase()!=='active') return null;
    return session;
  } catch(error) { console.error('validateClientSession error:',error); return null; }
}

function getClientUserFromSession(sessionId) {
  var session=validateClientSession(sessionId); if(!session) return null;
  var client=getClientById(session.clientId); if(!client && session.sheetId) client=getClientFromWorkspace(session.sheetId,session.clientId);
  return { email:session.email, name:session.user&&session.user.name?session.user.name:session.email, role:session.user&&session.user.role?session.user.role:'owner', clientId:session.clientId, businessId:session.businessId, clientName:client?(client.clientName||client.businessName):session.businessName, primaryColor:client&&client.primaryColor?client.primaryColor:'#5D2A86', logoUrl:client&&client.logoUrl?client.logoUrl:'', sheetId:session.sheetId, isClient:true, apiKey:client&&client.apiKey?client.apiKey:'' };
}

function clientLogout(sessionId) { try { if(sessionId) CacheService.getScriptCache().remove(sessionId); return {success:true,message:'Logged out successfully',code:'LOGOUT_SUCCESS'}; } catch(error) { return {success:false,message:error.message,code:'LOGOUT_ERROR'}; } }

function getClientDashboardData(sessionId) {
  try {
    var session=validateClientSession(sessionId); if(!session) return {success:false,message:'Session expired. Please login again.',code:'SESSION_EXPIRED'};
    var client=getClientById(session.clientId); if(!client) client=getClientFromWorkspace(session.sheetId,session.clientId);
    var sheetId=session.sheetId || (client&&client.sheetId); if(!sheetId) return {success:true,dashboard:getDefaultDashboardData(session.businessName),code:'NO_SHEET'};
    var ss=SpreadsheetApp.openById(sheetId), financeSheet=ss.getSheetByName('Financial_Data'), revenue=0, expenses=0, recentTransactions=[];
    if(financeSheet){ var data=financeSheet.getDataRange().getValues(); if(data.length>1){ var headers=data[0], typeCol=headers.indexOf('Type'), amountCol=headers.indexOf('Amount'), dateCol=headers.indexOf('Date'), descCol=headers.indexOf('Description'); for(var i=1;i<data.length;i++){var amount=parseFloat(data[i][amountCol])||0,type=String(data[i][typeCol]||'').toLowerCase();if(type==='revenue'||type==='income')revenue+=Math.abs(amount);else if(type==='expense')expenses+=Math.abs(amount);} if(dateCol!==-1) recentTransactions=data.slice(1).filter(function(r){return r[dateCol];}).sort(function(a,b){return new Date(b[dateCol])-new Date(a[dateCol]);}).slice(0,5).map(function(r){return {date:r[dateCol],description:descCol!==-1?r[descCol]:'',amount:amountCol!==-1?r[amountCol]:0,type:typeCol!==-1?r[typeCol]:''};}); } }
    var profit=revenue-expenses, margin=revenue>0?(profit/revenue)*100:0, names=['Finance','Ecommerce','Sales','CRM','HR','Logistics','Tax','Agro','Productivity','POS','Attendance','Warehouse'];
    var modules=names.map(function(name){var sheet=ss.getSheetByName(name+'_Data')||ss.getSheetByName(name);return {name:name,label:getModuleLabel(name),icon:getModuleIcon(name),count:sheet&&sheet.getLastRow()>1?sheet.getLastRow()-1:0,isAccessible:true};});
    return {success:true,dashboard:{kpis:{revenue:Math.round(revenue*100)/100,expenses:Math.round(expenses*100)/100,profit:Math.round(profit*100)/100,margin:Math.round(margin*10)/10},modules:modules,recentTransactions:recentTransactions,clientName:client?(client.clientName||client.businessName):session.businessName,primaryColor:client&&client.primaryColor?client.primaryColor:'#5D2A86',logoUrl:client&&client.logoUrl?client.logoUrl:'',sheetId:sheetId},code:'SUCCESS'};
  } catch(error) { console.error('getClientDashboardData error:',error); return {success:false,message:error.message,code:'DASHBOARD_ERROR'}; }
}

function getClientModuleData(moduleName,sessionId){try{var session=validateClientSession(sessionId);if(!session)return{success:false,message:'Session expired',code:'SESSION_EXPIRED'};if(!session.sheetId)return{success:true,data:[],code:'NO_SHEET'};var ss=SpreadsheetApp.openById(session.sheetId),sheet=ss.getSheetByName(moduleName+'_Data')||ss.getSheetByName(moduleName);if(!sheet)return{success:true,data:[],code:'SHEET_NOT_FOUND'};var values=sheet.getDataRange().getValues();if(values.length<2)return{success:true,data:[],code:'NO_DATA'};var headers=values[0],records=values.slice(1).map(function(row){var obj={};headers.forEach(function(h,i){if(h&&String(h).trim())obj[String(h).trim()]=row[i]!==undefined?row[i]:'';});return obj;});return{success:true,data:records,code:'SUCCESS'};}catch(error){return{success:false,message:error.message,code:'FETCH_ERROR'};}}

function saveClientModuleRecord(moduleName,record,sessionId){try{var session=validateClientSession(sessionId);if(!session)return{success:false,message:'Session expired',code:'SESSION_EXPIRED'};if(!session.sheetId)return{success:false,message:'No sheet found for client',code:'NO_SHEET'};var ss=SpreadsheetApp.openById(session.sheetId),sheet=ss.getSheetByName(moduleName+'_Data')||ss.getSheetByName(moduleName);if(!sheet){sheet=ss.insertSheet(moduleName+'_Data');sheet.appendRow(['Date','Description','Amount','Created_At']);}var headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];sheet.appendRow(headers.map(function(h){return record[h]!==undefined?record[h]:'';}));return{success:true,message:'Record saved successfully',code:'SAVE_SUCCESS'};}catch(error){return{success:false,message:error.message,code:'SAVE_ERROR'};}}

function getDefaultDashboardData(clientName){var names=['Finance','Ecommerce'];return{kpis:{revenue:0,expenses:0,profit:0,margin:0},modules:names.map(function(n){return{name:n,label:getModuleLabel(n),icon:getModuleIcon(n),count:0};}),recentTransactions:[],clientName:clientName||'My Business',primaryColor:'#5D2A86',logoUrl:''};}
function getModuleLabel(n){var x={Finance:'Financial Management',Ecommerce:'E-commerce',Sales:'Sales Pipeline',CRM:'Customer Relations',HR:'Human Resources',Logistics:'Logistics',Tax:'Tax Compliance',Agro:'Agriculture',Productivity:'Productivity Suite',POS:'Point of Sale',Attendance:'Staff Attendance',Warehouse:'Warehouse Management'};return x[n]||n;}
function getModuleIcon(n){var x={Finance:'fa-calculator',Ecommerce:'fa-cart-shopping',Sales:'fa-chart-line',CRM:'fa-users',HR:'fa-briefcase',Logistics:'fa-truck',Tax:'fa-file-invoice',Agro:'fa-seedling',Productivity:'fa-clock',POS:'fa-cash-register',Attendance:'fa-fingerprint',Warehouse:'fa-warehouse'};return x[n]||'fa-folder';}