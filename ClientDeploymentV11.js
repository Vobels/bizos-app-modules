// ============================================================
// ClientDeploymentV11.js - REAL BizOS CLIENT PACKAGE
// ============================================================
// Step 2 package builder.
// This does not alter the existing V10 generator or provisioning path.
// It copies the real BizOS UI/source from the current master project at
// generation time, then adds a client-scoped runtime/auth/workspace layer.
// ============================================================

var CLIENT_V11_UI_FILES_ = [
  'index','landing','dashboard','BusinessHubHTML','RecordsHTML','profile',
  'settings','staff','cms','notification','scripts','styles',
  'settings-controller','forms','ModalFix','ZZZ_UpgradeModalFix',
  'ZZZ_ProfilePageLayoutFix','UpgradeReviewFlow','SessionRouteGuard',
  'ZZZ_StaffInvitationUI'
];

var CLIENT_V11_SERVER_FILES_ = [
  'BusinessHubGS','Dashboard','Modules','Records',
  'Reference-data','Export','Profile'
];

function generateClientCodeSafelyV11(settings) {
  settings = settings || {};
  if (!settings.clientId || !settings.sheetId || !settings.email) {
    throw new Error('clientId, sheetId and email are required');
  }

  var config = {
    clientId: String(settings.clientId),
    clientName: String(settings.clientName || 'My Business'),
    primaryColor: String(settings.primaryColor || '#5D2A86'),
    logoUrl: String(settings.logoUrl || ''),
    email: String(settings.email || '').trim().toLowerCase(),
    sheetId: String(settings.sheetId),
    businessId: String(settings.businessId || settings.clientId),
    masterApiUrl: String(settings.masterApiUrl || getMasterApiUrl()),
    applicationUrl: String(settings.applicationUrl || ''),
    landingUrl: String(settings.landingUrl || ''),
    customDomain: String(settings.customDomain || '')
  };

  var masterFiles = fetchMasterProjectContentV11_();
  var selected = selectClientSourceFilesV11_(masterFiles);
  var files = [{
    name: 'Code',
    type: 'SERVER_JS',
    source: buildClientRuntimeV11_(config)
  }];

  selected.forEach(function(file) {
    var source = String(file.source || '');
    if (file.name === 'Modules') {
      source = renameFunctionV11_(source, 'getModuleSummary', 'getModuleSummaryLegacy_');
    }
    files.push({name:file.name,type:file.type,source:source});
  });

  files = files.map(function(file) {
    if (file.name === 'scripts') file.source = patchClientScriptsV11_(file.source, config);
    return file;
  });

  files.push({
    name: 'appsscript',
    type: 'JSON',
    source: JSON.stringify(buildClientManifestV11_(masterFiles))
  });

  return {
    success:true,
    version:'11.0',
    config:config,
    files:files,
    source:{
      masterScriptId:ScriptApp.getScriptId(),
      copiedUiFiles:CLIENT_V11_UI_FILES_.slice(),
      copiedServerFiles:CLIENT_V11_SERVER_FILES_.slice()
    }
  };
}

function fetchMasterProjectContentV11_() {
  var masterScriptId = String(ScriptApp.getScriptId() || '').trim();
  if (!masterScriptId) throw new Error('Master script ID is unavailable.');

  var url = 'https://script.googleapis.com/v1/projects/' +
    encodeURIComponent(masterScriptId) + '/content';

  var response = UrlFetchApp.fetch(url, {
    method:'get',
    headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},
    muteHttpExceptions:true
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error('Unable to read the master BizOS source package: ' + response.getContentText());
  }

  var result = JSON.parse(response.getContentText() || '{}');
  if (!Array.isArray(result.files)) throw new Error('Master BizOS source package returned no files.');
  return result.files;
}

function selectClientSourceFilesV11_(masterFiles) {
  var byName = {};
  (masterFiles || []).forEach(function(file) {
    if (!file || !file.name) return;
    var rawName = String(file.name);
    var normalizedName = rawName.replace(/\.(?:js|gs|html)$/i, '');
    byName[rawName] = file;
    byName[normalizedName] = file;
  });

  function getMasterFile_(name) {
    var key = String(name || '');
    return byName[key] || byName[key.replace(/\.(?:js|gs|html)$/i, '')] || null;
  }

  var selected = [];

  CLIENT_V11_UI_FILES_.forEach(function(name) {
    var file = getMasterFile_(name);
    if (!file || String(file.type || '') !== 'HTML') {
      throw new Error('Required BizOS UI file is missing from master: ' + name);
    }
    selected.push({
      name:String(name).replace(/\.html$/i,''),
      type:'HTML',
      source:String(file.source || '')
    });
  });

  CLIENT_V11_SERVER_FILES_.forEach(function(name) {
    var file = getMasterFile_(name);
    if (!file || String(file.type || '') !== 'SERVER_JS') {
      throw new Error('Required BizOS client-safe server file is missing from master: ' + name);
    }
    selected.push({
      name:String(name).replace(/\.js$/i,'').replace(/\.gs$/i,''),
      type:'SERVER_JS',
      source:String(file.source || '')
    });
  });

  return selected;
}

function renameFunctionV11_(source, oldName, newName) {
  var re = new RegExp('function\\s+' + oldName + '\\s*\\(', 'g');
  return String(source || '').replace(re, 'function ' + newName + '(');
}

function patchClientScriptsV11_(source, config) {
  var clientConfig = '<script>window.__CLIENT=' + JSON.stringify({
    isClient:true,id:config.clientId,name:config.clientName,
    primaryColor:config.primaryColor,logoUrl:config.logoUrl,
    applicationUrl:config.applicationUrl
  }) + ';</script>';

  var bridge = '<script>(function(){' +
    'var map={' +
      'getModuleData:"getClientModuleData",' +
      'getModuleDataWithSync:"getClientModuleDataWithSync",' +
      'saveRecord:"saveClientRecord",' +
      'updateRecord:"updateClientRecord",' +
      'updateRecordField:"updateClientRecordField",' +
      'deleteRecord:"deleteClientRecord",' +
      'getDashboardKPIs:"getClientDashboardData",' +
      'getEnhancedDashboardData:"getClientDashboardData",' +
      'getRecentTransactions:"getClientRecentTransactions",' +
      'getModuleSummary:"getClientModuleSummary",' +
      'getRecordsDirectory:"getClientRecordsDirectory",' +
      'getRecordRows:"getClientRecordRows"' +
    '};' +
    'var original=window.callGoogleScript;' +
    'if(typeof original==="function"){' +
      'window.callGoogleScript=function(name,params,ok,fail){' +
        'return original.call(this,map[name]||name,params||[],ok,fail);' +
      '};' +
    '}' +
  '})();</script>';

  return clientConfig + bridge + String(source || '');
}

function buildClientManifestV11_(masterFiles) {
  var manifest = {
    timeZone:'Africa/Lagos',
    exceptionLogging:'STACKDRIVER',
    runtimeVersion:'V8',
    webapp:{executeAs:'USER_DEPLOYING',access:'ANYONE_ANONYMOUS'},
    oauthScopes:[
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive',
      'https://www.googleapis.com/auth/script.external_request',
      'https://www.googleapis.com/auth/script.send_mail',
      'https://www.googleapis.com/auth/script.projects.readonly'
    ]
  };

  try {
    var masterManifest = (masterFiles || []).find(function(f){
      return f && String(f.name)==='appsscript';
    });
    if(masterManifest && masterManifest.source){
      var parsed=JSON.parse(masterManifest.source);
      if(parsed.timeZone)manifest.timeZone=parsed.timeZone;
      if(parsed.exceptionLogging)manifest.exceptionLogging=parsed.exceptionLogging;
    }
  } catch(ignore) {}

  return manifest;
}

function buildClientRuntimeV11_(c) {
  return [
'var CLIENT_CONFIG='+JSON.stringify(c)+';',
'var MODULES={Finance:{sheet:"Financial_Data",label:"Financial Management",icon:"calculator"},Sales:{sheet:"Sales_Data",label:"Sales Pipeline",icon:"chart-line"},Ecommerce:{sheet:"Ecommerce_Data",label:"E-commerce",icon:"cart"},CRM:{sheet:"CRM_Data",label:"Customer Relations",icon:"users"},HR:{sheet:"HR_Data",label:"Human Resources",icon:"briefcase"},Logistics:{sheet:"Logistics_Data",label:"Logistics",icon:"truck"},Tax:{sheet:"Tax_Data",label:"Tax Compliance",icon:"file-invoice"},Agro:{sheet:"Agro_Data",label:"Agriculture",icon:"seedling"},Productivity:{sheet:"Productivity_Data",label:"Productivity Suite",icon:"clock"},POS:{sheet:"POS_Data",label:"Point of Sale",icon:"cash-register"},Attendance:{sheet:"Attendance_Data",label:"Staff Attendance",icon:"fingerprint"},Warehouse:{sheet:"Warehouse_Data",label:"Warehouse Management",icon:"warehouse"}};',
'function doGet(e){var p=String(e&&e.parameter&&e.parameter.page||""),initialRoute="/";if(p==="dashboard")initialRoute="/dashboard";else if(p==="profile")initialRoute="/profile";else if(p==="staff")initialRoute="/staff";else if(p==="settings")initialRoute="/settings";var t=HtmlService.createTemplateFromFile("index");t.showLogin=true;t.routeMode="client";t.routeTarget=p||"landing";t.initialRoute=initialRoute;t.openBilling=false;t.editUpgrade=false;t.editUpgradeRequestId="";return t.evaluate().setTitle(CLIENT_CONFIG.clientName+" - BizOS").addMetaTag("viewport","width=device-width,initial-scale=1").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}',
'function include(filename){return HtmlService.createHtmlOutputFromFile(filename).getContent();}',
'function getWorkspaceFile(sid){requireClientSession_(sid);var ss=SpreadsheetApp.openById(CLIENT_CONFIG.sheetId),info=ss.getSheetByName("Client_Info");if(!info)throw new Error("Client workspace is not configured.");var v=info.getDataRange().getValues(),m={};for(var i=1;i<v.length;i++){if(v[i][0])m[String(v[i][0])]=v[i][1];}if(String(m.Client_ID||"")!==String(CLIENT_CONFIG.clientId))throw new Error("Client workspace identity mismatch.");if(String(m.Business_ID||CLIENT_CONFIG.businessId)!==String(CLIENT_CONFIG.businessId))throw new Error("Client business identity mismatch.");if(String(m.Status||"active").toLowerCase()!=="active")throw new Error("Client workspace is inactive.");return ss;}',
'function getOrCreateUserSheet(){var ss=SpreadsheetApp.openById(CLIENT_CONFIG.sheetId),s=ss.getSheetByName("Client_Users");if(!s){s=ss.insertSheet("Client_Users");s.appendRow(["Email","Password_Hash","Name","Phone","Role","Business_ID","Business_Name","Is_Demo","Department","Status","Profile_Photo_ID","Created_At","Updated_At","Is_Client"]);}return s;}',
'function getOrCreateBusinessSheet(){var ss=SpreadsheetApp.openById(CLIENT_CONFIG.sheetId),s=ss.getSheetByName("Client_Business");if(!s){s=ss.insertSheet("Client_Business");s.appendRow(["Business_ID","Business_Name","Owner_Email","Subscription_Tier","Workspace_ID","Status"]);s.appendRow([CLIENT_CONFIG.businessId,CLIENT_CONFIG.clientName,CLIENT_CONFIG.email,"sovereign",CLIENT_CONFIG.sheetId,"active"]);}return s;}',
'function addDefaultHeadersToModule(s,m){var h={Finance:["Transaction_ID","Date","Type","Category","Subcategory","Module_Source","Module_Record_ID","Description","Amount","Payment_Method","Reference","Status","Created_By","Created_At"],Sales:["Deal_ID","Date","Customer_Name","Customer_Email","Product_Service","Quantity","Unit_Price","Total_Amount","Stage","Probability","Expected_Close_Date","Sales_Person","Notes","Financial_Link_ID","Created_By"],Ecommerce:["Order_ID","Date","Customer_Email","Product_Name","SKU","Quantity","Unit_Price","Subtotal","Shipping_Cost","Tax","Total_Amount","Payment_Status","Fulfillment_Status","Tracking_Number","Financial_Link_ID","Created_By"],CRM:["Contact_ID","Name","Email","Phone","Company","Status","Last_Contact","Assigned_To","Notes","Created_By","Created_At"],HR:["Employee_ID","Name","Email","Phone","Department","Position","Employment_Type","Salary","Start_Date","Status","Created_By","Created_At"],Logistics:["Shipment_ID","Date","Customer_Name","Origin","Destination","Carrier","Tracking_Number","Status","Estimated_Delivery","Actual_Delivery","Cost","Created_By","Created_At"],Tax:["Tax_ID","Date","Type","Taxable_Amount","Rate","Tax_Amount","Status","Notes","Created_By","Created_At"],Agro:["Activity_ID","Date","Activity_Type","Crop","Quantity","Unit","Location","Cost","Revenue","Status","Notes","Created_By","Created_At"],Productivity:["Task_ID","Title","Description","Priority","Status","Due_Date","Assigned_To","Created_By","Created_At"],POS:["Sale_ID","Date","Customer_Name","Total_Amount","Payment_Method","Status","Created_By","Created_At"],Attendance:["Attendance_ID","Date","Employee_ID","Employee_Name","Check_In","Check_Out","Status","Created_By","Created_At"],Warehouse:["Item_ID","Name","SKU","Quantity","Unit_Cost","Location","Reorder_Level","Status","Created_By","Created_At"]};if(s.getLastRow()===0)s.appendRow(h[m]||["ID","Date","Description","Amount","Status","Created_By","Created_At"]);}',
'function getSheetHeadersBySheet(s){return s.getRange(1,1,1,s.getLastColumn()).getValues()[0];}',
'function getSheetHeaders(m,sid){var ss=getWorkspaceFile(sid),s=ss.getSheetByName(MODULES[m].sheet);if(!s){s=ss.insertSheet(MODULES[m].sheet);addDefaultHeadersToModule(s,m);}return getSheetHeadersBySheet(s);}',
'function generateTransactionId(){return "TXN-"+Utilities.getUuid().substring(0,8).toUpperCase();}',
'function clientHashV11_(p){return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(p||"")));}',
'function createSession(email,businessId){var sid="CLIENT_SESS_"+Utilities.getUuid(),now=Date.now(),u=findClientUser_(email,businessId),session={email:String(email).toLowerCase(),businessId:String(businessId||CLIENT_CONFIG.businessId),clientId:String(CLIENT_CONFIG.clientId),sheetId:String(CLIENT_CONFIG.sheetId),businessName:String(CLIENT_CONFIG.clientName),created:now,expiresAt:now+21600000,isClient:true,user:u};CacheService.getScriptCache().put(sid,JSON.stringify(session),21600);return sid;}',
'function validateClientLicenseV11_(s){try{var master=String(CLIENT_CONFIG.masterApiUrl||"");if(!master)return{success:false};var u="";try{u=String(ScriptApp.getService().getUrl()||"");}catch(ignore){}var match=u.match(/\\/s\\/([^\\/]+)/),deploymentId=match&&match[1]?match[1]:"";var r=UrlFetchApp.fetch(master,{method:"post",contentType:"application/json",muteHttpExceptions:true,payload:JSON.stringify({action:"validateClientLicense",payload:{clientId:s.clientId,businessId:s.businessId,sheetId:s.sheetId,scriptId:String(ScriptApp.getScriptId()||""),deploymentId:deploymentId}})});if(r.getResponseCode()<200||r.getResponseCode()>=300)return{success:false};var x=JSON.parse(r.getContentText()||"{}");return x&&x.success?x:{success:false};}catch(e){return{success:false};}}',
'function validateClientSession(sid){if(!sid)return null;try{var raw=CacheService.getScriptCache().get(String(sid));if(!raw)return null;var s=JSON.parse(raw),now=Date.now();if(!s||s.clientId!==String(CLIENT_CONFIG.clientId)||s.sheetId!==String(CLIENT_CONFIG.sheetId)||s.businessId!==String(CLIENT_CONFIG.businessId)||now>=Number(s.expiresAt||0)){CacheService.getScriptCache().remove(String(sid));return null;}var license=validateClientLicenseV11_(s);if(!license||!license.success){CacheService.getScriptCache().remove(String(sid));return null;}return s;}catch(e){return null;}}',
'function requireClientSession_(sid){var s=validateClientSession(sid);if(!s)throw new Error("Session expired. Please login again.");return s;}',
'function validateSession(sid){var s=validateClientSession(sid);return s?s.email:null;}',
'function getUserFromSession(sid){var s=requireClientSession_(sid);return s.user||findClientUser_(s.email,s.businessId);}',
'function findClientUser_(email,bid){var n=String(email||"").toLowerCase(),s=getOrCreateUserSheet(),v=s.getDataRange().getValues(),h=v[0]||[],ec=h.indexOf("Email"),bc=h.indexOf("Business_ID"),rc=h.indexOf("Role"),nc=h.indexOf("Name");for(var i=1;i<v.length;i++){if(String(v[i][ec]||"").toLowerCase()!==n)continue;if(bid&&bc>=0&&String(v[i][bc]||"")!==String(bid))continue;return{email:n,name:String(v[i][nc]||n),role:String(v[i][rc]||"owner"),businessId:String(bid||CLIENT_CONFIG.businessId),businessName:CLIENT_CONFIG.clientName,subscriptionTier:"sovereign",workspaceId:CLIENT_CONFIG.sheetId,isDemo:false,isClient:true,accessibleModules:Object.keys(MODULES)};}return{email:n,name:n,role:"owner",businessId:String(bid||CLIENT_CONFIG.businessId),businessName:CLIENT_CONFIG.clientName,subscriptionTier:"sovereign",workspaceId:CLIENT_CONFIG.sheetId,isDemo:false,isClient:true,accessibleModules:Object.keys(MODULES)};}',
'function userHasAccess(m,sid){var u=getUserFromSession(sid);return !!u&&u.accessibleModules.indexOf(String(m))>=0;}',
'function authenticateClient(email,password){try{email=String(email||"").trim().toLowerCase();password=String(password||"");var team=authenticateClientTeamMemberV11_(email,password);if(team&&team.success)return team;if(email!==String(CLIENT_CONFIG.email||"").toLowerCase())return{success:false,message:"This account is not a member of this business.",code:"NOT_CLIENT_ACCOUNT"};var r=UrlFetchApp.fetch(CLIENT_CONFIG.masterApiUrl,{method:"post",contentType:"application/json",muteHttpExceptions:true,payload:JSON.stringify({action:"clientLogin",payload:{clientId:CLIENT_CONFIG.clientId,email:email,password:password}})});if(r.getResponseCode()<200||r.getResponseCode()>=300)return{success:false,message:"Authentication service unavailable.",code:"AUTH_SERVICE"};var x=JSON.parse(r.getContentText()||"{}");if(!x||!x.success)return x||{success:false,message:"Invalid email or password"};if(!x.user||String(x.user.clientId||"")!==String(CLIENT_CONFIG.clientId))return{success:false,message:"This account is not authorized for this deployment.",code:"CLIENT_BINDING_MISMATCH"};var sid=createSession(email,CLIENT_CONFIG.businessId),cache=CacheService.getScriptCache(),session=JSON.parse(cache.get(sid));session.user={email:email,name:x.user.name||email.split("@")[0],role:x.user.role||"owner",businessId:CLIENT_CONFIG.businessId,businessName:CLIENT_CONFIG.clientName,subscriptionTier:x.user.subscriptionTier||"sovereign",workspaceId:CLIENT_CONFIG.sheetId,isDemo:false,isClient:true,accessibleModules:Object.keys(MODULES)};cache.put(sid,JSON.stringify(session),21600);return{success:true,sessionId:sid,user:session.user,message:"Login successful"};}catch(e){return{success:false,message:"Login error: "+e.message,code:"LOGIN_ERROR"};}}',
'function logout(sid){if(sid)CacheService.getScriptCache().remove(String(sid));return{success:true,message:"Logged out successfully"};}',
'function clientLogout(sid){return logout(sid);}',
'function getClientDashboard(sid,forceRefresh){\n  try{\n    var u=getUserFromSession(sid);\n    var cacheKey="BIZOS_CLIENT_DASH_"+String(CLIENT_CONFIG.clientId).replace(/[^A-Za-z0-9_]/g,"_");\n    if(!forceRefresh){\n      var cached=CacheService.getScriptCache().get(cacheKey);\n      if(cached){try{return{success:true,dashboard:JSON.parse(cached),cached:true};}catch(ignore){}}\n    }\n    var ss=getWorkspaceFile(sid),finance=ss.getSheetByName("Financial_Data");\n    var revenue=0,expenses=0,recent=[],daily={},expenseCategories={},financeCount=0;\n    if(finance&&finance.getLastRow()>1){\n      var v=finance.getDataRange().getValues(),h=v[0]||[];\n      var tc=h.indexOf("Type"),ac=h.indexOf("Amount"),dc=h.indexOf("Date"),xc=h.indexOf("Description"),cc=h.indexOf("Category");\n      financeCount=Math.max(0,v.length-1);\n      v.slice(1).forEach(function(r){\n        var a=Number(r[ac])||0,t=String(r[tc]||"").toLowerCase(),d=r[dc],key="";\n        if(d instanceof Date)key=Utilities.formatDate(d,Session.getScriptTimeZone(),"yyyy-MM-dd");\n        else if(d)key=String(d).slice(0,10);\n        if(t==="revenue"||t==="income"){revenue+=Math.abs(a);if(key){daily[key]=daily[key]||{revenue:0,expenses:0};daily[key].revenue+=Math.abs(a);}}\n        if(t==="expense"){expenses+=Math.abs(a);if(key){daily[key]=daily[key]||{revenue:0,expenses:0};daily[key].expenses+=Math.abs(a);}var cat=String(cc>=0?r[cc]:"Uncategorized")||"Uncategorized";expenseCategories[cat]=(expenseCategories[cat]||0)+Math.abs(a);}\n      });\n      recent=v.slice(1).filter(function(r){return dc>=0&&r[dc];}).slice(-10).reverse().map(function(r){\n        return{date:r[dc],description:xc>=0?r[xc]:"",amount:ac>=0?r[ac]:0,type:tc>=0?r[tc]:"",category:cc>=0?r[cc]:""};\n      });\n    }\n    var moduleCounts={},moduleNames=Object.keys(MODULES);\n    moduleNames.forEach(function(name){try{var sh=ss.getSheetByName(MODULES[name].sheet);moduleCounts[name]=sh?Math.max(0,sh.getLastRow()-1):0;}catch(e){moduleCounts[name]=0;}});\n    var teamSheet=ss.getSheetByName("Client_Team"),teamCount=teamSheet&&teamSheet.getLastRow()>1?teamSheet.getLastRow()-1:0;\n    var profit=revenue-expenses,margin=revenue?profit/revenue*100:0;\n    var trendKeys=Object.keys(daily).sort().slice(-7),trend={labels:[],revenue:[],expenses:[]};\n    trendKeys.forEach(function(k){trend.labels.push(k);trend.revenue.push(daily[k].revenue||0);trend.expenses.push(daily[k].expenses||0);});\n    var expPairs=Object.keys(expenseCategories).sort(function(a,b){return expenseCategories[b]-expenseCategories[a];}).slice(0,6);\n    var breakdown={labels:expPairs,values:expPairs.map(function(k){return expenseCategories[k];})};\n    var alerts=[],recommendations=[];\n    if(financeCount===0)recommendations.push({type:"start",title:"Start tracking your finances",message:"Add your first income or expense to unlock financial insights."});\n    if(financeCount>0&&revenue===0)recommendations.push({type:"info",title:"Record your income",message:"Add revenue entries so BizOS can calculate profit and business trends."});\n    if(revenue>0&&expenses>revenue)alerts.push({type:"warning",title:"Expenses are above revenue",message:"Your recorded expenses currently exceed recorded revenue."});\n    var health=null;\n    if(financeCount>0){var marginScore=Math.max(0,Math.min(100,50+margin));health={overall:Math.round(marginScore),metrics:[{label:"Profit margin",value:margin.toFixed(1)+"%"},{label:"Financial records",value:String(financeCount)},{label:"Team members",value:String(teamCount)}],recommendations:recommendations};}\n    else health={overall:null,metrics:[],recommendations:recommendations};\n    var dashboard={\n      kpis:{revenue:revenue,expenses:expenses,profit:profit,margin:margin},\n      modules:moduleNames,\n      moduleCounts:moduleCounts,\n      counts:{finance:financeCount,products:(moduleCounts.Ecommerce||0)+(moduleCounts.POS||0),modules:moduleNames.length,team:teamCount},\n      charts:{revenueVsExpenses:trend,expenseBreakdown:breakdown},\n      health:health,\n      alerts:alerts,\n      recommendations:recommendations,\n      recentActivity:recent.slice(0,8).map(function(x){return{date:x.date,title:x.description||x.type||"Financial entry",type:String(x.type||"").toLowerCase(),amount:x.amount};}),\n      recentTransactions:recent,\n      clientName:CLIENT_CONFIG.clientName,\n      primaryColor:CLIENT_CONFIG.primaryColor,\n      logoUrl:CLIENT_CONFIG.logoUrl,\n      workspaceId:CLIENT_CONFIG.sheetId,\n      userEmail:u.email,\n      generatedAt:new Date().toISOString()\n    };\n    var payload=JSON.stringify(dashboard);\n    try{CacheService.getScriptCache().put(cacheKey,payload,60);}catch(ignore2){}\n    return{success:true,dashboard:dashboard,cached:false};\n  }catch(e){return{success:false,message:e.message,code:"CLIENT_DASHBOARD_ERROR"};}\n}\nfunction getClientDashboardData(sid,forceRefresh){\n  var b=getClientDashboard(sid,forceRefresh);\n  if(!b.success)return b;\n  var d=b.dashboard;\n  return{success:true,cached:!!b.cached,dashboard:{\n    kpis:{revenue:{value:d.kpis.revenue,change:0},expenses:{value:d.kpis.expenses,change:0},netProfit:{value:d.kpis.profit,change:0},profitMargin:{value:d.kpis.margin,change:0}},\n    counts:d.counts,moduleCounts:d.moduleCounts,moduleSummary:d.modules,\n    health:d.health,charts:d.charts,recentActivity:d.recentActivity,alerts:d.alerts,recommendations:d.recommendations,recentTransactions:d.recentTransactions,\n    clientName:d.clientName,primaryColor:d.primaryColor,logoUrl:d.logoUrl,workspaceId:d.workspaceId,userEmail:d.userEmail,generatedAt:d.generatedAt\n  }};\n}\nfunction getClientModuleSummary(sid){var out={};Object.keys(MODULES).forEach(function(k){out[k]={label:MODULES[k].label,icon:MODULES[k].icon};});return out;}','function getModuleSummary(sid){return getClientModuleSummary(sid);}',
'function getClientModuleData(m,sid){try{var r=getModuleData(String(m||"Finance"),sid)||[];return r.map(function(x,i){x._row=i+2;return x;});}catch(e){return[];}}',
'function getClientModuleDataWithSync(m,sid){try{var r=getModuleData(String(m||"Finance"),sid)||[],h=getSheetHeaders(String(m||"Finance"),sid);return{success:true,headers:h,records:r.map(function(x,i){x._row=i+2;return x;})};}catch(e){return{success:false,message:e.message,records:[]};}}',
'function saveClientRecord(m,d,sid){return saveRecord(m,d,sid);}',
'function updateClientRecord(m,row,d,sid){return updateRecord(m,row,d,sid);}',
'function updateClientRecordField(m,row,field,value,sid){return updateRecordField(m,row,field,value,sid);}',
'function deleteClientRecord(m,row,sid){return deleteRecord(m,row,sid);}',
'function getClientRecentTransactions(limit,sid){try{return(getClientModuleData("Finance",sid).records||[]).slice(-Number(limit||10)).reverse();}catch(e){return[];}}',
'function getClientRecordsDirectory(sid){return getRecordsDirectory(sid);}',
'function getClientRecordRows(k,sid){return getRecordRows(k,sid);}',
'function getStaffModules(){return Object.keys(MODULES);}',
'function getAvailableModulesForStaff(){return Object.keys(MODULES);}',
'function ensureClientTeamSheetV11_(){var ss=SpreadsheetApp.openById(CLIENT_CONFIG.sheetId),s=ss.getSheetByName("Client_Team");if(!s){s=ss.insertSheet("Client_Team");s.appendRow(["User_ID","Email","Name","Phone","Role","Department","Status","Password_Hash","Must_Change_Password","Invited_By","Created_At","Updated_At","Last_Login","Assigned_Modules"]);}else{var lastCol=s.getLastColumn();if(lastCol<14)s.getRange(1,14).setValue("Assigned_Modules");}return s;}',
'function clientTeamRowsV11_(s){var v=s.getDataRange().getValues(),h=v[0]||[];return v.length<2?[]:v.slice(1).map(function(r,i){var o={row:i+2};h.forEach(function(x,j){if(x)o[String(x)]=r[j]===undefined?"":r[j];});return o;});}',
'function clientAssignedModulesV11_(team){var role=String(team&&team.Role||"staff").toLowerCase();if(role==="owner"||role==="admin")return Object.keys(MODULES);var raw=String(team&&team.Assigned_Modules||"").trim();if(!raw)return[];try{var parsed=JSON.parse(raw);if(!Array.isArray(parsed))return[];var allowed=Object.keys(MODULES);return parsed.map(function(m){return String(m||"").trim();}).filter(function(m){return m&&allowed.indexOf(m)>=0;});}catch(e){return[];}}\n\nfunction authenticateClientTeamMemberV11_(email,password){try{var s=ensureClientTeamSheetV11_(),t=clientTeamRowsV11_(s).find(function(r){return String(r.Email||"").toLowerCase()===String(email||"").toLowerCase();});if(!t)return null;if(String(t.Status||"").toLowerCase()==="suspended")return{success:false,message:"Your team access is suspended.",code:"TEAM_SUSPENDED"};if(String(t.Password_Hash||"")!==clientHashV11_(password))return{success:false,message:"Incorrect password",code:"AUTH_FAILED"};var license=validateClientLicenseV11_({clientId:CLIENT_CONFIG.clientId,businessId:CLIENT_CONFIG.businessId,sheetId:CLIENT_CONFIG.sheetId});if(!license||!license.success)return{success:false,message:"This client workspace is not currently authorized.",code:"CLIENT_LICENSE_INVALID"};var sid=createSession(email,CLIENT_CONFIG.businessId),cache=CacheService.getScriptCache(),session=JSON.parse(cache.get(sid));session.user={email:String(t.Email),name:String(t.Name||t.Email),role:String(t.Role||"staff"),businessId:CLIENT_CONFIG.businessId,businessName:CLIENT_CONFIG.clientName,subscriptionTier:"sovereign",workspaceId:CLIENT_CONFIG.sheetId,isDemo:false,isClient:true,accessibleModules:clientAssignedModulesV11_(t)};cache.put(sid,JSON.stringify(session),21600);return{success:true,sessionId:sid,user:session.user,message:"Login successful",mustChangePassword:String(t.Must_Change_Password||"").toUpperCase()==="YES"};}catch(e){return{success:false,message:e.message,code:"TEAM_LOGIN_ERROR"};}}',
'function getBusinessStaff(bid,sid){var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(bid))return{success:false,message:"Unauthorized"};return{success:true,members:clientTeamRowsV11_(ensureClientTeamSheetV11_()).map(function(r){return{email:r.Email,name:r.Name,role:r.Role,phone:r.Phone,department:r.Department,status:r.Status,assignedModules:clientAssignedModulesV11_(r),invitationStatus:r.Status};})};}',
'function inviteStaffMember(bid,email,name,role,invitedByEmail,assignedModules,sid){try{var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(bid)||["owner","admin"].indexOf(String(u.role).toLowerCase())<0)return{success:false,message:"Only the business owner or an admin can invite staff."};email=String(email||"").trim().toLowerCase();name=String(name||"").trim();if(!email||!name)return{success:false,message:"Name and email are required."};var s=ensureClientTeamSheetV11_(),rows=clientTeamRowsV11_(s);if(rows.some(function(r){return String(r.Email||"").toLowerCase()===email&&String(r.Status||"").toLowerCase()==="active";}))return{success:false,message:"This email is already an active team member."};var pass=Utilities.getUuid().replace(/-/g,"").slice(0,12)+"A!",now=new Date().toISOString(),uid="TEAM_"+Utilities.getUuid().slice(0,8).toUpperCase(),inviteRole=role==="admin"?"admin":"staff",rawModules=Array.isArray(assignedModules)?assignedModules:[],allowedModules=Object.keys(MODULES),storedModules=rawModules.map(function(m){return String(m||"").trim();}).filter(function(m,i,a){return m&&allowedModules.indexOf(m)>=0&&a.indexOf(m)===i;});if(inviteRole==="admin")storedModules=allowedModules.slice();s.appendRow([uid,email,name,"",inviteRole,"","invited",clientHashV11_(pass),"YES",u.email,now,now,"",JSON.stringify(storedModules)]);MailApp.sendEmail({to:email,subject:"You have been invited to "+CLIENT_CONFIG.clientName+" on BizOS",htmlBody:"<p>Hello <strong>"+name+"</strong>,</p><p>You have been invited to <strong>"+CLIENT_CONFIG.clientName+"</strong> on BizOS.</p><p><strong>Email:</strong> "+email+"<br><strong>Temporary password:</strong> "+pass+"</p><p>Please sign in and change your password from your profile.</p>"});return{success:true,message:"Invitation sent to "+email+".",userId:uid};}catch(e){return{success:false,message:e.message,code:"STAFF_INVITE_ERROR"};}}',
'function removeTeamMember(bid,email,sid){try{var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(bid)||["owner","admin"].indexOf(String(u.role).toLowerCase())<0)return{success:false,message:"Unauthorized"};var s=ensureClientTeamSheetV11_(),rows=clientTeamRowsV11_(s),t=rows.find(function(r){return String(r.Email||"").toLowerCase()===String(email||"").toLowerCase();});if(!t)return{success:false,message:"Team member not found."};s.deleteRow(t.row);return{success:true,message:"Team member removed."};}catch(e){return{success:false,message:e.message};}}',
'function resendStaffInvitation(){return{success:false,message:"Invitation resend is not enabled in the client package yet.",code:"NOT_YET_ENABLED"};}',
'function updateUserProfile(bid,name,phone,sid){try{var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(CLIENT_CONFIG.businessId))return{success:false,message:"Unauthorized"};var s=getOrCreateUserSheet(),v=s.getDataRange().getValues(),h=v[0],ec=h.indexOf("Email"),bc=h.indexOf("Business_ID"),nc=h.indexOf("Name"),pc=h.indexOf("Phone"),email=String(u.email||"").toLowerCase();for(var i=1;i<v.length;i++){if(String(v[i][ec]||"").toLowerCase()!==email)continue;if(bc>=0&&String(v[i][bc]||CLIENT_CONFIG.businessId)!==String(CLIENT_CONFIG.businessId))continue;if(nc>=0)s.getRange(i+1,nc+1).setValue(name);if(pc>=0)s.getRange(i+1,pc+1).setValue(phone);return{success:true};}return{success:false,message:"User profile not found"};}catch(e){return{success:false,message:e.message};}}',
'function changeUserPassword(email,currentPassword,newPassword,sid){try{var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(CLIENT_CONFIG.businessId))return{success:false,message:"Unauthorized"};var s=getOrCreateUserSheet(),v=s.getDataRange().getValues(),h=v[0],ec=h.indexOf("Email"),pc=h.indexOf("Password_Hash"),sessionEmail=String(u.email||"").toLowerCase();for(var i=1;i<v.length;i++){if(String(v[i][ec]||"").toLowerCase()!==sessionEmail)continue;if(String(v[i][pc]||"")!==clientHashV11_(currentPassword))return{success:false,message:"Current password is incorrect"};s.getRange(i+1,pc+1).setValue(clientHashV11_(newPassword));return{success:true};}return{success:false,message:"User password record not found"};}catch(e){return{success:false,message:e.message};}}',
'function getFeatureNotifications(){return{success:true,notifications:[]};}',
'function getBusinessDetails(sid){var u=getUserFromSession(sid);return{success:!!u,business:{businessId:CLIENT_CONFIG.businessId,businessName:CLIENT_CONFIG.clientName,ownerEmail:CLIENT_CONFIG.email,subscriptionTier:"sovereign",workspaceId:CLIENT_CONFIG.sheetId}};}',
'function getCurrentUpgradePaymentStatus(sid){requireClientSession_(sid);return{success:true,status:"active",message:"Your client workspace is active."};}',
'function getPendingUpgradeRequestStatus(sid){requireClientSession_(sid);return{success:true,pending:false};}',
'function handleUpgradeRequest(){return{success:false,message:"Billing/provisioning is managed by the BizOS master service.",code:"MASTER_ONLY"};}',
'function deleteBusinessAccount(){return{success:false,message:"Account deletion is managed by BizOS support.",code:"MASTER_ONLY"};}',
'function requestDataExport(email,bid,sid){try{getUserFromSession(sid);MailApp.sendEmail({to:CLIENT_CONFIG.email,subject:"BizOS data export request - "+CLIENT_CONFIG.businessId,body:"A data export was requested on "+new Date().toISOString()+"."});return{success:true};}catch(e){return{success:false,message:e.message};}}',
'function getFrontendConfig(){return{app:{name:"BizOS",tagline:"Business Operating System"},modules:{all:Object.keys(MODULES)},features:{enableCharts:true,enableTeamManagement:true},client:{defaultPrimaryColor:CLIENT_CONFIG.primaryColor,defaultTier:"sovereign"}};}',
'function getCountryRequirements(country){return{success:true,country:country||"Nigeria",requiredFields:[],fieldLabels:{},placeholders:{}};}',
'function getClientPublicUrlsV11(){return{applicationUrl:String(CLIENT_CONFIG.applicationUrl||ScriptApp.getService().getUrl()||""),landingUrl:String(CLIENT_CONFIG.landingUrl||""),customDomain:String(CLIENT_CONFIG.customDomain||"")};}'
  ].join(String.fromCharCode(10));
}

function validateClientDeploymentPackageV11_(pkg, expected) {
  if(!pkg||!Array.isArray(pkg.files)||!pkg.files.length)return{success:false,code:'PACKAGE_EMPTY',message:'Generated V11 client package is empty.'};
  var names=pkg.files.map(function(f){return String(f.name||'');});
  var required=['Code','index','landing','dashboard','BusinessHubHTML','BusinessHubGS','scripts','styles','appsscript'];
  var missing=required.filter(function(n){return names.indexOf(n)<0;});
  if(missing.length)return{success:false,code:'V11_PACKAGE_MISSING_FILES',message:'V11 package is missing: '+missing.join(', ')};
  var code=pkg.files.filter(function(f){return f.name==='Code';})[0].source||'';
  var forbidden=['function autoSetupClient','function redeployClientByBusinessId','function requireAdminSession_','function getBizOSMasterSpreadsheet_','function saveClientRecordV2','function createAndDeployClientScriptSafe'];
  var leaked=forbidden.filter(function(token){return code.indexOf(token)>=0;});
  if(leaked.length)return{success:false,code:'V11_MASTER_FUNCTION_LEAK',message:'V11 client runtime contains forbidden master functions: '+leaked.join(', ')};
  if(expected&&expected.clientId&&code.indexOf(String(expected.clientId))<0)return{success:false,code:'V11_CLIENT_ID_MISMATCH',message:'V11 runtime does not contain the expected client ID.'};
  if(expected&&expected.sheetId&&code.indexOf(String(expected.sheetId))<0)return{success:false,code:'V11_WORKSPACE_ID_MISMATCH',message:'V11 runtime does not contain the expected workspace ID.'};
  return{success:true,code:'V11_PACKAGE_VALID',version:'11.0',fileCount:pkg.files.length,files:names};
}
