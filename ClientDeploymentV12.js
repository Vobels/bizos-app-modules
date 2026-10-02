// ============================================================
// ClientDeploymentV12.js - DISTINCT CLIENT PRODUCT PACKAGE
// ============================================================
// V12 intentionally does NOT copy the Master BizOS UI.
// It reuses the proven client-safe server/runtime layer from V11,
// but serves an independently designed client-facing application.
// ============================================================

var CLIENT_V12_UI_FILES_ = ['ClientShellV12','ClientStylesV12'];

function generateClientCodeSafelyV12(settings) {
  settings = settings || {};
  if (!settings.clientId || !settings.sheetId || !settings.email) throw new Error('clientId, sheetId and email are required');

  var config = {
    clientId:String(settings.clientId),
    clientName:String(settings.clientName || 'My Business'),
    primaryColor:String(settings.primaryColor || '#2E7D32'),
    logoUrl:String(settings.logoUrl || ''),
    email:String(settings.email || '').trim().toLowerCase(),
    sheetId:String(settings.sheetId),
    businessId:String(settings.businessId || settings.clientId),
    masterApiUrl:String(settings.masterApiUrl || getMasterApiUrl()),
    applicationUrl:String(settings.applicationUrl || ''),
    landingUrl:String(settings.landingUrl || ''),
    customDomain:String(settings.customDomain || '')
  };
  config.businessProfile = getMasterBusinessProfileForV12_(config.businessId, config.email);

  var runtime = buildClientRuntimeV11_(config);
  var businessCenterSource = String(getMasterSourceForV12_('BusinessHubGS') || '');
  var staffSource = String(getMasterSourceForV12_('ClientStaffV12') || '');
  var appsSource = String(getMasterSourceForV12_('ClientAppsV12') || '');
  var settingsSource = String(getMasterSourceForV12_('ClientSettingsV12') || '');
  if(!staffSource) throw new Error('ClientStaffV12 source is missing.');
  if(!appsSource) throw new Error('ClientAppsV12 source is missing.');
  if(!settingsSource) throw new Error('ClientSettingsV12 source is missing.');
  if(!businessCenterSource) throw new Error('BusinessHubGS source is missing.');
  var businessCenterBridge = [
    'function clientSafeBusinessCenterResult_(result){try{return JSON.parse(JSON.stringify(result));}catch(e){return result;}}',
    'function getClientBusinessCenterData(sid){return clientSafeBusinessCenterResult_(getBusinessCenterData(sid));}',
    'function saveClientBusinessCenterProduct(data,sid){return saveBusinessCenterProduct(data,sid);}',
    'function updateClientBusinessCenterProduct(data,sid){return updateBusinessCenterProduct(data,sid);}',
    'function saveClientBusinessCenterCustomer(data,sid){return saveBusinessCenterCustomer(data,sid);}',
    'function getClientBusinessCenterSales(payload,sid){return clientSafeBusinessCenterResult_(getBusinessCenterSales(payload,sid));}',
    'function completeClientBusinessCenterSale(data,sid){return completeBusinessCenterSale(data,sid);}',
    "function adjustClientBusinessCenterStock(data,sid){return adjustBusinessCenterStock(data,sid);}",
    "function clientFinanceAccessV12_(sid,write){var user=getUserFromSession(sid);if(!user)return{ok:false,message:'Session expired. Please login again.'};var modules=Array.isArray(user.accessibleModules)?user.accessibleModules:[];if(modules.indexOf('Finance')===-1)return{ok:false,message:'Finance is not available for this account.'};if(write){var role=String(user.role||'staff').toLowerCase();if(user.isDemo==='YES')return{ok:false,message:'Demo accounts are view-only.'};if(role!=='owner'&&role!=='admin'&&role!=='staff')return{ok:false,message:'You have view-only access to Finance.'};}return{ok:true,user:user};}",
    "function getClientFinanceData(sid){var access=clientFinanceAccessV12_(sid,false);if(!access.ok)return{success:false,message:access.message};var result=getModuleDataWithSync('Finance',sid);if(!result||result.success===false)return result||{success:false,message:'Unable to load Finance.'};var records=Array.isArray(result.records)?result.records:[];var revenue=0,expenses=0;records.forEach(function(row){var type=String(row.Type||'').toLowerCase(),amount=Math.abs(Number(row.Amount)||0);if(type==='revenue'||type==='income')revenue+=amount;else if(type==='expense'||type==='expenses')expenses+=amount;});return{success:true,records:records.reverse(),summary:{revenue:revenue,expenses:expenses,profit:revenue-expenses,recordCount:records.length}};}",
    "function saveClientFinanceTransaction(data,sid){var access=clientFinanceAccessV12_(sid,true);if(!access.ok)return{success:false,message:access.message};data=data||{};data.Type=String(data.Type||'').toLowerCase();if(['revenue','expense'].indexOf(data.Type)===-1)return{success:false,message:'Transaction type must be income or expense.'};if(!(Number(data.Amount)>0))return{success:false,message:'Amount must be greater than zero.'};if(!String(data.Category||'').trim())return{success:false,message:'Category is required.'};return saveRecord('Finance',data,sid);}",
    "function deleteClientFinanceTransaction(recordId,sid){var access=clientFinanceAccessV12_(sid,true);if(!access.ok)return{success:false,message:access.message};recordId=String(recordId||'').trim();if(!recordId)return{success:false,message:'Transaction ID is required.'};var workspace=getWorkspaceFile(sid),sheet=workspace.getSheetByName(MODULES.Finance.sheet);if(!sheet||sheet.getLastRow()<2)return{success:false,message:'Transaction not found.'};var values=sheet.getDataRange().getValues(),headers=values[0],idCol=headers.indexOf('Transaction_ID');if(idCol<0)return{success:false,message:'Transaction ID column not found.'};for(var i=1;i<values.length;i++){if(String(values[i][idCol]||'')===recordId){sheet.deleteRow(i+1);return{success:true,message:'Transaction deleted successfully.'};}}return{success:false,message:'Transaction not found.'};}",
    "function clientSettingsAccessV12_(sid,write){var user=getUserFromSession(sid);if(!user)return{ok:false,message:'Session expired. Please login again.'};var role=String(user.role||'staff').toLowerCase();if(write&&['owner','admin'].indexOf(role)<0)return{ok:false,message:'Only the business owner or an administrator can change business settings.'};if(write&&user.isDemo==='YES')return{ok:false,message:'Demo accounts are view-only.'};return{ok:true,user:user};}",
    "function getClientAppsData(sid){var user=getUserFromSession(sid);if(!user)return{success:false,message:'Session expired. Please login again.'};return{success:true,summary:getModuleSummary(sid)};}",
    "function getClientSettingsData(sid){var access=clientSettingsAccessV12_(sid,false);if(!access.ok)return{success:false,message:access.message};var bp=(typeof CLIENT_CONFIG!=='undefined'&&CLIENT_CONFIG.businessProfile)||{};var settings={Business_Name:String(bp.businessName||CLIENT_CONFIG.clientName||''),Business_Type:String(bp.businessType||''),Phone:String(bp.ownerPhone||''),Website:String(bp.website||''),Address:String(bp.address||''),City:String(bp.city||''),Country:String(bp.country||''),Currency:String(bp.currency||'NGN'),Business_Description:String(bp.businessDescription||''),Display_Name:String(CLIENT_CONFIG.clientName||bp.businessName||''),Primary_Color:String(CLIENT_CONFIG.primaryColor||'#2E7D32'),Logo_Url:String(CLIENT_CONFIG.logoUrl||'')};var ws=getWorkspaceFile(sid),sheet=ws.getSheetByName('Client_Info');if(sheet&&sheet.getLastRow()>0){var rows=sheet.getDataRange().getValues(),h=rows[0]||[],p=h.indexOf('Property'),v=h.indexOf('Value');if(p>=0&&v>=0)for(var i=1;i<rows.length;i++){var k=String(rows[i][p]||'').trim(),value=rows[i][v];if(!k||value===undefined||value===null||String(value).trim()==='')continue;if(['Business_Name','Business_Type','Phone','Website','Address','City','Country','Currency','Business_Description','Display_Name','Primary_Color','Logo_Url'].indexOf(k)>=0)settings[k]=value;}}var u=access.user||{};u={email:String(u.email||bp.ownerEmail||CLIENT_CONFIG.email||''),name:String(u.name||bp.ownerName||''),phone:String(u.phone||bp.ownerPhone||''),role:String(u.role||'owner'),businessId:String(u.businessId||CLIENT_CONFIG.businessId),businessName:String(u.businessName||settings.Business_Name||CLIENT_CONFIG.clientName||'')};return{success:true,settings:settings,user:u,source:{canonical:'BizOS master business/user data',custom:'client workspace settings'}};}",
    "function saveClientSettings(data,sid){var access=clientSettingsAccessV12_(sid,true);if(!access.ok)return{success:false,message:access.message};data=data||{};var allowed=['Business_Name','Business_Type','Phone','Website','Address','City','Country','Currency','Business_Description','Display_Name','Primary_Color','Logo_Url'],clean={};allowed.forEach(function(k){if(data[k]!==undefined)clean[k]=String(data[k]||'').trim();});if(!clean.Business_Name)return{success:false,message:'Business name is required.'};if(clean.Currency&&!/^[A-Z]{3}$/.test(clean.Currency))return{success:false,message:'Currency must be a 3-letter code.'};if(clean.Primary_Color&&!/^#[0-9A-Fa-f]{6}$/.test(clean.Primary_Color))return{success:false,message:'Primary colour must be a 6-digit hex colour.'};var ws=getWorkspaceFile(sid),sheet=ws.getSheetByName('Client_Info');if(!sheet)return{success:false,message:'Client settings storage is unavailable.'};var rows=sheet.getDataRange().getValues(),h=rows[0]||[],p=h.indexOf('Property'),v=h.indexOf('Value'),idx={};if(p<0||v<0)return{success:false,message:'Client settings schema is unavailable.'};for(var i=1;i<rows.length;i++){var key=String(rows[i][p]||'').trim();if(key)idx[key]=i+1;}Object.keys(clean).forEach(function(k){if(idx[k])sheet.getRange(idx[k],v+1).setValue(clean[k]);else sheet.appendRow([k,clean[k]]);});return{success:true,settings:clean};}",
    "function updateClientProfile(data,sid){var access=clientSettingsAccessV12_(sid,false);if(!access.ok)return{success:false,message:access.message};data=data||{};var sheet=getOrCreateUserSheet(),rows=sheet.getDataRange().getValues(),h=rows[0]||[],emailCol=h.indexOf('Email'),nameCol=h.indexOf('Name'),phoneCol=h.indexOf('Phone'),roleCol=h.indexOf('Role'),businessIdCol=h.indexOf('Business_ID'),businessNameCol=h.indexOf('Business_Name'),statusCol=h.indexOf('Status'),clientCol=h.indexOf('Is_Client');if(emailCol<0)return{success:false,message:'User profile schema is unavailable.'};var wanted=String(access.user.email||'').toLowerCase(),row=-1;for(var i=1;i<rows.length;i++){if(String(rows[i][emailCol]||'').toLowerCase()===wanted){row=i+1;break;}}if(row<0){var newRow=[];for(var j=0;j<h.length;j++)newRow.push('');if(emailCol>=0)newRow[emailCol]=wanted;if(nameCol>=0)newRow[nameCol]=String(data.name||access.user.name||'');if(phoneCol>=0)newRow[phoneCol]=String(data.phone||access.user.phone||'');if(roleCol>=0)newRow[roleCol]=String(access.user.role||'owner');if(businessIdCol>=0)newRow[businessIdCol]=String(CLIENT_CONFIG.businessId||'');if(businessNameCol>=0)newRow[businessNameCol]=String(CLIENT_CONFIG.clientName||'');if(statusCol>=0)newRow[statusCol]='active';if(clientCol>=0)newRow[clientCol]='YES';sheet.appendRow(newRow);return{success:true};}if(nameCol>=0)sheet.getRange(row,nameCol+1).setValue(String(data.name||''));if(phoneCol>=0)sheet.getRange(row,phoneCol+1).setValue(String(data.phone||''));return{success:true};}",
    "function changeClientPassword(currentPassword,newPassword,sid){var access=clientSettingsAccessV12_(sid,false);if(!access.ok)return{success:false,message:access.message};return changeUserPassword(String(access.user.email||''),String(currentPassword||''),String(newPassword||''),sid);}"
  ].join(String.fromCharCode(10));
  runtime = runtime.replace(/\nfunction doGet\(e\)\{/, '\n'+businessCenterSource+'\n'+businessCenterBridge+'\nfunction doGet(e){');
  runtime += String.fromCharCode(10) + "function getClientModuleSummary(sid){var u=getUserFromSession(sid);if(!u)return{};var allowed={};var mods=Array.isArray(u.accessibleModules)?u.accessibleModules:[];mods.forEach(function(m){allowed[String(m)]=true;});var role=String(u.role||'staff').toLowerCase(),isAdmin=role==='owner'||role==='admin',out={},ws=null;try{ws=getWorkspaceFile(sid);}catch(e){ws=null;}Object.keys(MODULES).forEach(function(name){var def=MODULES[name]||{},count=0;try{var sh=ws&&ws.getSheetByName(def.sheet);count=sh&&sh.getLastRow()>1?sh.getLastRow()-1:0;}catch(e2){}out[name]={label:def.label||name,icon:def.icon||'',count:count,isAccessible:isAdmin||!!allowed[name],isReadOnly:false};});return out;}\nfunction getModuleSummary(sid){return getClientModuleSummary(sid);}" + String.fromCharCode(10);
  var doGetStart = runtime.indexOf('function doGet(e){');
  var includeStart = runtime.indexOf('function include', doGetStart);
  if (doGetStart < 0 || includeStart < 0) throw new Error('Unable to isolate the V12 client doGet runtime.');
  runtime = runtime.slice(0, doGetStart) +
    'function doGet(e){var t=HtmlService.createTemplateFromFile("ClientShellV12");return t.evaluate().setTitle(CLIENT_CONFIG.clientName+" - BizOS").addMetaTag("viewport","width=device-width,initial-scale=1").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}\n' +
    runtime.slice(includeStart);

  var files = [
    {name:'Code',type:'SERVER_JS',source:runtime},
    {name:'ClientShellV12',type:'HTML',source:buildClientShellV12_(config)},
    {name:'ClientBusinessCenterV12',type:'HTML',source:buildClientBusinessCenterV12_()},
    {name:'ClientFinanceV12',type:'HTML',source:buildClientFinanceV12_()},
    {name:'ClientStaffV12',type:'HTML',source:buildClientStaffV12_()},
    {name:'ClientAppsV12',type:'HTML',source:buildClientAppsV12_()},
    {name:'ClientSettingsV12',type:'HTML',source:buildClientSettingsV12_()},
    {name:'ClientStylesV12',type:'HTML',source:buildClientStylesV12_()},
    {name:'appsscript',type:'JSON',source:JSON.stringify(buildClientManifestV12_())}
  ];

  return {
    success:true,version:'12.0',config:config,files:files,
    source:{clientUi:'V12',copiedMasterUi:[],reusedBackendRuntime:'V11-safe-runtime',businessCenterBackend:'BusinessHubGS-safe-wrapped',businessCenterUi:'ClientBusinessCenterV12'}
  };
}

function getMasterBusinessProfileForV12_(businessId, email) {
  var out = {
    businessId:String(businessId || ''),
    businessName:'',
    ownerEmail:String(email || '').toLowerCase(),
    ownerName:'',
    ownerPhone:'',
    address:'',
    city:'',
    country:'',
    website:'',
    businessType:'',
    currency:'',
    businessDescription:''
  };
  try {
    var ss = getBizOSMasterSpreadsheet_();
    var businessSheet = ss.getSheetByName('Businesses');
    if (businessSheet && businessSheet.getLastRow() > 1) {
      var data = businessSheet.getDataRange().getValues();
      var h = data[0] || [];
      var idCol=h.indexOf('Business_ID'), nameCol=h.indexOf('Business_Name'), emailCol=h.indexOf('Owner_Email'),
          phoneCol=h.indexOf('Owner_Phone'), addressCol=h.indexOf('Address'), cityCol=h.indexOf('City'),
          countryCol=h.indexOf('Country'), websiteCol=h.indexOf('Website'), tierCol=h.indexOf('Subscription_Tier');
      var wantedId=String(businessId || '').trim(), wantedEmail=String(email || '').trim().toLowerCase(), row=null;
      for(var i=1;i<data.length;i++){
        var rid=idCol>=0?String(data[i][idCol]||'').trim():'';
        var remail=emailCol>=0?String(data[i][emailCol]||'').trim().toLowerCase():'';
        if((wantedId && rid===wantedId) || (!wantedId && wantedEmail && remail===wantedEmail)){ row=data[i]; break; }
      }
      if(row){
        out.businessId=idCol>=0?String(row[idCol]||out.businessId):out.businessId;
        out.businessName=nameCol>=0?String(row[nameCol]||''):''; 
        out.ownerEmail=emailCol>=0?String(row[emailCol]||out.ownerEmail).toLowerCase():out.ownerEmail;
        out.ownerPhone=phoneCol>=0?String(row[phoneCol]||''):'';
        out.address=addressCol>=0?String(row[addressCol]||''):'';
        out.city=cityCol>=0?String(row[cityCol]||''):'';
        out.country=countryCol>=0?String(row[countryCol]||''):'';
        out.website=websiteCol>=0?String(row[websiteCol]||''):'';
        out.subscriptionTier=tierCol>=0?String(row[tierCol]||''):'';
      }
    }
    var userSheet = ss.getSheetByName('Users');
    if(!userSheet) userSheet = ss.getSheetByName('User_Data');
    if(userSheet && userSheet.getLastRow() > 1){
      var uv=userSheet.getDataRange().getValues(), uh=uv[0]||[];
      var ue=uh.indexOf('Email'), ub=uh.indexOf('Business_ID'), un=uh.indexOf('Name'), up=uh.indexOf('Phone');
      var wantedUserEmail=String(out.ownerEmail||email||'').trim().toLowerCase(), urow=null;
      for(var j=1;j<uv.length;j++){
        var em=ue>=0?String(uv[j][ue]||'').trim().toLowerCase():'';
        var bid=ub>=0?String(uv[j][ub]||'').trim():'';
        if((wantedUserEmail && em===wantedUserEmail) && (!businessId || !ub || bid===String(businessId))){urow=uv[j];break;}
      }
      if(urow){
        if(un>=0) out.ownerName=String(urow[un]||'');
        if(up>=0 && !out.ownerPhone) out.ownerPhone=String(urow[up]||'');
        if(ue>=0 && !out.ownerEmail) out.ownerEmail=String(urow[ue]||'').toLowerCase();
      }
    }
  } catch(e) {
    console.warn('getMasterBusinessProfileForV12_ fallback:', e && e.message ? e.message : e);
  }
  return out;
}

function buildClientShellV12_(config) {
  var source = String(getMasterSourceForV12_('ClientShellV12') || '');
  if(!source) throw new Error('ClientShellV12 source is missing.');
  var clientConfig = JSON.stringify({
    isClient:true,id:config.clientId,name:config.clientName,primaryColor:config.primaryColor,logoUrl:config.logoUrl
  });
  return source.replace(/<head>/i,'<head><script>window.__CLIENT='+clientConfig+';</script>');
}

function buildClientBusinessCenterV12_() {
  var source = String(getMasterSourceForV12_('ClientBusinessCenterV12') || '');
  if(!source) throw new Error('ClientBusinessCenterV12 source is missing.');
  return source;
}

function buildClientFinanceV12_() {
  var source = String(getMasterSourceForV12_('ClientFinanceV12') || '');
  if(!source) throw new Error('ClientFinanceV12 source is missing.');
  return source;
}

function buildClientAppsV12_() { var source=String(getMasterSourceForV12_('ClientAppsV12')||''); if(!source) throw new Error('ClientAppsV12 source is missing.'); return source; }
function buildClientSettingsV12_() { var source=String(getMasterSourceForV12_('ClientSettingsV12')||''); if(!source) throw new Error('ClientSettingsV12 source is missing.'); return source; }
function buildClientStaffV12_() {
  var source = String(getMasterSourceForV12_('ClientStaffV12') || '');
  if(!source) throw new Error('ClientStaffV12 source is missing.');
  return source;
}

function buildClientStylesV12_() {
  var source = String(getMasterSourceForV12_('ClientStylesV12') || '');
  if(!source) throw new Error('ClientStylesV12 source is missing.');
  return source;
}

function getMasterSourceForV12_(fileName) {
  var files = fetchMasterProjectContentV11_();
  for(var i=0;i<files.length;i++) if(files[i] && String(files[i].name||'').replace(/\.(?:html|js|gs)$/i,'')===String(fileName)) return files[i].source||'';
  return '';
}

function buildClientManifestV12_() {
  return {
    timeZone:'Africa/Lagos',
    exceptionLogging:'STACKDRIVER',
    runtimeVersion:'V8',
    webapp:{executeAs:'USER_DEPLOYING',access:'ANYONE_ANONYMOUS'},
    oauthScopes:[
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive',
      'https://www.googleapis.com/auth/script.external_request',
      'https://www.googleapis.com/auth/script.send_mail'
    ]
  };
}

// ============================================================
// MASTER-SIDE NON-DEPLOY SMOKE TEST
// ============================================================
// Run this from the Master BizOS Apps Script project.
// It reads the active client from the Clients registry, generates
// and validates the V12 package, and makes NO client deployment,
// version, spreadsheet, or registry changes.
// Optional businessId lets you test a specific active client.
// With no argument it finds the active client by email.
// ============================================================

function testGenerateClientCodeSafelyV12(businessId) {
  var targetBusinessId = String(businessId || '').trim();
  var client = targetBusinessId
    ? getActiveClientByBusinessIdForRedeploy_(targetBusinessId)
    : getActiveClientByEmailForV12Test_('bizoshigroups@gmail.com');

  if (!client) {
    throw new Error(targetBusinessId
      ? 'No active client deployment was found for business ID: ' + targetBusinessId
      : 'No active client deployment was found for bizoshigroups@gmail.com.');
  }

  var settings = {
    clientId:client.clientId,
    clientName:client.clientName,
    primaryColor:client.primaryColor || '#2E7D32',
    logoUrl:client.logoUrl || '',
    email:client.email,
    sheetId:client.sheetId,
    businessId:client.businessId,
    masterApiUrl:getMasterApiUrl(),
    landingUrl:client.landingUrl || '',
    customDomain:client.customDomain || '',
    applicationUrl:client.webAppUrl || ''
  };

  var packageResult = generateClientCodeSafelyV12(settings);
  var packageCheck = validateClientDeploymentPackageV12_(packageResult,{
    clientId:client.clientId,
    sheetId:client.sheetId
  });

  var summary = {
    success:!!(packageResult && packageResult.success && packageCheck && packageCheck.success),
    test:'V12_PACKAGE_GENERATION_ONLY',
    deploymentChanged:false,
    clientRegistryChanged:false,
    clientSpreadsheetChanged:false,
    clientId:client.clientId,
    businessId:client.businessId,
    email:client.email,
    scriptId:client.scriptId,
    deploymentId:client.deploymentId,
    packageVersion:packageResult.version,
    packageCheck:packageCheck,
    files:(packageResult.files || []).map(function(file){
      return {name:String(file.name || ''),type:String(file.type || ''),sourceLength:String(file.source || '').length};
    })
  };

  console.log('V12 NON-DEPLOY SMOKE TEST RESULT:',JSON.stringify(summary));
  return summary;
}

function getActiveClientByEmailForV12Test_(email) {
  var ss = getBizOSMasterSpreadsheet_();
  var sheet = ss.getSheetByName('Clients');
  if (!sheet || sheet.getLastRow() < 2) return null;

  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function(h){ return String(h || '').trim(); });
  var emailCol = headers.indexOf('Email');
  var statusCol = headers.indexOf('Status');

  if (emailCol < 0) throw new Error('Client registry is missing Email.');
  if (statusCol < 0) throw new Error('Client registry is missing Status.');

  var wanted = String(email || '').trim().toLowerCase();

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (String(row[emailCol] || '').trim().toLowerCase() !== wanted) continue;
    if (String(row[statusCol] || '').trim().toLowerCase() !== 'active') continue;

    var client = {};
    headers.forEach(function(header,index){ client[header] = row[index] === undefined ? '' : row[index]; });

    return {
      clientId:String(client.Client_ID || ''),
      email:String(client.Email || '').toLowerCase(),
      clientName:String(client.Client_Name || ''),
      businessId:String(client.Business_ID || ''),
      sheetId:String(client.Sheet_ID || client.Workspace_ID || ''),
      scriptId:String(client.Script_ID || ''),
      deploymentId:String(client.Deployment_ID || ''),
      webAppUrl:String(client.Web_App_URL || ''),
      primaryColor:String(client.Primary_Color || ''),
      logoUrl:String(client.Logo_Url || ''),
      landingUrl:String(client.Landing_URL || ''),
      customDomain:String(client.Custom_Domain || ''),
      domain:String(client.Domain || '')
    };
  }

  return null;
}

function validateClientDeploymentPackageV12_(pkg, expected) {
  if(!pkg||!Array.isArray(pkg.files)||!pkg.files.length)return{success:false,code:'PACKAGE_EMPTY',message:'Generated V12 client package is empty.'};
  var names=pkg.files.map(function(f){return String(f.name||'');});
  var required=['Code','ClientShellV12','ClientBusinessCenterV12','ClientFinanceV12','ClientStaffV12','ClientAppsV12','ClientSettingsV12','ClientStylesV12','appsscript'];
  var missing=required.filter(function(n){return names.indexOf(n)<0;});
  if(missing.length)return{success:false,code:'V12_PACKAGE_MISSING_FILES',message:'V12 package is missing: '+missing.join(', ')};
  var shell=pkg.files.filter(function(f){return f.name==='ClientShellV12';})[0].source||'';
  var business=pkg.files.filter(function(f){return f.name==='ClientBusinessCenterV12';})[0].source||'';
  var finance=pkg.files.filter(function(f){return f.name==='ClientFinanceV12';})[0].source||'';
  var staff=pkg.files.filter(function(f){return f.name==='ClientStaffV12';})[0].source||'';
  var apps=pkg.files.filter(function(f){return f.name==='ClientAppsV12';})[0].source||'';
  var settings=pkg.files.filter(function(f){return f.name==='ClientSettingsV12';})[0].source||'';
  var styles=pkg.files.filter(function(f){return f.name==='ClientStylesV12';})[0].source||'';
  var code=pkg.files.filter(function(f){return f.name==='Code';})[0].source||'';
  var manifestSource=pkg.files.filter(function(f){return f.name==='appsscript';})[0].source||'';
  var forbiddenUi=['id="landing-page"','id="dashboard-page"','id="admin-dashboard"','Admin Login','Client Registry','Deployment Status'];
  var leakedUi=forbiddenUi.filter(function(token){return shell.indexOf(token)>=0;});
  if(leakedUi.length)return{success:false,code:'V12_MASTER_UI_LEAK',message:'V12 client UI contains Master UI markers: '+leakedUi.join(', ')};
  var forbiddenRuntime=['function autoSetupClient','function redeployClientByBusinessId','function requireAdminSession_','function getBizOSMasterSpreadsheet_','function createAndDeployClientScriptSafe'];
  var leakedRuntime=forbiddenRuntime.filter(function(token){return code.indexOf(token)>=0;});
  if(leakedRuntime.length)return{success:false,code:'V12_MASTER_RUNTIME_LEAK',message:'V12 client runtime contains forbidden Master functions: '+leakedRuntime.join(', ')};
  var requiredRuntime=['function doGet(e){','function clientFinanceAccessV12_','function getClientAppsData','function clientSettingsAccessV12_','function getClientSettingsData','function saveClientSettings','function changeClientPassword','function updateClientProfile','function getClientFinanceData','function saveClientFinanceTransaction','function deleteClientFinanceTransaction','function authenticateClient','function validateSession','function getUserFromSession','function clientLogout','function getClientDashboardData','function getClientDashboard(','function requireClientSession_','function getBusinessCenterData','function getClientBusinessCenterData','function saveClientBusinessCenterProduct','function saveClientBusinessCenterCustomer','function completeClientBusinessCenterSale','function adjustClientBusinessCenterStock'];
  var missingRuntime=requiredRuntime.filter(function(token){return code.indexOf(token)<0;});
  if(missingRuntime.length)return{success:false,code:'V12_RUNTIME_DEPENDENCY_MISSING',message:'V12 client runtime is missing required functions: '+missingRuntime.join(', ')};var securityRuntime=['function clientAssignedModulesV11_','CLIENT_LICENSE_INVALID','function removeTeamMember(bid,email,sid)','String(u.businessId)!==String(bid)','function updateUserProfile(bid,name,phone,sid)','function changeUserPassword(email,currentPassword,newPassword,sid)','function createClientSessionV11_'];var missingSecurity=securityRuntime.filter(function(token){return code.indexOf(token)<0;});if(missingSecurity.length)return{success:false,code:'V12_SECURITY_DEPENDENCY_MISSING',message:'V12 client runtime is missing hardened security markers: '+missingSecurity.join(', ')};if(code.indexOf('function createSession(')>=0)return{success:false,code:'V12_PUBLIC_SESSION_FACTORY_LEAK',message:'V12 client runtime still exposes the public session factory.'};
  var requiredShell=['ClientStylesV12','ClientStaffV12','ClientAppsV12','ClientSettingsV12','ClientBusinessCenterV12','authenticateClient','validateSession','getUserFromSession','clientLogout','getClientDashboardData','rich-kpi-grid','trendChart','breakdownChart','healthView'];
  var requiredStaff=['staff-v12-root','getBusinessStaff','inviteStaffMember','resendStaffInvitation','removeTeamMember','assignModuleToStaff','getAvailableModulesForStaff'];
  var missingStaff=requiredStaff.filter(function(token){return staff.indexOf(token)<0;});
  if(missingStaff.length)return{success:false,code:'V12_STAFF_DEPENDENCY_MISSING',message:'V12 Staff workspace is missing required markers: '+missingStaff.join(', ')};
  var requiredApps=['apps-v12-root','getClientAppsData'];var missingApps=requiredApps.filter(function(token){return apps.indexOf(token)<0;});if(missingApps.length)return{success:false,code:'V12_APPS_DEPENDENCY_MISSING',message:'V12 Apps workspace is missing required markers: '+missingApps.join(', ')};var requiredSettings=['settings-v12-root','getClientSettingsData','saveClientSettings','changeClientPassword'];var missingSettings=requiredSettings.filter(function(token){return settings.indexOf(token)<0;});if(missingSettings.length)return{success:false,code:'V12_SETTINGS_DEPENDENCY_MISSING',message:'V12 Settings workspace is missing required markers: '+missingSettings.join(', ')};var requiredBusiness=['bc-v12-root','getClientBusinessCenterData','saveClientBusinessCenterProduct','saveClientBusinessCenterCustomer','completeClientBusinessCenterSale','adjustClientBusinessCenterStock'];
  var missingBusiness=requiredBusiness.filter(function(token){return business.indexOf(token)<0;});
  if(missingBusiness.length)return{success:false,code:'V12_BUSINESS_CENTER_DEPENDENCY_MISSING',message:'V12 Business Center UI is missing required markers: '+missingBusiness.join(', ')};
  var missingShell=requiredShell.filter(function(token){return shell.indexOf(token)<0;});
  if(missingShell.length)return{success:false,code:'V12_SHELL_DEPENDENCY_MISSING',message:'V12 client shell is missing required UI/runtime markers: '+missingShell.join(', ')};
  var requiredStyles=['rich-kpi-grid','.bar-chart','.health-score','.mini-table'];
  var missingStyles=requiredStyles.filter(function(token){return styles.indexOf(token)<0;});
  if(missingStyles.length)return{success:false,code:'V12_STYLE_DEPENDENCY_MISSING',message:'V12 client styles are missing required dashboard styles: '+missingStyles.join(', ')};
  try{var manifest=JSON.parse(manifestSource);if(!manifest||manifest.runtimeVersion!=='V8'||!manifest.webapp)return{success:false,code:'V12_MANIFEST_INVALID',message:'V12 manifest is missing required V8/webapp configuration.'};}catch(e){return{success:false,code:'V12_MANIFEST_INVALID',message:'V12 manifest is not valid JSON: '+e.message};}
  if(expected&&expected.clientId&&code.indexOf(String(expected.clientId))<0)return{success:false,code:'V12_CLIENT_ID_MISMATCH',message:'V12 runtime does not contain the expected client ID.'};
  if(expected&&expected.sheetId&&code.indexOf(String(expected.sheetId))<0)return{success:false,code:'V12_WORKSPACE_ID_MISMATCH',message:'V12 runtime does not contain the expected workspace ID.'};
  return{success:true,code:'V12_PACKAGE_VALID',version:'12.0',fileCount:pkg.files.length,files:names,checks:{runtimeDependencies:true,shellDependencies:true,businessCenterDependencies:true,staffDependencies:true,appsDependencies:true,settingsDependencies:true,styleDependencies:true,manifest:true,masterUiLeak:false,masterRuntimeLeak:false}};
}
