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
  // F security hardening: internal helpers must not be exposed through google.script.run.
  runtime = runtime
    .replace(/\bgetOrCreateUserSheet\(/g,'getOrCreateUserSheet_(')
    .replace(/\bgetOrCreateBusinessSheet\(/g,'getOrCreateBusinessSheet_(')
    .replace(/\bgetBusinessStaff\(/g,'getBusinessStaffLegacy_(')
    .replace(/\bresendStaffInvitation\(/g,'resendStaffInvitationLegacy_(')
    .replace(/\bassignModuleToStaff\(/g,'assignModuleToStaffLegacy_(')
    .replace(/\bremoveModuleFromStaff\(/g,'removeModuleFromStaffLegacy_(')
    .replace(/\bgetClientModuleSummary\(/g,'getClientModuleSummaryLegacy_(')
    .replace(/\bremoveTeamMember\(/g,'removeTeamMemberLegacy_(');
  runtime = runtime.replace(
    'accessibleModules:Object.keys(MODULES)',
    'accessibleModules:clientModulesForEntitlementV12_(x.user&&x.user.role,x.user&&x.user.subscriptionTier)'
  );
  runtime = runtime.replace(
    'subscriptionTier:"sovereign",workspaceId:CLIENT_CONFIG.sheetId,isDemo:false,isClient:true,accessibleModules:clientAssignedModulesV11_(t)',
    'subscriptionTier:clientLicenseTierV12_(),workspaceId:CLIENT_CONFIG.sheetId,isDemo:false,isClient:true,accessibleModules:clientModulesForEntitlementV12_(t.Role,clientLicenseTierV12_(),clientAssignedModulesV11_(t))'
  );
  runtime = runtime.replace(
    'var moduleCounts={},moduleNames=Array.isArray(u.accessibleModules)?u.accessibleModules.slice():[];',
    'var moduleCounts={},moduleNames=Array.isArray(u.accessibleModules)?u.accessibleModules.slice():[];'
  );

  var v12SecurityBridge = [
    'function clientModulesForEntitlementV12_(role,tier,assigned){var r=String(role||"staff").toLowerCase(),t=String(tier||"").toLowerCase(),base=(t==="free"||t==="starter")?["Ecommerce"]:Object.keys(MODULES);if(r==="owner"||r==="admin")return base.slice();var a=Array.isArray(assigned)?assigned:[];return a.filter(function(m){return base.indexOf(String(m||""))!==-1;});}',
    'function clientLicenseTierV12_(){try{var r=validateClientLicenseV11_({clientId:CLIENT_CONFIG.clientId,businessId:CLIENT_CONFIG.businessId,sheetId:CLIENT_CONFIG.sheetId});return r&&r.success&&r.client?String(r.client.tier||""):"";}catch(e){return"";}}',
    'function getBusinessStaff(bid,sid){var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(bid)||["owner","admin"].indexOf(String(u.role||"").toLowerCase())<0)return{success:false,message:"Only the business owner or an administrator can view staff."};return{success:true,members:clientTeamRowsV11_(ensureClientTeamSheetV11_()).map(function(r){return{email:r.Email,name:r.Name,role:r.Role,phone:r.Phone,department:r.Department,status:r.Status,assignedModules:clientAssignedModulesV11_(r),invitationStatus:r.Status};})};}',
    'function resendStaffInvitation(bid,email,sid){try{var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(bid)||["owner","admin"].indexOf(String(u.role||"").toLowerCase())<0)return{success:false,message:"Only the business owner or an administrator can resend invitations."};email=String(email||"").trim().toLowerCase();var s=ensureClientTeamSheetV11_(),rows=clientTeamRowsV11_(s),t=rows.find(function(r){return String(r.Email||"").toLowerCase()===email;});if(!t)return{success:false,message:"Team member invitation not found."};if(String(t.Status||"").toLowerCase()==="accepted")return{success:false,message:"This team member has already accepted the invitation."};var pass=Utilities.getUuid().replace(/-/g,"").slice(0,12)+"A!",v=s.getDataRange().getValues(),h=v[0]||[],emailCol=h.indexOf("Email"),hashCol=h.indexOf("Password_Hash"),mustCol=h.indexOf("Must_Change_Password"),statusCol=h.indexOf("Status"),updatedCol=h.indexOf("Updated_At"),row=-1;for(var i=1;i<v.length;i++){if(String(v[i][emailCol]||"").toLowerCase()===email){row=i+1;break;}}if(row<0)return{success:false,message:"Team member invitation not found."};if(hashCol>=0)s.getRange(row,hashCol+1).setValue(clientHashV11_(pass));if(mustCol>=0)s.getRange(row,mustCol+1).setValue("YES");if(statusCol>=0)s.getRange(row,statusCol+1).setValue("invited");if(updatedCol>=0)s.getRange(row,updatedCol+1).setValue(new Date().toISOString());MailApp.sendEmail({to:email,subject:"Your BizOS invitation to "+CLIENT_CONFIG.clientName,htmlBody:"<p>Hello <strong>"+String(t.Name||email)+"</strong>,</p><p>Your BizOS invitation has been resent for <strong>"+CLIENT_CONFIG.clientName+"</strong>.</p><p><strong>Email:</strong> "+email+"<br><strong>Temporary password:</strong> "+pass+"</p><p>Please sign in and change your password from your profile.</p>"});return{success:true,message:"Invitation resent to "+email+"."};}catch(e){return{success:false,message:e.message,code:"STAFF_RESEND_ERROR"};}}',
    'function assignModuleToStaff(email,bid,moduleName,assignedBy,sid){try{var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(bid)||["owner","admin"].indexOf(String(u.role||"").toLowerCase())<0)return{success:false,message:"Only the business owner or an administrator can assign modules."};var mod=String(moduleName||"").trim(),tier=clientLicenseTierV12_(),allowed=(String(tier||"").toLowerCase()==="free"||String(tier||"").toLowerCase()==="starter")?["Ecommerce"]:Object.keys(MODULES);if(allowed.indexOf(mod)<0)return{success:false,message:"This module is not available for the current client plan."};var s=ensureClientTeamSheetV11_(),v=s.getDataRange().getValues(),h=v[0]||[],emailCol=h.indexOf("Email"),modulesCol=h.indexOf("Assigned_Modules"),row=-1;for(var i=1;i<v.length;i++){if(String(v[i][emailCol]||"").toLowerCase()===String(email||"").toLowerCase()){row=i+1;break;}}if(row<0)return{success:false,message:"Team member not found."};var raw=modulesCol>=0?String(s.getRange(row,modulesCol+1).getValue()||""):"",mods=[];try{mods=raw?JSON.parse(raw):[];}catch(e){mods=raw.split(",").map(function(x){return String(x||"").trim();}).filter(Boolean);}if(mods.indexOf(mod)<0)mods.push(mod);if(modulesCol<0){s.getRange(1,s.getLastColumn()+1).setValue("Assigned_Modules");modulesCol=s.getLastColumn()-1;}s.getRange(row,modulesCol+1).setValue(JSON.stringify(mods));return{success:true,message:"Module access updated."};}catch(e){return{success:false,message:e.message,code:"STAFF_MODULE_ASSIGN_ERROR"};}}',
    'function removeModuleFromStaff(email,bid,moduleName,sid){try{var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(bid)||["owner","admin"].indexOf(String(u.role||"").toLowerCase())<0)return{success:false,message:"Only the business owner or an administrator can remove modules."};var mod=String(moduleName||"").trim(),s=ensureClientTeamSheetV11_(),v=s.getDataRange().getValues(),h=v[0]||[],emailCol=h.indexOf("Email"),modulesCol=h.indexOf("Assigned_Modules"),row=-1;for(var i=1;i<v.length;i++){if(String(v[i][emailCol]||"").toLowerCase()===String(email||"").toLowerCase()){row=i+1;break;}}if(row<0)return{success:false,message:"Team member not found."};var raw=modulesCol>=0?String(s.getRange(row,modulesCol+1).getValue()||""):"",mods=[];try{mods=raw?JSON.parse(raw):[];}catch(e){mods=raw.split(",").map(function(x){return String(x||"").trim();}).filter(Boolean);}mods=mods.filter(function(x){return x!==mod;});if(modulesCol>=0)s.getRange(row,modulesCol+1).setValue(JSON.stringify(mods));return{success:true,message:"Module access updated."};}catch(e){return{success:false,message:e.message,code:"STAFF_MODULE_REMOVE_ERROR"};}}',
        'function removeTeamMember(bid,email,sid){try{var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(bid)||["owner","admin"].indexOf(String(u.role||"").toLowerCase())<0)return{success:false,message:"Unauthorized"};var s=ensureClientTeamSheetV11_(),rows=clientTeamRowsV11_(s),t=rows.find(function(r){return String(r.Email||"").toLowerCase()===String(email||"").toLowerCase();});if(!t)return{success:false,message:"Team member not found."};s.deleteRow(t.row);return{success:true,message:"Team member removed."};}catch(e){return{success:false,message:e.message};}}',
'function getClientModuleSummary(sid){var u=getUserFromSession(sid);if(!u)return{};var allowed={};(Array.isArray(u.accessibleModules)?u.accessibleModules:[]).forEach(function(m){allowed[String(m)]=true;});var out={},ws=null;try{ws=getWorkspaceFile(sid);}catch(e){ws=null;}Object.keys(MODULES).forEach(function(name){var def=MODULES[name]||{},count=0;try{var sh=ws&&ws.getSheetByName(def.sheet);count=sh&&sh.getLastRow()>1?sh.getLastRow()-1:0;}catch(e2){}out[name]={label:def.label||name,icon:def.icon||"",count:count,isAccessible:!!allowed[name],isReadOnly:false};});return out;}'
  ].join(String.fromCharCode(10));
  runtime += String.fromCharCode(10) + v12SecurityBridge + String.fromCharCode(10);
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
  // V12 security bridge provides the authoritative client module summary.
  runtime += String.fromCharCode(10) + "function getClientDashboard(sid,forceRefresh){\n  try{\n    var u=getUserFromSession(sid);\n    if(!u)return{success:false,message:\"Session expired.\",code:\"SESSION_EXPIRED\"};\n    var cacheKey=\"BIZOS_CLIENT_DASH_\"+String(CLIENT_CONFIG.clientId).replace(/[^A-Za-z0-9_]/g,\"_\");\n    if(!forceRefresh){\n      var cached=CacheService.getScriptCache().get(cacheKey);\n      if(cached){try{return{success:true,dashboard:JSON.parse(cached),cached:true};}catch(ignore){}}\n    }\n    var ss=getWorkspaceFile(sid),finance=ss.getSheetByName(\"Financial_Data\");\n    var revenue=0,expenses=0,recent=[],daily={},expenseCategories={},financeCount=0;\n    if(finance&&finance.getLastRow()>1){\n      var last=finance.getLastRow(),lastCol=finance.getLastColumn(),headers=finance.getRange(1,1,1,lastCol).getValues()[0]||[];\n      var tc=headers.indexOf(\"Type\"),ac=headers.indexOf(\"Amount\"),dc=headers.indexOf(\"Date\"),xc=headers.indexOf(\"Description\"),cc=headers.indexOf(\"Category\");\n      if(tc<0||ac<0)throw new Error(\"Financial_Data schema is missing Type or Amount.\");\n      financeCount=last-1;\n      var needed=[dc,tc,ac,xc,cc].filter(function(i){return i>=0;});\n      var minCol=Math.min.apply(null,needed),maxCol=Math.max.apply(null,needed),width=maxCol-minCol+1;\n      var rows=finance.getRange(2,minCol+1,financeCount,width).getValues();\n      rows.forEach(function(r){\n        var get=function(index){return index>=minCol?r[index-minCol]:'';};\n        var a=Number(get(ac))||0,t=String(get(tc)||\"\").toLowerCase(),d=get(dc),key=\"\";\n        if(d instanceof Date)key=Utilities.formatDate(d,Session.getScriptTimeZone(),\"yyyy-MM-dd\");else if(d)key=String(d).slice(0,10);\n        if(t===\"revenue\"||t===\"income\"){revenue+=Math.abs(a);if(key){daily[key]=daily[key]||{revenue:0,expenses:0};daily[key].revenue+=Math.abs(a);}}\n        if(t===\"expense\"||t===\"expenses\"){expenses+=Math.abs(a);if(key){daily[key]=daily[key]||{revenue:0,expenses:0};daily[key].expenses+=Math.abs(a);}var cat=String(get(cc)||\"Uncategorized\")||\"Uncategorized\";expenseCategories[cat]=(expenseCategories[cat]||0)+Math.abs(a);}\n      });\n      rows.slice(-10).reverse().forEach(function(r){\n        var get=function(index){return index>=minCol?r[index-minCol]:'';};\n        if(dc>=0&&get(dc))recent.push({date:get(dc),description:xc>=0?get(xc):\"\",amount:ac>=0?get(ac):0,type:tc>=0?get(tc):\"\",category:cc>=0?get(cc):\"\"});\n      });\n    }\n    var moduleCounts={},moduleNames=Array.isArray(u.accessibleModules)?u.accessibleModules.slice():[];\n    moduleNames.forEach(function(name){try{var sh=ss.getSheetByName(MODULES[name].sheet);moduleCounts[name]=sh?Math.max(0,sh.getLastRow()-1):0;}catch(e3){moduleCounts[name]=0;}});\n    var teamSheet=ss.getSheetByName(\"Client_Team\"),teamCount=teamSheet&&teamSheet.getLastRow()>1?teamSheet.getLastRow()-1:0;\n    var profit=revenue-expenses,margin=revenue?profit/revenue*100:0,trendKeys=Object.keys(daily).sort().slice(-7),trend={labels:[],revenue:[],expenses:[]};\n    trendKeys.forEach(function(k){trend.labels.push(k);trend.revenue.push(daily[k].revenue||0);trend.expenses.push(daily[k].expenses||0);});\n    var expPairs=Object.keys(expenseCategories).sort(function(a,b){return expenseCategories[b]-expenseCategories[a];}).slice(0,6),breakdown={labels:expPairs,values:expPairs.map(function(k){return expenseCategories[k];})};\n    var alerts=[],recommendations=[];\n    if(financeCount===0)recommendations.push({type:\"start\",title:\"Start tracking your finances\",message:\"Add your first income or expense to unlock financial insights.\"});\n    if(financeCount>0&&revenue===0)recommendations.push({type:\"info\",title:\"Record your income\",message:\"Add revenue entries so BizOS can calculate profit and business trends.\"});\n    if(revenue>0&&expenses>revenue)alerts.push({type:\"warning\",title:\"Expenses are above revenue\",message:\"Your recorded expenses currently exceed recorded revenue.\"});\n    var health=financeCount>0?{overall:Math.round(Math.max(0,Math.min(100,50+margin))),metrics:[{label:\"Profit margin\",value:margin.toFixed(1)+\"%\"},{label:\"Financial records\",value:String(financeCount)},{label:\"Team members\",value:String(teamCount)}],recommendations:recommendations}:{overall:null,metrics:[],recommendations:recommendations};\n    var dashboard={kpis:{revenue:revenue,expenses:expenses,profit:profit,margin:margin},modules:moduleNames,moduleCounts:moduleCounts,counts:{finance:financeCount,products:(moduleCounts.Ecommerce||0)+(moduleCounts.POS||0),modules:moduleNames.length,team:teamCount},charts:{revenueVsExpenses:trend,expenseBreakdown:breakdown},health:health,alerts:alerts,recommendations:recommendations,recentActivity:recent.slice(0,8).map(function(x){return{date:x.date,title:x.description||x.type||\"Financial entry\",type:String(x.type||\"\").toLowerCase(),amount:x.amount};}),recentTransactions:recent,clientName:CLIENT_CONFIG.clientName,primaryColor:CLIENT_CONFIG.primaryColor,logoUrl:CLIENT_CONFIG.logoUrl,workspaceId:CLIENT_CONFIG.sheetId,userEmail:u.email,generatedAt:new Date().toISOString()};\n    try{CacheService.getScriptCache().put(cacheKey,JSON.stringify(dashboard),60);}catch(ignore2){}\n    return{success:true,dashboard:dashboard,cached:false};\n  }catch(e){return{success:false,message:e.message,code:\"CLIENT_DASHBOARD_ERROR\"};}\n}" + String.fromCharCode(10);
  // F security bridge is appended before the V12 doGet isolation so its functions are part of the generated Code package.
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
  // F security markers: internal helpers are private and staff actions are server-authorized.

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
  if(missingRuntime.length)return{success:false,code:'V12_RUNTIME_DEPENDENCY_MISSING',message:'V12 client runtime is missing required functions: '+missingRuntime.join(', ')};var securityRuntime=['function clientAssignedModulesV11_','CLIENT_LICENSE_INVALID','function removeTeamMemberLegacy_(bid,email,sid)','function getBusinessStaff(bid,sid)','function resendStaffInvitation(bid,email,sid)','function assignModuleToStaff(email,bid,moduleName,assignedBy,sid)','function removeModuleFromStaff(email,bid,moduleName,sid)','String(u.businessId)!==String(bid)','function updateUserProfile(bid,name,phone,sid)','function changeUserPassword(email,currentPassword,newPassword,sid)','function createClientSessionV11_','function getOrCreateUserSheet_','function getOrCreateBusinessSheet_'];var missingSecurity=securityRuntime.filter(function(token){return code.indexOf(token)<0;});if(missingSecurity.length)return{success:false,code:'V12_SECURITY_DEPENDENCY_MISSING',message:'V12 client runtime is missing hardened security markers: '+missingSecurity.join(', ')};if(code.indexOf('function createSession(')>=0)return{success:false,code:'V12_PUBLIC_SESSION_FACTORY_LEAK',message:'V12 client runtime still exposes the public session factory.'};var forbiddenPublic=['function getOrCreateUserSheet(','function getOrCreateBusinessSheet(','function getBusinessStaffLegacy_(','function resendStaffInvitationLegacy_(','function assignModuleToStaffLegacy_(','function removeModuleFromStaffLegacy_(','function getClientModuleSummaryLegacy_('];var leakedPublic=forbiddenPublic.filter(function(token){return code.indexOf(token)>=0;});if(leakedPublic.length)return{success:false,code:'V12_INTERNAL_HELPER_EXPOSED',message:'V12 client runtime still exposes internal helper functions: '+leakedPublic.join(', ')};var requiredEntitlement=['function clientModulesForEntitlementV12_','function clientLicenseTierV12_','accessibleModules:clientModulesForEntitlementV12_'];var missingEntitlement=requiredEntitlement.filter(function(token){return code.indexOf(token)<0;});if(missingEntitlement.length)return{success:false,code:'V12_ENTITLEMENT_BRIDGE_MISSING',message:'V12 entitlement enforcement is incomplete: '+missingEntitlement.join(', ')};
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
