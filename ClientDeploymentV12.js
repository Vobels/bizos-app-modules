// ============================================================
// ClientDeploymentV12.js - DISTINCT CLIENT PRODUCT PACKAGE
// ============================================================
// V12 intentionally does NOT copy the Master BizOS UI.
// It reuses the proven client-safe server/runtime layer from V11,
// but serves an independently designed client-facing application.
// ============================================================

var CLIENT_V12_UI_FILES_ = ['ClientShellV12','ClientStylesV12'];
var CLIENT_V12_PACKAGE_VERSION_ = '12.2.0';
var CLIENT_V12_PACKAGE_RELEASE_ = 'V12.2.0';
var CLIENT_V12_PACKAGE_PROVENANCE_ = 'fix/full-client-bizos-package';

function generateClientCodeSafelyV12(settings) {
  settings = settings || {};
  if (!settings.clientId || !settings.sheetId || !settings.email) throw new Error('clientId, sheetId and email are required');

  // Branding/config values are data, never executable source. Keep the
  // object normalized here; the generated runtime embeds it with JSON.stringify
  // so apostrophes, quotes, backslashes, newlines and URLs cannot break JS.
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
    customDomain:String(settings.customDomain || ''),
    featureProfile:settings.featureProfile&&typeof settings.featureProfile==='object'?settings.featureProfile:{}
  };
  config.businessProfile = getMasterBusinessProfileForV12_(config.businessId, config.email);
  config.currency = String((config.businessProfile&&config.businessProfile.currency)||(config.businessProfile&&config.businessProfile.Currency)||'NGN').toUpperCase();

  var runtime = buildClientRuntimeV11_(config);

  // V12 entitlement, Apps and module-workspace logic requires the canonical
  // module registry to be present inside the generated client runtime.
  // Keep this dependency explicit so a future V11 runtime change cannot
  // silently produce a package that fails during client login.
  if (runtime.indexOf('var MODULES=') < 0) {
    runtime = 'var MODULES={Finance:{sheet:"Financial_Data",label:"Financial Management",icon:"calculator"},Sales:{sheet:"Sales_Data",label:"Sales Pipeline",icon:"chart-line"},Ecommerce:{sheet:"Ecommerce_Data",label:"E-commerce",icon:"cart"},CRM:{sheet:"CRM_Data",label:"Customer Relations",icon:"users"},HR:{sheet:"HR_Data",label:"Human Resources",icon:"briefcase"},Logistics:{sheet:"Logistics_Data",label:"Logistics",icon:"truck"},Tax:{sheet:"Tax_Data",label:"Tax Compliance",icon:"file-invoice"},Agro:{sheet:"Agro_Data",label:"Agriculture",icon:"seedling"},Productivity:{sheet:"Productivity_Data",label:"Productivity Suite",icon:"clock"},POS:{sheet:"POS_Data",label:"Point of Sale",icon:"cash-register"},Attendance:{sheet:"Attendance_Data",label:"Staff Attendance",icon:"fingerprint"},Warehouse:{sheet:"Warehouse_Data",label:"Warehouse Management",icon:"warehouse"}};\n' + runtime;
  }

  // Runtime-safe module accessor: login/session creation must never fail merely
  // because the generated MODULES registry is unavailable in a runtime scope.
  // When MODULES exists, it remains the single canonical registry.
  var moduleRegistryHelperV12_ = 'function clientModuleNamesV12_(){var fallback=["Finance","Sales","Ecommerce","CRM","HR","Logistics","Tax","Agro","Productivity","POS","Attendance","Warehouse"];try{return typeof MODULES!=="undefined"&&MODULES?Object.keys(MODULES):fallback;}catch(e){return fallback;}}\n';
  runtime = moduleRegistryHelperV12_ + runtime;

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
  // Apply entitlement filtering only to the master-authenticated owner session.
  // Do not replace the first generic accessibleModules:clientModuleNamesV12_()
  // occurrence: findClientUser_() also contains that expression and has no x variable.
  runtime = runtime.replace(
    'subscriptionTier:x.user.subscriptionTier||"sovereign",workspaceId:CLIENT_CONFIG.sheetId,isDemo:false,isClient:true,accessibleModules:clientModuleNamesV12_()',
    'subscriptionTier:x.user.subscriptionTier||"sovereign",workspaceId:CLIENT_CONFIG.sheetId,isDemo:false,isClient:true,accessibleModules:clientModulesForEntitlementV12_(x.user&&x.user.role,x.user&&x.user.subscriptionTier)'
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
    'function clientModulesForEntitlementV12_(role,tier,assigned){var r=String(role||"staff").toLowerCase(),t=String(tier||"").trim().toLowerCase(),base;if(!t)return[];base=(t==="free"||t==="starter")?["Finance","Ecommerce"]:clientModuleNamesV12_();if(r==="owner"||r==="admin")return base.slice();var a=Array.isArray(assigned)?assigned:[];return a.filter(function(m){return base.indexOf(String(m||""))!==-1;});}',
    'function clientLicenseTierV12_(){try{var r=validateClientLicenseV11_({clientId:CLIENT_CONFIG.clientId,businessId:CLIENT_CONFIG.businessId,sheetId:CLIENT_CONFIG.sheetId});if(!r||r.success!==true||!r.client)return"";var tier=String(r.client.tier||r.client.subscriptionTier||"").trim().toLowerCase();return tier;}catch(e){return"";}}',
        'function getClientUserFromSession(sid){try{var u=getUserFromSession(sid);if(!u||String(u.businessId||"")!==String(CLIENT_CONFIG.businessId||""))return null;var license=validateClientLicenseV11_({clientId:CLIENT_CONFIG.clientId,businessId:CLIENT_CONFIG.businessId,sheetId:CLIENT_CONFIG.sheetId});if(!license||license.success!==true||!license.client)return null;var tier=String(license.client.tier||license.client.subscriptionTier||"").trim();if(!tier)return null;u.isClient=true;u.clientId=CLIENT_CONFIG.clientId;u.subscriptionTier=tier;u.workspaceId=CLIENT_CONFIG.sheetId;u.accessibleModules=clientModulesForEntitlementV12_(u.role,tier,clientAssignedModulesV11_(u));return u;}catch(e){return null;}}',
'function getBusinessStaff(bid,sid){var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(bid)||["owner","admin"].indexOf(String(u.role||"").toLowerCase())<0)return{success:false,message:"Only the business owner or an administrator can view staff."};return{success:true,members:clientTeamRowsV11_(ensureClientTeamSheetV11_()).map(function(r){return{email:r.Email,name:r.Name,role:r.Role,phone:r.Phone,department:r.Department,status:r.Status,assignedModules:clientAssignedModulesV11_(r),invitationStatus:(String(r.Status||"").toLowerCase()==="invited"?"pending":String(r.Status||"").toLowerCase()==="accepted"?"accepted":String(r.Status||"").toLowerCase())};})};}',
    'function inviteStaffMember(bid,email,name,role,invitedBy,assignedModules,sid){try{var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(bid)||["owner","admin"].indexOf(String(u.role||"").toLowerCase())<0)return{success:false,message:"Only the business owner or an administrator can invite staff."};var result=inviteStaffMemberLegacy_(bid,email,name,role,invitedBy,assignedModules,sid);if(result&&result.success){logClientActivityV12_(u,"staff_invited","Invited "+String(email||"").toLowerCase()+" to the team.","staff");createClientNotificationV12_("all","staff","New team invitation",String(email||"").toLowerCase()+" was invited to the team.",u.email,"staff");}return result;}catch(e){return{success:false,message:e.message,code:"STAFF_INVITE_ERROR"};}}',
    'function resendStaffInvitation(bid,email,sid){try{var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(bid)||["owner","admin"].indexOf(String(u.role||"").toLowerCase())<0)return{success:false,message:"Only the business owner or an administrator can resend invitations."};email=String(email||"").trim().toLowerCase();var s=ensureClientTeamSheetV11_(),rows=clientTeamRowsV11_(s),t=rows.find(function(r){return String(r.Email||"").toLowerCase()===email;});if(!t)return{success:false,message:"Team member invitation not found."};if(String(t.Status||"").toLowerCase()==="accepted")return{success:false,message:"This team member has already accepted the invitation."};var pass=Utilities.getUuid().replace(/-/g,"").slice(0,12)+"A!",v=s.getDataRange().getValues(),h=v[0]||[],emailCol=h.indexOf("Email"),hashCol=h.indexOf("Password_Hash"),mustCol=h.indexOf("Must_Change_Password"),statusCol=h.indexOf("Status"),updatedCol=h.indexOf("Updated_At"),row=-1;for(var i=1;i<v.length;i++){if(String(v[i][emailCol]||"").toLowerCase()===email){row=i+1;break;}}if(row<0)return{success:false,message:"Team member invitation not found."};if(hashCol>=0)s.getRange(row,hashCol+1).setValue(clientHashV11_(pass));if(mustCol>=0)s.getRange(row,mustCol+1).setValue("YES");if(statusCol>=0)s.getRange(row,statusCol+1).setValue("invited");if(updatedCol>=0)s.getRange(row,updatedCol+1).setValue(new Date().toISOString());MailApp.sendEmail({to:email,subject:"Your BizOS invitation to "+CLIENT_CONFIG.clientName,htmlBody:"<p>Hello <strong>"+String(t.Name||email)+"</strong>,</p><p>Your BizOS invitation has been resent for <strong>"+CLIENT_CONFIG.clientName+"</strong>.</p><p><strong>Email:</strong> "+email+"<br><strong>Temporary password:</strong> "+pass+"</p><p>Please sign in and change your password from your profile.</p>"});logClientActivityV12_(u,"staff_invitation","Resent a staff invitation to "+email,"staff");createClientNotificationV12_("all","staff","Staff invitation resent","A staff invitation was resent to "+email+".",u.email,"staff");return{success:true,message:"Invitation resent to "+email+"."};}catch(e){return{success:false,message:e.message,code:"STAFF_RESEND_ERROR"};}}',
    'function assignModuleToStaff(email,bid,moduleName,assignedBy,sid){try{var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(bid)||["owner","admin"].indexOf(String(u.role||"").toLowerCase())<0)return{success:false,message:"Only the business owner or an administrator can assign modules."};var mod=String(moduleName||"").trim(),tier=clientLicenseTierV12_(),allowed=(String(tier||"").toLowerCase()==="free"||String(tier||"").toLowerCase()==="starter")?["Ecommerce"]:Object.keys(MODULES);if(allowed.indexOf(mod)<0)return{success:false,message:"This module is not available for the current client plan."};var s=ensureClientTeamSheetV11_(),v=s.getDataRange().getValues(),h=v[0]||[],emailCol=h.indexOf("Email"),modulesCol=h.indexOf("Assigned_Modules"),row=-1;for(var i=1;i<v.length;i++){if(String(v[i][emailCol]||"").toLowerCase()===String(email||"").toLowerCase()){row=i+1;break;}}if(row<0)return{success:false,message:"Team member not found."};var raw=modulesCol>=0?String(s.getRange(row,modulesCol+1).getValue()||""):"",mods=[];try{mods=raw?JSON.parse(raw):[];}catch(e){mods=raw.split(",").map(function(x){return String(x||"").trim();}).filter(Boolean);}if(mods.indexOf(mod)<0)mods.push(mod);if(modulesCol<0){s.getRange(1,s.getLastColumn()+1).setValue("Assigned_Modules");modulesCol=s.getLastColumn()-1;}s.getRange(row,modulesCol+1).setValue(JSON.stringify(mods));logClientActivityV12_(u,"staff_module","Granted "+mod+" access to "+email+".","staff");createClientNotificationV12_("all","staff","Staff access updated",mod+" access was granted to "+email+".",u.email,"staff");return{success:true,message:"Module access updated."};}catch(e){return{success:false,message:e.message,code:"STAFF_MODULE_ASSIGN_ERROR"};}}',
    'function removeModuleFromStaff(email,bid,moduleName,sid){try{var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(bid)||["owner","admin"].indexOf(String(u.role||"").toLowerCase())<0)return{success:false,message:"Only the business owner or an administrator can remove modules."};var mod=String(moduleName||"").trim(),s=ensureClientTeamSheetV11_(),v=s.getDataRange().getValues(),h=v[0]||[],emailCol=h.indexOf("Email"),modulesCol=h.indexOf("Assigned_Modules"),row=-1;for(var i=1;i<v.length;i++){if(String(v[i][emailCol]||"").toLowerCase()===String(email||"").toLowerCase()){row=i+1;break;}}if(row<0)return{success:false,message:"Team member not found."};var raw=modulesCol>=0?String(s.getRange(row,modulesCol+1).getValue()||""):"",mods=[];try{mods=raw?JSON.parse(raw):[];}catch(e){mods=raw.split(",").map(function(x){return String(x||"").trim();}).filter(Boolean);}mods=mods.filter(function(x){return x!==mod;});if(modulesCol>=0)s.getRange(row,modulesCol+1).setValue(JSON.stringify(mods));logClientActivityV12_(u,"staff_module","Removed "+mod+" access from "+email+".","staff");createClientNotificationV12_("all","staff","Staff access updated",mod+" access was removed from "+email+".",u.email,"staff");return{success:true,message:"Module access updated."};}catch(e){return{success:false,message:e.message,code:"STAFF_MODULE_REMOVE_ERROR"};}}',
        'function removeTeamMember(bid,email,sid){try{var u=getUserFromSession(sid);if(!u||String(u.businessId)!==String(bid)||["owner","admin"].indexOf(String(u.role||"").toLowerCase())<0)return{success:false,message:"Unauthorized"};var s=ensureClientTeamSheetV11_(),rows=clientTeamRowsV11_(s),t=rows.find(function(r){return String(r.Email||"").toLowerCase()===String(email||"").toLowerCase();});if(!t)return{success:false,message:"Team member not found."};s.deleteRow(t.row);logClientActivityV12_(u,"staff_removed","Removed team member "+email+".","staff");createClientNotificationV12_("all","staff","Team member removed",email+" was removed from the team.",u.email,"staff");return{success:true,message:"Team member removed."};}catch(e){return{success:false,message:e.message};}}',
'function getClientModuleSummary(sid){var u=getClientUserFromSession(sid);if(!u)return{};var allowed={};(Array.isArray(u.accessibleModules)?u.accessibleModules:[]).forEach(function(m){allowed[String(m)]=true;});var role=String(u.role||"staff").toLowerCase(),readOnly=role!=="owner"&&role!=="admin";var out={},ws=null;try{ws=getWorkspaceFile(sid);}catch(e){ws=null;}Object.keys(MODULES).forEach(function(name){var def=MODULES[name]||{},count=0;try{var sh=ws&&ws.getSheetByName(def.sheet);count=sh&&sh.getLastRow()>1?sh.getLastRow()-1:0;}catch(e2){}out[name]={label:def.label||name,icon:def.icon||"",count:count,isAccessible:!!allowed[name],isReadOnly:readOnly};});return out;}'
  ].join(String.fromCharCode(10));
  runtime += String.fromCharCode(10) + v12SecurityBridge + String.fromCharCode(10);
  // Live workspace validation: used by the client smoke audit to verify the
  // actual spreadsheet/database structure behind the authenticated session.
  // This is read-only and does not create or modify business data.
  var v12AccountBridge = [
    'function getClientV12WorkspaceHealth(sid){var u=getClientUserFromSession(sid);if(!u)return{success:false,message:"Session expired or client license is no longer valid.",code:"CLIENT_LICENSE_INVALID"};try{var ws=getWorkspaceFile(sid),info=ws.getSheetByName("Client_Info"),infoValues=info?info.getDataRange().getValues():[],meta={};if(!info)return{success:false,message:"Client_Info sheet is missing.",code:"CLIENT_INFO_MISSING"};for(var i=1;i<infoValues.length;i++){if(infoValues[i][0])meta[String(infoValues[i][0])]=infoValues[i][1];}var names=["Finance","Sales","Ecommerce","CRM","HR","Logistics","Tax","Agro","Productivity","POS","Attendance","Warehouse"],modules=[];names.forEach(function(name){var def=MODULES[name]||{},sheet=ws.getSheetByName(def.sheet),headers=[];try{if(sheet&&sheet.getLastColumn()>0)headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(x){return String(x||"");});}catch(e){}modules.push({module:name,sheet:String(def.sheet||""),exists:!!sheet,headersPresent:headers.length>0,recordCount:sheet&&sheet.getLastRow()>1?sheet.getLastRow()-1:0});});var ns=ws.getSheetByName("Client_Notifications"),as=ws.getSheetByName("Client_Activity"),notificationCheck=getClientNotificationsData(sid);return{success:true,clientInfo:{clientId:String(meta.Client_ID||""),businessId:String(meta.Business_ID||""),status:String(meta.Status||""),sheetId:String(CLIENT_CONFIG.sheetId||"")},moduleCount:modules.length,modules:modules,notificationsSheet:{exists:!!ns,headersPresent:!!(ns&&ns.getLastColumn()>0),readSuccess:!!(notificationCheck&&notificationCheck.success)},activitySheet:{exists:!!as,headersPresent:!!(as&&as.getLastColumn()>0),readSuccess:!!(notificationCheck&&notificationCheck.success)},notificationReadSuccess:!!(notificationCheck&&notificationCheck.success),notificationCount:notificationCheck&&Array.isArray(notificationCheck.notifications)?notificationCheck.notifications.length:0};}catch(e){return{success:false,message:String(e&&e.message||e),code:"WORKSPACE_HEALTH_ERROR"};}}',
    'function getClientV12ModuleSmokeAudit(sid){var u=getClientUserFromSession(sid);if(!u)return{success:false,message:"Session expired or client license is no longer valid.",code:"CLIENT_LICENSE_INVALID"};var names=["Finance","Sales","Ecommerce","CRM","HR","Logistics","Tax","Agro","Productivity","POS","Attendance","Warehouse"],out=[];var ws=getWorkspaceFile(sid);names.forEach(function(name){var def=MODULES[name]||{},sheet=ws&&ws.getSheetByName(def.sheet),accessible=Array.isArray(u.accessibleModules)&&u.accessibleModules.indexOf(name)>=0,headers=[];try{if(sheet&&sheet.getLastColumn()>0)headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(x){return String(x||"");});}catch(e){}var read=null;if(accessible){try{read=getModuleDataWithSync(name,sid); }catch(e2){read={success:false,message:e2.message};}}out.push({module:name,sheet:String(def.sheet||""),accessible:accessible,sheetExists:!!sheet,headers:headers,recordCount:read&&read.success!==false?Number(read.recordCount||0):null,readSuccess:accessible?!!(read&&read.success):true,message:read&&read.success===false?String(read.message||""):""});});return{success:true,packageVersion:CLIENT_V12_PACKAGE_VERSION_,modules:out};}',
    'function getClientFeatureProfileV12(sid){var u=getClientUserFromSession(sid);if(!u)return{success:false,message:"Session expired or client license is no longer valid."};var p=CLIENT_CONFIG.featureProfile&&typeof CLIENT_CONFIG.featureProfile==="object"?CLIENT_CONFIG.featureProfile:{};return{success:true,clientId:String(CLIENT_CONFIG.clientId||""),profile:p};}',
    'function getClientPackageInfo(sid){var u=getClientUserFromSession(sid);if(!u)return{success:false,message:"Session expired or client license is no longer valid.",code:"CLIENT_LICENSE_INVALID"};return{success:true,packageVersion:CLIENT_V12_PACKAGE_VERSION_,release:CLIENT_V12_PACKAGE_RELEASE_,provenance:CLIENT_V12_PACKAGE_PROVENANCE_,clientId:String(CLIENT_CONFIG.clientId||""),workspaceId:String(CLIENT_CONFIG.sheetId||""),generatedAt:String(CLIENT_CONFIG.generatedAt||"")};}',

    'function clientNotificationActorV12_(sid){var u=getUserFromSession(sid);if(!u)return{ok:false,message:"Session expired. Please login again."};return{ok:true,user:u};}',
    'function clientNotificationAudienceMatchesV12_(audience,user){var a=String(audience||"all").trim().toLowerCase(),email=String(user&&user.email||"").trim().toLowerCase(),role=String(user&&user.role||"staff").trim().toLowerCase();if(!a||a==="all")return true;if(a===email)return true;if(a==="owner_admin"||a==="owner-admin"||a==="admins")return role==="owner"||role==="admin";if(a==="staff")return role==="staff";return false;}',
    'function ensureClientNotificationsSheetV12_(ws){var s=ws.getSheetByName("Client_Notifications");if(!s){s=ws.insertSheet("Client_Notifications");s.appendRow(["Notification_ID","Audience","Type","Title","Message","Created_At","Created_By","Related_View","Read_By_JSON"]);}return s;}',
    'function ensureClientActivitySheetV12_(ws){var s=ws.getSheetByName("Client_Activity");if(!s){s=ws.insertSheet("Client_Activity");s.appendRow(["Activity_ID","Actor_Email","Actor_Name","Event_Type","Summary","Created_At","Related_View"]);}return s;}',
    'function clientNotificationReadMapV12_(raw){try{var x=JSON.parse(String(raw||"{}"));return x&&typeof x==="object"?x:{};}catch(e){return{};}}',
    'function createClientNotificationV12_(audience,type,title,message,createdBy,relatedView){try{var ws=getWorkspaceFileByConfigV12_(),s=ensureClientNotificationsSheetV12_(ws),id="NTF-"+Utilities.getUuid().replace(/-/g,"").slice(0,10).toUpperCase(),now=new Date().toISOString();s.appendRow([id,String(audience||"all"),String(type||"system"),String(title||"Workspace update"),String(message||""),now,String(createdBy||""),String(relatedView||""),"{}"]);return{id:id};}catch(e){return{error:e.message};}}',
    'function logClientActivityV12_(actor,eventType,summary,relatedView){try{var ws=getWorkspaceFileByConfigV12_(),s=ensureClientActivitySheetV12_(ws),id="ACT-"+Utilities.getUuid().replace(/-/g,"").slice(0,10).toUpperCase(),u=actor||{},now=new Date().toISOString();s.appendRow([id,String(u.email||""),String(u.name||u.email||"Workspace member"),String(eventType||"workspace"),String(summary||""),now,String(relatedView||"")]);return{id:id};}catch(e){return{error:e.message};}}',
    'function getWorkspaceFileByConfigV12_(){return SpreadsheetApp.openById(String(CLIENT_CONFIG.sheetId||""));}',
    'function getClientNotificationsData(sid){var a=clientNotificationActorV12_(sid);if(!a.ok)return{success:false,message:a.message};try{var ws=getWorkspaceFile(sid),ns=ensureClientNotificationsSheetV12_(ws),as=ensureClientActivitySheetV12_(ws),nv=ns.getDataRange().getValues(),nh=nv[0]||[],av=as.getDataRange().getValues(),ah=av[0]||[],email=String(a.user.email||"").toLowerCase(),notifications=[];for(var i=1;i<nv.length;i++){var row=nv[i],o={};nh.forEach(function(k,j){o[k]=row[j];});var audience=String(o.Audience||"all").toLowerCase(),reads=clientNotificationReadMapV12_(o.Read_By_JSON);if(clientNotificationAudienceMatchesV12_(audience,a.user)){notifications.push({id:String(o.Notification_ID||""),type:String(o.Type||"system"),title:String(o.Title||""),message:String(o.Message||""),createdAt:o.Created_At,relatedView:String(o.Related_View||""),read:!!reads[email]});}}notifications.sort(function(a,b){return new Date(b.createdAt)-new Date(a.createdAt);});var activity=[];for(var j=1;j<av.length;j++){var ar=av[j],ao={};ah.forEach(function(k,j2){ao[k]=ar[j2];});if(ao.Activity_ID)activity.push({id:String(ao.Activity_ID),actorEmail:String(ao.Actor_Email||""),actorName:String(ao.Actor_Name||"Workspace member"),eventType:String(ao.Event_Type||"workspace"),summary:String(ao.Summary||""),createdAt:ao.Created_At,relatedView:String(ao.Related_View||"")});}activity.sort(function(a,b){return new Date(b.createdAt)-new Date(a.createdAt);});if(!activity.length){var fs=ws.getSheetByName("Financial_Data");if(fs&&fs.getLastRow()>1){var fv=fs.getDataRange().getValues(),fh=fv[0]||[],dc=fh.indexOf("Date"),tc=fh.indexOf("Type"),xc=fh.indexOf("Description"),ac=fh.indexOf("Amount");for(var k=Math.max(1,fv.length-15);k<fv.length;k++){var rr=fv[k];activity.push({id:"FIN-"+k,actorEmail:"",actorName:"Workspace",eventType:"finance",summary:String(rr[xc>=0?xc:0]||rr[tc>=0?tc:0]||"Financial entry")+(ac>=0&&rr[ac]!==""?" · "+String(rr[ac]):""),createdAt:dc>=0?rr[dc]:""});}}}activity.sort(function(a,b){return new Date(b.createdAt)-new Date(a.createdAt);});return{success:true,notifications:notifications.slice(0,50),activity:activity.slice(0,50)};}catch(e){return{success:false,message:e.message,code:"CLIENT_NOTIFICATIONS_LOAD_ERROR"};}}',
    'function markClientNotificationRead(id,sid){var a=clientNotificationActorV12_(sid);if(!a.ok)return{success:false,message:a.message};id=String(id||"").trim();try{var ws=getWorkspaceFile(sid),s=ensureClientNotificationsSheetV12_(ws),v=s.getDataRange().getValues(),h=v[0]||[],idCol=h.indexOf("Notification_ID"),readCol=h.indexOf("Read_By_JSON"),email=String(a.user.email||"").toLowerCase();for(var i=1;i<v.length;i++){if(String(v[i][idCol]||"")!==id)continue;var audience=String(v[i][h.indexOf("Audience")]||"all");if(!clientNotificationAudienceMatchesV12_(audience,a.user))return{success:false,message:"Notification not found."};var reads=clientNotificationReadMapV12_(v[i][readCol]);reads[email]=new Date().toISOString();s.getRange(i+1,readCol+1).setValue(JSON.stringify(reads));return{success:true};}return{success:false,message:"Notification not found."};}catch(e){return{success:false,message:e.message};}}',
    'function markAllClientNotificationsRead(sid){var a=clientNotificationActorV12_(sid);if(!a.ok)return{success:false,message:a.message};try{var ws=getWorkspaceFile(sid),s=ensureClientNotificationsSheetV12_(ws),v=s.getDataRange().getValues(),h=v[0]||[],readCol=h.indexOf("Read_By_JSON"),email=String(a.user.email||"").toLowerCase(),now=new Date().toISOString();for(var i=1;i<v.length;i++){var audience=String(v[i][h.indexOf("Audience")]||"all").toLowerCase();if(!clientNotificationAudienceMatchesV12_(audience,a.user))continue;var reads=clientNotificationReadMapV12_(v[i][readCol]);reads[email]=now;s.getRange(i+1,readCol+1).setValue(JSON.stringify(reads));}return{success:true};}catch(e){return{success:false,message:e.message};}}',
    'function getClientProfileData(sid){var u=getUserFromSession(sid);if(!u)return{success:false,message:"Session expired. Please login again."};var license={};try{var r=validateClientLicenseV11_({clientId:CLIENT_CONFIG.clientId,businessId:CLIENT_CONFIG.businessId,sheetId:CLIENT_CONFIG.sheetId});if(r&&r.success&&r.client)license=r.client||{};}catch(e){}var tier=String(license.tier||u.subscriptionTier||"");var verification=String(license.verificationStatus||license.verification_status||"Not available");return{success:true,user:{name:String(u.name||""),email:String(u.email||CLIENT_CONFIG.email||""),phone:String(u.phone||""),role:String(u.role||"owner")},businessName:String(CLIENT_CONFIG.businessProfile&&CLIENT_CONFIG.businessProfile.businessName||u.businessName||CLIENT_CONFIG.clientName||""),businessId:String(CLIENT_CONFIG.businessId||u.businessId||""),subscriptionTier:tier,verificationStatus:verification,workspaceUrl:String(CLIENT_CONFIG.applicationUrl||CLIENT_CONFIG.landingUrl||"")};}',
    'function clientBillingAccessV12_(sid,write){var u=getUserFromSession(sid);if(!u)return{ok:false,message:"Session expired. Please login again."};if(["owner","admin"].indexOf(String(u.role||"").toLowerCase())<0)return{ok:false,message:"Only the business owner or an administrator can manage billing and maintenance."};if(write&&u.isDemo==="YES")return{ok:false,message:"Demo accounts cannot submit maintenance requests."};return{ok:true,user:u};}',
    'function ensureClientMaintenanceSheetV12_(ws){var s=ws.getSheetByName("Client_Maintenance_Requests");if(!s){s=ws.insertSheet("Client_Maintenance_Requests");s.appendRow(["Request_ID","Type","Details","Status","Created_At","Requested_By","Updated_At"]);}return s;}',
    'function ensureClientFeatureRequestSheetV12_(ws){var s=ws.getSheetByName("Client_Feature_Requests");if(!s){s=ws.insertSheet("Client_Feature_Requests");s.appendRow(["Request_ID","Title","Details","Priority","Status","Created_At","Requested_By","Updated_At"]);}return s;}',
    'function ensureClientQuotesSheetV12_(ws){var s=ws.getSheetByName("Client_Quotes");if(!s){s=ws.insertSheet("Client_Quotes");s.appendRow(["Quote_ID","Title","Details","Amount","Currency","Status","Created_At","Updated_At","Payment_Ref","Payment_Status","Paid_At"]);}else{var h=s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(function(x){return String(x||"");});["Payment_Ref","Payment_Status","Paid_At"].forEach(function(k){if(h.indexOf(k)<0){s.getRange(1,s.getLastColumn()+1).setValue(k);h.push(k);}});}return s;}',
    'function ensureClientServiceHistorySheetV12_(ws){var s=ws.getSheetByName("Client_Service_History");if(!s){s=ws.insertSheet("Client_Service_History");s.appendRow(["Service_ID","Type","Description","Status","Amount","Currency","Created_At","Updated_At"]);}return s;}',
    'function clientBillingRowsV12_(sheet,mapper){var v=sheet.getDataRange().getValues(),h=v[0]||[],out=[];for(var i=1;i<v.length;i++){var o={};h.forEach(function(k,j){o[String(k||"")]=v[i][j];});out.push(mapper(o));}return out;}',
    'function getClientQuotePaymentDetails(quoteId,sid){var a=clientBillingAccessV12_(sid,false);if(!a.ok)return{success:false,message:a.message};var ws=getWorkspaceFile(sid),s=ensureClientQuotesSheetV12_(ws),v=s.getDataRange().getValues(),h=v[0]||[],idc=h.indexOf("Quote_ID"),tc=h.indexOf("Title"),ac=h.indexOf("Amount"),cc=h.indexOf("Currency"),sc=h.indexOf("Status"),rc=h.indexOf("Payment_Ref"),psc=h.indexOf("Payment_Status");for(var i=1;i<v.length;i++){if(String(v[i][idc]||"")!==String(quoteId||""))continue;var status=String(v[i][sc]||"pending").toLowerCase();if(status==="paid")return{success:false,message:"This quote has already been paid."};var amount=Number(v[i][ac]||0);if(!amount||amount<=0)return{success:false,message:"This quote does not have a valid payment amount."};var currency=String(v[i][cc]||"NGN").toUpperCase(),ref="BIZOS-Q-"+String(quoteId).replace(/[^A-Za-z0-9.=-]/g,"-")+"-"+Utilities.getUuid().replace(/-/g,"").slice(0,8).toUpperCase();if(rc>=0)s.getRange(i+1,rc+1).setValue(ref);if(psc>=0)s.getRange(i+1,psc+1).setValue("pending");return{success:true,quoteId:String(quoteId),title:String(v[i][tc]||"BizOS service"),amount:Math.round(amount*100),displayAmount:amount,currency:currency,reference:ref,publicKey:getPaystackPublicKey_()};}return{success:false,message:"Quote not found."};}',
    'function verifyClientQuotePayment(quoteId,reference,sid){var a=clientBillingAccessV12_(sid,true);if(!a.ok)return{success:false,message:a.message};if(!reference)return{success:false,message:"Payment reference is required."};var ws=getWorkspaceFile(sid),s=ensureClientQuotesSheetV12_(ws),v=s.getDataRange().getValues(),h=v[0]||[],idc=h.indexOf("Quote_ID"),ac=h.indexOf("Amount"),cc=h.indexOf("Currency"),sc=h.indexOf("Status"),rc=h.indexOf("Payment_Ref"),psc=h.indexOf("Payment_Status"),pc=h.indexOf("Paid_At"),uc=h.indexOf("Updated_At");var row=-1;for(var i=1;i<v.length;i++){if(String(v[i][idc]||"")===String(quoteId||"")){row=i;break;}}if(row<0)return{success:false,message:"Quote not found."};var expected=rc>=0?String(v[row][rc]||""):"";if(expected&&expected!==String(reference))return{success:false,message:"This payment does not match the selected quote."};var amount=Number(v[row][ac]||0),currency=String(v[row][cc]||"NGN").toUpperCase();try{var response=UrlFetchApp.fetch("https://api.paystack.co/transaction/verify/"+encodeURIComponent(reference),{method:"get",headers:{Authorization:"Bearer "+getPaystackSecretKey_()},muteHttpExceptions:true});var data=JSON.parse(response.getContentText()||"{}");if(!data.status||!data.data||String(data.data.status||"").toLowerCase()!=="success")return{success:false,message:"Paystack has not confirmed this payment yet."};var paidKobo=Number(data.data.amount||0),expectedKobo=Math.round(amount*100);if(paidKobo!==expectedKobo||String(data.data.currency||"").toUpperCase()!==currency)return{success:false,message:"The payment amount or currency does not match this quote."};}catch(e){console.error("verifyClientQuotePayment error:",e);return{success:false,message:"We could not verify the payment right now. Please try again."};}var now=new Date().toISOString();if(sc>=0)s.getRange(row+1,sc+1).setValue("paid");if(psc>=0)s.getRange(row+1,psc+1).setValue("paid");if(pc>=0)s.getRange(row+1,pc+1).setValue(now);if(uc>=0)s.getRange(row+1,uc+1).setValue(now);logClientActivityV12_(a.user,"quote_payment","Paid quote "+String(quoteId),"billing");createClientNotificationV12_("all","payment","Quote payment received","Payment for quote "+String(quoteId)+" was confirmed.",a.user.email,"billing");return{success:true,quoteId:String(quoteId),message:"Payment confirmed."};}'
    'function getClientBillingData(sid){var a=clientBillingAccessV12_(sid,false);if(!a.ok)return{success:false,message:a.message};var license={};try{var r=validateClientLicenseV11_({clientId:CLIENT_CONFIG.clientId,businessId:CLIENT_CONFIG.businessId,sheetId:CLIENT_CONFIG.sheetId});if(r&&r.success&&r.client)license=r.client||{};}catch(e){}var tier=String(license.tier||a.user.subscriptionTier||"");var status=String(license.status||"active");var amount=license.amountPaid!=null?license.amountPaid:(license.amount_paid!=null?license.amount_paid:"");var currency=String(license.currency||"");var ws=getWorkspaceFile(sid),requests=clientBillingRowsV12_(ensureClientMaintenanceSheetV12_(ws),function(o){return{id:String(o.Request_ID||""),type:String(o.Type||""),details:String(o.Details||""),status:String(o.Status||"submitted"),createdAt:String(o.Created_At||""),requestedBy:String(o.Requested_By||""),updatedAt:String(o.Updated_At||"")};}),features=clientBillingRowsV12_(ensureClientFeatureRequestSheetV12_(ws),function(o){return{id:String(o.Request_ID||""),title:String(o.Title||""),details:String(o.Details||""),priority:String(o.Priority||"normal"),status:String(o.Status||"submitted"),createdAt:String(o.Created_At||""),requestedBy:String(o.Requested_By||""),updatedAt:String(o.Updated_At||"")};}),quotes=clientBillingRowsV12_(ensureClientQuotesSheetV12_(ws),function(o){return{id:String(o.Quote_ID||""),title:String(o.Title||""),details:String(o.Details||""),amount:o.Amount===""?"":Number(o.Amount||0),currency:String(o.Currency||currency||""),status:String(o.Status||"pending"),createdAt:String(o.Created_At||""),updatedAt:String(o.Updated_At||"")};}),services=clientBillingRowsV12_(ensureClientServiceHistorySheetV12_(ws),function(o){return{id:String(o.Service_ID||""),type:String(o.Type||""),description:String(o.Description||""),status:String(o.Status||"completed"),amount:o.Amount===""?"":Number(o.Amount||0),currency:String(o.Currency||currency||""),createdAt:String(o.Created_At||""),updatedAt:String(o.Updated_At||"")};});return{success:true,subscriptionTier:tier,status:status,amountPaid:amount,currency:currency,originalPayment:{amount:amount,currency:currency,status:amount!==""?"confirmed":"not_recorded"},paymentStatus:status,maintenancePlans:[{id:"occasional",name:"Occasional maintenance",description:"One-off support when you need it."},{id:"monthly",name:"Monthly maintenance",description:"Ongoing maintenance and support."},{id:"custom",name:"Custom support",description:"A tailored support arrangement for your workspace."}],requests:requests,featureRequests:features,quotes:quotes,serviceHistory:services};}',
    'function submitClientMaintenanceRequest(data,sid){var a=clientBillingAccessV12_(sid,true);if(!a.ok)return{success:false,message:a.message};data=data||{};var type=String(data.type||"occasional").trim().toLowerCase(),details=String(data.details||"").trim();if(["occasional","monthly","technical","update"].indexOf(type)<0)return{success:false,message:"Invalid maintenance request type."};if(!details)return{success:false,message:"Request details are required."};if(details.length>4000)return{success:false,message:"Request details are too long."};var ws=getWorkspaceFile(sid),s=ensureClientMaintenanceSheetV12_(ws),id="MNT-"+Utilities.getUuid().replace(/-/g,"").slice(0,10).toUpperCase(),now=new Date().toISOString();s.appendRow([id,type,details,"submitted",now,String(a.user.email||""),now]);logClientActivityV12_(a.user,"maintenance_request","Submitted "+type+" maintenance request "+id,"billing");createClientNotificationV12_("all","maintenance","Maintenance request submitted","A "+type+" maintenance request was submitted by "+String(a.user.name||a.user.email||"a workspace member")+".",a.user.email,"billing");return{success:true,requestId:id};}',
    'function submitClientFeatureRequest(data,sid){var a=clientBillingAccessV12_(sid,true);if(!a.ok)return{success:false,message:a.message};data=data||{};var title=String(data.title||"").trim(),details=String(data.details||"").trim(),priority=String(data.priority||"normal").trim().toLowerCase();if(!title)return{success:false,message:"Please provide a feature title."};if(!details)return{success:false,message:"Please describe the feature you need."};if(details.length>4000)return{success:false,message:"Feature details are too long."};if(["low","normal","high"].indexOf(priority)<0)priority="normal";var ws=getWorkspaceFile(sid),s=ensureClientFeatureRequestSheetV12_(ws),id="FTR-"+Utilities.getUuid().replace(/-/g,"").slice(0,10).toUpperCase(),now=new Date().toISOString();s.appendRow([id,title,details,priority,"submitted",now,String(a.user.email||""),now]);logClientActivityV12_(a.user,"feature_request","Submitted custom feature request "+id,"billing");createClientNotificationV12_("all","feature_request","Feature request submitted","Your custom feature request has been submitted for review.",a.user.email,"billing");return{success:true,requestId:id};}'
  ].join(String.fromCharCode(10));
  runtime += String.fromCharCode(10) + v12AccountBridge + String.fromCharCode(10);
  // V12 session durability bridge: CacheService remains the fast path, but the
  // session is also persisted server-side so normal cache eviction cannot log
  // an active client out. The persisted record contains only the session token
  // and non-secret workspace/session metadata.
  var v12SessionBridge = [
    'function v12SessionPropertyKey_(sid){return "V12_SESSION_"+String(sid||"");}',
    'function v12PersistSession_(sid,session){try{PropertiesService.getScriptProperties().setProperty(v12SessionPropertyKey_(sid),JSON.stringify(session));}catch(e){console.error("V12 session persistence failed:",e&&e.message||e);}}',
    'function v12ReadPersistedSession_(sid){try{var raw=PropertiesService.getScriptProperties().getProperty(v12SessionPropertyKey_(sid));return raw?JSON.parse(raw):null;}catch(e){return null;}}',
    'function v12DeletePersistedSession_(sid){try{PropertiesService.getScriptProperties().deleteProperty(v12SessionPropertyKey_(sid));}catch(e){}}',
    'function v12ValidateWorkspaceBinding_(s){try{var ss=SpreadsheetApp.openById(String(s.sheetId||""));var info=ss.getSheetByName("Client_Info");if(!info)return{success:false,code:"CLIENT_INFO_MISSING",message:"Client workspace is not configured."};var v=info.getDataRange().getValues(),m={};for(var i=1;i<v.length;i++){if(v[i][0])m[String(v[i][0])]=v[i][1];}if(String(m.Client_ID||"")!==String(s.clientId||""))return{success:false,code:"CLIENT_ID_MISMATCH",message:"Client workspace identity mismatch."};if(String(m.Business_ID||s.businessId||"")!==String(s.businessId||""))return{success:false,code:"BUSINESS_ID_MISMATCH",message:"Client business identity mismatch."};if(String(m.Status||"active").toLowerCase()!=="active")return{success:false,code:"CLIENT_INACTIVE",message:"Client workspace is inactive."};return{success:true,spreadsheet:ss};}catch(e){return{success:false,code:"WORKSPACE_VALIDATION_ERROR",message:String(e&&e.message||e)};}}',
    'function createClientSessionV11_(email,businessId){var sid="CLIENT_SESS_"+Utilities.getUuid(),now=Date.now(),u=findClientUser_(email,businessId),session={email:String(email).toLowerCase(),businessId:String(businessId||CLIENT_CONFIG.businessId),clientId:String(CLIENT_CONFIG.clientId),sheetId:String(CLIENT_CONFIG.sheetId),businessName:String(CLIENT_CONFIG.clientName),created:now,expiresAt:now+21600000,licenseValidatedAt:now,isClient:true,user:u};var raw=JSON.stringify(session);CacheService.getScriptCache().put(sid,raw,21600);v12PersistSession_(sid,session);return sid;}',
    'function validateClientSession(sid){if(!sid)return null;sid=String(sid);try{var cache=CacheService.getScriptCache(),raw=cache.get(sid),source="cache";if(!raw){var persisted=v12ReadPersistedSession_(sid);if(!persisted)return null;raw=JSON.stringify(persisted);source="persistent";}var s=JSON.parse(raw),now=Date.now();if(!s||s.clientId!==String(CLIENT_CONFIG.clientId)||s.sheetId!==String(CLIENT_CONFIG.sheetId)||s.businessId!==String(CLIENT_CONFIG.businessId)||now>=Number(s.expiresAt||0)){cache.remove(sid);v12DeletePersistedSession_(sid);return null;}var workspace=v12ValidateWorkspaceBinding_(s);if(!workspace.success){cache.remove(sid);v12DeletePersistedSession_(sid);return null;}if(source==="persistent")cache.put(sid,JSON.stringify(s),21600);var checked=Number(s.licenseValidatedAt||0),licenseWindow=300000;if(now-checked<=licenseWindow)return s;var license=validateClientLicenseV11_(s);if(license&&license.success){s.licenseValidatedAt=now;var updated=JSON.stringify(s);cache.put(sid,updated,21600);v12PersistSession_(sid,s);return s;}if(license&&license.validationStatus==="invalid"){cache.remove(sid);v12DeletePersistedSession_(sid);return null;}return s;}catch(e){return null;}}',
    'function logout(sid){try{if(sid){CacheService.getScriptCache().remove(String(sid));v12DeletePersistedSession_(String(sid));}return{success:true,message:"Logged out successfully"};}catch(e){return{success:false,message:e.message};}}',
    'function clientLogout(sid){return logout(sid);}'
  ].join(String.fromCharCode(10));
  runtime += String.fromCharCode(10) + v12SessionBridge + String.fromCharCode(10);

  var businessCenterSource = String(getMasterSourceForV12_('BusinessHubGS') || '');
  var staffSource = String(getMasterSourceForV12_('ClientStaffV12') || '');
  var appsSource = String(getMasterSourceForV12_('ClientAppsV12') || '');
  var settingsSource = String(getMasterSourceForV12_('ClientSettingsV12') || '');
  var profileSource = String(getMasterSourceForV12_('ClientProfileV12') || '');
  var notificationsSource = String(getMasterSourceForV12_('ClientNotificationsV12') || '');
  var billingSource = String(getMasterSourceForV12_('ClientBillingV12') || '');
  if(!profileSource) throw new Error('ClientProfileV12 source is missing.');
  if(!notificationsSource) throw new Error('ClientNotificationsV12 source is missing.');
  if(!billingSource) throw new Error('ClientBillingV12 source is missing.');
  if(!staffSource) throw new Error('ClientStaffV12 source is missing.');
  if(!appsSource) throw new Error('ClientAppsV12 source is missing.');
  if(!settingsSource) throw new Error('ClientSettingsV12 source is missing.');
  if(!businessCenterSource) throw new Error('BusinessHubGS source is missing.');
  var businessCenterBridge = [
    'function clientSafeBusinessCenterResult_(result){try{return JSON.parse(JSON.stringify(result));}catch(e){return result;}}',
    'function clientSafeBusinessCenterActorV12_(sid,write){var u=getUserFromSession(sid);if(!u)return{ok:false,message:"Session expired. Please login again."};if(String(u.businessId)!==String(CLIENT_CONFIG.businessId))return{ok:false,message:"Workspace access denied."};if(write&&String(u.isDemo||"").toUpperCase()==="YES")return{ok:false,message:"Demo accounts are view-only."};return{ok:true,user:u};}',
    'function getClientBusinessCenterConnectedData(sid){var a=clientSafeBusinessCenterActorV12_(sid,false);if(!a.ok)return{success:false,message:a.message};return clientSafeBusinessCenterResult_(getBusinessCenterConnectedData(sid));}',
    'function getClientBusinessCenterProductImages(productIds,sid){var a=clientSafeBusinessCenterActorV12_(sid,false);if(!a.ok)return{success:false,message:a.message};return clientSafeBusinessCenterResult_(getBusinessCenterProductImages(productIds,sid));}',
    'function saveClientBusinessCenterProductsBulk(data,sid){var a=clientSafeBusinessCenterActorV12_(sid,true);if(!a.ok)return{success:false,message:a.message};return saveBusinessCenterProductsBulk(data,sid);}',
    'function lookupClientBusinessCenterProduct(code,sid){var a=clientSafeBusinessCenterActorV12_(sid,false);if(!a.ok)return{success:false,message:a.message};return lookupBusinessCenterProduct(code,sid);}',
    'function getClientBusinessCenterCustomerLedger(customerId,sid){var a=clientSafeBusinessCenterActorV12_(sid,false);if(!a.ok)return{success:false,message:a.message};return clientSafeBusinessCenterResult_(getBusinessCenterCustomerLedger(customerId,sid));}',
    'function recordClientBusinessCenterCustomerPayment(data,sid){var a=clientSafeBusinessCenterActorV12_(sid,true);if(!a.ok)return{success:false,message:a.message};return recordBusinessCenterCustomerPayment(data,sid);}',
    'function getClientBusinessCenterSaleDetails(saleId,sid){var a=clientSafeBusinessCenterActorV12_(sid,false);if(!a.ok)return{success:false,message:a.message};return clientSafeBusinessCenterResult_(getBusinessCenterSaleDetails(saleId,sid));}',
    'function getClientBusinessCenterStockMovements(sid,productId){var a=clientSafeBusinessCenterActorV12_(sid,false);if(!a.ok)return{success:false,message:a.message};return clientSafeBusinessCenterResult_(getBusinessCenterStockMovements(sid,productId));}',
    'function getClientBusinessCenterFinanceData(sid){var a=clientSafeBusinessCenterActorV12_(sid,false);if(!a.ok)return{success:false,message:a.message};return clientSafeBusinessCenterResult_(getBusinessCenterFinanceData(sid));}',
    'function recordClientBusinessCenterExpense(data,sid){var a=clientSafeBusinessCenterActorV12_(sid,true);if(!a.ok)return{success:false,message:a.message};return recordBusinessCenterExpense(data,sid);}',
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
    "function ensureClientNotificationPreferencesSheetV12_(ws){var s=ws.getSheetByName('Client_Notification_Preferences');if(!s){s=ws.insertSheet('Client_Notification_Preferences');s.appendRow(['Preference_ID','Business_ID','Email','Feature_Updates','System_Alerts','Weekly_Reports','Security_Alerts','Marketing_Updates','Updated_At']);}return s;}",
    "function getClientNotificationPreferences(sid){var access=clientSettingsAccessV12_(sid,false);if(!access.ok)return{success:false,message:access.message};try{var ws=getWorkspaceFile(sid),s=ensureClientNotificationPreferencesSheetV12_(ws),v=s.getDataRange().getValues(),h=v[0]||[],idx={};h.forEach(function(k,i){idx[k]=i;});var email=String(access.user.email||'').trim().toLowerCase();for(var i=1;i<v.length;i++){if(String(v[i][idx.Business_ID]||'')===String(access.user.businessId||'')&&String(v[i][idx.Email]||'').trim().toLowerCase()===email)return{success:true,preferences:{features:v[i][idx.Feature_Updates]!==false,system:v[i][idx.System_Alerts]!==false,weekly:v[i][idx.Weekly_Reports]===true,security:v[i][idx.Security_Alerts]!==false,marketing:v[i][idx.Marketing_Updates]===true}};}return{success:true,preferences:{features:true,system:true,weekly:false,security:true,marketing:false}};}catch(e){return{success:false,message:'Could not load notification preferences.',code:'CLIENT_NOTIFICATION_PREFERENCES_LOAD_ERROR'};}}",
    "function saveClientNotificationPreferences(preferences,sid){var access=clientSettingsAccessV12_(sid,true);if(!access.ok)return{success:false,message:access.message};try{var ws=getWorkspaceFile(sid),s=ensureClientNotificationPreferencesSheetV12_(ws),v=s.getDataRange().getValues(),h=v[0]||[],idx={};h.forEach(function(k,i){idx[k]=i;});var email=String(access.user.email||'').trim().toLowerCase(),rowIndex=-1;for(var i=1;i<v.length;i++){if(String(v[i][idx.Business_ID]||'')===String(access.user.businessId||'')&&String(v[i][idx.Email]||'').trim().toLowerCase()===email){rowIndex=i;break;}}var p=preferences||{},values={Preference_ID:'NP_'+Utilities.getUuid().replace(/-/g,'').substring(0,20).toUpperCase(),Business_ID:String(access.user.businessId||''),Email:email,Feature_Updates:p.features!==false,System_Alerts:p.system!==false,Weekly_Reports:p.weekly===true,Security_Alerts:p.security!==false,Marketing_Updates:p.marketing===true,Updated_At:new Date().toISOString()};if(rowIndex<0)s.appendRow(h.map(function(k){return values[k]===undefined?'':values[k];}));else h.forEach(function(k,j){if(values[k]!==undefined)s.getRange(rowIndex+1,j+1).setValue(values[k]);});return{success:true,preferences:{features:values.Feature_Updates,system:values.System_Alerts,weekly:values.Weekly_Reports,security:values.Security_Alerts,marketing:values.Marketing_Updates}};}catch(e){return{success:false,message:'Could not save notification preferences.',code:'CLIENT_NOTIFICATION_PREFERENCES_SAVE_ERROR'};}}",
    "function getClientAppsData(sid){var user=getUserFromSession(sid);if(!user)return{success:false,message:'Session expired. Please login again.'};var allowed=Array.isArray(user.accessibleModules)?user.accessibleModules:[];var role=String(user.role||'staff').toLowerCase();var readOnly=user.isDemo==='YES'||['owner','admin'].indexOf(role)<0;var summary={};Object.keys(MODULES).forEach(function(name){var m=MODULES[name]||{};summary[name]={label:m.label||name,icon:m.icon||'fa-folder',isAccessible:allowed.indexOf(name)!==-1,count:0,isReadOnly:readOnly};try{var ss=getWorkspaceFile(sid),sheet=ss.getSheetByName(m.sheet);if(sheet)summary[name].count=Math.max(0,sheet.getLastRow()-1);}catch(ignore){}});return{success:true,summary:summary};}",
    "function clientMigrationActorV12_(sid){var u=getUserFromSession(sid);if(!u)return{ok:false,message:'Session expired. Please login again.'};var role=String(u.role||'').toLowerCase();if(['owner','admin'].indexOf(role)<0)return{ok:false,message:'Only the business owner or an administrator can manage migrated data.'};return{ok:true,user:u};}",
    "function ensureClientMigrationIndexV12_(ws){var s=ws.getSheetByName('Client_Migration_Index');if(!s){s=ws.insertSheet('Client_Migration_Index');s.appendRow(['Migration_ID','Module','Record_ID','Source_Workspace_ID','Source_Row','Migrated_At','Status','Archived_At','Record_Headers_JSON','Record_Row_JSON']);}return s;}",
    "function getClientMigratedData(sid){var a=clientMigrationActorV12_(sid);if(!a.ok)return{success:false,message:a.message};try{var ws=getWorkspaceFile(sid),s=ensureClientMigrationIndexV12_(ws),v=s.getDataRange().getValues(),h=v[0]||[],out=[];v.slice(1).forEach(function(r){var o={};h.forEach(function(k,i){o[k]=r[i];});if(String(o.Status||'active').toLowerCase()!=='deleted')out.push({migrationId:String(o.Migration_ID||''),module:String(o.Module||''),recordId:String(o.Record_ID||''),sourceRow:o.Source_Row,migratedAt:o.Migrated_At,status:String(o.Status||'active'),archivedAt:o.Archived_At});});return{success:true,records:out};}catch(e){return{success:false,message:e.message||'Migrated data could not be loaded.',code:'MIGRATED_DATA_LOAD_ERROR'};}}",
    "function setClientMigrationStatusV12(migrationId,status,sid){var a=clientMigrationActorV12_(sid);if(!a.ok)return{success:false,message:a.message};migrationId=String(migrationId||'').trim();status=String(status||'').toLowerCase();if(['archived','deleted'].indexOf(status)<0)return{success:false,message:'Invalid migration action.'};try{var ws=getWorkspaceFile(sid),idx=ensureClientMigrationIndexV12_(ws),v=idx.getDataRange().getValues(),h=v[0]||[],idCol=h.indexOf('Migration_ID'),statusCol=h.indexOf('Status'),archCol=h.indexOf('Archived_At'),moduleCol=h.indexOf('Module'),recordCol=h.indexOf('Record_ID'),headersCol=h.indexOf('Record_Headers_JSON'),rowCol=h.indexOf('Record_Row_JSON');for(var i=1;i<v.length;i++){if(String(v[i][idCol]||'')!==migrationId)continue;if(String(v[i][statusCol]||'').toLowerCase()==='deleted')return{success:false,message:'This migrated record has already been deleted.'};var moduleName=String(v[i][moduleCol]||''),recordId=String(v[i][recordCol]||''),sheet=ws.getSheetByName(MODULES[moduleName]&&MODULES[moduleName].sheet),idHeader=({Finance:'Transaction_ID',Ecommerce:'Order_ID',Sales:'Deal_ID',CRM:'Contact_ID',HR:'Employee_ID',Logistics:'Shipment_ID',Tax:'Tax_ID',Agro:'Record_ID',Productivity:'Record_ID',POS:'Transaction_ID',Attendance:'Record_ID',Warehouse:'Item_ID'})[moduleName];if(sheet&&recordId&&idHeader){var shv=sheet.getDataRange().getValues(),sh=shv[0]||[],ic=sh.indexOf(idHeader);if(ic>=0)for(var j=1;j<shv.length;j++)if(String(shv[j][ic]||'')===recordId){sheet.deleteRow(j+1);break;}}if(status==='archived'){var archive=ws.getSheetByName('Client_Migration_Archive');if(!archive){archive=ws.insertSheet('Client_Migration_Archive');archive.appendRow(['Migration_ID','Module','Record_ID','Archived_At','Record_Headers_JSON','Record_Row_JSON']);}archive.appendRow([migrationId,moduleName,recordId,new Date().toISOString(),String(v[i][headersCol]||'[]'),String(v[i][rowCol]||'[]')]);}if(statusCol>=0)idx.getRange(i+1,statusCol+1).setValue(status);if(archCol>=0)idx.getRange(i+1,archCol+1).setValue(status==='archived'?new Date().toISOString():'');return{success:true,status:status};}return{success:false,message:'Migrated record not found.'};}catch(e){return{success:false,message:e.message||'Migrated data could not be updated.',code:'MIGRATED_DATA_UPDATE_ERROR'};}}",
    "function restoreClientMigratedData(migrationId,sid){var a=clientMigrationActorV12_(sid);if(!a.ok)return{success:false,message:a.message};migrationId=String(migrationId||'').trim();try{var ws=getWorkspaceFile(sid),idx=ensureClientMigrationIndexV12_(ws),v=idx.getDataRange().getValues(),h=v[0]||[],idCol=h.indexOf('Migration_ID'),statusCol=h.indexOf('Status'),moduleCol=h.indexOf('Module'),headersCol=h.indexOf('Record_Headers_JSON'),rowCol=h.indexOf('Record_Row_JSON');for(var i=1;i<v.length;i++){if(String(v[i][idCol]||'')!==migrationId)continue;if(String(v[i][statusCol]||'').toLowerCase()!=='archived')return{success:false,message:'Only archived migrated records can be restored.'};var moduleName=String(v[i][moduleCol]||''),sheet=ws.getSheetByName(MODULES[moduleName]&&MODULES[moduleName].sheet);if(!sheet)return{success:false,message:'Target module sheet is unavailable.'};var row=JSON.parse(String(v[i][rowCol]||'[]')),headers=JSON.parse(String(v[i][headersCol]||'[]')),hc=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0],idHeader=({Finance:'Transaction_ID',Ecommerce:'Order_ID',Sales:'Deal_ID',CRM:'Contact_ID',HR:'Employee_ID',Logistics:'Shipment_ID',Tax:'Tax_ID',Agro:'Record_ID',Productivity:'Record_ID',POS:'Transaction_ID',Attendance:'Record_ID',Warehouse:'Item_ID'})[moduleName],ic=hc.indexOf(idHeader),restoredId=headers.indexOf(idHeader)>=0?String(row[headers.indexOf(idHeader)]||''):'';if(ic>=0&&restoredId&&sheet.getDataRange().getValues().slice(1).some(function(r){return String(r[ic]||'')===restoredId;}))return{success:false,message:'A record with this ID already exists.'};sheet.appendRow(hc.map(function(header){var k=headers.indexOf(header);return k>=0?row[k]:'';}));if(statusCol>=0)idx.getRange(i+1,statusCol+1).setValue('active');if(headersCol>=0&&rowCol>=0){idx.getRange(i+1,headersCol+1).setValue(JSON.stringify(hc));idx.getRange(i+1,rowCol+1).setValue(JSON.stringify(hc.map(function(header){var k=headers.indexOf(header);return k>=0?row[k]:'';})));}return{success:true,status:'active'};}return{success:false,message:'Archived migrated record not found.'};}catch(e){return{success:false,message:e.message||'Migrated record could not be restored.',code:'MIGRATED_DATA_RESTORE_ERROR'};}}",
    "function getClientSettingsData(sid){var access=clientSettingsAccessV12_(sid,false);if(!access.ok)return{success:false,message:access.message};var bp=(typeof CLIENT_CONFIG!=='undefined'&&CLIENT_CONFIG.businessProfile)||{};var settings={Business_Name:String(bp.businessName||CLIENT_CONFIG.clientName||''),Business_Type:String(bp.businessType||''),Phone:String(bp.ownerPhone||''),Website:String(bp.website||''),Address:String(bp.address||''),City:String(bp.city||''),Country:String(bp.country||''),Currency:String(bp.currency||'NGN'),Business_Description:String(bp.businessDescription||''),Display_Name:String(CLIENT_CONFIG.clientName||bp.businessName||''),Primary_Color:String(CLIENT_CONFIG.primaryColor||'#2E7D32'),Logo_Url:String(CLIENT_CONFIG.logoUrl||'')};var ws=getWorkspaceFile(sid),sheet=ws.getSheetByName('Client_Info');if(sheet&&sheet.getLastRow()>0){var rows=sheet.getDataRange().getValues(),h=rows[0]||[],p=h.indexOf('Property'),v=h.indexOf('Value');if(p>=0&&v>=0)for(var i=1;i<rows.length;i++){var k=String(rows[i][p]||'').trim(),value=rows[i][v];if(!k||value===undefined||value===null||String(value).trim()==='')continue;if(['Business_Name','Business_Type','Phone','Website','Address','City','Country','Currency','Business_Description','Display_Name','Primary_Color','Logo_Url'].indexOf(k)>=0)settings[k]=value;}}var u=access.user||{};u={email:String(u.email||bp.ownerEmail||CLIENT_CONFIG.email||''),name:String(u.name||bp.ownerName||''),phone:String(u.phone||bp.ownerPhone||''),role:String(u.role||'owner'),businessId:String(u.businessId||CLIENT_CONFIG.businessId),businessName:String(u.businessName||settings.Business_Name||CLIENT_CONFIG.clientName||'')};return{success:true,settings:settings,user:u,source:{canonical:'BizOS master business/user data',custom:'client workspace settings'}};}",
    "function saveClientSettings(data,sid){var access=clientSettingsAccessV12_(sid,true);if(!access.ok)return{success:false,message:access.message};data=data||{};var allowed=['Business_Name','Business_Type','Phone','Website','Address','City','Country','Currency','Business_Description','Display_Name','Primary_Color','Logo_Url'],clean={};allowed.forEach(function(k){if(data[k]!==undefined)clean[k]=String(data[k]||'').trim();});if(!clean.Business_Name)return{success:false,message:'Business name is required.'};if(clean.Currency&&!/^[A-Z]{3}$/.test(clean.Currency))return{success:false,message:'Currency must be a 3-letter code.'};if(clean.Primary_Color&&!/^#[0-9A-Fa-f]{6}$/.test(clean.Primary_Color))return{success:false,message:'Primary colour must be a 6-digit hex colour.'};var ws=getWorkspaceFile(sid),sheet=ws.getSheetByName('Client_Info');if(!sheet)return{success:false,message:'Client settings storage is unavailable.'};var rows=sheet.getDataRange().getValues(),h=rows[0]||[],p=h.indexOf('Property'),v=h.indexOf('Value'),idx={};if(p<0||v<0)return{success:false,message:'Client settings schema is unavailable.'};for(var i=1;i<rows.length;i++){var key=String(rows[i][p]||'').trim();if(key)idx[key]=i+1;}Object.keys(clean).forEach(function(k){if(idx[k])sheet.getRange(idx[k],v+1).setValue(clean[k]);else sheet.appendRow([k,clean[k]]);});return{success:true,settings:clean};}",
    "function updateClientProfile(data,sid){var access=clientSettingsAccessV12_(sid,false);if(!access.ok)return{success:false,message:access.message};data=data||{};var sheet=getOrCreateUserSheet(),rows=sheet.getDataRange().getValues(),h=rows[0]||[],emailCol=h.indexOf('Email'),nameCol=h.indexOf('Name'),phoneCol=h.indexOf('Phone'),roleCol=h.indexOf('Role'),businessIdCol=h.indexOf('Business_ID'),businessNameCol=h.indexOf('Business_Name'),statusCol=h.indexOf('Status'),clientCol=h.indexOf('Is_Client');if(emailCol<0)return{success:false,message:'User profile schema is unavailable.'};var wanted=String(access.user.email||'').toLowerCase(),row=-1;for(var i=1;i<rows.length;i++){if(String(rows[i][emailCol]||'').toLowerCase()===wanted){row=i+1;break;}}if(row<0){var newRow=[];for(var j=0;j<h.length;j++)newRow.push('');if(emailCol>=0)newRow[emailCol]=wanted;if(nameCol>=0)newRow[nameCol]=String(data.name||access.user.name||'');if(phoneCol>=0)newRow[phoneCol]=String(data.phone||access.user.phone||'');if(roleCol>=0)newRow[roleCol]=String(access.user.role||'owner');if(businessIdCol>=0)newRow[businessIdCol]=String(CLIENT_CONFIG.businessId||'');if(businessNameCol>=0)newRow[businessNameCol]=String(CLIENT_CONFIG.clientName||'');if(statusCol>=0)newRow[statusCol]='active';if(clientCol>=0)newRow[clientCol]='YES';sheet.appendRow(newRow);return{success:true};}if(nameCol>=0)sheet.getRange(row,nameCol+1).setValue(String(data.name||''));if(phoneCol>=0)sheet.getRange(row,phoneCol+1).setValue(String(data.phone||''));return{success:true};}",
    "function changeClientPassword(currentPassword,newPassword,sid){var access=clientSettingsAccessV12_(sid,false);if(!access.ok)return{success:false,message:access.message};return changeUserPassword(String(access.user.email||''),String(currentPassword||''),String(newPassword||''),sid);}"
  ].join(String.fromCharCode(10));
  runtime = runtime.replace(/\nfunction doGet\(e\)\{/, '\n'+businessCenterSource+'\n'+businessCenterBridge+'\nfunction doGet(e){');
  runtime += String.fromCharCode(10) + [
    'function renderV12RuntimeTrace_(token){',
    '  try{',
    '    var validation=UrlFetchApp.fetch(CLIENT_CONFIG.masterApiUrl,{method:"post",contentType:"application/json",muteHttpExceptions:true,payload:JSON.stringify({action:"validateV12RuntimeTrace",clientId:String(CLIENT_CONFIG.clientId||""),token:String(token||"")})});',
    '    var body=validation.getContentText()||"{}";',
    '    if(validation.getResponseCode()<200||validation.getResponseCode()>=300)return HtmlService.createHtmlOutput("<h2>Runtime trace unavailable</h2><pre>"+String(body).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")+"</pre>");',
    '    var auth=JSON.parse(body||"{}");',
    '    if(!auth||!auth.success)return HtmlService.createHtmlOutput("<h2>Runtime trace denied</h2><pre>"+String(auth&&auth.message||"Diagnostic token is invalid or expired.").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")+"</pre>");',
    '    var modulesType=typeof MODULES,moduleKeys=[];',
    '    if(modulesType==="object"&&MODULES)moduleKeys=Object.keys(MODULES);',
    '    var entitlement={success:false};',
    '    try{entitlement={success:true,modules:clientModulesForEntitlementV12_("owner","sovereign")};}catch(entErr){entitlement={success:false,message:String(entErr&&entErr.message||entErr)};}',
    '    var result={success:true,trace:"V12_CLIENT_RUNTIME_PROBE",clientId:String(CLIENT_CONFIG.clientId||""),packageVersion:String(CLIENT_V12_PACKAGE_VERSION_||""),runtime:{clientConfig:typeof CLIENT_CONFIG!=="undefined",modulesType:modulesType,moduleCount:moduleKeys.length,moduleKeys:moduleKeys,findClientUser:typeof findClientUser_==="function",createSession:typeof createClientSessionV11_==="function",entitlementBridge:typeof clientModulesForEntitlementV12_==="function",authenticateClient:typeof authenticateClient==="function",dashboardData:typeof getClientDashboardData==="function"},entitlement:entitlement};',
    '    var json=JSON.stringify(result,null,2).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");',
    '    return HtmlService.createHtmlOutput("<!doctype html><html><head><meta name=\\\"viewport\\\" content=\\\"width=device-width,initial-scale=1\\\"><title>BizOS Runtime Trace</title></head><body style=\\\"font-family:Arial,sans-serif;padding:24px\\\"><h2>BizOS V12 Runtime Trace</h2><pre style=\\\"white-space:pre-wrap\\\">"+json+"</pre></body></html>").setTitle("BizOS Runtime Trace");',
    '  }catch(e){return HtmlService.createHtmlOutput("<h2>Runtime trace failed</h2><pre>"+String(e&&e.message||e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")+"</pre>");}',
    '}'
  ].join(String.fromCharCode(10));
  // V12 security bridge provides the authoritative client module summary.
  runtime += String.fromCharCode(10) + "function getClientDashboard(sid,forceRefresh){\n  try{\n    var u=getClientUserFromSession(sid);\n    if(!u)return{success:false,message:'Session expired or client license is no longer valid.',code:'CLIENT_LICENSE_INVALID'};\n    var cacheKey='BIZOS_CLIENT_DASH_'+String(CLIENT_CONFIG.clientId).replace(/[^A-Za-z0-9_]/g,'_');\n    if(!forceRefresh){var cached=CacheService.getScriptCache().get(cacheKey);if(cached){try{return{success:true,dashboard:JSON.parse(cached),cached:true};}catch(ignore){}}}\n    var ss=getWorkspaceFile(sid),finance=ss.getSheetByName('Financial_Data'),now=new Date(),clientCurrency=String(CLIENT_CONFIG.currency||'NGN').toUpperCase();try{var cis=ss.getSheetByName('Client_Info');if(cis&&cis.getLastRow()>0){var civ=cis.getDataRange().getValues(),cih=civ[0]||[],cp=cih.indexOf('Property'),cv=cih.indexOf('Value');if(cp>=0&&cv>=0)for(var ci=1;ci<civ.length;ci++)if(String(civ[ci][cp]||'').trim()==='Currency'&&String(civ[ci][cv]||'').trim()){clientCurrency=String(civ[ci][cv]).trim().toUpperCase();break;}}}catch(ignoreCurrency){}\n    function empty(){return{revenue:0,expenses:0,count:0,categories:{}};}\n    function periodBounds_(p,n){var d=new Date(n);d.setHours(0,0,0,0),e=new Date(d),ps=new Date(d),pe=new Date(d);switch(String(p||'this_month')){case'today':e.setDate(e.getDate()+1);ps.setDate(ps.getDate()-1);pe=new Date(d);break;case'this_week':d.setDate(d.getDate()-d.getDay());e=new Date(d);e.setDate(e.getDate()+7);ps=new Date(d);ps.setDate(ps.getDate()-7);pe=new Date(d);break;case'this_year':d=new Date(d.getFullYear(),0,1);e=new Date(d.getFullYear()+1,0,1);ps=new Date(d.getFullYear()-1,0,1);pe=new Date(d);break;case'last_month':d=new Date(d.getFullYear(),d.getMonth()-1,1);e=new Date(d.getFullYear(),d.getMonth(),1);ps=new Date(d.getFullYear(),d.getMonth()-2,1);pe=new Date(d.getFullYear(),d.getMonth()-1,1);break;default:d=new Date(d.getFullYear(),d.getMonth(),1);e=new Date(d.getFullYear(),d.getMonth()+1,1);ps=new Date(d.getFullYear(),d.getMonth()-1,1);pe=new Date(d.getFullYear(),d.getMonth(),1);}return{start:d,end:e,previousStart:ps,previousEnd:pe};}\n    function readFinance_(sh,b){var out=empty();if(!sh||sh.getLastRow()<=1)return out;var a=sh.getDataRange().getValues(),h=a[0]||[],dc=h.indexOf('Date'),tc=h.indexOf('Type'),ac=h.indexOf('Amount'),cc=h.indexOf('Category');if(tc<0||ac<0)return out;for(var i=1;i<a.length;i++){var dt=dc>=0?new Date(a[i][dc]):new Date();if(dc>=0&&(isNaN(dt.getTime())||dt<b.start||dt>=b.end))continue;var n=Math.abs(Number(a[i][ac])||0),t=String(a[i][tc]||'').toLowerCase();out.count++;if(t==='revenue'||t==='income')out.revenue+=n;else if(t==='expense'||t==='expenses'){out.expenses+=n;var cat=String(cc>=0?a[i][cc]||'other':'other');out.categories[cat]=(out.categories[cat]||0)+n;}}return out;}\n    var period='this_month',b=periodBounds_(period,now),cur=readFinance_(finance,b),prev=readFinance_(finance,{start:b.previousStart,end:b.previousEnd}),pct=function(a,z){a=Number(a)||0;z=Number(z)||0;if(z===0)return a===0?0:100;return((a-z)/Math.abs(z))*100;},profit=cur.revenue-cur.expenses,prevProfit=prev.revenue-prev.expenses,margin=cur.revenue?profit/cur.revenue*100:0,prevMargin=prev.revenue?prevProfit/prev.revenue*100:0;\n    var moduleNames=Array.isArray(u.accessibleModules)?u.accessibleModules.slice():[],moduleSummary={},moduleCounts={},operational={},totalOperational=0,modulesWithData=0,teamMembers=0;\n    var defs={Finance:{count:'records'},Ecommerce:{count:'orders',value:'orderValue',sum:'Total_Amount'},Sales:{count:'deals',value:'pipelineValue',sum:'Total_Amount'},CRM:{count:'contacts'},HR:{count:'employees',value:'payroll',sum:'Salary'},Logistics:{count:'shipments'},Tax:{count:'records',value:'taxAmount',sum:'Tax_Amount'},Agro:{count:'records',revenue:'Revenue',costs:'Total_Cost'},Productivity:{count:'tasks',hours:'Hours_Spent'},POS:{count:'transactions',value:'sales',sum:'Total_Amount',items:'Quantity'},Attendance:{count:'records'},Warehouse:{count:'products',stockUnits:'Quantity',inventoryValue:'__inventory__'}};\n    moduleNames.forEach(function(name){var def=MODULES[name]||{},sh=ss.getSheetByName(def.sheet),count=sh&&sh.getLastRow()>1?sh.getLastRow()-1:0;moduleCounts[name]=Math.max(0,count);moduleSummary[name]={label:def.label||name,icon:def.icon||'',count:Math.max(0,count),isAccessible:true,isReadOnly:false};if(count>0)modulesWithData++;totalOperational+=Math.max(0,count);var spec=defs[name];if(!spec||!sh||count<=0)return;var vals=sh.getDataRange().getValues(),hh=vals[0]||[],idx=function(k){return hh.indexOf(k);},sum=function(k){var j=idx(k),z=0;if(j<0)return 0;for(var q=1;q<vals.length;q++)z+=Number(vals[q][j])||0;return z;},o={};o[spec.count]=count;if(spec.value)o[spec.value]=sum(spec.sum);if(name==='HR')teamMembers=count;if(spec.revenue)o.revenue=sum(spec.revenue);if(spec.costs)o.costs=sum(spec.costs);if(spec.hours)o.hours=sum(spec.hours);if(spec.items)o.items=sum(spec.items);if(spec.stockUnits)o.stockUnits=sum(spec.stockUnits);if(spec.inventoryValue){var qi=idx('Quantity'),ui=idx('Unit_Cost'),iv=0;if(qi>=0&&ui>=0)for(var z=1;z<vals.length;z++)iv+=(Number(vals[z][qi])||0)*(Number(vals[z][ui])||0);o.inventoryValue=iv;var mi=idx('Min_Stock'),low=0;if(qi>=0&&mi>=0)for(var y=1;y<vals.length;y++)if((Number(vals[y][qi])||0)<(Number(vals[y][mi])||0))low++;o.lowStock=low;}operational[name]=o;});\n    var bc=ss.getSheetByName('BusinessCenter_Sales'),bcCount=bc&&bc.getLastRow()>1?bc.getLastRow()-1:0;if(bcCount){var bv=bc.getDataRange().getValues(),bh=bv[0]||[],bi=function(k){return bh.indexOf(k);},sumBc=function(k){var j=bi(k),z=0;if(j<0)return 0;for(var q=1;q<bv.length;q++)z+=Number(bv[q][j])||0;return z;};operational.BusinessCenter={sales:bcCount,salesValue:sumBc('Total'),cashCollected:sumBc('Amount_Paid'),receivables:sumBc('Balance_Due')};}\n    var businessSnapshot={activeModules:moduleNames.length,modulesWithData:modulesWithData,activeRecords:totalOperational+bcCount,commerceRecords:(moduleCounts.Ecommerce||0)+(moduleCounts.POS||0)+bcCount,teamMembers:teamMembers,operational:operational};\n    var daily={},recent=[];if(finance&&finance.getLastRow()>1){var fv=finance.getDataRange().getValues(),fh=fv[0]||[],fd=fh.indexOf('Date'),ft=fh.indexOf('Type'),fa=fh.indexOf('Amount'),fx=fh.indexOf('Description'),fc=fh.indexOf('Category');for(var i=1;i<fv.length;i++){var raw=fv[i][fd],dt=raw instanceof Date?raw:new Date(raw);if(isNaN(dt.getTime()))continue;var key=Utilities.formatDate(dt,Session.getScriptTimeZone(),'yyyy-MM-dd'),amt=Math.abs(Number(fv[i][fa])||0),typ=String(fv[i][ft]||'').toLowerCase();daily[key]=daily[key]||{revenue:0,expenses:0};if(typ==='revenue'||typ==='income')daily[key].revenue+=amt;else if(typ==='expense'||typ==='expenses')daily[key].expenses+=amt;}for(var j=Math.max(1,fv.length-11);j<fv.length;j++){recent.push({date:fv[j][fd],description:fd>=0?fv[j][fx>=0?fx:ft]:'',amount:fa>=0?fv[j][fa]:0,type:ft>=0?fv[j][ft]:'',category:fc>=0?fv[j][fc]:''});}}\n    var labels=[],rv=[],ex=[];for(var d=29;d>=0;d--){var day=new Date(now);day.setDate(now.getDate()-d);var key=Utilities.formatDate(day,Session.getScriptTimeZone(),'yyyy-MM-dd');labels.push(key);rv.push(daily[key]?daily[key].revenue:0);ex.push(daily[key]?daily[key].expenses:0);}var cats=Object.keys(cur.categories).sort(function(a,z){return cur.categories[z]-cur.categories[a];}).slice(0,6);var weekly={};for(var w=7;w>=0;w--){var wsDate=new Date(now);wsDate.setHours(0,0,0,0);wsDate.setDate(wsDate.getDate()-w*7);var wk=Utilities.formatDate(wsDate,Session.getScriptTimeZone(),'yyyy-MM-dd'),wr=0,we=0;for(var di=0;di<7;di++){var dk=new Date(wsDate);dk.setDate(wsDate.getDate()+di);var dkey=Utilities.formatDate(dk,Session.getScriptTimeZone(),'yyyy-MM-dd');if(daily[dkey]){wr+=daily[dkey].revenue||0;we+=daily[dkey].expenses||0;}}weekly[wk]={revenue:wr,expenses:we};}\n    var activity=[];try{var as=ss.getSheetByName('Client_Activity');if(as&&as.getLastRow()>1){var av=as.getDataRange().getValues(),ah=av[0]||[],ai=function(k){return ah.indexOf(k);};for(var airow=Math.max(1,av.length-11);airow<av.length;airow++)activity.push({date:av[airow][ai('Created_At')],title:av[airow][ai('Summary')]||'Workspace activity',type:av[airow][ai('Event_Type')]||'workspace'});}}catch(ignoreActivity){}if(!activity.length)recent.forEach(function(x){activity.push({date:x.date,title:x.description||x.type||'Financial entry',type:String(x.type||'').toLowerCase(),amount:x.amount});});\n    var alerts=[],recommendations=[];if(cur.revenue===0)recommendations.push({type:'info',title:'Record your income',message:'Add revenue entries so BizOS can calculate profit and business trends.'});if(cur.expenses>cur.revenue&&cur.expenses>0)alerts.push({type:'warning',priority:'high',title:'Expenses are above revenue',message:'Recorded expenses exceed recorded revenue for this period.'});if(operational.Warehouse&&operational.Warehouse.lowStock>0)alerts.push({type:'warning',priority:'high',title:'Low stock',message:String(operational.Warehouse.lowStock)+' warehouse item(s) are below minimum stock.'});if(operational.Ecommerce&&operational.Ecommerce.orders>0)alerts.push({type:'info',priority:'medium',title:'Ecommerce activity',message:String(operational.Ecommerce.orders)+' ecommerce order(s) recorded.'});if(operational.POS&&operational.POS.transactions>0)alerts.push({type:'info',priority:'low',title:'POS activity',message:String(operational.POS.transactions)+' POS transaction(s) recorded.'});\n    var expenseScore=cur.revenue>0?Math.max(0,Math.min(100,(1-(cur.expenses/cur.revenue))*100)):(cur.expenses===0?100:0),marginScore=Math.max(0,Math.min(100,margin*5)),activityScore=moduleNames.length>0?Math.min(100,(modulesWithData/moduleNames.length)*100):(totalOperational>0?50:0),healthScore=Math.round((Math.max(0,expenseScore)+marginScore+activityScore)/3);if(cur.revenue===0)recommendations.push({priority:'high',type:'warning',title:'No revenue recorded',message:'Review finance, ecommerce, sales and POS activity.'});if(margin>0&&margin<10)recommendations.push({priority:'medium',type:'info',title:'Profit margin is below 10%',message:'Review pricing, costs and operating expenses.'});if(moduleNames.length>0&&modulesWithData===0)recommendations.push({priority:'medium',type:'info',title:'No operational module records yet',message:'Start with the modules that match your business workflow.'});var health={overall:Math.max(0,Math.min(100,healthScore)),metrics:[{label:'Revenue',value:cur.revenue>0?'Recorded':'None'},{label:'Expense control',value:cur.revenue>0?((cur.expenses/cur.revenue)*100).toFixed(1)+'% of revenue':'No revenue'},{label:'Profit margin',value:margin.toFixed(1)+'%'},{label:'Business activity',value:String(totalOperational+bcCount)+' records across '+String(modulesWithData)+' active modules'}],recommendations:recommendations};\n    var dashboard={kpis:{revenue:{value:cur.revenue,change:pct(cur.revenue,prev.revenue)},expenses:{value:cur.expenses,change:pct(cur.expenses,prev.expenses)},netProfit:{value:profit,change:pct(profit,prevProfit)},profitMargin:{value:margin,change:margin-prevMargin}},trends:{categoryBreakdown:cur.categories,daily:daily,weekly:weekly},moduleSummary:moduleSummary,businessSnapshot:businessSnapshot,moduleCounts:moduleCounts,counts:{finance:cur.count,products:(moduleCounts.Ecommerce||0)+(moduleCounts.POS||0),modules:moduleNames.length,team:(ss.getSheetByName('Client_Team')&&ss.getSheetByName('Client_Team').getLastRow()>1?ss.getSheetByName('Client_Team').getLastRow()-1:0)},charts:{revenueVsExpenses:{labels:labels,revenue:rv,expenses:ex},expenseBreakdown:{labels:cats,values:cats.map(function(k){return cur.categories[k];})}},health:health,alerts:alerts,recommendations:recommendations,recentActivity:activity.slice(0,10),recentTransactions:recent.slice(0,10),clientName:CLIENT_CONFIG.clientName,primaryColor:CLIENT_CONFIG.primaryColor,logoUrl:CLIENT_CONFIG.logoUrl,workspaceId:CLIENT_CONFIG.sheetId,userEmail:u.email,currency:clientCurrency,generatedAt:new Date().toISOString(),period:period};\n    try{CacheService.getScriptCache().put(cacheKey,JSON.stringify(dashboard),60);}catch(ignoreCache){}return{success:true,dashboard:dashboard,cached:false};\n  }catch(e){return{success:false,message:e.message||'Dashboard could not be loaded.',code:'CLIENT_DASHBOARD_ERROR'};}\n}\nfunction getClientDashboardData(sid,forceRefresh){return getClientDashboard(sid,forceRefresh);}" + String.fromCharCode(10);  // F security bridge is appended before the V12 doGet isolation so its functions are part of the generated Code package.
  var doGetStart = runtime.indexOf('function doGet(e){');
  var includeStart = runtime.indexOf('function include', doGetStart);
  if (doGetStart < 0 || includeStart < 0) throw new Error('Unable to isolate the V12 client doGet runtime.');
  runtime = runtime.slice(0, doGetStart) +
    'function doGet(e){var traceToken=String(e&&e.parameter&&e.parameter.v12TraceToken||"").trim();if(traceToken)return renderV12RuntimeTrace_(traceToken);var t=HtmlService.createTemplateFromFile("ClientShellV12");return t.evaluate().setTitle(CLIENT_CONFIG.clientName+" - BizOS").addMetaTag("viewport","width=device-width,initial-scale=1").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}\n' +
    runtime.slice(includeStart);

  var modulesSource = String(getMasterSourceForV12_('Modules') || '');
  var recordsSource = String(getMasterSourceForV12_('Records') || '');
  var businessCenterServerSource = String(getMasterSourceForV12_('BusinessHubGS') || '');
  if(!modulesSource) throw new Error('Modules source is missing.');
  if(!recordsSource) throw new Error('Records source is missing.');
  if(!businessCenterServerSource) throw new Error('BusinessHubGS source is missing.');
  // Proven BizOS server dependencies used by the real Business Center and module workspaces.
  modulesSource = modulesSource.replace(/function\s+getModuleSummary\s*\(/g,'function getModuleSummaryLegacy_(');

  var files = [
    {name:'Code',type:'SERVER_JS',source:runtime},
    {name:'Modules',type:'SERVER_JS',source:modulesSource},
    {name:'Records',type:'SERVER_JS',source:recordsSource},
    {name:'ClientShellV12',type:'HTML',source:buildClientShellV12_(config)},
    {name:'ClientBusinessCenterV12',type:'HTML',source:buildClientBusinessCenterV12_()},
    {name:'ClientFinanceV12',type:'HTML',source:buildClientFinanceV12_()},
    {name:'ClientStaffV12',type:'HTML',source:buildClientStaffV12_()},
    {name:'ClientAppsV12',type:'HTML',source:buildClientAppsV12_()},
    {name:'ClientModuleV12',type:'HTML',source:buildClientModuleV12_()},{name:'ClientRecordsV12',type:'HTML',source:buildClientRecordsV12_()},
    {name:'ClientSettingsV12',type:'HTML',source:buildClientSettingsV12_()},
    {name:'ClientProfileV12',type:'HTML',source:buildClientProfileV12_()},
    {name:'ClientNotificationsV12',type:'HTML',source:buildClientNotificationsV12_()},
    {name:'ClientBillingV12',type:'HTML',source:buildClientBillingV12_()},
    {name:'ClientStylesV12',type:'HTML',source:buildClientStylesV12_()},
    {name:'appsscript',type:'JSON',source:JSON.stringify(buildClientManifestV12_())}
  ];

  return {
    success:true,version:CLIENT_V12_PACKAGE_VERSION_,release:CLIENT_V12_PACKAGE_RELEASE_,provenance:CLIENT_V12_PACKAGE_PROVENANCE_,config:config,files:files,
    source:{clientUi:'V12-shell-with-BizOS-business-center',copiedMasterUi:['BusinessHubHTML'],reusedBackendRuntime:'V11-safe-runtime',businessCenterBackend:'BusinessHubGS-embedded-in-Code-and-safe-wrapped',businessCenterUi:'BizOS-BusinessHubHTML-adapted'}
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
          countryCol=h.indexOf('Country'), websiteCol=h.indexOf('Website'), currencyCol=h.indexOf('Currency'), tierCol=h.indexOf('Subscription_Tier');
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
        out.currency=currencyCol>=0?String(row[currencyCol]||'').trim().toUpperCase():'';
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
  var source = String(getMasterSourceForV12_('BusinessHubHTML') || '');
  if(!source) throw new Error('BusinessHubHTML source is missing.');

  // Reuse the actual BizOS Business Center UI. Only the API boundary and
  // client navigation are adapted; product wording/layout/features remain intact.
  var adapter = [
    '<script>',
    '(function(){',
    '  var actionMap={',
    '    getBusinessCenterData:"getClientBusinessCenterData",',
    '    getBusinessCenterConnectedData:"getClientBusinessCenterConnectedData",',
    '    getBusinessCenterProductImages:"getClientBusinessCenterProductImages",',
    '    saveBusinessCenterProduct:"saveClientBusinessCenterProduct",',
    '    saveBusinessCenterProductsBulk:"saveClientBusinessCenterProductsBulk",',
    '    saveBusinessCenterCustomer:"saveClientBusinessCenterCustomer",',
    '    lookupBusinessCenterProduct:"lookupClientBusinessCenterProduct",',
    '    completeBusinessCenterSale:"completeClientBusinessCenterSale",',
    '    getBusinessCenterCustomerLedger:"getClientBusinessCenterCustomerLedger",',
    '    recordBusinessCenterCustomerPayment:"recordClientBusinessCenterCustomerPayment",',
    '    getBusinessCenterSales:"getClientBusinessCenterSales",',
    '    getBusinessCenterSaleDetails:"getClientBusinessCenterSaleDetails",',
    '    updateBusinessCenterProduct:"updateClientBusinessCenterProduct",',
    '    adjustBusinessCenterStock:"adjustClientBusinessCenterStock",',
    '    getBusinessCenterStockMovements:"getClientBusinessCenterStockMovements",',
    '    getBusinessCenterFinanceData:"getClientBusinessCenterFinanceData",',
    '    recordBusinessCenterExpense:"recordClientBusinessCenterExpense"',
    '  };',
    '  window.callGoogleScript=function(action,payload,done,fail){',
    '    var name=actionMap[action]||action;',
    '    if(typeof google!=="undefined"&&google.script&&google.script.run){',
    '      google.script.run.withSuccessHandler(done||function(){}).withFailureHandler(fail||function(){} )[name].apply(null,payload||[]);',
    '    }else if(typeof window.bizosClientCall==="function"){',
    '      window.bizosClientCall(name,payload||[],done,fail);',
    '    }else if(fail)fail({message:"BizOS client runtime is still loading."});',
    '  };',
    '  window.bcV12Init=function(){',
    '    if(typeof window.showBusinessCenter==="function")window.showBusinessCenter();',
    '    else if(typeof window.loadBusinessCenter==="function")window.loadBusinessCenter();',
    '  };',
    '  window.closeBusinessCenter=function(){',
    '    if(typeof window.selectView==="function")window.selectView("overview");',
    '  };',
    '})();',
    '</script>'
  ].join(String.fromCharCode(10));
  return source + String.fromCharCode(10) + adapter;
}

function buildClientFinanceV12_() {
  var source = String(getMasterSourceForV12_('ClientFinanceV12') || '');
  if(!source) throw new Error('ClientFinanceV12 source is missing.');
  return source;
}

function buildClientAppsV12_() { var source=String(getMasterSourceForV12_('ClientAppsV12')||''); if(!source) throw new Error('ClientAppsV12 source is missing.'); return source; }
function buildClientModuleV12_() { var source=String(getMasterSourceForV12_('ClientModuleV12')||''); if(!source) throw new Error('ClientModuleV12 source is missing.'); return source; }
function buildClientRecordsV12_() { var source=String(getMasterSourceForV12_('ClientRecordsV12')||''); if(!source) throw new Error('ClientRecordsV12 source is missing.'); return source; }
function buildClientSettingsV12_() { var source=String(getMasterSourceForV12_('ClientSettingsV12')||''); if(!source) throw new Error('ClientSettingsV12 source is missing.'); return source; }
function buildClientProfileV12_() { var source=String(getMasterSourceForV12_('ClientProfileV12')||''); if(!source) throw new Error('ClientProfileV12 source is missing.'); return source; }
function buildClientNotificationsV12_() { var source=String(getMasterSourceForV12_('ClientNotificationsV12')||''); if(!source) throw new Error('ClientNotificationsV12 source is missing.'); return source; }
function buildClientBillingV12_() { var source=String(getMasterSourceForV12_('ClientBillingV12')||''); if(!source) throw new Error('ClientBillingV12 source is missing.'); return source; }
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
    : getActiveClientForV12Test_();

  if (!client) {
    throw new Error(targetBusinessId
      ? 'No active client deployment was found for business ID: ' + targetBusinessId
      : 'No active active client deployment was found in the Clients registry.');
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
    applicationUrl:client.webAppUrl || '',
    featureProfile:(function(){try{return JSON.parse(String(client.Feature_Profile_JSON||'{}'))||{};}catch(e){return {};}})()
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

function getActiveClientForV12Test_() {
  var ss = getBizOSMasterSpreadsheet_();
  var sheet = ss.getSheetByName('Clients');
  if (!sheet || sheet.getLastRow() < 2) return null;

  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function(h){ return String(h || '').trim(); });
  var statusCol = headers.indexOf('Status');
  var businessIdCol = headers.indexOf('Business_ID');
  if (statusCol < 0) throw new Error('Client registry is missing Status.');

  // Prefer the most recently updated active deployment when the caller
  // does not specify a Business_ID. This keeps the smoke test generic and
  // avoids coupling it to a test email or a specific customer.
  var updatedCol = headers.indexOf('Updated_At');
  var createdCol = headers.indexOf('Created_At');
  var best = null, bestTime = -1;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (String(row[statusCol] || '').trim().toLowerCase() !== 'active') continue;
    var stamp = updatedCol >= 0 ? row[updatedCol] : (createdCol >= 0 ? row[createdCol] : '');
    var time = stamp ? new Date(stamp).getTime() : i;
    if (!isFinite(time)) time = i;
    if (time >= bestTime) {
      bestTime = time;
      best = row;
    }
  }

  if (!best) return null;
  var client = {};
  headers.forEach(function(header,index){ client[header] = best[index] === undefined ? '' : best[index]; });

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

// ============================================================
// MASTER-SIDE V12 12-MODULE SMOKE AUDIT
// ============================================================
// Non-destructive audit of the selected active client's generated
// package plus its 12 module workspace sheets. It does not deploy,
// write client data, or modify the registry.
// ============================================================

function testV12TwelveModuleSmokeAudit(businessId) {
  var targetBusinessId = String(businessId || '').trim();
  var client = targetBusinessId
    ? getActiveClientByBusinessIdForRedeploy_(targetBusinessId)
    : getActiveClientForV12Test_();

  if (!client) throw new Error(targetBusinessId
    ? 'No active client deployment was found for business ID: ' + targetBusinessId
    : 'No active client deployment was found in the Clients registry.');

  var packageResult = generateClientCodeSafelyV12({
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
    applicationUrl:client.webAppUrl || '',
    featureProfile:(function(){try{return JSON.parse(String(client.Feature_Profile_JSON||'{}'))||{};}catch(e){return {};}})()
  });
  var packageCheck = validateClientDeploymentPackageV12_(packageResult,{
    clientId:client.clientId,
    sheetId:client.sheetId
  });

  var names=["Finance","Sales","Ecommerce","CRM","HR","Logistics","Tax","Agro","Productivity","POS","Attendance","Warehouse"];
  // Master-side audit must not depend on the client runtime MODULES global.
  // Keep this audit self-contained so it can execute in the master project.
  var masterModuleRegistry={
    Finance:{sheet:"Financial_Data"},
    Sales:{sheet:"Sales_Data"},
    Ecommerce:{sheet:"Ecommerce_Data"},
    CRM:{sheet:"CRM_Data"},
    HR:{sheet:"HR_Data"},
    Logistics:{sheet:"Logistics_Data"},
    Tax:{sheet:"Tax_Data"},
    Agro:{sheet:"Agro_Data"},
    Productivity:{sheet:"Productivity_Data"},
    POS:{sheet:"POS_Data"},
    Attendance:{sheet:"Attendance_Data"},
    Warehouse:{sheet:"Warehouse_Data"}
  };
  var ws=SpreadsheetApp.openById(String(client.sheetId || ''));
  var modules=[];
  names.forEach(function(name){
    var def=masterModuleRegistry[name] || {};
    var sh=def.sheet ? ws.getSheetByName(def.sheet) : null;
    var headers=[];
    var recordCount=0;
    if(sh){
      var lastCol=sh.getLastColumn(), lastRow=sh.getLastRow();
      if(lastCol>0) headers=sh.getRange(1,1,1,lastCol).getValues()[0].map(function(x){return String(x||'');});
      recordCount=Math.max(0,lastRow-1);
    }
    modules.push({
      module:name,
      expectedSheet:String(def.sheet || ''),
      sheetExists:!!sh,
      headersPresent:headers.length>0,
      recordCount:recordCount,
      headerCount:headers.length
    });
  });

  var codeFile=(packageResult.files||[]).filter(function(f){return f.name==='Code';})[0];
  var runtime=String(codeFile&&codeFile.source||'');
  var runtimeChecks={
    moduleRegistry:runtime.indexOf('var MODULES=')>=0,
    moduleReadPath:runtime.indexOf('getModuleDataWithSync(name,sid)')>=0,
    sessionGuard:runtime.indexOf('getClientUserFromSession(sid)')>=0,
    entitlementGuard:runtime.indexOf('clientModulesForEntitlementV12_')>=0,
    smokeAuditEmbedded:runtime.indexOf('function getClientV12ModuleSmokeAudit(sid)')>=0
  };

  var failures=modules.filter(function(m){return !m.sheetExists || !m.headersPresent;});
  var result={
    success:!!(packageCheck && packageCheck.success && runtimeChecks.moduleRegistry && runtimeChecks.moduleReadPath && runtimeChecks.sessionGuard && runtimeChecks.entitlementGuard && runtimeChecks.smokeAuditEmbedded && !failures.length),
    test:'V12_12_MODULE_SMOKE_AUDIT',
    liveClientRuntimeExecuted:false,
    note:'This master-side audit is non-destructive. It verifies the generated runtime and the client workspace data structure; the embedded client audit still requires a real client session to execute.',
    packageVersion:packageResult.version,
    clientId:client.clientId,
    businessId:client.businessId,
    email:client.email,
    scriptId:client.scriptId,
    deploymentId:client.deploymentId,
    packageCheck:packageCheck,
    runtimeChecks:runtimeChecks,
    modules:modules
  };
  console.log('V12 12-MODULE SMOKE AUDIT:',JSON.stringify(result));
  return result;
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
  var required=['Code','Modules','Records','ClientShellV12','ClientBusinessCenterV12','ClientFinanceV12','ClientStaffV12','ClientAppsV12','ClientModuleV12','ClientRecordsV12','ClientSettingsV12','ClientProfileV12','ClientNotificationsV12','ClientBillingV12','ClientStylesV12','appsscript'];
  var missing=required.filter(function(n){return names.indexOf(n)<0;});
  if(missing.length)return{success:false,code:'V12_PACKAGE_MISSING_FILES',message:'V12 package is missing: '+missing.join(', ')};
  var shell=pkg.files.filter(function(f){return f.name==='ClientShellV12';})[0].source||'';
  var business=pkg.files.filter(function(f){return f.name==='ClientBusinessCenterV12';})[0].source||'';
  var finance=pkg.files.filter(function(f){return f.name==='ClientFinanceV12';})[0].source||'';
  var staff=pkg.files.filter(function(f){return f.name==='ClientStaffV12';})[0].source||'';
  var apps=pkg.files.filter(function(f){return f.name==='ClientAppsV12';})[0].source||'';
  var moduleWorkspace=pkg.files.filter(function(f){return f.name==='ClientModuleV12';})[0].source||'';
  var settings=pkg.files.filter(function(f){return f.name==='ClientSettingsV12';})[0].source||'';
  var profile=pkg.files.filter(function(f){return f.name==='ClientProfileV12';})[0].source||'';
  var billing=pkg.files.filter(function(f){return f.name==='ClientBillingV12';})[0].source||'';
  var styles=pkg.files.filter(function(f){return f.name==='ClientStylesV12';})[0].source||'';
  var notifications=pkg.files.filter(function(f){return f.name==='ClientNotificationsV12';})[0].source||'';
  var code=pkg.files.filter(function(f){return f.name==='Code';})[0].source||'';
  var modulesServer=pkg.files.filter(function(f){return f.name==='Modules';})[0].source||'';
  var recordsServer=pkg.files.filter(function(f){return f.name==='Records';})[0].source||'';
  var businessCenterEmbedded=code.indexOf('function getBusinessCenterConnectedData(')>=0 || code.indexOf('BUSINESS_CENTER_PRODUCT_HEADERS')>=0;
  if(!businessCenterEmbedded)return{success:false,code:'V12_BUSINESS_CENTER_BACKEND_MISSING',message:'V12 package is missing the embedded BusinessHubGS backend in Code.'};
  var manifestSource=pkg.files.filter(function(f){return f.name==='appsscript';})[0].source||'';
  var forbiddenUi=['id="landing-page"','id="dashboard-page"','id="admin-dashboard"','Admin Login','Client Registry','Deployment Status'];
  var leakedUi=forbiddenUi.filter(function(token){return shell.indexOf(token)>=0;});
  if(leakedUi.length)return{success:false,code:'V12_MASTER_UI_LEAK',message:'V12 client UI contains Master UI markers: '+leakedUi.join(', ')};
  var forbiddenRuntime=['function autoSetupClient','function redeployClientByBusinessId','function requireAdminSession_','function getBizOSMasterSpreadsheet_','function createAndDeployClientScriptSafe'];
  var leakedRuntime=forbiddenRuntime.filter(function(token){return code.indexOf(token)>=0;});
  if(leakedRuntime.length)return{success:false,code:'V12_MASTER_RUNTIME_LEAK',message:'V12 client runtime contains forbidden Master functions: '+leakedRuntime.join(', ')};
  var requiredRuntime=['function doGet(e){','CLIENT_V12_PACKAGE_VERSION_','CLIENT_V12_PACKAGE_RELEASE_','CLIENT_V12_PACKAGE_PROVENANCE_','function v12PersistSession_','function v12ValidateWorkspaceBinding_','function getClientV12WorkspaceHealth','function clientFinanceAccessV12_','function getClientAppsData','function clientSettingsAccessV12_','function getClientSettingsData','function saveClientSettings','function changeClientPassword','function updateClientProfile','function getClientProfileData','function getClientBillingData','function submitClientMaintenanceRequest','function getClientNotificationsData','function markClientNotificationRead','function markAllClientNotificationsRead','function logClientActivityV12_','function createClientNotificationV12_','function getClientFinanceData','function saveClientFinanceTransaction','function deleteClientFinanceTransaction','function authenticateClient','function validateSession','function getUserFromSession','function clientLogout','function getClientDashboardData','function getClientDashboard(','function requireClientSession_','function getBusinessCenterData','function getClientBusinessCenterData','function getClientBusinessCenterConnectedData','function getClientBusinessCenterProductImages','function saveClientBusinessCenterProduct','function saveClientBusinessCenterProductsBulk','function saveClientBusinessCenterCustomer','function lookupClientBusinessCenterProduct','function completeClientBusinessCenterSale','function getClientBusinessCenterCustomerLedger','function recordClientBusinessCenterCustomerPayment','function getClientBusinessCenterSales','function getClientBusinessCenterSaleDetails','function updateClientBusinessCenterProduct','function adjustClientBusinessCenterStock','function getClientBusinessCenterStockMovements','function getClientBusinessCenterFinanceData','function recordClientBusinessCenterExpense'];
  var missingRuntime=requiredRuntime.filter(function(token){return code.indexOf(token)<0;});
  if(missingRuntime.length)return{success:false,code:'V12_RUNTIME_DEPENDENCY_MISSING',message:'V12 client runtime is missing required functions: '+missingRuntime.join(', ')};var securityRuntime=['function clientAssignedModulesV11_','CLIENT_LICENSE_INVALID','function removeTeamMemberLegacy_(bid,email,sid)','function getBusinessStaff(bid,sid)','function inviteStaffMember(bid,email,name,role,invitedBy,assignedModules,sid)','function resendStaffInvitation(bid,email,sid)','function assignModuleToStaff(email,bid,moduleName,assignedBy,sid)','function removeModuleFromStaff(email,bid,moduleName,sid)','String(u.businessId)!==String(bid)','function updateUserProfile(bid,name,phone,sid)','function changeUserPassword(email,currentPassword,newPassword,sid)','function createClientSessionV11_','function getOrCreateUserSheet_','function getOrCreateBusinessSheet_'];var missingSecurity=securityRuntime.filter(function(token){return code.indexOf(token)<0;});if(missingSecurity.length)return{success:false,code:'V12_SECURITY_DEPENDENCY_MISSING',message:'V12 client runtime is missing hardened security markers: '+missingSecurity.join(', ')};if(code.indexOf('function createSession(')>=0)return{success:false,code:'V12_PUBLIC_SESSION_FACTORY_LEAK',message:'V12 client runtime still exposes the public session factory.'};var forbiddenPublic=['function getOrCreateUserSheet(','function getOrCreateBusinessSheet('];var leakedPublic=forbiddenPublic.filter(function(token){return code.indexOf(token)>=0;});if(leakedPublic.length)return{success:false,code:'V12_INTERNAL_HELPER_EXPOSED',message:'V12 client runtime still exposes internal helper functions: '+leakedPublic.join(', ')};if(code.indexOf('var MODULES=')<0)return{success:false,code:'V12_MODULE_REGISTRY_MISSING',message:'V12 client runtime is missing the module registry.'};var requiredEntitlement=['function clientModulesForEntitlementV12_','function clientLicenseTierV12_','accessibleModules:clientModulesForEntitlementV12_'];var missingEntitlement=requiredEntitlement.filter(function(token){return code.indexOf(token)<0;});if(missingEntitlement.length)return{success:false,code:'V12_ENTITLEMENT_BRIDGE_MISSING',message:'V12 entitlement enforcement is incomplete: '+missingEntitlement.join(', ')};var requiredDashboard=['function getClientDashboard(sid,forceRefresh){','function getClientDashboardData(sid,forceRefresh){return getClientDashboard(sid,forceRefresh);}','function getClientUserFromSession(sid)','function getClientPackageInfo(sid)','function getClientV12ModuleSmokeAudit(sid)','businessSnapshot:businessSnapshot','modulesWithData','weekly:weekly','currency:clientCurrency','recentActivity:activity.slice','kpis:{revenue:{value:','profitMargin:{value:'];var missingDashboard=requiredDashboard.filter(function(token){return code.indexOf(token)<0;});if(missingDashboard.length)return{success:false,code:'V12_DASHBOARD_DEPENDENCY_MISSING',message:'V12 dashboard runtime is missing required parity markers: '+missingDashboard.join(', ')};
  var requiredShell=['ClientStylesV12','ClientStaffV12','ClientAppsV12','ClientSettingsV12','ClientProfileV12','ClientNotificationsV12','ClientBillingV12','ClientBusinessCenterV12','authenticateClient','validateSession','getUserFromSession','clientLogout','getClientDashboardData','bizos-kpi-grid','trendChart','weeklyTrendChart','operationalSnapshot','recentActivityView','OPERATIONAL SNAPSHOT','Workspace activity','breakdownChart','healthView'];
  var requiredModule=['client-module-v12-root','clientModuleV12Open','clientModuleV12RenderTable','clientModuleV12Refresh','clientModuleV12New','clientModuleV12Edit','clientModuleV12Delete','clientModuleV12SetMode','clientModuleV12Export','clientModuleV12Import','clientModuleV12SaveSheet','client-module-insight','Quick Insight','Spreadsheet View','Export','Import'];
  var requiredServerDeps=['function getModuleData(','function getModuleDataWithSync('];
  var missingServerDeps=requiredServerDeps.filter(function(token){return modulesServer.indexOf(token)<0;});
  if(missingServerDeps.length)return{success:false,code:'V12_MODULE_SERVER_DEPENDENCY_MISSING',message:'V12 module server dependency is missing: '+missingServerDeps.join(', ')};
  if(recordsServer.indexOf('function getRecordsDirectory(')<0||recordsServer.indexOf('function getRecordRows(')<0)return{success:false,code:'V12_RECORDS_SERVER_DEPENDENCY_MISSING',message:'V12 Records server dependency is incomplete.'};
  if(code.indexOf('function getBusinessCenterData(')<0||code.indexOf('function completeBusinessCenterSale(')<0)return{success:false,code:'V12_BUSINESS_CENTER_SERVER_DEPENDENCY_MISSING',message:'V12 embedded Business Center server dependency is incomplete.'};
  var missingModule=requiredModule.filter(function(token){return moduleWorkspace.indexOf(token)<0;});
  if(missingModule.length)return{success:false,code:'V12_MODULE_DEPENDENCY_MISSING',message:'V12 Module workspace is missing required markers: '+missingModule.join(', ')};
  var requiredStaff=['staff-v12-root','getBusinessStaff','inviteStaffMember','resendStaffInvitation','removeTeamMember','assignModuleToStaff','getAvailableModulesForStaff'];
  var missingStaff=requiredStaff.filter(function(token){return staff.indexOf(token)<0;});
  if(missingStaff.length)return{success:false,code:'V12_STAFF_DEPENDENCY_MISSING',message:'V12 Staff workspace is missing required markers: '+missingStaff.join(', ')};
  var requiredApps=['apps-v12-root','getClientAppsData'];var missingApps=requiredApps.filter(function(token){return apps.indexOf(token)<0;});if(missingApps.length)return{success:false,code:'V12_APPS_DEPENDENCY_MISSING',message:'V12 Apps workspace is missing required markers: '+missingApps.join(', ')};var requiredSettings=['settings-v12-root','getClientSettingsData','saveClientSettings','changeClientPassword'];var missingSettings=requiredSettings.filter(function(token){return settings.indexOf(token)<0;});if(missingSettings.length)return{success:false,code:'V12_SETTINGS_DEPENDENCY_MISSING',message:'V12 Settings workspace is missing required markers: '+missingSettings.join(', ')};var requiredProfile=['client-profile-v12','getClientProfileData','clientProfileInit','clientProfileSave','clientProfileChangePassword'];var missingProfile=requiredProfile.filter(function(token){return profile.indexOf(token)<0;});if(missingProfile.length)return{success:false,code:'V12_PROFILE_DEPENDENCY_MISSING',message:'V12 Profile workspace is missing required markers: '+missingProfile.join(', ')};var requiredBilling=['client-billing-v12','getClientBillingData','getClientQuotePaymentDetails','verifyClientQuotePayment','submitClientMaintenanceRequest','submitClientFeatureRequest','getClientFeatureProfileV12','clientBillingInit','clientBillingSubmit','clientBillingFeatureSubmit'];var requiredNotifications=['client-notifications-v12','getClientNotificationsData','markClientNotificationRead','markAllClientNotificationsRead','clientNotificationsInit','clientNotificationsTab'];var missingNotifications=requiredNotifications.filter(function(token){return notifications.indexOf(token)<0;});if(missingNotifications.length)return{success:false,code:'V12_NOTIFICATIONS_DEPENDENCY_MISSING',message:'V12 Notifications workspace is missing required markers: '+missingNotifications.join(', ')};var missingBilling=requiredBilling.filter(function(token){return billing.indexOf(token)<0;});if(missingBilling.length)return{success:false,code:'V12_BILLING_DEPENDENCY_MISSING',message:'V12 Billing workspace is missing required markers: '+missingBilling.join(', ')};var requiredBusiness=['business-center-page','bc-pos-page','openBCScanner','BarcodeDetector','getUserMedia','saveBusinessCenterProduct','saveBusinessCenterProductsBulk','lookupBusinessCenterProduct','completeBusinessCenterSale','getBusinessCenterCustomerLedger','recordBusinessCenterCustomerPayment','getBusinessCenterSaleDetails','getBusinessCenterFinanceData','recordBusinessCenterExpense','printBCReceipt'];
  var missingBusiness=requiredBusiness.filter(function(token){return business.indexOf(token)<0;});
  if(missingBusiness.length)return{success:false,code:'V12_BUSINESS_CENTER_DEPENDENCY_MISSING',message:'V12 Business Center UI is missing required markers: '+missingBusiness.join(', ')};
  var missingShell=requiredShell.filter(function(token){return shell.indexOf(token)<0;});
  if(missingShell.length)return{success:false,code:'V12_SHELL_DEPENDENCY_MISSING',message:'V12 client shell is missing required UI/runtime markers: '+missingShell.join(', ')};
  var requiredStyles=['bizos-kpi-grid','.bar-chart','.health-score','.mini-table'];
  var missingStyles=requiredStyles.filter(function(token){return styles.indexOf(token)<0;});
  if(missingStyles.length)return{success:false,code:'V12_STYLE_DEPENDENCY_MISSING',message:'V12 client styles are missing required dashboard styles: '+missingStyles.join(', ')};
  try{var manifest=JSON.parse(manifestSource);if(!manifest||manifest.runtimeVersion!=='V8'||!manifest.webapp)return{success:false,code:'V12_MANIFEST_INVALID',message:'V12 manifest is missing required V8/webapp configuration.'};}catch(e){return{success:false,code:'V12_MANIFEST_INVALID',message:'V12 manifest is not valid JSON: '+e.message};}
  if(expected&&expected.clientId&&code.indexOf(String(expected.clientId))<0)return{success:false,code:'V12_CLIENT_ID_MISMATCH',message:'V12 runtime does not contain the expected client ID.'};
  if(expected&&expected.sheetId&&code.indexOf(String(expected.sheetId))<0)return{success:false,code:'V12_WORKSPACE_ID_MISMATCH',message:'V12 runtime does not contain the expected workspace ID.'};
  return{success:true,code:'V12_PACKAGE_VALID',version:CLIENT_V12_PACKAGE_VERSION_,release:CLIENT_V12_PACKAGE_RELEASE_,provenance:CLIENT_V12_PACKAGE_PROVENANCE_,fileCount:pkg.files.length,files:names,checks:{runtimeDependencies:true,shellDependencies:true,businessCenterDependencies:true,staffDependencies:true,appsDependencies:true,settingsDependencies:true,notificationsDependencies:true,styleDependencies:true,manifest:true,masterUiLeak:false,masterRuntimeLeak:false}};
}

// ============================================================
// V12 DEPLOYED RUNTIME AUDIT
// ============================================================
// Read-only audit of the exact Apps Script version currently
// attached to the client's registered deployment.
// It does not publish, deploy, or modify client/registry data.
// ============================================================

function auditV12ClientDeploymentRuntime(businessId) {
  var targetBusinessId = String(businessId || '').trim();
  if (!targetBusinessId) {
    return {
      success:false,
      code:'BUSINESS_ID_REQUIRED',
      message:'Business ID is required.'
    };
  }

  var client = getActiveClientByBusinessIdForRedeploy_(targetBusinessId);
  if (!client) {
    return {
      success:false,
      code:'CLIENT_NOT_FOUND',
      message:'No active client deployment was found for business ID: ' + targetBusinessId
    };
  }

  var required = ['clientId','businessId','scriptId','deploymentId'];
  for (var i = 0; i < required.length; i++) {
    if (!String(client[required[i]] || '').trim()) {
      return {
        success:false,
        code:'CLIENT_DEPLOYMENT_METADATA_INCOMPLETE',
        message:'The client registry is missing required deployment metadata: ' + required[i] + '.',
        clientId:client.clientId
      };
    }
  }

  try {
    var headers = {
      Authorization:'Bearer ' + ScriptApp.getOAuthToken(),
      'Content-Type':'application/json'
    };
    var base = 'https://script.googleapis.com/v1/projects/' + encodeURIComponent(client.scriptId);

    // 1. Resolve the deployment to the exact version Google is serving.
    var deploymentResponse = UrlFetchApp.fetch(
      base + '/deployments/' + encodeURIComponent(client.deploymentId),
      {method:'get',headers:headers,muteHttpExceptions:true}
    );

    if (deploymentResponse.getResponseCode() < 200 || deploymentResponse.getResponseCode() >= 300) {
      return {
        success:false,
        code:'DEPLOYMENT_LOOKUP_FAILED',
        message:'Could not read the registered deployment: ' + deploymentResponse.getContentText(),
        clientId:client.clientId,
        scriptId:client.scriptId,
        deploymentId:client.deploymentId
      };
    }

    var deployment = JSON.parse(deploymentResponse.getContentText() || '{}');
    var config = deployment.deploymentConfig || {};
    var deployedVersion = String(config.versionNumber || '');

    // 2. Fetch the exact immutable version attached to the deployment.
    var versionContent = null;
    if (deployedVersion) {
      var versionResponse = UrlFetchApp.fetch(
        base + '/content?versionNumber=' + encodeURIComponent(deployedVersion),
        {method:'get',headers:headers,muteHttpExceptions:true}
      );

      if (versionResponse.getResponseCode() < 200 || versionResponse.getResponseCode() >= 300) {
        return {
          success:false,
          code:'DEPLOYED_VERSION_READ_FAILED',
          message:'The deployment points to version ' + deployedVersion + ', but that exact version could not be read: ' + versionResponse.getContentText(),
          clientId:client.clientId,
          scriptId:client.scriptId,
          deploymentId:client.deploymentId,
          deployedVersion:deployedVersion
        };
      }

      versionContent = JSON.parse(versionResponse.getContentText() || '{}');
    }

    // 3. Also read project HEAD so we can distinguish deployed-version drift
    // from a package-generation problem.
    var headResponse = UrlFetchApp.fetch(
      base + '/content',
      {method:'get',headers:headers,muteHttpExceptions:true}
    );

    if (headResponse.getResponseCode() < 200 || headResponse.getResponseCode() >= 300) {
      return {
        success:false,
        code:'HEAD_READ_FAILED',
        message:'The client project HEAD could not be read: ' + headResponse.getContentText(),
        clientId:client.clientId,
        scriptId:client.scriptId,
        deploymentId:client.deploymentId,
        deployedVersion:deployedVersion
      };
    }

    var headContent = JSON.parse(headResponse.getContentText() || '{}');

    function getCode_(project) {
      var files = project && Array.isArray(project.files) ? project.files : [];
      var codeFile = files.filter(function(file){
        return String(file.name || '') === 'Code';
      })[0];
      return {
        file:codeFile || null,
        source:String(codeFile && codeFile.source || '')
      };
    }

    function inspectRuntime_(source) {
      var s=String(source || '');
      return {
        codePresent:s.length>0,
        v12Marker:s.indexOf('CLIENT_V12_PACKAGE_VERSION_')>=0,
        moduleRegistry:s.indexOf('var MODULES=')>=0,
        clientId:s.indexOf(String(client.clientId))>=0,
        authenticateClient:s.indexOf('function authenticateClient(email,password)')>=0,
        createSession:s.indexOf('function createClientSessionV11_(')>=0,
        findClientUser:s.indexOf('function findClientUser_(')>=0,
        dashboardData:s.indexOf('function getClientDashboardData(sid,forceRefresh)')>=0,
        moduleSmokeAudit:s.indexOf('function getClientV12ModuleSmokeAudit(sid)')>=0,
        moduleLookupInLoginPath:s.indexOf('Object.keys(MODULES)')>=0,
        durableSession:s.indexOf('function v12PersistSession_')>=0,
        workspaceHealth:s.indexOf('function getClientV12WorkspaceHealth')>=0
      };
    }

    function digest_(source) {
      var bytes=Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        String(source || ''),
        Utilities.Charset.UTF_8
      );
      return bytes.map(function(b){
        var n=b<0 ? b+256 : b;
        return ('0'+n.toString(16)).slice(-2);
      }).join('');
    }

    var deployedCode=getCode_(versionContent);
    var headCode=getCode_(headContent);
    var deployedChecks=inspectRuntime_(deployedCode.source);
    var headChecks=inspectRuntime_(headCode.source);

    // Read-only live workspace/database validation. This is intentionally
    // independent of the generated client runtime so the audit can distinguish
    // a package problem from a malformed client spreadsheet.
    var workspaceHealth={success:false,clientInfo:false,modules:[],notificationsSheet:false,activitySheet:false};
    try{
      var ws=SpreadsheetApp.openById(String(client.sheetId||''));
      var info=ws.getSheetByName('Client_Info'),meta={};
      if(info){
        var iv=info.getDataRange().getValues();
        for(var mi=1;mi<iv.length;mi++)if(iv[mi][0])meta[String(iv[mi][0])]=iv[mi][1];
      }
      workspaceHealth.clientInfo=!!info &&
        String(meta.Client_ID||'')===String(client.clientId||'') &&
        String(meta.Business_ID||client.businessId||'')===String(client.businessId||'') &&
        String(meta.Status||'active').toLowerCase()==='active';
      var expectedModules=['Finance','Sales','Ecommerce','CRM','HR','Logistics','Tax','Agro','Productivity','POS','Attendance','Warehouse'];
      expectedModules.forEach(function(name){
        var sheetName=name+'_Data',sh=ws.getSheetByName(sheetName),headers=[];
        if(sh&&sh.getLastColumn()>0)headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
        workspaceHealth.modules.push({module:name,sheet:sheetName,exists:!!sh,headersPresent:headers.length>0,recordCount:sh&&sh.getLastRow()>1?sh.getLastRow()-1:0});
      });
      workspaceHealth.notificationsSheet=!!ws.getSheetByName('Client_Notifications');
      workspaceHealth.activitySheet=!!ws.getSheetByName('Client_Activity');
      workspaceHealth.success=workspaceHealth.clientInfo &&
        workspaceHealth.modules.every(function(m){return m.exists&&m.headersPresent;});
    }catch(workspaceError){
      workspaceHealth.error=String(workspaceError&&workspaceError.message||workspaceError);
    }

    var bindingOk =
      String(deployment.deploymentId || '') === String(client.deploymentId) &&
      String(config.scriptId || client.scriptId) === String(client.scriptId);

    var result={
      success:!!(bindingOk && deployedVersion && deployedChecks.codePresent),
      audit:'V12_DEPLOYED_RUNTIME_AUDIT',
      clientId:client.clientId,
      businessId:client.businessId,
      email:client.email,
      scriptId:client.scriptId,
      deploymentId:client.deploymentId,
      deploymentVersion:deployedVersion,
      deploymentDescription:String(config.description || ''),
      deploymentBindingOk:bindingOk,
      deployedRuntime:deployedChecks,
      headRuntime:headChecks,
      workspaceHealth:workspaceHealth,
      deployedCodeLength:deployedCode.source.length,
      headCodeLength:headCode.source.length,
      deployedCodeSha256:digest_(deployedCode.source),
      headCodeSha256:digest_(headCode.source),
      deployedMatchesHead:
        deployedCode.source.length === headCode.source.length &&
        digest_(deployedCode.source) === digest_(headCode.source),
      diagnosis:''
    };

    if (!bindingOk) {
      result.diagnosis='DEPLOYMENT_BINDING_MISMATCH: the registered deployment does not belong to the registered client script.';
    } else if (!deployedVersion) {
      result.success=false;
      result.diagnosis='DEPLOYMENT_VERSION_MISSING: the registered deployment has no version number.';
    } else if (!deployedChecks.codePresent) {
      result.success=false;
      result.diagnosis='DEPLOYED_CODE_MISSING: the exact deployed version has no readable Code file.';
    } else if (!deployedChecks.moduleRegistry) {
      result.success=false;
      result.diagnosis='DEPLOYED_RUNTIME_MISSING_MODULES: the exact version served by the deployment does not contain var MODULES=. This is the direct cause of the MODULES is not defined login failure.';
    } else if (!deployedChecks.authenticateClient || !deployedChecks.createSession || !deployedChecks.findClientUser) {
      result.success=false;
      result.diagnosis='DEPLOYED_LOGIN_PATH_INCOMPLETE: the deployed version is missing one or more functions in the authenticateClient -> createClientSessionV11_ -> findClientUser_ path.';
    } else if (!workspaceHealth.success) {
      result.success=false;
      result.diagnosis='CLIENT_WORKSPACE_VALIDATION_FAILED: Client_Info or one or more required module sheets/headers are missing or do not match the registered client.';
    } else if (!deployedChecks.durableSession || !deployedChecks.workspaceHealth || !headChecks.durableSession || !headChecks.workspaceHealth) {
      result.success=false;
      result.diagnosis='V12_SESSION_VALIDATION_INCOMPLETE: the deployed or HEAD runtime is missing the durable-session/workspace-validation bridge.';
    } else if (!headChecks.moduleRegistry) {
      result.success=false;
      result.diagnosis='GENERATOR_RUNTIME_MISSING_MODULES: project HEAD itself does not contain var MODULES=. The package generator/runtime assembly must be fixed before redeploying.';
    } else if (!result.deployedMatchesHead) {
      result.diagnosis='DEPLOYED_VERSION_DIFFERS_FROM_HEAD: the deployment is serving an older/different immutable version than project HEAD. Compare deploymentVersion with the latest version before changing client code.';
    } else {
      result.diagnosis='DEPLOYED_RUNTIME_CONTAINS_MODULES: source-level deployment inspection does not reproduce the reported MODULES error. The next step is runtime execution tracing/logging on the client login path, not another package-generation change.';
    }

    return result;
  } catch (error) {
    return {
      success:false,
      code:'DEPLOYED_RUNTIME_AUDIT_EXCEPTION',
      message:error && error.message ? error.message : 'V12 deployed runtime audit failed.',
      clientId:client.clientId,
      scriptId:client.scriptId,
      deploymentId:client.deploymentId
    };
  }
}

function auditV12ReleaseCompleteness(businessId) {
  var target=String(businessId||'').trim();
  var client=target ? getActiveClientByBusinessIdForRedeploy_(target) : getActiveClientForV12Test_();
  if(!client) throw new Error(target ? 'No active client deployment was found for business ID: '+target : 'No active client deployment was found.');

  var pkg=generateClientCodeSafelyV12({
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
    applicationUrl:client.webAppUrl || '',
    featureProfile:(function(){try{return JSON.parse(String(client.Feature_Profile_JSON||'{}'))||{};}catch(e){return {};}})()
  });
  var packageCheck=validateClientDeploymentPackageV12_(pkg,{clientId:client.clientId,sheetId:client.sheetId});
  var code=((pkg.files||[]).filter(function(f){return f.name==='Code';})[0]||{}).source||'';
  var staff=((pkg.files||[]).filter(function(f){return f.name==='ClientStaffV12';})[0]||{}).source||'';
  var dashboardSource=String(getMasterSourceForV12_('Dashboard')||'');

  var staffMarkers=[
    'function clientTeamRowsV11_',
    'function ensureClientTeamSheetV11_',
    'function getBusinessStaff(bid,sid)',
    'function inviteStaffMember(bid,email,name,role,invitedBy,assignedModules,sid)',
    'function resendStaffInvitation(bid,email,sid)',
    'function assignModuleToStaff(email,bid,moduleName,assignedBy,sid)',
    'function removeModuleFromStaff(bid,email,sid)',
    'function clientAssignedModulesV11_',
    'Assigned_Modules',
    'staff-v12-root'
  ];
  var staffChecks={};
  staffMarkers.forEach(function(token){staffChecks[token]=code.indexOf(token)>=0||staff.indexOf(token)>=0;});

  var dashboardMarkers=[
    'function getDashboardKPIs',
    'function getEnhancedDashboardData',
    'function getRecentTransactions',
    'function getChartDataOptimized',
    'function getDetailedFinancialMetrics',
    'function getDailyTrends',
    'function getWeeklyTrends',
    'businessSnapshot',
    'health',
    'charts',
    'recentActivity'
  ];
  var dashboardComparison={masterMarkers:{},clientMarkers:{},matched:0,total:dashboardMarkers.length};
  dashboardMarkers.forEach(function(token){
    var masterHas=dashboardSource.indexOf(token)>=0;
    var clientHas=code.indexOf(token)>=0;
    dashboardComparison.masterMarkers[token]=masterHas;
    dashboardComparison.clientMarkers[token]=clientHas;
    if(masterHas===clientHas) dashboardComparison.matched++;
  });

  var names=['Finance','Sales','Ecommerce','CRM','HR','Logistics','Tax','Agro','Productivity','POS','Attendance','Warehouse'];
  var ws=SpreadsheetApp.openById(String(client.sheetId||''));
  var modules=names.map(function(name){
    var sh=ws.getSheetByName(name+'_Data');
    var headers=sh&&sh.getLastColumn()>0?sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]:[];
    return {module:name,sheet:name+'_Data',sheetExists:!!sh,headersPresent:headers.length>0,recordCount:sh&&sh.getLastRow()>1?sh.getLastRow()-1:0};
  });

  var staffComplete=staffMarkers.every(function(token){return staffChecks[token];});
  var dashboardParity=dashboardComparison.matched===dashboardComparison.total;
  var modulesComplete=modules.every(function(m){return m.sheetExists&&m.headersPresent;});

  return {
    success:!!(packageCheck&&packageCheck.success&&staffComplete&&modulesComplete),
    audit:'V12_RELEASE_COMPLETENESS',
    liveClientRuntimeExecuted:false,
    note:'This is a non-destructive source/workspace audit. It closes the static Staff, Master-dashboard comparison, and 12-module workspace checks. A real client session is still required for live Staff actions and live module reads/writes.',
    package:{version:pkg.version,release:pkg.release,provenance:pkg.provenance,check:packageCheck},
    staffAudit:{complete:staffComplete,markers:staffChecks},
    dashboardComparison:{parity:dashboardParity,matched:dashboardComparison.matched,total:dashboardComparison.total,markers:dashboardComparison},
    modulesAudit:{complete:modulesComplete,count:modules.length,modules:modules}
  };
}
