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

  var runtime = buildClientRuntimeV11_(config);
  runtime = runtime.replace(
    /function doGet\(e\)\{[\s\S]*?\},'function include/,
    'function doGet(e){var t=HtmlService.createTemplateFromFile("ClientShellV12");return t.evaluate().setTitle(CLIENT_CONFIG.clientName+" - BizOS").addMetaTag("viewport","width=device-width,initial-scale=1").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}\nfunction include'
  );

  var files = [
    {name:'Code',type:'SERVER_JS',source:runtime},
    {name:'ClientShellV12',type:'HTML',source:buildClientShellV12_(config)},
    {name:'ClientStylesV12',type:'HTML',source:buildClientStylesV12_()},
    {name:'appsscript',type:'JSON',source:JSON.stringify(buildClientManifestV12_())}
  ];

  return {
    success:true,version:'12.0',config:config,files:files,
    source:{clientUi:'V12',copiedMasterUi:[],reusedBackendRuntime:'V11-safe-runtime'}
  };
}

function buildClientShellV12_(config) {
  var source = String(getMasterSourceForV12_('ClientShellV12') || '');
  if(!source) throw new Error('ClientShellV12 source is missing.');
  return source.replace(/__CLIENT_CONFIG_PLACEHOLDER__/g, JSON.stringify({
    isClient:true,id:config.clientId,name:config.clientName,primaryColor:config.primaryColor,logoUrl:config.logoUrl
  })).replace(/<head>/i,'<head><script>window.__CLIENT=__CLIENT_CONFIG_PLACEHOLDER__;</script>');
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
      'https://www.googleapis.com/auth/script.send_mail',
      'https://www.googleapis.com/auth/script.projects.readonly'
    ]
  };
}

function validateClientDeploymentPackageV12_(pkg, expected) {
  if(!pkg||!Array.isArray(pkg.files)||!pkg.files.length)return{success:false,code:'PACKAGE_EMPTY',message:'Generated V12 client package is empty.'};
  var names=pkg.files.map(function(f){return String(f.name||'');});
  var required=['Code','ClientShellV12','ClientStylesV12','appsscript'];
  var missing=required.filter(function(n){return names.indexOf(n)<0;});
  if(missing.length)return{success:false,code:'V12_PACKAGE_MISSING_FILES',message:'V12 package is missing: '+missing.join(', ')};
  var shell=pkg.files.filter(function(f){return f.name==='ClientShellV12';})[0].source||'';
  var code=pkg.files.filter(function(f){return f.name==='Code';})[0].source||'';
  var forbiddenUi=['id="landing-page"','id="dashboard-page"','id="admin-dashboard"','Admin Login','Client Registry','Deployment Status'];
  var leakedUi=forbiddenUi.filter(function(token){return shell.indexOf(token)>=0;});
  if(leakedUi.length)return{success:false,code:'V12_MASTER_UI_LEAK',message:'V12 client UI contains Master UI markers: '+leakedUi.join(', ')};
  var forbiddenRuntime=['function autoSetupClient','function redeployClientByBusinessId','function requireAdminSession_','function getBizOSMasterSpreadsheet_','function createAndDeployClientScriptSafe'];
  var leakedRuntime=forbiddenRuntime.filter(function(token){return code.indexOf(token)>=0;});
  if(leakedRuntime.length)return{success:false,code:'V12_MASTER_RUNTIME_LEAK',message:'V12 client runtime contains forbidden Master functions: '+leakedRuntime.join(', ')};
  if(expected&&expected.clientId&&code.indexOf(String(expected.clientId))<0)return{success:false,code:'V12_CLIENT_ID_MISMATCH',message:'V12 runtime does not contain the expected client ID.'};
  if(expected&&expected.sheetId&&code.indexOf(String(expected.sheetId))<0)return{success:false,code:'V12_WORKSPACE_ID_MISMATCH',message:'V12 runtime does not contain the expected workspace ID.'};
  return{success:true,code:'V12_PACKAGE_VALID',version:'12.0',fileCount:pkg.files.length,files:names};
}