// ============================================================
// ClientDeploymentV10.js - CLIENT PACKAGE MANIFEST HARDENING
// ============================================================
// Keeps V9 behavior unchanged and only ensures the generated client
// manifest explicitly permits invitation emails through MailApp.
// ============================================================

function buildClientUrlConfigV10_(settings){
  settings=settings||{};
  var applicationUrl=String(settings.applicationUrl||settings.webAppUrl||'').trim();
  var landingUrl=String(settings.landingUrl||'').trim();
  var customDomain=String(settings.customDomain||'').trim();
  return [
    '// ============================================================',
    '// CLIENT URL / LANDING CONFIGURATION - V10',
    '// ============================================================',
    'CLIENT_CONFIG.applicationUrl='+JSON.stringify(applicationUrl)+';',
    'CLIENT_CONFIG.landingUrl='+JSON.stringify(landingUrl)+';',
    'CLIENT_CONFIG.customDomain='+JSON.stringify(customDomain)+';',
    'function getClientPublicUrlsV10(){',
    '  var appUrl=String(CLIENT_CONFIG.applicationUrl||\'\').trim();',
    '  try{if(!appUrl)appUrl=String(ScriptApp.getService().getUrl()||\'\').trim();}catch(ignore){}',
    '  return {applicationUrl:appUrl,landingUrl:String(CLIENT_CONFIG.landingUrl||\'\').trim(),customDomain:String(CLIENT_CONFIG.customDomain||\'\').trim()};',
    '}'
  ].join(String.fromCharCode(10)) + String.fromCharCode(10);
}

function buildClientLicenseGuardV10_() {
  return `
  
// ============================================================
// CLIENT LICENSE / DEPLOYMENT BINDING - V10
// ============================================================
// A copied source package does not inherit the original Apps Script
// project identity. Every authenticated/session-protected operation
// therefore re-validates the actual runtime project + deployment
// against the master Clients registry.

function getClientRuntimeIdentityV10_() {
  var scriptId = '';
  var deploymentId = '';
  var webAppUrl = '';

  try {
    scriptId = String(ScriptApp.getScriptId() || '').trim();
  } catch (ignoreScriptId) {}

  try {
    webAppUrl = String(ScriptApp.getService().getUrl() || '').trim();
  } catch (ignoreServiceUrl) {}

  var match = webAppUrl.match(/\\/s\\/([^\\/]+)/);
  if (match && match[1]) deploymentId = String(match[1]).trim();

  return {
    scriptId: scriptId,
    deploymentId: deploymentId,
    webAppUrl: webAppUrl
  };
}

function validateClientLicenseV10_(session) {
  try {
    session = session || {};

    if (!session.clientId || !session.businessId || !session.sheetId) {
      return {
        success: false,
        code: 'LICENSE_SESSION_INVALID',
        message: 'Client license session is incomplete.'
      };
    }

    var runtime = getClientRuntimeIdentityV10_();
    if (!runtime.scriptId || !runtime.deploymentId) {
      return {
        success: false,
        code: 'LICENSE_RUNTIME_IDENTITY_UNAVAILABLE',
        message: 'This BizOS deployment could not verify its runtime identity.'
      };
    }

    var masterUrl = '';
    try {
      masterUrl = String((typeof CLIENT_CONFIG !== 'undefined' && CLIENT_CONFIG.masterApiUrl) || '').trim();
    } catch (ignoreConfig) {}

    if (!masterUrl) {
      return {
        success: false,
        code: 'LICENSE_MASTER_UNAVAILABLE',
        message: 'BizOS licensing service is not configured.'
      };
    }

    var response = UrlFetchApp.fetch(masterUrl, {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({
        action: 'validateClientLicense',
        payload: {
          clientId: String(session.clientId),
          businessId: String(session.businessId),
          sheetId: String(session.sheetId),
          scriptId: runtime.scriptId,
          deploymentId: runtime.deploymentId
        }
      })
    });

    var httpCode = response.getResponseCode();
    if (httpCode < 200 || httpCode >= 300) {
      return {
        success: false,
        code: 'LICENSE_SERVICE',
        message: 'BizOS licensing service is unavailable.'
      };
    }

    var result = JSON.parse(response.getContentText() || '{}');
    if (!result || !result.success) {
      return result || {
        success: false,
        code: 'LICENSE_REJECTED',
        message: 'This BizOS deployment is not authorized.'
      };
    }

    return result;
  } catch (error) {
    console.error('Client license validation error:', error);
    return {
      success: false,
      code: 'LICENSE_ERROR',
      message: error && error.message ? error.message : 'License validation failed.'
    };
  }
}

// Replace the generated local-session validator with a validator that
// also checks the master license binding. The local session remains the
// fast identity check; the master remains authoritative for entitlement.
function validateClientSession(sessionId) {
  if (!sessionId) return null;

  var cache = CacheService.getScriptCache();

  try {
    var raw = cache.get(String(sessionId));
    if (!raw) return null;

    var session = JSON.parse(raw);
    var now = Date.now();

    if (!session.created ||
        !session.clientId ||
        !session.sheetId ||
        !session.businessId ||
        (session.expiresAt && now >= Number(session.expiresAt)) ||
        now - Number(session.created) > 21600000) {
      cache.remove(String(sessionId));
      return null;
    }

    if (typeof CLIENT_CONFIG === 'undefined' ||
        String(session.clientId) !== String(CLIENT_CONFIG.clientId) ||
        String(session.sheetId) !== String(CLIENT_CONFIG.sheetId) ||
        String(session.businessId) !== String(CLIENT_CONFIG.businessId)) {
      cache.remove(String(sessionId));
      return null;
    }

    var license = validateClientLicenseV10_(session);
    if (!license || !license.success) {
      cache.remove(String(sessionId));
      return null;
    }

    return session;
  } catch (error) {
    try { cache.remove(String(sessionId)); } catch (ignore) {}
    return null;
  }
}

// The license guard is intentionally also exposed for internal smoke tests.
function checkClientLicenseV10() {
  try {
    var clientId = typeof CLIENT_CONFIG !== 'undefined' ? String(CLIENT_CONFIG.clientId || '') : '';
    var businessId = typeof CLIENT_CONFIG !== 'undefined' ? String(CLIENT_CONFIG.businessId || '') : '';
    var sheetId = typeof CLIENT_CONFIG !== 'undefined' ? String(CLIENT_CONFIG.sheetId || '') : '';

    return validateClientLicenseV10_({
      clientId: clientId,
      businessId: businessId,
      sheetId: sheetId
    });
  } catch (error) {
    return {
      success: false,
      code: 'LICENSE_ERROR',
      message: error && error.message ? error.message : 'License validation failed.'
    };
  }
}
`;
}

function buildClientNavigationV10_(){
  return `
function getClientApplicationUrlV10_(){
  try{
    var urls=getClientPublicUrlsV10();
    if(urls&&urls.applicationUrl)return String(urls.applicationUrl).replace(/\\/$/,'');
  }catch(ignore){}
  try{
    return String(ScriptApp.getService().getUrl()||'').replace(/\\/$/,'');
  }catch(ignoreUrl){}
  return '';
}

function clientNavigateV10_(page,params){
  var base=getClientApplicationUrlV10_();
  if(!base)throw new Error('Client application URL is unavailable.');
  var query=[];
  params=params||{};
  Object.keys(params).forEach(function(k){
    if(params[k]!==undefined&&params[k]!==null&&String(params[k])!==''){
      query.push(encodeURIComponent(k)+'='+encodeURIComponent(String(params[k])));
    }
  });
  var target=base+'?page='+encodeURIComponent(String(page||'login'))+(query.length?'&'+query.join('&'):'');
  return target;
}

function doGet(e){
  var p=String(e&&e.parameter&&e.parameter.page||'login'),t;
  if(p==='dashboard'){
    t=HtmlService.createTemplateFromFile('client-dashboard');
    t.sessionId=e&&e.parameter?e.parameter.sessionId||'':'';
  }else{
    t=HtmlService.createTemplateFromFile('client-landing');
  }
  t.clientId=CLIENT_CONFIG.clientId;
  t.clientName=CLIENT_CONFIG.clientName;
  t.primaryColor=CLIENT_CONFIG.primaryColor;
  t.logoUrl=CLIENT_CONFIG.logoUrl;
  t.applicationUrl=getClientApplicationUrlV10_();
  return t.evaluate()
    .setTitle(CLIENT_CONFIG.clientName+' - BizOS')
    .addMetaTag('viewport','width=device-width,initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function renderClientPageV10(page,sessionId){
  try{
    page=String(page||'login').toLowerCase();
    var t;
    if(page==='dashboard'){
      if(!sessionId)return{success:false,message:'Session is required.',code:'SESSION_REQUIRED'};
      var session=validateClientSession(String(sessionId));
      if(!session)return{success:false,message:'Your session has expired. Please sign in again.',code:'SESSION_INVALID'};
      t=HtmlService.createTemplateFromFile('client-dashboard');
      t.sessionId=String(sessionId);
    }else{
      t=HtmlService.createTemplateFromFile('client-landing');
      t.sessionId='';
    }
    t.clientId=CLIENT_CONFIG.clientId;
    t.clientName=CLIENT_CONFIG.clientName;
    t.primaryColor=CLIENT_CONFIG.primaryColor;
    t.logoUrl=CLIENT_CONFIG.logoUrl;
    t.applicationUrl=getClientApplicationUrlV10_();
    return{success:true,html:t.evaluate().getContent()};
  }catch(error){
    return{success:false,message:error&&error.message?error.message:'Unable to load BizOS page.',code:'PAGE_RENDER_ERROR'};
  }
}

`;
}

function generateClientCodeSafelyV10(settings){
  settings=settings||{};
  var base=generateClientCodeSafelyV9(settings);
  if(!base||!base.files)throw new Error('Client deployment package is empty.');
  base.files=base.files.map(function(file){
    if(!file)return file;

    if(file.name==='Code'){
      var code=typeof file.source==='string'?file.source:(typeof file.content==='string'?file.content:'');
      code+=buildClientUrlConfigV10_(settings);
      code+=buildClientLicenseGuardV10_();
      code+=buildClientNavigationV10_();
      file.source=code;
      if(typeof file.content==='string')file.content=code;
      return file;
    }

    if(file.name==='client-landing'){
      var landing=typeof file.source==='string'?file.source:(typeof file.content==='string'?file.content:'');
      var clientNavScript='<script>function openClientPage(page,params){params=params||{};var sid=String(params.sessionId||"");var b=document.getElementById("b");if(b){b.disabled=true;b.textContent="Opening workspace...";}google.script.run.withSuccessHandler(function(r){if(!r||!r.success){var m=document.getElementById("m");if(m)m.textContent=r&&r.message||"Unable to open BizOS.";if(b){b.disabled=false;b.textContent="Sign in";}return}document.open();document.write(r.html);document.close();}).withFailureHandler(function(){var m=document.getElementById("m");if(m)m.textContent="Unable to open BizOS. Please try again.";if(b){b.disabled=false;b.textContent="Sign in";}}).renderClientPageV10(page,sid);}</script><script>';
landing=landing.replace("<script>",clientNavScript);
      landing=landing.replace(
        "location.href='?page=dashboard&sessionId='+encodeURIComponent(r.sessionId)",
        "openClientPage('dashboard',{sessionId:r.sessionId})"
      );
      // Final login override: do not depend on an exact generated login string.
      // The client must remain inside its Apps Script page; dashboard rendering
      // is performed server-side and written into the current document.
      var safeLoginScript='<script>function login(){var e=document.getElementById("e").value.trim(),p=document.getElementById("p").value,b=document.getElementById("b"),m=document.getElementById("m");if(!e||!p){m.textContent="Enter your email and password.";return}b.disabled=true;b.textContent="Signing in...";google.script.run.withSuccessHandler(function(r){if(r&&r.success){try{localStorage.setItem("bizos_client_session",r.sessionId);localStorage.setItem("bizos_client_user",JSON.stringify(r.user||{}));}catch(ignore){}openClientPage("dashboard",{sessionId:r.sessionId});}else{m.textContent=r&&r.message||"Login failed";b.disabled=false;b.textContent="Sign in";}}).withFailureHandler(function(err){m.textContent=err&&err.message?"Login failed: "+err.message:"Login failed. Please try again.";b.disabled=false;b.textContent="Sign in";}).authenticateClient(e,p);}</script>';
      landing=landing.replace('</body>',safeLoginScript+'</body>');
      file.source=landing;
      if(typeof file.content==='string')file.content=landing;
      return file;
    }

    if(file.name==='client-dashboard'){
      var dashboard=typeof file.source==='string'?file.source:(typeof file.content==='string'?file.content:'');
            dashboard=dashboard.replace("<script>",'<script>function openClientPage(page){google.script.run.withSuccessHandler(function(r){if(!r||!r.success){document.body.innerHTML='<div style="font-family:Arial;padding:32px;text-align:center">Session ended. Please sign in again.</div>';return}document.open();document.write(r.html);document.close();}).withFailureHandler(function(){document.body.innerHTML='<div style="font-family:Arial;padding:32px;text-align:center">Please sign in again.</div>';}).renderClientPageV10(page,"");}</script><script>');
      dashboard=dashboard.replace(/location\.href='\?page=login'/g,"openClientPage('login')");
      file.source=dashboard;
      if(typeof file.content==='string')file.content=dashboard;
      return file;
    }

    if(file.name!=='appsscript')return file;
    var manifest={};
    try{manifest=JSON.parse(String(file.source||file.content||'{}'));}catch(e){throw new Error('Client appsscript manifest is invalid: '+e.message);}
    manifest.oauthScopes=manifest.oauthScopes||[];
    var mailScope='https://www.googleapis.com/auth/script.send_mail';
    if(manifest.oauthScopes.indexOf(mailScope)<0)manifest.oauthScopes.push(mailScope);
    file.source=JSON.stringify(manifest);
    if(typeof file.content==='string')file.content=file.source;
    return file;
  });
  return base;
}
