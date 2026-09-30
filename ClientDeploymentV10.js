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
  return '\n// ============================================================\\n// CLIENT URL / LANDING CONFIGURATION - V10\\n// ============================================================\\nCLIENT_CONFIG.applicationUrl='+JSON.stringify(applicationUrl)+';\\nCLIENT_CONFIG.landingUrl='+JSON.stringify(landingUrl)+';\\nCLIENT_CONFIG.customDomain='+JSON.stringify(customDomain)+';\\nfunction getClientPublicUrlsV10(){\\n  var appUrl=String(CLIENT_CONFIG.applicationUrl||'').trim();\\n  try{if(!appUrl)appUrl=String(ScriptApp.getService().getUrl()||'').trim();}catch(ignore){}\\n  return {applicationUrl:appUrl,landingUrl:String(CLIENT_CONFIG.landingUrl||'').trim(),customDomain:String(CLIENT_CONFIG.customDomain||'').trim()};\\n}\\n';
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
      file.source=code;
      if(typeof file.content==='string')file.content=code;
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
