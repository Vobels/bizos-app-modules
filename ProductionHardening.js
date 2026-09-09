// ============================================================
// ProductionHardening.js - BIZOS PRODUCTION HARDENING
// ============================================================
// Scope: paid-client upgrade/provisioning, isolation, metadata,
// branding, backup/recovery and pre-production diagnostics.
// ============================================================

var BIZOS_BACKUP_FOLDER = 'BizOS_Client_Backups';
var BIZOS_SESSION_TTL = 21600;

(function () {
  var previousAutoSetup = autoSetupClient;

  autoSetupClient = function (paymentData) {
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000);
    } catch (e) {
      return { success:false, code:'PROVISIONING_BUSY', message:'Another client provisioning operation is already running. Please try again.' };
    }

    try {
      paymentData = paymentData || {};
      var email = String(paymentData.email || '').trim().toLowerCase();
      var businessName = String(paymentData.businessName || paymentData.name || 'My Business').trim();
      if (!email) return { success:false, code:'INVALID_PAYMENT_DATA', message:'Client email is required.' };

      // Idempotency: an existing active deployment for this account is returned,
      // not duplicated. This is important when a payment webhook is retried.
      var existing = findActiveClientByEmail_(email);
      if (existing) {
        return {
          success:true,
          existing:true,
          clientId:existing.clientId,
          landingUrl:existing.webAppUrl || '',
          sheetUrl:existing.sheetId ? 'https://docs.google.com/spreadsheets/d/' + existing.sheetId + '/edit' : '',
          email:email,
          scriptId:existing.scriptId || '',
          deploymentId:existing.deploymentId || existing.Deployment_ID || '',
          message:'Existing BizOS client deployment returned. No duplicate was created.'
        };
      }

      var result = previousAutoSetup(paymentData);
      if (!result || !result.success) return result;

      // Persist deployment metadata after the safe provisioning routine succeeds.
      // This is deliberately a second metadata write so legacy Clients sheets can
      // be upgraded without replacing the existing provisioning implementation.
      var metadataResult = recordDeploymentMetadata_(result.clientId, {
        scriptId:result.scriptId || '',
        deploymentId:result.deploymentId || '',
        sheetId:result.sheetId || '',
        webAppUrl:result.landingUrl || result.webAppUrl || '',
        businessId:result.businessId || '',
        email:email,
        clientName:businessName,
        primaryColor:paymentData.primaryColor || '#2E7D32',
        logoUrl:paymentData.logoUrl || ''
      });

      if (!metadataResult.success) {
        console.error('Deployment metadata write failed:', metadataResult.message);
        return {
          success:true,
          warning:'Client was provisioned, but deployment metadata could not be fully recorded. Vobels should repair the Clients record before production use.',
          clientId:result.clientId,
          landingUrl:result.landingUrl,
          scriptId:result.scriptId || '',
          deploymentId:result.deploymentId || ''
        };
      }

      result.sheetId = metadataResult.sheetId || result.sheetId || '';
      result.businessId = metadataResult.businessId || result.businessId || '';
      result.deploymentId = metadataResult.deploymentId || result.deploymentId || '';
      return result;
    } finally {
      try { lock.releaseLock(); } catch (e) {}
    }
  };
})();

function findActiveClientByEmail_(email) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Clients');
    if (!sh || sh.getLastRow() < 2) return null;
    var values = sh.getDataRange().getValues();
    var headers = values[0].map(function (x) { return String(x || '').trim(); });
    var emailCol = headers.indexOf('Email');
    var statusCol = headers.indexOf('Status');
    if (emailCol < 0) return null;
    for (var i = 1; i < values.length; i++) {
      var rowEmail = String(values[i][emailCol] || '').trim().toLowerCase();
      var status = statusCol >= 0 ? String(values[i][statusCol] || '').trim().toLowerCase() : 'active';
      if (rowEmail === email && status === 'active') return rowToClient_(headers, values[i]);
    }
  } catch (e) {
    console.error('findActiveClientByEmail_ failed:', e);
  }
  return null;
}

function rowToClient_(headers, row) {
  var out = {};
  headers.forEach(function (h, i) { out[h] = row[i] == null ? '' : row[i]; });
  out.clientId = out.Client_ID || '';
  out.email = out.Email || '';
  out.clientName = out.Client_Name || '';
  out.sheetId = out.Sheet_ID || '';
  out.webAppUrl = out.Web_App_URL || '';
  out.scriptId = out.Script_ID || out.Script_ID || '';
  out.deploymentId = out.Deployment_ID || '';
  out.businessId = out.Business_ID || '';
  return out;
}

function recordDeploymentMetadata_(clientId, meta) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Clients');
    if (!sh || sh.getLastRow() < 1) return {success:false, message:'Clients sheet not found.'};

    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x || '').trim(); });
    var required = ['Script_ID','Deployment_ID','Sheet_ID','Web_App_URL','Business_ID'];
    required.forEach(function (name) {
      if (headers.indexOf(name) < 0) {
        sh.insertColumnAfter(sh.getLastColumn());
        sh.getRange(1, sh.getLastColumn()).setValue(name);
        headers.push(name);
      }
    });

    var idCol = headers.indexOf('Client_ID');
    if (idCol < 0) return {success:false, message:'Clients sheet has no Client_ID column.'};
    var rows = sh.getDataRange().getValues();
    var rowNumber = -1;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][idCol] || '') === String(clientId)) { rowNumber = i + 1; break; }
    }
    if (rowNumber < 2) return {success:false, message:'Client record not found: ' + clientId};

    var values = {
      Script_ID:meta.scriptId || '',
      Deployment_ID:meta.deploymentId || '',
      Sheet_ID:meta.sheetId || '',
      Web_App_URL:meta.webAppUrl || '',
      Business_ID:meta.businessId || ''
    };
    Object.keys(values).forEach(function (name) {
      sh.getRange(rowNumber, headers.indexOf(name) + 1).setValue(values[name]);
    });
    return {success:true, sheetId:meta.sheetId || '', businessId:meta.businessId || '', deploymentId:meta.deploymentId || ''};
  } catch (e) {
    return {success:false, message:e.message};
  }
}

// ------------------------------------------------------------
// Client isolation guard
// ------------------------------------------------------------
// A deployment should use its immutable CLIENT_CONFIG. Browser input must
// never be trusted for clientId/sheetId. This helper is intended for any
// future server endpoint that receives client context from the browser.
function assertClientWorkspace_(clientId, sheetId, businessId) {
  if (typeof CLIENT_CONFIG === 'undefined') throw new Error('Client configuration is unavailable.');
  if (String(clientId || '') !== String(CLIENT_CONFIG.clientId || '')) throw new Error('Client identity mismatch.');
  if (String(sheetId || '') !== String(CLIENT_CONFIG.sheetId || '')) throw new Error('Workspace identity mismatch.');
  if (businessId && String(businessId) !== String(CLIENT_CONFIG.businessId || '')) throw new Error('Business identity mismatch.');
  return true;
}

// ------------------------------------------------------------
// Branding update helper
// ------------------------------------------------------------
function updateClientBranding(clientId, branding) {
  try {
    branding = branding || {};
    var client = getClientById(clientId);
    if (!client) return {success:false, message:'Client not found.'};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('Clients');
    if (!sh) return {success:false, message:'Clients sheet not found.'};
    var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
    var idCol = headers.indexOf('Client_ID');
    var row = -1;
    var vals = sh.getDataRange().getValues();
    for (var i=1;i<vals.length;i++) if (String(vals[i][idCol]||'')===String(clientId)) { row=i+1; break; }
    if (row<2) return {success:false,message:'Client record not found.'};
    var map = {Client_Name:branding.clientName,Primary_Color:branding.primaryColor,Logo_Url:branding.logoUrl,Custom_Domain:branding.customDomain};
    Object.keys(map).forEach(function(k){ if(map[k]!==undefined && headers.indexOf(k)>=0) sh.getRange(row,headers.indexOf(k)+1).setValue(map[k]); });
    return {success:true,message:'Branding updated.'};
  } catch(e) { return {success:false,message:e.message}; }
}

// ------------------------------------------------------------
// Backup / recovery
// ------------------------------------------------------------
function getOrCreateBizOSBackupFolder_() {
  var folders = DriveApp.getFoldersByName(BIZOS_BACKUP_FOLDER);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(BIZOS_BACKUP_FOLDER);
}

function backupClientWorkspace(clientId) {
  try {
    var client = getClientById(clientId);
    if (!client || !client.sheetId) return {success:false,message:'Client workspace not found.'};
    var source = DriveApp.getFileById(client.sheetId);
    var root = getOrCreateBizOSBackupFolder_();
    var subFolders = root.getFoldersByName(clientId);
    var folder = subFolders.hasNext() ? subFolders.next() : root.createFolder(clientId);
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Africa/Lagos', 'yyyyMMdd_HHmmss');
    var copy = source.makeCopy(client.clientName + ' - BizOS Backup - ' + stamp, folder);
    return {success:true,clientId:clientId,backupFileId:copy.getId(),backupUrl:copy.getUrl(),createdAt:new Date().toISOString()};
  } catch(e) { return {success:false,message:e.message}; }
}

function backupAllClientWorkspaces() {
  var results=[];
  try {
    var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clients');
    if(!sh || sh.getLastRow()<2) return {success:true,count:0,results:[]};
    var values=sh.getDataRange().getValues(),headers=values[0].map(String),idCol=headers.indexOf('Client_ID'),statusCol=headers.indexOf('Status');
    for(var i=1;i<values.length;i++){
      var status=statusCol>=0?String(values[i][statusCol]||'').toLowerCase():'active';
      if(status!=='active') continue;
      var id=String(values[i][idCol]||'');
      if(id) results.push(backupClientWorkspace(id));
    }
    return {success:true,count:results.length,results:results};
  } catch(e) { return {success:false,message:e.message,results:results}; }
}

function installBizOSBackupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t){ if(t.getHandlerFunction()==='backupAllClientWorkspaces') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('backupAllClientWorkspaces').timeBased().everyDays(1).atHour(2).create();
  return {success:true,message:'Daily BizOS client backup trigger installed.'};
}

function listClientBackups(clientId) {
  try {
    var root=getOrCreateBizOSBackupFolder_(),folders=root.getFoldersByName(clientId),out=[];
    if(!folders.hasNext()) return {success:true,backups:[]};
    var folder=folders.next(),files=folder.getFiles();
    while(files.hasNext()){var f=files.next();out.push({id:f.getId(),name:f.getName(),url:f.getUrl(),updatedAt:f.getLastUpdated().toISOString()});}
    out.sort(function(a,b){return String(b.updatedAt).localeCompare(String(a.updatedAt));});
    return {success:true,backups:out};
  } catch(e){return {success:false,message:e.message,backups:[]};}
}

// Recovery is intentionally Vobels-controlled: a backup is restored by
// replacing the live workspace contents after verification, not by creating
// a second writable production workspace.
function getRecoveryInstructions(clientId) {
  return {success:true,clientId:clientId,steps:[
    'Disable client access temporarily.',
    'Identify the verified backup snapshot in BizOS_Client_Backups.',
    'Restore/replace the existing live workspace contents.',
    'Verify Client_Info Client_ID, Business_ID and Status.',
    'Re-run client isolation and dashboard/module checks.',
    'Re-enable client access and record the recovery event.'
  ]};
}

// ------------------------------------------------------------
// Pre-production diagnostics / final test suite
// ------------------------------------------------------------
function runBizOSProductionDiagnostics() {
  var checks=[];
  function check(name, fn){try{var r=fn();checks.push({name:name,pass:r===true||!!(r&&r.success),detail:r});}catch(e){checks.push({name:name,pass:false,detail:e.message});}}

  check('Clients sheet exists', function(){return !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clients');});
  check('Master API URL configured', function(){return !!getMasterApiUrl();});
  check('No hardcoded test secret in properties', function(){return String(PropertiesService.getScriptProperties().getProperty('MASTER_API_SECRET')||'').trim() !== 'TEST_SECRET_2024';});
  check('Backup folder accessible', function(){return !!getOrCreateBizOSBackupFolder_();});
  check('Client records have workspace IDs', function(){
    var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clients'); if(!sh||sh.getLastRow()<2)return true;
    var v=sh.getDataRange().getValues(),h=v[0].map(String),id=h.indexOf('Client_ID'),ws=h.indexOf('Sheet_ID'),st=h.indexOf('Status');
    for(var i=1;i<v.length;i++) if(String(v[i][st]||'').toLowerCase()==='active' && (!v[i][id]||!v[i][ws])) return false;
    return true;
  });
  return {success:checks.every(function(x){return x.pass;}),checks:checks,runAt:new Date().toISOString()};
}
