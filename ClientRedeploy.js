// ============================================================
// ClientRedeploy.js - SAFE EXISTING CLIENT REDEPLOYMENT
// ============================================================
// Updates an existing client Apps Script project in place.
// It never creates a second workspace, Script project, or deployment.
// Admin authentication is mandatory.
// ============================================================

function redeployClientByBusinessId(businessId, sessionId) {
  requireAdminSession_(sessionId);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return {success:false,code:'REDEPLOY_BUSY',message:'Another deployment operation is already in progress. Please try again shortly.'};
  }

  try {
    var wantedBusinessId = String(businessId || '').trim();
    if (!wantedBusinessId) {
      return {success:false,code:'BUSINESS_ID_REQUIRED',message:'Business ID is required.'};
    }

    var client = getActiveClientByBusinessIdForRedeploy_(wantedBusinessId);
    if (!client) {
      return {success:false,code:'CLIENT_NOT_FOUND',message:'No active client deployment was found for this business.'};
    }

    var required = ['clientId','email','clientName','businessId','sheetId','scriptId','deploymentId'];
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

    var settings = {
      clientId: client.clientId,
      clientName: client.clientName,
      primaryColor: client.primaryColor || '#2E7D32',
      logoUrl: client.logoUrl || '',
      email: client.email,
      sheetId: client.sheetId,
      businessId: client.businessId,
      masterApiUrl: getMasterApiUrl(),
      landingUrl: client.landingUrl || '',
      customDomain: client.customDomain || '',
      applicationUrl: client.webAppUrl || ''
    };

    console.log('CLIENT REDEPLOY START:', JSON.stringify({
      clientId: client.clientId,
      businessId: client.businessId,
      scriptId: client.scriptId,
      deploymentId: client.deploymentId
    }));

    var packageResult = generateClientCodeSafelyV10(settings);
    var packageCheck = validateClientDeploymentPackage_(packageResult, {
      clientId: client.clientId,
      sheetId: client.sheetId
    });

    if (!packageCheck || !packageCheck.success) {
      return {
        success:false,
        code:packageCheck && packageCheck.code ? packageCheck.code : 'PACKAGE_INVALID',
        message:packageCheck && packageCheck.message ? packageCheck.message : 'The generated V10 package failed validation.',
        clientId:client.clientId
      };
    }

    var files = packageResult.files.map(function(file) {
      return {
        name:String(file.name),
        type:String(file.type),
        source:String(file.source || file.content || '')
      };
    });

    var headers = {
      Authorization:'Bearer ' + ScriptApp.getOAuthToken(),
      'Content-Type':'application/json'
    };

    var base = 'https://script.googleapis.com/v1/projects/' + encodeURIComponent(client.scriptId);

    // Verify that the registered deployment still belongs to the registered
    // script project before touching project content.
    var deploymentResponse = UrlFetchApp.fetch(
      base + '/deployments/' + encodeURIComponent(client.deploymentId),
      {method:'get',headers:headers,muteHttpExceptions:true}
    );

    if (deploymentResponse.getResponseCode() < 200 || deploymentResponse.getResponseCode() >= 300) {
      return {
        success:false,
        code:'DEPLOYMENT_LOOKUP_FAILED',
        message:'The existing client deployment could not be verified: ' + deploymentResponse.getContentText(),
        clientId:client.clientId,
        scriptId:client.scriptId,
        deploymentId:client.deploymentId
      };
    }

    var deployment = JSON.parse(deploymentResponse.getContentText() || '{}');
    var deploymentConfig = deployment.deploymentConfig || {};

    if (String(deployment.deploymentId || '') !== String(client.deploymentId) ||
        String(deploymentConfig.scriptId || client.scriptId) !== String(client.scriptId)) {
      return {
        success:false,
        code:'DEPLOYMENT_BINDING_MISMATCH',
        message:'The registered deployment does not belong to the registered client script. No changes were made.',
        clientId:client.clientId,
        scriptId:client.scriptId,
        deploymentId:client.deploymentId
      };
    }

    // UpdateContent replaces the project's files. The V10 generator is
    // authoritative, so this intentionally publishes the complete package.
    var contentResponse = UrlFetchApp.fetch(base + '/content', {
      method:'put',
      headers:headers,
      payload:JSON.stringify({files:files}),
      muteHttpExceptions:true
    });

    if (contentResponse.getResponseCode() < 200 || contentResponse.getResponseCode() >= 300) {
      return {
        success:false,
        code:'PACKAGE_UPDATE_FAILED',
        message:'The existing client project could not be updated: ' + contentResponse.getContentText(),
        clientId:client.clientId,
        scriptId:client.scriptId,
        deploymentId:client.deploymentId
      };
    }

    // Create an immutable version from the updated project HEAD.
    var versionResponse = UrlFetchApp.fetch(base + '/versions', {
      method:'post',
      headers:headers,
      payload:JSON.stringify({
        description:'BizOS V10 redeploy ' + client.clientId + ' - ' + new Date().toISOString()
      }),
      muteHttpExceptions:true
    });

    if (versionResponse.getResponseCode() < 200 || versionResponse.getResponseCode() >= 300) {
      return {
        success:false,
        code:'VERSION_CREATE_FAILED',
        message:'The updated client project was saved, but a new version could not be created: ' + versionResponse.getContentText(),
        clientId:client.clientId,
        scriptId:client.scriptId,
        deploymentId:client.deploymentId,
        warning:'The existing deployment was not changed.'
      };
    }

    var versionResult = JSON.parse(versionResponse.getContentText() || '{}');
    var versionNumber = versionResult.versionNumber;
    if (!versionNumber) {
      return {
        success:false,
        code:'VERSION_NUMBER_MISSING',
        message:'Google did not return a version number. The existing deployment was not changed.',
        clientId:client.clientId,
        scriptId:client.scriptId,
        deploymentId:client.deploymentId
      };
    }

    // Update the SAME deployment. This preserves the deployment ID and URL.
    var updatePayload = {
      deploymentConfig:{
        scriptId:client.scriptId,
        versionNumber:versionNumber,
        manifestFileName:String(deploymentConfig.manifestFileName || 'appsscript'),
        description:'BizOS production deployment for ' + client.clientName
      }
    };

    var updateResponse = UrlFetchApp.fetch(
      base + '/deployments/' + encodeURIComponent(client.deploymentId),
      {
        method:'put',
        headers:headers,
        payload:JSON.stringify(updatePayload),
        muteHttpExceptions:true
      }
    );

    if (updateResponse.getResponseCode() < 200 || updateResponse.getResponseCode() >= 300) {
      return {
        success:false,
        code:'DEPLOYMENT_UPDATE_FAILED',
        message:'The new client version was created, but the existing deployment could not be updated: ' + updateResponse.getContentText(),
        clientId:client.clientId,
        scriptId:client.scriptId,
        deploymentId:client.deploymentId,
        versionNumber:versionNumber,
        warning:'The existing deployment remains on its previous version.'
      };
    }

    var expectedUrl = 'https://script.google.com/macros/s/' + client.deploymentId + '/exec';
    var updatedDeployment = JSON.parse(updateResponse.getContentText() || '{}');
    var updatedConfig = updatedDeployment.deploymentConfig || {};

    if (String(updatedDeployment.deploymentId || client.deploymentId) !== String(client.deploymentId) ||
        String(updatedConfig.scriptId || client.scriptId) !== String(client.scriptId) ||
        String(updatedConfig.versionNumber || '') !== String(versionNumber)) {
      return {
        success:false,
        code:'DEPLOYMENT_UPDATE_UNVERIFIED',
        message:'Google accepted the deployment update, but the returned deployment configuration could not be verified.',
        clientId:client.clientId,
        scriptId:client.scriptId,
        deploymentId:client.deploymentId,
        versionNumber:versionNumber
      };
    }

    // Keep the registry URL/IDs unchanged. Only touch the audit timestamp.
    updateClientRedeployAudit_(client.clientId, versionNumber, sessionId);

    return {
      success:true,
      code:'CLIENT_REDEPLOYED',
      clientId:client.clientId,
      businessId:client.businessId,
      scriptId:client.scriptId,
      deploymentId:client.deploymentId,
      versionNumber:versionNumber,
      webAppUrl:expectedUrl,
      message:'Client deployment updated successfully. The existing deployment URL was preserved.'
    };
  } catch (error) {
    console.error('redeployClientByBusinessId error:', error);
    return {
      success:false,
      code:'REDEPLOY_EXCEPTION',
      message:error && error.message ? error.message : 'Client redeployment failed.'
    };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function getActiveClientByBusinessIdForRedeploy_(businessId) {
  var ss = getBizOSMasterSpreadsheet_();
  var sheet = ss.getSheetByName('Clients');
  if (!sheet || sheet.getLastRow() < 2) return null;

  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function(h){ return String(h || '').trim(); });
  var businessIdCol = headers.indexOf('Business_ID');
  var statusCol = headers.indexOf('Status');

  if (businessIdCol < 0) throw new Error('Client registry is missing Business_ID.');
  if (statusCol < 0) throw new Error('Client registry is missing Status.');

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (String(row[businessIdCol] || '').trim() !== String(businessId).trim()) continue;
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
      tier:String(client.Tier || 'sovereign'),
      status:String(client.Status || ''),
      landingUrl:String(client.Landing_URL || ''),
      customDomain:String(client.Custom_Domain || ''),
      domain:String(client.Domain || '')
    };
  }

  return null;
}

function updateClientRedeployAudit_(clientId, versionNumber, sessionId) {
  try {
    var ss = getBizOSMasterSpreadsheet_();
    var sheet = ss.getSheetByName('Clients');
    if (!sheet || sheet.getLastRow() < 2) return;

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h || '').trim(); });
    var clientIdCol = headers.indexOf('Client_ID');
    var updatedAtCol = headers.indexOf('Updated_At');

    if (clientIdCol < 0) return;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][clientIdCol] || '').trim() !== String(clientId).trim()) continue;

      if (updatedAtCol >= 0) {
        sheet.getRange(i + 1, updatedAtCol + 1).setValue(new Date().toISOString());
      }

      try {
        logAdminAction('client_redeploy', {
          clientId:clientId,
          versionNumber:versionNumber,
          deploymentId:String(data[i][headers.indexOf('Deployment_ID')] || ''),
          scriptId:String(data[i][headers.indexOf('Script_ID')] || ''),
          performedBySession:sessionId || ''
        });
      } catch (ignoreLog) {}

      return;
    }
  } catch (error) {
    console.error('Redeploy audit update failed:', error);
  }
}
