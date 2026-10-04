/**
 * V12 client runtime execution trace.
 *
 * This is an ADMIN-ONLY diagnostic. It does not redeploy or modify the client.
 * It first inspects the existing client deployment entry points. If the target
 * project has an Apps Script API executable entry point and the Master is
 * authorized to execute it, it calls the deployed runtime using a harmless
 * function that directly touches MODULES.
 */
function traceClientRuntimeExecutionApiV12(businessId, sessionId) {
  requireAdminSession_(sessionId);

  var wanted = String(businessId || '').trim();
  if (!wanted) return {success:false, code:'BUSINESS_ID_REQUIRED', message:'Business ID is required.'};

  try {
    var client = getActiveClientByBusinessIdForRedeploy_(wanted);
    if (!client) return {success:false, code:'CLIENT_NOT_FOUND', message:'Active client deployment was not found.'};

    var scriptId = String(client.scriptId || '').trim();
    var deploymentId = String(client.deploymentId || '').trim();
    if (!scriptId || !deploymentId) {
      return {
        success:false,
        code:'DEPLOYMENT_METADATA_INCOMPLETE',
        message:'Client deployment metadata is missing Script ID or Deployment ID.',
        clientId:client.clientId,
        scriptId:scriptId,
        deploymentId:deploymentId
      };
    }

    var auth = {Authorization:'Bearer '+ScriptApp.getOAuthToken()};
    var deploymentUrl = 'https://script.googleapis.com/v1/projects/'+encodeURIComponent(scriptId)+'/deployments/'+encodeURIComponent(deploymentId);
    var dr = UrlFetchApp.fetch(deploymentUrl, {method:'GET', headers:auth, muteHttpExceptions:true});
    var deploymentBody = dr.getContentText() || '';
    if (dr.getResponseCode() < 200 || dr.getResponseCode() >= 300) {
      return {
        success:false,
        code:'DEPLOYMENT_LOOKUP_FAILED',
        message:'The client deployment could not be read.',
        httpStatus:dr.getResponseCode(),
        details:deploymentBody,
        clientId:client.clientId,
        scriptId:scriptId,
        deploymentId:deploymentId
      };
    }

    var deployment = JSON.parse(deploymentBody || '{}');
    var entryPoints = Array.isArray(deployment.entryPoints) ? deployment.entryPoints : [];
    var executionApi = entryPoints.filter(function(ep) {
      return ep && String(ep.entryPointType || '').toLowerCase() === 'executionapi';
    });

    var webApp = entryPoints.filter(function(ep) {
      return ep && String(ep.entryPointType || '').toLowerCase() === 'webapp';
    });

    var result = {
      success:false,
      trace:'V12_CLIENT_RUNTIME_EXECUTION_API',
      clientId:client.clientId,
      businessId:client.businessId,
      email:client.email,
      scriptId:scriptId,
      deploymentId:deploymentId,
      deploymentVersion:String((deployment.deploymentConfig || {}).versionNumber || ''),
      deploymentEntryPoints:entryPoints.map(function(ep) {
        return {
          type:String(ep && ep.entryPointType || ''),
          hasWebApp:!!(ep && ep.webApp),
          hasExecutionApi:!!(ep && ep.executionApi)
        };
      }),
      webAppEntryPointPresent:webApp.length > 0,
      executionApiEntryPointPresent:executionApi.length > 0
    };

    if (!executionApi.length) {
      result.code = 'NO_EXECUTION_API_ENTRY_POINT';
      result.message = 'The existing client deployment is a web-app deployment only; it has no Apps Script API executable entry point. No client code was changed and no deployment was created.';
      result.nextStep = 'If remote execution is required, add a controlled API-executable deployment path. Do not change the public web-app deployment just for this diagnostic.';
      return result;
    }

    var runUrl = 'https://script.googleapis.com/v1/scripts/'+encodeURIComponent(scriptId)+':run';
    var payload = {
      function:'getFrontendConfig',
      parameters:[],
      devMode:false
    };

    var rr = UrlFetchApp.fetch(runUrl, {
      method:'POST',
      headers:{
        Authorization:'Bearer '+ScriptApp.getOAuthToken(),
        'Content-Type':'application/json'
      },
      payload:JSON.stringify(payload),
      muteHttpExceptions:true
    });

    var responseText = rr.getContentText() || '';
    result.httpStatus = rr.getResponseCode();
    result.executionResponse = responseText;

    if (rr.getResponseCode() < 200 || rr.getResponseCode() >= 300) {
      result.code = 'EXECUTION_API_REQUEST_FAILED';
      result.message = 'The Apps Script Execution API request failed.';
      return result;
    }

    var runBody = JSON.parse(responseText || '{}');
    if (runBody.error) {
      result.code = 'CLIENT_RUNTIME_EXCEPTION';
      result.message = 'The client runtime executed but returned an execution error.';
      result.executionError = runBody.error;
      return result;
    }

    result.success = true;
    result.code = 'CLIENT_RUNTIME_EXECUTED';
    result.message = 'The deployed client runtime executed getFrontendConfig successfully. MODULES is therefore available in that execution context.';
    result.result = runBody.response && runBody.response.result !== undefined
      ? runBody.response.result
      : null;
    return result;

  } catch (e) {
    return {
      success:false,
      code:'CLIENT_RUNTIME_TRACE_EXCEPTION',
      message:String(e && e.message || e),
      stack:String(e && e.stack || '')
    };
  }
}
