/**
 * V12 client runtime execution trace.
 *
 * This is an ADMIN-ONLY diagnostic. It does not redeploy or modify the client.
 * It first inspects the existing client deployment entry points. If the target
 * project has an Apps Script API executable entry point and the Master is
 * authorized to execute it, it calls the deployed runtime using a harmless
 * function that directly touches MODULES.
 */
function cleanupExpiredV12RuntimeTraceProperties_() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var now = Date.now();
  var removed = 0;
  Object.keys(all).forEach(function(key) {
    if (key.indexOf('V12_RUNTIME_TRACE_') !== 0) return;
    try {
      var record = JSON.parse(String(all[key] || '{}'));
      var expiresAt = Number(record.expiresAt || 0);
      if (expiresAt && expiresAt <= now) {
        props.deleteProperty(key);
        removed++;
      }
    } catch (ignore) {
      // Leave malformed diagnostics untouched; hygiene must never delete
      // unknown or potentially useful state merely because it cannot be parsed.
    }
  });
  return removed;
}

function cleanupBizOSProvisionErrorProperties_(retentionDays) {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var now = Date.now();
  var retentionMs = Math.max(1, Number(retentionDays || 30)) * 24 * 60 * 60 * 1000;
  var removed = 0;
  Object.keys(all).forEach(function(key) {
    if (key.indexOf('BIZOS_PROVISION_ERROR_') !== 0) return;
    try {
      var record = JSON.parse(String(all[key] || '{}'));
      var at = Date.parse(String(record.at || ''));
      if (at && now - at >= retentionMs) {
        props.deleteProperty(key);
        removed++;
      }
    } catch (ignore) {
      // Keep malformed provisioning diagnostics for manual inspection.
    }
  });
  return removed;
}

function cleanupBizOSTemporaryScriptProperties(sessionId) {
  requireAdminSession_(sessionId);
  try {
    var tracesRemoved = cleanupExpiredV12RuntimeTraceProperties_();
    var provisionErrorsRemoved = cleanupBizOSProvisionErrorProperties_(30);
    return {
      success:true,
      v12RuntimeTraceRemoved:tracesRemoved,
      provisioningErrorsRemoved:provisionErrorsRemoved,
      emailIdempotencyFlagsPreserved:true,
      message:'Temporary BizOS diagnostics were cleaned up. Email delivery protection was preserved.'
    };
  } catch (error) {
    console.error('cleanupBizOSTemporaryScriptProperties error:', error);
    return {success:false,code:'PROPERTY_CLEANUP_ERROR',message:'We could not complete property cleanup right now.'};
  }
}

function runBizOSPropertyHygieneBestEffort_() {
  try {
    var cache = CacheService.getScriptCache();
    var gate = 'bizos_property_hygiene_last_run_v1';
    if (cache.get(gate)) return {success:true,skipped:true};
    var tracesRemoved = cleanupExpiredV12RuntimeTraceProperties_();
    var provisionErrorsRemoved = cleanupBizOSProvisionErrorProperties_(30);
    cache.put(gate, String(Date.now()), 6 * 60 * 60);
    return {success:true,skipped:false,v12RuntimeTraceRemoved:tracesRemoved,provisioningErrorsRemoved:provisionErrorsRemoved};
  } catch (error) {
    console.error('runBizOSPropertyHygieneBestEffort_ error:', error);
    return {success:false};
  }
}

function auditBizOSScriptPropertiesHygiene(sessionId) {
  requireAdminSession_(sessionId);
  try {
    var all = PropertiesService.getScriptProperties().getProperties();
    var now = Date.now();
    var counts = {
      total:Object.keys(all).length,
      v12RuntimeTrace:0,
      expiredV12RuntimeTrace:0,
      provisioningErrors:0,
      oldProvisioningErrors:0,
      paymentEmailFlags:0,
      deploymentEmailFlags:0,
      other:0
    };
    Object.keys(all).forEach(function(key) {
      if (key.indexOf('V12_RUNTIME_TRACE_') === 0) {
        counts.v12RuntimeTrace++;
        try {
          var trace = JSON.parse(String(all[key] || '{}'));
          if (Number(trace.expiresAt || 0) && Number(trace.expiresAt) <= now) counts.expiredV12RuntimeTrace++;
        } catch (ignoreTrace) {}
      } else if (key.indexOf('BIZOS_PROVISION_ERROR_') === 0) {
        counts.provisioningErrors++;
        try {
          var errorRecord = JSON.parse(String(all[key] || '{}'));
          var at = Date.parse(String(errorRecord.at || ''));
          if (at && now - at >= 30 * 24 * 60 * 60 * 1000) counts.oldProvisioningErrors++;
        } catch (ignoreError) {}
      } else if (key.indexOf('PAYMENT_CONFIRMATION_EMAIL_SENT_') === 0) {
        counts.paymentEmailFlags++;
      } else if (key.indexOf('DEPLOYMENT_CONFIRMATION_EMAIL_SENT_') === 0) {
        counts.deploymentEmailFlags++;
      } else {
        counts.other++;
      }
    });
    return {success:true,counts:counts,emailIdempotencyFlagsPreserved:true};
  } catch (error) {
    console.error('auditBizOSScriptPropertiesHygiene error:', error);
    return {success:false,code:'PROPERTY_AUDIT_ERROR',message:'We could not inspect Script Properties right now.'};
  }
}

function traceClientRuntimeExecutionApiV12(businessId, sessionId) {
  requireAdminSession_(sessionId);
  cleanupExpiredV12RuntimeTraceProperties_();
  var wanted=String(businessId||'').trim();
  if(!wanted)return{success:false,code:'BUSINESS_ID_REQUIRED',message:'Business ID is required.'};
  try{
    var client=getActiveClientByBusinessIdForRedeploy_(wanted);
    if(!client)return{success:false,code:'CLIENT_NOT_FOUND',message:'Active client deployment was not found.'};
    var webAppUrl=String(client.webAppUrl||'').trim();
    if(!webAppUrl)return{success:false,code:'WEB_APP_URL_MISSING',message:'Client deployment metadata is missing Web App URL.'};
    var token=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
    var record={clientId:String(client.clientId||''),businessId:String(client.businessId||''),expiresAt:Date.now()+10*60*1000};
    PropertiesService.getScriptProperties().setProperty('V12_RUNTIME_TRACE_'+token,JSON.stringify(record));
    var separator=webAppUrl.indexOf('?')>=0?'&':'?';
    return{success:true,trace:'V12_CLIENT_RUNTIME_WEBAPP',clientId:client.clientId,businessId:client.businessId,email:client.email,scriptId:client.scriptId,deploymentId:client.deploymentId,diagnosticUrl:webAppUrl+separator+'v12TraceToken='+encodeURIComponent(token),expiresInSeconds:600,message:'Short-lived runtime diagnostic URL created. It executes inside the existing client web-app deployment and does not redeploy or modify the client.'};
  }catch(e){return{success:false,code:'CLIENT_RUNTIME_TRACE_EXCEPTION',message:String(e&&e.message||e),stack:String(e&&e.stack||'')};}
}
