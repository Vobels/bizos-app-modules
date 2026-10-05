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
