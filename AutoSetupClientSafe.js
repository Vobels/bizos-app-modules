// ============================================================
// AutoSetupClientSafe.js - SAFE PAID CLIENT PROVISIONING
// ============================================================
// Paid clients are provisioned into a separate workspace and a
// separate generated application. The legacy fallback path is not used.
// ============================================================
(function () {
  autoSetupClient = function (paymentData) {
    var clientId = '', sheetResult = null, scriptResult = null, existingBusiness = null, businessId = null, businessName = '';
    try {
      paymentData = paymentData || {};
      if (!paymentData.email) return { success:false, code:'INVALID_PAYMENT_DATA', message:'Client email is required.' };
      businessName = paymentData.businessName || paymentData.name || 'My Business';
      clientId = 'CLIENT_' + Utilities.getUuid().substring(0,8).toUpperCase();
      var duplicateCheck = checkDuplicateClient(paymentData.email,businessName);
      if (duplicateCheck && duplicateCheck.isDuplicate) return {success:false,code:'DUPLICATE_'+duplicateCheck.type,message:duplicateCheck.message,details:duplicateCheck};
      existingBusiness = getBusinessByEmail(paymentData.email);
      businessId = existingBusiness ? existingBusiness.businessId : clientId;
      sheetResult = createClientSheetInClientDrive(paymentData.email,clientId,businessName,{businessId:businessId,tier:paymentData.tier||'sovereign',status:'provisioning',primaryColor:paymentData.primaryColor||'#2E7D32',logoUrl:paymentData.logoUrl||'',provisioningVersion:'3.0'});
      if (!sheetResult || !sheetResult.success || !sheetResult.sheetId) return {success:false,code:'WORKSPACE_CREATION_FAILED',message:'Client workspace could not be created: '+((sheetResult&&sheetResult.message)||'Unknown error')};
      var masterUrl = getMasterApiUrl();
      var clientCode = generateClientCodeSafelyV2({clientId:clientId,clientName:businessName,primaryColor:paymentData.primaryColor||'#2E7D32',logoUrl:paymentData.logoUrl||'',email:paymentData.email,sheetId:sheetResult.sheetId,businessId:businessId,masterApiUrl:masterUrl});
      if (!clientCode || !clientCode.files || !clientCode.files.length) throw new Error('Client deployment package is empty.');
      scriptResult = createAndDeployClientScript(paymentData.email,clientId,clientCode,businessName,paymentData.primaryColor||'#2E7D32',paymentData.logoUrl||'');
      if (!scriptResult || !scriptResult.success || !scriptResult.webAppUrl || !scriptResult.scriptId) {
        var deploymentMessage = scriptResult && scriptResult.message ? scriptResult.message : 'The client application could not be deployed.';
        cleanupFailedClientProvisioning_(sheetResult,clientId); cleanupFailedClientDeployment_(scriptResult);
        return {success:false,code:'DEPLOYMENT_FAILED',message:'Client provisioning failed: '+deploymentMessage,clientId:clientId,cleanedUp:true};
      }
      var saveResult = saveClientRecord({clientId:clientId,email:paymentData.email,clientName:businessName,domain:paymentData.domain||paymentData.customDomain||'',customDomain:paymentData.customDomain||'',primaryColor:paymentData.primaryColor||'#2E7D32',logoUrl:paymentData.logoUrl||'',tier:paymentData.tier||'sovereign',status:'active',sheetId:sheetResult.sheetId,webAppUrl:scriptResult.webAppUrl,landingUrl:scriptResult.webAppUrl,apiKey:Utilities.getUuid(),scriptId:scriptResult.scriptId,businessId:businessId});
      if (!saveResult || !saveResult.success) { cleanupFailedClientProvisioning_(sheetResult,clientId); cleanupFailedClientDeployment_(scriptResult); return {success:false,code:'CLIENT_RECORD_SAVE_FAILED',message:'Deployment succeeded, but the client record could not be saved. Provisioning was rolled back.',clientId:clientId,cleanedUp:true}; }
      if (existingBusiness && businessId) updateBusinessWithClientInfo(businessId,clientId,scriptResult.webAppUrl);
      sendClientWelcomeEmail(paymentData.email,businessName,scriptResult.webAppUrl,clientId);
      return {success:true,clientId:clientId,landingUrl:scriptResult.webAppUrl,sheetUrl:sheetResult.sheetUrl,email:paymentData.email,scriptId:scriptResult.scriptId,deploymentId:scriptResult.deploymentId,message:'Client setup complete.'};
    } catch (error) {
      console.error('SAFE CLIENT SETUP ERROR:',error);
      if (sheetResult && sheetResult.success) cleanupFailedClientProvisioning_(sheetResult,clientId);
      if (scriptResult && !scriptResult.success) cleanupFailedClientDeployment_(scriptResult);
      return {success:false,code:'PROVISIONING_FAILED',message:error&&error.message?error.message:'Client provisioning failed.',clientId:clientId||null,cleanedUp:!!sheetResult};
    }
  };
  function cleanupFailedClientProvisioning_(sheetResult,clientId){try{if(!sheetResult||!sheetResult.sheetId)return;DriveApp.getFileById(sheetResult.sheetId).setTrashed(true);}catch(error){console.error('Workspace cleanup failed for '+clientId,error);}}
  function cleanupFailedClientDeployment_(scriptResult){try{if(!scriptResult||!scriptResult.scriptId)return;var url='https://script.googleapis.com/v1/projects/'+encodeURIComponent(scriptResult.scriptId),response=UrlFetchApp.fetch(url,{method:'delete',headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});if(response.getResponseCode()<200||response.getResponseCode()>=300)console.error('Client Apps Script cleanup failed:',response.getResponseCode(),response.getContentText());}catch(error){console.error('Deployment cleanup failed:',error);}}
})();
