// ============================================================
// AutoSetupClientSafe.js - SAFE + IDEMPOTENT PAID CLIENT PROVISIONING
// ============================================================
// Explicit entry point. Does not override autoSetupClient().
// ============================================================

function autoSetupClientSafe(paymentData) {
  var lock = null;
  var clientId = '', sheetResult = null, scriptResult = null;
  var existingBusiness = null, businessId = null, businessName = '';
  var committed = false;

  try {
    paymentData = paymentData || {};
    if (!paymentData.email) return {success:false,code:'INVALID_PAYMENT_DATA',message:'Client email is required.'};
    businessName = paymentData.businessName || paymentData.name || 'My Business';

    lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) return {success:false,code:'PROVISIONING_BUSY',message:'Client provisioning is already in progress. Please retry shortly.'};

    console.log('SAFE PROVISIONING START:', paymentData.email, businessName);

    var existingClient = getExistingActiveClientDeployment_(paymentData.email, businessName);
    if (existingClient) {
      var existingResult = returnExistingClientDeployment_(existingClient);
      if (existingResult && existingResult.success) return existingResult;
      return existingResult || {success:false,code:'EXISTING_CLIENT_INCOMPLETE',message:'An active client already exists, but its deployment metadata is incomplete. No duplicate deployment was created.'};
    }

    var duplicateCheck = checkDuplicateClient(paymentData.email, businessName);
    if (duplicateCheck && duplicateCheck.isDuplicate) {
      var recheckedClient = getExistingActiveClientDeployment_(paymentData.email, businessName);
      if (recheckedClient) {
        var recheckedResult = returnExistingClientDeployment_(recheckedClient);
        if (recheckedResult) return recheckedResult;
      }
      return {success:false,code:'DUPLICATE_'+duplicateCheck.type,message:duplicateCheck.message,details:duplicateCheck};
    }

    clientId = 'CLIENT_' + Utilities.getUuid().substring(0,8).toUpperCase();
    existingBusiness = getBusinessByEmail(paymentData.email);
    businessId = existingBusiness ? existingBusiness.businessId : clientId;

    console.log('SAFE STEP 1: Creating workspace:', clientId);
    sheetResult = createClientSheetInClientDrive(paymentData.email, clientId, businessName, {
      businessId:businessId,
      tier:paymentData.tier || 'sovereign',
      status:'provisioning',
      primaryColor:paymentData.primaryColor || '#2E7D32',
      logoUrl:paymentData.logoUrl || '',
      provisioningVersion:'4.1'
    });
    if (!sheetResult || !sheetResult.success || !sheetResult.sheetId) {
      return {success:false,code:'WORKSPACE_CREATION_FAILED',message:'Client workspace could not be created: '+((sheetResult&&sheetResult.message)||'Unknown error')};
    }

    console.log('SAFE STEP 2: Generating V5 package');
    var clientCode = generateClientCodeSafelyV5({
      clientId:clientId,
      clientName:businessName,
      primaryColor:paymentData.primaryColor || '#2E7D32',
      logoUrl:paymentData.logoUrl || '',
      email:paymentData.email,
      sheetId:sheetResult.sheetId,
      businessId:businessId,
      masterApiUrl:getMasterApiUrl()
    });
    if (!clientCode || !clientCode.files || !clientCode.files.length) throw new Error('Client deployment package is empty.');

    var packageCheck = validateClientDeploymentPackage_(clientCode,{clientId:clientId,sheetId:sheetResult.sheetId});
    if (!packageCheck || !packageCheck.success) {
      cleanupFailedClientProvisioningSafe_(sheetResult,clientId);
      return {success:false,code:packageCheck&&packageCheck.code?packageCheck.code:'PACKAGE_INVALID',message:packageCheck&&packageCheck.message?packageCheck.message:'Generated client package failed validation.',clientId:clientId,cleanedUp:true};
    }

    console.log('SAFE STEP 3: Publishing exact V5 package');
    scriptResult = createAndDeployClientScriptSafe(paymentData.email,clientId,clientCode,businessName,paymentData.primaryColor || '#2E7D32',paymentData.logoUrl || '');
    if (!scriptResult || !scriptResult.success || !scriptResult.webAppUrl || !scriptResult.scriptId || !scriptResult.deploymentId) {
      var deploymentMessage = scriptResult && scriptResult.message ? scriptResult.message : 'The client application could not be deployed.';
      cleanupFailedClientProvisioningSafe_(sheetResult,clientId);
      cleanupFailedClientDeploymentSafe_(scriptResult);
      return {success:false,code:'DEPLOYMENT_FAILED',message:'Client provisioning failed: '+deploymentMessage,clientId:clientId,cleanedUp:true};
    }

    console.log('SAFE STEP 4: Saving full registry metadata');
    var saveResult = saveClientRecordV2({
      clientId:clientId,email:paymentData.email,clientName:businessName,
      domain:paymentData.domain || paymentData.customDomain || '',
      customDomain:paymentData.customDomain || '',
      primaryColor:paymentData.primaryColor || '#2E7D32',logoUrl:paymentData.logoUrl || '',
      tier:paymentData.tier || 'sovereign',status:'active',
      sheetId:sheetResult.sheetId,workspaceId:sheetResult.sheetId,
      webAppUrl:scriptResult.webAppUrl,landingUrl:scriptResult.webAppUrl,
      apiKey:Utilities.getUuid(),scriptId:scriptResult.scriptId,
      deploymentId:scriptResult.deploymentId,businessId:businessId,
      provisioningVersion:'4.1'
    });
    if (!saveResult || !saveResult.success) {
      cleanupFailedClientProvisioningSafe_(sheetResult,clientId);
      cleanupFailedClientDeploymentSafe_(scriptResult);
      return {success:false,code:'CLIENT_RECORD_SAVE_FAILED',message:'Deployment succeeded, but the client record could not be saved. Provisioning was rolled back.',clientId:clientId,cleanedUp:true};
    }

    committed = true;
    if (existingBusiness && businessId) updateBusinessWithClientInfo(businessId,clientId,scriptResult.webAppUrl);
    try { sendClientWelcomeEmail(paymentData.email,businessName,scriptResult.webAppUrl,clientId); }
    catch (emailError) { console.error('Client welcome email failed after successful provisioning:',emailError); }

    console.log('SAFE PROVISIONING COMPLETE:',clientId,scriptResult.webAppUrl);
    return {success:true,idempotent:false,clientId:clientId,landingUrl:scriptResult.webAppUrl,webAppUrl:scriptResult.webAppUrl,sheetUrl:sheetResult.sheetUrl,sheetId:sheetResult.sheetId,workspaceId:sheetResult.sheetId,email:paymentData.email,scriptId:scriptResult.scriptId,deploymentId:scriptResult.deploymentId,businessId:businessId,message:'Client setup complete.'};
  } catch (error) {
    console.error('SAFE CLIENT SETUP ERROR:',error);
    if (!committed) {
      if (sheetResult && sheetResult.success) cleanupFailedClientProvisioningSafe_(sheetResult,clientId);
      if (scriptResult && scriptResult.scriptId) cleanupFailedClientDeploymentSafe_(scriptResult);
    }
    return {success:false,code:'PROVISIONING_FAILED',message:error&&error.message?error.message:'Client provisioning failed.',clientId:clientId||null,cleanedUp:!committed};
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (ignore) {} }
  }
}

function cleanupFailedClientProvisioningSafe_(sheetResult,clientId) {
  try {
    if (!sheetResult || !sheetResult.sheetId) return;
    DriveApp.getFileById(sheetResult.sheetId).setTrashed(true);
  } catch (error) { console.error('Workspace cleanup failed for '+clientId,error); }
}

function cleanupFailedClientDeploymentSafe_(scriptResult) {
  try {
    if (!scriptResult || !scriptResult.scriptId) return;
    var url='https://script.googleapis.com/v1/projects/'+encodeURIComponent(scriptResult.scriptId);
    var response=UrlFetchApp.fetch(url,{method:'DELETE',headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});
    if (response.getResponseCode()<200 || response.getResponseCode()>=300) console.error('Client Apps Script cleanup failed:',response.getResponseCode(),response.getContentText());
  } catch (error) { console.error('Deployment cleanup failed:',error); }
}
