// ============================================================
// ProvisionConfirmedUpgrade.js
// Production bridge: confirmed upgrade -> safe client provisioning
// ============================================================

function provisionConfirmedUpgradeRequest(requestId) {
  if (!requestId) return {success:false,code:'REQUEST_ID_REQUIRED',message:'Request ID is required.'};

  return withClientProvisioningLock_(requestId, function() {
    try {
      var request = getUpgradeRequest(requestId);
      if (!request) return {success:false,code:'REQUEST_NOT_FOUND',message:'Upgrade request not found.'};

      var sheet = getBizOSMasterSpreadsheet_().getSheetByName('Upgrade_Requests');
      if (!sheet) return {success:false,code:'UPGRADE_SHEET_MISSING',message:'Upgrade request sheet not found.'};
      var data = sheet.getDataRange().getValues(),headers=data[0]||[];
      var requestIdCol=headers.indexOf('Request_ID'),statusCol=headers.indexOf('Status'),updatedCol=headers.indexOf('Updated_At');
      if(requestIdCol===-1||statusCol===-1)return{success:false,code:'UPGRADE_SCHEMA_INVALID',message:'Upgrade request sheet is missing required columns.'};

      var rowIndex=-1;
      for(var i=1;i<data.length;i++){if(String(data[i][requestIdCol])===String(requestId)){rowIndex=i;break;}}
      if(rowIndex===-1)return{success:false,code:'REQUEST_NOT_FOUND',message:'Upgrade request not found.'};
      var status=String(data[rowIndex][statusCol]||'').toLowerCase();

      // Always check for an already-provisioned client before starting a new
      // paid request. A customer can have more than one successful payment
      // (for example after a duplicate checkout), but that must never create
      // a second BizOS workspace.
      var existingEmail=request.workspaceEmail||request.email;
      var existing=getExistingActiveClientDeployment_(existingEmail,request.businessName);
      if(!existing && String(request.email||'').trim().toLowerCase() !== String(existingEmail||'').trim().toLowerCase()){
        existing=getExistingActiveClientDeployment_(request.email,request.businessName);
      }
      if(existing){
        var existingResult=returnExistingClientDeployment_(existing);
        if(existingResult&&existingResult.success){
          if(status!=='provisioned'){
            sheet.getRange(rowIndex+1,statusCol+1).setValue('provisioned');
            if(updatedCol!==-1)sheet.getRange(rowIndex+1,updatedCol+1).setValue(new Date().toISOString());
          }
          return{success:true,idempotent:true,reconciledDuplicatePayment:true,requestId:requestId,clientId:existingResult.clientId,webAppUrl:existingResult.webAppUrl,landingUrl:existingResult.landingUrl,deploymentId:existingResult.deploymentId,message:'Your payment is confirmed. Your existing BizOS workspace is ready. No second workspace was created.'};
        }
      }

      if(status!=='payment_confirmed'&&status!=='provisioning_failed'&&status!=='provisioning')return{success:false,code:'PAYMENT_NOT_CONFIRMED',message:'The payment for this upgrade has not been confirmed.'};

      var businessDetails=request.businessDetails||{};
      var workspaceEmail=String(request.workspaceEmail||request.email||'').trim().toLowerCase();
      var paymentData={
        email:workspaceEmail,
        accountEmail:request.email,
        businessName:request.businessName,
        tier:request.tier||'sovereign',
        domain:businessDetails.domain||businessDetails.Domain||'',
        customDomain:businessDetails.customDomain||businessDetails.Custom_Domain||'',
        primaryColor:businessDetails.primaryColor||businessDetails.Primary_Color||'#2E7D32',
        logoUrl:businessDetails.logoUrl||businessDetails.Logo_Url||''
      };

      sheet.getRange(rowIndex+1,statusCol+1).setValue('provisioning');
      if(updatedCol!==-1)sheet.getRange(rowIndex+1,updatedCol+1).setValue(new Date().toISOString());

      var result=autoSetupClientSafe(paymentData,true);
      if(!result||!result.success){
        sheet.getRange(rowIndex+1,statusCol+1).setValue('provisioning_failed');
        if(updatedCol!==-1)sheet.getRange(rowIndex+1,updatedCol+1).setValue(new Date().toISOString());
        try {
          PropertiesService.getScriptProperties().setProperty(
            'BIZOS_PROVISION_ERROR_' + requestId,
            JSON.stringify({
              at:new Date().toISOString(),
              code:result&&result.code||'PROVISIONING_FAILED',
              message:result&&result.message||'Client provisioning failed.',
              details:result&&result.details||null
            })
          );
        } catch(ignoreProvisionError){}
        return{success:false,code:'PROVISIONING_FAILED',paymentConfirmed:true,provisioning:result,message:'Your payment is confirmed, but your BizOS workspace could not be finished yet. You can continue setup without paying again.'};
      }

      sheet.getRange(rowIndex+1,statusCol+1).setValue('provisioned');
      if(updatedCol!==-1)sheet.getRange(rowIndex+1,updatedCol+1).setValue(new Date().toISOString());
      return{success:true,idempotent:!!result.idempotent,requestId:requestId,clientId:result.clientId,webAppUrl:result.webAppUrl,landingUrl:result.landingUrl,deploymentId:result.deploymentId,message:'Your BizOS workspace is ready.'};
    } catch(error){
      console.error('provisionConfirmedUpgradeRequest error:',error);
      try {
        PropertiesService.getScriptProperties().setProperty(
          'BIZOS_PROVISION_ERROR_' + requestId,
          JSON.stringify({
            at:new Date().toISOString(),
            code:'PROVISIONING_BRIDGE_ERROR',
            message:error&&error.message||'We could not finish setting up your BizOS workspace yet.'
          })
        );
      } catch(ignoreProvisionError){}
      return{success:false,code:'PROVISIONING_BRIDGE_ERROR',message:'We could not finish setting up your BizOS workspace yet.'};
    }
  });
}

function isUpgradeProvisioningWorkerInstalled_() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'processConfirmedUpgradeProvisioningQueue_') return true;
    }
  } catch (error) {
    console.error('isUpgradeProvisioningWorkerInstalled_ error:', error);
  }
  return false;
}

function ensureUpgradeProvisioningWorker_() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'processConfirmedUpgradeProvisioningQueue_') return true;
    }
    ScriptApp.newTrigger('processConfirmedUpgradeProvisioningQueue_').timeBased().everyMinutes(1).create();
    return true;
  } catch (error) {
    console.error('ensureUpgradeProvisioningWorker_ error:', error);
    return false;
  }
}

// Run this ONCE from the Apps Script editor as the deployment owner.
// This is intentionally separate from customer requests because Apps Script
// installable triggers require authorization from the account that creates them.
function setupUpgradeProvisioningWorker(spreadsheetId) {
  try {
    var spreadsheetSetup = configureBizOSMasterSpreadsheet_(spreadsheetId);
    if (!spreadsheetSetup.success) return spreadsheetSetup;
    if (typeof ScriptApp.requireScopes === 'function') {
      ScriptApp.requireScopes(ScriptApp.AuthMode.FULL, [
        'https://www.googleapis.com/auth/script.scriptapp',
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/script.send_mail'
      ]);
    }
    var ready = ensureUpgradeProvisioningWorker_();
    if (!ready) return {success:false,code:'WORKER_INSTALL_FAILED',message:'The provisioning worker could not be installed. Check the Apps Script execution log.'};
    return {success:true,message:'BizOS upgrade provisioning worker is installed and will process confirmed payments automatically.'};
  } catch (error) {
    console.error('setupUpgradeProvisioningWorker error:', error);
    return {success:false,code:'WORKER_INSTALL_AUTH_REQUIRED',message:'Please authorize the Apps Script project once, then run setupUpgradeProvisioningWorker again.'};
  }
}

function queueConfirmedUpgradeProvisioning(requestId) {
  if (!requestId) return {success:false,code:'REQUEST_ID_REQUIRED',message:'Request ID is required.'};

  // Customer requests must never create installable triggers. The worker is
  // installed once by the deployment owner and then processes all queued
  // confirmed payments, including payments made before this recovery flow.
  var workerReady = isUpgradeProvisioningWorkerInstalled_();
  return {
    success:workerReady,
    processing:workerReady,
    workerAvailable:workerReady,
    requestId:String(requestId),
    message:workerReady
      ? 'Your payment is confirmed. Workspace setup has been queued and will continue automatically.'
      : 'Your payment is confirmed. Workspace setup is not yet active. You do not need to pay again.'
  };
}

function reconcileSuccessfulUpgradePaymentsForProvisioning_() {
  try {
    ensureUpgradeRequestsSchema_();
    ensurePaymentsSchema_();

    var ss=getBizOSMasterSpreadsheet_();
    var upgradeSheet=ss.getSheetByName('Upgrade_Requests');
    var paymentSheet=ss.getSheetByName('Payments');
    if(!upgradeSheet||!paymentSheet)return 0;

    var upgradeData=upgradeSheet.getDataRange().getValues();
    var upgradeHeaders=upgradeData[0]||[];
    var reqCol=upgradeHeaders.indexOf('Request_ID');
    var statusCol=upgradeHeaders.indexOf('Status');
    var updatedCol=upgradeHeaders.indexOf('Updated_At');
    if(reqCol===-1||statusCol===-1)return 0;

    var paymentData=paymentSheet.getDataRange().getValues();
    var paymentHeaders=paymentData[0]||[];
    var pReqCol=paymentHeaders.indexOf('Request_ID');
    var pStatusCol=paymentHeaders.indexOf('Status');
    if(pReqCol===-1||pStatusCol===-1)return 0;

    var successful={};
    for(var p=1;p<paymentData.length;p++){
      var paymentStatus=String(paymentData[p][pStatusCol]||'').trim().toLowerCase();
      var paymentRequestId=String(paymentData[p][pReqCol]||'').trim();
      if(paymentStatus==='success'&&paymentRequestId)successful[paymentRequestId]=true;
    }

    var repaired=0;
    for(var i=1;i<upgradeData.length;i++){
      var requestId=String(upgradeData[i][reqCol]||'').trim();
      var status=String(upgradeData[i][statusCol]||'').trim().toLowerCase();
      if(!requestId||!successful[requestId])continue;
      if(status==='pending_payment'||!status){
        upgradeSheet.getRange(i+1,statusCol+1).setValue('payment_confirmed');
        if(updatedCol!==-1)upgradeSheet.getRange(i+1,updatedCol+1).setValue(new Date().toISOString());
        repaired++;
      }
    }
    return repaired;
  } catch(error){
    console.error('reconcileSuccessfulUpgradePaymentsForProvisioning_ error:',error);
    return 0;
  }
}

function processConfirmedUpgradeProvisioningQueue_() {
  try {
    // First recover historical successful payments whose upgrade row still
    // says pending_payment. This makes recovery automatic even when the
    // customer never revisits the original payment page.
    reconcileSuccessfulUpgradePaymentsForProvisioning_();

    var ss=getBizOSMasterSpreadsheet_();
    if(!ss) return;
    ensureUpgradeRequestsSchema_();
    var sheet=ss.getSheetByName('Upgrade_Requests');
    if(!sheet)return;
    var data=sheet.getDataRange().getValues(),headers=data[0]||[];
    var reqCol=headers.indexOf('Request_ID'),statusCol=headers.indexOf('Status');
    if(reqCol===-1||statusCol===-1)return;

    // Only successful payment records are eligible for background provisioning.
    // This prevents a stale/manual payment_confirmed status from ever creating
    // a client workspace without a matching successful payment.
    ensurePaymentsSchema_();
    var paymentSheet=ss.getSheetByName('Payments');
    var successfulPayments={};
    if(paymentSheet){
      var paymentData=paymentSheet.getDataRange().getValues(),paymentHeaders=paymentData[0]||[];
      var paymentReqCol=paymentHeaders.indexOf('Request_ID'),paymentStatusCol=paymentHeaders.indexOf('Status');
      if(paymentReqCol!==-1&&paymentStatusCol!==-1){
        for(var p=1;p<paymentData.length;p++){
          var paymentRequestId=String(paymentData[p][paymentReqCol]||'').trim();
          var paymentStatus=String(paymentData[p][paymentStatusCol]||'').trim().toLowerCase();
          if(paymentRequestId&&paymentStatus==='success')successfulPayments[paymentRequestId]=true;
        }
      }
    }

    // The actual provisioning function owns the per-request lock.
    // Do not hold the same ScriptLock here or provisioning would deadlock.
    var processed=0;
    for(var i=1;i<data.length&&processed<2;i++){
      var status=String(data[i][statusCol]||'').trim().toLowerCase();
      if(status!=='payment_confirmed'&&status!=='provisioning')continue;
      var requestId=String(data[i][reqCol]||'').trim();
      if(!requestId||!successfulPayments[requestId])continue;
      processed++;
      var result=provisionConfirmedUpgradeRequest(requestId);
      if(result&&result.success){
        try {
          if(typeof sendDeploymentConfirmationEmailOnce_==='function'){
            var req=getUpgradeRequest(requestId);
            sendDeploymentConfirmationEmailOnce_(req&&req.email||'',req&&req.businessName||'',requestId,result.webAppUrl||result.landingUrl||'');
          }
        } catch(emailError){console.error('Deployment confirmation email error:',emailError);}
      }
    }
  } catch(error) {
    console.error('processConfirmedUpgradeProvisioningQueue_ error:',error);
  }
}

function getDeploymentStatusPageData(requestId, sessionId) {
  try {
    var user = validateUpgradeSession(sessionId);
    if (!user) return {success:false,code:'UNAUTHORIZED',message:'Please sign in to view your deployment status.'};

    // When a status link contains a requestId, resolve that exact request.
    // Do not compare it with the user's newest request: a later unpaid request
    // must never hide an older confirmed payment/deployment.
    var latest = requestId ? getUpgradeRequest(String(requestId).trim()) : getLatestPaidUpgradeRequestForUser_(user.email);
    if (!latest) return {success:true,active:false,message:'No upgrade deployment was found for this account.'};

    if (String(latest.email || '').trim().toLowerCase() !== String(user.email || '').trim().toLowerCase()) {
      return {success:false,code:'REQUEST_ACCESS_DENIED',message:'This deployment request does not belong to the current BizOS account.'};
    }

    // Reconcile a successful payment row even if the upgrade row was not yet
    // updated. This keeps the deployment page consistent with the Payments sheet.
    if (requestId) {
      try {
        ensurePaymentsSchema_();
        var paymentSheet = getBizOSMasterSpreadsheet_().getSheetByName('Payments');
        if (paymentSheet) {
          var paymentData = paymentSheet.getDataRange().getValues();
          var paymentHeaders = paymentData[0] || [];
          var pReq = paymentHeaders.indexOf('Request_ID');
          var pStatus = paymentHeaders.indexOf('Status');
          if (pReq !== -1 && pStatus !== -1) {
            for (var pi = 1; pi < paymentData.length; pi++) {
              if (String(paymentData[pi][pReq] || '').trim() === String(latest.requestId || '').trim() &&
                  String(paymentData[pi][pStatus] || '').trim().toLowerCase() === 'success') {
                if (String(latest.status || '').trim().toLowerCase() === 'pending_payment') {
                  var upgradeSheet = getBizOSMasterSpreadsheet_().getSheetByName('Upgrade_Requests');
                  if (upgradeSheet) {
                    var upgradeData = upgradeSheet.getDataRange().getValues();
                    var uh = upgradeData[0] || [];
                    var ur = uh.indexOf('Request_ID'), us = uh.indexOf('Status');
                    if (ur !== -1 && us !== -1) {
                      for (var ui = 1; ui < upgradeData.length; ui++) {
                        if (String(upgradeData[ui][ur] || '').trim() === String(latest.requestId || '').trim()) {
                          upgradeSheet.getRange(ui + 1, us + 1).setValue('payment_confirmed');
                          latest.status = 'payment_confirmed';
                          break;
                        }
                      }
                    }
                  }
                }
                break;
              }
            }
          }
        }
      } catch (reconcileError) {
        console.error('Deployment status payment reconciliation error:', reconcileError);
      }
    }

    var status = String(latest.status || '').trim().toLowerCase();
    var provisioningError = '';
    if (status === 'provisioning_failed') {
      try {
        var rawProvisionError = PropertiesService.getScriptProperties().getProperty('BIZOS_PROVISION_ERROR_' + latest.requestId);
        if (rawProvisionError) {
          var parsedProvisionError = JSON.parse(rawProvisionError);
          provisioningError = String(parsedProvisionError.message || '');
        }
      } catch(ignoreProvisionStatusError){}
    }
    var result = {
      success:true,
      active:true,
      requestId:latest.requestId,
      paymentId:latest.paymentId,
      businessName:latest.businessName,
      tier:latest.tier,
      amount:latest.amount,
      currency:latest.currency,
      country:latest.country,
      workspaceEmail:latest.workspaceEmail || user.email,
      createdAt:latest.createdAt,
      status:status,
      paymentConfirmed:['payment_confirmed','provisioning','provisioning_failed','provisioned','active'].indexOf(status) !== -1,
      ready:false,
      processing:status === 'provisioning',
      provisioningError:provisioningError
    };

    if (status === 'provisioned' || status === 'active') {
      var existing = getExistingActiveClientDeployment_(latest.workspaceEmail || latest.email || user.email, latest.businessName);
      if (existing) {
        var ready = returnExistingClientDeployment_(existing);
        if (ready && ready.success) {
          result.ready = true;
          result.webAppUrl = ready.webAppUrl || '';
          result.landingUrl = ready.landingUrl || ready.webAppUrl || '';
          result.deploymentId = ready.deploymentId || '';
        }
      }
    }

    if (status === 'provisioning') {
      result.message = 'Your payment is confirmed and your BizOS workspace is currently being prepared.';
    } else if (status === 'provisioning_failed') {
      result.retryable = true;
      result.message = 'Your payment is confirmed, but workspace setup needs to be continued. You do not need to pay again.';
    } else if (result.ready) {
      result.message = 'Your payment is confirmed and your BizOS workspace is ready.';
    } else if (result.paymentConfirmed) {
      result.message = 'Your payment is confirmed. Your BizOS workspace is being prepared. You do not need to pay again.';
    } else if (status === 'pending_payment') {
      result.message = 'This upgrade request is waiting for payment.';
    } else {
      result.message = 'Your upgrade request is currently ' + status + '.';
    }

    return result;
  } catch (error) {
    console.error('getDeploymentStatusPageData error:', error);
    return {success:false,code:'DEPLOYMENT_STATUS_ERROR',message:'We could not load your deployment status right now.'};
  }
}

function getUpgradeProvisioningStatus(requestId,sessionId,accessToken){
  if(!requestId)return{success:false,code:'REQUEST_ID_REQUIRED',message:'Your setup request could not be found.'};
  try{
    var user=sessionId?validateUpgradeSession(sessionId):null;
    if(!user&&accessToken){var access=validatePaymentAccessToken_(accessToken,requestId);if(access)user={email:access.email};}
    if(!user)return{success:false,code:'UNAUTHORIZED',message:'Your payment-page access has expired. Please return to BizOS and reopen the payment request.'};
    var request=getUpgradeRequest(requestId);
    if(!request)return{success:false,code:'REQUEST_NOT_FOUND',message:'Your setup request could not be found.'};
    if(String(request.email||'').trim().toLowerCase()!==String(user.email||'').trim().toLowerCase())return{success:false,code:'REQUEST_ACCESS_DENIED',message:'This setup request does not belong to the current BizOS account.'};
    var status=String(request.status||'').trim().toLowerCase();
    if(status==='provisioned'||status==='active'){
      var existing=getExistingActiveClientDeployment_(request.workspaceEmail||request.email,request.businessName);
      if(existing){var ready=returnExistingClientDeployment_(existing);if(ready&&ready.success)return ready;}
      return{success:false,code:'PROVISIONED_URL_UNAVAILABLE',status:status,message:'Your workspace is provisioned, but its launch link is not available yet.'};
    }
    if(status==='provisioning')return{success:true,ready:false,processing:true,status:status,message:'Your payment is confirmed and your BizOS workspace is still being prepared.'};
    if(status==='payment_confirmed'){
      var queuedStatus = queueConfirmedUpgradeProvisioning(requestId);
      return {
        success:true,
        ready:false,
        processing:!!queuedStatus.success,
        status:status,
        workerAvailable:!!queuedStatus.success,
        setupQueued:!!queuedStatus.success,
        message:queuedStatus.success
          ? 'Your payment is confirmed. Workspace setup has been queued and will continue automatically.'
          : 'Your payment is confirmed. Workspace setup worker is not active yet. You do not need to pay again.'
      };
    }
    if(status==='provisioning_failed')return{success:true,ready:false,processing:false,retryable:true,status:status,message:'Your payment is confirmed, but workspace setup needs to be continued.'};
    return{success:false,code:'PAYMENT_NOT_CONFIRMED',status:status,message:'Your payment has not been confirmed for setup yet.'};
  }catch(error){console.error('getUpgradeProvisioningStatus error:',error);return{success:false,code:'PROVISIONING_STATUS_ERROR',message:'We could not check workspace setup status right now.'};}
}

function continueConfirmedUpgradeSetup(requestId,sessionId,accessToken){
  if(!requestId)return{success:false,code:'REQUEST_ID_REQUIRED',message:'Your setup request could not be found.'};
  try{
    var user=sessionId?validateUpgradeSession(sessionId):null;
    if(!user&&accessToken){
      var access=validatePaymentAccessToken_(accessToken,requestId);
      if(access)user={email:access.email};
    }
    if(!user)return{success:false,code:'UNAUTHORIZED',message:'Your payment-page access has expired. Please return to BizOS and reopen the payment request.'};

    var request=getUpgradeRequest(requestId);
    if(!request)return{success:false,code:'REQUEST_NOT_FOUND',message:'Your setup request could not be found.'};
    if(String(request.email||'').trim().toLowerCase()!==String(user.email||'').trim().toLowerCase()){
      return{success:false,code:'REQUEST_ACCESS_DENIED',message:'This setup request does not belong to the current BizOS account.'};
    }

    var status=String(request.status||'').toLowerCase();
    if(status==='provisioned'||status==='active'){
      var existingResult = provisionConfirmedUpgradeRequest(requestId);
      return existingResult;
    }
    if(status==='payment_confirmed'||status==='provisioning_failed'||status==='provisioning'){
      if(status==='provisioning_failed'){
        var upgradeSheet=getBizOSMasterSpreadsheet_().getSheetByName('Upgrade_Requests');
        if(upgradeSheet){
          var ud=upgradeSheet.getDataRange().getValues(), uh=ud[0]||[], ur=uh.indexOf('Request_ID'), us=uh.indexOf('Status'), uu=uh.indexOf('Updated_At');
          for(var ui=1;ui<ud.length;ui++){
            if(String(ud[ui][ur]||'').trim()===String(requestId).trim()){
              if(us!==-1)upgradeSheet.getRange(ui+1,us+1).setValue('payment_confirmed');
              if(uu!==-1)upgradeSheet.getRange(ui+1,uu+1).setValue(new Date().toISOString());
              break;
            }
          }
        }
      }
      var queued=queueConfirmedUpgradeProvisioning(requestId);
      return {
        success:true,
        paymentConfirmed:true,
        processing:true,
        requestId:requestId,
        message:queued.message||'Your payment is confirmed. Workspace setup has been started. Keep this page open or return later to check your launch link.'
      };
    }
    return{success:false,code:'PAYMENT_NOT_CONFIRMED',message:'We have not confirmed this payment yet. Please wait a moment and try again.'};
  }catch(error){
    console.error('continueConfirmedUpgradeSetup error:',error);
    return{success:false,code:'SETUP_RETRY_ERROR',message:'We could not continue setup right now. Please try again.'};
  }
}
