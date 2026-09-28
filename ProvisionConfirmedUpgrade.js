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

      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Upgrade_Requests');
      if (!sheet) return {success:false,code:'UPGRADE_SHEET_MISSING',message:'Upgrade request sheet not found.'};
      var data = sheet.getDataRange().getValues(),headers=data[0]||[];
      var requestIdCol=headers.indexOf('Request_ID'),statusCol=headers.indexOf('Status'),updatedCol=headers.indexOf('Updated_At');
      if(requestIdCol===-1||statusCol===-1)return{success:false,code:'UPGRADE_SCHEMA_INVALID',message:'Upgrade request sheet is missing required columns.'};

      var rowIndex=-1;
      for(var i=1;i<data.length;i++){if(String(data[i][requestIdCol])===String(requestId)){rowIndex=i;break;}}
      if(rowIndex===-1)return{success:false,code:'REQUEST_NOT_FOUND',message:'Upgrade request not found.'};
      var status=String(data[rowIndex][statusCol]||'').toLowerCase();

      if(status==='provisioned'||status==='active'||status==='provisioning'){
        var existingEmail=request.workspaceEmail||request.email;
        var existing=getExistingActiveClientDeployment_(existingEmail,request.businessName);
        if(!existing&&existingEmail!==request.email)existing=getExistingActiveClientDeployment_(request.email,request.businessName);
        if(existing){
          var existingResult=returnExistingClientDeployment_(existing);
          if(existingResult&&existingResult.success){
            if(status!=='provisioned'){sheet.getRange(rowIndex+1,statusCol+1).setValue('provisioned');if(updatedCol!==-1)sheet.getRange(rowIndex+1,updatedCol+1).setValue(new Date().toISOString());}
            return{success:true,idempotent:true,requestId:requestId,clientId:existingResult.clientId,webAppUrl:existingResult.webAppUrl,landingUrl:existingResult.landingUrl,deploymentId:existingResult.deploymentId,message:'Your BizOS workspace is ready.'};
          }
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
        return{success:false,code:'PROVISIONING_FAILED',paymentConfirmed:true,provisioning:result,message:'Your payment is confirmed, but your BizOS workspace could not be finished yet. You can continue setup without paying again.'};
      }

      sheet.getRange(rowIndex+1,statusCol+1).setValue('provisioned');
      if(updatedCol!==-1)sheet.getRange(rowIndex+1,updatedCol+1).setValue(new Date().toISOString());
      return{success:true,idempotent:!!result.idempotent,requestId:requestId,clientId:result.clientId,webAppUrl:result.webAppUrl,landingUrl:result.landingUrl,deploymentId:result.deploymentId,message:'Your BizOS workspace is ready.'};
    } catch(error){
      console.error('provisionConfirmedUpgradeRequest error:',error);
      return{success:false,code:'PROVISIONING_BRIDGE_ERROR',message:'We could not finish setting up your BizOS workspace yet.'};
    }
  });
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
        var paymentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payments');
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
                  var upgradeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Upgrade_Requests');
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
      processing:status === 'provisioning'
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
    if(status==='payment_confirmed')return{success:true,ready:false,processing:false,status:status,message:'Your payment is confirmed. Workspace setup has not finished yet.'};
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
    if(status==='provisioned'||status==='active'||status==='payment_confirmed'||status==='provisioning_failed'||status==='provisioning'){
      var result=provisionConfirmedUpgradeRequest(requestId);
      if(result&&result.success)return result;
      return {
        success:false,
        code:result&&result.code||'SETUP_RETRY_ERROR',
        paymentConfirmed:true,
        message:result&&result.message||'Your payment is safe, but setup is not finished yet. Please try again in a moment.'
      };
    }
    return{success:false,code:'PAYMENT_NOT_CONFIRMED',message:'We have not confirmed this payment yet. Please wait a moment and try again.'};
  }catch(error){
    console.error('continueConfirmedUpgradeSetup error:',error);
    return{success:false,code:'SETUP_RETRY_ERROR',message:'We could not continue setup right now. Please try again.'};
  }
}
