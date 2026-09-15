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

function continueConfirmedUpgradeSetup(requestId){
  if(!requestId)return{success:false,code:'REQUEST_ID_REQUIRED',message:'Your setup request could not be found.'};
  try{
    var request=getUpgradeRequest(requestId);
    if(!request)return{success:false,code:'REQUEST_NOT_FOUND',message:'Your setup request could not be found.'};
    var status=String(request.status||'').toLowerCase();
    if(status==='provisioned'||status==='active')return provisionConfirmedUpgradeRequest(requestId);
    if(status!=='payment_confirmed'&&status!=='provisioning_failed'&&status!=='provisioning')return{success:false,code:'PAYMENT_NOT_CONFIRMED',message:'We have not confirmed this payment yet. Please wait a moment and try again.'};
    return provisionConfirmedUpgradeRequest(requestId);
  }catch(error){console.error('continueConfirmedUpgradeSetup error:',error);return{success:false,code:'SETUP_RETRY_ERROR',message:'We could not continue setup right now. Please try again.'};}
}
