// ============================================================
// ProvisionConfirmedUpgrade.js
// Production bridge: confirmed upgrade -> safe client provisioning
// ============================================================

function provisionConfirmedUpgradeRequest(requestId) {
  var lock = null;
  try {
    if (!requestId) {
      return { success:false, code:'REQUEST_ID_REQUIRED', message:'Request ID is required.' };
    }

    lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
      return { success:false, code:'PROVISIONING_BUSY', message:'Provisioning is already in progress.' };
    }

    var request = getUpgradeRequest(requestId);
    if (!request) {
      return { success:false, code:'REQUEST_NOT_FOUND', message:'Upgrade request not found.' };
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Upgrade_Requests');
    if (!sheet) {
      return { success:false, code:'UPGRADE_SHEET_MISSING', message:'Upgrade request sheet not found.' };
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0] || [];
    var requestIdCol = headers.indexOf('Request_ID');
    var statusCol = headers.indexOf('Status');
    var updatedCol = headers.indexOf('Updated_At');
    if (requestIdCol === -1 || statusCol === -1) {
      return { success:false, code:'UPGRADE_SCHEMA_INVALID', message:'Upgrade request sheet is missing required columns.' };
    }

    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][requestIdCol]) === String(requestId)) {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex === -1) {
      return { success:false, code:'REQUEST_NOT_FOUND', message:'Upgrade request not found.' };
    }

    var status = String(data[rowIndex][statusCol] || '').toLowerCase();

    // Idempotent: an already-provisioned request must never create a second deployment.
    if (status === 'provisioned' || status === 'active') {
      var existing = getExistingActiveClientDeployment_(request.email, request.businessName);
      if (existing) {
        var existingResult = returnExistingClientDeployment_(existing);
        return {
          success:true,
          idempotent:true,
          clientId:existingResult.clientId,
          webAppUrl:existingResult.webAppUrl,
          landingUrl:existingResult.landingUrl,
          deploymentId:existingResult.deploymentId,
          message:'Client deployment already exists.'
        };
      }
    }

    // Only a server-side confirmed payment may enter provisioning.
    if (status !== 'payment_confirmed' && status !== 'provisioning_failed') {
      return {
        success:false,
        code:'PAYMENT_NOT_CONFIRMED',
        message:'The upgrade request has not reached a confirmed-payment state.'
      };
    }

    var businessDetails = request.businessDetails || {};
    var paymentData = {
      email:request.email,
      businessName:request.businessName,
      tier:request.tier || 'sovereign',
      domain:businessDetails.domain || businessDetails.Domain || '',
      customDomain:businessDetails.customDomain || businessDetails.Custom_Domain || '',
      primaryColor:businessDetails.primaryColor || businessDetails.Primary_Color || '#2E7D32',
      logoUrl:businessDetails.logoUrl || businessDetails.Logo_Url || ''
    };

    sheet.getRange(rowIndex + 1, statusCol + 1).setValue('provisioning');
    if (updatedCol !== -1) sheet.getRange(rowIndex + 1, updatedCol + 1).setValue(new Date().toISOString());

    var result = autoSetupClientSafe(paymentData);
    if (!result || !result.success) {
      sheet.getRange(rowIndex + 1, statusCol + 1).setValue('provisioning_failed');
      if (updatedCol !== -1) sheet.getRange(rowIndex + 1, updatedCol + 1).setValue(new Date().toISOString());
      return {
        success:false,
        code:'PROVISIONING_FAILED',
        paymentConfirmed:true,
        provisioning:result,
        message:'Payment is confirmed, but client provisioning failed. The request remains retryable.'
      };
    }

    sheet.getRange(rowIndex + 1, statusCol + 1).setValue('provisioned');
    if (updatedCol !== -1) sheet.getRange(rowIndex + 1, updatedCol + 1).setValue(new Date().toISOString());

    return {
      success:true,
      idempotent:!!result.idempotent,
      requestId:requestId,
      clientId:result.clientId,
      webAppUrl:result.webAppUrl,
      landingUrl:result.landingUrl,
      deploymentId:result.deploymentId,
      message:result.idempotent ? 'Existing client deployment returned.' : 'Client provisioned successfully.'
    };
  } catch (error) {
    console.error('provisionConfirmedUpgradeRequest error:', error);
    return { success:false, code:'PROVISIONING_BRIDGE_ERROR', message:error.message || 'Provisioning failed.' };
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (ignore) {} }
  }
}
