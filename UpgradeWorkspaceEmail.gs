/**
 * Update the destination email for an upgrade after the existing request
 * has been created. The account/session email remains unchanged.
 */
function updateUpgradeRequestWorkspaceEmail(requestId, workspaceEmail) {
  try {
    var email = String(workspaceEmail || '').trim().toLowerCase();
    if (!requestId) return {success:false, message:'Upgrade request not found.'};
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return {success:false, message:'Please provide a valid workspace email.'};
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var upgradeSheet = ss.getSheetByName('Upgrade_Requests');
    if (!upgradeSheet) return {success:false, message:'Upgrade request could not be found.'};

    var data = upgradeSheet.getDataRange().getValues();
    if (!data.length) return {success:false, message:'Upgrade request could not be found.'};
    var headers = data[0];
    var requestCol = headers.indexOf('Request_ID');
    var emailCol = headers.indexOf('Email');
    var updatedCol = headers.indexOf('Updated_At');
    if (requestCol === -1 || emailCol === -1) return {success:false, message:'Upgrade request format is invalid.'};

    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][requestCol]) === String(requestId)) {
        upgradeSheet.getRange(i + 1, emailCol + 1).setValue(email);
        if (updatedCol !== -1) upgradeSheet.getRange(i + 1, updatedCol + 1).setValue(new Date().toISOString());
        found = true;
        break;
      }
    }
    if (!found) return {success:false, message:'Upgrade request could not be found.'};

    var paymentSheet = ss.getSheetByName('Payments');
    if (paymentSheet) {
      var paymentData = paymentSheet.getDataRange().getValues();
      if (paymentData.length) {
        var ph = paymentData[0];
        var preq = ph.indexOf('Request_ID');
        var pemail = ph.indexOf('Email');
        if (preq !== -1 && pemail !== -1) {
          for (var j = 1; j < paymentData.length; j++) {
            if (String(paymentData[j][preq]) === String(requestId)) {
              paymentSheet.getRange(j + 1, pemail + 1).setValue(email);
              break;
            }
          }
        }
      }
    }

    return {success:true, email:email};
  } catch (error) {
    console.error('updateUpgradeRequestWorkspaceEmail error:', error);
    return {success:false, message:error.message || 'Unable to update workspace email.'};
  }
}
