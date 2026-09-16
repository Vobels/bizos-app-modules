/**
 * Update the destination email for an existing upgrade request.
 * The authenticated account email remains unchanged.
 */
function updateUpgradeRequestWorkspaceEmail(requestId, workspaceEmail, sessionId) {
  try {
    var email = String(workspaceEmail || '').trim().toLowerCase();
    if (!requestId) return {success:false,message:'Upgrade request not found.'};
    if (!/^\S+@\S+\.\S+$/.test(email)) return {success:false,message:'Please provide a valid workspace email.'};

    var user = validateUpgradeSession(sessionId);
    if (!user) return {success:false,message:'Session expired. Please login again.'};

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var upgradeSheet = ss.getSheetByName('Upgrade_Requests');
    if (!upgradeSheet) return {success:false,message:'Upgrade request could not be found.'};
    var data = upgradeSheet.getDataRange().getValues();
    if (!data.length) return {success:false,message:'Upgrade request could not be found.'};
    var headers = data[0];
    var requestCol = headers.indexOf('Request_ID');
    var accountEmailCol = headers.indexOf('Email');
    var workspaceEmailCol = headers.indexOf('Workspace_Email');
    var updatedCol = headers.indexOf('Updated_At');
    if (requestCol === -1 || accountEmailCol === -1) return {success:false,message:'Upgrade request format is invalid.'};
    if (workspaceEmailCol === -1) {
      workspaceEmailCol = headers.length;
      upgradeSheet.insertColumnAfter(headers.length);
      upgradeSheet.getRange(1, workspaceEmailCol + 1).setValue('Workspace_Email');
    }

    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][requestCol]) === String(requestId)) {
        var ownerEmail = String(data[i][accountEmailCol] || '').trim().toLowerCase();
        if (ownerEmail !== String(user.email || '').trim().toLowerCase()) return {success:false,message:'You do not have permission to update this upgrade request.'};
        foundRow = i + 1;
        break;
      }
    }
    if (foundRow === -1) return {success:false,message:'Upgrade request could not be found.'};

    upgradeSheet.getRange(foundRow, workspaceEmailCol + 1).setValue(email);
    if (updatedCol !== -1) upgradeSheet.getRange(foundRow, updatedCol + 1).setValue(new Date().toISOString());

    var paymentSheet = ss.getSheetByName('Payments');
    if (paymentSheet) {
      var paymentData = paymentSheet.getDataRange().getValues();
      if (paymentData.length) {
        var ph = paymentData[0], preq = ph.indexOf('Request_ID'), pworkspace = ph.indexOf('Workspace_Email');
        if (preq !== -1) {
          if (pworkspace === -1) { pworkspace = ph.length; paymentSheet.insertColumnAfter(ph.length); paymentSheet.getRange(1,pworkspace+1).setValue('Workspace_Email'); }
          for (var j = 1; j < paymentData.length; j++) {
            if (String(paymentData[j][preq]) === String(requestId)) { paymentSheet.getRange(j + 1,pworkspace + 1).setValue(email); break; }
          }
        }
      }
    }
    return {success:true,email:email};
  } catch (error) {
    console.error('updateUpgradeRequestWorkspaceEmail error:',error);
    return {success:false,message:error.message || 'Unable to update workspace email.'};
  }
}
