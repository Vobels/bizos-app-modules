// ============================================================
// MasterSecurity.js - authoritative paid-client authentication
// ============================================================
// Security rule:
//   The Clients registry is authoritative for paid-client identity,
//   business association, and workspace binding. Request-supplied
//   clientId/sheetId/businessId are never used to select a workspace.
// ============================================================

function getActiveClientByEmail_(email) {
  try {
    var normalizedEmail = String(email || '').trim().toLowerCase();
    if (!normalizedEmail) return { success:false, code:'INVALID_EMAIL', message:'Email is required.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Clients');
    if (!sheet || sheet.getLastRow() < 2) {
      return { success:false, code:'CLIENT_NOT_FOUND', message:'No active paid client is registered for this account.' };
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h){ return String(h || '').trim(); });
    var emailCol = headers.indexOf('Email');
    var statusCol = headers.indexOf('Status');
    if (emailCol < 0) return { success:false, code:'CLIENT_REGISTRY_INVALID', message:'Client registry is missing the Email column.' };

    var matches = [];
    for (var i=1; i<data.length; i++) {
      var row = data[i];
      var rowEmail = String(row[emailCol] || '').trim().toLowerCase();
      var status = statusCol >= 0 ? String(row[statusCol] || '').trim().toLowerCase() : '';
      if (rowEmail !== normalizedEmail || status !== 'active') continue;

      var client = {};
      headers.forEach(function(header,index){ client[header] = row[index] === undefined ? '' : row[index]; });
      matches.push(client);
    }

    if (matches.length === 0) return { success:false, code:'CLIENT_NOT_FOUND', message:'No active paid client is registered for this account.' };
    if (matches.length > 1) return { success:false, code:'MULTIPLE_ACTIVE_CLIENTS', message:'Multiple active client deployments are associated with this account. Please contact support.' };

    var c = matches[0];
    var client = {
      clientId:String(c.Client_ID || ''),
      email:String(c.Email || '').trim().toLowerCase(),
      clientName:String(c.Client_Name || ''),
      businessId:String(c.Business_ID || ''),
      sheetId:String(c.Workspace_ID || c.Sheet_ID || ''),
      scriptId:String(c.Script_ID || ''),
      deploymentId:String(c.Deployment_ID || ''),
      webAppUrl:String(c.Web_App_URL || ''),
      landingUrl:String(c.Landing_URL || c.Web_App_URL || ''),
      tier:String(c.Tier || 'sovereign'),
      status:String(c.Status || '').toLowerCase(),
      primaryColor:String(c.Primary_Color || '#5D2A86'),
      logoUrl:String(c.Logo_Url || '')
    };

    if (!client.clientId || !client.email || !client.businessId || !client.sheetId) {
      return { success:false, code:'CLIENT_REGISTRY_INCOMPLETE', message:'The client registry record is incomplete. Please contact support.' };
    }

    return { success:true, client:client };
  } catch (error) {
    console.error('getActiveClientByEmail_ error:', error);
    return { success:false, code:'CLIENT_REGISTRY_ERROR', message:'Unable to verify client registry.' };
  }
}
