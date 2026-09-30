// ============================================================
// ClientRegistryHardening.js
// Client registry + provisioning idempotency helpers.
// ============================================================

function getExistingActiveClientDeployment_(email, businessName) {
  try {
    var ss = getBizOSMasterSpreadsheet_();
    var sheet = ss.getSheetByName('Clients');
    if (!sheet || sheet.getLastRow() < 2) return null;

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function (h) { return String(h || '').trim(); });
    var emailCol = headers.indexOf('Email');
    var nameCol = headers.indexOf('Client_Name');
    var statusCol = headers.indexOf('Status');
    if (emailCol < 0) return null;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowEmail = String(row[emailCol] || '').trim().toLowerCase();
      var rowName = nameCol >= 0 ? String(row[nameCol] || '').trim() : '';
      var status = statusCol >= 0 ? String(row[statusCol] || '').trim().toLowerCase() : '';

      if (rowEmail === String(email || '').trim().toLowerCase() &&
          (!businessName || rowName === String(businessName).trim()) &&
          status === 'active') {
        var client = {};
        headers.forEach(function (header, index) {
          client[header] = row[index] === undefined ? '' : row[index];
        });
        return client;
      }
    }
    return null;
  } catch (error) {
    console.error('getExistingActiveClientDeployment_ error:', error);
    return null;
  }
}

function normalizeExistingClientDeployment_(client) {
  if (!client) return null;
  return {
    success: true,
    clientId: String(client.Client_ID || client.clientId || ''),
    email: String(client.Email || client.email || ''),
    clientName: String(client.Client_Name || client.clientName || ''),
    businessId: String(client.Business_ID || client.businessId || ''),
    sheetId: String(client.Workspace_ID || client.Sheet_ID || client.sheetId || ''),
    scriptId: String(client.Script_ID || client.scriptId || ''),
    deploymentId: String(client.Deployment_ID || client.deploymentId || ''),
    webAppUrl: String(client.Web_App_URL || client.webAppUrl || ''),
    landingUrl: String(client.Landing_URL || client.landingUrl || client.Web_App_URL || ''),
    tier: String(client.Tier || client.tier || 'sovereign'),
    status: String(client.Status || client.status || '')
  };
}

function ensureClientRegistryColumns_() {
  var ss = getBizOSMasterSpreadsheet_();
  var sheet = ss.getSheetByName('Clients');
  if (!sheet) {
    sheet = ss.insertSheet('Clients');
    sheet.appendRow([
      'Client_ID','Email','Client_Name','Business_ID','Domain','Primary_Color',
      'Logo_Url','Custom_Domain','Tier','Status','Sheet_ID','Workspace_ID',
      'Web_App_URL','Landing_URL','API_Key','Created_At','Updated_At',
      'Script_ID','Deployment_ID','Provisioning_Version'
    ]);
    return sheet;
  }

  var headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0]
    .map(function (h) { return String(h || '').trim(); });
  var required = [
    'Client_ID','Email','Client_Name','Business_ID','Domain','Primary_Color',
    'Logo_Url','Custom_Domain','Tier','Status','Sheet_ID','Workspace_ID',
    'Web_App_URL','Landing_URL','API_Key','Created_At','Updated_At',
    'Script_ID','Deployment_ID','Provisioning_Version'
  ];

  required.forEach(function (header) {
    if (headers.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      headers.push(header);
    }
  });
  return sheet;
}

function saveClientRecordV2(settings) {
  try {
    settings = settings || {};
    var sheet = ensureClientRegistryColumns_();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(function (h) { return String(h || '').trim(); });
    var data = sheet.getDataRange().getValues();
    var clientIdCol = headers.indexOf('Client_ID');
    var existingRow = -1;

    for (var i = 1; i < data.length; i++) {
      if (clientIdCol >= 0 && String(data[i][clientIdCol] || '') === String(settings.clientId || '')) {
        existingRow = i + 1;
        break;
      }
    }

    var now = new Date().toISOString();
    var rowData = headers.map(function (header) {
      switch (header) {
        case 'Client_ID': return settings.clientId || '';
        case 'Email': return settings.email || '';
        case 'Client_Name': return settings.clientName || '';
        case 'Business_ID': return settings.businessId || '';
        case 'Domain': return settings.domain || settings.customDomain || '';
        case 'Primary_Color': return settings.primaryColor || '#2E7D32';
        case 'Logo_Url': return settings.logoUrl || '';
        case 'Custom_Domain': return settings.customDomain || '';
        case 'Tier': return settings.tier || 'sovereign';
        case 'Status': return settings.status || 'active';
        case 'Sheet_ID': return settings.sheetId || '';
        case 'Workspace_ID': return settings.workspaceId || settings.sheetId || '';
        case 'Web_App_URL': return settings.webAppUrl || settings.landingUrl || '';
        case 'Landing_URL': return settings.landingUrl || settings.webAppUrl || '';
        case 'API_Key': return settings.apiKey || Utilities.getUuid();
        case 'Created_At': return existingRow > 0 ? data[existingRow - 1][headers.indexOf('Created_At')] : now;
        case 'Updated_At': return now;
        case 'Script_ID': return settings.scriptId || '';
        case 'Deployment_ID': return settings.deploymentId || '';
        case 'Provisioning_Version': return settings.provisioningVersion || '4.1';
        default: return '';
      }
    });

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, rowData.length).setValues([rowData]);
    }

    return { success: true, message: 'Client saved successfully', row: existingRow > 0 ? existingRow : sheet.getLastRow() };
  } catch (error) {
    console.error('saveClientRecordV2 error:', error);
    return { success: false, message: error && error.message ? error.message : String(error) };
  }
}

function returnExistingClientDeployment_(client) {
  var existing = normalizeExistingClientDeployment_(client);
  if (!existing || !existing.clientId) return null;

  if (!existing.sheetId || !existing.scriptId || !existing.deploymentId || !existing.webAppUrl) {
    return {
      success: false,
      code: 'EXISTING_CLIENT_INCOMPLETE',
      message: 'An active client record already exists, but its deployment metadata is incomplete. No second deployment was created.',
      clientId: existing.clientId,
      deploymentId: existing.deploymentId || null,
      scriptId: existing.scriptId || null,
      sheetId: existing.sheetId || null,
      webAppUrl: existing.webAppUrl || null
    };
  }

  return {
    success: true,
    idempotent: true,
    existing: true,
    clientId: existing.clientId,
    email: existing.email,
    clientName: existing.clientName,
    businessId: existing.businessId,
    sheetId: existing.sheetId,
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(existing.sheetId) + '/edit',
    scriptId: existing.scriptId,
    deploymentId: existing.deploymentId,
    landingUrl: existing.landingUrl || existing.webAppUrl,
    webAppUrl: existing.webAppUrl,
    tier: existing.tier,
    message: 'Existing active client deployment returned; no duplicate deployment created.'
  };
}
