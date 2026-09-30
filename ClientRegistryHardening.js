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

function getActiveClientByEmail_(email) {
  try {
    var wantedEmail = String(email || '').trim().toLowerCase();
    if (!wantedEmail) {
      return {
        success: false,
        code: 'INVALID_EMAIL',
        message: 'Email is required.'
      };
    }

    var ss = getBizOSMasterSpreadsheet_();
    var sheet = ss.getSheetByName('Clients');
    if (!sheet || sheet.getLastRow() < 2) {
      return {
        success: false,
        code: 'CLIENT_REGISTRY_UNAVAILABLE',
        message: 'The client registry is unavailable.'
      };
    }

    var values = sheet.getDataRange().getValues();
    var headers = values[0].map(function (h) {
      return String(h || '').trim();
    });

    var emailCol = headers.indexOf('Email');
    var statusCol = headers.indexOf('Status');

    if (emailCol < 0) {
      return {
        success: false,
        code: 'CLIENT_REGISTRY_SCHEMA_INVALID',
        message: 'The client registry is missing the Email column.'
      };
    }

    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var rowEmail = String(row[emailCol] || '').trim().toLowerCase();
      var status = statusCol >= 0
        ? String(row[statusCol] || '').trim().toLowerCase()
        : '';

      if (rowEmail !== wantedEmail || status !== 'active') continue;

      var client = {};
      headers.forEach(function (header, index) {
        client[header] = row[index] === undefined ? '' : row[index];
      });

      return {
        success: true,
        client: {
          clientId: String(client.Client_ID || ''),
          email: String(client.Email || ''),
          clientName: String(client.Client_Name || ''),
          businessId: String(client.Business_ID || ''),
          sheetId: String(client.Workspace_ID || client.Sheet_ID || ''),
          scriptId: String(client.Script_ID || ''),
          deploymentId: String(client.Deployment_ID || ''),
          webAppUrl: String(client.Web_App_URL || ''),
          tier: String(client.Tier || ''),
          status: String(client.Status || ''),
          primaryColor: String(client.Primary_Color || ''),
          logoUrl: String(client.Logo_Url || ''),
          apiKey: String(client.API_Key || '')
        }
      };
    }

    return {
      success: false,
      code: 'CLIENT_NOT_FOUND',
      message: 'No active paid client is registered for this account.'
    };
  } catch (error) {
    console.error('getActiveClientByEmail_ error:', error);
    return {
      success: false,
      code: 'CLIENT_LOOKUP_ERROR',
      message: error && error.message ? error.message : String(error)
    };
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


// ============================================================
// 🔐 CLIENT LICENSE VALIDATION - V10 RUNTIME BINDING
// ============================================================
// Master-authoritative validation for generated client deployments.
// The caller cannot choose a workspace by supplying arbitrary IDs:
// every identity supplied by the client runtime must match one active
// Clients registry row, including the actual Apps Script project and
// deployment that are executing the request.
function validateClientLicense(payload) {
  try {
    payload = payload || {};

    var clientId = String(payload.clientId || '').trim();
    var businessId = String(payload.businessId || '').trim();
    var sheetId = String(payload.sheetId || '').trim();
    var scriptId = String(payload.scriptId || '').trim();
    var deploymentId = String(payload.deploymentId || '').trim();

    if (!clientId || !businessId || !sheetId || !scriptId || !deploymentId) {
      return {
        success: false,
        code: 'LICENSE_BINDING_INCOMPLETE',
        message: 'Client deployment binding is incomplete.'
      };
    }

    var ss = getBizOSMasterSpreadsheet_();
    var sheet = ss.getSheetByName('Clients');
    if (!sheet || sheet.getLastRow() < 2) {
      return {
        success: false,
        code: 'CLIENT_REGISTRY_UNAVAILABLE',
        message: 'The client registry is unavailable.'
      };
    }

    var values = sheet.getDataRange().getValues();
    var headers = values[0].map(function (h) { return String(h || '').trim(); });
    var required = ['Client_ID','Business_ID','Sheet_ID','Workspace_ID','Status','Script_ID','Deployment_ID','Web_App_URL'];
    for (var r = 0; r < required.length; r++) {
      if (headers.indexOf(required[r]) < 0) {
        return {
          success: false,
          code: 'CLIENT_REGISTRY_SCHEMA_INVALID',
          message: 'The client registry is missing required deployment metadata.'
        };
      }
    }

    var idx = {};
    required.forEach(function (name) { idx[name] = headers.indexOf(name); });

    var matched = null;
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (String(row[idx.Client_ID] || '').trim() !== clientId) continue;

      matched = {
        clientId: String(row[idx.Client_ID] || '').trim(),
        businessId: String(row[idx.Business_ID] || '').trim(),
        sheetId: String(row[idx.Sheet_ID] || row[idx.Workspace_ID] || '').trim(),
        workspaceId: String(row[idx.Workspace_ID] || '').trim(),
        status: String(row[idx.Status] || '').trim().toLowerCase(),
        scriptId: String(row[idx.Script_ID] || '').trim(),
        deploymentId: String(row[idx.Deployment_ID] || '').trim(),
        webAppUrl: String(row[idx.Web_App_URL] || '').trim()
      };
      break;
    }

    if (!matched) {
      return {
        success: false,
        code: 'CLIENT_NOT_FOUND',
        message: 'This BizOS client is not registered.'
      };
    }

    if (matched.status !== 'active') {
      return {
        success: false,
        code: 'CLIENT_INACTIVE',
        message: 'This BizOS client workspace is inactive.'
      };
    }

    // All runtime bindings are mandatory and must match the same registry row.
    if (matched.businessId !== businessId ||
        matched.sheetId !== sheetId ||
        !matched.workspaceId ||
        matched.workspaceId !== sheetId ||
        matched.scriptId !== scriptId ||
        matched.deploymentId !== deploymentId) {
      console.warn('Client deployment binding mismatch for ' + clientId);
      return {
        success: false,
        code: 'CLIENT_DEPLOYMENT_MISMATCH',
        message: 'This BizOS deployment is not authorized for this client.'
      };
    }

    // The registry URL is another persisted representation of the same
    // deployment. Require it to point at the registered deployment ID.
    var expectedUrl = 'https://script.google.com/macros/s/' + deploymentId + '/exec';
    if (matched.webAppUrl !== expectedUrl) {
      return {
        success: false,
        code: 'CLIENT_URL_BINDING_MISMATCH',
        message: 'The registered BizOS deployment URL does not match the runtime deployment.'
      };
    }

    return {
      success: true,
      code: 'LICENSE_VALID',
      clientId: matched.clientId,
      businessId: matched.businessId,
      sheetId: matched.sheetId,
      scriptId: matched.scriptId,
      deploymentId: matched.deploymentId
    };
  } catch (error) {
    console.error('validateClientLicense error:', error);
    return {
      success: false,
      code: 'LICENSE_VALIDATION_ERROR',
      message: error && error.message ? error.message : 'Client license validation failed.'
    };
  }
}
