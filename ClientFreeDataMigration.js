// ============================================================
// ClientFreeDataMigration.js - FREE -> PAID WORKSPACE MIGRATION
// ============================================================
// Migrates an existing user's workspace data into the newly provisioned
// paid client workspace. Finance + Ecommerce are the free-tier modules,
// but the migration is defensive and also carries forward any records
// already present in the other BizOS modules, plus the business's staff.
// ============================================================

function migrateFreeWorkspaceDataToPaidClient(sourceWorkspaceId, targetWorkspaceId, businessId, businessName, ownerEmail) {
  try {
    if (!sourceWorkspaceId || !targetWorkspaceId || String(sourceWorkspaceId) === String(targetWorkspaceId)) {
      return {success:true,migrated:false,reason:'NO_FREE_WORKSPACE'};
    }

    var source = SpreadsheetApp.openById(sourceWorkspaceId);
    var target = SpreadsheetApp.openById(targetWorkspaceId);
    // Finance + Ecommerce are the free-tier modules, but the migration is
    // intentionally defensive: if a legacy/previously-paid workspace contains
    // records in any other BizOS module, those records must not be silently lost.
    var modules = [
      'Finance','Ecommerce','Sales','CRM','HR','Logistics',
      'Tax','Agro','Productivity','POS','Attendance','Warehouse'
    ];
    var results = {};
    var totalCopied = 0;

    modules.forEach(function(moduleName) {
      results[moduleName] = migrateFreeModuleSheet_(source, target, moduleName);
      totalCopied += results[moduleName].copied || 0;
    });

    var staffResult = migrateFreeWorkspaceStaffToPaidClient_(businessId, target, ownerEmail);
    migrateFreeBusinessMetadata_(source, target, businessId, businessName, ownerEmail, staffResult);

    return {
      success:true,
      migrated:true,
      migrationVersion:'2.0',
      totalCopied:totalCopied,
      modules:results,
      staff:staffResult
    };
  } catch (error) {
    console.error('Free-to-paid migration failed:', error);
    return {success:false,code:'FREE_DATA_MIGRATION_FAILED',message:error.message||'Free workspace data could not be migrated.'};
  }
}

function migrateFreeModuleSheet_(source, target, moduleName) {
  var sheetName = moduleName === 'Finance' ? 'Financial_Data' : 'Ecommerce_Data';
  var sourceSheet = source.getSheetByName(sheetName);
  var targetSheet = target.getSheetByName(sheetName);

  if (!sourceSheet) return {copied:0,skipped:0,reason:'SOURCE_SHEET_MISSING'};
  if (!targetSheet) throw new Error('Target sheet missing: '+sheetName);

  var sourceValues = sourceSheet.getDataRange().getValues();
  if (sourceValues.length < 2) return {copied:0,skipped:0,reason:'NO_SOURCE_DATA'};

  var sourceHeaders = sourceValues[0].map(function(h){return String(h||'').trim();});
  var targetHeaders = targetSheet.getRange(1,1,1,targetSheet.getLastColumn()).getValues()[0].map(function(h){return String(h||'').trim();});
  var targetIdHeader = moduleName === 'Finance' ? 'Transaction_ID' : 'Order_ID';
  var targetIdCol = targetHeaders.indexOf(targetIdHeader);

  // Build an index of existing target records so retries cannot duplicate data.
  var existing = {};
  var targetLastRow = targetSheet.getLastRow();
  if (targetLastRow > 1) {
    var targetValues = targetSheet.getRange(2,1,targetLastRow-1,targetSheet.getLastColumn()).getValues();
    targetValues.forEach(function(row) {
      var id = targetIdCol !== -1 ? String(row[targetIdCol]||'').trim() : '';
      if (id) existing['ID:'+id] = true;
      existing['FP:'+migrationRowFingerprint_(targetHeaders,row)] = true;
    });
  }

  var rowsToAppend = [];
  var skipped = 0;
  sourceValues.slice(1).forEach(function(sourceRow, sourceIndex) {
    var targetRow = targetHeaders.map(function(header) {
      var col = sourceHeaders.indexOf(header);
      return col === -1 ? '' : sourceRow[col];
    });

    var sourceIdCol = sourceHeaders.indexOf(targetIdHeader);
    var sourceId = sourceIdCol === -1 ? '' : String(sourceRow[sourceIdCol]||'').trim();

    // Legacy rows without IDs get a deterministic ID based on their stable
    // source location. Retrying the migration therefore produces the same ID,
    // while identical but separate source rows still receive different IDs.
    if (!sourceId && targetIdCol !== -1) {
      var migrationId = generateDeterministicMigrationId_(source.getId(), moduleName, sourceIndex + 2, targetIdHeader);
      if (existing['ID:'+migrationId]) {
        skipped++;
        return;
      }
      targetRow[targetIdCol] = migrationId;
    }

    var fingerprint = migrationRowFingerprint_(targetHeaders,targetRow);
    if ((sourceId && existing['ID:'+sourceId]) || existing['FP:'+fingerprint]) {
      skipped++;
      return;
    }

    rowsToAppend.push(targetRow);
    if (targetIdCol !== -1 && targetRow[targetIdCol]) existing['ID:'+String(targetRow[targetIdCol])] = true;
    existing['FP:'+migrationRowFingerprint_(targetHeaders,targetRow)] = true;
  });

  if (rowsToAppend.length) {
    targetSheet.getRange(targetSheet.getLastRow()+1,1,rowsToAppend.length,targetHeaders.length).setValues(rowsToAppend);
  }

  return {copied:rowsToAppend.length,skipped:skipped,sourceRows:sourceValues.length-1};
}

function generateDeterministicMigrationId_(sourceWorkspaceId, moduleName, sourceRowNumber, idHeader) {
  var seed = [String(sourceWorkspaceId||''),String(moduleName||''),String(sourceRowNumber||'')].join('|');
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed);
  var hex = digest.map(function(byte) {
    var value = byte < 0 ? byte + 256 : byte;
    return ('0' + value.toString(16)).slice(-2);
  }).join('').substring(0,12).toUpperCase();

  return (idHeader === 'Transaction_ID' ? 'TXN-MIGRATED-' : 'ORD-MIGRATED-') + hex;
}

function migrationRowFingerprint_(headers,row) {
  var values = headers.map(function(header,i) {
    if (header === 'Created_At' || header === 'Last_Login') return '';
    return String(row[i] === undefined || row[i] === null ? '' : row[i]);
  });
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,values.join('\u001f'))
  );
}

function migrateFreeWorkspaceStaffToPaidClient_(businessId, target, ownerEmail) {
  try {
    var userSheet = getOrCreateUserSheet();
    if (!userSheet) return {copied:0,updated:0,skipped:0,reason:'MASTER_USERS_SHEET_MISSING'};

    var values = userSheet.getDataRange().getValues();
    if (values.length < 2) return {copied:0,updated:0,skipped:0,reason:'NO_STAFF_DATA'};

    var headers = values[0] || [];
    var emailCol = headers.indexOf('Email');
    var businessIdCol = headers.indexOf('Business_ID');
    var roleCol = headers.indexOf('Role');
    var nameCol = headers.indexOf('Name');
    var phoneCol = headers.indexOf('Phone');
    var departmentCol = headers.indexOf('Department');
    var statusCol = headers.indexOf('Invitation_Status');
    var passwordCol = headers.indexOf('Password_Hash');
    var createdCol = headers.indexOf('Created_At');
    var assignedCol = headers.indexOf('Assigned_Modules');

    if (emailCol < 0 || businessIdCol < 0) {
      return {copied:0,updated:0,skipped:0,reason:'MASTER_USERS_SCHEMA_INCOMPLETE'};
    }

    var team = target.getSheetByName('Client_Team');
    if (!team) {
      team = target.insertSheet('Client_Team');
      team.appendRow([
        'User_ID','Email','Name','Phone','Role','Department','Status',
        'Password_Hash','Must_Change_Password','Invited_By','Created_At',
        'Updated_At','Last_Login','Assigned_Modules'
      ]);
    }

    var teamValues = team.getDataRange().getValues();
    var teamHeaders = teamValues[0] || [];
    var teamEmailCol = teamHeaders.indexOf('Email');
    if (teamEmailCol < 0) throw new Error('Client_Team schema is missing Email.');

    var existingByEmail = {};
    teamValues.slice(1).forEach(function(row, index) {
      var email = String(row[teamEmailCol] || '').trim().toLowerCase();
      if (email) existingByEmail[email] = index + 2;
    });

    var copied = 0, updated = 0, skipped = 0;
    values.slice(1).forEach(function(row, sourceIndex) {
      if (String(row[businessIdCol] || '') !== String(businessId || '')) return;

      var email = String(row[emailCol] || '').trim().toLowerCase();
      var role = String(roleCol >= 0 ? row[roleCol] || 'staff' : 'staff').trim().toLowerCase();
      if (!email || role === 'owner') return;

      var invitationStatus = String(statusCol >= 0 ? row[statusCol] || '' : '').trim().toLowerCase();
      var status = invitationStatus === 'pending' ? 'invited' :
                   invitationStatus === 'suspended' ? 'suspended' : 'active';
      var assigned = [];
      if (assignedCol >= 0 && row[assignedCol]) {
        try {
          assigned = JSON.parse(String(row[assignedCol]));
          if (!Array.isArray(assigned)) assigned = [];
        } catch (ignore) {
          assigned = String(row[assignedCol]).split(',').map(function(x){return String(x || '').trim();}).filter(Boolean);
        }
      }
      if (role === 'admin') assigned = Object.keys(MODULES);

      var now = new Date().toISOString();
      var rowData = {
        User_ID: 'TEAM_MIGRATED_' + Utilities.base64EncodeWebSafe(
          Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(businessId) + '|' + email)
        ).replace(/[^A-Za-z0-9]/g,'').substring(0,12).toUpperCase(),
        Email: email,
        Name: nameCol >= 0 ? row[nameCol] || email : email,
        Phone: phoneCol >= 0 ? row[phoneCol] || '' : '',
        Role: role,
        Department: departmentCol >= 0 ? row[departmentCol] || '' : '',
        Status: status,
        Password_Hash: passwordCol >= 0 ? row[passwordCol] || '' : '',
        Must_Change_Password: passwordCol >= 0 && row[passwordCol] ? 'YES' : 'YES',
        Invited_By: ownerEmail || '',
        Created_At: createdCol >= 0 ? row[createdCol] || now : now,
        Updated_At: now,
        Last_Login: '',
        Assigned_Modules: JSON.stringify(assigned)
      };

      var targetRow = existingByEmail[email];
      if (targetRow) {
        var targetData = team.getRange(targetRow,1,1,team.getLastColumn()).getValues()[0];
        teamHeaders.forEach(function(header, col) {
          if (rowData[header] !== undefined && rowData[header] !== '') {
            // Preserve a live client password if the record was already migrated
            // and subsequently changed by the client.
            if (header === 'Password_Hash' && targetData[col]) return;
            team.getRange(targetRow,col+1).setValue(rowData[header]);
          }
        });
        updated++;
      } else {
        var output = teamHeaders.map(function(header){return rowData[header] === undefined ? '' : rowData[header];});
        team.appendRow(output);
        copied++;
      }
    });

    return {copied:copied,updated:updated,skipped:skipped,sourceRows:values.length-1};
  } catch (error) {
    console.error('Free-to-paid staff migration failed:', error);
    return {success:false,copied:0,updated:0,skipped:0,reason:'STAFF_MIGRATION_FAILED',message:error.message || 'Staff migration failed.'};
  }
}

function migrateFreeBusinessMetadata_(source, target, businessId, businessName, ownerEmail, staffResult) {
  var targetInfo = target.getSheetByName('Client_Info');
  if (!targetInfo) return;

  var metadata = {
    Source_Free_Workspace_ID: source.getId(),
    Business_ID: businessId || '',
    Business_Name: businessName || '',
    Owner_Email: String(ownerEmail || '').toLowerCase(),
    Data_Migrated_At: new Date().toISOString(),
    Data_Migration_Version: '2.0',
    Migrated_Module_Data: 'All BizOS modules (defensive migration)',
    Migrated_Staff_Count: staffResult ? Number(staffResult.copied || 0) + Number(staffResult.updated || 0) : 0
  };

  // Copy useful business profile fields from the free workspace when present.
  var sourceInfo = source.getSheetByName('Client_Info') || source.getSheetByName('Business_Info');
  if (sourceInfo && sourceInfo.getLastRow() > 1) {
    var values = sourceInfo.getDataRange().getValues();
    values.slice(1).forEach(function(row) {
      var key = String(row[0]||'').trim();
      if (key && row.length > 1 && row[1] !== '') metadata[key] = row[1];
    });
  }

  var values = targetInfo.getDataRange().getValues();
  var headers = values[0] || ['Property','Value'];
  if (headers.indexOf('Property') === -1 || headers.indexOf('Value') === -1) return;

  var existing = {};
  values.slice(1).forEach(function(row,i){existing[String(row[0]||'')] = i+2;});
  Object.keys(metadata).forEach(function(key) {
    if (existing[key]) targetInfo.getRange(existing[key],2).setValue(metadata[key]);
    else targetInfo.appendRow([key,metadata[key]]);
  });
}