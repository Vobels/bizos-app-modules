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
      results[moduleName] = migrateFreeModuleSheet_(source, target, moduleName, businessId);
      totalCopied += results[moduleName].copied || 0;
    });

    var staffResult = migrateFreeWorkspaceStaffToPaidClient_(businessId, target, ownerEmail);
    if (staffResult && staffResult.success === false) {
      throw new Error(staffResult.message || 'Staff migration failed.');
    }
    var notificationResult = migrateBizOSNotificationsAndActivityToPaidClient_(businessId, target, ownerEmail);
    if (notificationResult && notificationResult.success === false) {
      throw new Error(notificationResult.message || 'BizOS notifications and activity could not be migrated.');
    }
    migrateFreeBusinessMetadata_(source, target, businessId, businessName, ownerEmail, staffResult);

    return {
      success:true,
      migrated:true,
      migrationVersion:'2.0',
      totalCopied:totalCopied,
      modules:results,
      staff:staffResult,
      notifications:notificationResult
    };
  } catch (error) {
    console.error('Free-to-paid migration failed:', error);
    return {success:false,code:'FREE_DATA_MIGRATION_FAILED',message:error.message||'Free workspace data could not be migrated.'};
  }
}

function migrateFreeModuleSheet_(source, target, moduleName, businessId) {
  var moduleSheetMap = {
    Finance:'Financial_Data', Ecommerce:'Ecommerce_Data', Sales:'Sales_Data', CRM:'CRM_Data',
    HR:'HR_Data', Logistics:'Logistics_Data', Tax:'Tax_Data', Agro:'Agro_Data',
    Productivity:'Productivity_Data', POS:'POS_Data', Attendance:'Attendance_Data', Warehouse:'Warehouse_Data'
  };
  var sheetName = moduleSheetMap[moduleName] || (moduleName + '_Data');
  var sourceSheet = source.getSheetByName(sheetName);
  var targetSheet = target.getSheetByName(sheetName);

  if (!sourceSheet) return {copied:0,skipped:0,reason:'SOURCE_SHEET_MISSING'};
  if (!targetSheet) throw new Error('Target sheet missing: '+sheetName);

  var sourceValues = sourceSheet.getDataRange().getValues();
  if (sourceValues.length < 2) return {copied:0,skipped:0,reason:'NO_SOURCE_DATA'};

  var sourceHeaders = sourceValues[0].map(function(h){return String(h||'').trim();});
  var targetHeaders = targetSheet.getRange(1,1,1,targetSheet.getLastColumn()).getValues()[0].map(function(h){return String(h||'').trim();});
  var targetIdHeaderMap = {
    Finance:'Transaction_ID', Ecommerce:'Order_ID', Sales:'Deal_ID', CRM:'Contact_ID',
    HR:'Employee_ID', Logistics:'Shipment_ID', Tax:'Tax_ID', Agro:'Activity_ID',
    Productivity:'Task_ID', POS:'Sale_ID', Attendance:'Attendance_ID', Warehouse:'Item_ID'
  };
  var targetIdHeader = targetIdHeaderMap[moduleName] || '';
  var targetIdCol = targetHeaders.indexOf(targetIdHeader);
  var sourceBusinessIdCol = sourceHeaders.indexOf('Business_ID');

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
  var migrationEntries = [];
  var skipped = 0;
  sourceValues.slice(1).forEach(function(sourceRow, sourceIndex) {
    // A workspace should normally contain only one business, but if a legacy
    // sheet ever contains Business_ID values, enforce the exact business scope.
    if (sourceBusinessIdCol >= 0 && String(sourceRow[sourceBusinessIdCol] || '') !== String(businessId || '')) return;

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
    migrationEntries.push({
      moduleName:moduleName,
      recordId:targetIdCol !== -1 ? String(targetRow[targetIdCol] || '') : '',
      sourceRow:sourceIndex + 2,
      headers:targetHeaders.slice(),
      row:targetRow.slice()
    });
    if (targetIdCol !== -1 && targetRow[targetIdCol]) existing['ID:'+String(targetRow[targetIdCol])] = true;
    existing['FP:'+migrationRowFingerprint_(targetHeaders,targetRow)] = true;
  });

  if (rowsToAppend.length) {
    targetSheet.getRange(targetSheet.getLastRow()+1,1,rowsToAppend.length,targetHeaders.length).setValues(rowsToAppend);
    recordClientMigrationEntries_(target, source.getId(), migrationEntries);
  }

  return {copied:rowsToAppend.length,skipped:skipped,sourceRows:sourceValues.length-1};
}

function ensureClientMigrationIndex_(target) {
  var s=target.getSheetByName('Client_Migration_Index');
  if(!s){
    s=target.insertSheet('Client_Migration_Index');
    s.appendRow(['Migration_ID','Module','Record_ID','Source_Workspace_ID','Source_Row','Migrated_At','Status','Archived_At','Record_Headers_JSON','Record_Row_JSON']);
  }
  return s;
}

function recordClientMigrationEntries_(target, sourceWorkspaceId, entries) {
  if(!entries || !entries.length)return;
  var s=ensureClientMigrationIndex_(target), now=new Date().toISOString();
  var rows=entries.map(function(e){
    var migrationId='MIG_'+Utilities.getUuid().replace(/-/g,'').substring(0,16).toUpperCase();
    return [migrationId,e.moduleName,e.recordId,String(sourceWorkspaceId||''),e.sourceRow,now,'active','',JSON.stringify(e.headers||[]),JSON.stringify(e.row||[])];
  });
  s.getRange(s.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
}

function generateDeterministicMigrationId_(sourceWorkspaceId, moduleName, sourceRowNumber, idHeader) {
  var seed = [String(sourceWorkspaceId||''),String(moduleName||''),String(sourceRowNumber||'')].join('|');
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed);
  var hex = digest.map(function(byte) {
    var value = byte < 0 ? byte + 256 : byte;
    return ('0' + value.toString(16)).slice(-2);
  }).join('').substring(0,12).toUpperCase();

  var prefixMap = {
    Transaction_ID:'TXN', Order_ID:'ORD', Deal_ID:'DEAL', Contact_ID:'CRM', Employee_ID:'EMP',
    Shipment_ID:'SHIP', Tax_ID:'TAX', Activity_ID:'ACT', Task_ID:'TASK', Sale_ID:'SALE',
    Attendance_ID:'ATT', Item_ID:'ITEM'
  };
  return (prefixMap[idHeader] || 'REC') + '-MIGRATED-' + hex;
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

function migrateBizOSNotificationsAndActivityToPaidClient_(businessId, target, ownerEmail) {
  var result = {success:true,notificationsCopied:0,activityCopied:0,notificationsSkipped:0,activitySkipped:0};
  try {
    var master = getBizOSMasterSpreadsheet_();
    var notificationSheet = master.getSheetByName('BizOS_Notifications');
    var activitySheet = master.getSheetByName('BizOS_Activity');
    var targetNotifications = ensureClientNotificationsSheetV12_(target);
    var targetActivity = ensureClientActivitySheetV12_(target);

    if (notificationSheet && notificationSheet.getLastRow() > 1) {
      var nv=notificationSheet.getDataRange().getValues(), nh=nv[0]||[], ni={};
      nh.forEach(function(h,i){ni[h]=i;});
      var tv=targetNotifications.getDataRange().getValues(), th=tv[0]||[], ti={};
      th.forEach(function(h,i){ti[h]=i;});
      var existing={};
      for(var r=1;r<tv.length;r++) existing[String(tv[r][ti.Notification_ID]||'')]=true;
      for(var i=1;i<nv.length;i++){
        if(String(nv[i][ni.Business_ID]||'')!==String(businessId||''))continue;
        var id=String(nv[i][ni.Notification_ID]||'');
        if(!id||existing[id]){result.notificationsSkipped++;continue;}
        var row=new Array(th.length).fill('');
        row[ti.Notification_ID]=id;
        row[ti.Audience]=String(nv[i][ni.Audience]||'all');
        row[ti.Type]=String(nv[i][ni.Type]||'system');
        row[ti.Title]=String(nv[i][ni.Title]||'');
        row[ti.Message]=String(nv[i][ni.Message]||'');
        row[ti.Created_At]=nv[i][ni.Created_At]||'';
        row[ti.Created_By]=String(nv[i][ni.Created_By]||'BizOS');
        row[ti.Related_View]=String(nv[i][ni.Related_View]||'');
        row[ti.Read_By_JSON]=String(nv[i][ni.Read_By_JSON]||'[]');
        targetNotifications.appendRow(row);
        existing[id]=true;
        result.notificationsCopied++;
      }
    }

    if (activitySheet && activitySheet.getLastRow() > 1) {
      var av=activitySheet.getDataRange().getValues(), ah=av[0]||[], ai={};
      ah.forEach(function(h,i){ai[h]=i;});
      var tav=targetActivity.getDataRange().getValues(), tah=tav[0]||[], tai={};
      tah.forEach(function(h,i){tai[h]=i;});
      var existingActivity={};
      for(var a=1;a<tav.length;a++) existingActivity[String(tav[a][tai.Activity_ID]||'')]=true;
      for(var j=1;j<av.length;j++){
        if(String(av[j][ai.Business_ID]||'')!==String(businessId||''))continue;
        var activityId=String(av[j][ai.Activity_ID]||'');
        if(!activityId||existingActivity[activityId]){result.activitySkipped++;continue;}
        var activityRow=new Array(tah.length).fill('');
        activityRow[tai.Activity_ID]=activityId;
        activityRow[tai.Actor_Email]=String(av[j][ai.Actor_Email]||'');
        activityRow[tai.Actor_Name]=String(av[j][ai.Actor_Name]||'');
        activityRow[tai.Event_Type]=String(av[j][ai.Event_Type]||'activity');
        activityRow[tai.Summary]=String(av[j][ai.Summary]||'');
        activityRow[tai.Created_At]=av[j][ai.Created_At]||'';
        activityRow[tai.Related_View]=String(av[j][ai.Related_View]||'');
        targetActivity.appendRow(activityRow);
        existingActivity[activityId]=true;
        result.activityCopied++;
      }
    }
    return result;
  } catch(error) {
    console.error('migrateBizOSNotificationsAndActivityToPaidClient_ error:',error);
    return {success:false,code:'NOTIFICATION_ACTIVITY_MIGRATION_FAILED',message:error.message,notificationsCopied:0,activityCopied:0};
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