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

    var businessCenterResult = migrateFreeBusinessCenterDataToPaidClient_(source, target);
    totalCopied += Number(businessCenterResult.copied || 0);

    var staffResult = migrateFreeWorkspaceStaffToPaidClient_(businessId, target, ownerEmail);
    if (staffResult && staffResult.success === false) {
      throw new Error(staffResult.message || 'Staff migration failed.');
    }
    var notificationResult = migrateBizOSNotificationsAndActivityToPaidClient_(businessId, target, ownerEmail);
    if (notificationResult && notificationResult.success === false) {
      throw new Error(notificationResult.message || 'BizOS notifications and activity could not be migrated.');
    }
    var preferenceResult = migrateBizOSNotificationPreferencesToPaidClient_(businessId, target, ownerEmail);
    if (preferenceResult && preferenceResult.success === false) {
      throw new Error(preferenceResult.message || 'Notification preferences could not be migrated.');
    }
    migrateFreeBusinessMetadata_(source, target, businessId, businessName, ownerEmail, staffResult);

    // Create a BizOS-owned recovery snapshot only after the paid workspace
    // contains the migrated data. This is recovery/backup storage, not the
    // live client workspace.
    var recoveryBackup = createClientRecoveryBackup_(target, businessName, businessId);
    if (!recoveryBackup || !recoveryBackup.success) {
      throw new Error(recoveryBackup && recoveryBackup.message
        ? recoveryBackup.message
        : 'The BizOS recovery backup could not be created.');
    }

    return {
      success:true,
      migrated:true,
      migrationVersion:'2.1',
      totalCopied:totalCopied,
      modules:results,
      staff:staffResult,
      notifications:notificationResult,
      businessCenter:businessCenterResult,
      recoveryBackup:recoveryBackup
    };
  } catch (error) {
    console.error('Free-to-paid migration failed:', error);
    return {success:false,code:'FREE_DATA_MIGRATION_FAILED',message:error.message||'Free workspace data could not be migrated.'};
  }
}

function createClientRecoveryBackup_(target, businessName, businessId) {
  try {
    var sourceFile = DriveApp.getFileById(target.getId());
    var recoveryRoot = null;
    var folders = DriveApp.getFoldersByName('BizOS_Client_Recovery');
    recoveryRoot = folders.hasNext() ? folders.next() : DriveApp.createFolder('BizOS_Client_Recovery');

    var safeName = String(businessName || businessId || 'Client').replace(/[^a-zA-Z0-9]/g, '_');
    var clientFolders = recoveryRoot.getFoldersByName(safeName + '_' + String(businessId || '').replace(/[^a-zA-Z0-9]/g, '_'));
    var clientFolder = clientFolders.hasNext()
      ? clientFolders.next()
      : recoveryRoot.createFolder(safeName + '_' + String(businessId || '').replace(/[^a-zA-Z0-9]/g, '_'));

    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Africa/Lagos', 'yyyyMMdd_HHmmss');
    var backupFile = sourceFile.makeCopy('BizOS Recovery - ' + safeName + ' - ' + stamp, clientFolder);

    // Business Center product images are external Drive files referenced by
    // Image_File_ID, so preserve image copies alongside the spreadsheet backup.
    var copiedImages = 0;
    try {
      var parents = sourceFile.getParents();
      if (parents.hasNext()) {
        var workspaceFolder = parents.next();
        var files = workspaceFolder.getFiles();
        while (files.hasNext()) {
          var file = files.next();
          if (file.getId() === sourceFile.getId()) continue;
          var mime = String(file.getMimeType() || '').toLowerCase();
          if (mime.indexOf('image/') !== 0) continue;
          file.makeCopy('Recovery - ' + file.getName(), clientFolder);
          copiedImages++;
        }
      }
    } catch (imageError) {
      console.error('Recovery image backup warning:', imageError);
    }

    return {
      success:true,
      backupId:backupFile.getId(),
      backupUrl:backupFile.getUrl(),
      recoveryFolderId:clientFolder.getId(),
      recoveryFolderUrl:clientFolder.getUrl(),
      copiedImages:copiedImages,
      createdAt:new Date().toISOString()
    };
  } catch (error) {
    console.error('createClientRecoveryBackup_ error:', error);
    return {
      success:false,
      code:'RECOVERY_BACKUP_FAILED',
      message:'BizOS could not create the client recovery backup: ' + (error.message || error)
    };
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
    HR:'Employee_ID', Logistics:'Shipment_ID', Tax:'Tax_ID', Agro:'Record_ID',
    Productivity:'Record_ID', POS:'Transaction_ID', Attendance:'Record_ID', Warehouse:'Item_ID'
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

  // Translate known legacy source headers into the canonical paid schema.
  // This is intentionally limited to documented legacy Finance names.
  var sourceHeaderAliases = {
    Finance: {
      Transaction_ID: ['Transaction_ID', 'Transaction ID', 'Reference'],
      Type: ['Type', 'Type(Income/Expenses)', 'Type (Income/Expenses)', 'Income/Expenses'],
      Payment_Method: ['Payment_Method', 'Payment Method'],
      Created_By: ['Created_By', 'Created By'],
      Created_At: ['Created_At', 'Created At']
    }
  };

  function sourceColumnForTargetHeader_(header) {
    var direct = sourceHeaders.indexOf(header);
    if (direct >= 0) return direct;
    var aliases = sourceHeaderAliases[moduleName] && sourceHeaderAliases[moduleName][header];
    if (!aliases) return -1;
    for (var ai = 0; ai < aliases.length; ai++) {
      var aliasIndex = sourceHeaders.indexOf(aliases[ai]);
      if (aliasIndex >= 0) return aliasIndex;
    }
    return -1;
  }

  sourceValues.slice(1).forEach(function(sourceRow, sourceIndex) {
    // A workspace should normally contain only one business, but if a legacy
    // sheet ever contains Business_ID values, enforce the exact business scope.
    if (sourceBusinessIdCol >= 0 && String(sourceRow[sourceBusinessIdCol] || '') !== String(businessId || '')) return;

    var targetRow = targetHeaders.map(function(header) {
      var col = sourceColumnForTargetHeader_(header);
      return col === -1 ? '' : sourceRow[col];
    });

    if (moduleName === 'Finance') {
      var typeCol = targetHeaders.indexOf('Type');
      if (typeCol >= 0) {
        var legacyType = String(targetRow[typeCol] || '').trim().toLowerCase();
        if (legacyType === 'income' || legacyType === 'revenue') targetRow[typeCol] = 'revenue';
        else if (legacyType === 'expenses' || legacyType === 'expense') targetRow[typeCol] = 'expense';
      }
    }

    var sourceIdCol = sourceColumnForTargetHeader_(targetIdHeader);
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

function migrateFreeBusinessCenterDataToPaidClient_(source, target) {
  var sheetNames = [
    'BusinessCenter_Products',
    'BusinessCenter_Customers',
    'BusinessCenter_Sales',
    'BusinessCenter_Sale_Items',
    'BusinessCenter_Stock_Movements',
    'BusinessCenter_Customer_Ledger',
    'BusinessCenter_Finance_Transactions'
  ];
  var results = {success:true,copied:0,skipped:0,sheets:{}};
  var targetFile = DriveApp.getFileById(target.getId());
  var targetFolder = null;
  try { targetFolder = targetFile.getParents().hasNext() ? targetFile.getParents().next() : null; } catch(ignore) {}

  sheetNames.forEach(function(name) {
    var sourceSheet = source.getSheetByName(name);
    if (!sourceSheet || sourceSheet.getLastRow() < 1) {
      results.sheets[name] = {copied:0,skipped:0,reason:'SOURCE_SHEET_MISSING'};
      return;
    }
    var sourceValues = sourceSheet.getDataRange().getValues();
    var headers = sourceValues[0] || [];
    var targetSheet = target.getSheetByName(name);
    if (!targetSheet) {
      targetSheet = target.insertSheet(name);
      targetSheet.getRange(1,1,1,headers.length).setValues([headers]);
    }

    var targetValues = targetSheet.getDataRange().getValues();
    var targetHeaders = targetValues[0] || headers;
    var idHeaders = ['Product_ID','Customer_ID','Sale_ID','Movement_ID','Entry_ID','Transaction_ID'];
    var idHeader = idHeaders.find(function(h){return headers.indexOf(h)>=0 && targetHeaders.indexOf(h)>=0;}) || '';
    var targetIdCol = idHeader ? targetHeaders.indexOf(idHeader) : -1;
    var existing = {};
    if (targetValues.length > 1 && targetIdCol >= 0) {
      targetValues.slice(1).forEach(function(row){ var id=String(row[targetIdCol]||'').trim(); if(id)existing[id]=true; });
    }

    var copied=0, skipped=0, rows=[];
    sourceValues.slice(1).forEach(function(sourceRow) {
      var sourceBusinessCol=headers.indexOf('Business_ID');
      if(sourceBusinessCol>=0 && String(sourceRow[sourceBusinessCol]||'')!==String(source.getId()||'') && String(sourceRow[sourceBusinessCol]||'')!=='') {
        // Business Center sheets normally do not contain Business_ID. If they do,
        // the source workspace/business boundary is still enforced by the workspace itself.
      }
      var row=targetHeaders.map(function(h){var i=headers.indexOf(h);return i>=0?sourceRow[i]:'';});
      var id=targetIdCol>=0?String(row[targetIdCol]||'').trim():'';
      if(id && existing[id]){skipped++;return;}
      rows.push(row); if(id)existing[id]=true; copied++;
    });
    if(rows.length)targetSheet.getRange(targetSheet.getLastRow()+1,1,rows.length,targetHeaders.length).setValues(rows);
    results.sheets[name]={copied:copied,skipped:skipped};
    results.copied+=copied; results.skipped+=skipped;
  });

  // Product images live in Drive rather than in the sheet. Copy referenced
  // files into the paid workspace folder when accessible and rewrite the IDs.
  var srcProducts=source.getSheetByName('BusinessCenter_Products');
  var dstProducts=target.getSheetByName('BusinessCenter_Products');
  if(srcProducts && dstProducts && targetFolder) {
    var sv=srcProducts.getDataRange().getValues(), sh=sv[0]||[], imageCol=sh.indexOf('Image_File_ID'), productCol=sh.indexOf('Product_ID');
    var dv=dstProducts.getDataRange().getValues(), dh=dv[0]||[], dImage=dh.indexOf('Image_File_ID'), dProduct=dh.indexOf('Product_ID');
    if(imageCol>=0 && productCol>=0 && dImage>=0 && dProduct>=0) {
      for(var r=1;r<sv.length;r++){
        var oldId=String(sv[r][imageCol]||'').trim(), pid=String(sv[r][productCol]||'').trim();
        if(!oldId||!pid)continue;
        var targetRow=-1; for(var dr=1;dr<dv.length;dr++){if(String(dv[dr][dProduct]||'')===pid){targetRow=dr+1;break;}}
        if(targetRow<0)continue;
        try {
          var copiedFile=DriveApp.getFileById(oldId).makeCopy(DriveApp.getFileById(oldId).getName(),targetFolder);
          dstProducts.getRange(targetRow,dImage+1).setValue(copiedFile.getId());
        } catch(imageError) {
          console.warn('Business Center image migration skipped for '+pid+': '+imageError.message);
        }
      }
    }
  }
  return results;
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
    Shipment_ID:'SHIP', Tax_ID:'TAX', Record_ID:'REC', Item_ID:'ITEM'
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

    // Preserve the same legacy module-assignment fallback used by the master
    // Staff runtime. Older Users sheets may not have Assigned_Modules yet and
    // keep staff access in Staff_Modules instead.
    var legacyStaffModules = {};
    if (assignedCol < 0) {
      var legacySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Staff_Modules');
      if (legacySheet && legacySheet.getLastRow() > 1) {
        var legacyValues = legacySheet.getDataRange().getValues();
        var legacyHeaders = legacyValues[0] || [];
        var legacyEmailCol = legacyHeaders.indexOf('Staff_Email');
        var legacyBusinessCol = legacyHeaders.indexOf('Business_ID');
        var legacyModuleCol = legacyHeaders.indexOf('Module_Access');
        if (legacyEmailCol >= 0 && legacyBusinessCol >= 0 && legacyModuleCol >= 0) {
          legacyValues.slice(1).forEach(function(legacyRow) {
            var legacyEmail = String(legacyRow[legacyEmailCol] || '').trim().toLowerCase();
            var legacyBusiness = String(legacyRow[legacyBusinessCol] || '');
            var legacyModule = String(legacyRow[legacyModuleCol] || '').trim();
            if (!legacyEmail || legacyBusiness !== String(businessId || '') || !legacyModule) return;
            legacyStaffModules[legacyEmail] = legacyStaffModules[legacyEmail] || [];
            if (legacyStaffModules[legacyEmail].indexOf(legacyModule) === -1) {
              legacyStaffModules[legacyEmail].push(legacyModule);
            }
          });
        }
      }
    }

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
      sourceRows++;

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
      } else if (legacyStaffModules[email]) {
        assigned = legacyStaffModules[email].slice();
      }
      var allowedModules = Object.keys(MODULES);
      assigned = assigned.map(function(moduleName) {
        return String(moduleName || '').trim();
      }).filter(function(moduleName, index, list) {
        return moduleName && allowedModules.indexOf(moduleName) >= 0 && list.indexOf(moduleName) === index;
      });
      if (role === 'admin') assigned = allowedModules.slice();

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
            if (
              (header === 'Password_Hash' ||
               header === 'Status' ||
               header === 'Assigned_Modules' ||
               header === 'Last_Login') &&
              targetData[col]
            ) return;
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

    return {copied:copied,updated:updated,skipped:skipped,sourceRows:sourceRows};
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

function ensureClientNotificationPreferencesSheetV12_(target) {
  var s=target.getSheetByName('Client_Notification_Preferences');
  if(!s){s=target.insertSheet('Client_Notification_Preferences');s.appendRow(['Preference_ID','Business_ID','Email','Feature_Updates','System_Alerts','Weekly_Reports','Security_Alerts','Marketing_Updates','Updated_At']);}
  return s;
}

function migrateBizOSNotificationPreferencesToPaidClient_(businessId, target, ownerEmail) {
  try {
    var source=getBizOSMasterSpreadsheet_(), sourceSheet=source.getSheetByName('BizOS_Notification_Preferences');
    if(!sourceSheet || sourceSheet.getLastRow()<2)return{success:true,copied:0,updated:0,skipped:0};
    var values=sourceSheet.getDataRange().getValues(),h=values[0]||[],idx={};h.forEach(function(k,i){idx[k]=i;});
    if(idx.Business_ID===undefined||idx.Email===undefined)return{success:true,copied:0,updated:0,skipped:0,reason:'SOURCE_SCHEMA_MISSING'};
    var email=String(ownerEmail||'').trim().toLowerCase(),matches=values.slice(1).filter(function(row){return String(row[idx.Business_ID]||'')===String(businessId||'')&&String(row[idx.Email]||'').trim().toLowerCase()===email;});
    if(!matches.length)return{success:true,copied:0,updated:0,skipped:0,reason:'NO_PREFERENCES'};
    var targetSheet=ensureClientNotificationPreferencesSheetV12_(target),tv=targetSheet.getDataRange().getValues(),th=tv[0]||[],tmap={};th.forEach(function(k,i){tmap[k]=i;});
    var existingRow=-1;for(var i=1;i<tv.length;i++){if(String(tv[i][tmap.Email]||'').trim().toLowerCase()===email&&String(tv[i][tmap.Business_ID]||'')===String(businessId||'')){existingRow=i+1;break;}}
    var row=matches[matches.length-1],out={Preference_ID:String(row[idx.Preference_ID]||'NP_MIGRATED_'+Utilities.getUuid().replace(/-/g,'').slice(0,12).toUpperCase()),Business_ID:String(businessId||''),Email:email,Feature_Updates:row[idx.Feature_Updates]!==false,System_Alerts:row[idx.System_Alerts]!==false,Weekly_Reports:row[idx.Weekly_Reports]===true,Security_Alerts:row[idx.Security_Alerts]!==false,Marketing_Updates:row[idx.Marketing_Updates]===true,Updated_At:row[idx.Updated_At]||new Date().toISOString()};
    if(existingRow){th.forEach(function(k,j){if(out[k]!==undefined)targetSheet.getRange(existingRow,j+1).setValue(out[k]);});return{success:true,copied:0,updated:1,skipped:0};}
    targetSheet.appendRow(th.map(function(k){return out[k]===undefined?'':out[k];}));
    return{success:true,copied:1,updated:0,skipped:0};
  }catch(error){console.error('Notification preference migration failed:',error);return{success:false,code:'NOTIFICATION_PREFERENCE_MIGRATION_FAILED',message:error.message};}
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