// ============================================================
// ClientFreeDataMigration.js - FREE -> PAID WORKSPACE MIGRATION
// ============================================================
// Migrates the free user's Finance + Ecommerce data into the
// newly provisioned paid client workspace. The target workspace
// remains the customer's isolated paid workspace.
// ============================================================

function migrateFreeWorkspaceDataToPaidClient(sourceWorkspaceId, targetWorkspaceId, businessId, businessName, ownerEmail) {
  try {
    if (!sourceWorkspaceId || !targetWorkspaceId || String(sourceWorkspaceId) === String(targetWorkspaceId)) {
      return {success:true,migrated:false,reason:'NO_FREE_WORKSPACE'};
    }

    var source = SpreadsheetApp.openById(sourceWorkspaceId);
    var target = SpreadsheetApp.openById(targetWorkspaceId);
    var modules = ['Finance','Ecommerce'];
    var results = {};
    var totalCopied = 0;

    modules.forEach(function(moduleName) {
      results[moduleName] = migrateFreeModuleSheet_(source, target, moduleName);
      totalCopied += results[moduleName].copied || 0;
    });

    migrateFreeBusinessMetadata_(source, target, businessId, businessName, ownerEmail);

    return {success:true,migrated:true,totalCopied:totalCopied,modules:results};
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
  sourceValues.slice(1).forEach(function(sourceRow) {
    var targetRow = targetHeaders.map(function(header) {
      var col = sourceHeaders.indexOf(header);
      return col === -1 ? '' : sourceRow[col];
    });

    var sourceIdCol = sourceHeaders.indexOf(targetIdHeader);
    var sourceId = sourceIdCol === -1 ? '' : String(sourceRow[sourceIdCol]||'').trim();
    var fingerprint = migrationRowFingerprint_(targetHeaders,targetRow);
    if ((sourceId && existing['ID:'+sourceId]) || existing['FP:'+fingerprint]) {
      skipped++;
      return;
    }

    // Preserve existing IDs/relationships when present. If legacy data has no
    // ID, generate one once during migration so future retries can identify it.
    if (!sourceId && targetIdCol !== -1) {
      targetRow[targetIdCol] = moduleName === 'Finance'
        ? 'TXN-MIGRATED-'+Utilities.getUuid().substring(0,12).toUpperCase()
        : 'ORD-MIGRATED-'+Utilities.getUuid().substring(0,12).toUpperCase();
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

function migrationRowFingerprint_(headers,row) {
  var values = headers.map(function(header,i) {
    if (header === 'Created_At' || header === 'Last_Login') return '';
    return String(row[i] === undefined || row[i] === null ? '' : row[i]);
  });
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,values.join('\u001f'))
  );
}

function migrateFreeBusinessMetadata_(source, target, businessId, businessName, ownerEmail) {
  var targetInfo = target.getSheetByName('Client_Info');
  if (!targetInfo) return;

  var metadata = {
    Source_Free_Workspace_ID: source.getId(),
    Business_ID: businessId || '',
    Business_Name: businessName || '',
    Owner_Email: String(ownerEmail || '').toLowerCase(),
    Data_Migrated_At: new Date().toISOString()
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
