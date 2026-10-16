/**
 * BizOS Audit History
 * Append-only activity history for record changes in master and client workspaces.
 * Audit identities and timestamps are derived server-side from authenticated sessions.
 */
var BIZOS_AUDIT_HEADERS_ = [
  'Event_ID', 'Timestamp', 'Business_ID', 'Workspace_ID', 'Actor_Email',
  'Actor_Name', 'Actor_Role', 'App', 'Action', 'Record_ID',
  'Changed_Fields_JSON', 'Before_JSON', 'After_JSON'
];

function auditSafeSnapshot_(record) {
  if (!record || typeof record !== 'object') return {};
  var blocked = /password|token|secret|api.?key|credential|authorization|session.?id/i;
  var safe = {};
  Object.keys(record).forEach(function(key) {
    if (key.charAt(0) === '_' || blocked.test(key)) return;
    var value = record[key];
    if (value instanceof Date) value = value.toISOString();
    if (typeof value === 'function' || value === undefined) return;
    safe[key] = value;
  });
  return safe;
}

function auditChangedFields_(beforeRecord, afterRecord) {
  var before = auditSafeSnapshot_(beforeRecord || {});
  var after = auditSafeSnapshot_(afterRecord || {});
  var keys = Object.keys(before).concat(Object.keys(after)).filter(function(key, i, all) {
    return all.indexOf(key) === i;
  });
  return keys.filter(function(key) {
    return JSON.stringify(before[key] === undefined ? null : before[key]) !==
      JSON.stringify(after[key] === undefined ? null : after[key]);
  }).map(function(key) {
    return {
      field: key,
      before: before[key] === undefined ? null : before[key],
      after: after[key] === undefined ? null : after[key]
    };
  });
}

function auditRecordChange_(workspace, sessionId, moduleName, action, recordId, beforeRecord, afterRecord, clientSession) {
  try {
    if (!workspace) throw new Error('Audit workspace is unavailable.');
    var actor = null, businessId = '', workspaceId = '';
    if (clientSession) {
      actor = clientSession.user || {};
      businessId = String(clientSession.businessId || '');
      workspaceId = String(clientSession.sheetId || workspace.getId() || '');
    } else {
      actor = getUserFromSession(sessionId);
      if (!actor) throw new Error('Audit actor session is unavailable.');
      businessId = String(actor.businessId || '');
      workspaceId = String(workspace.getId() || '');
    }
    var sheet = workspace.getSheetByName('Activity_History');
    if (!sheet) {
      sheet = workspace.insertSheet('Activity_History');
      sheet.getRange(1, 1, 1, BIZOS_AUDIT_HEADERS_.length).setValues([BIZOS_AUDIT_HEADERS_]);
      sheet.setFrozenRows(1);
    } else if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, BIZOS_AUDIT_HEADERS_.length).setValues([BIZOS_AUDIT_HEADERS_]);
      sheet.setFrozenRows(1);
    }
    var before = auditSafeSnapshot_(beforeRecord || {});
    var after = auditSafeSnapshot_(afterRecord || {});
    var changed = auditChangedFields_(before, after);
    sheet.appendRow([
      Utilities.getUuid(),
      new Date().toISOString(),
      businessId,
      workspaceId,
      String(actor.email || (clientSession && clientSession.email) || ''),
      String(actor.name || ''),
      String(actor.role || ''),
      String(moduleName || ''),
      String(action || ''),
      String(recordId == null ? '' : recordId),
      JSON.stringify(changed),
      JSON.stringify(before),
      JSON.stringify(after)
    ]);
    return true;
  } catch (error) {
    // Keep the main data operation's response stable, but surface audit failures
    // in server logs so they are not silently mistaken for successful auditing.
    console.error('BizOS audit history write failed:', String(error && error.message || error));
    return false;
  }
}

function getRecordAuditHistory(moduleName, recordId, sessionId) {
  try {
    if (!userHasAccess(moduleName, sessionId)) return {success:false, message:'Access denied'};
    var workspace = getWorkspaceFile(sessionId);
    if (!workspace) return {success:false, message:'Workspace is unavailable.'};
    var sheet = workspace.getSheetByName('Activity_History');
    if (!sheet || sheet.getLastRow() < 2) return {success:true, history:[]};
    var values = sheet.getDataRange().getValues(), headers = values[0].map(String);
    var moduleCol = headers.indexOf('App'), recordCol = headers.indexOf('Record_ID');
    if (moduleCol < 0 || recordCol < 0) return {success:true, history:[]};
    var history = [];
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][moduleCol]) !== String(moduleName) ||
          String(values[i][recordCol]) !== String(recordId)) continue;
      var item = {};
      headers.forEach(function(header, j) { item[header] = values[i][j]; });
      ['Changed_Fields_JSON','Before_JSON','After_JSON'].forEach(function(key) {
        try { item[key] = JSON.parse(item[key] || (key === 'Changed_Fields_JSON' ? '[]' : '{}')); }
        catch (ignore) { item[key] = key === 'Changed_Fields_JSON' ? [] : {}; }
      });
      history.push(item);
    }
    history.reverse();
    return {success:true, history:history};
  } catch (error) {
    console.error('getRecordAuditHistory failed:', error);
    return {success:false, message:'We could not load this record’s history.'};
  }
}

function getActivityHistory(moduleName, limit, sessionId) {
  try {
    var user = getUserFromSession(sessionId);
    if (!user) return {success:false, message:'Session expired. Please sign in again.'};
    var requestedModule = String(moduleName || '').trim();
    if (requestedModule && !userHasAccess(requestedModule, sessionId)) {
      return {success:false, message:'Access denied'};
    }
    var workspace = getWorkspaceFile(sessionId);
    var sheet = workspace && workspace.getSheetByName('Activity_History');
    if (!sheet || sheet.getLastRow() < 2) return {success:true, history:[]};
    var values = sheet.getDataRange().getValues(), headers = values[0].map(String);
    var maxRows = Math.max(1, Math.min(Number(limit) || 100, 500));
    var result = [];
    for (var i = values.length - 1; i >= 1 && result.length < maxRows; i--) {
      var item = {};
      headers.forEach(function(header, j) { item[header] = values[i][j]; });
      if (requestedModule && String(item.App) !== requestedModule) continue;
      ['Changed_Fields_JSON','Before_JSON','After_JSON'].forEach(function(key) {
        try { item[key] = JSON.parse(item[key] || (key === 'Changed_Fields_JSON' ? '[]' : '{}')); }
        catch (ignore) { item[key] = key === 'Changed_Fields_JSON' ? [] : {}; }
      });
      result.push(item);
    }
    return {success:true, history:result};
  } catch (error) {
    console.error('getActivityHistory failed:', error);
    return {success:false, message:'We could not load activity history.'};
  }
}

function getClientRecordAuditHistory(moduleName, recordId, sessionId) {
  try {
    var session = validateClientSession(sessionId);
    if (!session) return {success:false,message:'Session expired. Please sign in again.'};
    if (getClientAccessibleModulesV12_(session).indexOf(String(moduleName)) === -1) {
      return {success:false,message:'Access denied'};
    }
    var workspace = SpreadsheetApp.openById(session.sheetId);
    var sheet = workspace.getSheetByName('Activity_History');
    if (!sheet || sheet.getLastRow() < 2) return {success:true,history:[]};
    var values = sheet.getDataRange().getValues(), headers = values[0].map(String);
    var moduleCol = headers.indexOf('App'), recordCol = headers.indexOf('Record_ID');
    if (moduleCol < 0 || recordCol < 0) return {success:true,history:[]};
    var history = [];
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][moduleCol]) !== String(moduleName) || String(values[i][recordCol]) !== String(recordId)) continue;
      var item = {};
      headers.forEach(function(header,j){item[header]=values[i][j];});
      ['Changed_Fields_JSON','Before_JSON','After_JSON'].forEach(function(key){
        try { item[key]=JSON.parse(item[key] || (key==='Changed_Fields_JSON'?'[]':'{}')); }
        catch(ignore){item[key]=key==='Changed_Fields_JSON'?[]:{};}
      });
      history.push(item);
    }
    history.reverse();
    return {success:true,history:history};
  } catch (error) {
    console.error('getClientRecordAuditHistory failed:',error);
    return {success:false,message:'We could not load this record’s history.'};
  }
}
