// ============================================================
// Records.js - BizOS business record directory
// ============================================================

function getRecordsDirectory(sessionId) {
  try {
    const user = getUserFromSession(sessionId);
    if (!user) return { success: false, message: 'Session expired. Please login again.' };

    const modules = Array.isArray(user.accessibleModules) ? user.accessibleModules : [];
    const workspace = getWorkspaceFile(sessionId);
    const result = [];

    modules.forEach(function(moduleName) {
      if (!MODULES[moduleName]) return;
      const sheet = workspace.getSheetByName(MODULES[moduleName].sheet);
      result.push({
        key: moduleName,
        name: MODULES[moduleName].label || moduleName,
        module: moduleName,
        count: sheet && sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 0,
        source: 'module'
      });
    });

    if (modules.indexOf('Ecommerce') !== -1) {
      [
        ['Products', 'BusinessCenter_Products'],
        ['Customers', 'BusinessCenter_Customers'],
        ['Sales', 'BusinessCenter_Sales'],
        ['Sale Items', 'BusinessCenter_Sale_Items'],
        ['Stock Movements', 'BusinessCenter_Stock_Movements'],
        ['Business Center Finance', 'BusinessCenter_Finance_Transactions']
      ].forEach(function(item) {
        const sheet = workspace.getSheetByName(item[1]);
        result.push({
          key: 'BC:' + item[1],
          name: item[0],
          module: 'Business Center',
          count: sheet && sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 0,
          source: 'business_center'
        });
      });
    }

    return { success: true, records: result };
  } catch (error) {
    console.error('Records directory error:', error);
    return { success: false, message: error.message };
  }
}

function getRecordRows(recordKey, sessionId) {
  try {
    const user = getUserFromSession(sessionId);
    if (!user) return { success: false, message: 'Session expired. Please login again.' };

    const workspace = getWorkspaceFile(sessionId);
    let sheetName = '';
    let moduleName = '';

    if (String(recordKey).indexOf('BC:') === 0) {
      if (user.accessibleModules.indexOf('Ecommerce') === -1) {
        return { success: false, message: 'You do not have access to Business Center records.' };
      }
      sheetName = String(recordKey).substring(3);
      moduleName = 'Business Center';
    } else {
      moduleName = String(recordKey || '');
      if (!MODULES[moduleName] || user.accessibleModules.indexOf(moduleName) === -1) {
        return { success: false, message: 'You do not have access to this record type.' };
      }
      sheetName = MODULES[moduleName].sheet;
    }

    const sheet = workspace.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, recordType: moduleName, records: [], headers: [] };
    }

    const limit = 500;
    const lastRow = sheet.getLastRow();
    const startRow = Math.max(2, lastRow - limit + 1);
    const width = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, width).getValues()[0];
    const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, width).getValues();
    const records = values.reverse().map(function(row) {
      const obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i]; });
      return obj;
    });

    return { success: true, recordType: moduleName, headers: headers, records: records };
  } catch (error) {
    console.error('Record rows error:', error);
    return { success: false, message: error.message };
  }
}
