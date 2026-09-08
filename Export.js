// export.gs
function exportModuleData(moduleName, sessionId) {
  try {
    if (!userHasAccess(moduleName, sessionId)) throw new Error("Access denied");
    
    const workspace = getWorkspaceFile(sessionId);
    const sheet = workspace.getSheetByName(MODULES[moduleName].sheet);
    if (!sheet) return { success: false, message: "Sheet not found" };
    
    const data = sheet.getDataRange().getValues();
    const csv = data.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')).join('\n');
    
    return {
      success: true,
      csv: csv,
      filename: `${moduleName}_export_${new Date().toISOString().split('T')[0]}.csv`
    };
  } catch (error) {
    console.error('Export error:', error);
    return { success: false, message: error.message };
  }
}

function importCSV(moduleName, rows, sessionId) {
  try {
    if (!userHasAccess(moduleName, sessionId)) throw new Error("Access denied");
    
    const workspace = getWorkspaceFile(sessionId);
    const sheet = workspace.getSheetByName(MODULES[moduleName].sheet);
    if (!sheet) return { success: false, message: "Sheet not found" };
    
    const headers = getSheetHeaders(moduleName, sessionId);
    
    for (const row of rows) {
      const rowData = headers.map((_, i) => row[i] || "");
      sheet.appendRow(rowData);
    }
    
    return { success: true, message: `Imported ${rows.length} records successfully` };
  } catch (error) {
    console.error('Import error:', error);
    return { success: false, message: error.message };
  }
}
