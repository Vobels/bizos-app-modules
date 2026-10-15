// ============================================================
// 📁 ClientData.gs - Client Data Operations
// ============================================================

// ============================================================
// 📊 GET CLIENT SHEET
// ============================================================

function getClientSheetById(sheetId) {
  try {
    return SpreadsheetApp.openById(sheetId);
  } catch (error) {
    console.error('Error opening client sheet:', error);
    return null;
  }
}

// ============================================================
// 📊 GET CLIENT MODULE DATA
// ============================================================

function getClientModuleData(moduleName, sessionId) {
  try {
    // Validate session
    const session = validateClientSession(sessionId);
    if (!session) {
      return { success: false, message: 'Session expired' };
    }
    
    const sheetId = session.sheetId;
    const ss = getClientSheetById(sheetId);
    if (!ss) {
      return { success: false, message: 'Sheet not found' };
    }
    
    const sheetName = moduleName + '_Data';
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: true, data: [] };
    }
    
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { success: true, data: [] };
    
    const headers = values[0];
    const records = values.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { 
        // Only include non-empty columns
        if (h && h.trim()) {
          obj[h.trim()] = row[i] || '';
        }
      });
      return obj;
    });
    
    return { success: true, data: records };
    
  } catch (error) {
    console.error('Get client module data error:', error);
    return { success: false, message: error.message };
  }
}

// ============================================================
// 📊 GET CLIENT DASHBOARD DATA
// ============================================================

function getClientDashboardData(sessionId) {
  // Keep one canonical client dashboard implementation. The richer
  // ClientDashboard.js response matches the master dashboard UI contract
  // (KPI objects with value/change, charts, health, and recent transactions).
  return getClientDashboard(sessionId);
}

// ============================================================
// 🏠 DEFAULT DASHBOARD
// ============================================================

function getDefaultDashboard(clientName, primaryColor) {
  return {
    kpis: { revenue: 0, expenses: 0, profit: 0, margin: 0 },
    modules: ['Finance', 'Ecommerce'],
    clientName: clientName || 'My Business',
    primaryColor: primaryColor || '#1a5276',
    logoUrl: ''
  };
}