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
  try {
    // Validate session
    const session = validateClientSession(sessionId);
    if (!session) {
      return { success: false, message: 'Session expired' };
    }
    
    const client = getClientById(session.clientId);
    if (!client) {
      return { success: false, message: 'Client not found' };
    }
    
    const sheetId = session.sheetId;
    const ss = getClientSheetById(sheetId);
    if (!ss) {
      return { 
        success: false, 
        dashboard: getDefaultDashboard(client.clientName, client.primaryColor) 
      };
    }
    
    // Get financial data
    const financeSheet = ss.getSheetByName('Financial_Data');
    let revenue = 0, expenses = 0;
    
    if (financeSheet) {
      const data = financeSheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0];
        const typeCol = headers.indexOf('Type');
        const amountCol = headers.indexOf('Amount');
        
        if (typeCol !== -1 && amountCol !== -1) {
          for (let i = 1; i < data.length; i++) {
            const amount = parseFloat(data[i][amountCol]) || 0;
            const type = data[i][typeCol] || '';
            if (type.toLowerCase() === 'revenue' || type.toLowerCase() === 'income') {
              revenue += Math.abs(amount);
            } else if (type.toLowerCase() === 'expense') {
              expenses += Math.abs(amount);
            }
          }
        }
      }
    }
    
    const profit = revenue - expenses;
    const margin = revenue > 0 ? (profit / revenue) * 100 : 0;
    
    // All 12 modules for sovereign tier
    const allModules = ['Finance', 'Ecommerce', 'Sales', 'CRM', 'HR', 'Logistics', 'Tax', 'Agro', 'Productivity', 'POS', 'Attendance', 'Warehouse'];
    
    return {
      success: true,
      dashboard: {
        kpis: { 
          revenue: Math.round(revenue * 100) / 100, 
          expenses: Math.round(expenses * 100) / 100, 
          profit: Math.round(profit * 100) / 100, 
          margin: Math.round(margin * 10) / 10 
        },
        modules: allModules,
        clientName: client.clientName,
        primaryColor: client.primaryColor || '#1a5276',
        logoUrl: client.logoUrl || ''
      }
    };
    
  } catch (error) {
    console.error('Get client dashboard error:', error);
    return { 
      success: false, 
      dashboard: getDefaultDashboard('My Business', '#1a5276') 
    };
  }
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