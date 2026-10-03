// ============================================================
// 📁 ClientDashboard.gs - Client Dashboard Serving
// ============================================================
// NOTE: validateClientSession() is now in ClientAuth.gs
// NOTE: Helper functions (getModuleLabel, getModuleIcon, etc.) are now in ClientAuth.gs

// ============================================================
// 🎯 GET CLIENT DASHBOARD DATA
// ============================================================

function getClientDashboard(sessionId) {
  try {
    // Use the validateClientSession from ClientAuth.gs (NO prefix)
    const session = validateClientSession(sessionId);
    if (!session) {
      return { success: false, message: 'Session expired' };
    }
    
    const client = getClientById(session.clientId);
    if (!client) {
      return { success: false, message: 'Client not found' };
    }
    
    const sheetId = session.sheetId || client.sheetId;
    if (!sheetId) {
      return { success: false, dashboard: getDefaultDashboardData(client.clientName) };
    }
    
    let ss;
    try {
      ss = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      console.error('Could not open client sheet:', e);
      return { success: false, dashboard: getDefaultDashboardData(client.clientName) };
    }
    
    // Match the main BizOS dashboard KPI period: current month,
    // with percentage change against the previous month.
    const financeSheet = ss.getSheetByName('Financial_Data');
    let revenue = 0, expenses = 0, previousRevenue = 0, previousExpenses = 0;
    
    if (financeSheet) {
      const data = financeSheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0];
        const dateCol = headers.indexOf('Date');
        const typeCol = headers.indexOf('Type');
        const amountCol = headers.indexOf('Amount');
        const now = new Date();
        const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
        const previousMonthStart = new Date(now.getFullYear(), now.getMonth() - 1, 1);
        const previousMonthEnd = monthStart;
        
        if (typeCol !== -1 && amountCol !== -1) {
          for (let i = 1; i < data.length; i++) {
            const rowDate = dateCol !== -1 ? new Date(data[i][dateCol]) : null;
            if (!rowDate || isNaN(rowDate.getTime())) continue;
            const amount = parseFloat(data[i][amountCol]) || 0;
            const type = String(data[i][typeCol] || '').toLowerCase();
            const isRevenue = type === 'revenue' || type === 'income';
            const isExpense = type === 'expense';
            
            if (rowDate >= monthStart) {
              if (isRevenue) revenue += Math.abs(amount);
              else if (isExpense) expenses += Math.abs(amount);
            } else if (rowDate >= previousMonthStart && rowDate < previousMonthEnd) {
              if (isRevenue) previousRevenue += Math.abs(amount);
              else if (isExpense) previousExpenses += Math.abs(amount);
            }
          }
        }
      }
    }
    
    const profit = revenue - expenses;
    const previousProfit = previousRevenue - previousExpenses;
    const margin = revenue > 0 ? (profit / revenue) * 100 : 0;
    const previousMargin = previousRevenue > 0 ? (previousProfit / previousRevenue) * 100 : 0;
    const percentageChange = function(current, previous) {
      if (previous === 0) return current === 0 ? 0 : 100;
      return ((current - previous) / Math.abs(previous)) * 100;
    };
    
    // Get module counts
    const modules = ['Finance', 'Ecommerce', 'Sales', 'CRM', 'HR', 'Logistics', 'Tax', 'Agro', 'Productivity', 'POS', 'Attendance', 'Warehouse'];
    const moduleSummary = {};
    
    modules.forEach(moduleName => {
      const sheet = ss.getSheetByName(moduleName + '_Data');
      if (sheet) {
        const lastRow = sheet.getLastRow();
        moduleSummary[moduleName] = lastRow > 1 ? lastRow - 1 : 0;
      } else {
        moduleSummary[moduleName] = 0;
      }
    });
    
    // Get recent transactions
    let recentTransactions = [];
    if (financeSheet) {
      const data = financeSheet.getDataRange().getValues();
      const headers = data[0];
      if (data.length > 1) {
        const dateCol = headers.indexOf('Date');
        const descCol = headers.indexOf('Description');
        const amountCol = headers.indexOf('Amount');
        const typeCol = headers.indexOf('Type');
        
        recentTransactions = data.slice(1)
          .filter(row => row[dateCol])
          .sort((a, b) => new Date(b[dateCol]) - new Date(a[dateCol]))
          .slice(0, 5)
          .map(row => ({
            date: row[dateCol],
            description: row[descCol] || '',
            amount: row[amountCol] || 0,
            type: row[typeCol] || ''
          }));
      }
    }
    
    return {
      success: true,
      dashboard: {
        kpis: {
          revenue: Math.round(revenue * 100) / 100,
          expenses: Math.round(expenses * 100) / 100,
          profit: Math.round(profit * 100) / 100,
          margin: Math.round(margin * 10) / 10
        },
        modules: modules.map(name => ({
          name: name,
          label: getModuleLabel(name),
          icon: getModuleIcon(name),
          count: moduleSummary[name] || 0
        })),
        recentTransactions: recentTransactions,
        clientName: client.clientName || 'My Business',
        primaryColor: client.primaryColor || '#5D2A86',
        logoUrl: client.logoUrl || '',
        sheetId: sheetId
      }
    };
    
  } catch (error) {
    console.error('Get client dashboard error:', error);
    return { success: false, message: error.message };
  }
}

// ============================================================
// 🗑️ DELETE CLIENT RECORD
// ============================================================

function deleteClientRecord(moduleName, rowIndex, sessionId) {
  try {
    const session = validateClientSession(sessionId);
    if (!session) {
      return { success: false, message: 'Session expired' };
    }
    
    const sheetId = session.sheetId;
    if (!sheetId) {
      return { success: false, message: 'Sheet not found' };
    }
    
    let ss;
    try {
      ss = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      return { success: false, message: 'Could not open client sheet' };
    }
    
    const sheetName = moduleName + '_Data';
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, message: 'Module not found' };
    }
    
    // +2 because row 1 is headers, and index is 0-based
    sheet.deleteRow(rowIndex + 2);
    
    return { success: true, message: 'Record deleted' };
    
  } catch (error) {
    console.error('Delete client record error:', error);
    return { success: false, message: error.message };
  }
}

// ============================================================
// 📊 GET CLIENT SHEET - Keep this one
// ============================================================

function getClientSheetById(sheetId) {
  try {
    if (!sheetId) return null;
    return SpreadsheetApp.openById(sheetId);
  } catch (error) {
    console.error('Error opening client sheet:', error);
    return null;
  }
}