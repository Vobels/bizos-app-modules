// ============================================================
// 📁 ClientAuth.gs - FIXED VERSION
// ============================================================

/**
 * ✅ CLIENT LOGIN - IMPROVED WITH BETTER ERROR HANDLING
 */
function clientLogin(clientId, email, password, sheetId, businessId, businessName) {
  try {
    console.log('🔐 ========================================');
    console.log('🔐 CLIENT LOGIN ATTEMPT');
    console.log('🔐 ========================================');
    console.log('🆔 Client ID:', clientId);
    console.log('📧 Email:', email);
    console.log('🏢 Business Name:', businessName);

    // ===== VALIDATE INPUT =====
    if (!email || !password) {
      console.error('❌ Missing email or password');
      return { 
        success: false, 
        message: 'Email and password are required',
        code: 'INVALID_INPUT'
      };
    }

    // ===== GET CLIENT INFO =====
    const client = getClientById(clientId);
    if (!client) {
      console.error('❌ Client not found:', clientId);
      return { 
        success: false, 
        message: 'Client not found. Please contact support.',
        code: 'CLIENT_NOT_FOUND'
      };
    }

    console.log('✅ Client found:', client.clientName);

    // ===== AUTHENTICATE USER (Against main app database) =====
    console.log('🔐 Authenticating user...');
    const authResult = authenticateUser(email, password, null);

    if (!authResult.success) {
      console.error('❌ Authentication failed:', authResult.message);
      return {
        success: false,
        message: authResult.message,
        code: 'AUTH_FAILED'
      };
    }

    console.log('✅ User authenticated:', email);

    // ===== VERIFY USER IS CLIENT =====
    const user = authResult.user;
    if (!user.isClient) {
      console.warn('⚠️ User is not marked as client. Allowing access anyway.');
    }

    // ===== CREATE CLIENT SESSION =====
    const sessionId = 'CLIENT_SESS_' + Utilities.getUuid();
    const cache = CacheService.getScriptCache();

    const clientSession = {
      email: email,
      clientId: clientId,
      sheetId: sheetId || client.sheetId || '',
      businessId: businessId || client.businessId || clientId,
      businessName: businessName || client.clientName || client.businessName || 'My Business',
      created: Date.now(),
      isClient: true,
      user: user,
      apiKey: client.apiKey || ''
    };

    cache.put(sessionId, JSON.stringify(clientSession), 86400);
    console.log('✅ Client session created:', sessionId.substring(0, 20) + '...');

    // ===== ALSO CREATE REGULAR SESSION FOR COMPATIBILITY =====
    const regularSessionId = createSession(email.toLowerCase());

    // ===== RETURN SUCCESS =====
    const response = {
      success: true,
      sessionId: sessionId,
      regularSessionId: regularSessionId,
      message: 'Login successful',
      clientName: client.clientName || client.businessName || 'My Business',
      primaryColor: client.primaryColor || '#5D2A86',
      logoUrl: client.logoUrl || '',
      user: {
        email: user.email,
        name: user.name,
        role: user.role,
        clientId: clientId,
        businessId: businessId || client.businessId || clientId
      },
      code: 'SUCCESS'
    };

    console.log('✅ ========================================');
    console.log('✅ CLIENT LOGIN SUCCESSFUL');
    console.log('✅ ========================================');
    console.log('📧 Email:', email);
    console.log('🏢 Client:', client.clientName);
    console.log('🆔 Session:', sessionId.substring(0, 20) + '...');

    return response;

  } catch (error) {
    console.error('❌ ========================================');
    console.error('❌ CLIENT LOGIN ERROR');
    console.error('❌ ========================================');
    console.error('Error message:', error.message);
    console.error('Error stack:', error.stack);

    return {
      success: false,
      message: 'Login error: ' + error.message,
      code: 'LOGIN_ERROR',
      details: error.stack
    };
  }
}

/**
 * ✅ VALIDATE CLIENT SESSION
 */
function validateClientSession(sessionId) {
  if (!sessionId) {
    console.log('⚠️ validateClientSession: No sessionId provided');
    return null;
  }

  try {
    const cache = CacheService.getScriptCache();
    const data = cache.get(sessionId);

    if (!data) {
      console.log('⚠️ validateClientSession: No session in cache for ID:', sessionId.substring(0, 20) + '...');
      return null;
    }

    const session = JSON.parse(data);

    if (Date.now() - session.created > 86400000) {
      cache.remove(sessionId);
      console.log('⚠️ validateClientSession: Session expired');
      return null;
    }

    const client = getClientById(session.clientId);
    if (!client) {
      console.log('⚠️ validateClientSession: Client not found');
      return null;
    }

    if (client.status && client.status !== 'active') {
      console.log('⚠️ validateClientSession: Client inactive');
      return null;
    }

    console.log('✅ Client session validated for:', session.email);
    return session;

  } catch (error) {
    console.error('❌ validateClientSession error:', error);
    return null;
  }
}

/**
 * ✅ GET CLIENT USER FROM SESSION
 */
function getClientUserFromSession(sessionId) {
  const session = validateClientSession(sessionId);
  if (!session) {
    console.log('⚠️ getClientUserFromSession: Invalid session');
    return null;
  }

  const user = getUserFromSession(session.regularSessionId || sessionId);

  if (!user) {
    console.log('ℹ️ Using cached user data from session');
    return {
      email: session.email,
      name: session.user?.name || session.email,
      role: session.user?.role || 'owner',
      clientId: session.clientId,
      businessId: session.businessId,
      isClient: true
    };
  }

  const client = getClientById(session.clientId);
  const enrichedUser = {
    ...user,
    clientId: session.clientId,
    clientName: client?.clientName || client?.businessName || 'My Business',
    primaryColor: client?.primaryColor || '#5D2A86',
    logoUrl: client?.logoUrl || '',
    sheetId: session.sheetId,
    isClient: true,
    apiKey: client?.apiKey || ''
  };

  console.log('✅ Client user loaded from session:', enrichedUser.email);
  return enrichedUser;
}

/**
 * ✅ CLIENT LOGOUT
 */
function clientLogout(sessionId) {
  try {
    if (sessionId) {
      const cache = CacheService.getScriptCache();
      cache.remove(sessionId);
      console.log('🚪 Client session logged out:', sessionId.substring(0, 20) + '...');
    }
    return { 
      success: true, 
      message: "Logged out successfully",
      code: 'LOGOUT_SUCCESS'
    };
  } catch (error) {
    console.error('❌ Logout error:', error);
    return { 
      success: false, 
      message: error.message,
      code: 'LOGOUT_ERROR'
    };
  }
}

/**
 * ✅ GET CLIENT DASHBOARD DATA - FIXED VERSION
 */
function getClientDashboardData(sessionId) {
  try {
    const session = validateClientSession(sessionId);
    if (!session) {
      return { 
        success: false, 
        message: "Session expired. Please login again.",
        code: 'SESSION_EXPIRED'
      };
    }

    const user = getClientUserFromSession(sessionId);
    const client = getClientById(session.clientId);

    const sheetId = session.sheetId || client?.sheetId;
    if (!sheetId) {
      return { 
        success: true, 
        dashboard: getDefaultDashboardData(user?.clientName || 'My Business'),
        code: 'NO_SHEET'
      };
    }

    let ss;
    try {
      ss = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      console.error('⚠️ Could not open client sheet:', e);
      return { 
        success: true, 
        dashboard: getDefaultDashboardData(user?.clientName || 'My Business'),
        code: 'SHEET_NOT_ACCESSIBLE'
      };
    }

    // Get KPIs
    let revenue = 0, expenses = 0;
    const financeSheet = ss.getSheetByName('Financial_Data');

    if (financeSheet) {
      const data = financeSheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0];
        const typeCol = headers.indexOf('Type');
        const amountCol = headers.indexOf('Amount');

        if (typeCol !== -1 && amountCol !== -1) {
          for (let i = 1; i < data.length; i++) {
            const amount = parseFloat(data[i][amountCol]) || 0;
            const type = (data[i][typeCol] || '').toLowerCase();
            if (type === 'revenue' || type === 'income') {
              revenue += Math.abs(amount);
            } else if (type === 'expense') {
              expenses += Math.abs(amount);
            }
          }
        }
      }
    }

    const profit = revenue - expenses;
    const margin = revenue > 0 ? (profit / revenue) * 100 : 0;

    // ✅ FIX: Use getModuleSummary from modules.gs (returns proper structure with label & icon)
    let moduleSummary = {};
    try {
      // Try to get from modules.gs
      moduleSummary = getModuleSummary(sessionId);
    } catch (e) {
      console.error('⚠️ getModuleSummary error, using fallback:', e);
      // Fallback: build manually
      const modules = ['Finance', 'Ecommerce', 'Sales', 'CRM', 'HR', 'Logistics', 'Tax', 'Agro', 'Productivity', 'POS', 'Attendance', 'Warehouse'];
      modules.forEach(name => {
        const sheet = ss.getSheetByName(name + '_Data');
        const count = sheet && sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 0;
        moduleSummary[name] = {
          label: getModuleLabel(name),
          icon: getModuleIcon(name),
          count: count,
          isAccessible: true
        };
      });
    }

    // Convert the summary object to an array
    const moduleArray = [];
    const moduleNames = ['Finance', 'Ecommerce', 'Sales', 'CRM', 'HR', 'Logistics', 'Tax', 'Agro', 'Productivity', 'POS', 'Attendance', 'Warehouse'];
    
    moduleNames.forEach(name => {
      if (moduleSummary[name]) {
        moduleArray.push({
          name: name,
          label: moduleSummary[name].label || name,
          icon: moduleSummary[name].icon || 'folder',
          count: moduleSummary[name].count || 0,
          isAccessible: moduleSummary[name].isAccessible !== undefined ? moduleSummary[name].isAccessible : true
        });
      } else {
        // Fallback if module not in summary
        const sheet = ss.getSheetByName(name + '_Data');
        const count = sheet && sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 0;
        moduleArray.push({
          name: name,
          label: getModuleLabel(name),
          icon: getModuleIcon(name),
          count: count,
          isAccessible: true
        });
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

        if (dateCol !== -1 && amountCol !== -1) {
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
    }

    const clientName = client?.clientName || client?.businessName || user?.clientName || 'My Business';
    const primaryColor = client?.primaryColor || '#5D2A86';

    return {
      success: true,
      dashboard: {
        kpis: {
          revenue: Math.round(revenue * 100) / 100,
          expenses: Math.round(expenses * 100) / 100,
          profit: Math.round(profit * 100) / 100,
          margin: Math.round(margin * 10) / 10
        },
        modules: moduleArray,  // ✅ Now passing an ARRAY with label, icon, count
        recentTransactions: recentTransactions,
        clientName: clientName,
        primaryColor: primaryColor,
        logoUrl: client?.logoUrl || '',
        sheetId: sheetId
      },
      code: 'SUCCESS'
    };

  } catch (error) {
    console.error('❌ Get client dashboard error:', error);
    return { 
      success: false, 
      message: error.message,
      code: 'DASHBOARD_ERROR'
    };
  }
}


/**
 * ✅ GET CLIENT MODULE DATA
 */
function getClientModuleData(moduleName, sessionId) {
  try {
    const session = validateClientSession(sessionId);
    if (!session) {
      return { 
        success: false, 
        message: "Session expired",
        code: 'SESSION_EXPIRED'
      };
    }

    const sheetId = session.sheetId;
    if (!sheetId) {
      return { success: true, data: [], code: 'NO_SHEET' };
    }

    let ss;
    try {
      ss = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      console.error('⚠️ Could not open client sheet:', e);
      return { success: true, data: [], code: 'SHEET_NOT_ACCESSIBLE' };
    }

    const sheetName = moduleName + '_Data';
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.getSheetByName(moduleName);
    }

    if (!sheet) {
      return { success: true, data: [], code: 'SHEET_NOT_FOUND' };
    }

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { success: true, data: [], code: 'NO_DATA' };

    const headers = values[0];
    const records = values.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        if (h && h.trim()) {
          obj[h.trim()] = row[i] !== undefined ? row[i] : '';
        }
      });
      return obj;
    });

    console.log('✅ Loaded', records.length, 'records from', moduleName);
    return { success: true, data: records, code: 'SUCCESS' };

  } catch (error) {
    console.error('❌ Get client module data error:', error);
    return { 
      success: false, 
      message: error.message,
      code: 'FETCH_ERROR'
    };
  }
}

/**
 * ✅ SAVE CLIENT MODULE RECORD
 * (Renamed from saveClientRecord to avoid collision with
 *  Master Backend.gs's saveClientRecord(settings), which
 *  writes to the Clients sheet during client provisioning.)
 */
function saveClientModuleRecord(moduleName, record, sessionId) {
  try {
    const session = validateClientSession(sessionId);
    if (!session) {
      return { 
        success: false, 
        message: "Session expired",
        code: 'SESSION_EXPIRED'
      };
    }

    const sheetId = session.sheetId;
    if (!sheetId) {
      return { 
        success: false, 
        message: "No sheet found for client",
        code: 'NO_SHEET'
      };
    }

    let ss;
    try {
      ss = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      return { 
        success: false, 
        message: "Could not open client sheet",
        code: 'SHEET_NOT_ACCESSIBLE'
      };
    }

    const sheetName = moduleName + '_Data';
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const defaultHeaders = ['Date', 'Description', 'Amount', 'Created_At'];
      sheet.appendRow(defaultHeaders);
      const headerRange = sheet.getRange(1, 1, 1, defaultHeaders.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f4f6');
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];

    const rowData = [];
    headers.forEach(h => {
      rowData.push(record[h] || '');
    });

    if (!record.Created_At && headers.includes('Created_At')) {
      const idx = headers.indexOf('Created_At');
      if (idx !== -1 && !rowData[idx]) {
        rowData[idx] = new Date().toISOString();
      }
    }

    sheet.appendRow(rowData);
    console.log('✅ Saved record to', moduleName);

    return { 
      success: true, 
      message: "Record saved successfully",
      code: 'SAVE_SUCCESS'
    };

  } catch (error) {
    console.error('❌ Save client record error:', error);
    return { 
      success: false, 
      message: error.message,
      code: 'SAVE_ERROR'
    };
  }
}

// ============================================================
// 🛠 HELPER FUNCTIONS
// ============================================================

function getModuleLabel(moduleName) {
  const labels = {
    'Finance': 'Financial Management',
    'Ecommerce': 'E-commerce',
    'Sales': 'Sales Pipeline',
    'CRM': 'Customer Relations',
    'HR': 'Human Resources',
    'Logistics': 'Logistics',
    'Tax': 'Tax Compliance',
    'Agro': 'Agriculture',
    'Productivity': 'Productivity Suite',
    'POS': 'Point of Sale',
    'Attendance': 'Staff Attendance',
    'Warehouse': 'Warehouse Management'
  };
  return labels[moduleName] || moduleName;
}

function getModuleIcon(moduleName) {
  const icons = {
    'Finance': 'fa-calculator',
    'Ecommerce': 'fa-cart-shopping',
    'Sales': 'fa-chart-line',
    'CRM': 'fa-users',
    'HR': 'fa-briefcase',
    'Logistics': 'fa-truck',
    'Tax': 'fa-file-invoice',
    'Agro': 'fa-seedling',
    'Productivity': 'fa-clock',
    'POS': 'fa-cash-register',
    'Attendance': 'fa-fingerprint',
    'Warehouse': 'fa-warehouse'
  };
  return icons[moduleName] || 'fa-folder';
}

function getDefaultDashboardData(clientName) {
  return {
    kpis: { revenue: 0, expenses: 0, profit: 0, margin: 0 },
    modules: ['Finance', 'Ecommerce'].map(name => ({
      name: name,
      label: getModuleLabel(name),
      icon: getModuleIcon(name),
      count: 0
    })),
    recentTransactions: [],
    clientName: clientName || 'My Business',
    primaryColor: '#5D2A86',
    logoUrl: ''
  };
}