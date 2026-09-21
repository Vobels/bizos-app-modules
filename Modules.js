// modules.gs
function getModuleData(moduleName, sessionId) {
  try {
    if (!userHasAccess(moduleName, sessionId)) throw new Error("Access denied");
    
    const workspace = getWorkspaceFile(sessionId);
    const sheet = workspace.getSheetByName(MODULES[moduleName].sheet);
    
    if (!sheet) {
      console.log(`Sheet ${MODULES[moduleName].sheet} not found in workspace`);
      return [];
    }
    
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return [];
    
    const headers = values[0];
    return values.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    }).filter(row => Object.keys(row).length > 0);
    
  } catch (error) {
    console.error('Error getting module data:', error);
    return [];
  }
}



function getModuleDataWithSync(moduleName, sessionId) {
  // Handle case where parameters are swapped
  if (typeof moduleName === 'string' && moduleName.length > 30 && sessionId === undefined) {
    sessionId = moduleName;
    moduleName = 'Finance';
  }
  
  try {
    if (!userHasAccess(moduleName, sessionId)) throw new Error("Access denied");
    
    const user = getUserFromSession(sessionId);
    const workspace = getWorkspaceFile(sessionId);
    const sheet = workspace.getSheetByName(MODULES[moduleName].sheet);
    if (!sheet) return { success: false, message: "Module not found" };
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, records: [], financialSummary: { totalRevenue: 0, totalExpenses: 0 }, recordCount: 0 };
    
    // Limit to last 1000 rows for performance
    const startRow = Math.max(1, lastRow - 1000);
    const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn()).getValues();
    const headers = values[0];
    
    const records = values.slice(1).map(row => {
      const record = {};
      headers.forEach((h, i) => { record[h] = row[i]; });
      return record;
    });
    
    // Get financial summary (also limited)
    const financeSheet = workspace.getSheetByName('Financial_Data');
    let financialSummary = { totalRevenue: 0, totalExpenses: 0 };
    
    if (financeSheet) {
      const financeLastRow = financeSheet.getLastRow();
      const financeStartRow = Math.max(1, financeLastRow - 5000);
      const financeData = financeSheet.getRange(financeStartRow, 1, financeLastRow - financeStartRow + 1, financeSheet.getLastColumn()).getValues();
      const financeHeaders = financeData[0];
      const moduleSourceCol = financeHeaders.indexOf('Module_Source');
      const typeCol = financeHeaders.indexOf('Type');
      const amountCol = financeHeaders.indexOf('Amount');
      
      for (const row of financeData.slice(1)) {
        if (row[moduleSourceCol] === moduleName) {
          const amount = parseFloat(row[amountCol]) || 0;
          if (row[typeCol] === 'revenue') {
            financialSummary.totalRevenue += amount;
          } else if (row[typeCol] === 'expense') {
            financialSummary.totalExpenses += Math.abs(amount);
          }
        }
      }
    }
    
    return {
      success: true,
      records: records,
      financialSummary: financialSummary,
      recordCount: records.length
    };
    
  } catch (error) {
    console.error('Get module data error:', error);
    return { success: false, message: error.message };
  }
}

function getModuleSummary(sessionId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `module_summary_${sessionId}`;
  const cached = cache.get(cacheKey);
  if (cached) {
    console.log("Returning cached module summary");
    return JSON.parse(cached);
  }
  
  try {
    let email = null;
    let subscriptionTier = 'free';
    let isDemo = false;
    let userRole = 'staff';
    let businessId = null;
    let workspace = SpreadsheetApp.getActiveSpreadsheet();
    let staffModules = [];
    
    try {
      email = validateSession(sessionId);
      if (email) {
        const userSheet = getOrCreateUserSheet();
        const userData = userSheet.getDataRange().getValues();
        const userHeaders = userData[0];
        
        const emailCol = userHeaders.indexOf('Email');
        const businessIdCol = userHeaders.indexOf('Business_ID');
        const roleCol = userHeaders.indexOf('Role');
        const isDemoCol = userHeaders.indexOf('Is_Demo');
        
        let userRow = null;
        for (let i = 1; i < userData.length; i++) {
          if (userData[i][emailCol] && userData[i][emailCol].toLowerCase() === email.toLowerCase()) {
            userRow = userData[i];
            break;
          }
        }
        
        if (userRow) {
          businessId = userRow[businessIdCol];
          userRole = userRow[roleCol] || 'staff';
          isDemo = userRow[isDemoCol] === 'YES';
          
          // Get business info
          const businessSheet = getOrCreateBusinessSheet();
          const businessData = businessSheet.getDataRange().getValues();
          const businessHeaders = businessData[0];
          
          const bizIdCol = businessHeaders.indexOf('Business_ID');
          const tierCol = businessHeaders.indexOf('Subscription_Tier');
          const workspaceIdCol = businessHeaders.indexOf('Workspace_ID');
          
          for (let i = 1; i < businessData.length; i++) {
            if (businessData[i][bizIdCol] === businessId) {
              subscriptionTier = businessData[i][tierCol] || 'free';
              const workspaceId = businessData[i][workspaceIdCol];
              if (workspaceId) {
                try {
                  workspace = SpreadsheetApp.openById(workspaceId);
                } catch(e) {
                  workspace = SpreadsheetApp.getActiveSpreadsheet();
                }
              }
              break;
            }
          }
          
          // ===== KEY CHANGE: Get staff assigned modules =====
          // If user is staff, get their assigned modules
          if (userRole === 'staff' || userRole === 'viewer' || userRole === 'manager') {
            staffModules = getStaffModules(email, businessId);
            console.log('Staff modules for', email, ':', staffModules);
          }
        }
      }
    } catch(e) {
      console.log("Session validation error, using defaults:", e);
    }
    
    console.log("Module summary - Tier:", subscriptionTier, "isDemo:", isDemo, "Role:", userRole);
    
    // ===== DEFINE ALL MODULES =====
    const allModules = {
      'Finance': { sheet: 'Financial_Data', label: 'Financial Management', icon: 'calculator' },
      'Ecommerce': { sheet: 'Ecommerce_Data', label: 'E-commerce', icon: 'cart' },
      'Sales': { sheet: 'Sales_Data', label: 'Sales Pipeline', icon: 'chart-line' },
      'CRM': { sheet: 'CRM_Data', label: 'Customer Relations', icon: 'users' },
      'HR': { sheet: 'HR_Data', label: 'Human Resources', icon: 'briefcase' },
      'Logistics': { sheet: 'Logistics_Data', label: 'Logistics', icon: 'truck' },
      'Tax': { sheet: 'Tax_Data', label: 'Tax Compliance', icon: 'file-invoice' },
      'Agro': { sheet: 'Agro_Data', label: 'Agriculture', icon: 'seedling' },
      'Productivity': { sheet: 'Productivity_Data', label: 'Productivity Suite', icon: 'clock' },
      'POS': { sheet: 'POS_Data', label: 'Point of Sale', icon: 'cash-register' },
      'Attendance': { sheet: 'Attendance_Data', label: 'Staff Attendance', icon: 'fingerprint' },
      'Warehouse': { sheet: "Warehouse_Data", label: "Warehouse Management", icon: "warehouse" }
    };
    
    // ===== DETERMINE ACCESSIBLE MODULES =====
    let accessibleModules = [];
    const isOwnerOrAdmin = userRole === 'owner' || userRole === 'admin';
    
    // Owner/Admin: Get modules based on subscription tier
    if (isOwnerOrAdmin) {
      if (isDemo) {
        // Demo: All modules but read-only
        accessibleModules = Object.keys(allModules);
      } else if (subscriptionTier === 'free' || subscriptionTier === 'starter') {
        // Free: Only Finance + Ecommerce
        accessibleModules = ['Finance', 'Ecommerce'];
      } else {
        // Sovereign/Enterprise: All modules
        accessibleModules = Object.keys(allModules);
      }
    } else {
      // Staff/Manager/Viewer: Only assigned modules
      accessibleModules = staffModules.length > 0 ? staffModules : [];
      // If no modules assigned, they get Finance by default (view only)
      if (accessibleModules.length === 0) {
        accessibleModules = ['Finance'];
      }
    }
    
    // ===== BUILD SUMMARY =====
    const summary = {};
    for (const [key, module] of Object.entries(allModules)) {
      // Check if this module is accessible
      const isAccessible = accessibleModules.includes(key);
      
      // Get record count if accessible
      let count = 0;
      if (isAccessible) {
        try {
          const sheet = workspace.getSheetByName(module.sheet);
          if (sheet && sheet.getLastRow() > 1) {
            count = sheet.getLastRow() - 1;
          }
        } catch (e) {
          count = 0;
        }
      }
      
      summary[key] = {
        label: module.label,
        icon: module.icon,
        count: count,
        isAccessible: isAccessible,
        isReadOnly: isDemo && isAccessible
      };
    }
    
    cache.put(cacheKey, JSON.stringify(summary), 300);
    return summary;
    
  } catch (error) {
    console.error('Error in getModuleSummary:', error);
    return {
      'Finance': { label: 'Financial Management', icon: 'calculator', count: 0, isAccessible: true },
      'Ecommerce': { label: 'E-commerce', icon: 'cart', count: 0, isAccessible: true }
    };
  }
}


function autoSyncToFinancial(moduleName, recordData, workspaceId, sessionId) {
  try {
    const workspace = SpreadsheetApp.openById(workspaceId);
    const financeSheet = workspace.getSheetByName('Financial_Data');
    if (!financeSheet) return { success: false, message: "Financial_Data sheet not found" };
    
    const user = getUserFromSession(sessionId);
    const now = new Date().toISOString();
    
    let financialEntries = [];
    
    // Determine financial impact based on module and record type
    switch(moduleName) {
      case 'Sales':
        if (recordData.Stage === 'closed_won' && recordData.Total_Amount > 0) {
          financialEntries.push({
            Transaction_ID: generateTransactionId(),
            Date: recordData.Date || now.split('T')[0],
            Type: 'revenue',
            Category: 'sales',
            Subcategory: recordData.Product_Service || 'product',
            Module_Source: 'Sales',
            Module_Record_ID: recordData.Deal_ID,
            Description: `Sale: ${recordData.Product_Service} to ${recordData.Customer_Name}`,
            Amount: parseFloat(recordData.Total_Amount),
            Payment_Method: recordData.Payment_Method || 'pending',
            Reference: recordData.Deal_ID,
            Status: 'completed',
            Created_By: user?.email || recordData.Created_By,
            Created_At: now
          });
        }
        break;
        
      case 'Ecommerce':
        if (recordData.Total_Amount > 0) {
          // Revenue entry
          financialEntries.push({
            Transaction_ID: generateTransactionId(),
            Date: recordData.Date || now.split('T')[0],
            Type: 'revenue',
            Category: 'ecommerce',
            Subcategory: 'product_sales',
            Module_Source: 'Ecommerce',
            Module_Record_ID: recordData.Order_ID,
            Description: `Order ${recordData.Order_ID}: ${recordData.Product_Name}`,
            Amount: parseFloat(recordData.Total_Amount),
            Payment_Method: recordData.Payment_Status === 'paid' ? 'completed' : 'pending',
            Reference: recordData.Order_ID,
            Status: recordData.Payment_Status || 'pending',
            Created_By: user?.email || recordData.Created_By,
            Created_At: now
          });
          
          // Shipping expense if applicable
          if (recordData.Shipping_Cost && parseFloat(recordData.Shipping_Cost) > 0) {
            financialEntries.push({
              Transaction_ID: generateTransactionId(),
              Date: recordData.Date || now.split('T')[0],
              Type: 'expense',
              Category: 'shipping',
              Subcategory: 'fulfillment',
              Module_Source: 'Ecommerce',
              Module_Record_ID: recordData.Order_ID,
              Description: `Shipping for order ${recordData.Order_ID}`,
              Amount: -Math.abs(parseFloat(recordData.Shipping_Cost)),
              Payment_Method: 'completed',
              Reference: recordData.Order_ID,
              Status: 'completed',
              Created_By: user?.email || recordData.Created_By,
              Created_At: now
            });
          }
        }
        break;
        
      case 'Agro':
        if (recordData.Activity === 'sale' && recordData.Revenue > 0) {
          financialEntries.push({
            Transaction_ID: generateTransactionId(),
            Date: recordData.Date || now.split('T')[0],
            Type: 'revenue',
            Category: 'agro',
            Subcategory: recordData.Type === 'crop' ? 'crop_sales' : 'livestock_sales',
            Module_Source: 'Agro',
            Module_Record_ID: recordData.Record_ID,
            Description: `Sold ${recordData.Quantity} ${recordData.Crop_Name || recordData.Livestock_Type}`,
            Amount: parseFloat(recordData.Revenue),
            Payment_Method: 'completed',
            Reference: recordData.Record_ID,
            Status: 'completed',
            Created_By: user?.email || recordData.Created_By,
            Created_At: now
          });
        } else if (recordData.Total_Cost > 0) {
          financialEntries.push({
            Transaction_ID: generateTransactionId(),
            Date: recordData.Date || now.split('T')[0],
            Type: 'expense',
            Category: 'agro',
            Subcategory: recordData.Type === 'crop' ? 'farming_cost' : 'animal_husbandry',
            Module_Source: 'Agro',
            Module_Record_ID: recordData.Record_ID,
            Description: `${recordData.Activity}: ${recordData.Quantity} ${recordData.Crop_Name || recordData.Livestock_Type}`,
            Amount: -Math.abs(parseFloat(recordData.Total_Cost)),
            Payment_Method: 'completed',
            Reference: recordData.Record_ID,
            Status: 'completed',
            Created_By: user?.email || recordData.Created_By,
            Created_At: now
          });
        }
        break;
        
      case 'HR':
        if (recordData.Transaction_Type === 'salary' && recordData.Amount > 0) {
          financialEntries.push({
            Transaction_ID: generateTransactionId(),
            Date: recordData.Date || now.split('T')[0],
            Type: 'expense',
            Category: 'salary',
            Subcategory: 'payroll',
            Module_Source: 'HR',
            Module_Record_ID: recordData.Employee_ID,
            Description: `Salary: ${recordData.Name} - ${recordData.Period || 'monthly'}`,
            Amount: -Math.abs(parseFloat(recordData.Amount)),
            Payment_Method: 'completed',
            Reference: recordData.Employee_ID,
            Status: 'completed',
            Created_By: user?.email || recordData.Created_By,
            Created_At: now
          });
        }
        break;
    }
    
    // Append all financial entries
    const headers = getSheetHeadersBySheet(financeSheet);
    for (const entry of financialEntries) {
      const rowData = headers.map(header => entry[header] || '');
      financeSheet.appendRow(rowData);
      
      // Update source record with Financial_Link_ID
      if (entry.Module_Record_ID) {
        updateSourceRecordWithLink(workspace, moduleName, entry.Module_Record_ID, entry.Transaction_ID);
      }
    }
    
    // Recalculate dashboard metrics after sync
    recalculateDashboardMetrics(workspaceId);
    
    return { success: true, syncedCount: financialEntries.length, entries: financialEntries };
    
  } catch (error) {
    console.error('Auto-sync error:', error);
    return { success: false, message: error.message };
  }
}


function saveRecord(moduleName, data, sessionId) {
  try {
    if (!userHasAccess(moduleName, sessionId)) throw new Error("Access denied");
    
    const user = getUserFromSession(sessionId);
    
    // IMPORTANT: Get the user's workspace file, not the main sheet
    const workspace = getWorkspaceFile(sessionId);
    
    // Check if workspace exists and is different from main sheet
    if (!workspace || workspace.getId() === SpreadsheetApp.getActiveSpreadsheet().getId()) {
      console.log('No dedicated workspace found, using main sheet');
    }
    
    const sheet = workspace.getSheetByName(MODULES[moduleName].sheet);
    if (!sheet) {
      // Create sheet if it doesn't exist
      workspace.insertSheet(MODULES[moduleName].sheet);
      const newSheet = workspace.getSheetByName(MODULES[moduleName].sheet);
      addDefaultHeadersToModule(newSheet, moduleName);
      return saveRecord(moduleName, data, sessionId);
    }
    
    const headers = getSheetHeadersBySheet(sheet);
    
    // Add created by and timestamp
    data.Created_By = user.name || user.email;
    data.Created_At = new Date().toISOString();
    
    // Generate ID if needed
    if (!data.Transaction_ID && moduleName === 'Finance') {
      data.Transaction_ID = generateTransactionId();
    }
    if (!data.Deal_ID && moduleName === 'Sales') {
      data.Deal_ID = `DEAL-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    if (!data.Order_ID && moduleName === 'Ecommerce') {
      data.Order_ID = `ORD-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    if (!data.Record_ID && moduleName === 'Agro') {
      data.Record_ID = `AGRO-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    if (!data.Employee_ID && moduleName === 'HR') {
      data.Employee_ID = `EMP-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    
    // Save record to module sheet
    const rowData = headers.map(header => data[header] || "");
    sheet.appendRow(rowData);
    
    // Auto-sync to financial data (if applicable and not already finance)
    if (moduleName !== 'Finance') {
      const syncResult = autoSyncToFinancial(moduleName, data, workspace.getId(), sessionId);
      if (!syncResult.success) {
        console.warn('Auto-sync warning:', syncResult.message);
      }
    }
    
    // ===== TRIGGER AUTO-RESPONSE =====
    // Only trigger for Finance module transactions with email
    if (moduleName === 'Finance' && data.Email) {
      const transactionData = {
        Name: data.Name || data.Customer_Name || '',
        Email: data.Email,
        Amount: data.Amount,
        Category: data.Category,
        Description: data.Description,
        Reference: data.Reference,
        Date: data.Date,
        Transaction_ID: data.Transaction_ID
      };
      
      // Process auto-response asynchronously (don't block)
      try {
        processAutoResponse(transactionData, user.businessId, workspace.getId());
      } catch (e) {
        console.error('Auto-response error:', e);
      }
    }
    
    // SINGLE return statement at the end
    return { 
      success: true, 
      message: "Record saved successfully",
      recordId: data.Transaction_ID || data.Deal_ID || data.Order_ID || data.Record_ID || data.Employee_ID
    };
    
  } catch (error) {
    console.error('Save record error:', error);
    return { success: false, message: error.message };
  }
}

function updateRecord(moduleName, data, recordId, sessionId) {
  try {
    if (!userHasAccess(moduleName, sessionId)) throw new Error("Access denied");
    
    const user = getUserFromSession(sessionId);
    const workspace = getWorkspaceFile(sessionId);
    const sheet = workspace.getSheetByName(MODULES[moduleName].sheet);
    if (!sheet) return { success: false, message: "Sheet not found" };
    
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    
    // Find the ID column
    const idCol = headers.findIndex(h => h === 'Deal_ID' || h === 'Order_ID' || h === 'Record_ID' || h === 'Employee_ID');
    if (idCol === -1) return { success: false, message: "ID column not found" };
    
    // Find the row
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][idCol] === recordId) {
        rowIndex = i;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, message: "Record not found" };
    
    // Update each column
    for (const [key, value] of Object.entries(data)) {
      if (key === '_record_id') continue;
      const colIndex = headers.indexOf(key);
      if (colIndex !== -1) {
        sheet.getRange(rowIndex + 1, colIndex + 1).setValue(value);
      }
    }
    
    return { success: true, message: "Record updated successfully" };
    
  } catch (error) {
    console.error('Update record error:', error);
    return { success: false, message: error.message };
  }
}

function deleteRecord(moduleName, index, sessionId) {
  try {
    if (!userHasAccess(moduleName, sessionId)) throw new Error("Access denied");
    
    const workspace = getWorkspaceFile(sessionId);
    const sheet = workspace.getSheetByName(MODULES[moduleName].sheet);
    if (!sheet) return { success: false, message: "Sheet not found" };
    
    // +2 because row 1 is headers, and index is 0-based
    sheet.deleteRow(index + 2);
    
    return { success: true, message: "Record deleted successfully" };
  } catch (error) {
    console.error('Delete error:', error);
    return { success: false, message: error.message };
  }
}


function getEcommerceUsage(sessionId) {
  const user = getUserFromSession(sessionId);
  if (user.subscriptionTier !== 'free') {
    return { tier: 'tier2', unlimited: true };
  }
  
  const workspace = getWorkspaceFile(sessionId);
  const sheet = workspace.getSheetByName('Ecommerce_Data');
  
  if (!sheet) {
    return { tier: 'free', usage: { products: 0, maxProducts: 10, ordersThisMonth: 0, maxOrdersPerMonth: 50 } };
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const productCol = headers.indexOf('Product');
  const uniqueProducts = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][productCol]) uniqueProducts.add(data[i][productCol].toString());
  }
  
  const dateCol = headers.indexOf('Date');
  const currentMonth = new Date().getMonth();
  const currentYear = new Date().getFullYear();
  let ordersThisMonth = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][dateCol]) {
      const orderDate = new Date(data[i][dateCol]);
      if (orderDate.getMonth() === currentMonth && orderDate.getFullYear() === currentYear) {
        ordersThisMonth++;
      }
    }
  }
  
  return {
    tier: 'free',
    usage: {
      products: uniqueProducts.size,
      maxProducts: 10,
      ordersThisMonth: ordersThisMonth,
      maxOrdersPerMonth: 50,
      productsRemaining: Math.max(0, 10 - uniqueProducts.size),
      ordersRemaining: Math.max(0, 50 - ordersThisMonth)
    }
  };
}

// ==================== MODULE ACCESS CONTROL ====================

function getModuleAccess(licenseResult, moduleName) {
  if (!licenseResult || !licenseResult.valid) {
    return { hasAccess: false, isReadOnly: true, message: 'Demo Mode' };
  }
  
  const accessibleModules = licenseResult.accessibleModules || [];
  const isReadOnly = licenseResult.isReadOnly || false;
  
  if (!accessibleModules.includes(moduleName)) {
    return {
      hasAccess: false,
      isReadOnly: true,
      message: `${moduleName} requires ${licenseResult.tier === LICENSE_TIERS.STARTER ? 'Sovereign' : 'upgrade'} license`
    };
  }
  
  return {
    hasAccess: true,
    isReadOnly: isReadOnly,
    message: 'Access granted'
  };
}

function executeWithLicenseCheck(operation, moduleName, data, sessionId) {
  const user = getUserFromSession(sessionId);
  
  if (!user) {
    return {
      success: false,
      message: 'User not found. Please login again.'
    };
  }
  
  const userRole = user.role || 'staff';
  const isOwnerOrAdmin = userRole === 'owner' || userRole === 'admin';
  
  // ===== STAFF: Check if module is assigned =====
  if (!isOwnerOrAdmin) {
    const accessibleModules = user.accessibleModules || [];
    if (!accessibleModules.includes(moduleName)) {
      return {
        success: false,
        message: 'You do not have access to this module'
      };
    }
    // Staff are read-only by default
    if (operation === 'save' || operation === 'delete' || operation === 'update') {
      return {
        success: false,
        isReadOnly: true,
        message: 'You have view-only access to this module'
      };
    }
  }
  
  // ===== OWNER/ADMIN: Check demo mode =====
  if (user.isDemo === 'YES') {
    return {
      success: false,
      isReadOnly: true,
      message: 'This is a Read-Only Demo. Upgrade to Sovereign to save, edit, or delete records.'
    };
  }
  
  // Check module access based on subscription tier
  const tier = user.subscriptionTier || 'free';
  const allowedModules = tier === 'free' ? ['Finance', 'Ecommerce'] : 
                         ['Finance', 'Ecommerce', 'Sales', 'CRM', 'HR', 'Logistics', 'Tax', 'Agro', 'Productivity', 'POS', 'Attendance', 'Warehouse'];
  
  if (!allowedModules.includes(moduleName)) {
    return {
      success: false,
      message: `${moduleName} requires upgrade. Please upgrade to Sovereign plan to access this module.`
    };
  }
  
  // Check record limits for Free tier (500 records)
  if (tier === 'free') {
    const currentCount = getRecordCount(moduleName, sessionId);
    if (currentCount >= 500) {
      return {
        success: false,
        message: 'Free tier limit of 500 records reached. Upgrade to Sovereign for unlimited records.'
      };
    }
  }
  
  return executeOperation(operation, moduleName, data, sessionId);
}