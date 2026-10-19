function getBizOSModuleDefinition_(moduleName) {
  var definitions = {
    Finance:{sheet:'Financial_Data',label:'Financial Management'},
    Sales:{sheet:'Sales_Data',label:'Sales Pipeline'},
    Ecommerce:{sheet:'Ecommerce_Data',label:'E-commerce'},
    CRM:{sheet:'CRM_Data',label:'Customer Relations'},
    HR:{sheet:'HR_Data',label:'Human Resources'},
    Logistics:{sheet:'Logistics_Data',label:'Logistics'},
    Tax:{sheet:'Tax_Data',label:'Tax Compliance'},
    Agro:{sheet:'Agro_Data',label:'Agriculture'},
    Productivity:{sheet:'Productivity_Data',label:'Productivity Suite'},
    POS:{sheet:'POS_Data',label:'Point of Sale'},
    Attendance:{sheet:'Attendance_Data',label:'Staff Attendance'},
    Warehouse:{sheet:'Warehouse_Data',label:'Warehouse Management'}
  };
  return definitions[String(moduleName || '')] || null;
}

// Resolve a module sheet consistently for legacy workspaces that predate
// the full module-sheet set. Creation is idempotent and uses canonical names.
function ensureModuleFinancialColumns_(sheet, moduleName) {
  if (!sheet || moduleName !== 'Logistics') return;
  const required = ['Date', 'Fuel_Cost', 'Repair_Cost', 'Toll_Cost', 'Transport_Income', 'Financial_Status', 'Financial_Link_ID'];
  const lastColumn = sheet.getLastColumn();
  const headers = lastColumn > 0 ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(h) { return String(h || '').trim(); }) : [];
  const missing = required.filter(function(header) { return headers.indexOf(header) < 0; });
  if (missing.length) sheet.getRange(1, lastColumn + 1, 1, missing.length).setValues([missing]);
}

// Append-only migration for the optional shared Business Center customer reference.
// Existing columns and record values are never reordered or rewritten.
function ensureModuleCustomerReferenceColumns_(sheet, moduleName) {
  if (!sheet || !['CRM', 'Sales', 'Ecommerce', 'Agro', 'POS'].includes(moduleName)) return;
  let lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    // Recover an existing but empty CRM sheet using the canonical headers first.
    addDefaultHeadersToModule(sheet, moduleName);
    lastColumn = sheet.getLastColumn();
  }
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  if (headers.indexOf('Customer_ID') === -1) {
    sheet.getRange(1, lastColumn + 1).setValue('Customer_ID');
  }
}

// Append the customer reference to Finance only when a source explicitly carries
// one. Existing finance columns and transaction rows are never reordered.
function ensureFinanceCustomerReferenceColumn_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  if (headers.indexOf('Customer_ID') === -1) {
    sheet.getRange(1, lastColumn + 1).setValue('Customer_ID');
  }
}

function getBusinessCenterCustomerById_(workspace, customerId) {
  const id = String(customerId || '').trim();
  if (!workspace || !id) return null;
  const sheet = workspace.getSheetByName('BusinessCenter_Customers');
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return null;
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(h) { return String(h || '').trim(); });
  const idCol = headers.indexOf('Customer_ID');
  if (idCol < 0) return null;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol] || '').trim() === id) {
      const customer = {};
      headers.forEach(function(header, j) { customer[header] = values[i][j]; });
      return customer;
    }
  }
  return null;
}

// Read registered customers for CRM, Sales, or Ecommerce forms. The workspace
// and module access are checked server-side; only minimal customer fields return.
function getBusinessCenterCustomerReferences(moduleName, sessionId) {
  try {
    const allowedModules = ['CRM', 'Sales', 'Ecommerce', 'Agro', 'POS'];
    if (!allowedModules.includes(moduleName)) return { success: false, message: 'Customer references are not enabled for this module.' };
    if (!userHasAccess(moduleName, sessionId)) return { success: false, message: 'Access denied' };
    const workspace = getWorkspaceFile(sessionId);
    if (!workspace) return { success: false, message: 'Workspace is unavailable.' };
    const sheet = workspace.getSheetByName('BusinessCenter_Customers');
    if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return { success: true, customers: [] };
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(function(h) { return String(h || '').trim(); });
    const idCol = headers.indexOf('Customer_ID');
    const nameCol = headers.indexOf('Name');
    const phoneCol = headers.indexOf('Phone');
    const emailCol = headers.indexOf('Email');
    if (idCol < 0) return { success: true, customers: [] };
    const customers = values.slice(1).map(function(row) {
      return {
        Customer_ID: String(row[idCol] || '').trim(),
        Name: nameCol < 0 ? '' : String(row[nameCol] || ''),
        Phone: phoneCol < 0 ? '' : String(row[phoneCol] || ''),
        Email: emailCol < 0 ? '' : String(row[emailCol] || '')
      };
    }).filter(function(customer) { return !!customer.Customer_ID; });
    return { success: true, customers: customers };
  } catch (error) {
    console.error('Business Center customer reference lookup failed:', error);
    return { success: false, message: error.message || 'Could not load customer references.' };
  }
}

// Backwards-compatible CRM lookup name for any older UI callers.
function getCRMCustomerReferences(sessionId) {
  return getBusinessCenterCustomerReferences('CRM', sessionId);
}

function getOrCreateModuleSheet_(workspace, moduleName) {
  if (!workspace) throw new Error('Workspace is unavailable.');
  const definition = getBizOSModuleDefinition_(moduleName);
  if (!definition) return null;

  let sheet = workspace.getSheetByName(definition.sheet);
  if (!sheet) {
    sheet = workspace.insertSheet(definition.sheet);
    addDefaultHeadersToModule(sheet, moduleName);
  }
  // Append-only schema migrations preserve existing workspace data.
  ensureModuleFinancialColumns_(sheet, moduleName);
  ensureModuleCustomerReferenceColumns_(sheet, moduleName);
  return sheet;
}

// modules.gs
function getModuleData(moduleName, sessionId) {
  try {
    if (!userHasAccess(moduleName, sessionId)) {
      throw new Error("Access denied");
    }
    
    const workspace = getWorkspaceFile(sessionId);
    const definition = getBizOSModuleDefinition_(moduleName);
    if (!definition) throw new Error("Module not found");
    const sheet = getOrCreateModuleSheet_(workspace, moduleName);
    if (!sheet) throw new Error("Module not found");
    
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
    if (!userHasAccess(moduleName, sessionId)) {
      // Temporary safe diagnostics for the shared record-list failure.
      // Never log session IDs, email addresses, or credentials.
      let diagnosticUser = null;
      let sessionValid = false;
      try {
        if (typeof validateBizOSSession === "function") {
          const validation = validateBizOSSession(sessionId);
          sessionValid = !!(validation && validation.success);
          diagnosticUser = sessionValid ? validation.user : null;
        } else {
          diagnosticUser = getUserFromSession(sessionId);
          sessionValid = !!diagnosticUser;
        }
      } catch (diagnosticError) {
        console.warn("Module access diagnostic: session resolution threw", String(diagnosticError && diagnosticError.message || diagnosticError));
      }

      const role = String(diagnosticUser && diagnosticUser.role || "unknown").trim().toLowerCase();
      const tier = String(diagnosticUser && diagnosticUser.subscriptionTier || "unknown").trim().toLowerCase();
      const ownerOrAdmin = role === "owner" || role === "admin";
      let allowedModules = [];
      try {
        allowedModules = ownerOrAdmin
          ? (getAllowedModulesForTier_(tier) || [])
          : ((diagnosticUser && Array.isArray(diagnosticUser.accessibleModules))
              ? diagnosticUser.accessibleModules
              : ["Finance"]);
      } catch (diagnosticError) {
        console.warn("Module access diagnostic: could not resolve allowed modules");
      }

      console.warn("Module data authorization denied", JSON.stringify({
        requestedModule: String(moduleName || ""),
        sessionValid: sessionValid,
        role: role,
        tier: tier,
        ownerOrAdmin: ownerOrAdmin,
        requestedModuleAllowed: Array.isArray(allowedModules) && allowedModules.includes(moduleName),
        allowedModules: Array.isArray(allowedModules) ? allowedModules : []
      }));
      throw new Error("Access denied");
    }
    
    // userHasAccess() already validates the session. Avoid a second
    // user lookup here before resolving the workspace.
    const workspace = getWorkspaceFile(sessionId);
    const definition = getBizOSModuleDefinition_(moduleName);
    if (!definition) return { success: false, message: "Module not found" };
    const sheet = getOrCreateModuleSheet_(workspace, moduleName);
    if (!sheet) return { success: false, message: "Module not found" };
    
    const lastRow = sheet.getLastRow();
    const headers = getSheetHeadersBySheet(sheet);
    if (lastRow <= 1) return { success: true, headers: headers, records: [], financialSummary: { totalRevenue: 0, totalExpenses: 0 }, recordCount: 0 };
    
    // Keep the latest 1000 data rows while always reading headers from row 1.
    const startRow = Math.max(2, lastRow - 999);
    const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn()).getValues();
    
    const records = values.map((row, offset) => {
      const record = {};
      headers.forEach((h, i) => {
        const value = row[i];
        // google.script.run cannot serialize Date objects in returned data.
        // Convert spreadsheet dates to stable strings before returning records.
        if (value instanceof Date && !isNaN(value.getTime())) {
          const timePart = Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm:ss');
          record[h] = timePart === '00:00:00'
            ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd')
            : Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
        } else {
          record[h] = value;
        }
      });
      record._row = startRow + offset;
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
      headers: headers,
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
  // Resolve the same session-scoped user/business record used by server-side
  // authorization. Do not independently select the first Users row by email.
  const sessionValidation = typeof validateBizOSSession === "function"
    ? validateBizOSSession(sessionId)
    : null;
  const sessionUser = sessionValidation
    ? (sessionValidation.success ? sessionValidation.user : null)
    : getUserFromSession(sessionId);
  if (!sessionUser) {
    console.warn("Module summary denied: no valid session user.");
    return {};
  }

  const businessId = String(sessionUser.businessId || "");
  const subscriptionTier = normalizeSubscriptionTier(sessionUser.subscriptionTier);
  const userRole = String(sessionUser.role || "staff").toLowerCase();
  const isDemo = sessionUser.isDemo === true || sessionUser.isDemo === "YES";
  const email = sessionUser.email;
  const cache = CacheService.getScriptCache();

  // Include the resolved business and tier so a changed entitlement cannot
  // reuse a summary cached for a different business or plan. v3 bypasses old
  // v2 entries created by the previous email-only lookup.
  const cacheKey = `module_summary_v3_${sessionId}_${businessId}_${subscriptionTier}`;
  const cached = cache.get(cacheKey);
  if (cached) {
    console.log("Returning cached module summary for business:", businessId, "tier:", subscriptionTier);
    return JSON.parse(cached);
  }

  try {
    // Reuse the canonical workspace resolver. Paid accounts must never have
    // their module counts read from the master spreadsheet as a fallback.
    const workspace = getWorkspaceFile(sessionId);
    let staffModules = [];

    if (userRole !== "owner" && userRole !== "admin") {
      staffModules = getStaffModules(email, businessId);
      console.log("Staff modules for", email, ":", staffModules);
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
      } else if (normalizeSubscriptionTier(subscriptionTier) === 'free') {
        // Free/Starter: Finance + Ecommerce are the core BizOS apps
        accessibleModules = ['Finance', 'Ecommerce'];
      } else {
        // Sovereign/Enterprise: All modules
        accessibleModules = Object.keys(allModules);
      }
    } else {
      // Staff/Manager/Viewer: Only assigned modules
      accessibleModules = staffModules.length > 0 ? staffModules : [];
      // No implicit module access for unassigned staff.
      if (accessibleModules.length === 0) {
        accessibleModules = [];
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
    if (moduleName === 'BusinessCenter' || String(recordData && recordData.Customer_ID || '').trim()) {
      ensureFinanceCustomerReferenceColumn_(financeSheet);
    }

    const user = getUserFromSession(sessionId);
    const now = new Date().toISOString();
    
    let financialEntries = [];
    
    // Determine financial impact based on module and record type
    switch(moduleName) {
      case 'Sales':
        if (String(recordData.Stage || '').trim().toLowerCase().replace(/[\\s-]+/g, '_') === 'closed_won' && parseFloat(recordData.Total_Amount) > 0) {
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
        const ecommercePaymentStatus = String(recordData.Payment_Status || '').toLowerCase();
        const ecommerceOrderStatus = String(recordData.Order_Status || recordData.Status || '').toLowerCase();
        const ecommerceRevenueRecognized = ['paid', 'completed', 'success', 'succeeded'].includes(ecommercePaymentStatus) ||
          ['delivered', 'completed', 'fulfilled'].includes(ecommerceOrderStatus);
        if (parseFloat(recordData.Total_Amount) > 0 && ecommerceRevenueRecognized) {
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
        
      case 'Logistics': {
        const logisticsStatus = String(recordData.Financial_Status || '').trim().toLowerCase();
        const logisticsSettled = ['paid', 'received', 'completed', 'settled'].includes(logisticsStatus);
        const fuelCost = Math.max(0, parseFloat(recordData.Fuel_Cost) || 0);
        const repairCost = Math.max(0, parseFloat(recordData.Repair_Cost) || 0);
        const tollCost = Math.max(0, parseFloat(recordData.Toll_Cost) || 0);
        const logisticsExpense = fuelCost + repairCost + tollCost;
        const transportIncome = Math.max(0, parseFloat(recordData.Transport_Income) || 0);
        if (logisticsSettled && logisticsExpense > 0) {
          financialEntries.push({
            Transaction_ID: generateTransactionId(), Date: recordData.Date || now.split('T')[0],
            Type: 'expense', Category: 'logistics', Subcategory: 'transport_operating_costs',
            Module_Source: 'Logistics', Module_Record_ID: recordData.Shipment_ID,
            Description: 'Logistics costs for shipment ' + (recordData.Shipment_ID || '') +
              ' (fuel ' + fuelCost + ', repairs ' + repairCost + ', tolls ' + tollCost + ')',
            Amount: -Math.abs(logisticsExpense), Payment_Method: 'completed',
            Reference: recordData.Shipment_ID, Status: 'completed',
            Created_By: user?.email || recordData.Created_By, Created_At: now
          });
        }
        if (logisticsSettled && transportIncome > 0) {
          financialEntries.push({
            Transaction_ID: generateTransactionId(), Date: recordData.Date || now.split('T')[0],
            Type: 'revenue', Category: 'logistics', Subcategory: 'transport_income',
            Module_Source: 'Logistics', Module_Record_ID: recordData.Shipment_ID,
            Description: 'Transport income for shipment ' + (recordData.Shipment_ID || ''),
            Amount: transportIncome, Payment_Method: 'completed',
            Reference: recordData.Shipment_ID, Status: 'completed',
            Created_By: user?.email || recordData.Created_By, Created_At: now
          });
        }
        break;
      }

      case 'Agro':
        if (String(recordData.Activity || '').trim().toLowerCase() === 'sale' && parseFloat(recordData.Revenue) > 0) {
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
        
      case 'POS':
        const posStatus = String(recordData.Status || 'completed').toLowerCase();
        const posIsReversed = ['cancelled', 'canceled', 'refunded', 'refund', 'void', 'voided', 'failed', 'reversed'].includes(posStatus);
        if (parseFloat(recordData.Total_Amount) > 0 && !posIsReversed) {
          financialEntries.push({
            Transaction_ID: generateTransactionId(),
            Date: recordData.Date || now.split('T')[0],
            Type: 'revenue',
            Category: 'pos',
            Subcategory: 'point_of_sale',
            Module_Source: 'POS',
            Module_Record_ID: recordData.Transaction_ID,
            Description: 'POS sale ' + recordData.Transaction_ID + ': ' + (recordData.Product_Name || ''),
            Amount: parseFloat(recordData.Total_Amount),
            Payment_Method: recordData.Payment_Method || 'completed',
            Reference: recordData.Transaction_ID,
            Status: recordData.Status || 'completed',
            Created_By: user?.email || recordData.Created_By,
            Created_At: now
          });
        }
        break;

      case 'Tax':
        if (String(recordData.Status || '').toLowerCase() === 'paid' && parseFloat(recordData.Tax_Amount) > 0 && recordData.Tax_ID) {
          financialEntries.push({
            Transaction_ID: generateTransactionId(),
            Date: recordData.Date || now.split('T')[0],
            Type: 'expense',
            Category: 'tax',
            Subcategory: recordData.Type || 'tax_payment',
            Module_Source: 'Tax',
            Module_Record_ID: recordData.Tax_ID,
            Description: 'Tax payment: ' + (recordData.Type || recordData.Reference || recordData.Tax_ID),
            Amount: -Math.abs(parseFloat(recordData.Tax_Amount)),
            Payment_Method: recordData.Payment_Method || 'completed',
            Reference: recordData.Reference || recordData.Tax_ID,
            Status: 'completed',
            Created_By: user?.email || recordData.Created_By,
            Created_At: now
          });
        }
        break;

      case 'BusinessCenterExpense':
        if (parseFloat(recordData.Total) > 0 && recordData.Transaction_ID) {
          financialEntries.push({
            Transaction_ID: generateTransactionId(),
            Date: recordData.Date || now.split('T')[0],
            Type: 'expense',
            Category: recordData.Category || 'business_center',
            Subcategory: 'business_center_expense',
            Module_Source: 'BusinessCenterExpense',
            Module_Record_ID: recordData.Transaction_ID,
            Description: recordData.Description || 'Business Center expense',
            Amount: -Math.abs(parseFloat(recordData.Total)),
            Payment_Method: recordData.Payment_Method || 'completed',
            Reference: recordData.Transaction_ID,
            Status: 'completed',
            Created_By: user?.email || recordData.Created_By,
            Created_At: now
          });
        }
        break;

      case 'BusinessCenter':
        if (parseFloat(recordData.Total) > 0 && recordData.Sale_ID) {
          financialEntries.push({
            Transaction_ID: generateTransactionId(),
            Date: recordData.Date || now.split('T')[0],
            Type: 'revenue',
            Category: 'business_center',
            Subcategory: 'point_of_sale',
            Module_Source: 'BusinessCenter',
            Module_Record_ID: recordData.Sale_ID,
            Description: 'Business Center sale ' + recordData.Sale_ID,
            Amount: parseFloat(recordData.Total),
            Payment_Method: recordData.Payment_Method || 'pending',
            Reference: recordData.Sale_ID,
            Status: recordData.Status || 'completed',
            Created_By: user?.email || recordData.Created_By,
            Created_At: now
          });
        }
        break;
        
      // HR records contain employee profile data, not payroll transactions.
      // Payroll should only be synced from an explicit payroll transaction source.
    }
    
    // Carry explicit shared-customer identity to linked ledger entries. Do not
    // infer or backfill customer identity from names/emails on historical rows.
    const linkedCustomerId = String(recordData && recordData.Customer_ID || '').trim();
    if (linkedCustomerId) {
      financialEntries.forEach(function(entry) { entry.Customer_ID = linkedCustomerId; });
    }

    // Reconcile existing linked financial events for this source record.
    // A source edit can remove an event (for example, a Sales deal changing
    // from closed_won to closed_lost, or an Ecommerce shipping cost becoming 0).
    // Only rows explicitly linked to this module record are touched.
    const sourceRecordId = recordData.Deal_ID || recordData.Order_ID || recordData.Sale_ID || recordData.Record_ID || recordData.Transaction_ID || recordData.Tax_ID || recordData.Shipment_ID;
    const headers = getSheetHeadersBySheet(financeSheet);
    if (sourceRecordId) {
      const sourceCol = headers.indexOf('Module_Source');
      const sourceRecordCol = headers.indexOf('Module_Record_ID');
      const typeCol = headers.indexOf('Type');

      if (sourceCol >= 0 && sourceRecordCol >= 0 && typeCol >= 0) {
        const desiredTypes = {};
        financialEntries.forEach(entry => { desiredTypes[String(entry.Type || '').toLowerCase()] = true; });

        const currentValues = financeSheet.getDataRange().getValues();
        const keptTypes = {};

        for (let i = 1; i < currentValues.length; i++) {
          const sameSource = String(currentValues[i][sourceCol] || '') === moduleName;
          const sameRecord = String(currentValues[i][sourceRecordCol] || '') === String(sourceRecordId);
          if (!sameSource || !sameRecord) continue;

          const type = String(currentValues[i][typeCol] || '').toLowerCase();
          const statusCol = headers.indexOf('Status');
          const amountCol = headers.indexOf('Amount');

          // Keep the ledger audit trail. Superseded or duplicate source-linked
          // events are voided in place instead of being physically deleted.
          if (!desiredTypes[type] || keptTypes[type]) {
            if (type !== 'void' && type !== 'voided') {
              financeSheet.getRange(i + 1, typeCol + 1).setValue('void');
              if (statusCol >= 0) financeSheet.getRange(i + 1, statusCol + 1).setValue('voided');
              if (amountCol >= 0) financeSheet.getRange(i + 1, amountCol + 1).setValue(0);
            }
          } else {
            keptTypes[type] = true;
          }
        }
      }
    }

    // Reload after reconciliation so row indexes remain accurate.
    const financeValues = financeSheet.getDataRange().getValues();
    const sourceCol = headers.indexOf('Module_Source');
    const sourceRecordCol = headers.indexOf('Module_Record_ID');
    const typeCol = headers.indexOf('Type');
    const transactionIdCol = headers.indexOf('Transaction_ID');

    for (const entry of financialEntries) {
      let existingRow = -1;
      if (sourceCol >= 0 && sourceRecordCol >= 0 && typeCol >= 0 && entry.Module_Record_ID) {
        for (let i = 1; i < financeValues.length; i++) {
          if (String(financeValues[i][sourceCol] || '') === moduleName &&
              String(financeValues[i][sourceRecordCol] || '') === String(entry.Module_Record_ID) &&
              String(financeValues[i][typeCol] || '').toLowerCase() === String(entry.Type || '').toLowerCase()) {
            existingRow = i + 1;
            if (transactionIdCol >= 0 && financeValues[i][transactionIdCol]) {
              entry.Transaction_ID = financeValues[i][transactionIdCol];
            }
            break;
          }
        }
      }

      const rowData = headers.map((header, index) => {
        if (entry[header] !== undefined) return entry[header];
        return existingRow > 0 ? financeValues[existingRow - 1][index] : '';
      });
      if (existingRow > 0) {
        financeSheet.getRange(existingRow, 1, 1, headers.length).setValues([rowData]);
      } else {
        financeSheet.appendRow(rowData);
      }
      
      // Update source record with Financial_Link_ID
      if (entry.Module_Record_ID && typeof updateSourceRecordWithLink === 'function') {
        updateSourceRecordWithLink(workspace, moduleName, entry.Module_Record_ID, entry.Transaction_ID);
      }
    }
    
    // Invalidate the short-lived per-session KPI cache after ledger changes.
    try {
      if (sessionId) CacheService.getScriptCache().remove('dashboard_fast_' + sessionId);
    } catch (cacheError) {
      console.warn('Dashboard KPI cache invalidation warning:', cacheError.message);
    }

    // Recalculate dashboard metrics after sync
    recalculateDashboardMetrics(workspaceId);
    
    return { success: true, syncedCount: financialEntries.length, entries: financialEntries };
    
  } catch (error) {
    console.error('Auto-sync error:', error);
    return { success: false, message: error.message };
  }
}


function assertModuleWriteAccess_(moduleName, sessionId, operation) {
  if (!userHasAccess(moduleName, sessionId)) throw new Error("Access denied");
  const user = getUserFromSession(sessionId);
  if (!user) throw new Error("Session expired. Please login again.");
  const role = String(user.role || "staff").toLowerCase();
  const isPrivileged = role === "owner" || role === "admin";
  const isDemo = user.isDemo === true || String(user.isDemo || "").toUpperCase() === "YES";
  if (isDemo) throw new Error("Demo accounts are read-only.");
  // Staff may delete records only in modules assigned to them.
  // Creating and editing remain owner/admin actions; demo accounts stay read-only.
  if (!isPrivileged && !(role === "staff" && operation === "delete")) {
    throw new Error(operation === "delete" ? "You do not have permission to delete records." : "This module is read-only for your account.");
  }
  return user;
}

function saveRecord(moduleName, data, sessionId) {
  try {
    const user = assertModuleWriteAccess_(moduleName, sessionId, "write");
    
    // IMPORTANT: Get the user's workspace file, not the main sheet
    const workspace = getWorkspaceFile(sessionId);
    const definition = getBizOSModuleDefinition_(moduleName);
    if (!definition) return { success: false, message: "Module not found" };
    
    // Do not call getActiveSpreadsheet() from a web request: it may be null
    // outside an interactive spreadsheet UI. getWorkspaceFile() already
    // resolves the correct master or dedicated workspace.
    if (!workspace) {
      throw new Error('Workspace is unavailable.');
    }
    
    const sheet = workspace.getSheetByName(definition.sheet);
    if (sheet) {
      ensureModuleFinancialColumns_(sheet, moduleName);
      ensureModuleCustomerReferenceColumns_(sheet, moduleName);
    }
    if (!sheet) {
      // Create sheet if it doesn't exist
      workspace.insertSheet(definition.sheet);
      const newSheet = workspace.getSheetByName(definition.sheet);
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
    if (!data.Transaction_ID && moduleName === 'POS') {
      data.Transaction_ID = `POS-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    if (!data.Deal_ID && moduleName === 'Sales') {
      data.Deal_ID = `DEAL-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    if (!data.Order_ID && moduleName === 'Ecommerce') {
      data.Order_ID = `ORD-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    if (!data.Contact_ID && moduleName === 'CRM') {
      data.Contact_ID = `CONT-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    if (!data.Record_ID && moduleName === 'Agro') {
      data.Record_ID = `AGRO-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    if (!data.Employee_ID && moduleName === 'HR') {
      data.Employee_ID = `EMP-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    if (!data.Shipment_ID && moduleName === 'Logistics') {
      data.Shipment_ID = `SHIP-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    if (!data.Tax_ID && moduleName === 'Tax') {
      data.Tax_ID = `TAX-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    if (!data.Record_ID && moduleName === 'Productivity') {
      data.Record_ID = `TASK-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    if (!data.Record_ID && moduleName === 'Attendance') {
      data.Record_ID = `ATT-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    if (!data.Transaction_ID && moduleName === 'POS') {
      data.Transaction_ID = `POS-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    if (!data.Item_ID && moduleName === 'Warehouse') {
      data.Item_ID = `ITEM-${Utilities.getUuid().substring(0, 8).toUpperCase()}`;
    }
    
    // CRM may optionally link a contact to a registered Business Center customer.
    // Never infer identity from name/email/phone during save; the explicit ID is authoritative.
    if (['CRM', 'Sales', 'Ecommerce', 'Agro', 'POS'].includes(moduleName)) {
      data.Customer_ID = String(data.Customer_ID || '').trim();
      if (data.Customer_ID) {
        const customer = getBusinessCenterCustomerById_(workspace, data.Customer_ID);
        if (!customer) {
          return { success: false, message: 'The selected Business Center customer could not be found. Refresh the customer list and try again.' };
        }
        if (moduleName === 'CRM') {
          data.Name = String(customer.Name || '');
          data.Email = String(customer.Email || '');
          data.Phone = String(customer.Phone || '');
        } else if (['Sales', 'Agro', 'POS'].includes(moduleName)) {
          data.Customer_Name = String(customer.Name || '');
          if (['Sales', 'POS'].includes(moduleName)) data.Customer_Email = String(customer.Email || '');
        } else if (moduleName === 'Ecommerce') {
          data.Customer_Email = String(customer.Email || '');
        }
      }
    }

    // Save record to module sheet
    const rowData = headers.map(header => data[header] || "");
    sheet.appendRow(rowData);
    const savedRecord = {};
    headers.forEach(function(header, index) { savedRecord[header] = rowData[index]; });
    const savedRecordId = data.Transaction_ID || data.Deal_ID || data.Order_ID || data.Contact_ID || data.Employee_ID || data.Shipment_ID || data.Tax_ID || data.Item_ID || data.Record_ID || '';
    auditRecordChange_(workspace, sessionId, moduleName, 'CREATE', savedRecordId, {}, savedRecord);
    
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
      recordId: data.Transaction_ID || data.Deal_ID || data.Order_ID || data.Contact_ID || data.Employee_ID || data.Shipment_ID || data.Tax_ID || data.Item_ID || data.Record_ID
    };
    
  } catch (error) {
    console.error('Save record error:', error);
    return { success: false, message: error.message };
  }
}

function updateRecord(moduleName, data, recordId, sessionId) {
  try {
    const user = assertModuleWriteAccess_(moduleName, sessionId, "write");
    const workspace = getWorkspaceFile(sessionId);
    const definition = getBizOSModuleDefinition_(moduleName);
    if (!definition) return { success: false, message: "Module not found" };
    const sheet = workspace.getSheetByName(definition.sheet);
    if (!sheet) return { success: false, message: "Sheet not found" };
    ensureModuleFinancialColumns_(sheet, moduleName);
    ensureModuleCustomerReferenceColumns_(sheet, moduleName);
    
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    
    // Prefer the physical row captured when the user clicked Edit. This
    // targets the exact record even if older data already contains duplicate IDs.
    // Legacy callers may continue to pass a scalar record ID.
    let rowIndex = -1;
    const requestedRow = recordId && typeof recordId === "object"
      ? Number(recordId.rowNumber || recordId._row || 0)
      : 0;
    if (Number.isInteger(requestedRow) && requestedRow >= 2 && requestedRow <= sheet.getLastRow()) {
      rowIndex = requestedRow - 1; // values includes the header at index 0
    } else {
      const idCol = headers.findIndex(h => h === 'Deal_ID' || h === 'Order_ID' || h === 'Contact_ID' || h === 'Record_ID' || h === 'Employee_ID' || h === 'Transaction_ID' || h === 'Shipment_ID' || h === 'Tax_ID' || h === 'Item_ID');
      if (idCol === -1) return { success: false, message: "ID column not found" };
      for (let i = 1; i < values.length; i++) {
        if (values[i][idCol] === recordId) {
          rowIndex = i;
          break;
        }
      }
    }

    if (rowIndex === -1) return { success: false, message: "Record not found" };
    
    // Validate explicit shared-customer links before updating any cells.
    if (['CRM', 'Sales', 'Ecommerce', 'Agro', 'POS'].includes(moduleName) &&
        Object.prototype.hasOwnProperty.call(data, 'Customer_ID')) {
      data.Customer_ID = String(data.Customer_ID || '').trim();
      if (data.Customer_ID) {
        const customer = getBusinessCenterCustomerById_(workspace, data.Customer_ID);
        if (!customer) {
          return { success: false, message: 'The selected Business Center customer could not be found. Refresh the customer list and try again.' };
        }
        if (moduleName === 'CRM') {
          data.Name = String(customer.Name || '');
          data.Email = String(customer.Email || '');
          data.Phone = String(customer.Phone || '');
        } else if (['Sales', 'Agro', 'POS'].includes(moduleName)) {
          data.Customer_Name = String(customer.Name || '');
          if (['Sales', 'POS'].includes(moduleName)) data.Customer_Email = String(customer.Email || '');
        } else if (moduleName === 'Ecommerce') {
          data.Customer_Email = String(customer.Email || '');
        }
      }
    }

    // Capture the full previous row for the audit trail before applying edits.
    const beforeRecord = {};
    headers.forEach(function(header, i) { beforeRecord[header] = values[rowIndex][i]; });

    // Update each column. The audit actor is derived from the authenticated session,
    // never from client-supplied Created_By / Updated_By fields.
    for (const [key, value] of Object.entries(data)) {
      if (key === '_record_id' || key === 'Created_By' || key === 'Created_At' ||
          key === 'Updated_By' || key === 'Updated_At') continue;
      const colIndex = headers.indexOf(key);
      if (colIndex !== -1) {
        sheet.getRange(rowIndex + 1, colIndex + 1).setValue(value);
      }
    }
    const updatedValuesForAudit = sheet.getRange(rowIndex + 1, 1, 1, headers.length).getValues()[0];
    const afterRecord = {};
    headers.forEach(function(header, i) { afterRecord[header] = updatedValuesForAudit[i]; });
    const auditRecordId = beforeRecord.Transaction_ID || beforeRecord.Deal_ID || beforeRecord.Order_ID ||
      beforeRecord.Contact_ID || beforeRecord.Employee_ID || beforeRecord.Shipment_ID ||
      beforeRecord.Tax_ID || beforeRecord.Item_ID || beforeRecord.Record_ID || recordId;
    auditRecordChange_(workspace, sessionId, moduleName, 'UPDATE', auditRecordId, beforeRecord, afterRecord);

    // Keep linked financial events in sync with the edited source record.
    // This only applies to modules that already have an authoritative financial
    // sync path; ordinary Finance entries and unrelated module records are untouched.
    if (moduleName !== 'Finance' &&
        ['Sales', 'Ecommerce', 'Agro', 'POS', 'Tax', 'Logistics'].includes(moduleName)) {
      const updatedValues = sheet.getRange(rowIndex + 1, 1, 1, headers.length).getValues()[0];
      const updatedRecord = {};
      headers.forEach((header, i) => { updatedRecord[header] = updatedValues[i]; });
      const syncResult = autoSyncToFinancial(moduleName, updatedRecord, workspace.getId(), sessionId);
      if (!syncResult.success) {
        console.warn('Financial sync warning after update:', syncResult.message);
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
    assertModuleWriteAccess_(moduleName, sessionId, "delete");
    
    const workspace = getWorkspaceFile(sessionId);
    const definition = getBizOSModuleDefinition_(moduleName);
    if (!definition) return { success: false, message: "Module not found" };
    const sheet = workspace.getSheetByName(definition.sheet);
    if (!sheet) return { success: false, message: "Sheet not found" };
    
    // +2 because row 1 is headers, and index is 0-based.
    // Capture the record before removal so the audit log retains its context.
    const rowNumber = Number(index) + 2;
    if (!Number.isInteger(Number(index)) || Number(index) < 0 || rowNumber > sheet.getLastRow()) {
      return { success: false, message: "Record not found" };
    }
    const headers = getSheetHeadersBySheet(sheet);
    const rowValues = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
    const beforeRecord = {};
    headers.forEach(function(header, i) { beforeRecord[header] = rowValues[i]; });
    const recordId = beforeRecord.Transaction_ID || beforeRecord.Deal_ID || beforeRecord.Order_ID ||
      beforeRecord.Contact_ID || beforeRecord.Employee_ID || beforeRecord.Shipment_ID ||
      beforeRecord.Tax_ID || beforeRecord.Item_ID || beforeRecord.Record_ID || '';
    auditRecordChange_(workspace, sessionId, moduleName, 'DELETE', recordId, beforeRecord, {});
    sheet.deleteRow(rowNumber);

    return { success: true, message: "Record deleted successfully" };
  } catch (error) {
    console.error('Delete error:', error);
    return { success: false, message: error.message };
  }
}


function getEcommerceUsage(sessionId) {
  const user = getUserFromSession(sessionId);
  if (normalizeSubscriptionTier(user.subscriptionTier) === 'sovereign') {
    return { tier: 'sovereign', unlimited: true };
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

function getLicenseModuleAccess_(licenseResult, moduleName) {
  if (!licenseResult || !licenseResult.valid) {
    return { hasAccess: false, isReadOnly: true, message: 'Demo Mode' };
  }
  
  const accessibleModules = licenseResult.accessibleModules || [];
  const isReadOnly = licenseResult.isReadOnly || false;
  
  if (!accessibleModules.includes(moduleName)) {
    return {
      hasAccess: false,
      isReadOnly: true,
      message: `${moduleName} requires a Sovereign plan`
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
  
  const userRole = String(user.role || 'staff').toLowerCase();
  const isOwnerOrAdmin = userRole === 'owner' || userRole === 'admin';
  const isDemo = user.isDemo === true || String(user.isDemo || '').toUpperCase() === 'YES';

  // Demo accounts are read-only regardless of role or assigned modules.
  if (isDemo) {
    return {
      success: false,
      isReadOnly: true,
      message: 'This is a Read-Only Demo. You cannot change or delete records.'
    };
  }

  // Non-owner roles must have explicit module assignment. Staff may delete
  // assigned-module records, but creating and editing remain owner/admin actions.
  if (!isOwnerOrAdmin) {
    const accessibleModules = Array.isArray(user.accessibleModules) ? user.accessibleModules : [];
    if (!accessibleModules.includes(moduleName)) {
      return {
        success: false,
        message: 'You do not have access to this module'
      };
    }
    if (!(userRole === 'staff' && operation === 'delete')) {
      return {
        success: false,
        isReadOnly: true,
        message: operation === 'delete'
          ? 'You do not have permission to delete records.'
          : 'You have view-only access to this module'
      };
    }
  }
  
  // Check module access based on subscription tier
  const tier = normalizeSubscriptionTier(user.subscriptionTier);
  const allowedModules = getAllowedModulesForTier_(tier);
  
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