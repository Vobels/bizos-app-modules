//utils.gs
// ==================== HELPER FUNCTIONS ==================

function userHasAccess(moduleName, sessionId) {
  const user = getUserFromSession(sessionId);
  if (!user) return false;
  
  const userRole = user.role || 'staff';
  const isOwnerOrAdmin = userRole === 'owner' || userRole === 'admin';
  
  // Owner/Admin: Check subscription tier
  if (isOwnerOrAdmin) {
    const module = MODULES[moduleName];
    return module ? module.tiers.includes(user.subscriptionTier) : false;
  }
  
  // Staff: Check assigned modules
  const accessibleModules = user.accessibleModules || ['Finance'];
  return accessibleModules.includes(moduleName);
}

function generateTransactionId() {
  const date = new Date();
  const dateStr = date.toISOString().slice(0, 10).replace(/-/g, '');
  const random = Math.random().toString(36).substring(2, 8).toUpperCase();
  return `TXN-${dateStr}-${random}`;
}

function getSheetHeaders(moduleName, sessionId) {
  try {
    if (!userHasAccess(moduleName, sessionId)) return [];
    const workspace = getWorkspaceFile(sessionId);
    const sheet = workspace.getSheetByName(MODULES[moduleName].sheet);
    if (!sheet) return [];
    return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  } catch (error) {
    console.error('Error getting headers:', error);
    return [];
  }
}


function getCategoriesByType(type) {
  if (type === 'Income') {
    return [
      'Sales Revenue', 'Service Income', 'Product Sales', 'Consulting',
      'Rental Income', 'Interest Income', 'Investment Income',
      'Commission', 'Refunds', 'Other Income'
    ];
  } else if (type === 'Expense') {
    return [
      'Operations', 'Marketing', 'Salaries', 'Rent & Utilities',
      'Software & Subscriptions', 'Legal & Compliance', 'Taxes',
      'Travel & Entertainment', 'Equipment', 'Inventory',
      'Shipping & Logistics', 'Maintenance', 'Insurance',
      'Bank Fees', 'Miscellaneous', 'Other Expense'
    ];
  }
  return [];
}

function getSubcategories(category) {
  const subcategories = {
    'Operations': ['Office Supplies', 'Printing', 'Postage', 'Cleaning', 'Security'],
    'Marketing': ['Advertising', 'Social Media', 'SEO', 'Content Creation', 'Email Marketing', 'Events'],
    'Salaries': ['Regular Pay', 'Bonuses', 'Benefits', 'Contractors', 'Overtime', 'Commissions'],
    'Rent & Utilities': ['Rent', 'Electricity', 'Water', 'Internet', 'Phone', 'Gas'],
    'Software & Subscriptions': ['SaaS', 'Cloud Services', 'Domain & Hosting', 'App Subscriptions'],
    'Legal & Compliance': ['Legal Fees', 'Licenses', 'Permits', 'Audit Fees'],
    'Taxes': ['Income Tax', 'Sales Tax', 'VAT', 'Property Tax', 'Payroll Tax'],
    'Travel & Entertainment': ['Flights', 'Hotels', 'Meals', 'Transportation', 'Client Entertainment'],
    'Equipment': ['Computers', 'Furniture', 'Machinery', 'Repairs', 'Maintenance'],
    'Inventory': ['Raw Materials', 'Finished Goods', 'Packaging', 'Supplies'],
    'Shipping & Logistics': ['Courier', 'Freight', 'Customs', 'Warehousing'],
    'Maintenance': ['Repairs', 'Servicing', 'Upkeep', 'Renovations'],
    'Insurance': ['Health Insurance', 'Business Insurance', 'Liability Insurance'],
    'Bank Fees': ['Transaction Fees', 'Monthly Fees', 'Wire Transfer Fees', 'Overdraft'],
    'Miscellaneous': ['Training', 'Charitable', 'Gifts', 'Staff Welfare'],
    'Other Expense': ['Custom Entry'],
    'Sales Revenue': ['Product Sales', 'Service Revenue', 'Subscription Revenue'],
    'Service Income': ['Consulting', 'Maintenance', 'Support'],
    'Product Sales': ['Retail', 'Wholesale', 'Online Sales'],
    'Consulting': ['Strategy', 'Implementation', 'Training'],
    'Other Income': ['Interest', 'Dividends', 'Grants']
  };
  
  return subcategories[category] || ['Other'];
}

function getAgroActivities(type) {
  if (type === 'Crop') return REFERENCE_DATA.Agro.cropActivities;
  if (type === 'Livestock') return REFERENCE_DATA.Agro.livestockActivities;
  return [];
}

// ==================== LICENSE HELPER FUNCTIONS ====================
// Add these to your utils.gs file

function getBusinessIdFromSession(sessionId) {
  const user = getUserFromSession(sessionId);
  return user ? user.businessId : null;
}

function getDeploymentIdFromSession(sessionId) {
  const user = getUserFromSession(sessionId);
  if (user && user.businessId) {
    const businessSheet = getOrCreateBusinessSheet();
    const data = businessSheet.getDataRange().getValues();
    const headers = data[0];
    const bizIdCol = headers.indexOf('Business_ID');
    const workspaceIdCol = headers.indexOf('Workspace_ID');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][bizIdCol] === user.businessId) {
        return data[i][workspaceIdCol] || 'main';
      }
    }
  }
  return 'main';
}

function getRecordCount(moduleName, sessionId) {
  try {
    const workspace = getWorkspaceFile(sessionId);
    const sheet = workspace.getSheetByName(MODULES[moduleName]?.sheet);
    if (!sheet) return 0;
    const lastRow = sheet.getLastRow();
    return lastRow > 1 ? lastRow - 1 : 0;
  } catch (error) {
    console.error('Error getting record count:', error);
    return 0;
  }
}

function executeOperation(operation, moduleName, data, sessionId) {
  switch(operation) {
    case 'save':
      return saveRecord(moduleName, data, sessionId);
    case 'delete':
      return deleteRecord(moduleName, data.index, sessionId);
    case 'update':
      return updateRecord(moduleName, data.updates, data.index, sessionId);
    default:
      return { success: false, message: 'Unknown operation' };
  }
}

// This is the CRITICAL missing function
function validateLicense(businessId, deploymentId, sessionId) {
  try {
    const user = getUserFromSession(sessionId);
    
    if (!user) {
      return {
        valid: true,
        tier: 'demo',
        isReadOnly: true,
        accessibleModules: ['Finance', 'Ecommerce', 'Sales', 'CRM', 'HR', 'Logistics', 'Tax', 'Agro'],
        maxRecords: 50,
        message: 'Demo Mode'
      };
    }
    
    // Map subscription tier to access
    let accessibleModules = ['Finance', 'Ecommerce'];
    let isReadOnly = false;
    let maxRecords = null;
    let tier = user.subscriptionTier || 'free';
    
    if (user.isDemo === 'YES') {
      isReadOnly = true;
      accessibleModules = ['Finance', 'Ecommerce', 'Sales', 'CRM', 'HR', 'Logistics', 'Tax', 'Agro'];
      maxRecords = 50;
    } else if (tier === 'free') {
      accessibleModules = ['Finance', 'Ecommerce'];
      maxRecords = 500;
    } else if (tier === 'tier2' || tier === 'professional') {
      accessibleModules = ['Finance', 'Ecommerce', 'Sales', 'CRM', 'HR', 'Logistics', 'Tax', 'Agro'];
      maxRecords = null;
    } else if (tier === 'enterprise') {
      accessibleModules = ['Finance', 'Ecommerce', 'Sales', 'CRM', 'HR', 'Logistics', 'Tax', 'Agro'];
      maxRecords = null;
    }
    
    return {
      valid: true,
      tier: tier,
      isReadOnly: isReadOnly,
      accessibleModules: accessibleModules,
      maxRecords: maxRecords,
      message: `${tier.toUpperCase()} License Active`
    };
    
  } catch (error) {
    console.error('validateLicense error:', error);
    return {
      valid: true,
      tier: 'demo',
      isReadOnly: true,
      accessibleModules: ['Finance', 'Ecommerce'],
      maxRecords: 50,
      message: 'Fallback Mode'
    };
  }
}

const LICENSE_TIERS = {
  DEMO: 'demo',
  STARTER: 'starter',
  SOVEREIGN: 'sovereign',
  ARCHITECT: 'architect'
};
