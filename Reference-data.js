// ============================================
// REFERENCE-DATA.GS
// All dropdown options, categories, and validation rules
// ============================================

const REFERENCE_DATA = {
  // ==================== FINANCE MODULE ====================
  Finance: {
    transactionTypes: ['Income', 'Expense'],  // Using "Income" instead of "revenue" for users
    
    categories: {
      'Income': [
        'Sales Revenue', 'Service Income', 'Product Sales', 'Consulting', 
        'Rental Income', 'Interest Income', 'Investment Income', 
        'Commission', 'Refunds', 'Other Income'
      ],
      'Expense': [
        'Operations', 'Marketing', 'Salaries', 'Rent & Utilities',
        'Software & Subscriptions', 'Legal & Compliance', 'Taxes',
        'Travel & Entertainment', 'Equipment', 'Inventory',
        'Shipping & Logistics', 'Maintenance', 'Insurance',
        'Bank Fees', 'Miscellaneous', 'Other Expense'
      ]
    },
    
    subcategories: {
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
      'Other Expense': ['Custom Entry']
    },
    
    paymentMethods: ['Cash', 'Bank Transfer', 'Credit Card', 'Debit Card', 'Mobile Money', 'Cheque', 'PayPal', 'Cryptocurrency', 'Other'],
    paymentStatus: ['Pending', 'Completed', 'Failed', 'Refunded'],
    recurring: ['One-time', 'Daily', 'Weekly', 'Monthly', 'Quarterly', 'Yearly']
  },

  // ==================== SALES MODULE ====================
  Sales: {
    stages: ['Lead', 'Qualified', 'Proposal Sent', 'Negotiation', 'Closed Won', 'Closed Lost'],
    sources: ['Website', 'Referral', 'Social Media', 'Email Campaign', 'Cold Call', 'Conference', 'Partner'],
    priorities: ['Low', 'Medium', 'High', 'Urgent'],
    products: [] // Will be populated from user's products
  },

  // ==================== ECOMMERCE MODULE ====================
  Ecommerce: {
    orderStatus: ['Pending', 'Processing', 'Shipped', 'Delivered', 'Cancelled', 'Refunded', 'On Hold'],
    paymentStatus: ['Paid', 'Pending', 'Failed', 'Refunded', 'Partially Paid'],
    fulfillmentStatus: ['Unfulfilled', 'Partially Fulfilled', 'Fulfilled', 'Ready for Pickup'],
    shippingMethods: ['Standard Shipping', 'Express Shipping', 'Next Day Delivery', 'Pickup', 'International'],
    productCategories: ['Electronics', 'Clothing', 'Food', 'Books', 'Beauty', 'Home & Garden', 'Sports', 'Toys', 'Other']
  },

  // ==================== CRM MODULE ====================
  CRM: {
    leadStatus: ['New', 'Contacted', 'Qualified', 'Proposal', 'Negotiation', 'Won', 'Lost', 'Inactive'],
    leadSources: ['Website', 'Referral', 'Cold Email', 'Social Media', 'Event', 'Advertisement', 'Partner'],
    industries: ['Technology', 'Healthcare', 'Finance', 'Retail', 'Manufacturing', 'Education', 'Real Estate', 'Agriculture', 'Transportation', 'Other'],
    contactRoles: ['Decision Maker', 'Influencer', 'User', 'Champion', 'Gatekeeper']
  },

  // ==================== HR MODULE ====================
  HR: {
    employmentTypes: ['Full-time', 'Part-time', 'Contract', 'Intern', 'Freelance', 'Remote', 'Temporary'],
    departments: ['Executive', 'Sales', 'Marketing', 'Engineering', 'Product', 'HR', 'Finance', 'Operations', 'Customer Support', 'Legal'],
    positions: ['Manager', 'Senior', 'Junior', 'Lead', 'Director', 'VP', 'C-Level', 'Associate', 'Intern'],
    employmentStatus: ['Active', 'On Leave', 'Probation', 'Terminated', 'Resigned', 'Suspended'],
    genders: ['Male', 'Female', 'Non-binary', 'Prefer not to say'],
    maritalStatus: ['Single', 'Married', 'Divorced', 'Widowed']
  },

  // ==================== LOGISTICS MODULE ====================
  Logistics: {
    shipmentStatus: ['Order Placed', 'Processing', 'Picked Up', 'In Transit', 'Out for Delivery', 'Delivered', 'Failed', 'Returned'],
    carriers: ['DHL', 'FedEx', 'UPS', 'USPS', 'Aramex', 'Local Courier', 'Other'],
    shipmentTypes: ['Standard', 'Express', 'Overnight', 'International', 'Freight'],
    warehouseZones: ['Zone A', 'Zone B', 'Zone C', 'North', 'South', 'East', 'West']
  },

  // ==================== TAX MODULE ====================
  Tax: {
    taxTypes: ['Sales Tax', 'VAT', 'Income Tax', 'Property Tax', 'Payroll Tax', 'Customs Duty', 'Excise Tax'],
    filingFrequencies: ['Monthly', 'Quarterly', 'Annually', 'Semi-annually'],
    taxStatus: ['Draft', 'Filed', 'Pending', 'Audited', 'Paid', 'Overdue']
  },

  // ==================== AGRO MODULE ====================
  Agro: {
    types: ['Crop', 'Livestock', 'Poultry', 'Fishery', 'Mixed'],
    
    cropActivities: ['Land Preparation', 'Planting', 'Fertilizing', 'Irrigation', 'Pest Control', 'Weeding', 'Harvesting', 'Storage', 'Sale', 'Processing'],
    cropTypes: ['Maize', 'Rice', 'Wheat', 'Soybeans', 'Cassava', 'Yam', 'Tomatoes', 'Onions', 'Peppers', 'Cabbage', 'Carrots', 'Beans', 'Groundnuts', 'Cotton', 'Cocoa', 'Coffee', 'Tea', 'Rubber', 'Other'],
    
    livestockActivities: ['Feeding', 'Vaccination', 'Breeding', 'Health Check', 'Milk Collection', 'Egg Collection', 'Sale', 'Purchase', 'Slaughter', 'Medical Treatment'],
    livestockTypes: ['Cattle', 'Goats', 'Sheep', 'Pigs', 'Chickens', 'Turkeys', 'Ducks', 'Rabbits', 'Fish', 'Bees', 'Other'],
    
    poultryActivities: ['Feeding', 'Vaccination', 'Egg Collection', 'Cleaning', 'Health Check', 'Sale'],
    poultryTypes: ['Broilers', 'Layers', 'Turkeys', 'Ducks', 'Geese', 'Quail'],
    
    fisheryActivities: ['Feeding', 'Water Treatment', 'Harvesting', 'Stocking', 'Health Check', 'Sale'],
    fisheryTypes: ['Tilapia', 'Catfish', 'Salmon', 'Trout', 'Carp', 'Shrimp', 'Prawns', 'Other'],
    
    units: ['kg', 'grams', 'tons', 'pieces', 'dozens', 'liters', 'acres', 'hectares', 'bags', 'crates', 'bundles']
  }
};

// ==================== HELPER FUNCTIONS ====================

function getReferenceData() {
  return REFERENCE_DATA;
}

function getModuleReferenceData(moduleName) {
  return REFERENCE_DATA[moduleName] || {};
}

function getTransactionTypes() {
  return REFERENCE_DATA.Finance.transactionTypes;
}

function getIncomeCategories() {
  return REFERENCE_DATA.Finance.categories.Income;
}

function getExpenseCategories() {
  return REFERENCE_DATA.Finance.categories.Expense;
}

function getSubcategories(category) {
  return REFERENCE_DATA.Finance.subcategories[category] || ['Other'];
}

function getPaymentMethods() {
  return REFERENCE_DATA.Finance.paymentMethods;
}

function getSalesStages() {
  return REFERENCE_DATA.Sales.stages;
}

function getOrderStatus() {
  return REFERENCE_DATA.Ecommerce.orderStatus;
}

function getEmploymentTypes() {
  return REFERENCE_DATA.HR.employmentTypes;
}

function getDepartments() {
  return REFERENCE_DATA.HR.departments;
}

function getAgroTypes() {
  return REFERENCE_DATA.Agro.types;
}

function getAgroActivities(type) {
  if (type === 'Crop') return REFERENCE_DATA.Agro.cropActivities;
  if (type === 'Livestock') return REFERENCE_DATA.Agro.livestockActivities;
  if (type === 'Poultry') return REFERENCE_DATA.Agro.poultryActivities;
  if (type === 'Fishery') return REFERENCE_DATA.Agro.fisheryActivities;
  return [];
}

function getAgroUnits() {
  return REFERENCE_DATA.Agro.units;
}

// ==================== UNIVERSAL TRANSACTIONS SHEET ====================
// This sheet works for both manual entry AND API integration (Paystack, Stripe, etc.)

const UNIVERSAL_TRANSACTIONS_HEADERS = [
  'Date',           // Transaction date (YYYY-MM-DD)
  'Name',           // Customer name or business name
  'Email',          // Customer email (for API matching)
  'Amount',         // Transaction amount (positive for income, negative for expense)
  'Category',       // Transaction category (Sales, Operations, etc.)
  'Description',    // Transaction description/reference
  'Transaction_ID', // Unique ID (from Paystack or auto-generated)
  'Payment_Method', // Card, Bank Transfer, USSD, etc.
  'Status',         // success, pending, failed
  'Reference',      // Paystack reference or external reference
  'Source',         // 'paystack', 'manual', 'stripe', 'flutterwave'
  'Created_At',     // Timestamp
  'Synced_At'       // When synced to Finance sheet
];

// Create or get Universal Transactions sheet
function getOrCreateUniversalTransactionsSheet(workspaceId) {
  const workspace = SpreadsheetApp.openById(workspaceId);
  let sheet = workspace.getSheetByName('Universal_Transactions');
  
  if (!sheet) {
    sheet = workspace.insertSheet('Universal_Transactions');
    sheet.appendRow(UNIVERSAL_TRANSACTIONS_HEADERS);
    
    const headerRange = sheet.getRange(1, 1, 1, UNIVERSAL_TRANSACTIONS_HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f4f6');
  }
  
  return sheet;
}

// ==================== PAYSTACK API INTEGRATION ====================
// Users can call this function with their own Paystack secret key

function syncPaystackTransactions(paystackSecretKey, sessionId) {
  try {
    const user = getUserFromSession(sessionId);
    if (!user) return { success: false, message: "Unauthorized" };
    
    const workspace = getWorkspaceFile(sessionId);
    const universalSheet = getOrCreateUniversalTransactionsSheet(workspace.getId());
    const financeSheet = workspace.getSheetByName('Financial_Data');
    
    // Call Paystack API to get transactions
    const url = 'https://api.paystack.co/transaction?perPage=100';
    const options = {
      headers: {
        'Authorization': `Bearer ${paystackSecretKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    
    if (!result.status) {
      return { success: false, message: "Paystack API error: " + result.message };
    }
    
    let syncedCount = 0;
    let newTransactions = [];
    
    result.data.forEach(transaction => {
      // Check if transaction already exists
      const existingData = universalSheet.getDataRange().getValues();
      const existingRefs = existingData.slice(1).map(row => row[9]); // Reference column
      
      if (!existingRefs.includes(transaction.reference)) {
        const transactionData = {
          Date: new Date(transaction.transaction_date).toISOString().split('T')[0],
          Name: transaction.customer?.business_name || transaction.customer?.customer_name || 'Customer',
          Email: transaction.customer?.email || '',
          Amount: transaction.amount / 100, // Convert from kobo
          Category: transaction.metadata?.category || 'Sales Revenue',
          Description: transaction.metadata?.description || transaction.reason || 'Paystack Payment',
          Transaction_ID: transaction.id,
          Payment_Method: transaction.channel || 'card',
          Status: transaction.status,
          Reference: transaction.reference,
          Source: 'paystack',
          Created_At: new Date().toISOString(),
          Synced_At: ''
        };
        
        newTransactions.push(transactionData);
        universalSheet.appendRow([
          transactionData.Date,
          transactionData.Name,
          transactionData.Email,
          transactionData.Amount,
          transactionData.Category,
          transactionData.Description,
          transactionData.Transaction_ID,
          transactionData.Payment_Method,
          transactionData.Status,
          transactionData.Reference,
          transactionData.Source,
          transactionData.Created_At,
          ''
        ]);
        
        syncedCount++;
      }
    });
    
    // Auto-sync to Finance sheet
    syncUniversalToFinance(workspace.getId(), newTransactions);
    
    return {
      success: true,
      message: `Synced ${syncedCount} new transactions from Paystack`,
      transactions: newTransactions
    };
    
  } catch (error) {
    console.error('Paystack sync error:', error);
    return { success: false, message: error.message };
  }
}

// Sync Universal Transactions to Financial_Data sheet
function syncUniversalToFinance(workspaceId, transactions) {
  try {
    const workspace = SpreadsheetApp.openById(workspaceId);
    const financeSheet = workspace.getSheetByName('Financial_Data');
    
    if (!financeSheet) return;
    
    const financeHeaders = financeSheet.getRange(1, 1, 1, financeSheet.getLastColumn()).getValues()[0];
    
    transactions.forEach(transaction => {
      const financeEntry = {
        Transaction_ID: generateTransactionId(),
        Date: transaction.Date,
        Type: transaction.Amount >= 0 ? 'revenue' : 'expense',
        Category: transaction.Category,
        Subcategory: transaction.Description?.substring(0, 50) || '',
        Module_Source: 'paystack_api',
        Module_Record_ID: transaction.Transaction_ID,
        Description: transaction.Description,
        Amount: Math.abs(transaction.Amount),
        Payment_Method: transaction.Payment_Method,
        Reference: transaction.Reference,
        Status: transaction.Status === 'success' ? 'completed' : transaction.Status,
        Created_By: 'paystack_sync',
        Created_At: new Date().toISOString()
      };
      
      const rowData = financeHeaders.map(header => financeEntry[header] || '');
      financeSheet.appendRow(rowData);
    });
    
    return { success: true, syncedCount: transactions.length };
    
  } catch (error) {
    console.error('Sync to finance error:', error);
    return { success: false, message: error.message };
  }
}

// ==================== WEBHOOK FOR REAL-TIME PAYSTACK SYNC ====================
// This can be exposed as a webhook endpoint for Paystack to call

function handlePaystackWebhook(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const event = postData.event;
    
    if (event === 'charge.success') {
      const data = postData.data;
      
      // Find which business this belongs to (you'd need to map by customer email)
      const businessSheet = getOrCreateBusinessSheet();
      const businessData = businessSheet.getDataRange().getValues();
      const businessHeaders = businessData[0];
      const ownerEmailCol = businessHeaders.indexOf('Owner_Email');
      
      // Find business by customer email
      let targetBusiness = null;
      for (let i = 1; i < businessData.length; i++) {
        if (businessData[i][ownerEmailCol] === data.customer?.email) {
          targetBusiness = businessData[i];
          break;
        }
      }
      
      if (targetBusiness) {
        const workspaceId = targetBusiness[businessHeaders.indexOf('Workspace_ID')];
        if (workspaceId) {
          const transaction = [{
            Date: new Date().toISOString().split('T')[0],
            Name: data.customer?.business_name || data.customer?.customer_name,
            Email: data.customer?.email,
            Amount: data.amount / 100,
            Category: data.metadata?.category || 'Sales Revenue',
            Description: data.metadata?.description || data.reason,
            Transaction_ID: data.id,
            Payment_Method: data.channel,
            Status: data.status,
            Reference: data.reference,
            Source: 'paystack_webhook',
            Created_At: new Date().toISOString()
          }];
          
          const universalSheet = getOrCreateUniversalTransactionsSheet(workspaceId);
          transaction.forEach(t => {
            universalSheet.appendRow(Object.values(t));
          });
          
          syncUniversalToFinance(workspaceId, transaction);
        }
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==================== GET UNIVERSAL TRANSACTIONS (for frontend) ====================

function getUniversalTransactions(sessionId, source = null) {
  try {
    const user = getUserFromSession(sessionId);
    if (!user) return { success: false, message: "Unauthorized" };
    
    const workspace = getWorkspaceFile(sessionId);
    const sheet = workspace.getSheetByName('Universal_Transactions');
    
    if (!sheet) return { success: true, transactions: [] };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    let transactions = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
    
    if (source) {
      transactions = transactions.filter(t => t.Source === source);
    }
    
    return { success: true, transactions: transactions };
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ==================== ADD SINGLE TRANSACTION (for manual entry or API) ====================

function addUniversalTransaction(transactionData, sessionId) {
  try {
    const user = getUserFromSession(sessionId);
    if (!user) return { success: false, message: "Unauthorized" };
    
    const workspace = getWorkspaceFile(sessionId);
    const universalSheet = getOrCreateUniversalTransactionsSheet(workspace.getId());
    
    // Validate required fields
    if (!transactionData.Date || !transactionData.Amount) {
      return { success: false, message: "Date and Amount are required" };
    }
    
    const newTransaction = {
      Date: transactionData.Date,
      Name: transactionData.Name || '',
      Email: transactionData.Email || '',
      Amount: parseFloat(transactionData.Amount),
      Category: transactionData.Category || 'Uncategorized',
      Description: transactionData.Description || '',
      Transaction_ID: transactionData.Transaction_ID || generateTransactionId(),
      Payment_Method: transactionData.Payment_Method || 'manual',
      Status: transactionData.Status || 'completed',
      Reference: transactionData.Reference || '',
      Source: transactionData.Source || 'manual',
      Created_At: new Date().toISOString(),
      Synced_At: ''
    };
    
    universalSheet.appendRow([
      newTransaction.Date,
      newTransaction.Name,
      newTransaction.Email,
      newTransaction.Amount,
      newTransaction.Category,
      newTransaction.Description,
      newTransaction.Transaction_ID,
      newTransaction.Payment_Method,
      newTransaction.Status,
      newTransaction.Reference,
      newTransaction.Source,
      newTransaction.Created_At,
      ''
    ]);
    
    // Sync to Finance sheet
    syncUniversalToFinance(workspace.getId(), [newTransaction]);
    
    return { success: true, transaction: newTransaction };
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}



// ==================== DYNAMIC EMAIL AUTO-RESPONDER SYSTEM ====================
// This system is CONFIGURABLE by each business - no hardcoding!

// Get auto-responder rules for a business
function getAutoResponderRules(businessId) {
  try {
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    let rulesSheet = mainSheet.getSheetByName('Email_AutoResponders');
    
    if (!rulesSheet) {
      // Create default rules sheet
      rulesSheet = mainSheet.insertSheet('Email_AutoResponders');
      rulesSheet.appendRow([
        'Business_ID', 'Category', 'Trigger_Keyword', 'Subject_Template', 
        'Body_Template', 'Attachment_URL', 'Send_To', 'Is_Active', 'Created_At'
      ]);
      
      // Add default template (businesses can edit)
      rulesSheet.appendRow([
        businessId, 'General', '', 
        'Thank you for your payment!',
        'Hello {Name},\n\nThank you for your payment of ₦{Amount}. Your transaction has been confirmed.\n\nReference: {Reference}\n\nBest regards,\n{Business_Name}',
        '', 'customer', 'YES', new Date().toISOString()
      ]);
    }
    
    const data = rulesSheet.getDataRange().getValues();
    const headers = data[0];
    
    const businessIdCol = headers.indexOf('Business_ID');
    const categoryCol = headers.indexOf('Category');
    const keywordCol = headers.indexOf('Trigger_Keyword');
    const subjectCol = headers.indexOf('Subject_Template');
    const bodyCol = headers.indexOf('Body_Template');
    const attachmentCol = headers.indexOf('Attachment_URL');
    const sendToCol = headers.indexOf('Send_To');
    const activeCol = headers.indexOf('Is_Active');
    
    const rules = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][businessIdCol] === businessId && data[i][activeCol] === 'YES') {
        rules.push({
          category: data[i][categoryCol],
          keyword: data[i][keywordCol],
          subjectTemplate: data[i][subjectCol],
          bodyTemplate: data[i][bodyCol],
          attachmentUrl: data[i][attachmentCol],
          sendTo: data[i][sendToCol]
        });
      }
    }
    
    return rules;
    
  } catch (error) {
    console.error('Error getting auto-responder rules:', error);
    return [];
  }
}

// Process and send auto-response based on transaction
function processAutoResponse(transactionData, businessId, workspaceId) {
  try {
    const rules = getAutoResponderRules(businessId);
    if (rules.length === 0) return;
    
    // Get business info
    const businessSheet = getOrCreateBusinessSheet();
    const businessData = businessSheet.getDataRange().getValues();
    const businessHeaders = businessData[0];
    const bizIdCol = businessHeaders.indexOf('Business_ID');
    const bizNameCol = businessHeaders.indexOf('Business_Name');
    const ownerEmailCol = businessHeaders.indexOf('Owner_Email');
    
    let businessName = '';
    let businessEmail = '';
    for (let i = 1; i < businessData.length; i++) {
      if (businessData[i][bizIdCol] === businessId) {
        businessName = businessData[i][bizNameCol];
        businessEmail = businessData[i][ownerEmailCol];
        break;
      }
    }
    
    // Find matching rule
    let matchedRule = null;
    for (const rule of rules) {
      // Check category match
      if (rule.category && transactionData.Category === rule.category) {
        matchedRule = rule;
        break;
      }
      // Check keyword match in description
      if (rule.keyword && transactionData.Description && 
          transactionData.Description.toLowerCase().includes(rule.keyword.toLowerCase())) {
        matchedRule = rule;
        break;
      }
    }
    
    // If no specific rule, use general rule (first one)
    if (!matchedRule) {
      matchedRule = rules.find(r => r.category === 'General');
    }
    
    if (!matchedRule) return;
    
    // Determine recipient
    let recipient = '';
    if (matchedRule.sendTo === 'customer') {
      recipient = transactionData.Email;
    } else if (matchedRule.sendTo === 'business') {
      recipient = businessEmail;
    } else if (matchedRule.sendTo === 'both') {
      recipient = transactionData.Email + ',' + businessEmail;
    }
    
    if (!recipient) return;
    
    // Replace template variables
    const subject = replaceTemplateVariables(matchedRule.subjectTemplate, transactionData, businessName);
    let body = replaceTemplateVariables(matchedRule.bodyTemplate, transactionData, businessName);
    
    // Add footer
    body += `\n\n---\nThis is an automated message from ${businessName}.`;
    
    // Send email
    const options = { name: businessName };
    if (matchedRule.attachmentUrl) {
      options.attachments = [UrlFetchApp.fetch(matchedRule.attachmentUrl).getBlob()];
    }
    
    GmailApp.sendEmail(recipient, subject, body, options);
    
    // Log the auto-response
    logAutoResponse(businessId, transactionData, subject);
    
    return { success: true, rule: matchedRule };
    
  } catch (error) {
    console.error('Auto-response error:', error);
    return { success: false, message: error.message };
  }
}

// Replace variables in templates
function replaceTemplateVariables(template, data, businessName) {
  let result = template;
  
  const variables = {
    '{Name}': data.Name || '',
    '{Email}': data.Email || '',
    '{Amount}': data.Amount ? `₦${parseFloat(data.Amount).toLocaleString()}` : '',
    '{Category}': data.Category || '',
    '{Description}': data.Description || '',
    '{Reference}': data.Reference || '',
    '{Date}': data.Date || new Date().toISOString().split('T')[0],
    '{Business_Name}': businessName || '',
    '{Transaction_ID}': data.Transaction_ID || ''
  };
  
  for (const [key, value] of Object.entries(variables)) {
    result = result.split(key).join(value);
  }
  
  return result;
}

// Log auto-response history
function logAutoResponse(businessId, transactionData, subject) {
  try {
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = mainSheet.getSheetByName('AutoResponse_Log');
    
    if (!logSheet) {
      logSheet = mainSheet.insertSheet('AutoResponse_Log');
      logSheet.appendRow([
        'Timestamp', 'Business_ID', 'Customer_Email', 'Category', 
        'Amount', 'Subject_Sent', 'Status'
      ]);
    }
    
    logSheet.appendRow([
      new Date().toISOString(),
      businessId,
      transactionData.Email,
      transactionData.Category,
      transactionData.Amount,
      subject,
      'sent'
    ]);
    
  } catch (error) {
    console.error('Log error:', error);
  }
}

// ==================== BUSINESS CAN CONFIGURE THEIR OWN RULES ====================

// Save/Update auto-responder rule (businesses can call this from frontend)
function saveAutoResponderRule(ruleData, sessionId) {
  try {
    const user = getUserFromSession(sessionId);
    if (!user || user.role !== 'owner') {
      return { success: false, message: "Only business owners can configure auto-responders" };
    }
    
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    let rulesSheet = mainSheet.getSheetByName('Email_AutoResponders');
    
    if (!rulesSheet) {
      rulesSheet = mainSheet.insertSheet('Email_AutoResponders');
      rulesSheet.appendRow([
        'Business_ID', 'Category', 'Trigger_Keyword', 'Subject_Template', 
        'Body_Template', 'Attachment_URL', 'Send_To', 'Is_Active', 'Created_At'
      ]);
    }
    
    rulesSheet.appendRow([
      user.businessId,
      ruleData.category,
      ruleData.keyword || '',
      ruleData.subjectTemplate,
      ruleData.bodyTemplate,
      ruleData.attachmentUrl || '',
      ruleData.sendTo || 'customer',
      'YES',
      new Date().toISOString()
    ]);
    
    return { success: true, message: "Auto-responder rule saved" };
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Get all rules for a business (for frontend display)
function getBusinessAutoResponderRules(sessionId) {
  try {
    const user = getUserFromSession(sessionId);
    if (!user) return { success: false, message: "Unauthorized" };
    
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    let rulesSheet = mainSheet.getSheetByName('Email_AutoResponders');
    
    if (!rulesSheet) return { success: true, rules: [] };
    
    const data = rulesSheet.getDataRange().getValues();
    const headers = data[0];
    
    const businessIdCol = headers.indexOf('Business_ID');
    const categoryCol = headers.indexOf('Category');
    const keywordCol = headers.indexOf('Trigger_Keyword');
    const subjectCol = headers.indexOf('Subject_Template');
    const bodyCol = headers.indexOf('Body_Template');
    const attachmentCol = headers.indexOf('Attachment_URL');
    const sendToCol = headers.indexOf('Send_To');
    const activeCol = headers.indexOf('Is_Active');
    
    const rules = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][businessIdCol] === user.businessId) {
        rules.push({
          category: data[i][categoryCol],
          keyword: data[i][keywordCol],
          subjectTemplate: data[i][subjectCol],
          bodyTemplate: data[i][bodyCol],
          attachmentUrl: data[i][attachmentCol],
          sendTo: data[i][sendToCol],
          isActive: data[i][activeCol]
        });
      }
    }
    
    return { success: true, rules: rules };
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// Delete a rule
function deleteAutoResponderRule(category, sessionId) {
  try {
    const user = getUserFromSession(sessionId);
    if (!user || user.role !== 'owner') {
      return { success: false, message: "Unauthorized" };
    }
    
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet();
    const rulesSheet = mainSheet.getSheetByName('Email_AutoResponders');
    
    if (!rulesSheet) return { success: false, message: "No rules found" };
    
    const data = rulesSheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === user.businessId && data[i][1] === category) {
        rulesSheet.deleteRow(i + 1);
        return { success: true, message: "Rule deleted" };
      }
    }
    
    return { success: false, message: "Rule not found" };
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}

