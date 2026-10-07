//workspace.gs
function createWorkspaceForBusiness(businessId, businessName, ownerEmail) {
  try {
    console.log("Creating workspace for:", businessName);
    
    const workspaceName = businessName + " - BizOS Workspace";
    const newWorkspace = SpreadsheetApp.create(workspaceName);
    const workspaceFile = DriveApp.getFileById(newWorkspace.getId());
    
    // Get the folder where YOUR main Business OS v1.0.2 file lives
    const mainFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    const parents = mainFile.getParents();
    
    const parentFolder = getBizOSFolder();

    // Create subfolder per business
    const businessFolder = getOrCreateBusinessSubfolder(
      parentFolder,
      businessId,
      businessName
    );

    // Move workspace into that subfolder
    workspaceFile.moveTo(businessFolder);
    
    // Create sheets
    Object.keys(MODULES).forEach(moduleName => {
      const sheet = newWorkspace.insertSheet(MODULES[moduleName].sheet);
      addDefaultHeadersToModule(sheet, moduleName);
    });
    
    const usersSheet = newWorkspace.insertSheet('Workspace_Users');
    usersSheet.appendRow(['Email', 'Name', 'Role', 'Business_ID', 'Added_At', 'Status']);
    usersSheet.appendRow([ownerEmail, 'Business Owner', 'owner', businessId, new Date().toISOString(), 'active']);
    
    newWorkspace.addEditor(ownerEmail);
    
    return {
      success: true,
      workspaceId: newWorkspace.getId(),
      workspaceUrl: newWorkspace.getUrl()
    };
    
  } catch (error) {
    console.error('Workspace creation error:', error);
    return { success: false, message: 'We could not create your workspace right now. Please try again.' };
  }
}

function getBizOSFolder() {
  const folderName = "Business OS";
  let folder;
  
  // Try to find existing folder
  const folders = DriveApp.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    folder = folders.next();
    console.log('Found existing Business OS folder:', folder.getId());
  } else {
    folder = DriveApp.createFolder(folderName);
    console.log('Created new Business OS folder:', folder.getId());
  }
  
  return folder;
}

function addDefaultHeadersToModule(sheet, moduleName) {
  const defaultHeaders = {
    Finance: [
      'Transaction_ID', 'Date', 'Type', 'Category', 'Subcategory', 
      'Module_Source', 'Module_Record_ID', 'Description', 'Amount', 
      'Payment_Method', 'Reference', 'Status', 'Created_By', 'Created_At'
    ],
    Sales: [
      'Deal_ID', 'Date', 'Customer_Name', 'Customer_Email', 'Product_Service',
      'Quantity', 'Unit_Price', 'Total_Amount', 'Stage', 'Probability',
      'Expected_Close_Date', 'Sales_Person', 'Notes', 'Financial_Link_ID', 'Created_By'
    ],
    Ecommerce: [
      'Order_ID', 'Date', 'Customer_Email', 'Product_Name', 'SKU', 'Quantity',
      'Unit_Price', 'Subtotal', 'Shipping_Cost', 'Tax', 'Total_Amount',
      'Payment_Status', 'Fulfillment_Status', 'Tracking_Number', 'Financial_Link_ID', 'Created_By'
    ],
    CRM: [
      'Contact_ID', 'Name', 'Email', 'Phone', 'Company', 'Status', 
      'Last_Contact', 'Assigned_To', 'Notes', 'Created_By', 'Created_At'
    ],
    HR: [
      'Employee_ID', 'Name', 'Position', 'Department', 'Hire_Date', 
      'Salary', 'Manager', 'Status', 'Created_By', 'Created_At'
    ],
    Logistics: [
      'Shipment_ID', 'Order_ID', 'Status', 'Location', 'Carrier', 
      'Tracking', 'ETA', 'Dispatched_By', 'Created_By', 'Created_At'
    ],
    Tax: [
      'Tax_ID', 'Date', 'Type', 'Amount', 'Tax_Rate', 'Tax_Amount', 
      'Reference', 'Filed_By', 'Status', 'Created_By', 'Created_At'
    ],
    Agro: [
      'Record_ID', 'Date', 'Type', 'Activity', 'Crop_Name', 'Livestock_Type',
      'Quantity', 'Unit', 'Unit_Cost', 'Total_Cost', 'Selling_Price', 
      'Revenue', 'Field_Plot', 'Health_Status', 'Notes', 'Financial_Link_ID', 'Created_By'
    ],
    // Add to the defaultHeaders object
    Productivity: [
      'Record_ID', 'Staff_Email', 'Date', 'Task', 'Hours_Spent', 
      'Apps_Used', 'Notes', 'Created_By', 'Created_At'
    ],
    POS: [
      'Transaction_ID', 'Date', 'Customer_Name', 'Customer_Email', 
      'Product_Name', 'Quantity', 'Unit_Price', 'Total_Amount', 
      'Payment_Method', 'Status', 'Receipt_Number', 'Created_By', 'Created_At'
    ],
    Attendance: [
      'Record_ID', 'Staff_Email', 'Date', 'Clock_In', 'Clock_Out', 
      'Hours_Worked', 'Status', 'Created_By', 'Created_At'
    ],
    Warehouse: [
  'Item_ID', 'Item_Name', 'SKU', 'Category', 'Quantity', 
  'Min_Stock', 'Max_Stock', 'Location', 'Supplier', 'Last_Restock_Date',
  'Unit_Cost', 'Selling_Price', 'Status', 'Created_By', 'Created_At'
]
  };
  
  const headers = defaultHeaders[moduleName] || ['ID', 'Name', 'Value', 'Created_By', 'Created_At'];
  sheet.appendRow(headers);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f3f4f6');
}

function getSheetHeadersBySheet(sheet) {
  if (!sheet) return [];
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}


function getWorkspaceFile(sessionId) {
  const user = getUserFromSession(sessionId);
  if (!user) {
    throw new Error('Unauthorized: valid BizOS session required.');
  }

  // Free/starter accounts intentionally use the master workspace.
  // Paid accounts must have an isolated workspace; never silently fall back
  // to the master spreadsheet if their workspace is missing or unavailable.
  const tier = String(user.subscriptionTier || 'starter').toLowerCase();
  const isPaidTier = ['sovereign', 'professional', 'tier2', 'enterprise'].includes(tier);

  if (!user.workspaceId) {
    if (!isPaidTier) {
      console.log('Using master spreadsheet for free/starter account.');
      return SpreadsheetApp.getActiveSpreadsheet();
    }
    throw new Error('Workspace is not configured for this paid account.');
  }

  try {
    const workspace = SpreadsheetApp.openById(user.workspaceId);
    console.log('Using isolated workspace:', workspace.getName(), 'ID:', workspace.getId());
    return workspace;
  } catch (error) {
    console.error('Error opening isolated workspace:', error.message);
    if (!isPaidTier) {
      console.log('Using master spreadsheet for free/starter account after workspace lookup failure.');
      return SpreadsheetApp.getActiveSpreadsheet();
    }
    throw new Error('Unable to open the isolated workspace for this paid account.');
  }
}

function updateSourceRecordWithLink(workspace, moduleName, recordId, financialId) {
  try {
    const sheet = workspace.getSheetByName(MODULES[moduleName]?.sheet);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    
    const headers = data[0];
    const idCol = headers.findIndex(h => h === 'Deal_ID' || h === 'Order_ID' || h === 'Record_ID' || h === 'Employee_ID');
    const linkCol = headers.findIndex(h => h === 'Financial_Link_ID');
    
    if (idCol === -1 || linkCol === -1) return;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === recordId) {
        sheet.getRange(i + 1, linkCol + 1).setValue(financialId);
        break;
      }
    }
  } catch (error) {
    console.error('Update link error:', error);
  }
}



function forceDrivePermission() {
  try {
    // This will trigger Drive permission request
    const testFolder = DriveApp.createFolder("BizOS_Permission_Test");
    console.log("Created test folder:", testFolder.getName());
    
    // Create a subfolder to test full permissions
    const subFolder = testFolder.createFolder("Test_Subfolder");
    console.log("Created subfolder:", subFolder.getName());
    
    // Clean up
    testFolder.setTrashed(true);
    console.log("All Drive permissions working!");
    
    return "Success - Drive permissions granted";
  } catch (error) {
    console.error("Permission error:", error.toString());
    return "Error: " + error.toString();
  }
}

function authorizeDrive() {
  // This simple line will trigger the Drive permission dialog
  const folder = DriveApp.getRootFolder();
  Logger.log("Drive access granted! Root folder ID: " + folder.getId());
}



function getOrCreateBusinessSubfolder(parentFolder, businessId, businessName) {
  // Create a clean folder name (remove special characters)
  const safeBusinessName = businessName.replace(/[^a-zA-Z0-9]/g, '_');
  const subfolderName = `${businessId} - ${safeBusinessName}`;
  
  // Check if subfolder already exists
  const existingFolders = parentFolder.getFoldersByName(subfolderName);
  
  if (existingFolders.hasNext()) {
    console.log('Using existing subfolder:', subfolderName);
    return existingFolders.next();
  } else {
    const newSubfolder = parentFolder.createFolder(subfolderName);
    console.log('Created new subfolder:', subfolderName);
    return newSubfolder;
  }
}




