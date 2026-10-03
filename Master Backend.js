// ============================================================
// 📁 Master Backend.gs - COMPLETE MASTER API (FIXED)
// ============================================================

const MASTER_SECRET = 'TEST_SECRET_2024';

// ============================================================
// 🔗 GET MASTER API URL - CORRECT VERSION (NO DUPLICATE)
// ============================================================
function getMasterApiUrl() {
  // Your main master API URL - this is the correct one
  const mainMasterUrl = 'https://script.google.com/macros/s/AKfycby9J1h6m-hCeFYSCc5IgTM7DmVKRgKclUNdCqN3-dIiL34RWVQpTr10LBtdnUad0m-v/exec';
  
  // Try to get from script properties first
  let url = PropertiesService.getScriptProperties().getProperty('MASTER_API_URL');
  if (url) {
    console.log('📡 Using MASTER_API_URL from properties:', url);
    return url;
  }
  
  // Cache it in properties for faster access
  PropertiesService.getScriptProperties().setProperty('MASTER_API_URL', mainMasterUrl);
  
  console.log('📡 Using master API URL:', mainMasterUrl);
  return mainMasterUrl;
}

// ============================================================
// 📋 GET DEPLOYMENTS
// ============================================================

function getDeployments() {
  try {
    const url = 'https://script.googleapis.com/v1/projects/' + ScriptApp.getScriptId() + '/deployments';
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      console.error('❌ Failed to get deployments:', response.getContentText());
      return [];
    }
    
    const result = JSON.parse(response.getContentText());
    return result.deployments || [];
  } catch (error) {
    console.error('❌ getDeployments error:', error);
    return [];
  }
}

// ============================================================
// 🔐 GET CLIENT BY ID
// ============================================================

function getClientById(clientId) {
  try {
    console.log('🔍 getClientById called for:', clientId);
    
    const ss = getBizOSMasterSpreadsheet_();
    const sheet = ss.getSheetByName('Clients');
    
    if (!sheet) {
      console.log('⚠️ Clients sheet not found');
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      console.log('⚠️ Clients sheet has no data');
      return null;
    }
    
    const headers = data[0];
    const clientIdCol = headers.indexOf('Client_ID');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row && row[clientIdCol] === clientId) {
        console.log(`✅ Found client: ${clientId} at row ${i+1}`);
        
        const client = {};
        headers.forEach((header, index) => {
          const value = row[index] || '';
          switch(header) {
            case 'Client_ID': client.clientId = value; break;
            case 'Email': client.email = value; break;
            case 'Client_Name': client.clientName = value; break;
            case 'Domain': client.domain = value; break;
            case 'Primary_Color': client.primaryColor = value; break;
            case 'Logo_Url': client.logoUrl = value; break;
            case 'Custom_Domain': client.customDomain = value; break;
            case 'Tier': client.tier = value; break;
            case 'Status': client.status = value; break;
            case 'Sheet_ID': client.sheetId = value; break;
            case 'Web_App_URL': client.webAppUrl = value; break;
            case 'API_Key': client.apiKey = value; break;
            case 'Created_At': client.createdAt = value; break;
            default: client[header] = value;
          }
        });
        
        return client;
      }
    }
    
    console.log(`❌ Client not found: ${clientId}`);
    return null;
  } catch (error) {
    console.error('❌ getClientById error:', error);
    return null;
  }
}

// ============================================================
// 🔐 GET CLIENT FROM BUSINESS SHEET - FALLBACK
// ============================================================

function getClientFromBusinessSheet(clientId) {
  try {
    const ss = getBizOSMasterSpreadsheet_();
    const sheet = ss.getSheetByName('Businesses');
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return null;
    
    const headers = data[0];
    const idCol = headers.indexOf('Business_ID');
    const emailCol = headers.indexOf('Owner_Email');
    const nameCol = headers.indexOf('Business_Name');
    const tierCol = headers.indexOf('Subscription_Tier');
    const workspaceCol = headers.indexOf('Workspace_ID');
    const statusCol = headers.indexOf('Status');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === clientId) {
        return {
          clientId: data[i][idCol] || '',
          email: data[i][emailCol] || '',
          clientName: data[i][nameCol] || 'My Business',
          primaryColor: '#1a5276',
          logoUrl: '',
          customDomain: '',
          tier: data[i][tierCol] || 'sovereign',
          sheetId: '',
          webAppUrl: '',
          status: data[i][statusCol] || 'active',
          scriptId: '',
          businessId: data[i][idCol] || ''
        };
      }
    }
    
    return null;
  } catch (error) {
    console.error('Get client from business error:', error);
    return null;
  }
}

// ============================================================
// 💾 SAVE CLIENT RECORD
// ============================================================

// ============================================================
// 💾 SAVE CLIENT RECORD - FIXED (No session validation)
// ============================================================

function saveClientRecord(settings) {
  try {
    console.log('💾 saveClientRecord called for:', settings.clientId);

    const ss = getBizOSMasterSpreadsheet_();
    let sheet = ss.getSheetByName('Clients');

    if (!sheet) {
      console.log('📋 Creating Clients sheet...');
      sheet = ss.insertSheet('Clients');
      const headers = [
        'Client_ID', 'Email', 'Client_Name', 'Domain', 'Primary_Color',
        'Logo_Url', 'Custom_Domain', 'Tier', 'Status', 'Sheet_ID',
        'Web_App_URL', 'API_Key', 'Created_At', 'Script_ID'
      ];
      sheet.appendRow(headers);
      console.log('✅ Clients sheet created');
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    console.log('📋 Headers:', headers.join(' | '));

    const rowData = [];
    headers.forEach(header => {
      let value = '';
      switch(header) {
        case 'Client_ID': value = settings.clientId || ''; break;
        case 'Email': value = settings.email || ''; break;
        case 'Client_Name': value = settings.clientName || ''; break;
        case 'Domain': value = settings.domain || settings.customDomain || ''; break;
        case 'Primary_Color': value = settings.primaryColor || '#2E7D32'; break;
        case 'Logo_Url': value = settings.logoUrl || ''; break;
        case 'Custom_Domain': value = settings.customDomain || ''; break;
        case 'Tier': value = settings.tier || 'sovereign'; break;
        case 'Status': value = 'active'; break;
        case 'Sheet_ID': value = settings.sheetId || ''; break;
        case 'Web_App_URL': value = settings.webAppUrl || settings.landingUrl || ''; break;
        case 'API_Key': value = settings.apiKey || Utilities.getUuid(); break;
        case 'Created_At': value = new Date().toISOString(); break;
        case 'Script_ID': value = settings.scriptId || ''; break;
        default: value = '';
      }
      rowData.push(value);
    });

    // Check if client exists
    const data = sheet.getDataRange().getValues();
    const clientIdCol = headers.indexOf('Client_ID');
    let exists = false;
    let existingRow = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i] && data[i][clientIdCol] === settings.clientId) {
        exists = true;
        existingRow = i + 1;
        break;
      }
    }

    if (exists && existingRow > 0) {
      for (let col = 0; col < rowData.length; col++) {
        sheet.getRange(existingRow, col + 1).setValue(rowData[col]);
      }
      console.log(`✅ Client ${settings.clientId} updated at row ${existingRow}`);
      console.log(`✅ Script ID saved: ${settings.scriptId}`);
    } else {
      sheet.appendRow(rowData);
      console.log(`✅ Client ${settings.clientId} added as new row`);
      console.log(`✅ Script ID saved: ${settings.scriptId}`);
    }

    return { success: true, message: 'Client saved successfully' };

  } catch (error) {
    console.error('❌ saveClientRecord error:', error);
    return { success: false, message: error.message };
  }
}


// ============================================================
// ✅ IMPROVED autoSetupClient - Now checks for duplicates
// ============================================================
function autoSetupClient(paymentData) {
  try {
    console.log('🚀 Setting up client:', paymentData.email);

    const businessName = paymentData.businessName || paymentData.name || 'My Business';
    
    // ✅ CHECK FOR DUPLICATES FIRST
    console.log('🔍 Checking for duplicate client...');
    const duplicateCheck = checkDuplicateClient(paymentData.email, businessName);
    
    if (duplicateCheck.isDuplicate) {
      console.error('❌ DUPLICATE FOUND:', duplicateCheck.message);
      return {
        success: false,
        message: duplicateCheck.message,
        code: 'DUPLICATE_' + duplicateCheck.type,
        details: duplicateCheck
      };
    }

    console.log('✅ No duplicates found - proceeding with setup');

    const clientId = 'CLIENT_' + Utilities.getUuid().substring(0, 8).toUpperCase();
    console.log('🆔 Generated Client ID:', clientId);

    // Continue with rest of setup...
    let existingBusiness = getBusinessByEmail(paymentData.email);
    let businessId = existingBusiness ? existingBusiness.businessId : null;
    let existingData = null;

    if (existingBusiness) {
      console.log('📋 Existing user found:', existingBusiness.businessName);
      businessId = existingBusiness.businessId;
      existingData = getUserDataForMigration(paymentData.email, businessId);
    }

    // Create user
    console.log('👤 Ensuring user exists in Users sheet...');
    createUserIfNotExists(
      paymentData.email,
      paymentData.password || 'default123',
      businessName,
      'owner',
      businessId || clientId,
      true,
      true
    );

    // Create client sheet
    console.log('📊 Creating client sheet...');
    const sheetResult = createClientSheetInClientDrive(
      paymentData.email,
      clientId,
      businessName,
      existingData
    );

    if (!sheetResult.success) {
      return { success: false, message: 'Sheet creation failed: ' + sheetResult.message };
    }

    // Deploy client script
    let webAppUrl = '';
    let scriptId = '';

    try {
      console.log('📤 Attempting to deploy client script...');
      const masterUrl = getMasterApiUrl();
      const clientCode = generateClientCodeSafely({
        clientId: clientId,
        clientName: businessName,
        primaryColor: paymentData.primaryColor || '#2E7D32',
        logoUrl: paymentData.logoUrl || '',
        customDomain: paymentData.customDomain || '',
        email: paymentData.email,
        sheetId: sheetResult.sheetId,
        businessId: businessId || clientId,
        masterApiUrl: masterUrl
      });

      const scriptResult = createAndDeployClientScript(
        paymentData.email,
        clientId,
        clientCode,
        businessName,
        paymentData.primaryColor || '#2E7D32',
        paymentData.logoUrl || ''
      );

      if (scriptResult && scriptResult.success) {
        webAppUrl = scriptResult.webAppUrl;
        scriptId = scriptResult.scriptId;
        console.log('✅ Client script deployed successfully!');
        console.log('🔗 Web App URL:', webAppUrl);
      } else {
        console.log('⚠️ Deployment failed, using fallback mode:', scriptResult ? scriptResult.message : 'Unknown error');
        webAppUrl = createFallbackLandingPage(clientId, businessName, paymentData);
        scriptId = '';
      }
    } catch (deployError) {
      console.log('⚠️ Deployment error, using fallback mode:', deployError.message);
      webAppUrl = createFallbackLandingPage(clientId, businessName, paymentData);
      scriptId = '';
    }

    // Save client record
    console.log('💾 Saving client record...');
    const saveResult = saveClientRecord({
      clientId: clientId,
      email: paymentData.email,
      clientName: businessName,
      domain: paymentData.domain || paymentData.customDomain || '',
      customDomain: paymentData.customDomain || '',
      primaryColor: paymentData.primaryColor || '#2E7D32',
      logoUrl: paymentData.logoUrl || '',
      tier: paymentData.tier || 'sovereign',
      status: 'active',
      sheetId: sheetResult.sheetId,
      webAppUrl: webAppUrl,
      landingUrl: webAppUrl,
      apiKey: Utilities.getUuid(),
      scriptId: scriptId
    });

    console.log('📊 Save result:', saveResult);

    // Update business if exists
    if (existingBusiness) {
      updateBusinessWithClientInfo(businessId, clientId, webAppUrl);
    }

    // Send welcome email
    sendClientWelcomeEmail(
      paymentData.email,
      businessName,
      webAppUrl,
      clientId
    );

    console.log('✅ ========================================');
    console.log('✅ CLIENT SETUP COMPLETE!');
    console.log('✅ ========================================');
    console.log('🆔 Client ID:', clientId);
    console.log('🔗 Landing URL:', webAppUrl);
    console.log('📊 Sheet URL:', sheetResult.sheetUrl);

    return {
      success: true,
      clientId: clientId,
      landingUrl: webAppUrl,
      sheetUrl: sheetResult.sheetUrl,
      email: paymentData.email,
      password: paymentData.password || 'default123',
      scriptId: scriptId,
      message: 'Client setup complete!'
    };

  } catch (error) {
    console.error('❌ Auto setup error:', error);
    console.error('❌ Stack:', error.stack);
    return { success: false, message: error.message };
  }
}

// ============================================================
// 👤 CREATE USER IF NOT EXISTS
// ============================================================

function createUserIfNotExists(email, password, name, role, businessId, isVerified, isClient) {
  try {
    console.log(`👤 Creating/checking user: ${email}`);
    
    const userSheet = getOrCreateUserSheet();
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    
    // Check if Is_Client column exists, if not add it
    let isClientCol = headers.indexOf('Is_Client');
    if (isClientCol === -1) {
      const newCol = userSheet.getLastColumn() + 1;
      userSheet.getRange(1, newCol).setValue('Is_Client');
      isClientCol = newCol;
      console.log('✅ Added Is_Client column');
    }
    
    for (let i = 1; i < data.length; i++) {
      if (data[i] && data[i][emailCol] === email) {
        console.log(`👤 User ${email} already exists`);
        
        if (password) {
          const passwordCol = headers.indexOf('Password_Hash');
          if (passwordCol !== -1) {
            const hashedPassword = hashPassword(password);
            userSheet.getRange(i + 1, passwordCol + 1).setValue(hashedPassword);
            console.log(`✅ Password updated for ${email}`);
          }
        }
        return { success: true, message: 'User already exists' };
      }
    }
    
    const nameCol = headers.indexOf('Name');
    const roleCol = headers.indexOf('Role');
    const businessIdCol = headers.indexOf('Business_ID');
    const isVerifiedCol = headers.indexOf('Is_Verified');
    const passwordCol = headers.indexOf('Password_Hash');
    const createdCol = headers.indexOf('Created_At');
    
    const hashedPassword = hashPassword(password || 'default123');
    console.log(`🔑 Password hashed for ${email}`);
    
    const rowData = [];
    headers.forEach(header => {
      switch(header) {
        case 'Email': rowData.push(email); break;
        case 'Password_Hash': rowData.push(hashedPassword); break;
        case 'Name': rowData.push(name || ''); break;
        case 'Role': rowData.push(role || 'owner'); break;
        case 'Business_ID': rowData.push(businessId || ''); break;
        case 'Is_Verified': rowData.push(isVerified ? 'YES' : 'NO'); break;
        case 'Created_At': rowData.push(new Date().toISOString()); break;
        case 'Is_Client': rowData.push(isClient ? 'YES' : 'NO'); break;
        default: rowData.push('');
      }
    });
    
    userSheet.appendRow(rowData);
    console.log(`👤 User ${email} created successfully`);
    return { success: true, message: 'User created' };
    
  } catch (error) {
    console.error('❌ createUserIfNotExists error:', error);
    return { success: false, message: error.message };
  }
}

// ============================================================
// 📊 CREATE CLIENT SHEET - WITH MIGRATION
// ============================================================

function createClientSheetInClientDrive(email, clientId, businessName, existingData) {
  try {
    console.log('📊 Creating client sheet for:', businessName);
    
    const sheetName = businessName + ' - BizOS Data';
    const newSheet = SpreadsheetApp.create(sheetName);
    const sheetId = newSheet.getId();
    const sheetUrl = newSheet.getUrl();
    
    const modules = {
      'Financial_Data': ['Date', 'Type', 'Category', 'Subcategory', 'Amount', 'Description', 'Status', 'Created_By', 'Created_At'],
      'Ecommerce_Data': ['Order_ID', 'Date', 'Customer_Email', 'Product', 'Quantity', 'Price', 'Status', 'Created_At'],
      'Sales_Data': ['Deal_ID', 'Date', 'Customer_Name', 'Product', 'Amount', 'Stage', 'Created_At'],
      'CRM_Data': ['Contact_ID', 'Name', 'Email', 'Phone', 'Company', 'Status', 'Created_At'],
      'HR_Data': ['Employee_ID', 'Name', 'Position', 'Department', 'Salary', 'Status', 'Created_At'],
      'Logistics_Data': ['Shipment_ID', 'Order_ID', 'Status', 'Location', 'Carrier', 'Created_At'],
      'Tax_Data': ['Tax_ID', 'Date', 'Type', 'Amount', 'Status', 'Created_At'],
      'Agro_Data': ['Record_ID', 'Date', 'Type', 'Activity', 'Quantity', 'Cost', 'Revenue', 'Created_At'],
      'Productivity_Data': ['Record_ID', 'Staff', 'Date', 'Task', 'Hours', 'Created_At'],
      'POS_Data': ['Transaction_ID', 'Date', 'Customer', 'Product', 'Quantity', 'Total', 'Created_At'],
      'Attendance_Data': ['Record_ID', 'Staff', 'Date', 'Clock_In', 'Clock_Out', 'Status', 'Created_At'],
      'Warehouse_Data': ['Item_ID', 'Item_Name', 'SKU', 'Quantity', 'Location', 'Status', 'Created_At']
    };
    
    Object.entries(modules).forEach(([name, headers]) => {
      const sheet = newSheet.insertSheet(name);
      sheet.appendRow(headers);
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f3f4f6');
    });
    
    const infoSheet = newSheet.insertSheet('Client_Info');
    infoSheet.appendRow(['Property', 'Value']);
    infoSheet.appendRow(['Client_ID', clientId]);
    infoSheet.appendRow(['Business_Name', businessName]);
    infoSheet.appendRow(['Created_At', new Date().toISOString()]);
    infoSheet.appendRow(['Tier', 'sovereign']);
    
    if (existingData) {
      console.log('📥 Migrating existing data for:', email);
      
      Object.entries(existingData).forEach(([moduleName, records]) => {
        if (records && records.length > 0) {
          const sheet = newSheet.getSheetByName(moduleName + '_Data');
          if (sheet) {
            const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
            records.forEach(record => {
              const rowData = headers.map(h => record[h] || '');
              sheet.appendRow(rowData);
            });
            console.log(`  ✅ Migrated ${records.length} records to ${moduleName}`);
          }
        }
      });
    }
    
    try {
      newSheet.addEditor(email);
      console.log('📤 Sheet shared with:', email);
    } catch (shareError) {
      console.log('⚠️ Could not share sheet:', shareError.message);
    }
    
    try {
      const folder = getOrCreateClientFolder(email, businessName);
      const file = DriveApp.getFileById(sheetId);
      file.moveTo(folder);
      console.log('📁 Sheet moved to folder:', folder.getName());
    } catch (folderError) {
      console.log('⚠️ Could not move to folder:', folderError.message);
    }
    
    return {
      success: true,
      sheetId: sheetId,
      sheetUrl: sheetUrl
    };
    
  } catch (error) {
    console.error('Create sheet error:', error);
    return { success: false, message: error.message };
  }
}

// ============================================================
// 📋 GET USER DATA FOR MIGRATION
// ============================================================

function getUserDataForMigration(email, businessId) {
  try {
    console.log('📋 Fetching user data for migration:', email);
    
    const ss = getBizOSMasterSpreadsheet_();
    const modules = ['Finance', 'Ecommerce', 'Sales', 'CRM', 'HR', 'Logistics', 'Tax', 'Agro', 'Productivity', 'POS', 'Attendance', 'Warehouse'];
    const userData = {};
    
    modules.forEach(moduleName => {
      const sheetName = moduleName + '_Data';
      const sheet = ss.getSheetByName(sheetName);
      
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return;
        
        const headers = data[0];
        const records = [];
        
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const emailCol = headers.indexOf('Email') !== -1 ? headers.indexOf('Email') : headers.indexOf('Created_By');
          const bizIdCol = headers.indexOf('Business_ID');
          
          let matches = false;
          if (emailCol !== -1 && row[emailCol] === email) {
            matches = true;
          } else if (bizIdCol !== -1 && row[bizIdCol] === businessId) {
            matches = true;
          }
          
          if (matches) {
            const record = {};
            headers.forEach((h, idx) => {
              record[h] = row[idx];
            });
            records.push(record);
          }
        }
        
        if (records.length > 0) {
          userData[moduleName] = records;
        }
      }
    });
    
    console.log('📊 Found data for migration:', Object.keys(userData).length, 'modules');
    return userData;
    
  } catch (error) {
    console.error('Get user data error:', error);
    return null;
  }
}


// ============================================================
// 📝 GENERATE CLIENT CODE - WITH BETTER ERROR HANDLING
// ============================================================

function generateClientCode(clientId, businessName, email, sheetId, masterUrl) {
  return `
const CLIENT_CONFIG = {
  clientId: '${clientId}',
  clientName: '${businessName}',
  primaryColor: '#7C3AED',
  logoUrl: '',
  customDomain: '',
  sheetId: '${sheetId}',
  email: '${email}',
  businessId: '${clientId}',
  masterApiUrl: '${masterUrl}'
};

// ============================================================
// 📁 CLIENT SCRIPT - FIXED authenticateClient()
// Add this to your generated client code
// ============================================================

function authenticateClient(email, password, businessName) {
  try {
    console.log('🔐 ========================================');
    console.log('🔐 CLIENT LOGIN ATTEMPT');
    console.log('🔐 ========================================');
    console.log('📧 Email:', email);
    console.log('🏢 Business Name:', businessName);
    console.log('📡 Master API URL:', CLIENT_CONFIG.masterApiUrl);

    const payload = {
      action: 'clientLogin',
      clientId: CLIENT_CONFIG.clientId,
      email: email,
      password: password,
      sheetId: CLIENT_CONFIG.sheetId,
      businessId: CLIENT_CONFIG.businessId,
      businessName: businessName || CLIENT_CONFIG.clientName
    };

    console.log('📤 Payload:', JSON.stringify(payload));

    const response = UrlFetchApp.fetch(CLIENT_CONFIG.masterApiUrl, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 30000
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    console.log('📊 Response Code:', responseCode);
    console.log('📝 Response Text (first 200 chars):', responseText.substring(0, 200));

    // ============================================================
    // CHECK FOR HTML RESPONSE (error page)
    // ============================================================
    if (responseText.trim().startsWith('<')) {
      console.error('❌ SERVER RETURNED HTML INSTEAD OF JSON');
      console.error('This usually means:');
      console.error('  1. Master API URL is wrong');
      console.error('  2. Master API is down or not accessible');
      console.error('  3. There is a redirect happening');
      console.error('Response:', responseText.substring(0, 500));
      
      return {
        success: false,
        message: 'Server error: Master API not responding correctly. Please check the configuration.',
        details: 'Expected JSON but got HTML. Master API URL may be incorrect.'
      };
    }

    // ============================================================
    // PARSE JSON RESPONSE
    // ============================================================
    if (responseCode !== 200) {
      console.error('❌ Server returned error code:', responseCode);
      return {
        success: false,
        message: 'Server error: ' + responseCode + ' - ' + responseText
      };
    }

    let result;
    try {
      result = JSON.parse(responseText);
    } catch (parseError) {
      console.error('❌ Failed to parse response:', parseError);
      console.error('Response was:', responseText);
      return {
        success: false,
        message: 'Server returned invalid response. Please try again later.',
        details: parseError.message
      };
    }

    console.log('📊 Result:', result);

    if (result.success && result.sessionId) {
      console.log('✅ Login successful!');
      result.redirectUrl = ScriptApp.getService().getUrl() + '?page=dashboard&sessionId=' + result.sessionId;
    } else {
      console.log('❌ Login failed:', result.message);
    }

    return result;

  } catch (error) {
    console.error('❌ ========================================');
    console.error('❌ LOGIN ERROR');
    console.error('❌ ========================================');
    console.error('Error message:', error.message);
    console.error('Error stack:', error.stack);

    let message = error.message;
    if (message.includes('timeout')) {
      message = 'Request timed out. Master API is not responding. Check your connection and try again.';
    } else if (message.includes('authorization')) {
      message = 'Authorization required. The app needs additional permissions.';
    } else if (message.includes('404')) {
      message = 'Master API not found. The server URL may be incorrect.';
    } else if (message.includes('403')) {
      message = 'Access forbidden. Check the Master API permissions.';
    } else if (message.includes('Unexpected token')) {
      message = 'Server returned invalid data. The Master API may be down.';
    }

    return {
      success: false,
      message: message,
      details: error.stack
    };
  }
}

function getClientSettings() {
  return {
    clientName: CLIENT_CONFIG.clientName,
    primaryColor: CLIENT_CONFIG.primaryColor,
    logoUrl: CLIENT_CONFIG.logoUrl,
    customDomain: CLIENT_CONFIG.customDomain
  };
}

function doGet(e) {
  try {
    const params = e?.parameter || {};
    if (params.auth === '1') {
      return HtmlService.createHtmlOutput(\`
        <html>
          <head><title>Authorization</title></head>
          <body style="font-family: Arial; text-align: center; padding: 50px;">
            <h2 style="color: #7C3AED;">Authorizing...</h2>
            <p>Please grant permissions when prompted.</p>
            <script>
              google.script.run.getClientSettings();
            <\/script>
          </body>
        </html>
      \`).setTitle('Authorization');
    }
    const template = HtmlService.createTemplateFromFile('client-landing');
    template.settings = getClientSettings();
    template.clientId = CLIENT_CONFIG.clientId;
    template.clientName = CLIENT_CONFIG.clientName;
    return template.evaluate()
      .setTitle(CLIENT_CONFIG.clientName + ' - Business OS')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('Error serving page:', error);
    return HtmlService.createHtmlOutput('<h1>Error</h1><p>' + error.message + '</p>');
  }
}

function doPost(e) {
  // ✅ ADD ERROR HANDLING FIRST
  if (!e || !e.postData || !e.postData.contents) {
    console.error('❌ Invalid request: No postData received');
    return ContentService
      .createTextOutput(JSON.stringify({ 
        success: false, 
        message: 'Invalid request format' 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const data = JSON.parse(e.postData.contents);
    console.log('📨 doPost action:', data.action);
    
    // ✅ FIX: Change 'login' to 'clientLogin'
    if (data.action === 'clientLogin') {
      const result = authenticateClient(
        data.email, 
        data.password, 
        data.businessName  // ✅ Include businessName
      );
      console.log('📨 Returning result:', result);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (data.action === 'getSettings') {
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, settings: getClientSettings() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: 'Unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('❌ doPost error:', error);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

console.log('✅ ${config.clientName} client loaded');
  `;
}




// ============================================================
// 🏠 CREATE FALLBACK LANDING PAGE (No deployment needed)
// ============================================================

function createFallbackLandingPage(clientId, businessName, paymentData) {
  try {
    console.log('📄 Creating fallback landing page for:', businessName);
    
    // Get master URL
    const masterUrl = getMasterApiUrl();
    
    const primaryColor = paymentData.primaryColor || '#2E7D32';
    const logoUrl = paymentData.logoUrl || '';
    
    const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${businessName} - Login</title>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    :root { --primary: ${primaryColor}; --primary-dark: ${primaryColor}dd; }
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      background: linear-gradient(135deg, #f5f7fa 0%, #e4e8f0 100%);
      padding: 20px;
    }
    .container {
      max-width: 420px;
      width: 100%;
      padding: 40px 30px;
      background: white;
      border-radius: 16px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.1);
      text-align: center;
    }
    .logo-placeholder {
      width: 80px;
      height: 80px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 36px;
      font-weight: bold;
      color: white;
      margin: 0 auto 16px;
      background: var(--primary);
    }
    ${logoUrl ? `.logo-img { width: 80px; height: 80px; border-radius: 50%; object-fit: cover; margin: 0 auto 16px; display: block; }` : ''}
    h1 { font-size: 28px; font-weight: 700; color: var(--primary); margin-bottom: 4px; }
    .subtitle { color: #6b7280; font-size: 14px; }
    .form-group { margin-bottom: 16px; text-align: left; }
    .form-group label { display: block; font-size: 14px; font-weight: 500; color: #374151; margin-bottom: 6px; }
    .form-group input {
      width: 100%; padding: 12px 16px;
      border: 2px solid #e5e7eb; border-radius: 10px;
      font-size: 14px; outline: none;
    }
    .form-group input:focus { border-color: var(--primary); }
    .login-btn {
      width: 100%; padding: 14px;
      background: var(--primary); color: white;
      border: none; border-radius: 10px;
      font-size: 16px; font-weight: 600; cursor: pointer;
    }
    .login-btn:hover { background: var(--primary-dark); }
    .login-btn:disabled { opacity: 0.7; cursor: not-allowed; }
    .error { color: #ef4444; font-size: 14px; text-align: center; margin-top: 12px; display: none; padding: 10px; background: #fef2f2; border-radius: 8px; }
    .loading { display: none; text-align: center; margin-top: 12px; }
    .loading i { color: var(--primary); }
    .footer { margin-top: 24px; font-size: 12px; color: #9ca3af; }
  </style>
</head>
<body>
  <div class="container">
    ${logoUrl ? `<img src="${logoUrl}" alt="${businessName}" class="logo-img">` : `<div class="logo-placeholder">${businessName.charAt(0).toUpperCase()}</div>`}
    <h1>${businessName}</h1>
    <p class="subtitle">Business Operating System</p>
    
    <div style="margin-top: 30px;">
      <h2 style="font-size: 18px; font-weight: 600; color: #1f2937; text-align: center; margin-bottom: 20px;">Welcome Back</h2>
      <form onsubmit="handleLogin(event)">
        <div class="form-group">
          <label><i class="fas fa-envelope"></i> Email Address</label>
          <input type="email" id="login-email" placeholder="you@company.com" required>
        </div>
        <div class="form-group">
          <label><i class="fas fa-lock"></i> Password</label>
          <input type="password" id="login-password" placeholder="••••••••" required>
        </div>
        <button type="submit" class="login-btn" id="login-btn">
          <i class="fas fa-sign-in-alt"></i> Sign In
        </button>
        <div id="login-error" class="error"></div>
        <div class="loading" id="login-loading">
          <i class="fas fa-spinner fa-spin"></i> Signing in...
        </div>
      </form>
    </div>
    <div class="footer">Powered by BizOS</div>
  </div>
  
  <script>
    const MASTER_URL = '${masterUrl}';
    const CLIENT_ID = '${clientId}';
    
    function handleLogin(event) {
      event.preventDefault();
      
      const email = document.getElementById('login-email').value.trim();
      const password = document.getElementById('login-password').value;
      const errorEl = document.getElementById('login-error');
      const btn = document.getElementById('login-btn');
      const loading = document.getElementById('login-loading');
      
      errorEl.style.display = 'none';
      loading.style.display = 'block';
      btn.disabled = true;
      btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Signing in...';
      
      fetch(MASTER_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'clientLogin',
          clientId: CLIENT_ID,
          email: email,
          password: password,
          sheetId: '',
          businessId: CLIENT_ID
        })
      })
      .then(response => response.json())
      .then(result => {
        loading.style.display = 'none';
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-sign-in-alt"></i> Sign In';
        
        if (result.success) {
          localStorage.setItem('bizos_session', result.sessionId);
          localStorage.setItem('bizos_user', JSON.stringify(result.user));
          localStorage.setItem('bizos_client_name', result.clientName || '${businessName}');
          window.location.href = window.location.href + '?dashboard=true';
        } else {
          errorEl.textContent = result.message || 'Login failed';
          errorEl.style.display = 'block';
        }
      })
      .catch(error => {
        loading.style.display = 'none';
        btn.disabled = false;
        btn.innerHTML = '<i class="fas fa-sign-in-alt"></i> Sign In';
        errorEl.textContent = 'Network error. Please try again.';
        errorEl.style.display = 'block';
        console.error('Login error:', error);
      });
    }
  </script>
</body>
</html>`;
    
    // Save to Drive
    const folder = getOrCreateClientFolder(paymentData.email, businessName);
    const file = folder.createFile('landing.html', html, 'text/html');
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Return the file URL as the landing URL
    const fileUrl = file.getUrl();
    console.log('✅ Fallback landing page created:', fileUrl);
    
    return fileUrl;
    
  } catch (error) {
    console.error('❌ Fallback landing error:', error);
    // Ultimate fallback - return the master URL with client param
    return getMasterApiUrl() + '?client=' + clientId;
  }
}

// ============================================================
// 🚀 CREATE AND DEPLOY CLIENT SCRIPT
// ============================================================

function createAndDeployClientScript(email, clientId, code, businessName, primaryColor, logoUrl) {
  try {
    console.log('🚀 Creating client script for:', businessName);
    
    const landingHtml = getClientLandingHtml(clientId, businessName, primaryColor, logoUrl);
    const appsscriptJson = getAppScriptJson();
    
    // STEP 1: Create the script project
    const createUrl = 'https://script.googleapis.com/v1/projects';
    const createPayload = {
      title: businessName + ' - BizOS'
    };
    
    console.log('📤 Creating script project...');
    
    const createResponse = UrlFetchApp.fetch(createUrl, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(createPayload),
      muteHttpExceptions: true
    });
    
    const createStatus = createResponse.getResponseCode();
    
    if (createStatus !== 200) {
      const errorText = createResponse.getContentText();
      console.error('❌ Create project failed:', errorText);
      if (errorText.includes('403') || errorText.includes('permission')) {
        return { 
          success: false, 
          message: 'Apps Script API not enabled or insufficient permissions.' 
        };
      }
      return { success: false, message: 'Failed to create project: ' + createStatus };
    }
    
    const createResult = JSON.parse(createResponse.getContentText());
    const scriptId = createResult.scriptId;
    console.log('✅ Script created with ID:', scriptId);
    
    // STEP 2: Upload files
    const files = [
      { name: 'Code', type: 'SERVER_JS', source: code },
      { name: 'client-landing', type: 'HTML', source: landingHtml },
      // { name: 'client-dashboard', type: 'HTML', source: getClientDashboardHtml() },
      { name: 'appsscript', type: 'JSON', source: appsscriptJson }
    ];
    
    const updateUrl = `https://script.googleapis.com/v1/projects/${scriptId}/content`;
    console.log('📤 Uploading files...');
    
    const updateResponse = UrlFetchApp.fetch(updateUrl, {
      method: 'PUT',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ files: files }),
      muteHttpExceptions: true
    });
    
    if (updateResponse.getResponseCode() !== 200) {
      console.error('❌ Upload failed:', updateResponse.getContentText());
      return { success: false, message: 'Failed to upload files' };
    }
    console.log('✅ Files uploaded');
    
    // STEP 3: Create a version
    const versionUrl = `https://script.googleapis.com/v1/projects/${scriptId}/versions`;
    const versionResponse = UrlFetchApp.fetch(versionUrl, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ description: 'Initial version' }),
      muteHttpExceptions: true
    });
    
    if (versionResponse.getResponseCode() !== 200) {
      console.error('❌ Version creation failed:', versionResponse.getContentText());
      return { success: false, message: 'Failed to create version' };
    }
    
    const versionResult = JSON.parse(versionResponse.getContentText());
    const versionNumber = versionResult.versionNumber || 1;
    console.log('✅ Version created:', versionNumber);
    
    // STEP 4: Deploy as web app
    const deployUrl = `https://script.googleapis.com/v1/projects/${scriptId}/deployments`;
    
    const deployPayload = {
      versionNumber: versionNumber,
      manifestFileName: 'appsscript',
      description: 'Web app for ' + businessName
    };
    
    console.log('📤 Deploying as web app...');
    
    const deployResponse = UrlFetchApp.fetch(deployUrl, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(deployPayload),
      muteHttpExceptions: true
    });
    
    const deployStatus = deployResponse.getResponseCode();
    
    if (deployStatus !== 200) {
      console.error('❌ Deploy failed:', deployResponse.getContentText());
      return { 
        success: false, 
        scriptId: scriptId,
        editUrl: `https://script.google.com/d/${scriptId}/edit`,
        message: 'Deploy failed: ' + deployStatus
      };
    }
    
    const deployResult = JSON.parse(deployResponse.getContentText());
    const deploymentId = deployResult.deploymentId;
    console.log('✅ Deployed successfully!');
    console.log('📋 Deployment ID:', deploymentId);
    
    // STEP 5: Get the CORRECT web app URL using deployment ID
    const webAppUrl = `https://script.google.com/macros/s/${deploymentId}/exec`;
    console.log('🔗 CORRECT Web App URL:', webAppUrl);
    
    // STEP 6: Share script with client
    try {
      const file = DriveApp.getFileById(scriptId);
      file.addEditor(email);
      console.log('📤 Script shared with:', email);
    } catch (shareError) {
      console.log('⚠️ Could not share script:', shareError.message);
    }
    

    // STEP 7: Wait for deployment to propagate
    console.log('⏳ Waiting 10 seconds for deployment to propagate...');
    Utilities.sleep(10000);
    
    return {
      success: true,
      scriptId: scriptId,
      deploymentId: deploymentId,
      webAppUrl: webAppUrl,
      editUrl: `https://script.google.com/d/${scriptId}/edit`,
      message: 'Client script deployed as web app'
    };
    
  } catch (error) {
    console.error('❌ Create/deploy error:', error);
    return { success: false, message: error.message };
  }
}

// ============================================================
// 📝 GENERATE CLIENT CODE SAFELY - COMPLETE FIXED VERSION
// ============================================================

function generateClientCodeSafely(config) {
  const masterUrl = config.masterApiUrl || getMasterApiUrl();

  return `// ============================================================
// 📁 Client App.gs - ${config.clientName}
// Client ID: ${config.clientId}
// Created: ${new Date().toISOString()}
// ============================================================

const CLIENT_CONFIG = {
  clientId: '${config.clientId}',
  clientName: '${config.clientName}',
  primaryColor: '${config.primaryColor || '#1a5276'}',
  logoUrl: '${config.logoUrl || ''}',
  customDomain: '${config.customDomain || ''}',
  sheetId: '${config.sheetId}',
  email: '${config.email}',
  businessId: '${config.businessId || config.clientId}',
  masterApiUrl: '${masterUrl}'
};

// ============================================================
// 🔐 AUTHENTICATE CLIENT
// ============================================================
function authenticateClient(email, password, businessName) {
  try {
    if (!email || !password || !businessName) {
      return { success: false, message: 'Email, password, and business name are required' };
    }

    var payload = {
      action: 'clientLogin',
      clientId: CLIENT_CONFIG.clientId,
      email: email.trim(),
      password: password,
      sheetId: CLIENT_CONFIG.sheetId,
      businessId: CLIENT_CONFIG.businessId,
      businessName: businessName.trim()
    };

    var response = UrlFetchApp.fetch(CLIENT_CONFIG.masterApiUrl, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 30000
    });

    var responseText = response.getContentText();
    if (responseText.trim().startsWith('<')) {
      return { success: false, message: 'Server error: Master API not responding correctly.' };
    }
    if (response.getResponseCode() !== 200) {
      return { success: false, message: 'Server error: ' + response.getResponseCode() };
    }

    var result;
    try {
      result = JSON.parse(responseText);
    } catch (e) {
      return { success: false, message: 'Server returned invalid response' };
    }

    // ✅ Just return the result - HTML will handle showing dashboard
    return result;

  } catch (error) {
    return { success: false, message: 'Login error: ' + error.message };
  }
}

// ============================================================
// 📁 GET CLIENT SETTINGS
// ============================================================
function getClientSettings() {
  return {
    clientName: CLIENT_CONFIG.clientName,
    primaryColor: CLIENT_CONFIG.primaryColor,
    logoUrl: CLIENT_CONFIG.logoUrl,
    customDomain: CLIENT_CONFIG.customDomain
  };
}

// ============================================================
// 📁 PROXY FUNCTIONS - Call Master API
// ============================================================

function proxyCheckSession(sessionId) {
  try {
    var response = UrlFetchApp.fetch(CLIENT_CONFIG.masterApiUrl, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({ 
        action: 'checkClientSession', 
        sessionId: sessionId,
        clientId: CLIENT_CONFIG.clientId 
      }),
      muteHttpExceptions: true,
      timeout: 30000
    });
    var text = response.getContentText();
    if (text.trim().startsWith('<')) return { success: false, message: 'Master API error' };
    return JSON.parse(text);
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function proxyGetDashboard(sessionId) {
  try {
    var response = UrlFetchApp.fetch(CLIENT_CONFIG.masterApiUrl, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({ 
        action: 'getClientDashboard', 
        sessionId: sessionId,
        clientId: CLIENT_CONFIG.clientId 
      }),
      muteHttpExceptions: true,
      timeout: 30000
    });
    var text = response.getContentText();
    if (text.trim().startsWith('<')) return { success: false, message: 'Master API error' };
    return JSON.parse(text);
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function proxyLogout(sessionId) {
  try {
    UrlFetchApp.fetch(CLIENT_CONFIG.masterApiUrl, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({ 
        action: 'clientLogout', 
        sessionId: sessionId,
        clientId: CLIENT_CONFIG.clientId 
      }),
      muteHttpExceptions: true,
      timeout: 15000
    });
  } catch (e) {}
  return { success: true };
}

// ============================================================
// 📁 DO GET - Serve combined landing + dashboard
// ============================================================
function doGet(e) {
  try {
    var params = e?.parameter || {};

    // Handle authorization callback
    if (params.auth === '1') {
      return HtmlService.createHtmlOutput(
        '<html><body style="font-family: Arial; text-align: center; padding: 50px;">' +
        '<h2 style="color: #5D2A86;">✅ Authorization Successful</h2>' +
        '<p>You can now close this tab and try logging in again.</p>' +
        '</body></html>'
      ).setTitle('Authorization Complete');
    }

    // Serve the landing page (contains both login AND dashboard)
    var template = HtmlService.createTemplateFromFile('client-landing');
    template.settings = getClientSettings();
    template.clientId = CLIENT_CONFIG.clientId;
    template.clientName = CLIENT_CONFIG.clientName;

    return template.evaluate()
      .setTitle(CLIENT_CONFIG.clientName + ' - Business OS')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    console.error('doGet error:', error);
    return HtmlService.createHtmlOutput('<h1>Error</h1><p>' + error.message + '</p>');
  }
}

// ============================================================
// 📁 DO POST - Handle login and other actions
// ============================================================
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: false, 
        message: 'Invalid request' 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var data = JSON.parse(e.postData.contents);

    if (data.action === 'login') {
      var result = authenticateClient(data.email, data.password, data.businessName);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'getSettings') {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        settings: getClientSettings() 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ 
      success: false, 
      message: 'Unknown action' 
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('doPost error:', error);
    return ContentService.createTextOutput(JSON.stringify({ 
      success: false, 
      message: error.message 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

console.log('✅ Client app loaded successfully');
`;
}

// ============================================================
// 🔍 CHECK FOR DUPLICATE CLIENT (NEW - PREVENTS DUPLICATES)
// ============================================================
function checkDuplicateClient(email, businessName) {
  try {
    const ss = getBizOSMasterSpreadsheet_();
    const sheet = ss.getSheetByName('Clients');
    
    if (!sheet) {
      return { isDuplicate: false };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    const nameCol = headers.indexOf('Client_Name');

    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][emailCol];
      const rowName = data[i][nameCol];

      // Check both email AND business name
      if (rowEmail && rowEmail.toLowerCase() === email.toLowerCase()) {
        console.log('❌ Duplicate email found:', email);
        return {
          isDuplicate: true,
          type: 'EMAIL',
          message: 'A client with this email already exists',
          existingClient: rowName
        };
      }

      if (rowName && rowName.toLowerCase() === businessName.toLowerCase()) {
        console.log('❌ Duplicate business name found:', businessName);
        return {
          isDuplicate: true,
          type: 'BUSINESS_NAME',
          message: 'A client with this business name already exists',
          existingEmail: rowEmail
        };
      }
    }

    return { isDuplicate: false };

  } catch (error) {
    console.error('❌ checkDuplicateClient error:', error);
    return { isDuplicate: false, error: error.message };
  }
}
// ============================================================
// 📄 GET CLIENT LANDING + DASHBOARD HTML
// ============================================================

function getClientLandingHtml(clientId, businessName, primaryColor, logoUrl) {
  const color = primaryColor || '#1a5276';
  const masterUrl = getMasterApiUrl();

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${businessName} - Business OS</title>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    :root { --primary: ${color}; --primary-dark: ${color}dd; }
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      min-height: 100vh;
      background: #f5f7fa;
    }
    
    /* ===== LANDING PAGE STYLES ===== */
    #landing-page {
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      background: linear-gradient(135deg, #f5f7fa 0%, #e4e8f0 100%);
      padding: 20px;
    }
    #landing-page .container {
      max-width: 420px;
      width: 100%;
      padding: 40px 30px;
      background: white;
      border-radius: 16px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.1);
      text-align: center;
    }
    #landing-page .logo-placeholder {
      width: 80px;
      height: 80px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 36px;
      font-weight: bold;
      color: white;
      margin: 0 auto 16px;
      background: var(--primary);
    }
    #landing-page h1 { font-size: 28px; font-weight: 700; color: var(--primary); margin-bottom: 4px; }
    #landing-page .subtitle { color: #6b7280; font-size: 14px; }
    #landing-page .login-form { text-align: left; margin-top: 30px; }
    #landing-page .login-form h2 { font-size: 18px; font-weight: 600; color: #1f2937; margin-bottom: 20px; text-align: center; }
    #landing-page .form-group { margin-bottom: 16px; }
    #landing-page .form-group label { display: block; font-size: 14px; font-weight: 500; color: #374151; margin-bottom: 6px; }
    #landing-page .form-group input {
      width: 100%; padding: 12px 16px; border: 2px solid #e5e7eb; border-radius: 10px;
      font-size: 14px; outline: none; transition: border-color 0.2s;
    }
    #landing-page .form-group input:focus { border-color: var(--primary); }
    #landing-page .login-btn {
      width: 100%; padding: 14px; background: var(--primary); color: white;
      border: none; border-radius: 10px; font-size: 16px; font-weight: 600;
      cursor: pointer; transition: background 0.2s;
    }
    #landing-page .login-btn:hover { background: var(--primary-dark); }
    #landing-page .login-btn:disabled { opacity: 0.7; cursor: not-allowed; }
    #landing-page .error { color: #ef4444; font-size: 14px; text-align: center; margin-top: 12px; display: none; padding: 10px; background: #fef2f2; border-radius: 8px; }
    #landing-page .success { display: none; color: #10b981; font-size: 14px; text-align: center; margin-top: 12px; padding: 10px; background: #f0fdf4; border-radius: 8px; }
    #landing-page .loading { display: none; text-align: center; margin-top: 12px; }
    #landing-page .loading i { color: var(--primary); }
    #landing-page .powered-by { margin-top: 24px; font-size: 12px; color: #9ca3af; text-align: center; }
    #landing-page .powered-by a { color: var(--primary); text-decoration: none; }
    
    /* ===== DASHBOARD STYLES ===== */
    #dashboard-page {
      display: none;
      min-height: 100vh;
      background: #f5f7fa;
      padding: 20px;
    }
    #dashboard-page .header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      background: white;
      padding: 20px;
      border-radius: 12px;
      margin-bottom: 20px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    #dashboard-page .header h1 { font-size: 20px; color: #1f2937; }
    #dashboard-page .header .client-badge {
      background: var(--primary);
      color: white;
      padding: 4px 12px;
      border-radius: 20px;
      font-size: 12px;
    }
    #dashboard-page .logout-btn {
      padding: 8px 16px;
      background: #ef4444;
      color: white;
      border: none;
      border-radius: 8px;
      cursor: pointer;
      font-size: 14px;
    }
    #dashboard-page .logout-btn:hover { background: #dc2626; }
    #dashboard-page .kpis {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 16px;
      margin-bottom: 20px;
    }
    #dashboard-page .kpi-card {
      background: white;
      padding: 20px;
      border-radius: 12px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    #dashboard-page .kpi-card .label { font-size: 13px; color: #6b7280; margin-bottom: 6px; }
    #dashboard-page .kpi-card .value { font-size: 24px; font-weight: 700; color: #1f2937; }
    #dashboard-page .modules {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
      gap: 12px;
      margin-top: 12px;
    }
    #dashboard-page .module-card {
      background: white;
      padding: 16px;
      border-radius: 10px;
      text-align: center;
      box-shadow: 0 2px 8px rgba(0,0,0,0.05);
      cursor: pointer;
      transition: transform 0.2s, box-shadow 0.2s;
    }
    #dashboard-page .module-card:hover {
      transform: translateY(-2px);
      box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    #dashboard-page .module-card .icon { font-size: 24px; margin-bottom: 8px; display: block; }
    #dashboard-page .module-card .name { font-weight: 600; font-size: 14px; color: #1f2937; }
    #dashboard-page .module-card .count { font-size: 12px; color: #9ca3af; margin-top: 4px; }
    #dashboard-page .loading { text-align: center; padding: 60px; color: #6b7280; }
    #dashboard-page .error-msg { text-align: center; padding: 40px; color: #ef4444; }
    #dashboard-page .section-title {
      font-size: 18px;
      font-weight: 600;
      color: #1f2937;
      margin: 20px 0 12px 0;
    }
    #dashboard-page .transactions {
      background: white;
      border-radius: 12px;
      overflow: hidden;
      box-shadow: 0 2px 8px rgba(0,0,0,0.05);
      margin-top: 12px;
    }
    #dashboard-page .transactions table { width: 100%; border-collapse: collapse; }
    #dashboard-page .transactions th {
      text-align: left;
      padding: 12px 16px;
      background: #f9fafb;
      font-size: 12px;
      font-weight: 600;
      color: #6b7280;
      text-transform: uppercase;
    }
    #dashboard-page .transactions td {
      padding: 12px 16px;
      border-top: 1px solid #f3f4f6;
      font-size: 14px;
    }
    #dashboard-page .transactions .amount.positive { color: #10b981; }
    #dashboard-page .transactions .amount.negative { color: #ef4444; }
    #dashboard-page .hidden { display: none !important; }
  </style>
</head>
<body>

  <!-- ===== LANDING PAGE ===== -->
  <div id="landing-page">
    <div class="container">
      <div class="logo-placeholder">${businessName.charAt(0).toUpperCase()}</div>
      <h1>${businessName}</h1>
      <p class="subtitle">Business Operating System</p>
      
      <div class="login-form">
        <h2>Sign In</h2>
        <form onsubmit="handleLogin(event)">
          <div class="form-group">
            <label><i class="fas fa-building"></i> Business Name</label>
            <input type="text" id="login-business-name" placeholder="Enter your business name" required>
          </div>
          <div class="form-group">
            <label><i class="fas fa-envelope"></i> Email Address</label>
            <input type="email" id="login-email" placeholder="you@company.com" required>
          </div>
          <div class="form-group">
            <label><i class="fas fa-lock"></i> Password</label>
            <input type="password" id="login-password" placeholder="••••••••" required>
          </div>
          <button type="submit" class="login-btn" id="login-btn">
            <i class="fas fa-sign-in-alt"></i> Sign In
          </button>
          <div id="login-error" class="error"></div>
          <div id="login-success" class="success">
            <i class="fas fa-check-circle"></i> Login successful! Loading dashboard...
          </div>
          <div class="loading" id="login-loading">
            <i class="fas fa-spinner fa-spin"></i> Signing in...
          </div>
        </form>
      </div>
      
      <div class="powered-by">Powered by <a href="#">BizOS</a></div>
    </div>
  </div>

  <!-- ===== DASHBOARD PAGE ===== -->
  <div id="dashboard-page">
    <div class="header">
      <div>
        <h1 id="client-name">${businessName}</h1>
        <span class="client-badge">Client Dashboard</span>
      </div>
      <button class="logout-btn" onclick="doLogout()">Logout</button>
    </div>
    
    <div id="dashboard-content" class="loading">
      <i class="fas fa-spinner fa-spin" style="font-size:32px;"></i>
      <p style="margin-top:16px;">Loading your dashboard...</p>
    </div>
  </div>

  <script>
    // ============================================================
    // HANDLE LOGIN
    // ============================================================
    
    function handleLogin(event) {
      event.preventDefault();
      event.stopPropagation();
      
      const businessName = document.getElementById('login-business-name').value.trim();
      const email = document.getElementById('login-email').value.trim();
      const password = document.getElementById('login-password').value;
      const errorEl = document.getElementById('login-error');
      const successEl = document.getElementById('login-success');
      const btn = document.getElementById('login-btn');
      const loading = document.getElementById('login-loading');
      
      errorEl.style.display = 'none';
      errorEl.textContent = '';
      successEl.style.display = 'none';
      
      if (!businessName || !email || !password) {
        errorEl.textContent = 'All fields are required';
        errorEl.style.display = 'block';
        return;
      }
      
      loading.style.display = 'block';
      btn.disabled = true;
      btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Signing in...';
      
      console.log('🔐 Login attempt:', { businessName, email });
      
      google.script.run
        .withSuccessHandler(function(result) {
          console.log('✅ Login response:', result);
          
          loading.style.display = 'none';
          btn.disabled = false;
          btn.innerHTML = '<i class="fas fa-sign-in-alt"></i> Sign In';
          
          if (result && result.success && result.sessionId) {
            // ✅ Store session
            localStorage.setItem('bizos_client_session', result.sessionId);
            localStorage.setItem('bizos_client_user', JSON.stringify(result.user));
            localStorage.setItem('bizos_client_name', result.clientName || '${businessName}');
            localStorage.setItem('bizos_client_id', '${clientId}');
            
            // ✅ Show success message
            successEl.style.display = 'block';
            
            // ✅ Switch to dashboard
            setTimeout(function() {
              showDashboard(result.sessionId);
            }, 500);
            
          } else {
            errorEl.textContent = result?.message || 'Login failed. Please try again.';
            errorEl.style.display = 'block';
          }
        })
        .withFailureHandler(function(error) {
          console.error('❌ Login error:', error);
          loading.style.display = 'none';
          btn.disabled = false;
          btn.innerHTML = '<i class="fas fa-sign-in-alt"></i> Sign In';
          errorEl.textContent = 'Error: ' + (error.message || 'Unknown error');
          errorEl.style.display = 'block';
        })
        .authenticateClient(email, password, businessName);
    }
    
    // ============================================================
    // SHOW DASHBOARD
    // ============================================================
    
    function showDashboard(sessionId) {
      // Hide landing page, show dashboard
      document.getElementById('landing-page').style.display = 'none';
      document.getElementById('dashboard-page').style.display = 'block';
      
      // Load dashboard data
      loadDashboardData(sessionId);
    }
    
    function loadDashboardData(sessionId) {
      var content = document.getElementById('dashboard-content');
      content.innerHTML = '<div class="loading"><i class="fas fa-spinner fa-spin" style="font-size:32px;"></i><p style="margin-top:16px;">Loading your data...</p></div>';
      
      google.script.run
        .withSuccessHandler(function(result) {
          console.log('📊 Dashboard data:', result);
          
          if (!result || !result.success) {
            content.innerHTML = '<div class="error-msg">Failed to load dashboard. <a href="#" onclick="location.reload()">Try again</a></div>';
            return;
          }
          
          renderDashboard(result.dashboard);
        })
        .withFailureHandler(function(error) {
          console.error('❌ Dashboard error:', error);
          content.innerHTML = '<div class="error-msg">Failed to load dashboard. Please try again.</div>';
        })
        .proxyGetDashboard(sessionId);
    }
    
    function renderDashboard(d) {
      var content = document.getElementById('dashboard-content');
      
      // Build KPIs
      var html = '<div class="kpis">';
      html += '<div class="kpi-card"><div class="label">Revenue</div><div class="value">$' + (d.kpis?.revenue || 0).toLocaleString() + '</div></div>';
      html += '<div class="kpi-card"><div class="label">Expenses</div><div class="value">$' + (d.kpis?.expenses || 0).toLocaleString() + '</div></div>';
      html += '<div class="kpi-card"><div class="label">Profit</div><div class="value">$' + (d.kpis?.profit || 0).toLocaleString() + '</div></div>';
      html += '<div class="kpi-card"><div class="label">Margin</div><div class="value">' + (d.kpis?.margin || 0) + '%</div></div>';
      html += '</div>';
      
      // Build Modules
      html += '<div class="section-title"><i class="fas fa-th-large" style="margin-right:8px;"></i>Your Modules</div>';
      html += '<div class="modules">';
      if (d.modules && d.modules.length > 0) {
        d.modules.forEach(function(m) {
          html += '<div class="module-card">';
          html += '<span class="icon"><i class="fas fa-' + (m.icon || 'folder') + '"></i></span>';
          html += '<div class="name">' + m.label + '</div>';
          html += '<div class="count">' + m.count + ' records</div>';
          html += '</div>';
        });
      } else {
        html += '<div class="module-card" style="grid-column:span 4;">No modules available</div>';
      }
      html += '</div>';
      
      // Build Recent Transactions
      html += '<div class="section-title" style="margin-top:24px;"><i class="fas fa-clock" style="margin-right:8px;"></i>Recent Transactions</div>';
      html += '<div class="transactions">';
      if (d.recentTransactions && d.recentTransactions.length > 0) {
        html += '<table><thead><tr><th>Date</th><th>Description</th><th>Type</th><th style="text-align:right;">Amount</th></tr></thead><tbody>';
        d.recentTransactions.forEach(function(t) {
          var isPositive = t.type === 'revenue' || t.type === 'income';
          html += '<tr>';
          html += '<td>' + (t.date || '') + '</td>';
          html += '<td>' + (t.description || '') + '</td>';
          html += '<td>' + (t.type || '') + '</td>';
          html += '<td class="amount ' + (isPositive ? 'positive' : 'negative') + '" style="text-align:right;">$' + (t.amount || 0).toLocaleString() + '</td>';
          html += '</tr>';
        });
        html += '</tbody></table>';
      } else {
        html += '<div style="padding:20px;text-align:center;color:#9ca3af;">No recent transactions</div>';
      }
      html += '</div>';
      
      content.innerHTML = html;
    }
    
    // ============================================================
    // LOGOUT
    // ============================================================
    
    function doLogout() {
      var sessionId = localStorage.getItem('bizos_client_session');
      if (sessionId) {
        google.script.run.proxyLogout(sessionId);
      }
      localStorage.removeItem('bizos_client_session');
      localStorage.removeItem('bizos_client_user');
      localStorage.removeItem('bizos_client_name');
      
      // Show landing page, hide dashboard
      document.getElementById('landing-page').style.display = 'flex';
      document.getElementById('dashboard-page').style.display = 'none';
      
      // Clear form
      document.getElementById('login-business-name').value = '';
      document.getElementById('login-email').value = '';
      document.getElementById('login-password').value = '';
    }
    
    // ============================================================
    // CHECK EXISTING SESSION ON LOAD
    // ============================================================
    
    (function checkExistingSession() {
      var sessionId = localStorage.getItem('bizos_client_session');
      var clientName = localStorage.getItem('bizos_client_name');
      
      if (sessionId && clientName) {
        console.log('🔄 Found existing session, loading dashboard...');
        // Update client name
        document.getElementById('client-name').textContent = clientName;
        // Show dashboard
        showDashboard(sessionId);
      }
    })();
    
    console.log('✅ Client app loaded');
    console.log('🆔 Client ID:', '${clientId}');
    console.log('🏢 Client Name:', '${businessName}');
  </script>
</body>
</html>`;
}

// ============================================================
// 📋 GET APPSCRIPT JSON
// ============================================================

function getAppScriptJson() {
  return JSON.stringify({
    "timeZone": "Africa/Lagos",
    "dependencies": {},
    "exceptionLogging": "STACKDRIVER",
    "runtimeVersion": "V8",
    "webapp": {
      "executeAs": "USER_DEPLOYING",
      "access": "ANYONE_ANONYMOUS"
    },
    "oauthScopes": [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive.file",
      "https://www.googleapis.com/auth/script.send_mail",
      "https://www.googleapis.com/auth/userinfo.email",
      "https://www.googleapis.com/auth/script.external_request",
      "https://www.googleapis.com/auth/drive",
      "https://www.googleapis.com/auth/script.scriptapp",
      "https://www.googleapis.com/auth/script.deployments",
      "https://www.googleapis.com/auth/script.projects"
    ]
  });
}

// ============================================================
// 🔍 GET BUSINESS BY EMAIL
// ============================================================

function getBusinessByEmail(email) {
  try {
    const ss = getBizOSMasterSpreadsheet_();
    const sheet = ss.getSheetByName('Businesses');
    if (!sheet) {
      console.log('⚠️ Businesses sheet not found');
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      console.log('⚠️ Businesses sheet has no data');
      return null;
    }
    
    const headers = data[0];
    const emailCol = headers.indexOf('Owner_Email');
    const idCol = headers.indexOf('Business_ID');
    const nameCol = headers.indexOf('Business_Name');
    const tierCol = headers.indexOf('Subscription_Tier');
    
    if (emailCol === -1) {
      console.log('⚠️ Owner_Email column not found in Businesses sheet');
      return null;
    }
    
    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][emailCol];
      if (rowEmail && rowEmail.toLowerCase() === email.toLowerCase()) {
        console.log(`✅ Found business for email: ${email}`);
        return {
          businessId: data[i][idCol] || '',
          businessName: data[i][nameCol] || '',
          subscriptionTier: data[i][tierCol] || 'free',
          workspaceId: workspaceCol !== -1 ? (data[i][workspaceCol] || '') : ''
        };
      }
    }
    
    console.log(`ℹ️ No business found for email: ${email}`);
    return null;
  } catch (error) {
    console.error('❌ getBusinessByEmail error:', error);
    return null;
  }
}

// ============================================================
// 🔄 UPDATE BUSINESS WITH CLIENT INFO
// ============================================================

function updateBusinessWithClientInfo(businessId, clientId, landingUrl) {
  try {
    const ss = getBizOSMasterSpreadsheet_();
    const sheet = ss.getSheetByName('Businesses');
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('Business_ID');
    const clientIdCol = headers.indexOf('Client_ID');
    const landingUrlCol = headers.indexOf('Landing_URL');
    const tierCol = headers.indexOf('Subscription_Tier');
    
    if (clientIdCol === -1) {
      sheet.appendRow([...headers, 'Client_ID', 'Landing_URL']);
      const newData = sheet.getDataRange().getValues();
      const newHeaders = newData[0];
      const newClientIdCol = newHeaders.indexOf('Client_ID');
      const newLandingUrlCol = newHeaders.indexOf('Landing_URL');
      for (let i = 1; i < newData.length; i++) {
        if (newData[i][idCol] === businessId) {
          sheet.getRange(i + 1, newClientIdCol + 1).setValue(clientId);
          sheet.getRange(i + 1, newLandingUrlCol + 1).setValue(landingUrl);
          break;
        }
      }
      return;
    }
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === businessId) {
        if (clientIdCol !== -1) sheet.getRange(i + 1, clientIdCol + 1).setValue(clientId);
        if (landingUrlCol !== -1) sheet.getRange(i + 1, landingUrlCol + 1).setValue(landingUrl);
        if (tierCol !== -1) sheet.getRange(i + 1, tierCol + 1).setValue('sovereign');
        break;
      }
    }
    
    console.log('✅ Business updated with client info');
  } catch (error) {
    console.error('Update business error:', error);
  }
}

// ============================================================
// 📧 SEND WELCOME EMAIL
// ============================================================

function sendClientWelcomeEmail(email, businessName, landingUrl, clientId) {
  try {
    const subject = `🎉 Welcome to ${businessName}'s BizOS Workspace!`;
    
    const body = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h1 style="color: #1a5276;">Welcome to BizOS!</h1>
        <p>Your business <strong>${businessName}</strong> has been successfully set up.</p>
        
        <h2>Your Branded Landing Page</h2>
        <p>Access your business portal at:</p>
        <p><strong><a href="${landingUrl}" target="_blank">${landingUrl}</a></strong></p>
        
        <h2>Login Credentials</h2>
        <p><strong>Email:</strong> ${email}</p>
        <p><strong>Password:</strong> Your existing password (same as before)</p>
        
        <h2>What's Next?</h2>
        <ul>
          <li>Visit your branded landing page</li>
          <li>Login with your email and password</li>
          <li>Access all 12 business modules</li>
          <li>Your data is stored securely in your own Google Sheet</li>
        </ul>
        
        <hr>
        <p style="color: #666; font-size: 12px;">This is an automated message from BizOS.</p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: body
    });
    
    console.log('📧 Welcome email sent to:', email);
  } catch (error) {
    console.error('Welcome email error:', error);
  }
}

// ============================================================
// 📁 GET OR CREATE CLIENT FOLDER
// ============================================================

function getOrCreateClientFolder(email, businessName) {
  try {
    const folders = DriveApp.getFoldersByName('BizOS_Client_Workspaces');
    let folder;
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder('BizOS_Client_Workspaces');
    }
    
    const cleanName = businessName.replace(/[^a-zA-Z0-9]/g, '_');
    const subfolderName = cleanName + '_' + email.split('@')[0];
    
    const subFolders = folder.getFoldersByName(subfolderName);
    if (subFolders.hasNext()) {
      return subFolders.next();
    } else {
      return folder.createFolder(subfolderName);
    }
    
  } catch (error) {
    console.error('Folder error:', error);
    return DriveApp.createFolder('BizOS_' + businessName.replace(/[^a-zA-Z0-9]/g, '_'));
  }
}

console.log('✅ Master Backend loaded');
console.log('🔗 Master API URL:', getMasterApiUrl());