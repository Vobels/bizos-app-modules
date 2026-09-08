// ============================================================
// 📁 Tests.gs - SMART TEST SUITE WITH FORCE REGENERATION
// ============================================================

/**
 * ✅ CLEANUP CLIENT - Delete from Clients sheet
 */
function cleanupClient() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Clients');
    if (!sheet) {
      console.log('No Clients sheet found');
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const clientIdCol = headers.indexOf('Client_ID');
    const emailCol = headers.indexOf('Email');
    
    let deleted = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (!row || !row[emailCol]) continue;
      
      const rowEmail = String(row[emailCol]).toLowerCase().trim();
      if (rowEmail === 'otpglobal.ltd@gmail.com') {
        sheet.deleteRow(i + 1);
        console.log('✅ Deleted client:', row[clientIdCol] || 'Unknown');
        deleted++;
      }
    }
    
    console.log('✅ Cleanup complete. Deleted:', deleted, 'client(s)');
  } catch (error) {
    console.error('Cleanup error:', error);
  }
}

/**
 * ✅ DELETE CLIENT SCRIPT FROM DRIVE
 */
function deleteClientScriptFromDrive() {
  try {
    const searchQuery = "title contains 'Otp global company - BizOS' and mimeType contains 'script'";
    const files = DriveApp.searchFiles(searchQuery);
    
    let deleted = 0;
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const fileId = file.getId();
      
      try {
        file.removeEditor('otpglobal.ltd@gmail.com');
        console.log('  ✅ Removed editor from:', fileName);
      } catch (e) {
        console.log('  ⚠️ Could not remove editor:', e.message);
      }
      
      file.setTrashed(true);
      console.log('  ✅ Deleted from Drive:', fileName, '(ID:', fileId + ')');
      deleted++;
    }
    
    console.log('✅ Deleted', deleted, 'script(s) from Drive');
    return deleted;
  } catch (error) {
    console.error('Error deleting script:', error);
    return 0;
  }
}

/**
 * ✅ FULL CLEANUP - Delete client record AND script
 */
function fullCleanupOtpClient() {
  console.log('🧹 ========================================');
  console.log('🧹 FULL CLEANUP - OTP GLOBAL CLIENT');
  console.log('🧹 ========================================');
  
  console.log('\n📌 STEP 1: Deleting from Clients sheet...');
  cleanupClient();
  
  console.log('\n📌 STEP 2: Deleting from Google Drive...');
  deleteClientScriptFromDrive();
  
  console.log('\n✅ Full cleanup complete!');
  console.log('📌 Run setupOtpGlobalClient(true) to recreate.');
}

/**
 * ✅ SETUP OTP GLOBAL CLIENT - WITH FORCE REGENERATION
 * 
 * @param {boolean} force - If true, forces regeneration even if client exists
 */
function setupOtpGlobalClient(force = false) {
  console.log('🚀 ========================================');
  console.log('🚀 SETTING UP OTP GLOBAL CLIENT');
  console.log('🚀 ========================================');
  
  const email = 'otpglobal.ltd@gmail.com';
  const businessName = 'Otp global company';
  const domain = 'bizos.otpgloballimited.com';
  
  // STEP 1: Check if client already exists (skip if force=true)
  if (!force) {
    console.log('\n📌 STEP 1: Checking if client already exists...');
    const existingClient = findExistingClient(email);
    
    if (existingClient) {
      console.log('✅ Client already exists!');
      console.log('\n✅ ========================================');
      console.log('✅ OTP GLOBAL CLIENT FOUND');
      console.log('✅ ========================================');
      console.log('🆔 Client ID:', existingClient.clientId);
      console.log('📧 Email:', email);
      console.log('🏢 Business:', existingClient.clientName || businessName);
      console.log('🌐 Domain:', existingClient.domain || domain);
      console.log('📊 Sheet URL:', existingClient.sheetUrl);
      console.log('🔗 Web App URL (exec):', existingClient.webAppUrl);
      console.log('✏️  Edit URL:', existingClient.editUrl);
      console.log('📋 Script ID:', existingClient.scriptId);
      console.log('✅ ========================================');
      console.log('\n📝 LOGIN CREDENTIALS:');
      console.log('  Email:', email);
      console.log('  Password: otpglobal123');
      console.log('  Business Name:', existingClient.clientName || businessName);
      console.log('\n✅ ========================================');
      console.log('\n💡 To force regenerate, run: setupOtpGlobalClient(true)');
      
      return {
        success: true,
        alreadyExists: true,
        clientId: existingClient.clientId,
        email: email,
        password: 'otpglobal123',
        businessName: existingClient.clientName || businessName,
        webAppUrl: existingClient.webAppUrl,
        editUrl: existingClient.editUrl,
        scriptId: existingClient.scriptId,
        sheetUrl: existingClient.sheetUrl
      };
    }
  } else {
    console.log('\n🔄 FORCE REGENERATION ENABLED');
    console.log('📌 STEP 1: Cleaning up existing client...');
    cleanupClient();
    console.log('📌 STEP 2: Deleting script from Drive...');
    deleteClientScriptFromDrive();
    console.log('✅ Cleanup complete. Proceeding with fresh setup...');
  }
  
  // STEP 2: Call existing autoSetupClient() function
  console.log('\n📌 STEP 3: Creating new client using autoSetupClient()...');
  
  const paymentData = {
    email: email,
    password: 'otpglobal123',
    businessName: businessName,
    domain: domain,
    customDomain: domain,
    primaryColor: '#1F2937',
    logoUrl: '',
    tier: 'sovereign'
  };
  
  const setupResult = autoSetupClient(paymentData);
  
  if (!setupResult || !setupResult.success) {
    console.error('❌ Client setup failed:', setupResult?.message);
    return {
      success: false,
      message: setupResult?.message || 'Setup failed'
    };
  }
  
  console.log('✅ Client setup successful!');
  
  // STEP 3: Get the script ID and edit URL
  console.log('\n📌 STEP 4: Retrieving script details...');
  
  const scriptId = setupResult.scriptId || getClientScriptId(setupResult.clientId);
  const editUrl = scriptId ? `https://script.google.com/d/${scriptId}/edit` : 'Not available';
  
  console.log('✅ Script ID:', scriptId || 'Not available (normal for fallback mode)');
  console.log('✅ Edit URL:', editUrl);
  
  // STEP 4: Return complete information
  console.log('\n✅ ========================================');
  console.log('✅ OTP GLOBAL CLIENT CREATED SUCCESSFULLY!');
  console.log('✅ ========================================');
  console.log('🆔 Client ID:', setupResult.clientId);
  console.log('📧 Email:', email);
  console.log('🏢 Business:', businessName);
  console.log('🌐 Domain:', domain);
  console.log('📊 Sheet URL:', setupResult.sheetUrl);
  console.log('🔗 Web App URL (exec):', setupResult.landingUrl || setupResult.webAppUrl);
  console.log('✏️  Edit URL:', editUrl);
  console.log('📋 Script ID:', scriptId || 'N/A');
  console.log('✅ ========================================');
  console.log('\n📝 LOGIN CREDENTIALS:');
  console.log('  Email:', email);
  console.log('  Password: otpglobal123');
  console.log('  Business Name:', businessName);
  console.log('\n📌 URLS:');
  console.log('  🔗 User Login (exec):', setupResult.landingUrl || setupResult.webAppUrl);
  console.log('  ✏️  Dev/Edit (for admins):', editUrl);
  console.log('  📊 Sheet (data):', setupResult.sheetUrl);
  console.log('✅ ========================================');
  
  return {
    success: true,
    alreadyExists: false,
    clientId: setupResult.clientId,
    email: email,
    password: 'otpglobal123',
    businessName: businessName,
    webAppUrl: setupResult.landingUrl || setupResult.webAppUrl,
    editUrl: editUrl,
    scriptId: scriptId,
    sheetUrl: setupResult.sheetUrl
  };
}

/**
 * 🔍 FIND EXISTING CLIENT - Check if already setup
 * Returns full client info including business name
 */
function findExistingClient(email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Clients');
    
    if (!sheet) {
      console.log('  ℹ️ Clients sheet not found (first time setup)');
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const emailCol = headers.indexOf('Email');
    const idCol = headers.indexOf('Client_ID');
    const nameCol = headers.indexOf('Client_Name');
    const urlCol = headers.indexOf('Web_App_URL');
    const sheetIdCol = headers.indexOf('Sheet_ID');
    const scriptIdCol = headers.indexOf('Script_ID');
    const domainCol = headers.indexOf('Domain');
    const colorCol = headers.indexOf('Primary_Color');
    const tierCol = headers.indexOf('Tier');
    const statusCol = headers.indexOf('Status');
    
    if (emailCol === -1) {
      console.log('  ℹ️ Clients sheet exists but email column not found');
      return null;
    }
    
    // Check if client with this email exists
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || !row[emailCol]) continue;
      
      const rowEmail = String(row[emailCol]).toLowerCase().trim();
      const searchEmail = String(email).toLowerCase().trim();
      
      if (rowEmail === searchEmail) {
        console.log('  ✅ Found existing client with email:', email);
        
        const clientId = row[idCol] || '';
        const clientName = row[nameCol] || '';
        const webAppUrl = row[urlCol] || '';
        const sheetId = row[sheetIdCol] || '';
        const scriptId = row[scriptIdCol] || '';
        const domain = row[domainCol] || '';
        const primaryColor = row[colorCol] || '#1a5276';
        const tier = row[tierCol] || 'sovereign';
        const status = row[statusCol] || 'active';
        
        // Get sheet URL
        let sheetUrl = '';
        if (sheetId) {
          try {
            sheetUrl = SpreadsheetApp.openById(sheetId).getUrl();
          } catch (e) {
            console.log('  ⚠️ Could not open sheet:', e.message);
          }
        }
        
        return {
          clientId: clientId,
          clientName: clientName,
          webAppUrl: webAppUrl,
          sheetId: sheetId,
          scriptId: scriptId,
          sheetUrl: sheetUrl,
          domain: domain,
          primaryColor: primaryColor,
          tier: tier,
          status: status,
          editUrl: scriptId ? `https://script.google.com/d/${scriptId}/edit` : null
        };
      }
    }
    
    console.log('  ℹ️ No existing client found for:', email);
    return null;
    
  } catch (error) {
    console.error('  ❌ Error finding existing client:', error);
    return null;
  }
}

/**
 * 🔍 GET CLIENT SCRIPT ID - Retrieve from Clients sheet
 */
function getClientScriptId(clientId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Clients');
    
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('Client_ID');
    const scriptCol = headers.indexOf('Script_ID');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row) continue;
      if (row[idCol] === clientId && scriptCol !== -1) {
        return row[scriptCol] || null;
      }
    }
    
    return null;
  } catch (error) {
    console.error('Error getting script ID:', error);
    return null;
  }
}

// ============================================================
// 🧪 TEST FUNCTIONS
// ============================================================

/**
 * 🧪 TEST: Verify autoSetupClient() works
 */
function testAutoSetupClientFunction() {
  console.log('🧪 ========================================');
  console.log('🧪 TESTING autoSetupClient() FUNCTION');
  console.log('🧪 ========================================');
  
  const testEmail = 'test.client.' + Date.now() + '@example.com';
  const testData = {
    email: testEmail,
    password: 'Test123456',
    businessName: 'Test Business ' + Date.now(),
    primaryColor: '#7C3AED',
    logoUrl: '',
    customDomain: '',
    tier: 'sovereign'
  };
  
  console.log('\n📌 Calling autoSetupClient() with:');
  console.log('  Email:', testData.email);
  console.log('  Business:', testData.businessName);
  
  const result = autoSetupClient(testData);
  
  console.log('\n📊 Result:');
  console.log('  Success:', result.success ? '✅' : '❌');
  console.log('  Message:', result.message);
  console.log('  Client ID:', result.clientId || 'N/A');
  console.log('  Landing URL:', result.landingUrl || result.webAppUrl || 'N/A');
  console.log('  Sheet URL:', result.sheetUrl || 'N/A');
  
  if (result.success) {
    console.log('\n✅ autoSetupClient() works correctly!');
  } else {
    console.log('\n❌ autoSetupClient() failed');
  }
  
  return result;
}

/**
 * 🧪 TEST: Verify clientLogin() works
 */
function testClientLoginFunction() {
  console.log('🧪 ========================================');
  console.log('🧪 TESTING clientLogin() FUNCTION');
  console.log('🧪 ========================================');
  
  const email = 'otpglobal.ltd@gmail.com';
  const password = 'otpglobal123';
  
  // First check if client exists
  const existingClient = findExistingClient(email);
  
  if (!existingClient) {
    console.log('❌ OTP Global client not found');
    console.log('📌 Run setupOtpGlobalClient() first');
    return { success: false, message: 'Client not found' };
  }
  
  console.log('\n📌 Testing clientLogin() with:');
  console.log('  Client ID:', existingClient.clientId);
  console.log('  Email:', email);
  console.log('  Password:', password);
  
  const result = clientLogin(
    existingClient.clientId,
    email,
    password,
    existingClient.sheetId || '',
    existingClient.clientId,
    existingClient.clientName || 'Otp global company'
  );
  
  console.log('\n📊 Result:');
  console.log('  Success:', result.success ? '✅' : '❌');
  console.log('  Message:', result.message);
  console.log('  Session ID:', result.sessionId ? result.sessionId.substring(0, 20) + '...' : 'N/A');
  console.log('  User Email:', result.user?.email || 'N/A');
  console.log('  Is Client:', result.user?.isClient ? '✅' : '❌');
  
  if (result.success) {
    console.log('\n✅ clientLogin() works correctly!');
  } else {
    console.log('\n❌ clientLogin() failed');
  }
  
  return result;
}

/**
 * 🧪 TEST: Verify authenticateUser() works with client
 */
function testAuthenticateUserFunction() {
  console.log('🧪 ========================================');
  console.log('🧪 TESTING authenticateUser() FUNCTION');
  console.log('🧪 ========================================');
  
  const email = 'otpglobal.ltd@gmail.com';
  const password = 'otpglobal123';
  
  console.log('\n📌 Testing authenticateUser() with:');
  console.log('  Email:', email);
  console.log('  Password:', password);
  console.log('  Business Name: null (client users should skip this check)');
  
  const result = authenticateUser(email, password, null);
  
  console.log('\n📊 Result:');
  console.log('  Success:', result.success ? '✅' : '❌');
  console.log('  Message:', result.message);
  console.log('  User Email:', result.user?.email || 'N/A');
  console.log('  Is Client:', result.user?.isClient ? '✅' : '❌');
  console.log('  Session ID:', result.sessionId ? result.sessionId.substring(0, 20) + '...' : 'N/A');
  
  if (result.success) {
    console.log('\n✅ authenticateUser() works correctly!');
  } else {
    console.log('\n❌ authenticateUser() failed');
  }
  
  return result;
}

/**
 * 🧪 TEST: Verify getClientById() works
 */
function testGetClientByIdFunction() {
  console.log('🧪 ========================================');
  console.log('🧪 TESTING getClientById() FUNCTION');
  console.log('🧪 ========================================');
  
  const email = 'otpglobal.ltd@gmail.com';
  
  const existingClient = findExistingClient(email);
  
  if (!existingClient) {
    console.log('❌ OTP Global client not found');
    return { success: false, message: 'Client not found' };
  }
  
  console.log('\n📌 Testing getClientById() with:');
  console.log('  Client ID:', existingClient.clientId);
  
  const result = getClientById(existingClient.clientId);
  
  console.log('\n📊 Result:');
  console.log('  Found:', result ? '✅' : '❌');
  if (result) {
    console.log('  Client Name:', result.clientName || result.businessName || 'N/A');
    console.log('  Email:', result.email || 'N/A');
    console.log('  Status:', result.status || 'N/A');
    console.log('  Tier:', result.tier || 'N/A');
    console.log('  Web App URL:', result.webAppUrl ? '✅' : '❌');
    console.log('  Sheet ID:', result.sheetId ? '✅' : '❌');
  }
  
  if (result) {
    console.log('\n✅ getClientById() works correctly!');
  } else {
    console.log('\n❌ getClientById() failed');
  }
  
  return result;
}

// ============================================================
// 🧪 COMPLETE TEST SUITE
// ============================================================

function runAllTests() {
  console.log('🧪 ========================================');
  console.log('🧪 RUNNING COMPLETE TEST SUITE');
  console.log('🧪 ========================================');
  
  console.log('\n\n📌 TEST 1/5: Setup OTP Global Client');
  console.log('═'.repeat(50));
  const setupResult = setupOtpGlobalClient();
  
  console.log('\n\n📌 TEST 2/5: Verify getClientById()');
  console.log('═'.repeat(50));
  const getResult = testGetClientByIdFunction();
  
  console.log('\n\n📌 TEST 3/5: Verify authenticateUser()');
  console.log('═'.repeat(50));
  const authResult = testAuthenticateUserFunction();
  
  console.log('\n\n📌 TEST 4/5: Verify clientLogin()');
  console.log('═'.repeat(50));
  const loginResult = testClientLoginFunction();
  
  console.log('\n\n📌 TEST 5/5: Test autoSetupClient() (optional)');
  console.log('═'.repeat(50));
  console.log('⏭️  Skipped (only run if testing new deployment)');
  
  console.log('\n\n✅ ========================================');
  console.log('✅ TEST SUITE COMPLETE');
  console.log('✅ ========================================');
  console.log('\n📊 SUMMARY:');
  console.log('  ✅ Setup:', setupResult.success ? '✅ PASS' : '❌ FAIL');
  console.log('  ✅ getClientById:', getResult ? '✅ PASS' : '❌ FAIL');
  console.log('  ✅ authenticateUser:', authResult.success ? '✅ PASS' : '❌ FAIL');
  console.log('  ✅ clientLogin:', loginResult.success ? '✅ PASS' : '❌ FAIL');
  console.log('\n✅ ========================================');
}

// ============================================================
// 🚀 QUICK COMMANDS
// ============================================================

/*
RUN THESE TO TEST:

1. FORCE REGENERATE OTP GLOBAL CLIENT (USE THIS!):
   setupOtpGlobalClient(true);

2. FIRST TIME SETUP (creates client if doesn't exist):
   setupOtpGlobalClient();

3. FULL CLEANUP (delete client + script):
   fullCleanupOtpClient();

4. VERIFY ALL FUNCTIONS WORK:
   runAllTests();

5. TEST SPECIFIC FUNCTIONS:
   testGetClientByIdFunction();
   testAuthenticateUserFunction();
   testClientLoginFunction();

6. TEST NEW CLIENT SETUP (uses quota):
   testAutoSetupClientFunction();
*/