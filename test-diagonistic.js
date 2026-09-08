// ============================================================
// 🔍 DIAGNOSTIC: Test Code Generation & Deployment
// ============================================================
// Run this to verify the generated client code is valid

function testCodeGeneration() {
  try {
    console.log('🧪 ========================================');
    console.log('🧪 TESTING CODE GENERATION');
    console.log('🧪 ========================================');

    const testConfig = {
      clientId: 'CLIENT_TEST123',
      clientName: 'Test Business',
      primaryColor: '#5D2A86',
      logoUrl: '',
      customDomain: '',
      sheetId: 'SHEET123',
      email: 'test@example.com',
      businessId: 'BIZ123',
      masterApiUrl: 'https://example.com/api'
    };

    console.log('📝 Test Config:', JSON.stringify(testConfig, null, 2));

    // Generate code
    console.log('\n🔄 Generating client code...');
    const generatedCode = generateClientCodeSafely(testConfig);

    console.log('✅ Code generated, length:', generatedCode.length, 'characters');

    // Verify code contains required functions
    console.log('\n✅ CHECKING FOR REQUIRED FUNCTIONS:');
    
    const requiredFunctions = ['doGet', 'doPost', 'authenticateClient', 'getClientSettings', 'CLIENT_CONFIG'];
    let allFound = true;

    requiredFunctions.forEach(func => {
      const found = generatedCode.includes('function ' + func) || generatedCode.includes('const ' + func);
      const status = found ? '✅' : '❌';
      console.log(status + ' ' + func + ':', found);
      if (!found) allFound = false;
    });

    if (!allFound) {
      console.error('❌ CRITICAL: Generated code is MISSING required functions!');
      console.log('\n📋 First 500 characters of generated code:');
      console.log(generatedCode.substring(0, 500));
      return { success: false, message: 'Generated code missing required functions' };
    }

    console.log('\n✅ All required functions found!');

    // Check structure
    console.log('\n📋 CODE STRUCTURE:');
    console.log('- Starts with comment:', generatedCode.startsWith('//'));
    console.log('- Contains CLIENT_CONFIG:', generatedCode.includes('CLIENT_CONFIG'));
    console.log('- Contains authenticateClient:', generatedCode.includes('authenticateClient'));
    console.log('- Contains doGet:', generatedCode.includes('doGet'));
    console.log('- Contains doPost:', generatedCode.includes('doPost'));
    console.log('- Ends with console.log:', generatedCode.includes("console.log('✅ Client app loaded successfully')"));

    // Show first 1000 chars
    console.log('\n📖 Generated Code (first 1000 chars):');
    console.log(generatedCode.substring(0, 1000));

    console.log('\n📖 Generated Code (last 500 chars):');
    console.log(generatedCode.substring(generatedCode.length - 500));

    console.log('\n✅ ========================================');
    console.log('✅ CODE GENERATION TEST PASSED');
    console.log('✅ ========================================');

    return {
      success: true,
      codeLength: generatedCode.length,
      hasDoGet: generatedCode.includes('function doGet'),
      hasDoPost: generatedCode.includes('function doPost'),
      hasAuthenticate: generatedCode.includes('authenticateClient')
    };

  } catch (error) {
    console.error('❌ CODE GENERATION ERROR:', error.message);
    console.error('Stack:', error.stack);
    return { success: false, message: error.message, stack: error.stack };
  }
}

// ============================================================
// 🔍 CHECK LAST DEPLOYED CLIENT SCRIPT
// ============================================================

function checkLastDeployedClientScript() {
  try {
    console.log('🔍 ========================================');
    console.log('🔍 CHECKING LAST DEPLOYED CLIENT');
    console.log('🔍 ========================================');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const clientsSheet = ss.getSheetByName('Clients');

    if (!clientsSheet) {
      console.log('❌ Clients sheet not found');
      return { success: false, message: 'Clients sheet not found' };
    }

    const data = clientsSheet.getDataRange().getValues();
    if (data.length < 2) {
      console.log('❌ No clients in Clients sheet');
      return { success: false, message: 'No clients found' };
    }

    const headers = data[0];
    const scriptIdCol = headers.indexOf('Script_ID');
    const webAppCol = headers.indexOf('Web_App_URL');
    const emailCol = headers.indexOf('Email');

    if (scriptIdCol === -1) {
      console.log('❌ Script_ID column not found');
      return { success: false, message: 'Script_ID column missing' };
    }

    // Get last client
    const lastRow = data[data.length - 1];
    const scriptId = lastRow[scriptIdCol];
    const webAppUrl = lastRow[webAppCol];
    const email = lastRow[emailCol];

    console.log('📊 Last Client:');
    console.log('- Email:', email);
    console.log('- Web App URL:', webAppUrl);
    console.log('- Script ID:', scriptId);

    if (!scriptId || scriptId === '') {
      console.log('⚠️ No Script ID saved (likely deployment failed)');
      return { 
        success: false, 
        message: 'Last client has no Script ID - deployment may have failed',
        client: { email, webAppUrl, scriptId }
      };
    }

    try {
      const script = Apps.getScriptProject(scriptId);
      console.log('✅ Script project found:', scriptId);

      // Try to get the script content
      const content = script.getSourceCode();
      console.log('📝 Script content length:', content.length);
      console.log('- Contains doGet:', content.includes('doGet'));
      console.log('- Contains doPost:', content.includes('doPost'));
      console.log('- Contains CLIENT_CONFIG:', content.includes('CLIENT_CONFIG'));

      if (!content.includes('doGet')) {
        console.error('❌ PROBLEM: Script exists but has no doGet function!');
        console.log('\n📖 Current script content (first 500 chars):');
        console.log(content.substring(0, 500));
        return {
          success: false,
          message: 'Deployed script missing doGet function',
          details: { scriptId, hasDoGet: false, contentLength: content.length }
        };
      }

      return {
        success: true,
        message: 'Script is valid and has doGet',
        details: { scriptId, email, hasDoGet: true, hasDoPost: true }
      };

    } catch (scriptError) {
      console.error('⚠️ Could not access script:', scriptError.message);
      return {
        success: false,
        message: 'Could not access script: ' + scriptError.message,
        details: { scriptId, email, error: scriptError.message }
      };
    }

  } catch (error) {
    console.error('❌ CHECK ERROR:', error.message);
    return { success: false, message: error.message };
  }
}

// ============================================================
// 🔧 FIX: Re-deploy Last Client Script
// ============================================================

function redeployLastClientScript() {
  try {
    console.log('🔧 ========================================');
    console.log('🔧 RE-DEPLOYING LAST CLIENT');
    console.log('🔧 ========================================');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const clientsSheet = ss.getSheetByName('Clients');

    if (!clientsSheet) {
      console.error('❌ Clients sheet not found');
      return { success: false, message: 'Clients sheet not found' };
    }

    const data = clientsSheet.getDataRange().getValues();
    const headers = data[0];

    // Get last client row
    const lastRow = data[data.length - 1];
    const emailCol = headers.indexOf('Email');
    const nameCol = headers.indexOf('Client_Name');
    const colorCol = headers.indexOf('Primary_Color');
    const logoCol = headers.indexOf('Logo_Url');
    const idCol = headers.indexOf('Client_ID');

    const email = lastRow[emailCol];
    const clientName = lastRow[nameCol];
    const primaryColor = lastRow[colorCol] || '#2E7D32';
    const logoUrl = lastRow[logoCol] || '';
    const clientId = lastRow[idCol];

    console.log('📊 Re-deploying client:');
    console.log('- Email:', email);
    console.log('- Name:', clientName);
    console.log('- ID:', clientId);

    // Generate fresh code
    console.log('\n🔄 Generating fresh client code...');
    const masterUrl = getMasterApiUrl();
    const clientCode = generateClientCodeSafely({
      clientId: clientId,
      clientName: clientName,
      primaryColor: primaryColor,
      logoUrl: logoUrl,
      customDomain: '',
      email: email,
      sheetId: '', // Will be fetched from sheet
      businessId: clientId,
      masterApiUrl: masterUrl
    });

    console.log('✅ Code generated:', clientCode.length, 'characters');
    console.log('✅ Contains doGet:', clientCode.includes('function doGet'));
    console.log('✅ Contains doPost:', clientCode.includes('function doPost'));

    if (!clientCode.includes('function doGet')) {
      console.error('❌ CRITICAL: Generated code missing doGet!');
      return { success: false, message: 'Generated code is invalid' };
    }

    // Get current Script ID (if exists)
    const scriptIdCol = headers.indexOf('Script_ID');
    const currentScriptId = lastRow[scriptIdCol];

    if (!currentScriptId || currentScriptId === '') {
      console.log('⚠️ No current script ID - cannot re-deploy');
      console.log('ℹ️ Run setupOtpGlobalClient() to create a fresh client instead');
      return { 
        success: false, 
        message: 'No script ID to re-deploy. Run setupOtpGlobalClient() instead',
        client: { email, clientName, clientId }
      };
    }

    console.log('\n📤 Current Script ID:', currentScriptId);
    console.log('ℹ️ Manual re-deployment would require updating the Apps Script project');
    console.log('ℹ️ Since the project was already created, we would need to update its files');

    return {
      success: false,
      message: 'Current script exists. To fix: Delete this client and run setupOtpGlobalClient() again',
      recommendation: {
        step1: 'Delete the client row from Clients sheet',
        step2: 'Delete the script project from Apps Script home',
        step3: 'Run: setupOtpGlobalClient()',
        note: 'This ensures clean deployment with all functions'
      }
    };

  } catch (error) {
    console.error('❌ RE-DEPLOY ERROR:', error.message);
    return { success: false, message: error.message };
  }
}

// ============================================================
// 🎯 RUN ALL DIAGNOSTICS
// ============================================================

function runAllDiagnostics() {
  console.log('\n\n╔════════════════════════════════════════════════════════════╗');
  console.log('║           BIZOS CLIENT DEPLOYMENT DIAGNOSTICS             ║');
  console.log('╚════════════════════════════════════════════════════════════╝\n');

  console.log('\n📋 TEST 1: CODE GENERATION');
  console.log('═══════════════════════════════════════════════════════════');
  const genTest = testCodeGeneration();

  console.log('\n\n📋 TEST 2: CHECK LAST DEPLOYED CLIENT');
  console.log('═══════════════════════════════════════════════════════════');
  const deployTest = checkLastDeployedClientScript();

  console.log('\n\n📋 SUMMARY');
  console.log('═══════════════════════════════════════════════════════════');
  console.log('Code Generation:', genTest.success ? '✅ PASS' : '❌ FAIL');
  console.log('Last Deployed:', deployTest.success ? '✅ PASS' : '❌ FAIL');

  if (!genTest.success) {
    console.log('\n⚠️ ACTION REQUIRED: Fix code generation');
    console.log('- generateClientCodeSafely() is returning invalid code');
    console.log('- Check the function definition for syntax errors');
  }

  if (!deployTest.success && deployTest.message.includes('missing doGet')) {
    console.log('\n⚠️ ACTION REQUIRED: Re-deploy client');
    console.log('Solution: Delete and recreate the client');
    console.log('1. Delete row from Clients sheet');
    console.log('2. Delete script from Apps Script home');
    console.log('3. Run: setupOtpGlobalClient()');
  }

  console.log('\n');
  return {
    codeGeneration: genTest,
    lastDeployed: deployTest
  };
}