// ============================================================
// 📁 Code.gs - FIXED doPost() - Handle clientLogin FIRST
// ============================================================
function doGet(e) {
  try {
    var params = e?.parameter || {};

    if (params.auth === '1') {
      return HtmlService.createHtmlOutput(
        '<html><body style="font-family: Arial; text-align: center; padding: 50px;">' +
        '<h2 style="color: #5D2A86;">✅ Authorization Successful</h2>' +
        '<p>You can now close this tab and try logging in again.</p>' +
        '</body></html>'
      ).setTitle('Authorization Complete');
    }

    // ✅ If page=dashboard, redirect to MAIN APP dashboard
    if (params.page === 'dashboard' && params.sessionId) {
      var mainAppUrl = CLIENT_CONFIG.masterApiUrl;
      return HtmlService.createHtmlOutput(
        '<html><head><meta http-equiv="refresh" content="0; url=' + mainAppUrl + '?page=dashboard&sessionId=' + params.sessionId + '&clientId=' + CLIENT_CONFIG.clientId + '"></head>' +
        '<body style="font-family: Arial; text-align: center; padding: 50px;"><p>Redirecting to dashboard...</p></body></html>'
      ).setTitle('Redirecting...');
    }

    // Serve the landing page
    var template = HtmlService.createTemplateFromFile('client-landing');
    template.settings = getClientSettings();
    template.clientId = CLIENT_CONFIG.clientId;
    template.clientName = CLIENT_CONFIG.clientName;

    return template.evaluate()
      .setTitle(CLIENT_CONFIG.clientName + ' - Business OS')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    return HtmlService.createHtmlOutput('<h1>Error</h1><p>' + error.message + '</p>');
  }
}


function doPost(e) {
  // ===== RATE LIMITING =====
  const rateLimitKey = `rate_limit_${e?.postData?.contents?.length || 'unknown'}`;
  const cache = CacheService.getScriptCache();
  const lastRequest = cache.get(rateLimitKey);
  if (lastRequest && (Date.now() - parseInt(lastRequest)) < 500) {
    console.log("⚠️ Rate limit - request too fast");
    return ContentService.createTextOutput(JSON.stringify({
      success: false, 
      message: "Please wait before next action"
    })).setMimeType(ContentService.MimeType.JSON);
  }
  cache.put(rateLimitKey, Date.now().toString(), 5);
  
  console.log("═══════════════════════════════════════");
  console.log("📨 doPost START");
  console.log("═══════════════════════════════════════");
  
  if (e === undefined) {
    console.error("❌ ERROR: No data received");
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "No data received"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const postData = JSON.parse(e.postData.contents);
    console.log("📨 Action:", postData.action);
    
    const action = postData.action;
    const payload = postData.payload || postData;
    const sessionId = postData.sessionId;

    let result;

    // ============================================================
    // ✅ PART 1: PUBLIC ACTIONS (No session required)
    // ============================================================
    
    // ⭐ CRITICAL: Handle clientLogin FIRST, before any session checks
    if (action === 'clientLogin') {
      console.log("📨 ========================================");
      console.log("📨 CLIENT LOGIN - PUBLIC ACTION");
      console.log("📨 ========================================");
      console.log('📊 Payload:', JSON.stringify(payload));
      
      result = clientLogin(
        payload.clientId || '',
        payload.email || '',
        payload.password || '',
        payload.sheetId || '',
        payload.businessId || payload.clientId || '',
        payload.businessName || ''
      );
      
      console.log('📊 Result:', result.success ? '✅ SUCCESS' : '❌ FAILED', result.message);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // ⭐ NEW: Check client session action
    if (action === 'checkClientSession') {
      console.log("📨 CHECK CLIENT SESSION - PUBLIC ACTION");
      const session = validateClientSession(payload.sessionId || '');
      result = session
        ? { 
            success: true, 
            session: { 
              email: session.email, 
              clientId: session.clientId, 
              businessName: session.businessName,
              sheetId: session.sheetId
            } 
          }
        : { success: false, message: 'Session expired' };
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Other public actions
    if (action === 'REGISTER_BUSINESS') {
      console.log("📨 REGISTER_BUSINESS - PUBLIC ACTION");
      const businessData = payload;
      const businessId = generateBusinessId();
      const bizSheet = getOrCreateBusinessSheet();
      bizSheet.appendRow([
        businessId, businessData.businessName, businessData.cac || "", "",
        "pending", new Date().toISOString(), "",
        businessData.ownerEmail, businessData.ownerName, businessData.ownerPhone || "",
        "", "", businessData.country || "Nigeria", "",
        "free", "active", "", new Date().toISOString(), "", "YES"
      ]);
      
      const userSheet = getOrCreateUserSheet();
      const hashedPassword = hashPassword(businessData.password || "temporary123");
      userSheet.appendRow([
        businessData.ownerEmail.toLowerCase(), hashedPassword, businessData.ownerName,
        businessData.ownerPhone || "", "owner", businessId, businessData.businessName,
        "NO", "", "pending", "", new Date().toISOString(), "", "YES", "NO"
      ]);
      
      const token = Utilities.getUuid();
      let verificationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email_Verifications');
      if (!verificationSheet) {
        verificationSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Email_Verifications');
        verificationSheet.appendRow(['Email', 'Token', 'Business_Name', 'Business_Id', 'Created_At', 'Expires_At', 'Is_Used']);
      }
      const expiresAt = new Date();
      expiresAt.setHours(expiresAt.getHours() + 24);
      verificationSheet.appendRow([businessData.ownerEmail.toLowerCase(), token, businessData.businessName, businessId, new Date().toISOString(), expiresAt.toISOString(), 'NO']);
      sendVerificationEmail(businessData.ownerEmail, businessData.businessName, businessId, token);
      
      result = { success: true, message: "Account created! Please check your email to verify your account.", businessId: businessId, requiresVerification: true };
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'VERIFY_EMAIL') {
      console.log("📨 VERIFY_EMAIL - PUBLIC ACTION");
      result = verifyEmailToken(payload.token);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'LOGIN') {
      console.log("📨 LOGIN - PUBLIC ACTION");
      result = authenticateUser(payload.email, payload.password, payload.businessName);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'resetPassword') {
      console.log("📨 RESET PASSWORD - PUBLIC ACTION");
      result = resetPassword(payload.email || '');
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'confirmPasswordReset') {
      console.log("📨 CONFIRM PASSWORD RESET - PUBLIC ACTION");
      result = confirmPasswordReset(payload.token || '', payload.newPassword || '');
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'resendVerificationEmail') {
      console.log("📨 RESEND VERIFICATION - PUBLIC ACTION");
      result = resendVerificationEmail(payload.email || '');
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }

    // ============================================================
    // ✅ PART 2: PROTECTED ACTIONS (Session required)
    // ============================================================
    
    console.log("🔐 Checking session for protected action:", action);
    const user = getUserFromSession(sessionId);
    
    if (!user) {
      console.error("❌ UNAUTHORIZED - No valid session for action:", action);
      return ContentService.createTextOutput(JSON.stringify({
        success: false, 
        message: "Unauthorized - please login",
        code: 'UNAUTHORIZED'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    console.log("✅ User authenticated:", user.email);

    // Route to appropriate function based on action
    switch (action) {
      // Dashboard functions
      case 'getDashboardKPIs':
        const period = payload[0] || 'this_month';
        result = getDashboardKPIs(period, sessionId);
        break;
      case 'getRecentTransactions':
        const limit = payload[0] || 10;
        result = getRecentTransactions(limit, sessionId);
        break;
      case 'getModuleSummary':
        result = getModuleSummary(sessionId);
        break;
      
      // Business Center Part 1
      case 'getBusinessCenterData':
        result = getBusinessCenterData(sessionId);
        break;
      case 'getBusinessCenterProductImages':
        result = getBusinessCenterProductImages(payload[0] || [], sessionId);
        break;
      case 'saveBusinessCenterProduct':
        result = saveBusinessCenterProduct(payload[0] || {}, sessionId);
        break;
      case 'saveBusinessCenterProductsBulk':
        result = saveBusinessCenterProductsBulk(payload[0] || [], sessionId);
        break;
      case 'saveBusinessCenterCustomer':
        result = saveBusinessCenterCustomer(payload[0] || {}, sessionId);
        break;
      // Business Center Part 2: POS + Inventory
      case 'lookupBusinessCenterProduct':
        result = lookupBusinessCenterProduct(payload[0] || '', sessionId);
        break;
      case 'completeBusinessCenterSale':
        result = completeBusinessCenterSale(payload[0] || {}, sessionId);
        break;
      case 'updateBusinessCenterProduct':
        result = updateBusinessCenterProduct(payload[0] || {}, sessionId);
        break;
      case 'adjustBusinessCenterStock':
        result = adjustBusinessCenterStock(payload[0] || {}, sessionId);
        break;
      case 'getBusinessCenterStockMovements':
        result = getBusinessCenterStockMovements(sessionId, payload[0] || '');
        break;
      // Business Center Part 4: Customer Ledger
      case 'getBusinessCenterCustomerLedger':
        result = getBusinessCenterCustomerLedger(payload[0] || '', sessionId);
        break;
      case 'recordBusinessCenterCustomerPayment':
        result = recordBusinessCenterCustomerPayment(payload[0] || {}, sessionId);
        break;
      // Business Center Part 3: Sales History + Receipts
      case 'getBusinessCenterSales':
        result = getBusinessCenterSales(payload[0] || {}, sessionId);
        break;
      case 'getBusinessCenterSaleDetails':
        result = getBusinessCenterSaleDetails(payload[0] || '', sessionId);
        break;
      case 'getBusinessCenterFinanceData':
        result=getBusinessCenterFinanceData(sessionId); break;
      case 'recordBusinessCenterExpense':
        result=recordBusinessCenterExpense(payload[0]||{},sessionId); break;
      case 'getEnhancedDashboardData':
        const dashPeriod = payload[0] || 'this_month';
        result = getEnhancedDashboardData(dashPeriod, sessionId);
        break;
      
      // Business Center Part 1: safe existing-data connection (read-only)
      case 'getBusinessCenterConnectedData':
        result = getBusinessCenterConnectedData(sessionId);
        break;
      
      // Module functions
      case 'getModuleData':
        const moduleName = payload[0];
        result = getModuleData(moduleName, sessionId);
        break;
      case 'getModuleDataWithSync':
        const syncModuleName = payload[0];
        result = getModuleDataWithSync(syncModuleName, sessionId);
        break;
      case 'saveRecord':
        const saveModuleName = payload[0];
        const saveData = payload[1];
        result = saveRecord(saveModuleName, saveData, sessionId);
        break;
      case 'deleteRecord':
        const deleteModuleName = payload[0];
        const deleteIndex = payload[1];
        result = deleteRecord(deleteModuleName, deleteIndex, sessionId);
        break;
      case 'updateRecord':
        const updateModuleName = payload[0];
        const updateData = payload[1];
        const recordId = payload[2];
        result = updateRecord(updateModuleName, updateData, recordId, sessionId);
        break;
      case 'updateRecordField':
        const fieldModuleName = payload[0];
        const rowIndex = payload[1];
        const columnName = payload[2];
        const newValue = payload[3];
        result = updateRecordField(fieldModuleName, rowIndex, columnName, newValue, sessionId);
        break;
      
      // Reference data functions
      case 'getModuleReferenceData':
        const refModuleName = payload[0];
        result = getModuleReferenceData(refModuleName);
        break;
      case 'getCategoriesByType':
        const type = payload[0];
        result = getCategoriesByType(type);
        break;
      case 'getSubcategories':
        const category = payload[0];
        result = getSubcategories(category);
        break;
      case 'getAgroActivities':
        const agroType = payload[0];
        result = getAgroActivities(agroType);
        break;
      
      // Export/Import functions
      case 'exportModuleData':
        const exportModuleName = payload[0];
        result = exportModuleData(exportModuleName, sessionId);
        break;
      case 'importCSV':
        const importModuleName = payload[0];
        const rows = payload[1];
        result = importCSV(importModuleName, rows, sessionId);
        break;
      
      // Staff functions
      case 'inviteStaffMember':
        const businessId = payload[0];
        const email = payload[1];
        const name = payload[2];
        const role = payload[3];
        const invitedBy = payload[4];
        result = inviteStaffMember(businessId, email, name, role, invitedBy);
        break;
      case 'getBusinessStaff':
        const staffBusinessId = payload[0];
        result = getBusinessStaff(staffBusinessId);
        break;
      case 'removeTeamMember':
        const removeBusinessId = payload[0];
        const removeEmail = payload[1];
        result = removeTeamMember(removeBusinessId, removeEmail);
        break;
      
      // Profile functions
      case 'getBusinessDetails':
        const detailsBusinessId = payload[0];
        result = getBusinessDetails(detailsBusinessId);
        break;
      case 'updateUserProfile':
        const profileBusinessId = payload[0];
        const profileName = payload[1];
        const profilePhone = payload[2];
        result = updateUserProfile(profileBusinessId, profileName, profilePhone);
        break;
      case 'changeUserPassword':
        const passwordEmail = payload[0];
        const currentPassword = payload[1];
        const newPassword = payload[2];
        result = changeUserPassword(passwordEmail, currentPassword, newPassword);
        break;
      case 'uploadProfilePhoto':
        const photoEmail = payload[0];
        const imageData = payload[1];
        result = uploadProfilePhoto(photoEmail, imageData);
        break;
      case 'requestDataExport':
        const exportEmail = payload[0];
        const exportBusinessId = payload[1];
        result = requestDataExport(exportEmail, exportBusinessId);
        break;
      case 'deleteBusinessAccount':
        const deleteBizId = payload[0];
        const deleteBizEmail = payload[1];
        result = deleteBusinessAccount(deleteBizId, deleteBizEmail);
        break;
      
      // Upgrade functions
      case 'getPendingUpgradeRequestStatus':
        const upgradeBusinessId = payload[0] || '';
        result = getPendingUpgradeRequestStatus(upgradeBusinessId, sessionId);
        break;
      case 'submitUpgradeRequest':
        const submitBusinessId = payload[0];
        const upgradeData = payload[1];
        result = submitUpgradeRequest(submitBusinessId, upgradeData, sessionId);
        break;
      case 'getCountryRequirements':
        const country = payload[0];
        result = getCountryRequirements(country);
        break;
      
      // Auth functions
      case 'logout':
        result = logout(sessionId);
        break;
      
      // Generic fallback
      case 'SAVE_TRANSACTION':
        result = saveRecord('Finance', payload, sessionId);
        break;

      // Admin functions
      case 'getAdminStats':
        result = getAdminStats();
        break;
      case 'getAllBusinesses':
        result = getAllBusinesses();
        break;
      case 'getAllUsers':
        result = getAllUsers();
        break;
      case 'getAllFeatures':
        result = getAllFeatures();
        break;
      case 'getAllModules':
        result = getAllModules();
        break;
      case 'addFeature':
        result = addFeature(payload[0]);
        break;
      case 'toggleFeature':
        result = toggleFeature(payload[0]);
        break;
      case 'deleteFeature':
        result = deleteFeature(payload[0]);
        break;
      case 'deleteBusiness':
        result = deleteBusiness(payload[0]);
        break;
      case 'toggleBusinessStatus':
        result = toggleBusinessStatus(payload[0]);
        break;
      case 'exportBusinessData':
        result = exportBusinessData();
        break;
      case 'clearSystemCache':
        result = clearSystemCache();
        break;
      case 'runSystemHealthCheck':
        result = runSystemHealthCheck();
        break;
      case 'getSystemLogs':
        result = getSystemLogs();
        break;
      case 'adminLogout':
        result = adminLogout();
        break;
      case 'getClientConfig':
        result = getClientConfig();
        break;
      case 'saveBusinessPreferences':
        const prefBusinessId = payload[0];
        const preferences = payload[1];
        result = saveBusinessPreferences(prefBusinessId, preferences);
        break;
      case 'getBusinessPreferences':
        const getPrefBusinessId = payload[0];
        result = getBusinessPreferences(getPrefBusinessId);
        break;
      case 'getWhiteLabelConfig':
        result = getWhiteLabelConfig();
        break;
      case 'setWhiteLabelConfig':
        const wlClientName = payload[0];
        const wlClientDomain = payload[1];
        const wlCustomCss = payload[2] || '';
        const wlCustomLogo = payload[3] || '';
        const wlCustomFavicon = payload[4] || '';
        result = setWhiteLabelConfig(wlClientName, wlClientDomain, wlCustomCss, wlCustomLogo, wlCustomFavicon);
        break;
      case 'UPDATE_PROFILE':
        result = updateUserProfile(user.businessId, payload.name, payload.phone);
        break;

      // Client Dashboard functions (Protected - require session)
      case 'getClientDashboard':
        result = getClientDashboardData(sessionId);
        break;

      case 'getClientModuleData':
        result = getClientModuleData(payload.moduleName || '', sessionId);
        break;

      case 'saveClientRecord':
        result = saveClientModuleRecord(payload.moduleName || '', payload.record || {}, sessionId);
        break;

      case 'deleteClientRecord':
        result = deleteClientRecord(payload.moduleName || '', payload.rowIndex || 0, sessionId);
        break;

      case 'clientLogout':
        result = clientLogout(sessionId);
        break;
      
      default:
        console.error("❌ Unknown action:", action);
        result = { success: false, message: "Action not supported: " + action, code: 'UNKNOWN_ACTION' };
    }
    
    console.log("📊 Action result - Success:", result.success, "Message:", result.message);
    console.log("═══════════════════════════════════════");
    console.log("📨 doPost END");
    console.log("═══════════════════════════════════════");
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("═══════════════════════════════════════");
    console.error("❌ ERROR in doPost");
    console.error("═══════════════════════════════════════");
    console.error("Error:", error.toString());
    console.error("Stack:", error.stack);
    console.error("═══════════════════════════════════════");
    
    return ContentService.createTextOutput(JSON.stringify({
      success: false, 
      message: error.toString(),
      code: 'INTERNAL_ERROR'
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// 📌 Helper: Get Script URL
// ============================================================

function getScriptUrl() {
  const url = ScriptApp.getService().getUrl();
  console.log("🔗 Script URL:", url);
  return url;
}

// ============================================================
// 📌 Helper: Include HTML files
// ============================================================

function include(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

function getUrl() {
  return ScriptApp.getService().getUrl();
}