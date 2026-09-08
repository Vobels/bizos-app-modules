// ============================================================
// AutoSetupClientSafe.js - SAFE PAID CLIENT PROVISIONING
// ============================================================
// Overrides the legacy autoSetupClient() flow without modifying
// the existing Master Backend source yet.
//
// IMPORTANT: A client is NOT recorded as active unless the private
// deployment succeeds. Failed provisioning cleans up the workspace
// created during the attempt.
// ============================================================

(function () {
  var legacyAutoSetupClient = autoSetupClient;

  autoSetupClient = function (paymentData) {
    var clientId = '';
    var sheetResult = null;
    var scriptResult = null;
    var existingBusiness = null;
    var businessId = null;
    var businessName = '';

    try {
      paymentData = paymentData || {};

      if (!paymentData.email) {
        return { success: false, code: 'INVALID_PAYMENT_DATA', message: 'Client email is required.' };
      }

      businessName = paymentData.businessName || paymentData.name || 'My Business';
      clientId = 'CLIENT_' + Utilities.getUuid().substring(0, 8).toUpperCase();

      console.log('🚀 SAFE CLIENT SETUP:', paymentData.email);
      console.log('🆔 Client ID:', clientId);

      // --------------------------------------------------------
      // 1. Duplicate protection
      // --------------------------------------------------------
      var duplicateCheck = checkDuplicateClient(paymentData.email, businessName);
      if (duplicateCheck && duplicateCheck.isDuplicate) {
        return {
          success: false,
          code: 'DUPLICATE_' + duplicateCheck.type,
          message: duplicateCheck.message,
          details: duplicateCheck
        };
      }

      // --------------------------------------------------------
      // 2. Resolve existing BizOS business identity
      // --------------------------------------------------------
      existingBusiness = getBusinessByEmail(paymentData.email);
      businessId = existingBusiness ? existingBusiness.businessId : clientId;

      // --------------------------------------------------------
      // 3. Create the isolated client workspace
      // --------------------------------------------------------
      sheetResult = createClientSheetInClientDrive(
        paymentData.email,
        clientId,
        businessName,
        {
          businessId: businessId,
          tier: paymentData.tier || 'sovereign',
          status: 'provisioning',
          primaryColor: paymentData.primaryColor || '#2E7D32',
          logoUrl: paymentData.logoUrl || '',
          provisioningVersion: '2.0'
        }
      );

      if (!sheetResult || !sheetResult.success || !sheetResult.sheetId) {
        return {
          success: false,
          code: 'WORKSPACE_CREATION_FAILED',
          message: 'Client workspace could not be created: ' + ((sheetResult && sheetResult.message) || 'Unknown error')
        };
      }

      // --------------------------------------------------------
      // 4. Deploy the private client application
      // --------------------------------------------------------
      var masterUrl = getMasterApiUrl();
      var clientCode = generateClientCodeSafely({
        clientId: clientId,
        clientName: businessName,
        primaryColor: paymentData.primaryColor || '#2E7D32',
        logoUrl: paymentData.logoUrl || '',
        customDomain: paymentData.customDomain || '',
        email: paymentData.email,
        sheetId: sheetResult.sheetId,
        businessId: businessId,
        masterApiUrl: masterUrl
      });

      if (!clientCode || !clientCode.files || !clientCode.files.length) {
        throw new Error('Client deployment package is empty.');
      }

      scriptResult = createAndDeployClientScript(
        paymentData.email,
        clientId,
        clientCode,
        businessName,
        paymentData.primaryColor || '#2E7D32',
        paymentData.logoUrl || ''
      );

      // NO FALLBACK LANDING PAGE HERE.
      // A paid client must receive a real isolated deployment.
      if (!scriptResult || !scriptResult.success || !scriptResult.webAppUrl || !scriptResult.scriptId) {
        var deploymentMessage = scriptResult && scriptResult.message
          ? scriptResult.message
          : 'The client application could not be deployed.';

        cleanupFailedClientProvisioning_(sheetResult, clientId);

        return {
          success: false,
          code: 'DEPLOYMENT_FAILED',
          message: 'Client provisioning failed: ' + deploymentMessage,
          clientId: clientId,
          cleanedUp: true
        };
      }

      console.log('✅ Deployment successful:', scriptResult.webAppUrl);

      // --------------------------------------------------------
      // 5. ONLY NOW create the master Clients record
      // --------------------------------------------------------
      var saveResult = saveClientRecord({
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
        webAppUrl: scriptResult.webAppUrl,
        landingUrl: scriptResult.webAppUrl,
        apiKey: Utilities.getUuid(),
        scriptId: scriptResult.scriptId
      });

      if (!saveResult || !saveResult.success) {
        cleanupFailedClientProvisioning_(sheetResult, clientId);
        cleanupFailedClientDeployment_(scriptResult);

        return {
          success: false,
          code: 'CLIENT_RECORD_SAVE_FAILED',
          message: 'Deployment succeeded, but the client record could not be saved. Provisioning was rolled back.',
          clientId: clientId,
          cleanedUp: true
        };
      }

      // --------------------------------------------------------
      // 6. Update existing business only after successful record
      // --------------------------------------------------------
      if (existingBusiness && businessId) {
        updateBusinessWithClientInfo(businessId, clientId, scriptResult.webAppUrl);
      }

      // --------------------------------------------------------
      // 7. Welcome email only after everything is successful
      // --------------------------------------------------------
      sendClientWelcomeEmail(
        paymentData.email,
        businessName,
        scriptResult.webAppUrl,
        clientId
      );

      console.log('✅ SAFE CLIENT PROVISIONING COMPLETE');

      return {
        success: true,
        clientId: clientId,
        landingUrl: scriptResult.webAppUrl,
        sheetUrl: sheetResult.sheetUrl,
        email: paymentData.email,
        scriptId: scriptResult.scriptId,
        deploymentId: scriptResult.deploymentId,
        message: 'Client setup complete.'
      };

    } catch (error) {
      console.error('❌ SAFE CLIENT SETUP ERROR:', error);

      // If a workspace was created but the workflow failed before the
      // success record was committed, remove the workspace.
      if (sheetResult && sheetResult.success) {
        cleanupFailedClientProvisioning_(sheetResult, clientId);
      }

      return {
        success: false,
        code: 'PROVISIONING_FAILED',
        message: error && error.message ? error.message : 'Client provisioning failed.',
        clientId: clientId || null,
        cleanedUp: !!sheetResult
      };
    }
  };

  // ------------------------------------------------------------
  // Roll back the workspace created during a failed provisioning
  // attempt. This does not touch unrelated client workspaces.
  // ------------------------------------------------------------
  function cleanupFailedClientProvisioning_(sheetResult, clientId) {
    try {
      if (!sheetResult || !sheetResult.sheetId) return;

      var file = DriveApp.getFileById(sheetResult.sheetId);
      file.setTrashed(true);

      console.log('🧹 Rolled back client workspace:', clientId, sheetResult.sheetId);
    } catch (error) {
      // Never hide the original provisioning failure because cleanup failed.
      console.error('⚠️ Workspace cleanup failed for ' + clientId + ':', error);
    }
  }

  // A deployed Apps Script project is not safely deleted here by default.
  // Keeping this helper isolated lets us add authenticated project deletion
  // later without weakening the provisioning transaction.
  function cleanupFailedClientDeployment_(scriptResult) {
    try {
      if (!scriptResult || !scriptResult.scriptId) return;
      console.warn('⚠️ Orphaned deployment requires cleanup:', scriptResult.scriptId);
    } catch (error) {
      console.error('⚠️ Deployment cleanup check failed:', error);
    }
  }
})();
