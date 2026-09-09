// ============================================================
// ProvisioningSafeTest.js
// Explicit integration test for the authoritative provisioning path.
// This does NOT override any legacy function.
// ============================================================

function setupOtpGlobalClientSafe(force) {
  console.log('🚀 ========================================');
  console.log('🚀 SAFE OTP GLOBAL CLIENT PROVISIONING');
  console.log('🚀 ========================================');

  var email = 'otpglobal.ltd@gmail.com';
  var businessName = 'Otp global company';
  var domain = 'bizos.otpgloballimited.com';

  if (force) {
    console.log('🔄 FORCE TEST: cleaning existing OTP test artifacts...');
    cleanupClient();
    deleteClientScriptFromDrive();
  }

  var paymentData = {
    email: email,
    password: 'otpglobal123',
    businessName: businessName,
    domain: domain,
    customDomain: domain,
    primaryColor: '#1F2937',
    logoUrl: '',
    tier: 'sovereign'
  };

  console.log('📌 CALLING EXPLICIT autoSetupClientSafe()');
  var result = autoSetupClientSafe(paymentData);

  console.log('📊 SAFE RESULT:', JSON.stringify({
    success: result && result.success,
    idempotent: result && result.idempotent,
    code: result && result.code,
    clientId: result && result.clientId,
    scriptId: result && result.scriptId,
    deploymentId: result && result.deploymentId,
    webAppUrl: result && result.webAppUrl,
    sheetId: result && result.sheetId,
    workspaceId: result && result.workspaceId,
    cleanedUp: result && result.cleanedUp,
    message: result && result.message
  }, null, 2));

  return result;
}
