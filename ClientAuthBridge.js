// Compatibility bridge for the client login page.
// Keeps the existing clientLogin implementation while exposing the
// authenticateClient name used by client-landing.html.
function authenticateClient(email, password, businessName) {
  var config = getClientDeploymentConfig();
  if (!config || !config.success) {
    return { success:false, message:'Client deployment configuration is unavailable.', code:'NO_CLIENT_CONFIG' };
  }
  return clientLogin(
    config.clientId,
    email,
    password,
    config.sheetId,
    config.businessId,
    businessName || config.clientName || ''
  );
}
