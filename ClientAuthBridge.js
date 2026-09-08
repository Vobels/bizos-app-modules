// Compatibility bridge for the paid-client login page.
// The client deployment already knows its own clientId, workspace and business name.
// Users authenticate with the same BizOS email + password they used before payment.
function authenticateClient(email, password) {
  var config = getClientDeploymentConfig();
  if (!config || !config.success) {
    return { success:false, message:'Client deployment configuration is unavailable.', code:'NO_CLIENT_CONFIG' };
  }

  if (!email || !password) {
    return { success:false, message:'Email and password are required.', code:'INVALID_INPUT' };
  }

  return clientLogin(
    config.clientId,
    email,
    password,
    config.sheetId,
    config.businessId,
    config.clientName
  );
}
