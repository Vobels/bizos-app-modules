// ============================================================
// ZZZ_MasterEntryPoint.js - MASTER WEB ENTRY POINT
// ============================================================
function getPublicBizOSUrl_() {
  var base = '';
  try { base = CONFIG && CONFIG.URLS ? CONFIG.URLS.base : ''; } catch (e) {}
  if (!base) { try { base = ScriptApp.getService().getUrl() || ''; } catch (e) {} }
  return String(base).replace(/\/$/, '');
}
function getAppUrl() { return getPublicBizOSUrl_(); }
function getAuthoritativeBizOSWebAppUrl_() {
  try {
    var url = ScriptApp.getService().getUrl();
    if (url) return String(url).replace(/\/$/, '');
  } catch (e) {}
  return '';
}
function getInitialPublicRoute_(params) {
  var requested = params ? String(params.path || params.route || params.page || '').trim().toLowerCase() : '';
  var aliases = {landing:'/',login:'/login',dashboard:'/dashboard',profile:'/profile',staff:'/staff',settings:'/settings'};
  if (aliases[requested]) return aliases[requested];
  if (requested.charAt(0) !== '/') requested = '/' + requested;
  requested = requested.replace(/\/+$/, '') || '/';
  var allowed = {'/':true,'/login':true,'/dashboard':true,'/profile':true,'/staff':true,'/settings':true};
  return allowed[requested] ? requested : '/';
}
function doGet(e) {
  try {
    // Public GET requests must not initialize or repair the private master database.
    // Database access belongs in server-side operations that explicitly need it.
    var params = (e && e.parameter) ? e.parameter : {};
    var pathInfo = (e && e.pathInfo) ? String(e.pathInfo) : '';
    var pathRoute = pathInfo.replace(/^\/+|\/+$/g, '').toLowerCase();
    var requestedPage = String(params.page || pathRoute || '').toLowerCase();
    if (String(params.admin || '').toLowerCase() === 'true') requestedPage = 'admin';
    if (requestedPage === 'admin/settings') requestedPage = 'admin-settings';
    if (params.auth === '1') return HtmlService.createHtmlOutput('<html><body style="font-family:Arial;text-align:center;padding:50px"><h2 style="color:#5D2A86">Authorization Successful</h2><p>You can now close this tab and return to BizOS.</p></body></html>').setTitle('Authorization Complete');
    var invitationCode = String(params.invite || '').trim();
    if (invitationCode || requestedPage === 'invite' || requestedPage === 'invitation') {
      var invitationTemplate = HtmlService.createTemplateFromFile('staff-invitation');
      invitationTemplate.invitationCode = invitationCode;
      return invitationTemplate.evaluate().setTitle('Join BizOS').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    if (requestedPage === 'payment-method' || requestedPage === 'paymentmethod') {
      var methodRequestId = String(params.requestId || '').trim();
      var methodAccessToken = String(params.accessToken || '').trim();
      var methodTemplate = HtmlService.createTemplateFromFile('PaymentMethod');
      methodTemplate.requestId = methodRequestId;
      methodTemplate.accessToken = methodAccessToken;
      methodTemplate.publicBizOSUrl = getAuthoritativeBizOSWebAppUrl_() || getPublicBizOSUrl_();
      return methodTemplate.evaluate().setTitle('Choose Payment Method - BizOS').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    if (requestedPage === 'flutterwave-callback') {
      var callbackRequestId = String(params.requestId || '').trim();
      var transactionId = String(params.transaction_id || params.transactionId || '').trim();
      var callbackStatus = String(params.status || '').trim();
      var callbackResult = handleFlutterwaveCallback(callbackRequestId, transactionId, callbackStatus);
      var publicUrl = getPublicBizOSUrl_();
      if (callbackResult && callbackResult.success) {
        var launchUrl = callbackResult.landingUrl || callbackResult.webAppUrl || publicUrl;
        return HtmlService.createHtmlOutput('<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>Payment Confirmed - BizOS</title><meta http-equiv="refresh" content="2;url=' + String(launchUrl).replace(/"/g,'&quot;') + '"></head><body style="font-family:Arial;text-align:center;padding:60px"><h1 style="color:#5D2A86">Payment confirmed</h1><p>Your payment has been confirmed and your BizOS workspace is being prepared.</p><p>Opening your workspace...</p></body></html>').setTitle('Payment Confirmed - BizOS');
      }
      var message = callbackResult && callbackResult.message ? callbackResult.message : 'We could not confirm this payment.';
      return HtmlService.createHtmlOutput('<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>Payment Update - BizOS</title></head><body style="font-family:Arial;text-align:center;padding:60px"><h1 style="color:#5D2A86">Payment update</h1><p>' + String(message).replace(/</g,'&lt;').replace(/>/g,'&gt;') + '</p><p><a href="' + publicUrl.replace(/"/g,'&quot;') + '">Return to BizOS</a></p></body></html>').setTitle('Payment Update - BizOS');
    }
    if (requestedPage === 'deployment-status' || requestedPage === 'deployment' || requestedPage === 'deployment-confirmation') {
      var deploymentTemplate = HtmlService.createTemplateFromFile('DeploymentStatus');
      deploymentTemplate.requestId = String(params.requestId || '').trim();
      deploymentTemplate.confirmationMode = requestedPage === 'deployment-confirmation';
      deploymentTemplate.publicBizOSUrl = getAuthoritativeBizOSWebAppUrl_() || getPublicBizOSUrl_();
      return deploymentTemplate.evaluate().setTitle(deploymentTemplate.confirmationMode ? 'BizOS Deployment Confirmation' : 'BizOS Deployment Status').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (requestedPage === 'paystack-int' || requestedPage === 'payment') {
      var requestId = String(params.requestId || '').trim();
      var paymentAccessToken = String(params.accessToken || '').trim();
      var paymentTemplate = HtmlService.createTemplateFromFile('paystack-payment');
      paymentTemplate.requestId = requestId;
      paymentTemplate.accessToken = paymentAccessToken;
      paymentTemplate.publicBizOSUrl = getAuthoritativeBizOSWebAppUrl_() || getPublicBizOSUrl_();
      return paymentTemplate.evaluate().setTitle('Complete Payment - BizOS').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    if (requestedPage === 'admin-reset') {
      var adminResetTemplate = HtmlService.createTemplateFromFile('admin-reset');
      adminResetTemplate.publicBizOSUrl = getAuthoritativeBizOSWebAppUrl_() || getPublicBizOSUrl_();
      return adminResetTemplate.evaluate().setTitle('Reset Admin Password - BizOS').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (requestedPage === 'admin' || requestedPage === 'admin-settings') {
      var adminSessionId = String(params.session || '').trim();
      if (typeof validateAdminSession !== 'function' || !validateAdminSession(adminSessionId)) {
        var adminLoginTemplate = HtmlService.createTemplateFromFile('admin-login');
        adminLoginTemplate.publicBizOSUrl = getAuthoritativeBizOSWebAppUrl_() || getPublicBizOSUrl_();
    adminLoginTemplate.adminUrl = getAuthoritativeBizOSWebAppUrl_() || getPublicBizOSUrl_();
        return adminLoginTemplate.evaluate().setTitle('Admin Login - BizOS').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      var adminTemplate = HtmlService.createTemplateFromFile('admin');
      // admin.html uses publicBizOSUrl for its post-login/navigation links.
      // Supply it explicitly; otherwise Apps Script template evaluation throws
      // a ReferenceError and the generic doGet() fallback displays "Unable to load BizOS".
      adminTemplate.publicBizOSUrl = getAuthoritativeBizOSWebAppUrl_() || getPublicBizOSUrl_();
      adminTemplate.initialAdminTab = requestedPage === 'admin-settings' ? 'settings' : 'overview';
      var adminOutput = adminTemplate.evaluate().setTitle(requestedPage === 'admin-settings' ? 'BizOS Admin Settings' : 'BizOS Admin Dashboard').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      if (requestedPage === 'admin-settings') adminOutput.append('<script>window.addEventListener("load",function(){if(typeof switchTab==="function")switchTab("settings");});</script>');
      return adminOutput;
    }
    var verificationToken = params.verify || params.token;
    if (verificationToken) {
      var verificationResult = verifyEmailToken(String(verificationToken));
      var ok = verificationResult && verificationResult.success;
      var publicUrl = getPublicBizOSUrl_();
      var message = verificationResult && verificationResult.message ? verificationResult.message : (ok ? 'Your email has been verified.' : 'We could not verify this email address.');
      var title = ok ? 'Email Verified' : 'Verification Failed';
      return HtmlService.createHtmlOutput('<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="refresh" content="0;url=' + publicUrl + '"><title>' + title + ' - BizOS</title><script>window.top.location.replace(' + JSON.stringify(publicUrl) + ');</script></head><body style="font-family:Arial;text-align:center;padding:50px"><h1 style="color:#5D2A86">' + title + '</h1><p>' + message + '</p></body></html>').setTitle(title + ' - BizOS');
    }
    var openBilling = requestedPage === 'billing' || String(params.billing || '').toLowerCase() === 'true' || String(params.billing || '') === '1';
    var editUpgrade = String(params.editUpgrade || '').toLowerCase() === 'true' || String(params.editUpgrade || '') === '1';
    var editUpgradeRequestId = String(params.requestId || '').trim();
    if (requestedPage === 'billing') params = Object.assign({}, params, {page: 'dashboard'});
    var initialRoute = getInitialPublicRoute_(params);
    var template = HtmlService.createTemplateFromFile('index');
    template.showLogin = initialRoute === '/login';
    template.openBilling = openBilling;
    template.editUpgrade = editUpgrade;
    template.editUpgradeRequestId = editUpgradeRequestId;
    template.initialRoute = initialRoute;
    template.routeMode = 'public';
    template.routeTarget = initialRoute === '/' || initialRoute === '/login' ? 'landing' : 'app';
    return template.evaluate().setTitle('BizOS - Business Operating System').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('Master doGet error:', error);
    return HtmlService.createHtmlOutput('<h1>BizOS</h1><p>Unable to load BizOS.</p>');
  }
}