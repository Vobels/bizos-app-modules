// ============================================================
// ZZZ_MasterEntryPoint.js - MASTER WEB ENTRY POINT
// ============================================================
// One web entry point for the master BizOS application.
// Internal client navigation is owned by index.html + the existing
// show...Page() functions. Netlify supplies the public URL shell.
// main is untouched.
// ============================================================

function getPublicBizOSUrl_() {
  var base = '';
  try { base = CONFIG && CONFIG.URLS ? CONFIG.URLS.base : ''; } catch (e) {}
  if (!base) base = 'https://bizos.higroups.com';
  return String(base).replace(/\/$/, '');
}
function getAppUrl() { return getPublicBizOSUrl_(); }
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
    try { ensureBizOSReady_(); } catch (dbError) { console.error('Database readiness check failed:', dbError); }
    var params = (e && e.parameter) ? e.parameter : {};
    var pathInfo = (e && e.pathInfo) ? String(e.pathInfo) : '';
    var pathRoute = pathInfo.replace(/^\/+|\/+$/g, '').toLowerCase();
    var requestedPage = String(params.page || pathRoute || '').toLowerCase();
    if (requestedPage === 'admin/settings') requestedPage = 'admin-settings';

    if (params.auth === '1') return HtmlService.createHtmlOutput('<html><body style="font-family:Arial;text-align:center;padding:50px"><h2 style="color:#5D2A86">Authorization Successful</h2><p>You can now close this tab and return to BizOS.</p></body></html>').setTitle('Authorization Complete');

    // Staff invitations must never fall through to the normal landing/login UI.
    // The existing email format is ?invite=CODE, so both old and new links
    // are supported without requiring previously sent invitations to be resent.
    var invitationCode = String(params.invite || '').trim();
    if (invitationCode || requestedPage === 'invite' || requestedPage === 'invitation') {
      var invitationTemplate = HtmlService.createTemplateFromFile('staff-invitation');
      invitationTemplate.invitationCode = invitationCode;
      return invitationTemplate.evaluate().setTitle('Join BizOS').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (requestedPage === 'paystack-int' || requestedPage === 'payment') {
      var requestId = String(params.requestId || '').trim();
      var paymentTemplate = HtmlService.createTemplateFromFile('paystack-payment');
      paymentTemplate.requestId = requestId;
      return paymentTemplate.evaluate().setTitle('Complete Payment - BizOS').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (requestedPage === 'admin' || requestedPage === 'admin-settings') {
      var adminTemplate = HtmlService.createTemplateFromFile('admin');
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
      return HtmlService.createHtmlOutput('<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="refresh" content="0;url=' + publicUrl + '"><title>' + title + ' - BizOS</title><script>window.top.location.replace(' + JSON.stringify(publicUrl) + ');</script></head><body style="margin:0;background:#f8fafc;font-family:Arial,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:24px;box-sizing:border-box;"><div style="width:100%;max-width:520px;background:#fff;border:1px solid #e5e7eb;border-radius:18px;padding:36px;text-align:center;box-shadow:0 10px 30px rgba(0,0,0,.08);box-sizing:border-box;"><h1 style="color:#5D2A86;margin:0 0 12px;">' + title + '</h1><p style="color:#4b5563;line-height:1.6;margin:0;">' + message + '</p><p style="color:#6b7280;margin-top:18px;">Returning to BizOS...</p></div></body></html>').setTitle(title + ' - BizOS');
    }

    // Normal client routes all render the same application document.
    // index.html applies initialRoute using the existing internal page
    // functions. There are no separate Dashboard/Profile/Staff/Settings
    // public HTML pages.
    var initialRoute = getInitialPublicRoute_(params);
    var template = HtmlService.createTemplateFromFile('index');
    template.showLogin = initialRoute === '/login';
    template.initialRoute = initialRoute;
    template.routeMode = 'public';
    template.routeTarget = initialRoute === '/' || initialRoute === '/login' ? 'landing' : 'app';
    return template.evaluate().setTitle('BizOS - Business Operating System').addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('Master doGet error:', error);
    return HtmlService.createHtmlOutput('<h1>BizOS</h1><p>Unable to load BizOS.</p>');
  }
}