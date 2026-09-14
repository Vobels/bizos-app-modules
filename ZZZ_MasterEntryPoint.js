// ============================================================
// ZZZ_MasterEntryPoint.js - AUTHORITATIVE MASTER WEB ENTRY POINT
// ============================================================
// This is the final master doGet assignment in the source set.
// It serves index.html for the master BizOS application. Paid client
// deployments have their own generated doGet and are not routed here.
// ============================================================

function getPublicBizOSUrl_() {
  var base = '';
  try { base = CONFIG && CONFIG.URLS ? CONFIG.URLS.base : ''; } catch (e) {}
  if (!base) base = 'https://bizos.higroups.com';
  return String(base).replace(/\/$/, '');
}

function getAppUrl() { return getPublicBizOSUrl_(); }

function getInitialPublicRoute_(params) {
  var requested = '';
  if (params) requested = params.path || params.route || params.page || '';
  requested = String(requested || '').trim().toLowerCase();
  if (!requested) return '/';

  var aliases = {
    landing: '/', dashboard: '/dashboard', profile: '/profile', staff: '/staff',
    admin: '/admin', 'admin-settings': '/admin/settings', settings: '/admin/settings'
  };
  if (aliases[requested]) return aliases[requested];
  if (requested.charAt(0) !== '/') requested = '/' + requested;
  requested = requested.replace(/\/+$/, '') || '/';

  var allowed = {
    '/': true, '/dashboard': true, '/profile': true, '/staff': true,
    '/admin': true, '/admin/settings': true
  };
  return allowed[requested] ? requested : '/';
}

doGet = function(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var requestedPage = String(params.page || '').toLowerCase();

    if (params.auth === '1') {
      return HtmlService.createHtmlOutput(
        '<html><body style="font-family:Arial;text-align:center;padding:50px">' +
        '<h2 style="color:#5D2A86">Authorization Successful</h2>' +
        '<p>You can now close this tab and return to BizOS.</p>' +
        '</body></html>'
      ).setTitle('Authorization Complete');
    }

    // Preserve the reconstructed frontend's dedicated Admin entry while
    // keeping the existing full admin.html implementation from this branch.
    if (requestedPage === 'admin' || requestedPage === 'admin-settings') {
      return HtmlService.createTemplateFromFile('admin').evaluate()
        .setTitle(requestedPage === 'admin-settings' ? 'BizOS Admin Settings' : 'BizOS Admin Dashboard')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    var verificationToken = params.verify || params.token;
    if (verificationToken) {
      var verificationResult = verifyEmailToken(String(verificationToken));
      var ok = verificationResult && verificationResult.success;
      var publicUrl = getPublicBizOSUrl_();
      var message = verificationResult && verificationResult.message
        ? verificationResult.message
        : (ok ? 'Your email has been verified.' : 'We could not verify this email address.');
      var title = ok ? 'Email Verified' : 'Verification Failed';

      return HtmlService.createHtmlOutput(
        '<!doctype html><html><head>' +
        '<meta name="viewport" content="width=device-width,initial-scale=1">' +
        '<meta http-equiv="refresh" content="0;url=' + publicUrl + '">' +
        '<title>' + title + ' - BizOS</title>' +
        '<script>window.top.location.replace(' + JSON.stringify(publicUrl) + ');</script>' +
        '</head><body style="margin:0;background:#f8fafc;font-family:Arial,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:24px;box-sizing:border-box;">' +
        '<div style="width:100%;max-width:520px;background:#fff;border:1px solid #e5e7eb;border-radius:18px;padding:36px;text-align:center;box-shadow:0 10px 30px rgba(0,0,0,.08);box-sizing:border-box;">' +
        '<h1 style="color:#5D2A86;margin:0 0 12px;">' + title + '</h1>' +
        '<p style="color:#4b5563;line-height:1.6;margin:0;">' + message + '</p>' +
        '<p style="color:#6b7280;margin-top:18px;">Returning to BizOS...</p>' +
        '</div></body></html>'
      ).setTitle(title + ' - BizOS');
    }

    var template = HtmlService.createTemplateFromFile('index');
    template.showLogin = requestedPage === 'login';
    template.initialRoute = getInitialPublicRoute_(params);
    template.routeMode = template.initialRoute === '/dashboard' ? 'dashboard' : 'landing';
    template.routeTarget = template.initialRoute === '/' ? 'landing' : template.initialRoute.substring(1);

    return template.evaluate()
      .setTitle('BizOS - Business Operating System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('Master doGet error:', error);
    var publicUrl = getPublicBizOSUrl_();
    return HtmlService.createHtmlOutput(
      '<!doctype html><html><head>' +
      '<meta http-equiv="refresh" content="0;url=' + publicUrl + '">' +
      '<script>window.top.location.replace(' + JSON.stringify(publicUrl) + ');</script>' +
      '</head><body style="font-family:Arial;text-align:center;padding:50px">' +
      '<p>Returning to BizOS...</p></body></html>'
    );
  }
};