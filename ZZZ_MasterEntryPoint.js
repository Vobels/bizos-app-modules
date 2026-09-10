// ============================================================
// ZZZ_MasterEntryPoint.js - AUTHORITATIVE MASTER WEB ENTRY POINT
// ============================================================
// This is the final master doGet assignment in the source set.
// It serves index.html for the master BizOS application. Paid client
// deployments have their own generated doGet and are not routed here.
// ============================================================

doGet = function(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};

    if (params.auth === '1') {
      return HtmlService.createHtmlOutput(
        '<html><body style="font-family:Arial;text-align:center;padding:50px">' +
        '<h2 style="color:#5D2A86">Authorization Successful</h2>' +
        '<p>You can now close this tab and return to BizOS.</p>' +
        '</body></html>'
      ).setTitle('Authorization Complete');
    }

    // Email verification is handled here, but every user-facing destination
    // remains on the public BizOS domain. Apps Script is backend-only.
    if (params.verify) {
      var verificationResult = verifyEmailToken(String(params.verify));
      var ok = verificationResult && verificationResult.success;
      var message = verificationResult && verificationResult.message
        ? verificationResult.message
        : (ok ? 'Your email has been verified.' : 'We could not verify this email address.');
      var color = ok ? '#10B981' : '#EF4444';
      var title = ok ? 'Email Verified' : 'Verification Failed';
      var publicUrl = getPublicBizOSUrl_();
      var action = ok
        ? '<a href="' + publicUrl + '/login" style="display:inline-block;background:#5D2A86;color:#fff;padding:12px 22px;border-radius:8px;text-decoration:none;font-weight:600;">Sign in to BizOS</a>'
        : '<a href="' + publicUrl + '" style="display:inline-block;background:#5D2A86;color:#fff;padding:12px 22px;border-radius:8px;text-decoration:none;font-weight:600;">Return to BizOS</a>';

      return HtmlService.createHtmlOutput(
        '<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1">' +
        '<title>' + title + ' - BizOS</title></head>' +
        '<body style="margin:0;background:#f8fafc;font-family:Arial,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:24px;box-sizing:border-box;">' +
          '<div style="width:100%;max-width:520px;background:#fff;border:1px solid #e5e7eb;border-radius:18px;padding:36px;text-align:center;box-shadow:0 10px 30px rgba(0,0,0,.08);box-sizing:border-box;">' +
            '<div style="width:64px;height:64px;border-radius:50%;background:' + color + ';color:#fff;display:flex;align-items:center;justify-content:center;margin:0 auto 20px;font-size:30px;">' + (ok ? '✓' : '!') + '</div>' +
            '<h1 style="color:#5D2A86;margin:0 0 12px;">' + title + '</h1>' +
            '<p style="color:#4b5563;line-height:1.6;margin:0 0 26px;">' + message + '</p>' +
            action +
          '</div>' +
        '</body></html>'
      ).setTitle(title + ' - BizOS');
    }

    var template = HtmlService.createTemplateFromFile('index');
    template.showLogin = String(params.page || '').toLowerCase() === 'login';

    return template.evaluate()
      .setTitle('BizOS - Business Operating System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('Master doGet error:', error);
    // Never leave a normal user on an Apps Script error/fallback page.
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
