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

    var template = HtmlService.createTemplateFromFile('index');
    template.showLogin = String(params.page || '').toLowerCase() === 'login';

    return template.evaluate()
      .setTitle('BizOS - Business Operating System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('Master doGet error:', error);
    return HtmlService.createHtmlOutput('<h1>BizOS</h1><p>Unable to load BizOS.</p>');
  }
};
