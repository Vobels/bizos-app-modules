// ============================================================
// ZZZ_Z_FinalEmailFix.js
// FINAL verification-email override
// ============================================================
// This file is intentionally last alphabetically so the final
// verification-email implementation wins over older Email.js code.

(function() {
  function publicBizOSUrl_() {
    var base = '';
    try {
      base = CONFIG && CONFIG.URLS ? CONFIG.URLS.base : '';
    } catch (e) {}
    if (!base) base = 'https://bizos.higroups.com';
    return String(base).replace(/\/$/, '');
  }

  sendVerificationEmail = function(email, businessName, businessId, token) {
    var verificationUrl = publicBizOSUrl_() + '/verify?token=' + encodeURIComponent(token);
    var content =
      '<p>Welcome to BizOS, <strong>' + String(businessName || 'there') + '</strong>!</p>' +
      '<p>Please verify your email address to activate your account.</p>' +
      '<p>This link expires in <strong>24 hours</strong>.</p>' +
      '<p style="background:#f3f4f6;padding:12px;border-radius:8px;color:#4b5563;font-size:13px;">To make sure you continue receiving BizOS messages, please mark this email as <strong>Not Spam</strong> and add BizOS to your trusted contacts or address book.</p>';

    return sendBrandedEmail(
      email,
      'Verify your email - ' + String(businessName || 'BizOS'),
      'Verify Your Email',
      content,
      'Verify Email',
      verificationUrl
    );
  };

  console.log('ZZZ final verification email override loaded');
})();
