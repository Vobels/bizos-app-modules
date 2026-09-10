// ============================================================
// ZZZ_Z_PublicUrlFix.js
// FINAL USER-FACING URL + REGISTRATION/VERIFICATION COPY HARDENING
// ============================================================
// This file is intentionally last alphabetically so these final
// assignments override older implementations without rewriting them.
// ============================================================

// Single public BizOS URL source.
getPublicBizOSUrl_ = function() {
  var base = '';
  try {
    base = CONFIG && CONFIG.URLS ? CONFIG.URLS.base : '';
  } catch (e) {}
  if (!base) base = 'https://bizos.higroups.com';
  return String(base).replace(/\/$/, '');
};

// Backward-compatible URL helper used by older email/registration code.
getAppUrl = function() {
  return getPublicBizOSUrl_();
};

// Final verification email: keep the useful anti-spam instruction,
// but never tell a person who is reading the email that they "didn't see" it.
sendVerificationEmail = function(email, businessName, businessId, token) {
  var verificationUrl = getPublicBizOSUrl_() + '/verify?token=' + encodeURIComponent(token);
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

// Improve the post-signup message without replacing the registration logic.
(function() {
  var originalRegisterBusiness = registerBusiness;
  registerBusiness = function(businessData) {
    var response = originalRegisterBusiness(businessData);
    if (response && response.success && response.requiresVerification) {
      var email = String((businessData && businessData.ownerEmail) || '').trim();
      response.message = 'Account created! We\'ve sent a verification link to ' + email + '. Please check your inbox or Spam/Junk folder, then click the link in the email to activate your account.';
    }
    return response;
  };
})();

console.log('ZZZ final public URL and user-flow fixes loaded');
