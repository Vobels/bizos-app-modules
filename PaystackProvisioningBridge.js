// ============================================================
// PaystackProvisioningBridge.js
// Production bridge from verified payment to safe provisioning.
// ============================================================

function verifyPaystackPaymentAndProvision(reference, requestId) {
  try {
    if (!reference || !requestId) {
      return { success:false, code:'INVALID_PAYMENT_COMPLETION', message:'Payment reference and request ID are required.' };
    }

    // Existing verifier remains the authoritative payment verifier.
    var verification = verifyPaystackPayment(reference, requestId);
    if (!verification || !verification.success) {
      return verification || { success:false, code:'PAYMENT_VERIFICATION_FAILED', message:'Payment verification failed.' };
    }

    // The verifier has confirmed the payment server-side. Provisioning now uses
    // only the server-side upgrade request and the safe/idempotent provisioning path.
    var provisioning = provisionConfirmedUpgradeRequest(requestId);
    if (!provisioning || !provisioning.success) {
      return {
        success:false,
        code:'PROVISIONING_FAILED_AFTER_PAYMENT',
        paymentVerified:true,
        message:'Payment was verified, but client provisioning could not be completed. No incomplete client deployment was activated.',
        provisioning:provisioning
      };
    }

    return {
      success:true,
      paymentVerified:true,
      idempotent:!!provisioning.idempotent,
      requestId:requestId,
      clientId:provisioning.clientId,
      webAppUrl:provisioning.webAppUrl,
      landingUrl:provisioning.landingUrl,
      deploymentId:provisioning.deploymentId,
      message:provisioning.message || 'Payment confirmed and client provisioned successfully.'
    };
  } catch (error) {
    console.error('verifyPaystackPaymentAndProvision error:', error);
    return { success:false, code:'PAYMENT_PROVISIONING_ERROR', message:error.message || 'Payment completion failed.' };
  }
}
