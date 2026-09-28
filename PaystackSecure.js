// ============================================================
// PaystackSecure.js
// Secure Paystack checkout + server-side verification.
// Secrets are read only from Script Properties.
// ============================================================

function getPaystackSecretKey_() {
  var key = PropertiesService.getScriptProperties().getProperty('PAYSTACK_SECRET_KEY');
  if (!key) throw new Error('PAYSTACK_SECRET_KEY is not configured in Script Properties.');
  return String(key).trim();
}

function getPaystackPublicKey_() {
  var key = PropertiesService.getScriptProperties().getProperty('PAYSTACK_PUBLIC_KEY');
  if (!key) throw new Error('PAYSTACK_PUBLIC_KEY is not configured in Script Properties.');
  return String(key).trim();
}

function getUpgradePaymentDetails(requestId, sessionId, accessToken) {
  try {
    if (!requestId) return {success:false,code:'REQUEST_ID_REQUIRED',message:'Upgrade request ID is required.'};

    var user = sessionId ? validateUpgradeSession(sessionId) : null;
    if (!user && accessToken) {
      var access = validatePaymentAccessToken_(accessToken, requestId);
      if (access) user = {email:access.email};
    }
    if (!user) return {success:false,code:'UNAUTHORIZED',message:'Your payment-page access has expired. Please return to BizOS and reopen the payment request.'};

    var request = getUpgradeRequest(requestId);

    // The requestId in a resumed payment URL is only a routing hint. If the
    // direct lookup misses it, recover the authenticated user's latest valid
    // pending request instead of forcing the user to submit a duplicate
    // upgrade request. The fallback remains account-bound through the
    // validated session and pending-request lookup.
    if (!request) {
      var latestPending = getLatestUpgradeRequestForUser_(user.email);
      if (latestPending && latestPending.status === 'pending_payment' && !latestPending.expired) {
        requestId = latestPending.requestId;
        request = getUpgradeRequest(requestId);
      }
    }

    if (!request) return {success:false,code:'REQUEST_NOT_FOUND',message:'Upgrade request not found.'};
    if (String(request.email || '').trim().toLowerCase() !== String(user.email || '').trim().toLowerCase()) return {success:false,code:'REQUEST_ACCESS_DENIED',message:'This upgrade request does not belong to the current BizOS account.'};

    var requestEmail = String(request.email || '').trim().toLowerCase();
    var sessionEmail = String(user.email || '').trim().toLowerCase();
    if (!requestEmail || requestEmail !== sessionEmail) {
      return {success:false,code:'REQUEST_ACCESS_DENIED',message:'This upgrade request does not belong to the current BizOS account.'};
    }

    var status = String(request.status || '').trim().toLowerCase();
    if (status === 'pending_payment' && isUpgradeRequestExpired_(request.createdAt)) {
      markUpgradeRequestExpired_(requestId);
      return {success:false,code:'REQUEST_EXPIRED',message:'This upgrade request expired after 10 days. Please submit a new upgrade request.'};
    }
    if (status !== 'pending_payment') {
      return {success:false,code:'INVALID_PAYMENT_STATE',message:'This upgrade request is no longer awaiting payment.'};
    }

    var amount = Number(request.amount || 0);
    if (!amount || amount <= 0) {
      return {success:false,code:'INVALID_PAYMENT_AMOUNT',message:'Upgrade request has an invalid payment amount.'};
    }

    return {
      success:true,
      requestId:String(request.requestId || requestId),
      email:String(request.email || ''),
      businessName:String(request.businessName || ''),
      tier:String(request.tier || 'sovereign'),
      displayAmount:amount,
      currency:String(request.currency || 'NGN').toUpperCase(),
      country:String(request.country || '')
    };
  } catch (error) {
    console.error('getUpgradePaymentDetails error:', error);
    return {success:false,code:'PAYMENT_DETAILS_ERROR',message:error.message || 'Unable to load upgrade payment details.'};
  }
}

function getPaystackCheckoutDetails(requestId, sessionId, accessToken) {
  try {
    if (!requestId) return {success:false,code:'REQUEST_ID_REQUIRED',message:'Request ID is required.'};
    var user = sessionId ? validateUpgradeSession(sessionId) : null;
    if (!user && accessToken) {
      var access = validatePaymentAccessToken_(accessToken, requestId);
      if (access) user = {email:access.email};
    }
    if (!user) return {success:false,code:'UNAUTHORIZED',message:'Your payment-page access has expired. Please return to BizOS and reopen the payment request.'};

    var request = getUpgradeRequest(requestId);
    if (!request) return {success:false,code:'REQUEST_NOT_FOUND',message:'Upgrade request not found.'};
    if (String(request.email || '').trim().toLowerCase() !== String(user.email || '').trim().toLowerCase()) return {success:false,code:'REQUEST_ACCESS_DENIED',message:'This upgrade request does not belong to the current BizOS account.'};

    var status = String(request.status || '').toLowerCase();
    if (status === 'pending_payment' && isUpgradeRequestExpired_(request.createdAt)) {
      markUpgradeRequestExpired_(requestId);
      return {success:false,code:'REQUEST_EXPIRED',message:'This upgrade request expired after 10 days. Please submit a new upgrade request.'};
    }
    if (status !== 'pending_payment') {
      // The customer may be returning to the payment URL after Paystack already
      // confirmed the transaction. Never present a confirmed payment as payable again.
      if (status === 'payment_confirmed' || status === 'provisioning' || status === 'provisioning_failed' || status === 'provisioned' || status === 'active') {
        var setupStatus = null;
        try {
          setupStatus = getUpgradeProvisioningStatus(requestId, sessionId, accessToken);
        } catch (statusError) {
          console.error('getPaystackCheckoutDetails setup status error:', statusError);
        }
        if (setupStatus && setupStatus.success && (setupStatus.webAppUrl || setupStatus.landingUrl)) {
          return {
            success:true,
            requestId:requestId,
            paymentConfirmed:true,
            ready:true,
            idempotent:true,
            webAppUrl:setupStatus.webAppUrl || '',
            landingUrl:setupStatus.landingUrl || setupStatus.webAppUrl || '',
            status:status,
            message:'Your payment was already confirmed and your BizOS workspace is ready.'
          };
        }
        return {
          success:true,
          requestId:requestId,
          paymentConfirmed:true,
          ready:false,
          retryable:status === 'provisioning_failed',
          status:status,
          message: status === 'provisioning_failed'
            ? 'Your payment is confirmed, but workspace setup needs to be continued. You do not need to pay again.'
            : 'Your payment is confirmed. Your BizOS workspace is still being prepared. You do not need to pay again.'
        };
      }
      return {success:false,code:'INVALID_PAYMENT_STATE',message:'This upgrade request is not awaiting payment.'};
    }

    var amount = Number(request.amount || 0);
    if (!amount || amount <= 0) {
      return {success:false,code:'INVALID_PAYMENT_AMOUNT',message:'Upgrade request has an invalid payment amount.'};
    }

    var currency = String(request.currency || 'NGN').toUpperCase();
    var reference = 'BIZOS-' + String(requestId).replace(/[^A-Za-z0-9.=-]/g,'-') + '-' + Utilities.getUuid().replace(/-/g,'').substring(0,8);

    savePaystackReference_(requestId, reference);

    return {
      success:true,
      requestId:requestId,
      email:String(request.email || ''),
      businessName:String(request.businessName || ''),
      amount:Math.round(amount * 100),
      displayAmount:amount,
      currency:currency,
      reference:reference,
      publicKey:getPaystackPublicKey_()
    };
  } catch (error) {
    console.error('getPaystackCheckoutDetails error:', error);
    return {success:false,code:'PAYSTACK_CHECKOUT_CONFIG_ERROR',message:error.message || 'Unable to prepare payment.'};
  }
}

function savePaystackReference_(requestId, reference) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payments');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  if (!data.length) return;
  var headers = data[0];
  var requestIdCol = headers.indexOf('Request_ID');
  var refCol = headers.indexOf('Transaction_Ref');
  if (requestIdCol === -1 || refCol === -1) return;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][requestIdCol]) === String(requestId)) {
      sheet.getRange(i + 1, refCol + 1).setValue(reference);
      return;
    }
  }
}

function verifyPaystackPaymentAndProvisionSecure(reference, requestId, sessionId, accessToken) {
  try {
    var user = sessionId ? validateUpgradeSession(sessionId) : null;
    if (!user && accessToken) {
      var access = validatePaymentAccessToken_(accessToken, requestId);
      if (access) user = {email:access.email};
    }
    if (!user) return {success:false,code:'UNAUTHORIZED',message:'Your payment-page access has expired. Please return to BizOS and reopen the payment request.'};
    if (!reference || !requestId) {
      return {success:false,code:'INVALID_PAYMENT_COMPLETION',message:'Payment reference and request ID are required.'};
    }

    // Do not hold a Script Lock while calling Paystack. Provisioning owns the
    // single lock, which prevents duplicate client work without blocking the
    // payment verification network request.
    var request = getUpgradeRequest(requestId);
    if (!request) return {success:false,code:'REQUEST_NOT_FOUND',message:'Upgrade request not found.'};
    if (String(request.email || '').trim().toLowerCase() !== String(user.email || '').trim().toLowerCase()) return {success:false,code:'REQUEST_ACCESS_DENIED',message:'This upgrade request does not belong to the current BizOS account.'};
    if (String(request.status || '').trim().toLowerCase() === 'pending_payment' && isUpgradeRequestExpired_(request.createdAt)) {
      markUpgradeRequestExpired_(requestId);
      return {success:false,code:'REQUEST_EXPIRED',message:'This upgrade request expired after 10 days. Please submit a new upgrade request.'};
    }

    var response = UrlFetchApp.fetch(
      'https://api.paystack.co/transaction/verify/' + encodeURIComponent(reference),
      {
        method:'get',
        headers:{
          'Authorization':'Bearer ' + getPaystackSecretKey_(),
          'Content-Type':'application/json'
        },
        muteHttpExceptions:true
      }
    );

    var result = JSON.parse(response.getContentText() || '{}');
    var tx = result.data || {};
    if (!result.status || tx.status !== 'success') {
      return {success:false,code:'PAYMENT_VERIFICATION_FAILED',message:'Paystack did not confirm this transaction.'};
    }

    var expectedAmount = Math.round(Number(request.amount || 0) * 100);
    var paidAmount = Number(tx.amount || 0);
    var expectedCurrency = String(request.currency || '').toUpperCase();
    var paidCurrency = String(tx.currency || '').toUpperCase();
    var expectedEmail = String(request.email || '').trim().toLowerCase();
    var paidEmail = String((tx.customer && tx.customer.email) || '').trim().toLowerCase();

    if (paidAmount !== expectedAmount) {
      return {success:false,code:'PAYMENT_AMOUNT_MISMATCH',message:'Payment amount does not match the upgrade amount.'};
    }
    if (expectedCurrency && paidCurrency !== expectedCurrency) {
      return {success:false,code:'PAYMENT_CURRENCY_MISMATCH',message:'Payment currency does not match the upgrade currency.'};
    }
    if (expectedEmail && paidEmail && paidEmail !== expectedEmail) {
      return {success:false,code:'PAYMENT_EMAIL_MISMATCH',message:'Payment email does not match the upgrade request.'};
    }

    markPaymentConfirmed_(requestId, reference, tx);

    var provisioning = provisionConfirmedUpgradeRequest(requestId);
    if (!provisioning || !provisioning.success) {
      return {
        success:false,
        code:'PROVISIONING_FAILED_AFTER_PAYMENT',
        paymentVerified:true,
        message:'Payment was verified, but your BizOS workspace could not be finished yet. You can continue setup without paying again.',
        provisioning:provisioning
      };
    }

    return {
      success:true,
      paymentVerified:true,
      idempotent:!!provisioning.idempotent,
      requestId:requestId,
      reference:reference,
      clientId:provisioning.clientId,
      webAppUrl:provisioning.webAppUrl,
      landingUrl:provisioning.landingUrl,
      deploymentId:provisioning.deploymentId,
      message:provisioning.message || 'Your BizOS workspace is ready.'
    };
  } catch (error) {
    console.error('verifyPaystackPaymentAndProvisionSecure error:', error);
    return {success:false,code:'PAYMENT_PROVISIONING_ERROR',message:error.message || 'Payment completion failed.'};
  }
}

function markPaymentConfirmed_(requestId, reference, transaction) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var upgrade = ss.getSheetByName('Upgrade_Requests');
  if (upgrade) {
    var data = upgrade.getDataRange().getValues();
    var headers = data[0] || [];
    var requestIdCol = headers.indexOf('Request_ID');
    var statusCol = headers.indexOf('Status');
    var updatedCol = headers.indexOf('Updated_At');
    if (requestIdCol !== -1 && statusCol !== -1) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][requestIdCol]) === String(requestId)) {
          upgrade.getRange(i + 1, statusCol + 1).setValue('payment_confirmed');
          if (updatedCol !== -1) upgrade.getRange(i + 1, updatedCol + 1).setValue(new Date().toISOString());
          break;
        }
      }
    }
  }

  var payments = ss.getSheetByName('Payments');
  if (payments) {
    var pdata = payments.getDataRange().getValues();
    var ph = pdata[0] || [];
    var reqCol = ph.indexOf('Request_ID');
    var statusCol2 = ph.indexOf('Status');
    var refCol = ph.indexOf('Transaction_Ref');
    var completedCol = ph.indexOf('Completed_At');
    if (reqCol !== -1 && statusCol2 !== -1) {
      for (var j = 1; j < pdata.length; j++) {
        if (String(pdata[j][reqCol]) === String(requestId)) {
          payments.getRange(j + 1, statusCol2 + 1).setValue('success');
          if (refCol !== -1) payments.getRange(j + 1, refCol + 1).setValue(reference);
          if (completedCol !== -1) payments.getRange(j + 1, completedCol + 1).setValue(new Date().toISOString());
          break;
        }
      }
    }
  }
}
