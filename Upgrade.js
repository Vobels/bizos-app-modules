//upgrade.gs
// ==================== UPGRADE.GS - COMPLETE WORKING VERSION ====================
// Paystack credentials are stored in Apps Script Project Settings > Script Properties.

function getLegacyPaystackSecretKey_() {
  var key = PropertiesService.getScriptProperties().getProperty('PAYSTACK_SECRET_KEY');
  if (!key) throw new Error('PAYSTACK_SECRET_KEY is not configured in Script Properties.');
  return String(key).trim();
}

function verifyPaystackPayment(reference, requestId) {
  return verifyPaystackPaymentAndProvisionSecure(reference, requestId);
}

function sendPaymentConfirmationEmailOnce_(email, businessName, requestId) {
  try {
    var key = 'PAYMENT_CONFIRMATION_EMAIL_SENT_' + String(requestId || '').trim();
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty(key) === '1') return;
    sendPaymentConfirmationEmail(email, businessName, requestId);
    props.setProperty(key, '1');
  } catch (error) {
    console.error('sendPaymentConfirmationEmailOnce_ error:', error);
  }
}

function sendPaymentConfirmationEmail(email, businessName, requestId) {
  try {
    var subject = 'BizOS - Payment Received for ' + businessName;
    var body = '<div style="font-family: Arial, sans-serif;">' +
      '<h2>Payment Received! 🎉</h2>' +
      '<p>We have received your payment for <strong>' + businessName + '</strong>.</p>' +
      '<p><strong>Request ID:</strong> ' + requestId + '</p>' +
      '<p>Your BizOS workspace is now being prepared. You will receive your workspace access once provisioning is complete.</p>' +
      '<p><a href="' + (getAuthoritativeBizOSWebAppUrl_() || getPublicBizOSUrl_()) + '?page=deployment-status&requestId=' + encodeURIComponent(requestId) + '" style="display:inline-block;padding:11px 16px;background:#2563eb;color:#ffffff;text-decoration:none;border-radius:8px;font-weight:600;">View Deployment Status</a></p>' +
      '<p>You do not need to pay again. You can return to this page at any time to see your deployment status and launch link when your workspace is ready.</p>' +
      '<br><p>Thank you for choosing BizOS!</p></div>';
    MailApp.sendEmail({to: email, subject: subject, htmlBody: body});
  } catch (error) {
    console.error('Email error:', error);
  }
}

function sendDeploymentConfirmationEmailOnce_(email, businessName, requestId, launchUrl) {
  try {
    email = String(email || '').trim();
    requestId = String(requestId || '').trim();
    launchUrl = String(launchUrl || '').trim();
    if (!email || !requestId || !launchUrl) return;
    var key = 'DEPLOYMENT_CONFIRMATION_EMAIL_SENT_' + requestId;
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty(key) === '1') return;
    var publicUrl = getAuthoritativeBizOSWebAppUrl_() || getPublicBizOSUrl_();
    var statusUrl = publicUrl + '?page=deployment-confirmation&requestId=' + encodeURIComponent(requestId);
    var safeBusiness = String(businessName || 'your business').replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});
    var safeLaunch = launchUrl.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
    var body = '<div style="font-family:Arial,sans-serif;max-width:640px;margin:auto;color:#1f2937">' +
      '<h2>BizOS Workspace Deployed</h2>' +
      '<p>Your BizOS workspace for <strong>' + safeBusiness + '</strong> has been successfully deployed.</p>' +
      '<p><strong>Your payment has been confirmed and your workspace is ready.</strong></p>' +
      '<p><a href="' + safeLaunch + '" style="display:inline-block;padding:12px 18px;background:#2563eb;color:#fff;text-decoration:none;border-radius:8px;font-weight:600">Launch My BizOS</a></p>' +
      '<p><a href="' + statusUrl + '" style="color:#2563eb">View Deployment Confirmation</a></p>' +
      '<p>Request ID: ' + requestId + '</p>' +
      '</div>';
    MailApp.sendEmail({to:email,subject:'BizOS Workspace Deployed - '+(businessName||'Your Business'),htmlBody:body});
    props.setProperty(key,'1');
  } catch(error) {
    console.error('sendDeploymentConfirmationEmailOnce_ error:', error);
  }
}

function getBankDetails() {
  return {success: true, bankName: 'Example Bank', accountName: 'BizOS Global', accountNumber: '1234567890'};
}

/**
 * Handle upgrade request submission from frontend.
 * Account email remains the authenticated login identity.
 * Workspace_Email is the Google/Workspace destination used for provisioning.
 */
function handleUpgradeRequest(upgradeData, sessionId) {
  var upgradeLock = null;
  try {
    upgradeData = upgradeData || {};
    if (upgradeData.editRequestId) return updatePendingUpgradeRequest_(upgradeData, sessionId, String(upgradeData.editRequestId));
    var user = validateUpgradeSession(sessionId);
    if (!user) return {success:false, message:'Session expired. Please login again.'};

    if (user.subscriptionTier !== 'free' && user.subscriptionTier !== 'starter') {
      return {success:false, message:'You are already on a higher tier!'};
    }

    upgradeData = upgradeData || {};
    var country = upgradeData.country || upgradeData.originalCountry;
    var businessDetails = upgradeData.businessDetails || {};
    var tier = upgradeData.tier || 'sovereign';
    var certificateBase64 = upgradeData.certificateBase64;
    var certificateName = upgradeData.certificateName;
    var workspaceEmail = String(upgradeData.workspaceEmail || user.email || '').trim().toLowerCase();

    if (!country) return {success:false, message:'Please select your business country'};
    if (!/^\S+@\S+\.\S+$/.test(workspaceEmail)) return {success:false, message:'Please provide a valid workspace email.'};

    var countryConfig = getCountryRequirements(country);
    if (countryConfig.success && countryConfig.requiredFields) {
      for (var i = 0; i < countryConfig.requiredFields.length; i++) {
        var field = countryConfig.requiredFields[i];
        if (!businessDetails[field]) {
          return {success:false, message:'Missing required field: ' + ((countryConfig.fieldLabels && countryConfig.fieldLabels[field]) || field)};
        }
      }
    }

    // Serialize the final duplicate check + request creation so two rapid
    // submissions cannot create two open upgrade requests for the same account.
    upgradeLock = LockService.getScriptLock();
    upgradeLock.waitLock(10000);

    // Repair/validate the transactional schemas before duplicate detection.
    // This also restores a deleted header row without overwriting request data.
    ensureUpgradeRequestsSchema_();
    ensurePaymentsSchema_();

    // A confirmed/paid request must never create a second checkout.
    // This check runs before the pending-request check so even a previously
    // failed/provisioning request cannot be paid for a second time.
    var paidRequest = getLatestPaidUpgradeRequestForUser_(user.email);
    if (paidRequest) {
      var paidAccessToken = createPaymentAccessToken_(paidRequest.requestId, user.email);
      var paidPublicUrl = getAuthoritativeBizOSWebAppUrl_() || getPublicBizOSUrl_();
      var paidStatusUrl = paidPublicUrl + '?page=deployment-status&requestId=' +
        encodeURIComponent(paidRequest.requestId) + '&accessToken=' + encodeURIComponent(paidAccessToken);
      return {
        success:true,
        alreadyPaid:true,
        paymentConfirmed:true,
        requestId:paidRequest.requestId,
        paymentId:paidRequest.paymentId,
        amount:paidRequest.amount,
        currency:paidRequest.currency,
        status:paidRequest.status,
        redirectUrl:paidStatusUrl,
        message:'Your payment is already confirmed. You do not need to pay again. Continue with your existing BizOS setup.'
      };
    }

    var latestRequest = getLatestUpgradeRequestForUser_(user.email);
    if (latestRequest && latestRequest.status === 'pending_payment' && latestRequest.expired) {
      markUpgradeRequestExpired_(latestRequest.requestId);
      latestRequest.status = 'expired';
    }
    if (latestRequest && latestRequest.status === 'pending_payment' && !latestRequest.expired) {
      var pendingPublicUrl = getAuthoritativeBizOSWebAppUrl_() || getPublicBizOSUrl_();
      var pendingPaymentAccessToken = createPaymentAccessToken_(latestRequest.requestId, user.email);
      var pendingPaymentPageUrl = pendingPublicUrl + '?page=payment-method&requestId=' + encodeURIComponent(latestRequest.requestId) +
        '&country=' + encodeURIComponent(latestRequest.country || country) + '&tier=' + encodeURIComponent(latestRequest.tier || tier) +
        '&accessToken=' + encodeURIComponent(pendingPaymentAccessToken);
      return {
        success:true,
        message:'You already have an upgrade in progress. Continue with your existing payment request.',
        requestId:latestRequest.requestId,
        paymentId:latestRequest.paymentId,
        amount:latestRequest.amount,
        currency:latestRequest.currency,
        redirectUrl:pendingPaymentPageUrl,
        resumed:true
      };
    }

    var certificateUrl = '';
    var uploadedCertificateName = '';
    if (certificateBase64 && certificateName) {
      var uploadResult = uploadBusinessCertificate(certificateBase64, certificateName, country, user.businessId, user.businessName);
      if (uploadResult.success) {
        certificateUrl = uploadResult.fileUrl;
        uploadedCertificateName = uploadResult.fileName;
      } else {
        console.error('Certificate upload failed:', uploadResult.message);
      }
    }

    var requestId = 'UPG_' + Utilities.getUuid().substring(0, 8).toUpperCase();
    var paymentId = 'PAY_' + Utilities.getUuid().substring(0, 8).toUpperCase();
    var timestamp = new Date();
    var requestedCurrency = String(upgradeData.currency || '').trim().toUpperCase();
    var pricing = getLocalizedPricing(country, requestedCurrency);
    var amount = pricing.sovereign && pricing.sovereign.price || 500;
    var currency = pricing.sovereign && pricing.sovereign.code || 'USD';
    var symbol = pricing.sovereign && pricing.sovereign.currency || '$';

    var ss = getBizOSMasterSpreadsheet_();
    ensureUpgradeRequestsSchema_();
    var upgradeSheet = ss.getSheetByName('Upgrade_Requests');

    upgradeSheet.appendRow([
      requestId, user.email, workspaceEmail, user.name || user.email, user.businessName || '', country,
      JSON.stringify(businessDetails), tier, paymentId, amount, currency, 'pending_payment',
      certificateUrl, uploadedCertificateName, timestamp.toISOString(), ''
    ]);

    ensurePaymentsSchema_();
    var paymentSheet = ss.getSheetByName('Payments');
    paymentSheet.appendRow([
      paymentId, requestId, user.email, workspaceEmail, amount, currency, 'pending', '', timestamp.toISOString(), ''
    ]);
    snapshotPaymentFx_(requestId, pricing);

    if (upgradeLock) {
      upgradeLock.releaseLock();
      upgradeLock = null;
    }

    sendUpgradeRequestEmail(user.email, user.name, requestId, country, tier, amount, symbol);
    sendAdminUpgradeNotification(user.email, user.name, requestId, country, businessDetails, certificateUrl);

    // Keep the payment flow on the Apps Script deployment that owns the
    // Upgrade_Requests sheet. This preserves requestId and guarantees the
    // payment page reads the same backend/workspace that created the request.
    var publicUrl = getAuthoritativeBizOSWebAppUrl_() || getPublicBizOSUrl_();
    var paymentAccessToken = createPaymentAccessToken_(requestId, user.email);
    var paymentPageUrl = publicUrl + '?page=payment-method&requestId=' + encodeURIComponent(requestId) +
      '&country=' + encodeURIComponent(country) + '&tier=' + encodeURIComponent(tier) +
      '&accessToken=' + encodeURIComponent(paymentAccessToken);

    return {
      success:true,
      message:'Upgrade request submitted! Please complete payment to activate your Sovereign tier.',
      requestId:requestId,
      paymentId:paymentId,
      amount:amount,
      currency:symbol,
      redirectUrl:paymentPageUrl
    };
  } catch (error) {
    console.error('Upgrade request error:', error);
    return {success:false, message:'We could not start your upgrade right now. Please try again.'};
  } finally {
    if (upgradeLock) {
      try { upgradeLock.releaseLock(); } catch (ignore) {}
    }
  }
}


/**
 * Return the latest upgrade request for an authenticated user, regardless of status.
 * The latest row is authoritative: only a latest pending_payment request can
 * be resumed or block creation of a new request.
 */
var UPGRADE_REQUEST_TTL_MS = 10 * 24 * 60 * 60 * 1000;

function isUpgradeRequestExpired_(createdAt) {
  var createdDate = createdAt instanceof Date ? createdAt : new Date(createdAt);
  if (isNaN(createdDate.getTime())) return true;
  return Date.now() - createdDate.getTime() >= UPGRADE_REQUEST_TTL_MS;
}

function markUpgradeRequestExpired_(requestId) {
  try {
    ensureUpgradeRequestsSchema_();
    var sheet = getBizOSMasterSpreadsheet_().getSheetByName('Upgrade_Requests');
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    if (!data.length) return;
    var headers = data[0] || [];
    var requestIdCol = headers.indexOf('Request_ID');
    var statusCol = headers.indexOf('Status');
    var updatedCol = headers.indexOf('Updated_At');
    if (requestIdCol === -1 || statusCol === -1) return;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][requestIdCol]) === String(requestId)) {
        if (String(data[i][statusCol] || '').toLowerCase() === 'pending_payment') {
          sheet.getRange(i + 1, statusCol + 1).setValue('expired');
          if (updatedCol !== -1) sheet.getRange(i + 1, updatedCol + 1).setValue(new Date().toISOString());
        }
        return;
      }
    }
  } catch (error) {
    console.error('markUpgradeRequestExpired_ error:', error);
  }
}

function getLatestUpgradeRequestForUser_(email) {
  try {
    ensureUpgradeRequestsSchema_();
    var sheet = getBizOSMasterSpreadsheet_().getSheetByName('Upgrade_Requests');
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return null;
    var headers = data[0];
    var emailCol = headers.indexOf('Email');
    var statusCol = headers.indexOf('Status');
    var requestIdCol = headers.indexOf('Request_ID');
    var paymentIdCol = headers.indexOf('Payment_ID');
    var workspaceEmailCol = headers.indexOf('Workspace_Email');
    var businessNameCol = headers.indexOf('Business_Name');
    var tierCol = headers.indexOf('Tier');
    var amountCol = headers.indexOf('Amount');
    var currencyCol = headers.indexOf('Currency');
    var countryCol = headers.indexOf('Country');
    var createdAtCol = headers.indexOf('Created_At');
    if (emailCol === -1 || statusCol === -1 || requestIdCol === -1) return null;

    var targetEmail = String(email || '').trim().toLowerCase();
    if (!targetEmail) return null;

    for (var i = data.length - 1; i >= 1; i--) {
      var rowEmail = String(data[i][emailCol] || '').trim().toLowerCase();
      var status = String(data[i][statusCol] || '').trim().toLowerCase();
      if (rowEmail === targetEmail) {
        return {
          requestId: requestIdCol !== -1 ? String(data[i][requestIdCol] || '') : '',
          paymentId: paymentIdCol !== -1 ? String(data[i][paymentIdCol] || '') : '',
          workspaceEmail: workspaceEmailCol !== -1 ? String(data[i][workspaceEmailCol] || '') : '',
          businessName: businessNameCol !== -1 ? String(data[i][businessNameCol] || '') : '',
          tier: tierCol !== -1 ? String(data[i][tierCol] || 'sovereign') : 'sovereign',
          amount: amountCol !== -1 ? data[i][amountCol] : '',
          currency: currencyCol !== -1 ? String(data[i][currencyCol] || '') : '',
          country: countryCol !== -1 ? String(data[i][countryCol] || '') : '',
          createdAt: createdAtCol !== -1 ? data[i][createdAtCol] : '',
          status: status,
          expiresAt: (function () {
            var created = data[i][createdAtCol];
            var createdDate = created instanceof Date ? created : new Date(created);
            if (isNaN(createdDate.getTime())) return '';
            return new Date(createdDate.getTime() + UPGRADE_REQUEST_TTL_MS).toISOString();
          })(),
          expired: status === 'pending_payment' && isUpgradeRequestExpired_(data[i][createdAtCol])
        };
      }
    }
    return null;
  } catch (error) {
    console.error('getLatestUpgradeRequestForUser_ error:', error);
    return null;
  }
}

/**
 * Get the authenticated user's latest upgrade. A pending upgrade is
 * resumable only when it is the latest request for that account.
 */
function getLatestPaidUpgradeRequestForUser_(email) {
  try {
    ensureUpgradeRequestsSchema_();
    ensurePaymentsSchema_();
    var ss = getBizOSMasterSpreadsheet_();
    var upgradeSheet = ss.getSheetByName('Upgrade_Requests');
    var paymentSheet = ss.getSheetByName('Payments');
    if (!upgradeSheet || !paymentSheet) return null;

    var targetEmail = String(email || '').trim().toLowerCase();
    if (!targetEmail) return null;

    var upgradeData = upgradeSheet.getDataRange().getValues();
    var uh = upgradeData[0] || [];
    var uEmail = uh.indexOf('Email'), uRequest = uh.indexOf('Request_ID'), uPayment = uh.indexOf('Payment_ID');
    var uStatus = uh.indexOf('Status'), uWorkspace = uh.indexOf('Workspace_Email'), uBusiness = uh.indexOf('Business_Name');
    var uTier = uh.indexOf('Tier'), uAmount = uh.indexOf('Amount'), uCurrency = uh.indexOf('Currency');
    var uCountry = uh.indexOf('Country'), uCreated = uh.indexOf('Created_At');
    if (uEmail === -1 || uRequest === -1 || uStatus === -1) return null;

    var paidRequestIds = {};
    var pd = paymentSheet.getDataRange().getValues();
    var ph = pd[0] || [];
    var pRequest = ph.indexOf('Request_ID'), pStatus = ph.indexOf('Status');
    if (pRequest !== -1 && pStatus !== -1) {
      for (var p = 1; p < pd.length; p++) {
        if (String(pd[p][pStatus] || '').trim().toLowerCase() === 'success') {
          paidRequestIds[String(pd[p][pRequest] || '').trim()] = true;
        }
      }
    }

    var tracked = ['payment_confirmed','provisioning','provisioning_failed','provisioned','active'];
    var candidates = [];
    for (var i = 1; i < upgradeData.length; i++) {
      if (String(upgradeData[i][uEmail] || '').trim().toLowerCase() !== targetEmail) continue;
      var requestId = String(upgradeData[i][uRequest] || '').trim();
      var status = String(upgradeData[i][uStatus] || '').trim().toLowerCase();
      if (!requestId) continue;

      var paid = tracked.indexOf(status) !== -1 || !!paidRequestIds[requestId];
      if (!paid) continue;

      var createdValue = uCreated !== -1 ? upgradeData[i][uCreated] : '';
      var createdDate = createdValue instanceof Date ? createdValue : new Date(createdValue);
      var sortTime = isNaN(createdDate.getTime()) ? i : createdDate.getTime();

      candidates.push({
        rowNumber:i + 1,
        requestId:requestId,
        paymentId:uPayment !== -1 ? String(upgradeData[i][uPayment] || '') : '',
        workspaceEmail:uWorkspace !== -1 ? String(upgradeData[i][uWorkspace] || '') : '',
        businessName:uBusiness !== -1 ? String(upgradeData[i][uBusiness] || '') : '',
        tier:uTier !== -1 ? String(upgradeData[i][uTier] || 'sovereign') : 'sovereign',
        amount:uAmount !== -1 ? upgradeData[i][uAmount] : '',
        currency:uCurrency !== -1 ? String(upgradeData[i][uCurrency] || '') : '',
        country:uCountry !== -1 ? String(upgradeData[i][uCountry] || '') : '',
        createdAt:createdValue,
        status:status,
        sortTime:sortTime,
        reconciledFromPayment:!!paidRequestIds[requestId] && tracked.indexOf(status) === -1
      });
    }

    if (!candidates.length) return null;
    candidates.sort(function(a,b){ return b.sortTime - a.sortTime; });
    var selected = candidates[0];

    // Keep the upgrade row truthful when the payment row already says success.
    if (selected.reconciledFromPayment) {
      if (uStatus !== -1) upgradeSheet.getRange(selected.rowNumber, uStatus + 1).setValue('payment_confirmed');
      selected.status = 'payment_confirmed';
      selected.reconciledFromPayment = false;
    }
    return selected;
  } catch (error) {
    console.error('getLatestPaidUpgradeRequestForUser_ error:', error);
    return null;
  }
}

function getCurrentUpgradePaymentStatus(sessionId) {
  try {
    var user = validateUpgradeSession(sessionId);
    if (!user) return {success:false,code:'UNAUTHORIZED',message:'Session expired.'};

    // Look across the account's upgrade/payment history, not only the newest
    // upgrade row. This is important when a customer has an older confirmed
    // payment and later created/resumed another request.
    var latest = getLatestPaidUpgradeRequestForUser_(user.email);
    if (!latest) return {success:true,active:false};

    var status = String(latest.status || '').trim().toLowerCase();
    var accessToken = createPaymentAccessToken_(latest.requestId, user.email);
    var resumeUrl = (getAuthoritativeBizOSWebAppUrl_() || getPublicBizOSUrl_()) +
      '?page=deployment-status&requestId=' + encodeURIComponent(latest.requestId);

    var setup = getUpgradeProvisioningStatus(latest.requestId, sessionId, accessToken);
    return {
      success:true,
      active:true,
      requestId:latest.requestId,
      status:status,
      paymentConfirmed:true,
      ready:!!(setup && setup.success && (setup.webAppUrl || setup.landingUrl)),
      webAppUrl:setup && setup.webAppUrl || '',
      landingUrl:setup && setup.landingUrl || '',
      resumeUrl:resumeUrl,
      message: status === 'provisioned' || status === 'active'
        ? 'Your payment is confirmed and your BizOS workspace is ready.'
        : status === 'provisioning_failed'
          ? 'Your payment is confirmed, but workspace setup still needs to be continued. You do not need to pay again.'
          : 'Your payment is confirmed. Your BizOS workspace is being prepared. You do not need to pay again.'
    };
  } catch (error) {
    console.error('getCurrentUpgradePaymentStatus error:', error);
    return {success:false,code:'PAYMENT_STATUS_ERROR',message:'We could not check your upgrade payment status right now.'};
  }
}
function getPendingUpgradeRequestStatus(businessId, sessionId) {
  try {
    var user = validateUpgradeSession(sessionId);
    if (!user) return {success:false, message:'Session expired. Please login again.'};
    if (businessId && user.businessId && String(businessId) !== String(user.businessId)) {
      return {success:false, message:'Business does not match the current session.'};
    }
    var latestRequest = getLatestUpgradeRequestForUser_(user.email);
    if (!latestRequest || latestRequest.status !== 'pending_payment' || latestRequest.expired) {
      if (latestRequest && latestRequest.status === 'pending_payment' && latestRequest.expired) markUpgradeRequestExpired_(latestRequest.requestId);
      return {success:true, hasPending:false};
    }
    var pending = latestRequest;
    return {
      success:true,
      hasPending:true,
      requestId:pending.requestId,
      paymentId:pending.paymentId,
      businessName:pending.businessName || user.businessName || '',
      tier:pending.tier || 'sovereign',
      amount:pending.amount,
      currency:pending.currency,
      country:pending.country,
      workspaceEmail:pending.workspaceEmail,
      businessDetails:(function(){var full=getUpgradeRequest(pending.requestId);return full&&full.businessDetails?full.businessDetails:{};})(),
      certificateUrl:(function(){var full=getUpgradeRequest(pending.requestId);return full&&full.certificateUrl?full.certificateUrl:'';})(),
      certificateName:(function(){var full=getUpgradeRequest(pending.requestId);return full&&full.certificateName?full.certificateName:'';})(),
      createdAt:pending.createdAt,
      status:pending.status,
      redirectUrl: (getAuthoritativeBizOSWebAppUrl_() || getPublicBizOSUrl_()) + '?page=payment-method&requestId=' + encodeURIComponent(pending.requestId) +
        '&country=' + encodeURIComponent(pending.country || '') + '&tier=' + encodeURIComponent(pending.tier || 'sovereign') +
        '&accessToken=' + encodeURIComponent(createPaymentAccessToken_(pending.requestId, user.email))
    };
  } catch (error) {
    console.error('getPendingUpgradeRequestStatus error:', error);
    return {success:false, message:error.message || 'Unable to check upgrade status.'};
  }
}

/**
 * Update an existing pending upgrade request in place.
 * Keeps the one-active-request rule and recalculates pricing server-side.
 */
function updatePendingUpgradeRequest_(upgradeData, sessionId, requestId) {
  var editLock = null;
  try {
    var user = validateUpgradeSession(sessionId);
    if (!user) return {success:false, message:'Session expired. Please login again.'};
    if (user.subscriptionTier !== 'free' && user.subscriptionTier !== 'starter') return {success:false, message:'You are already on a higher tier!'};

    upgradeData = upgradeData || {};
    var country = String(upgradeData.country || '').trim();
    var businessDetails = upgradeData.businessDetails || {};
    var tier = String(upgradeData.tier || 'sovereign').trim() || 'sovereign';
    var workspaceEmail = String(upgradeData.workspaceEmail || user.email || '').trim().toLowerCase();
    if (!country) return {success:false, message:'Please select your business country'};
    if (!/^\S+@\S+\.\S+$/.test(workspaceEmail)) return {success:false, message:'Please provide a valid workspace email.'};

    var countryConfig = getCountryRequirements(country);
    if (countryConfig && countryConfig.requiredFields) {
      for (var i = 0; i < countryConfig.requiredFields.length; i++) {
        var field = countryConfig.requiredFields[i];
        if (!businessDetails[field]) return {success:false, message:'Missing required field: ' + ((countryConfig.fieldLabels && countryConfig.fieldLabels[field]) || field)};
      }
    }

    editLock = LockService.getScriptLock();
    editLock.waitLock(10000);
    ensureUpgradeRequestsSchema_();
    ensurePaymentsSchema_();

    var request = getUpgradeRequest(requestId);
    if (!request) return {success:false, message:'Upgrade request not found.'};
    if (String(request.email || '').trim().toLowerCase() !== String(user.email || '').trim().toLowerCase()) return {success:false, message:'This upgrade request does not belong to the current BizOS account.'};
    if (String(request.status || '').toLowerCase() !== 'pending_payment') return {success:false, message:'This upgrade request is no longer editable.'};
    if (isUpgradeRequestExpired_(request.createdAt)) {
      markUpgradeRequestExpired_(requestId);
      return {success:false, message:'This upgrade request has expired. Please start a new upgrade request.'};
    }

    var requestedCurrency = String(upgradeData.currency || '').trim().toUpperCase();
    var pricing = getLocalizedPricing(country, requestedCurrency);
    var amount = pricing.sovereign && Number(pricing.sovereign.price) || 500;
    var currency = pricing.sovereign && String(pricing.sovereign.code || 'USD').toUpperCase() || 'USD';

    var ss = getBizOSMasterSpreadsheet_();
    var upgradeSheet = ss.getSheetByName('Upgrade_Requests');
    var rows = upgradeSheet.getDataRange().getValues();
    var headers = rows[0] || [];
    var idCol=headers.indexOf('Request_ID'), workspaceCol=headers.indexOf('Workspace_Email'), countryCol=headers.indexOf('Country');
    var detailsCol=headers.indexOf('Business_Details'), tierCol=headers.indexOf('Tier'), amountCol=headers.indexOf('Amount');
    var currencyCol=headers.indexOf('Currency'), certificateUrlCol=headers.indexOf('Certificate_URL'), certificateNameCol=headers.indexOf('Certificate_Name'), updatedCol=headers.indexOf('Updated_At');
    var rowNumber=-1;
    for(var r=1;r<rows.length;r++){if(idCol!==-1&&String(rows[r][idCol]||'')===String(requestId)){rowNumber=r+1;break;}}
    if(rowNumber===-1) return {success:false, message:'Upgrade request not found.'};

    if(workspaceCol!==-1) upgradeSheet.getRange(rowNumber,workspaceCol+1).setValue(workspaceEmail);
    if(countryCol!==-1) upgradeSheet.getRange(rowNumber,countryCol+1).setValue(country);
    if(detailsCol!==-1) upgradeSheet.getRange(rowNumber,detailsCol+1).setValue(JSON.stringify(businessDetails));
    if(tierCol!==-1) upgradeSheet.getRange(rowNumber,tierCol+1).setValue(tier);
    if(amountCol!==-1) upgradeSheet.getRange(rowNumber,amountCol+1).setValue(amount);
    if(currencyCol!==-1) upgradeSheet.getRange(rowNumber,currencyCol+1).setValue(currency);

    // A country change invalidates a previously uploaded country-specific
    // certificate unless the user supplies a replacement in this edit.
    var newCertificateUrl = '';
    var newCertificateName = '';
    if (upgradeData.certificateBase64 && upgradeData.certificateName) {
      var uploadResult = uploadBusinessCertificate(upgradeData.certificateBase64, upgradeData.certificateName, country, user.businessId, user.businessName);
      if (!uploadResult.success) return {success:false, message:uploadResult.message || 'Unable to upload the replacement certificate.'};
      newCertificateUrl = uploadResult.fileUrl || '';
      newCertificateName = uploadResult.fileName || String(upgradeData.certificateName);
    }
    if (country !== String(request.country || '').trim() || newCertificateUrl) {
      if(certificateUrlCol!==-1) upgradeSheet.getRange(rowNumber,certificateUrlCol+1).setValue(newCertificateUrl);
      if(certificateNameCol!==-1) upgradeSheet.getRange(rowNumber,certificateNameCol+1).setValue(newCertificateName);
    }
    if(updatedCol!==-1) upgradeSheet.getRange(rowNumber,updatedCol+1).setValue(new Date().toISOString());

    var paymentSheet=ss.getSheetByName('Payments');
    if(paymentSheet){
      var pdata=paymentSheet.getDataRange().getValues(), ph=pdata[0]||[];
      var pReq=ph.indexOf('Request_ID'),pAmount=ph.indexOf('Amount'),pCurrency=ph.indexOf('Currency'),pStatus=ph.indexOf('Status'),pRef=ph.indexOf('Transaction_Ref'),pCompleted=ph.indexOf('Completed_At');
      for(var p=1;p<pdata.length;p++) if(pReq!==-1&&String(pdata[p][pReq]||'')===String(requestId)){
        var prow=p+1;
        if(pAmount!==-1) paymentSheet.getRange(prow,pAmount+1).setValue(amount);
        if(pCurrency!==-1) paymentSheet.getRange(prow,pCurrency+1).setValue(currency);
        if(pStatus!==-1) paymentSheet.getRange(prow,pStatus+1).setValue('pending');
        if(pRef!==-1) paymentSheet.getRange(prow,pRef+1).setValue('');
        if(pCompleted!==-1) paymentSheet.getRange(prow,pCompleted+1).setValue('');
        break;
      }
    }

    var publicUrl=getAuthoritativeBizOSWebAppUrl_()||getPublicBizOSUrl_();
    var paymentAccessToken=createPaymentAccessToken_(requestId,user.email);
    var paymentPageUrl=publicUrl+'?page=payment-method&requestId='+encodeURIComponent(requestId)+'&country='+encodeURIComponent(country)+'&tier='+encodeURIComponent(tier)+'&accessToken='+encodeURIComponent(paymentAccessToken);
    return {success:true,message:'Your upgrade details were updated. Please review the payment amount before continuing.',requestId:requestId,paymentId:request.paymentId,amount:amount,currency:currency,redirectUrl:paymentPageUrl,updated:true};
  } catch(error) {
    console.error('updatePendingUpgradeRequest_ error:',error);
    return {success:false,message:error.message||'Unable to update your upgrade request.'};
  } finally { if(editLock){try{editLock.releaseLock();}catch(ignore){}} }
}

function sendAdminUpgradeNotification(userEmail, userName, requestId, country, businessDetails, certificateUrl) {
  try {
    var subject = 'New Sovereign Upgrade Request - ' + requestId;
    var certificateHtml = certificateUrl
      ? '<p><strong>Certificate:</strong> <a href="' + certificateUrl + '" target="_blank">View Uploaded Document</a></p>'
      : '<p><strong>Certificate:</strong> Not uploaded</p>';
    var htmlBody = '<div style="font-family: Arial, sans-serif;">' +
      '<h2>New Upgrade Request</h2>' +
      '<p><strong>User:</strong> ' + userName + ' (' + userEmail + ')</p>' +
      '<p><strong>Request ID:</strong> ' + requestId + '</p>' +
      '<p><strong>Country:</strong> ' + country + '</p>' + certificateHtml +
      '<p><strong>Business Details:</strong></p><pre>' + JSON.stringify(businessDetails, null, 2) + '</pre>' +
      '<p><strong>Status:</strong> Pending Payment</p>' +
      '<p><strong>Time:</strong> ' + new Date().toLocaleString() + '</p></div>';
    MailApp.sendEmail({to:'bizoshigroups@gmail.com', subject:subject, htmlBody:htmlBody});
  } catch (error) {
    console.error('Admin email error:', error);
  }
}

function getUpgradeRequest(requestId) {
  try {
    ensureUpgradeRequestsSchema_();
    var sheet = getBizOSMasterSpreadsheet_().getSheetByName('Upgrade_Requests');
    if (!sheet) return null;
    var data = sheet.getDataRange().getValues();
    if (!data.length) return null;
    var headers = data[0];
    var col = {
      requestId:headers.indexOf('Request_ID'), email:headers.indexOf('Email'), workspaceEmail:headers.indexOf('Workspace_Email'),
      name:headers.indexOf('Name'), businessName:headers.indexOf('Business_Name'), country:headers.indexOf('Country'),
      businessDetails:headers.indexOf('Business_Details'), tier:headers.indexOf('Tier'), paymentId:headers.indexOf('Payment_ID'),
      amount:headers.indexOf('Amount'), currency:headers.indexOf('Currency'), status:headers.indexOf('Status'),
      certificateUrl:headers.indexOf('Certificate_URL'), certificateName:headers.indexOf('Certificate_Name'), createdAt:headers.indexOf('Created_At')
    };
    if (col.requestId === -1) return null;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][col.requestId]) === String(requestId)) {
        var details = {};
        if (col.businessDetails !== -1 && data[i][col.businessDetails]) {
          try { details = JSON.parse(data[i][col.businessDetails]); } catch (ignore) { details = {}; }
        }
        return {
          requestId:col.requestId !== -1 ? data[i][col.requestId] : '',
          email:col.email !== -1 ? data[i][col.email] : '',
          workspaceEmail:col.workspaceEmail !== -1 ? data[i][col.workspaceEmail] : '',
          name:col.name !== -1 ? data[i][col.name] : '',
          businessName:col.businessName !== -1 ? data[i][col.businessName] : '',
          country:col.country !== -1 ? data[i][col.country] : '',
          businessDetails:details,
          tier:col.tier !== -1 ? data[i][col.tier] : 'sovereign',
          paymentId:col.paymentId !== -1 ? data[i][col.paymentId] : '',
          amount:col.amount !== -1 ? data[i][col.amount] : '',
          currency:col.currency !== -1 ? data[i][col.currency] : '',
          status:col.status !== -1 ? data[i][col.status] : '',
          certificateUrl:col.certificateUrl !== -1 ? data[i][col.certificateUrl] : '',
          certificateName:col.certificateName !== -1 ? data[i][col.certificateName] : '',
          createdAt:col.createdAt !== -1 ? data[i][col.createdAt] : ''
        };
      }
    }
    return null;
  } catch (error) {
    console.error('Error getting upgrade request:', error);
    return null;
  }
}

function uploadBusinessCertificate(base64Data, fileName, country, businessId, businessName) {
  try {
    if (!businessId || !businessName) throw new Error('Missing businessId or businessName');
    var cleanBase64 = base64Data.indexOf(',') !== -1 ? base64Data.split(',')[1] : base64Data;
    var rootFolder = getBizOSFolder();
    var businessFolder = getOrCreateBusinessSubfolder(rootFolder, businessId, businessName);
    var certFolder;
    var existing = businessFolder.getFoldersByName('Certificates');
    certFolder = existing.hasNext() ? existing.next() : businessFolder.createFolder('Certificates');
    var blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), getMimeType(fileName), fileName);
    var file = certFolder.createFile(blob);
    // Certificates are private business documents. BizOS stores the file URL for
    // internal verification; public link sharing is not required and would expose
    // sensitive registration documents to anyone who obtains the URL.
    return {success:true, fileUrl:file.getUrl(), fileId:file.getId(), fileName:fileName};
  } catch (error) {
    console.error('Error uploading certificate:', error);
    return {success:false, message:error.message};
  }
}

function getMimeType(fileName) {
  var ext = String(fileName).split('.').pop().toLowerCase();
  var mimeTypes = {'pdf':'application/pdf','png':'image/png','jpg':'image/jpeg','jpeg':'image/jpeg'};
  return mimeTypes[ext] || 'application/octet-stream';
}

function getLocalizedPricing(country, requestedCurrency) {
  var name = String(country || '').trim();
  var wanted = String(requestedCurrency || '').trim().toUpperCase();
  if (!wanted) {
    var countryDefaults = (CONFIG.PRICING.payment && CONFIG.PRICING.payment.countryDefaultCurrencies) || {};
    wanted = countryDefaults[name] || (CONFIG.PRICING.payment && CONFIG.PRICING.payment.defaultCurrency) || 'USD';
  }

  var converted = convertBizOSBasePriceToCurrency_(wanted);
  if (!converted.success) return converted;

  return {
    success:true,
    sovereign:{
      price:converted.price,
      currency:converted.symbol,
      code:converted.currency,
      name:converted.name,
      basePrice:converted.baseAmount,
      baseCurrency:converted.baseCurrency,
      exchangeRate:converted.rate,
      exchangeRateSource:converted.rateSource,
      exchangeRateUpdatedAt:converted.rateUpdatedAt,
      decimals:converted.decimals
    }
  };
}

function getAvailableBizOSCheckoutCurrencies_(country) {
  var rates = getPaymentExchangeRates_();
  var currencies = [];
  Object.keys(rates).forEach(function(code) {
    var cfg = getPaymentCurrencyConfig_(code);
    if (!cfg) return;
    var restrictions = (CONFIG.PRICING.payment && CONFIG.PRICING.payment.currencyCountryRestrictions) || {};
    if (restrictions[code] && restrictions[code].indexOf(String(country || '').trim()) === -1) return;
    currencies.push({code:cfg.code,name:cfg.name,symbol:cfg.symbol,rate:cfg.rate,decimals:cfg.decimals});
  });
  currencies.sort(function(a,b) {
    if (a.code === 'USD') return -1;
    if (b.code === 'USD') return 1;
    if (a.code === 'NGN') return -1;
    if (b.code === 'NGN') return 1;
    return a.code.localeCompare(b.code);
  });
  return currencies;
}

function getUpgradeCurrencyOptions(requestId, sessionId, accessToken) {
  try {
    var user = sessionId ? validateUpgradeSession(sessionId) : null;
    if (!user && accessToken) {
      var access = validatePaymentAccessToken_(accessToken, requestId);
      if (access) user = {email:access.email};
    }
    if (!user) return {success:false,code:'UNAUTHORIZED',message:'Your payment-page access has expired. Please return to BizOS and reopen the payment request.'};

    var request = getUpgradeRequest(requestId);
    if (!request) return {success:false,code:'REQUEST_NOT_FOUND',message:'Upgrade request not found.'};
    if (String(request.email || '').trim().toLowerCase() !== String(user.email || '').trim().toLowerCase()) {
      return {success:false,code:'REQUEST_ACCESS_DENIED',message:'This upgrade request does not belong to the current BizOS account.'};
    }

    var currencies = getAvailableBizOSCheckoutCurrencies_(request.country);
    return {
      success:true,
      currentCurrency:String(request.currency || getLocalizedPricing(request.country, '').sovereign.code).toUpperCase(),
      options:currencies
    };
  } catch (error) {
    console.error('getUpgradeCurrencyOptions error:', error);
    return {success:false,message:'We could not load the available currencies right now.'};
  }
}

function ensurePaymentFxColumns_(sheet) {
  if (!sheet) return;
  var headers = sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn())).getValues()[0] || [];
  var required = ['Base_Amount','Base_Currency','FX_Rate','FX_Rate_Source','FX_Rate_Updated_At'];
  required.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      headers.push(header);
    }
  });
}

function snapshotPaymentFx_(requestId, pricing) {
  try {
    var ss = getBizOSMasterSpreadsheet_();
    var sheets = [ss.getSheetByName('Upgrade_Requests'), ss.getSheetByName('Payments')];
    sheets.forEach(function(sheet) {
      if (!sheet) return;
      ensurePaymentFxColumns_(sheet);
      var data = sheet.getDataRange().getValues();
      var headers = data[0] || [];
      var idCol = headers.indexOf('Request_ID');
      if (idCol === -1) return;
      for (var i=1;i<data.length;i++) {
        if (String(data[i][idCol] || '').trim() !== String(requestId || '').trim()) continue;
        var values = {
          Base_Amount:pricing.sovereign.basePrice,
          Base_Currency:pricing.sovereign.baseCurrency,
          FX_Rate:pricing.sovereign.exchangeRate,
          FX_Rate_Source:pricing.sovereign.exchangeRateSource || 'admin',
          FX_Rate_Updated_At:pricing.sovereign.exchangeRateUpdatedAt || new Date().toISOString()
        };
        Object.keys(values).forEach(function(header) {
          var col=headers.indexOf(header);
          if(col!==-1) sheet.getRange(i+1,col+1).setValue(values[header]);
        });
        return;
      }
    });
  } catch(error) {
    console.error('snapshotPaymentFx_ error:',error);
  }
}

function updateUpgradeRequestCurrency(requestId, currency, sessionId, accessToken) {
  try {
    var user = sessionId ? validateUpgradeSession(sessionId) : null;
    if (!user && accessToken) {
      var access = validatePaymentAccessToken_(accessToken, requestId);
      if (access) user = {email:access.email};
    }
    if (!user) return {success:false,message:'Your payment-page access has expired. Please return to BizOS and reopen the payment request.'};

    var request = getUpgradeRequest(requestId);
    if (!request) return {success:false,message:'Upgrade request not found.'};
    if (String(request.email || '').trim().toLowerCase() !== String(user.email || '').trim().toLowerCase()) return {success:false,message:'This upgrade request does not belong to the current BizOS account.'};
    if (String(request.status || '').toLowerCase() !== 'pending_payment') return {success:false,message:'This upgrade request is no longer editable.'};

    var pricing = getLocalizedPricing(request.country, currency);
    if (!pricing || !pricing.success || !pricing.sovereign) return {success:false,message:(pricing && pricing.message)||'This currency is not currently available for BizOS checkout.'};

    var ss=getBizOSMasterSpreadsheet_();
    var sheet=ss.getSheetByName('Upgrade_Requests');
    var rows=sheet.getDataRange().getValues(), headers=rows[0]||[];
    var idCol=headers.indexOf('Request_ID'), amountCol=headers.indexOf('Amount'), currencyCol=headers.indexOf('Currency'), updatedCol=headers.indexOf('Updated_At');
    for(var i=1;i<rows.length;i++){
      if(idCol!==-1&&String(rows[i][idCol])===String(requestId)){
        if(amountCol!==-1) sheet.getRange(i+1,amountCol+1).setValue(pricing.sovereign.price);
        if(currencyCol!==-1) sheet.getRange(i+1,currencyCol+1).setValue(pricing.sovereign.code);
        if(updatedCol!==-1) sheet.getRange(i+1,updatedCol+1).setValue(new Date().toISOString());
        var payments=ss.getSheetByName('Payments');
        if(payments){
          var pdata=payments.getDataRange().getValues(), ph=pdata[0]||[];
          var pReq=ph.indexOf('Request_ID'),pAmount=ph.indexOf('Amount'),pCurrency=ph.indexOf('Currency');
          for(var p=1;p<pdata.length;p++) if(pReq!==-1&&String(pdata[p][pReq])===String(requestId)){
            if(pAmount!==-1) payments.getRange(p+1,pAmount+1).setValue(pricing.sovereign.price);
            if(pCurrency!==-1) payments.getRange(p+1,pCurrency+1).setValue(pricing.sovereign.code);
            break;
          }
        }
        snapshotPaymentFx_(requestId, pricing);
        return {success:true,currency:pricing.sovereign.code,amount:pricing.sovereign.price,symbol:pricing.sovereign.currency,baseAmount:pricing.sovereign.basePrice,baseCurrency:pricing.sovereign.baseCurrency,exchangeRate:pricing.sovereign.exchangeRate};
      }
    }
    return {success:false,message:'Upgrade request not found.'};
  } catch(error) {
    console.error('updateUpgradeRequestCurrency error:',error);
    return {success:false,message:'We could not update the payment currency right now. Please try again.'};
  }
}

function validateUpgradeSession(sessionId) {
  try {
    if (!sessionId) return null;
    var cache = CacheService.getScriptCache();
    var sessionData = cache.get(sessionId);
    if (!sessionData) return null;
    var session = JSON.parse(sessionData);
    var email = session.email;
    if (!email) return null;

    var userSheet = getOrCreateUserSheet();
    var data = userSheet.getDataRange().getValues();
    var headers = data[0];
    var emailCol = headers.indexOf('Email');
    var nameCol = headers.indexOf('Name');
    var businessIdCol = headers.indexOf('Business_ID');
    var businessNameCol = headers.indexOf('Business_Name');
    if (emailCol === -1) return null;

    var userRow = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][emailCol] && String(data[i][emailCol]).toLowerCase() === String(email).toLowerCase()) {
        userRow = data[i];
        break;
      }
    }
    if (!userRow) return null;

    var businessId = businessIdCol !== -1 ? userRow[businessIdCol] : '';
    var businessName = businessNameCol !== -1 ? userRow[businessNameCol] : '';
    var subscriptionTier = 'free';

    if (businessId) {
      var businessSheet = getOrCreateBusinessSheet();
      var businessData = businessSheet.getDataRange().getValues();
      var bh = businessData[0] || [];
      var bid = bh.indexOf('Business_ID');
      var tierCol = bh.indexOf('Subscription_Tier');
      var bn = bh.indexOf('Business_Name');
      for (var j = 1; j < businessData.length; j++) {
        if (bid !== -1 && String(businessData[j][bid]) === String(businessId)) {
          if (tierCol !== -1 && businessData[j][tierCol]) subscriptionTier = businessData[j][tierCol];
          if (bn !== -1 && businessData[j][bn]) businessName = businessData[j][bn];
          break;
        }
      }
    }

    return {
      email:email,
      name:nameCol !== -1 ? userRow[nameCol] : email,
      businessId:businessId,
      businessName:businessName,
      subscriptionTier:String(subscriptionTier).toLowerCase()
    };
  } catch (error) {
    console.error('validateUpgradeSession error:', error);
    return null;
  }
}

/**
 * Short-lived payment-page access handoff.
 * The token is stored only in Script Cache and is passed in the URL fragment,
 * so it is not sent to the server as a query parameter or exposed as a referrer.
 */
var PAYMENT_ACCESS_TOKEN_TTL_SECONDS = 2 * 60 * 60;

function createPaymentAccessToken_(requestId, email) {
  var token = Utilities.getUuid().replace(/-/g, '');
  var payload = {
    requestId: String(requestId || ''),
    email: String(email || '').trim().toLowerCase(),
    createdAt: Date.now()
  };
  CacheService.getScriptCache().put(
    'bizos_payment_access_' + token,
    JSON.stringify(payload),
    PAYMENT_ACCESS_TOKEN_TTL_SECONDS
  );
  return token;
}

function validatePaymentAccessToken_(token, requestId) {
  if (!token) return null;
  var raw = CacheService.getScriptCache().get('bizos_payment_access_' + String(token));
  if (!raw) return null;
  try {
    var payload = JSON.parse(raw);
    if (!payload || !payload.requestId || !payload.email) return null;
    if (String(payload.requestId) !== String(requestId || '')) return null;
    return payload;
  } catch (error) {
    console.error('validatePaymentAccessToken_ error:', error);
    return null;
  }
}
