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

function sendPaymentConfirmationEmail(email, businessName, requestId) {
  try {
    var subject = 'BizOS - Payment Received for ' + businessName;
    var body = '<div style="font-family: Arial, sans-serif;">' +
      '<h2>Payment Received! 🎉</h2>' +
      '<p>We have received your payment for <strong>' + businessName + '</strong>.</p>' +
      '<p><strong>Request ID:</strong> ' + requestId + '</p>' +
      '<p>Your BizOS workspace is now being prepared. You will receive your workspace access once provisioning is complete.</p>' +
      '<br><p>Thank you for choosing BizOS!</p></div>';
    MailApp.sendEmail({to: email, subject: subject, htmlBody: body});
  } catch (error) {
    console.error('Email error:', error);
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

    var latestRequest = getLatestUpgradeRequestForUser_(user.email);
    if (latestRequest && latestRequest.status === 'pending_payment' && latestRequest.expired) {
      markUpgradeRequestExpired_(latestRequest.requestId);
      latestRequest.status = 'expired';
    }
    if (latestRequest && latestRequest.status === 'pending_payment' && !latestRequest.expired) {
      var pendingPublicUrl = getPublicBizOSUrl_();
      var pendingPaymentPageUrl = pendingPublicUrl + '/?page=payment-method&requestId=' + encodeURIComponent(latestRequest.requestId) +
        '&country=' + encodeURIComponent(latestRequest.country || country) + '&tier=' + encodeURIComponent(latestRequest.tier || tier);
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
    var pricing = getLocalizedPricing(country);
    var amount = pricing.sovereign && pricing.sovereign.price || 499;
    var currency = pricing.sovereign && pricing.sovereign.code || 'USD';
    var symbol = pricing.sovereign && pricing.sovereign.currency || '$';

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var upgradeSheet = ss.getSheetByName('Upgrade_Requests');
    if (!upgradeSheet) {
      upgradeSheet = ss.insertSheet('Upgrade_Requests');
      upgradeSheet.appendRow([
        'Request_ID','Email','Workspace_Email','Name','Business_Name','Country','Business_Details','Tier','Payment_ID',
        'Amount','Currency','Status','Certificate_URL','Certificate_Name','Created_At','Updated_At'
      ]);
    }

    upgradeSheet.appendRow([
      requestId, user.email, workspaceEmail, user.name || user.email, user.businessName || '', country,
      JSON.stringify(businessDetails), tier, paymentId, amount, currency, 'pending_payment',
      certificateUrl, uploadedCertificateName, timestamp.toISOString(), ''
    ]);

    var paymentSheet = ss.getSheetByName('Payments');
    if (!paymentSheet) {
      paymentSheet = ss.insertSheet('Payments');
      paymentSheet.appendRow([
        'Payment_ID','Request_ID','Email','Workspace_Email','Amount','Currency','Status','Transaction_Ref','Created_At','Completed_At'
      ]);
    }
    paymentSheet.appendRow([
      paymentId, requestId, user.email, workspaceEmail, amount, currency, 'pending', '', timestamp.toISOString(), ''
    ]);

    if (upgradeLock) {
      upgradeLock.releaseLock();
      upgradeLock = null;
    }

    sendUpgradeRequestEmail(user.email, user.name, requestId, country, tier, amount, symbol);
    sendAdminUpgradeNotification(user.email, user.name, requestId, country, businessDetails, certificateUrl);

    // Keep the payment flow on the Apps Script deployment that owns the
    // Upgrade_Requests sheet. This preserves requestId and guarantees the
    // payment page reads the same backend/workspace that created the request.
    var publicUrl = getPublicBizOSUrl_();
    var paymentPageUrl = publicUrl + '/?page=payment-method&requestId=' + encodeURIComponent(requestId) +
      '&country=' + encodeURIComponent(country) + '&tier=' + encodeURIComponent(tier);

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
    return {success:false, message:error.message || 'An error occurred. Please try again.'};
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
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Upgrade_Requests');
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
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Upgrade_Requests');
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
      createdAt:pending.createdAt,
      status:pending.status,
      redirectUrl: getPublicBizOSUrl_() + '/?page=payment-method&requestId=' + encodeURIComponent(pending.requestId) +
        '&country=' + encodeURIComponent(pending.country || '') + '&tier=' + encodeURIComponent(pending.tier || 'sovereign')
    };
  } catch (error) {
    console.error('getPendingUpgradeRequestStatus error:', error);
    return {success:false, message:error.message || 'Unable to check upgrade status.'};
  }
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
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Upgrade_Requests');
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
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
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

function getLocalizedPricing(country) {
  var paymentConfig = PAYMENT_COUNTRIES[country] || PAYMENT_COUNTRIES.default;
  var basePrices = {sovereign:499, enterprise:2990};
  var rate = paymentConfig.multiplier || 1;
  var exchangeRate = 1500;
  return {
    sovereign:{price:Math.round(basePrices.sovereign * exchangeRate * rate) / 100, currency:paymentConfig.symbol || '$', code:paymentConfig.currency || 'USD'},
    enterprise:{price:Math.round(basePrices.enterprise * exchangeRate * rate) / 100, currency:paymentConfig.symbol || '$', code:paymentConfig.currency || 'USD'}
  };
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
