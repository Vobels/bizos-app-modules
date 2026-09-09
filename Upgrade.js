//upgrade.gs
// ==================== UPGRADE.GS - COMPLETE WORKING VERSION ====================
// Paystack credentials are stored in Apps Script Project Settings > Script Properties.

function getLegacyPaystackSecretKey_() {
  var key = PropertiesService.getScriptProperties().getProperty('PAYSTACK_SECRET_KEY');
  if (!key) throw new Error('PAYSTACK_SECRET_KEY is not configured in Script Properties.');
  return String(key).trim();
}

function verifyPaystackPayment(reference, requestId) {
  // Keep the legacy function name for compatibility, but route it through
  // the secure production verification + provisioning flow.
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
 */
function handleUpgradeRequest(upgradeData, sessionId) {
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

    if (!country) return {success:false, message:'Please select your business country'};

    var countryConfig = getCountryRequirements(country);
    if (countryConfig.success && countryConfig.requiredFields) {
      for (var i = 0; i < countryConfig.requiredFields.length; i++) {
        var field = countryConfig.requiredFields[i];
        if (!businessDetails[field]) {
          return {success:false, message:'Missing required field: ' + ((countryConfig.fieldLabels && countryConfig.fieldLabels[field]) || field)};
        }
      }
    }

    var certificateUrl = '';
    var uploadedCertificateName = '';
    if (certificateBase64 && certificateName) {
      var uploadResult = uploadBusinessCertificate(
        certificateBase64, certificateName, country, user.businessId, user.businessName
      );
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
        'Request_ID','Email','Name','Business_Name','Country','Business_Details','Tier','Payment_ID',
        'Amount','Currency','Status','Certificate_URL','Certificate_Name','Created_At','Updated_At'
      ]);
    }

    upgradeSheet.appendRow([
      requestId, user.email, user.name || user.email, user.businessName || '', country,
      JSON.stringify(businessDetails), tier, paymentId, amount, currency, 'pending_payment',
      certificateUrl, uploadedCertificateName, timestamp.toISOString(), ''
    ]);

    var paymentSheet = ss.getSheetByName('Payments');
    if (!paymentSheet) {
      paymentSheet = ss.insertSheet('Payments');
      paymentSheet.appendRow([
        'Payment_ID','Request_ID','Email','Amount','Currency','Status','Transaction_Ref','Created_At','Completed_At'
      ]);
    }
    paymentSheet.appendRow([
      paymentId, requestId, user.email, amount, currency, 'pending', '', timestamp.toISOString(), ''
    ]);

    sendUpgradeRequestEmail(user.email, user.name, requestId, country, tier, amount, symbol);
    sendAdminUpgradeNotification(user.email, user.name, requestId, country, businessDetails, certificateUrl);

    var baseUrl = ScriptApp.getService().getUrl();
    var paymentPageUrl = baseUrl + '?page=paystack-int&requestId=' + requestId +
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
    var requestIdCol = headers.indexOf('Request_ID');
    if (requestIdCol === -1) return null;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][requestIdCol]) === String(requestId)) {
        return {
          requestId:data[i][0], email:data[i][1], name:data[i][2], businessName:data[i][3],
          country:data[i][4], businessDetails:data[i][5] ? JSON.parse(data[i][5]) : {}, tier:data[i][6],
          paymentId:data[i][7], amount:data[i][8], currency:data[i][9], status:data[i][10],
          certificateUrl:data[i][11], certificateName:data[i][12], createdAt:data[i][13]
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
