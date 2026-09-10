// ============================================================
// ZZZ_RegistrationEmailFix.js
// Registration + email reliability hardening
// ============================================================
// This file intentionally uses final global assignments so the
// authoritative implementations replace the older broken versions.

sendBrandedEmail = function(to, subject, title, content, buttonText, buttonUrl) {
  if (!to) return false;

  try {
    var emailConfig = (typeof CONFIG !== 'undefined' && CONFIG.EMAIL) ? CONFIG.EMAIL : {};
    var supportEmail = emailConfig.supportEmail || (CONFIG.APP && CONFIG.APP.supportEmail) || '';
    var fromName = emailConfig.fromName || 'BizOS';

    var htmlBody = '' +
      '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;padding:20px;">' +
        '<div style="background:#5D2A86;padding:20px;text-align:center;border-radius:12px 12px 0 0;">' +
          '<h1 style="color:white;margin:0;">BizOS</h1>' +
          '<p style="color:#e9d5ff;margin:5px 0 0;">Business Operating System</p>' +
        '</div>' +
        '<div style="background:white;padding:30px;border:1px solid #e5e7eb;">' +
          '<h2 style="color:#333;margin-top:0;">' + String(title || '') + '</h2>' +
          String(content || '') +
          (buttonText && buttonUrl ?
            '<div style="text-align:center;margin:30px 0;"><a href="' + String(buttonUrl) + '" style="background:#5D2A86;color:white;padding:12px 24px;text-decoration:none;border-radius:8px;display:inline-block;">' + String(buttonText) + ' →</a></div>' : '') +
          '<hr style="border:none;border-top:1px solid #e5e7eb;margin:20px 0;">' +
          (supportEmail ? '<p style="font-size:12px;color:#666;margin:0;">Need help? <a href="mailto:' + String(supportEmail) + '" style="color:#5D2A86;">Contact support</a></p>' : '') +
        '</div>' +
        '<div style="background:#f9fafb;padding:12px;text-align:center;border-radius:0 0 12px 12px;">' +
          '<p style="font-size:11px;color:#6b7280;margin:0;">© ' + new Date().getFullYear() + ' BizOS. All rights reserved.</p>' +
        '</div>' +
      '</div>';

    GmailApp.sendEmail(to, subject, 'Please view this message in an HTML-capable email client.', {
      htmlBody: htmlBody,
      name: fromName
    });

    console.log('Email sent to: ' + to);
    return true;
  } catch (error) {
    console.error('Email failed to ' + to + ': ' + error.message);
    return false;
  }
};

sendVerificationEmail = function(email, businessName, businessId, token) {
  var verificationUrl = getAppUrl() + '/verify?token=' + encodeURIComponent(token);
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

resendVerificationEmail = function(email) {
  try {
    email = String(email || '').trim().toLowerCase();
    if (!email) return { success:false, message:'Email address is required.' };

    var userSheet = getOrCreateUserSheet();
    var data = userSheet.getDataRange().getValues();
    var headers = data[0] || [];
    var emailCol = headers.indexOf('Email');
    var nameCol = headers.indexOf('Name');
    var businessNameCol = headers.indexOf('Business_Name');
    var businessIdCol = headers.indexOf('Business_ID');
    var verifiedCol = headers.indexOf('Is_Verified');

    var found = null;
    for (var i = 1; i < data.length; i++) {
      if (emailCol !== -1 && String(data[i][emailCol] || '').toLowerCase() === email) {
        found = { row:i + 1, data:data[i] };
        break;
      }
    }

    if (!found) return { success:false, message:'Email not found. Please sign up first.' };
    if (verifiedCol !== -1 && String(found.data[verifiedCol]).toUpperCase() === 'YES') {
      return { success:false, message:'This email is already verified. Please sign in.' };
    }

    var businessName = businessNameCol !== -1 ? String(found.data[businessNameCol] || '') : '';
    var businessId = businessIdCol !== -1 ? String(found.data[businessIdCol] || '') : '';
    var token = Utilities.getUuid();
    var now = new Date();
    var expires = new Date(now.getTime() + 24 * 60 * 60 * 1000);

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var verificationSheet = spreadsheet.getSheetByName('Email_Verifications');
    if (!verificationSheet) {
      verificationSheet = spreadsheet.insertSheet('Email_Verifications');
      verificationSheet.appendRow(['Email','Token','Business_Name','Business_Id','Created_At','Expires_At','Is_Used']);
    }

    verificationSheet.appendRow([
      email, token, businessName, businessId,
      now.toISOString(), expires.toISOString(), 'NO'
    ]);

    var sent = sendVerificationEmail(email, businessName, businessId, token);
    if (!sent) {
      var lastRow = verificationSheet.getLastRow();
      if (lastRow > 1) verificationSheet.deleteRow(lastRow);
      return { success:false, message:'We could not send the verification email. Please try again.' };
    }

    return { success:true, message:'Verification email sent. Please check your inbox and spam folder.' };
  } catch (error) {
    console.error('Resend verification error:', error);
    return { success:false, message:'Failed to send verification email. Please try again.' };
  }
};

registerBusiness = function(businessData) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  var businessSheet = null;
  var userSheet = null;
  var verificationSheet = null;
  var businessRow = -1;
  var userRow = -1;
  var verificationRow = -1;
  var verificationSheetCreated = false;

  try {
    businessData = businessData || {};

    var businessName = String(businessData.businessName || '').trim();
    var ownerEmail = String(businessData.ownerEmail || '').trim().toLowerCase();
    var password = String(businessData.password || '');

    if (!businessName || !ownerEmail || !password) {
      return { success:false, message:'Missing required fields' };
    }

    if (!/^\S+@\S+\.\S+$/.test(ownerEmail)) {
      return { success:false, message:'Please enter a valid email address.' };
    }

    if (password.length < 6) {
      return { success:false, message:'Password must be at least 6 characters.' };
    }

    businessSheet = getOrCreateBusinessSheet();
    userSheet = getOrCreateUserSheet();

    var businessDataRange = businessSheet.getDataRange().getValues();
    var businessHeaders = businessDataRange[0] || [];
    var businessNameCol = businessHeaders.indexOf('Business_Name');

    var existingBusiness = businessDataRange.find(function(row) {
      return businessNameCol !== -1 && row[businessNameCol] && String(row[businessNameCol]).trim().toLowerCase() === businessName.toLowerCase();
    });
    if (existingBusiness) {
      return { success:false, message:'Business name already registered' };
    }

    var userDataRange = userSheet.getDataRange().getValues();
    var userHeaders = userDataRange[0] || [];
    var emailCol = userHeaders.indexOf('Email');
    var existingUser = userDataRange.find(function(row) {
      return emailCol !== -1 && row[emailCol] && String(row[emailCol]).trim().toLowerCase() === ownerEmail;
    });
    if (existingUser) {
      return { success:false, message:'Email already registered. Please sign in instead.' };
    }

    var businessId = generateBusinessId();
    var hashedPassword = hashPassword(password);
    var ownerName = String(businessData.ownerName || businessName || 'Business Owner').trim();
    var ownerPhone = String(businessData.ownerPhone || '').trim();
    var address = String(businessData.address || '').trim();
    var city = String(businessData.city || '').trim();
    var country = String(businessData.country || 'Nigeria').trim();
    var website = String(businessData.website || '').trim();
    var now = new Date().toISOString();

    var newBusiness = [
      businessId, businessName, '', '', 'pending', now, 'auto',
      ownerEmail, ownerName, ownerPhone, address, city, country, website,
      'free', 'active', '', now, '', 'YES'
    ];
    businessSheet.appendRow(newBusiness);
    businessRow = businessSheet.getLastRow();

    var newUser = [
      ownerEmail, hashedPassword, ownerName, ownerPhone, 'owner',
      businessId, businessName, 'NO', '', 'pending', '', now, '', 'YES', 'NO'
    ];
    userSheet.appendRow(newUser);
    userRow = userSheet.getLastRow();

    var token = Utilities.getUuid();
    var expiresAt = new Date(Date.now() + 24 * 60 * 60 * 1000);
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    verificationSheet = spreadsheet.getSheetByName('Email_Verifications');
    if (!verificationSheet) {
      verificationSheet = spreadsheet.insertSheet('Email_Verifications');
      verificationSheetCreated = true;
      verificationSheet.appendRow(['Email','Token','Business_Name','Business_Id','Created_At','Expires_At','Is_Used']);
    }

    verificationSheet.appendRow([
      ownerEmail, token, businessName, businessId,
      now, expiresAt.toISOString(), 'NO'
    ]);
    verificationRow = verificationSheet.getLastRow();

    var emailSent = sendVerificationEmail(ownerEmail, businessName, businessId, token);
    if (!emailSent) {
      if (verificationRow > 1) verificationSheet.deleteRow(verificationRow);
      if (userRow > 1) userSheet.deleteRow(userRow);
      if (businessRow > 1) businessSheet.deleteRow(businessRow);

      return {
        success:false,
        code:'VERIFICATION_EMAIL_FAILED',
        message:'We could not send the verification email. Your account was not created. Please try again.'
      };
    }

    return {
      success:true,
      message:'Account created! A verification email has been sent to ' + ownerEmail + '. Please check your inbox and click the verification link to activate your account.',
      businessId:businessId,
      requiresVerification:true
    };
  } catch (error) {
    console.error('Business registration error:', error);

    try {
      if (verificationRow > 1 && verificationSheet) verificationSheet.deleteRow(verificationRow);
      if (userRow > 1 && userSheet) userSheet.deleteRow(userRow);
      if (businessRow > 1 && businessSheet) businessSheet.deleteRow(businessRow);
    } catch (rollbackError) {
      console.error('Registration rollback error:', rollbackError);
    }

    return { success:false, message:'Registration failed. Please try again.' };
  } finally {
    lock.releaseLock();
  }
};

console.log('ZZZ Registration/Email hardening loaded');
