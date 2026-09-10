// ==================== EMAIL.GS - ALL EMAIL FUNCTIONS ====================

// ==================== CONFIGURATION ====================
// Use your custom domain for all email links
const APP_BASE_URL = 'https://bizos.higroups.com';  // Your custom domain

function getAppUrl() {
  // Return your custom domain for professional branding
  return APP_BASE_URL;
}

// For sovereign clients who have their own deployment
function getAppUrlForBusiness(businessId) {
  try {
    const businessSheet = getOrCreateBusinessSheet();
    const data = businessSheet.getDataRange().getValues();
    const headers = data[0];
    const bizIdCol = headers.indexOf('Business_ID');
    const customUrlCol = headers.indexOf('Custom_URL');
    if (customUrlCol !== -1) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][bizIdCol] === businessId && data[i][customUrlCol]) return data[i][customUrlCol];
      }
    }
    return APP_BASE_URL;
  } catch (error) {
    return APP_BASE_URL;
  }
}

// ==================== UNIVERSAL BRANDED EMAIL TEMPLATE ====================
function sendBrandedEmail(to, subject, title, content, buttonText, buttonUrl) {
  if (!to) {
    console.log("No email address provided");
    return false;
  }
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
      <div style="background: #5D2A86; padding: 20px; text-align: center; border-radius: 12px 12px 0 0;">
        <h1 style="color: white; margin: 0;">BizOS</h1>
        <p style="color: #e9d5ff; margin: 5px 0 0;">Business Operating System</p>
      </div>
      <div style="background: white; padding: 30px; border-radius: 0 0 12px 12px; border: 1px solid #e5e7eb;">
        <h2 style="color: #333; margin-top: 0;">${title}</h2>
        ${content}
        ${buttonText && buttonUrl ? `
          <div style="text-align: center; margin: 30px 0;">
            <a href="${buttonUrl}" style="background: #5D2A86; color: white; padding: 12px 24px; text-decoration: none; border-radius: 8px; display: inline-block;">${buttonText} →</a>
          </div>` : ''}
        <hr style="border: none; border-top: 1px solid #e5e7eb; margin: 20px 0;">
        <p style="font-size: 12px; color: #666; margin: 0;">Need help? <a href="mailto:${CONFIG.EMAIL.supportEmail}" style="color: #5D2A86;">Contact support</a></p>
      </div>
      <div style="background: #f9fafb; padding: 12px; text-align: center; border-radius: 0 0 12px 12px;">
        <p style="font-size: 11px; color: #6b7280; margin: 0;">© ${new Date().getFullYear()} BizOS. All rights reserved.</p>
      </div>
    </div>`;
  try {
    GmailApp.sendEmail(to, subject, "", { htmlBody: htmlBody, name: "BizOS" });
    console.log("Email sent to: " + to);
    return true;
  } catch (error) {
    console.log("Email failed to " + to + ": " + error.message);
    return false;
  }
}

// ==================== VERIFICATION EMAILS ====================
function sendVerificationEmail(email, businessName, businessId, token) {
  const verificationUrl = `${getAppUrl()}/verify?token=${encodeURIComponent(token)}`;
  const content = `
    <p>Welcome to BizOS, <strong>${businessName}</strong>!</p>
    <p>Please verify your email address to activate your account.</p>
    <p>This link expires in <strong>24 hours</strong>.</p>
    <p style="background:#f9fafb;padding:12px;border-radius:8px;color:#4b5563;font-size:13px;">If you don't see this email in your inbox, please check your <strong>Spam</strong> or <strong>Junk</strong> folder and search for <strong>BizOS</strong>.</p>`;
  return sendBrandedEmail(email, `Verify your email - ${businessName}`, 'Verify Your Email', content, 'Verify Email', verificationUrl);
}

function resendVerificationEmail(email) {
  try {
    email = String(email || '').trim().toLowerCase();
    if (!email) return { success: false, message: "Please enter your email address." };
    const userSheet = getOrCreateUserSheet();
    const userData = userSheet.getDataRange().getValues();
    const userHeaders = userData[0];
    const emailCol = userHeaders.indexOf('Email');
    const businessNameCol = userHeaders.indexOf('Business_Name');
    const verifiedCol = userHeaders.indexOf('Is_Verified');
    let userFound = false;
    let businessName = '';
    for (let i = 1; i < userData.length; i++) {
      if (userData[i][emailCol] && String(userData[i][emailCol]).toLowerCase() === email) {
        userFound = true;
        businessName = userData[i][businessNameCol];
        if (String(userData[i][verifiedCol] || '').toUpperCase() === 'YES') {
          return { success: false, message: "This email is already verified. Please login." };
        }
        break;
      }
    }
    if (!userFound) return { success: false, message: "Email not found. Please sign up first." };

    const token = Utilities.getUuid();
    const expiresAt = new Date();
    expiresAt.setHours(expiresAt.getHours() + 24);
    let verificationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email_Verifications');
    if (!verificationSheet) {
      verificationSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Email_Verifications');
      verificationSheet.appendRow(['Email', 'Token', 'Business_Name', 'Business_Id', 'Created_At', 'Expires_At', 'Is_Used']);
    }
    const verificationData = verificationSheet.getDataRange().getValues();
    const verificationHeaders = verificationData[0];
    const vEmailCol = verificationHeaders.indexOf('Email');
    const vUsedCol = verificationHeaders.indexOf('Is_Used');
    if (vEmailCol !== -1 && vUsedCol !== -1) {
      for (let i = 1; i < verificationData.length; i++) {
        if (String(verificationData[i][vEmailCol] || '').toLowerCase() === email && String(verificationData[i][vUsedCol] || '').toUpperCase() !== 'YES') {
          verificationSheet.getRange(i + 1, vUsedCol + 1).setValue('YES');
        }
      }
    }
    const sent = sendVerificationEmail(email, businessName, '', token);
    if (!sent) return { success: false, message: "We couldn't send the verification email. Please try again." };
    verificationSheet.appendRow([email, token, businessName, '', new Date().toISOString(), expiresAt.toISOString(), 'NO']);
    return { success: true, message: "Verification email sent. Please check your inbox and your Spam/Junk folder." };
  } catch (error) {
    console.error('Resend error:', error);
    return { success: false, message: "Failed to send verification email. Please try again." };
  }
}

// ==================== WELCOME EMAILS ====================
function sendWelcomeEmail(email, businessName) {
  const content = `<p>Welcome to BizOS, <strong>${businessName}</strong>!</p><p>Your business has been verified and your BizOS workspace is ready!</p><p>You can now:</p><ul><li>Access your business dashboard</li><li>Invite team members</li><li>Start using all available modules</li><li>Track your business performance</li></ul>`;
  return sendBrandedEmail(email, `Welcome to BizOS - ${businessName}`, 'Welcome to BizOS!', content, 'Go to Dashboard', getAppUrl());
}
function sendWelcomeEmailFree(email, businessName) {
  const content = `<p>Welcome to BizOS, <strong>${businessName}</strong>!</p><p>Your free account has been created successfully. You can now:</p><ul><li>Track your finances</li><li>Manage your e-commerce store (up to 10 products, 50 orders/month)</li></ul><p style="background: #f3f4f6; padding: 10px; border-radius: 8px;"><strong>💡 Tip:</strong> Upgrade to Sovereign to unlock all 8 business modules!</p>`;
  return sendBrandedEmail(email, `Welcome to BizOS - ${businessName}`, 'Welcome to BizOS!', content, 'Go to Dashboard', getAppUrl());
}

// ==================== INVITATION EMAILS ====================
function sendInvitationEmail(email, name, invitationCode, role) {
  const acceptUrl = `${getAppUrl()}?invite=${invitationCode}`;
  const content = `<p>Hello <strong>${name}</strong>,</p><p>You have been invited to join a business on BizOS as a <strong>${role}</strong>.</p><p>Your invitation code: <strong>${invitationCode}</strong></p><p>This invitation will expire in <strong>7 days</strong>.</p>`;
  return sendBrandedEmail(email, `You've been invited to join BizOS`, 'Invitation to Join', content, 'Accept Invitation', acceptUrl);
}

// ==================== PASSWORD RESET EMAILS ====================
function sendPasswordResetEmail(email, resetToken) {
  const resetUrl = `${getAppUrl()}?reset=${resetToken}`;
  const content = `<p>We received a request to reset your password.</p><p>This link expires in <strong>1 hour</strong>.</p><p style="background: #fef2f2; padding: 10px; border-radius: 8px; color: #991b1b;"><strong>⚠️ Security Notice:</strong> If you didn't request this, please ignore this email.</p>`;
  return sendBrandedEmail(email, 'BizOS - Password Reset Request', 'Reset Your Password', content, 'Reset Password', resetUrl);
}

// ==================== UPGRADE & VERIFICATION EMAILS ====================
function sendUpgradeRequestConfirmation(email, businessName, country, requestId) {
  const content = `<p>Thank you for upgrading to <strong>Sovereign Tier</strong>!</p><p>Your request for <strong>${country}</strong> has been received.</p><p><strong>Request ID:</strong> ${requestId}</p><p>Our team will review your documents within <strong>24-48 hours</strong>.</p><p>You can continue using your free account during verification.</p>`;
  return sendBrandedEmail(email, `Upgrade Request Received - ${businessName}`, 'Upgrade Request Received', content, 'Return to Dashboard', getAppUrl());
}
function sendUpgradeApprovalEmail(email, businessName, adminEmail) {
  const content = `<p>Your business <strong>${businessName}</strong> has been upgraded to <strong>Sovereign Tier</strong>!</p><p>You now have access to:</p><ul><li>All 8 business modules (Finance, Sales, E-commerce, CRM, HR, Logistics, Tax, Agro)</li><li>Unlimited team members</li><li>Priority support</li><li>Isolated secure workspace</li></ul><hr><p style="font-size: 12px; color: #666;">Approved by: ${adminEmail || 'BizOS Admin'}</p>`;
  return sendBrandedEmail(email, `${businessName} - Upgraded to Sovereign`, 'Congratulations! 🎉', content, 'Access Your Dashboard', getAppUrl());
}
function sendUpgradeRequestEmail(email, businessName, targetTier, amount, paymentPlan, requestId) {
  const content = `<p>Your upgrade request to <strong>${targetTier.toUpperCase()}</strong> has been received.</p><p><strong>Business:</strong> ${businessName}</p><p><strong>Amount:</strong> $${amount} (${paymentPlan})</p><p><strong>Request ID:</strong> ${requestId}</p><p>Our team will contact you shortly with payment instructions.</p>`;
  sendBrandedEmail(email, `Upgrade Request - ${businessName} to ${targetTier}`, 'Upgrade Request Received', content, 'Return to Dashboard', getAppUrl());
  const adminEmail = Session.getActiveUser().getEmail() || "bizoshigroups@gmail.com";
  const adminContent = `<h3>New Upgrade Request</h3><p><strong>Business:</strong> ${businessName}</p><p><strong>Tier:</strong> ${targetTier}</p><p><strong>Amount:</strong> $${amount}</p><p><strong>Request ID:</strong> ${requestId}</p>`;
  sendBrandedEmail(adminEmail, `New Upgrade Request - ${businessName}`, 'New Upgrade Request', adminContent, 'Review Request', getAppUrl());
}
function sendTierUpgradeConfirmation(businessId, tier) {
  const businessSheet = getOrCreateBusinessSheet();
  const data = businessSheet.getDataRange().getValues();
  const headers = data[0];
  const bizIdCol = headers.indexOf('Business_ID');
  const ownerEmailCol = headers.indexOf('Owner_Email');
  const businessNameCol = headers.indexOf('Business_Name');
  const row = data.find(row => row[bizIdCol] === businessId);
  if (!row) return;
  const email = row[ownerEmailCol];
  const businessName = row[businessNameCol];
  const content = `<p>Your business <strong>${businessName}</strong> has been upgraded to <strong>${tier.toUpperCase()}</strong> tier!</p><p>Thank you for choosing BizOS!</p>`;
  return sendBrandedEmail(email, `${businessName} - Upgraded to ${tier}`, 'Congratulations on Your Upgrade! 🎉', content, 'Access Your Dashboard', getAppUrl());
}

// ==================== ADMIN NOTIFICATION EMAILS ====================
function notifyAdminUpgradeRequest(user, upgradeData, countryConfig, requestId) {
  const adminEmail = Session.getActiveUser().getEmail() || "bizoshigroups@gmail.com";
  let verificationDetails = '';
  for (let field of countryConfig.requiredFields) verificationDetails += `<li><strong>${countryConfig.fieldLabels[field]}:</strong> ${upgradeData[field]}</li>`;
  let documentsList = '';
  for (let doc of countryConfig.documentTypes) documentsList += `<li>${doc}</li>`;
  const content = `<h3>New Upgrade Request</h3><p><strong>Business:</strong> ${user.businessName}</p><p><strong>Country:</strong> ${upgradeData.country}</p><p><strong>Request ID:</strong> ${requestId}</p><h4>Verification Details:</h4><ul>${verificationDetails}</ul><h4>Required Documents:</h4><ul>${documentsList}</ul><p><strong>Action Required:</strong> Please verify this business to complete the upgrade.</p>`;
  return sendBrandedEmail(adminEmail, `Upgrade Request - ${user.businessName} (${upgradeData.country})`, 'New Upgrade Request - Action Required', content, 'Review Request', getAppUrl());
}
function notifyAdminPaymentReceived(paymentId) {
  const adminEmail = Session.getActiveUser().getEmail() || "bizoshigroups@gmail.com";
  const content = `<p><strong>Payment ID:</strong> ${paymentId}</p><p><strong>Status:</strong> Completed</p><p><strong>Time:</strong> ${new Date().toLocaleString()}</p><p>Please verify the upgrade request to grant access.</p>`;
  return sendBrandedEmail(adminEmail, `Payment Received - ${paymentId}`, 'Payment Confirmation', content, 'Review Request', getAppUrl());
}

// ==================== ENTERPRISE INQUIRY EMAILS ====================
function sendEnterpriseInquiry(name, email, company, message) {
  try {
    const adminContent = `<h3>New Enterprise Request</h3><table style="width: 100%;"><tr><td style="padding: 4px 0;"><strong>Name:</strong> ${escapeHtml(name)}</td></tr><tr><td style="padding: 4px 0;"><strong>Email:</strong> <a href="mailto:${email}">${escapeHtml(email)}</a></td></tr><tr><td style="padding: 4px 0;"><strong>Company:</strong> ${escapeHtml(company)}</td></tr><tr><td style="padding: 4px 0;"><strong>Message:</strong> ${escapeHtml(message) || 'No message'}</td></tr><tr><td style="padding: 4px 0;"><strong>Received:</strong> ${new Date().toLocaleString()}</td></tr></table><div style="margin-top: 20px;"><a href="mailto:${email}" style="background: #5D2A86; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Reply to ${escapeHtml(name)}</a></div>`;
    sendBrandedEmail('bizoshigroups@gmail.com', `New Enterprise Inquiry - ${company}`, 'Enterprise Request', adminContent);
    const customerContent = `<p>Thank you for your interest in BizOS Enterprise, <strong>${escapeHtml(name)}</strong>!</p><p>We have received your inquiry and will get back to you within <strong>24 hours</strong>.</p><div style="background: #f9fafb; padding: 15px; border-radius: 8px; margin: 20px 0;"><p><strong>Your request summary:</strong></p><ul><li>Company: ${escapeHtml(company)}</li><li>Email: ${escapeHtml(email)}</li></ul></div><p>In the meantime, you can explore our free plan to test the platform!</p>`;
    sendBrandedEmail(email, 'BizOS Enterprise - We received your inquiry', 'Thank You for Your Interest', customerContent, 'Explore BizOS', getAppUrl());
    return { success: true, message: "Inquiry sent successfully" };
  } catch (error) {
    console.error('Error sending enterprise inquiry:', error);
    return { success: false, message: error.message };
  }
}

function escapeHtml(text) {
  if (!text) return '';
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\"/g, '&quot;').replace(/'/g, '&#39;');
}