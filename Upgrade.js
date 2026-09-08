//upgrade.gs
// ==================== UPGRADE.GS - COMPLETE WORKING VERSION ====================
// Paystack Secret Key - Replace with your actual key
const PAYSTACK_SECRET_KEY = 'sk_test_d1c5a8b2e3f4g5h6i7j8k9l0m1n2o3p4'; // REPLACE THIS!

function verifyPaystackPayment(reference, requestId) {
  try {
    const url = 'https://api.paystack.co/transaction/verify/' + reference;
    const options = {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + PAYSTACK_SECRET_KEY,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    
    if (result.status && result.data.status === 'success') {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Upgrade_Requests");
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      
      const requestIdCol = headers.indexOf("Request_ID");
      const statusCol = headers.indexOf("Status");
      const paymentIdCol = headers.indexOf("Payment_ID");
      
      let rowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][requestIdCol] === requestId) {
          rowIndex = i;
          break;
        }
      }
      
      if (rowIndex !== -1) {
        sheet.getRange(rowIndex + 1, statusCol + 1).setValue("payment_confirmed");
        
        const email = data[rowIndex][1];
        const businessName = data[rowIndex][3];
        
        sendPaymentConfirmationEmail(email, businessName, requestId);
        
        return { success: true, message: "Payment verified successfully!" };
      }
    }
    
    return { success: false, message: "Payment verification failed" };
    
  } catch (error) {
    console.error("Paystack verification error:", error);
    return { success: false, message: error.message };
  }
}

function sendPaymentConfirmationEmail(email, businessName, requestId) {
  try {
    const subject = "BizOS - Payment Received for " + businessName;
    const body = `
      <div style="font-family: Arial, sans-serif;">
        <h2>Payment Received! 🎉</h2>
        <p>We have received your payment for <strong>${businessName}</strong>.</p>
        <p><strong>Request ID:</strong> ${requestId}</p>
        <p>Your account will be upgraded to Sovereign tier within 24 hours after verification.</p>
        <p>You will receive another email once your account is upgraded.</p>
        <br>
        <p>Thank you for choosing BizOS!</p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: body
    });
  } catch (error) {
    console.error("Email error:", error);
  }
}

function getBankDetails() {
  return {
    success: true,
    bankName: "Example Bank",
    accountName: "BizOS Global",
    accountNumber: "1234567890"
  };
}

/**
 * Handle upgrade request submission from frontend
 */
function handleUpgradeRequest(upgradeData, sessionId) {
  try {
    // ⭐ CHANGED: was validateSession(sessionId)
    const user = validateUpgradeSession(sessionId);
    if (!user) return { success: false, message: "Session expired. Please login again." };
    
    if (user.subscriptionTier !== 'free' && user.subscriptionTier !== 'starter') {
      return { success: false, message: "You are already on a higher tier!" };
    }
    
    const { country, originalCountry, businessDetails, tier = "sovereign", certificateBase64, certificateName } = upgradeData;
    const finalCountry = country || originalCountry;
    
    if (!finalCountry) {
      return { success: false, message: "Please select your business country" };
    }
    
    const countryConfig = getCountryRequirements(finalCountry);
    if (countryConfig.success && countryConfig.requiredFields) {
      for (const field of countryConfig.requiredFields) {
        if (!businessDetails[field]) {
          return { success: false, message: `Missing required field: ${countryConfig.fieldLabels[field] || field}` };
        }
      }
    }
    
    let certificateUrl = "";
    let uploadedCertificateName = "";
    
    if (certificateBase64 && certificateName) {
      const uploadResult = uploadBusinessCertificate(
        certificateBase64,
        certificateName,
        finalCountry,
        user.businessId,
        user.businessName
      );
      
      if (uploadResult.success) {
        certificateUrl = uploadResult.fileUrl;
        uploadedCertificateName = uploadResult.fileName;
        console.log("Certificate uploaded:", certificateUrl);
      } else {
        console.error("Certificate upload failed:", uploadResult.message);
      }
    }
    
    const requestId = "UPG_" + Utilities.getUuid().substring(0, 8).toUpperCase();
    const paymentId = "PAY_" + Utilities.getUuid().substring(0, 8).toUpperCase();
    const timestamp = new Date();
    
    const pricing = getLocalizedPricing(finalCountry);
    const amount = pricing.sovereign?.price || 499;
    const currency = pricing.sovereign?.code || "USD";
    const symbol = pricing.sovereign?.currency || "$";
    
    let upgradeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Upgrade_Requests");
    if (!upgradeSheet) {
      upgradeSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Upgrade_Requests");
      upgradeSheet.appendRow([
        "Request_ID", "Email", "Name", "Business_Name", "Country", 
        "Business_Details", "Tier", "Payment_ID", "Amount", "Currency",
        "Status", "Certificate_URL", "Certificate_Name", "Created_At", "Updated_At"
      ]);
    }
    
    upgradeSheet.appendRow([
      requestId,
      user.email,
      user.name || user.email,
      user.businessName || "",
      finalCountry,
      JSON.stringify(businessDetails),
      tier,
      paymentId,
      amount,
      currency,
      "pending_payment",
      certificateUrl,
      uploadedCertificateName,
      timestamp.toISOString(),
      ""
    ]);
    
    let paymentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payments");
    if (!paymentSheet) {
      paymentSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Payments");
      paymentSheet.appendRow([
        "Payment_ID", "Request_ID", "Email", "Amount", "Currency", 
        "Status", "Transaction_Ref", "Created_At", "Completed_At"
      ]);
    }
    paymentSheet.appendRow([
      paymentId, requestId, user.email, amount, currency, 
      "pending", "", timestamp.toISOString(), ""
    ]);
    
    sendUpgradeRequestEmail(user.email, user.name, requestId, finalCountry, tier, amount, symbol);
    sendAdminUpgradeNotification(user.email, user.name, requestId, finalCountry, businessDetails, certificateUrl);
    
    const baseUrl = ScriptApp.getService().getUrl();
    const paymentPageUrl = baseUrl + "?page=paystack-int&requestId=" + requestId + "&country=" + encodeURIComponent(finalCountry) + "&tier=" + tier;
    
    return {
      success: true,
      message: "Upgrade request submitted! Please complete payment to activate your Sovereign tier.",
      requestId: requestId,
      paymentId: paymentId,
      amount: amount,
      currency: symbol,
      redirectUrl: paymentPageUrl
    };
    
  } catch (error) {
    console.error("Upgrade request error:", error);
    return { success: false, message: error.message || "An error occurred. Please try again." };
  }
}

/**
 * Send notification to admin
 */
function sendAdminUpgradeNotification(userEmail, userName, requestId, country, businessDetails, certificateUrl) {
  try {
    const subject = "New Sovereign Upgrade Request - " + requestId;
    
    const certificateHtml = certificateUrl ? 
      `<p><strong>Certificate:</strong> <a href="${certificateUrl}" target="_blank">View Uploaded Document</a></p>` : 
      '<p><strong>Certificate:</strong> Not uploaded</p>';
    
    const htmlBody = `
      <div style="font-family: Arial, sans-serif;">
        <h2>New Upgrade Request</h2>
        <p><strong>User:</strong> ${userName} (${userEmail})</p>
        <p><strong>Request ID:</strong> ${requestId}</p>
        <p><strong>Country:</strong> ${country}</p>
        ${certificateHtml}
        <p><strong>Business Details:</strong></p>
        <pre>${JSON.stringify(businessDetails, null, 2)}</pre>
        <p><strong>Status:</strong> Pending Payment</p>
        <p><strong>Time:</strong> ${new Date().toLocaleString()}</p>
        <hr>
        <p>Login to admin dashboard to verify payment when received.</p>
      </div>
    `;
    
    MailApp.sendEmail({
      to: "bizoshigroups@gmail.com",
      subject: subject,
      htmlBody: htmlBody
    });
    
  } catch (error) {
    console.error("Admin email error:", error);
  }
}

/**
 * Get upgrade request by ID
 */
function getUpgradeRequest(requestId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Upgrade_Requests");
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const requestIdCol = headers.indexOf("Request_ID");
    if (requestIdCol === -1) return null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][requestIdCol] === requestId) {
        return {
          requestId: data[i][0],
          email: data[i][1],
          name: data[i][2],
          businessName: data[i][3],
          country: data[i][4],
          businessDetails: data[i][5] ? JSON.parse(data[i][5]) : {},
          tier: data[i][6],
          paymentId: data[i][7],
          amount: data[i][8],
          currency: data[i][9],
          status: data[i][10],
          certificateUrl: data[i][11],
          certificateName: data[i][12],
          createdAt: data[i][13]
        };
      }
    }
    return null;
    
  } catch (error) {
    console.error("Error getting upgrade request:", error);
    return null;
  }
}

/**
 * Upload business certificate to Google Drive
 */
function uploadBusinessCertificate(base64Data, fileName, country, businessId, businessName) {
  try {
    if (!businessId || !businessName) {
      throw new Error("Missing businessId or businessName");
    }

    const cleanBase64 = base64Data.includes(",")
      ? base64Data.split(",")[1]
      : base64Data;

    const rootFolder = getBizOSFolder();

    const businessFolder = getOrCreateBusinessSubfolder(
      rootFolder,
      businessId,
      businessName
    );

    let certFolder;
    const existing = businessFolder.getFoldersByName("Certificates");

    if (existing.hasNext()) {
      certFolder = existing.next();
    } else {
      certFolder = businessFolder.createFolder("Certificates");
    }

    const blob = Utilities.newBlob(
      Utilities.base64Decode(cleanBase64),
      getMimeType(fileName),
      fileName
    );

    const file = certFolder.createFile(blob);

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      success: true,
      fileUrl: file.getUrl(),
      fileId: file.getId(),
      fileName: fileName
    };

  } catch (error) {
    console.error("Error uploading certificate:", error);
    return { success: false, message: error.message };
  }
}

/**
 * Get MIME type from file name
 */
function getMimeType(fileName) {
  const ext = fileName.split('.').pop().toLowerCase();
  const mimeTypes = {
    'pdf': 'application/pdf',
    'png': 'image/png',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg'
  };
  return mimeTypes[ext] || 'application/octet-stream';
}

/**
 * Get localized pricing for a country
 */
function getLocalizedPricing(country) {
  const paymentConfig = PAYMENT_COUNTRIES[country] || PAYMENT_COUNTRIES.default;
  
  const basePrices = {
    sovereign: 499,
    enterprise: 2990
  };
  
  const rate = paymentConfig.multiplier || 1;
  const exchangeRate = 1500;
  
  return {
    sovereign: {
      price: Math.round(basePrices.sovereign * exchangeRate * rate) / 100,
      currency: paymentConfig.symbol || "$",
      code: paymentConfig.currency || "USD"
    },
    enterprise: {
      price: Math.round(basePrices.enterprise * exchangeRate * rate) / 100,
      currency: paymentConfig.symbol || "$",
      code: paymentConfig.currency || "USD"
    }
  };
}

/**
 * ⭐ RENAMED from validateSession to validateUpgradeSession
 * to avoid colliding with Auth.gs's validateSession(sessionId),
 * which returns an email string rather than a user object.
 * Uses CacheService (same as Auth.gs) instead of a Sessions sheet.
 */
function validateUpgradeSession(sessionId) {
  try {
    if (!sessionId) {
      console.log("validateUpgradeSession: No sessionId provided");
      return null;
    }
    
    const cache = CacheService.getScriptCache();
    const sessionData = cache.get(sessionId);
    
    if (!sessionData) {
      console.log("validateUpgradeSession: No session data in cache for ID:", sessionId.substring(0, 20) + "...");
      return null;
    }
    
    const session = JSON.parse(sessionData);
    const email = session.email;
    
    if (!email) {
      console.log("validateUpgradeSession: No email in session data");
      return null;
    }
    
    console.log("validateUpgradeSession: Session valid for email:", email);
    
    const userSheet = getOrCreateUserSheet();
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    const emailCol = headers.indexOf('Email');
    const nameCol = headers.indexOf('Name');
    const roleCol = headers.indexOf('Role');
    const businessIdCol = headers.indexOf('Business_ID');
    const businessNameCol = headers.indexOf('Business_Name');
    const isDemoCol = headers.indexOf('Is_Demo');
    
    if (emailCol === -1) {
      console.error("validateUpgradeSession: Email column not found");
      return null;
    }
    
    let userRow = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol] && data[i][emailCol].toLowerCase() === email.toLowerCase()) {
        userRow = data[i];
        break;
      }
    }
    
    if (!userRow) {
      console.log("validateUpgradeSession: User not found for email:", email);
      return null;
    }
    
    const businessId = userRow[businessIdCol];
    let subscriptionTier = 'free';
    
    if (businessId) {
      const businessSheet = getOrCreateBusinessSheet();
      const businessData = businessSheet.getDataRange().getValues();
      const businessHeaders = businessData[0];
      const bizIdCol = businessHeaders.indexOf('Business_ID');
      const tierCol = businessHeaders.indexOf('Subscription_Tier');
      
      if (bizIdCol !== -1) {
        for (let i = 1; i < businessData.length; i++) {
          if (businessData[i][bizIdCol] === businessId) {
            subscriptionTier = tierCol !== -1 ? businessData[i][tierCol] || 'free' : 'free';
            break;
          }
        }
      }
    }
    
    const isDemo = userRow[isDemoCol] === 'YES';
    
    return {
      email: userRow[emailCol],
      name: userRow[nameCol] || email,
      role: userRow[roleCol] || 'owner',
      businessId: businessId,
      businessName: userRow[businessNameCol] || '',
      subscriptionTier: isDemo ? 'demo' : subscriptionTier,
      isDemo: isDemo
    };
    
  } catch (error) {
    console.error("validateUpgradeSession error:", error);
    return null;
  }
}

function markPaymentInitiated(requestId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Upgrade_Requests");
    if (!sheet) return { success: false, message: "Request not found" };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const requestIdCol = headers.indexOf("Request_ID");
    const statusCol = headers.indexOf("Status");
    
    if (requestIdCol === -1 || statusCol === -1) {
      return { success: false, message: "Invalid sheet structure" };
    }
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][requestIdCol] === requestId) {
        rowIndex = i;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: "Request not found" };
    }
    
    sheet.getRange(rowIndex + 1, statusCol + 1).setValue("payment_initiated");
    
    return { success: true, message: "Payment recorded. We'll verify and upgrade your account soon." };
    
  } catch (error) {
    console.error("Error marking payment:", error);
    return { success: false, message: error.message };
  }
}

