//business.gs 
//==================== BUSINESS REGISTRATION & VERIFICATION ====================

function registerBusiness(businessData) {
  try {
    console.log('Business registration attempt:', businessData.businessName);
    
    // Validate required fields
    if (!businessData.businessName || !businessData.ownerEmail || !businessData.password) {
      return { success: false, message: "Missing required fields" };
    }
    
    const businessSheet = getOrCreateBusinessSheet();
    const userSheet = getOrCreateUserSheet();
    
    // SAFE: Convert to string and handle null/undefined
    const businessName = String(businessData.businessName || '').trim();
    const ownerEmail = String(businessData.ownerEmail || '').trim();
    
    if (!businessName || !ownerEmail) {
      return { success: false, message: "Business name and email are required" };
    }
    
    // Check if business name exists
    const businessData_range = businessSheet.getDataRange().getValues();
    const businessHeaders = businessData_range[0];
    const businessNameCol = businessHeaders.indexOf('Business_Name');
    const existingBusiness = businessData_range.find(row => 
      row[businessNameCol] && String(row[businessNameCol]).toLowerCase() === businessName.toLowerCase()
    );
    
    if (existingBusiness) {
      return { success: false, message: "Business name already registered" };
    }
    
    // Check if email already exists
    const userData_range = userSheet.getDataRange().getValues();
    const userHeaders = userData_range[0];
    const emailCol = userHeaders.indexOf('Email');
    const existingUser = userData_range.find(row => 
      row[emailCol] && String(row[emailCol]).toLowerCase() === ownerEmail.toLowerCase()
    );
    
    if (existingUser) {
      return { success: false, message: "Email already registered. Please login instead." };
    }
    
    // Create business
    const businessId = generateBusinessId();
    const hashedPassword = hashPassword(businessData.password);
    
    // Get other fields with safe defaults
    const ownerName = String(businessData.ownerName || businessData.businessName || 'Business Owner').trim();
    const ownerPhone = String(businessData.ownerPhone || '').trim();
    const address = String(businessData.address || '').trim();
    const city = String(businessData.city || '').trim();
    const country = String(businessData.country || 'Nigeria').trim();
    const website = String(businessData.website || '').trim();
    
    const newBusiness = [
      businessId,
      businessName,
      '', // CAC Number
      '', // Tax ID
      'pending',
      new Date().toISOString(),
      'auto',
      ownerEmail.toLowerCase(),
      ownerName,
      ownerPhone,
      address,
      city,
      country,
      website,
      'free',
      'active',
      '', // workspaceId
      new Date().toISOString(),
      '',
      'YES'
    ];
    
    businessSheet.appendRow(newBusiness);
    
    // Create user
    const newUser = [
      ownerEmail.toLowerCase(),
      hashedPassword,
      ownerName,
      ownerPhone,
      'owner',
      businessId,
      businessName,
      'NO', // Is_Verified
      '', // Invited_By
      'pending', // Invitation_Status
      '', // Invitation_Code
      new Date().toISOString(),
      '',
      'YES',
      'NO'
    ];
    
    userSheet.appendRow(newUser);
    
    // Generate verification token
    const token = Utilities.getUuid();
    const expiresAt = new Date();
    expiresAt.setHours(expiresAt.getHours() + 24);

    // Store token
    let verificationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email_Verifications');
    if (!verificationSheet) {
      verificationSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Email_Verifications');
      verificationSheet.appendRow(['Email', 'Token', 'Business_Name', 'Business_Id', 'Created_At', 'Expires_At', 'Is_Used']);
    }

    verificationSheet.appendRow([
      ownerEmail.toLowerCase(),
      token,
      businessName,
      businessId,
      new Date().toISOString(),
      expiresAt.toISOString(),
      'NO'
    ]);

    // Send verification email
    sendVerificationEmail(ownerEmail, businessName, businessId, token);
    
    return { 
      success: true, 
      message: "Account created! A verification email has been sent to " + ownerEmail + ". Please check your inbox and click the verification link to activate your account.",
      businessId: businessId,
      requiresVerification: true
    };
    
  } catch (error) {
    console.error('Business registration error:', error);
    return { success: false, message: "Registration failed: " + error.message };
  }
}

function generateBusinessId() {
  return 'BIZ-' + Utilities.getUuid().substring(0, 8).toUpperCase();
}

function getOrCreateBusinessSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let businessSheet = spreadsheet.getSheetByName('Businesses');
  
  if (!businessSheet) {
    businessSheet = spreadsheet.insertSheet('Businesses');
    
    const headers = [
      'Business_ID', 'Business_Name', 'CAC_Number', 'Tax_ID', 
      'Verification_Status', 'Verification_Date', 'Verified_By',
      'Owner_Email', 'Owner_Name', 'Owner_Phone',
      'Address', 'City', 'Country', 'Website',
      'Subscription_Tier', 'Subscription_Status', 'Workspace_ID',
      'Created_At', 'Last_Active', 'Is_Active',
      'Assigned_Modules'
    ];
    businessSheet.appendRow(headers);
    
    const headerRange = businessSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f4f6');
    
    console.info('Businesses sheet created');
  }
  
  return businessSheet;
}

function verifyBusiness(businessId, adminEmail) {
  try {
    const businessSheet = getOrCreateBusinessSheet();
    const data = businessSheet.getDataRange().getValues();
    const headers = data[0];
    
    const businessIdCol = headers.indexOf('Business_ID');
    const statusCol = headers.indexOf('Verification_Status');
    const verificationDateCol = headers.indexOf('Verification_Date');
    const verifiedByCol = headers.indexOf('Verified_By');
    const workspaceIdCol = headers.indexOf('Workspace_ID');
    
    const rowIndex = data.findIndex(row => row[businessIdCol] === businessId);
    
    if (rowIndex === -1) {
      return { success: false, message: "Business not found" };
    }
    
    // Update verification status
    businessSheet.getRange(rowIndex + 1, statusCol + 1).setValue('verified');
    businessSheet.getRange(rowIndex + 1, verificationDateCol + 1).setValue(new Date().toISOString());
    businessSheet.getRange(rowIndex + 1, verifiedByCol + 1).setValue(adminEmail);
    
    // Create workspace for verified business
    const businessName = data[rowIndex][headers.indexOf('Business_Name')];
    const ownerEmail = data[rowIndex][headers.indexOf('Owner_Email')];
    
    const workspaceResult = createWorkspaceForBusiness(businessId, businessName, ownerEmail);
    
    if (workspaceResult.success) {
      businessSheet.getRange(rowIndex + 1, workspaceIdCol + 1).setValue(workspaceResult.workspaceId);
      
      // Update user verification status
      updateUserVerification(ownerEmail, businessId);
      
      // Send welcome email
      sendWelcomeEmail(ownerEmail, businessName);
    }
    
    return { 
      success: true, 
      message: "Business verified successfully! Workspace created." 
    };
    
  } catch (error) {
    console.error('Business verification error:', error);
    return { success: false, message: error.message };
  }
}

function getBusinessDetails(businessId, sessionId) { 
  try {
    console.log("getBusinessDetails called for businessId:", businessId);
    console.log("SessionId provided:", sessionId ? "YES" : "NO");
    
    // Verify session FIRST
    if (sessionId) {
      const user = getUserFromSession(sessionId);
      if (!user) {
        console.log("Invalid session");
        return { success: false, message: "Unauthorized - Invalid session" };
      }
      console.log("Session verified for user:", user.email);
    }
    
    if (!businessId) {
      console.log("No businessId provided");
      return { success: false, message: "No business ID provided" };
    }
    
    const businessSheet = getOrCreateBusinessSheet();
    const data = businessSheet.getDataRange().getValues();
    
    if (data.length < 2) {
      console.log("No data in Businesses sheet");
      return { success: false, message: "No business data found" };
    }
    
    const headers = data[0];
    console.log("Headers found:", headers);
    
    // Find column indices (case insensitive)
    const bizIdCol = headers.findIndex(h => h && h.toString().toLowerCase() === 'business_id');
    const nameCol = headers.findIndex(h => h && h.toString().toLowerCase() === 'business_name');
    const tierCol = headers.findIndex(h => h && h.toString().toLowerCase() === 'subscription_tier');
    const statusCol = headers.findIndex(h => h && h.toString().toLowerCase() === 'verification_status');
    const workspaceIdCol = headers.findIndex(h => h && h.toString().toLowerCase() === 'workspace_id');
    
    console.log("Column indices - ID:", bizIdCol, "Name:", nameCol, "Tier:", tierCol, "Status:", statusCol);
    
    // Find the business row
    let businessRow = null;
    let rowNumber = -1;
    for (let i = 1; i < data.length; i++) {
      const rowBusinessId = data[i][bizIdCol];
      if (rowBusinessId && rowBusinessId.toString() === businessId.toString()) {
        businessRow = data[i];
        rowNumber = i + 1;
        console.log("Found business at row", rowNumber);
        break;
      }
    }
    
    if (!businessRow) {
      console.log("Business not found for ID:", businessId);
      console.log("Available business IDs in sheet:");
      for (let i = 1; i < data.length; i++) {
        console.log(`  - ${data[i][bizIdCol]}`);
      }
      return { success: false, message: "Business not found" };
    }
    
    let workspaceUrl = '';
    if (workspaceIdCol !== -1 && businessRow[workspaceIdCol]) {
      try {
        const workspace = SpreadsheetApp.openById(businessRow[workspaceIdCol]);
        workspaceUrl = workspace.getUrl();
      } catch(e) {
        console.log("Could not open workspace:", e.message);
      }
    }
    
    const result = {
      success: true,
      businessId: businessId,
      businessName: nameCol !== -1 ? businessRow[nameCol] || '' : '',
      subscriptionTier: tierCol !== -1 ? businessRow[tierCol] || 'free' : 'free',
      verificationStatus: statusCol !== -1 ? businessRow[statusCol] || 'pending' : 'pending',
      workspaceUrl: workspaceUrl
    };
    
    console.log("Returning result:", result);
    return result;
    
  } catch(error) {
    console.error('getBusinessDetails error:', error);
    return { success: false, message: error.toString() };
  }
}

function getOrCreateUserSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let userSheet = spreadsheet.getSheetByName('Users');
  
  if (!userSheet) {
    userSheet = spreadsheet.insertSheet('Users');
    
    const headers = [
      'Email', 'Password_Hash', 'Name', 'Phone', 'Role',
      'Business_ID', 'Business_Name', 'Is_Verified',
      'Invited_By', 'Invitation_Status', 'Invitation_Code',
      'Created_At', 'Last_Login', 'Is_Active', 'Is_Demo'
    ];
    userSheet.appendRow(headers);
    
    const headerRange = userSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f4f6');
    
   
    // Create workspace for demo business
    const demoBusinessId = generateBusinessId();
    businessSheet = getOrCreateBusinessSheet();

    // Create workspace for demo business
    const demoWorkspace = createWorkspaceForBusiness(demoBusinessId, 'Demo Business', 'demo@bizos.com');
    const demoWorkspaceId = demoWorkspace.success ? demoWorkspace.workspaceId : '';

    businessSheet.appendRow([
      demoBusinessId, 'Demo Business', 'DEMO123456', 'DEMO789012',
      'verified', new Date().toISOString(), 'admin@bizos.com',
      'owner@demobiz.com', 'Demo Owner', '+1234567890',
      '123 Demo St', 'Demo City', 'Demo Country', 'https://demobiz.com',
      'free', 'active', demoWorkspaceId,
      new Date().toISOString(), '', 'YES'
    ]);
    userSheet.appendRow([
      'demo@bizos.com', '', 'Demo User', '+1234567890', 'staff',
      demoBusinessId, 'Demo Business', 'YES',
      'system', 'accepted', '',
      new Date().toISOString(), '', 'YES', 'YES'
    ]);
    
    userSheet.appendRow([
      'premium@bizos.com', '', 'Premium User', '+1234567890', 'owner',
      demoBusinessId, 'Demo Business', 'YES',
      'system', 'accepted', '',
      new Date().toISOString(), '', 'YES', 'YES'
    ]);
    
    console.info('Users sheet created with demo accounts');
  }
  
  return userSheet;
}

function generateInvitationCode() {
  return Math.random().toString(36).substring(2, 10).toUpperCase();
}

function verifyEmailToken(token) {
  try {
    const verificationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email_Verifications');
    if (!verificationSheet) {
      return { success: false, message: "Invalid verification token." };
    }
    
    const data = verificationSheet.getDataRange().getValues();
    const headers = data[0];
    
    const tokenCol = headers.indexOf('Token');
    const emailCol = headers.indexOf('Email');
    const businessNameCol = headers.indexOf('Business_Name');
    const businessIdCol = headers.indexOf('Business_Id');
    const expiresAtCol = headers.indexOf('Expires_At');
    const isUsedCol = headers.indexOf('Is_Used');
    
    let rowIndex = -1;
    let tokenData = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][tokenCol] === token) {
        rowIndex = i;
        tokenData = data[i];
        break;
      }
    }
    
    if (!tokenData) {
      return { success: false, message: "Invalid verification token." };
    }
    
    if (tokenData[isUsedCol] === 'YES') {
      return { success: false, message: "This verification link has already been used." };
    }
    
    const expiresAt = new Date(tokenData[expiresAtCol]);
    if (new Date() > expiresAt) {
      return { success: false, message: "Verification link has expired.", expired: true };
    }
    
    // Mark token as used
    verificationSheet.getRange(rowIndex + 1, isUsedCol + 1).setValue('YES');
    
    // Get token data
    const email = tokenData[emailCol];
    const businessName = tokenData[businessNameCol];
    const businessId = tokenData[businessIdCol];
    
    // Update user as verified
    const userSheet = getOrCreateUserSheet();
    const userData = userSheet.getDataRange().getValues();
    const userHeaders = userData[0];
    
    const userEmailCol = userHeaders.indexOf('Email');
    const verifiedCol = userHeaders.indexOf('Is_Verified');
    const userBusinessIdCol = userHeaders.indexOf('Business_ID');
    
    let userFound = false;
    
    for (let i = 1; i < userData.length; i++) {
      const userEmail = userData[i][userEmailCol];
      if (userEmail && userEmail.toLowerCase() === email.toLowerCase()) {
        userSheet.getRange(i + 1, verifiedCol + 1).setValue('YES');
        userFound = true;
        console.log(`Verified user: ${email}`);
        break;
      }
    }
    
    if (!userFound) {
      console.error(`User not found for email: ${email}`);
      return { success: false, message: "User not found. Please sign up again." };
    }
    
    // Get business sheet reference
    const businessSheet = getOrCreateBusinessSheet();
    const businessData = businessSheet.getDataRange().getValues();
    const businessHeaders = businessData[0];
    
    const bizIdCol = businessHeaders.indexOf('Business_ID');
    const bizStatusCol = businessHeaders.indexOf('Verification_Status');
    const workspaceIdCol = businessHeaders.indexOf('Workspace_ID');
    
    // Create workspace for the verified business
    console.log(`Creating workspace for ${businessName} (${businessId})`);
    const workspaceResult = createWorkspaceForBusiness(businessId, businessName, email);
    
    if (workspaceResult.success) {
      console.log(`Workspace created: ${workspaceResult.workspaceId}`);
      
      // Update business sheet with workspace ID and verification status
      for (let i = 1; i < businessData.length; i++) {
        if (businessData[i][bizIdCol] === businessId) {
          businessSheet.getRange(i + 1, workspaceIdCol + 1).setValue(workspaceResult.workspaceId);
          businessSheet.getRange(i + 1, bizStatusCol + 1).setValue('verified');
          console.log(`Updated business ${businessId} with workspace and verified status`);
          break;
        }
      }
    } else {
      console.error(`Failed to create workspace: ${workspaceResult.message}`);
      // Still mark as verified even if workspace creation fails
      for (let i = 1; i < businessData.length; i++) {
        if (businessData[i][bizIdCol] === businessId) {
          businessSheet.getRange(i + 1, bizStatusCol + 1).setValue('verified');
          break;
        }
      }
    }
    
    // Return success with business name for better UX
    return { 
      success: true, 
      message: "Email verified successfully! You can now login.",
      businessName: businessName,
      email: email,
      workspaceCreated: workspaceResult.success
    };
    
  } catch (error) {
    console.error("Verification error:", error);
    return { success: false, message: "Verification failed: " + error.message };
  }
}

function fixEmailVerificationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Email_Verifications');
  
  if (!sheet) {
    sheet = ss.insertSheet('Email_Verifications');
    sheet.appendRow(['Email', 'Token', 'Business_Name', 'Business_Id', 'Created_At', 'Expires_At', 'Is_Used']);
  } else {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes('Business_Id')) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Business_Id');
    }
  }
  
  console.log('Email_Verifications sheet is ready');
}

function sendBusinessVerificationNotice(email, businessName, businessId) {
  try {
    const subject = `BizOS - Business Verification Required: ${businessName}`;
    const body = `
      <h2>Thank you for registering ${businessName} with BizOS!</h2>
      <p>Your business registration has been received. Our team will verify your business documents within 24-48 hours.</p>
      <p><strong>Business ID:</strong> ${businessId}</p>
      <p>Once verified, you'll receive a welcome email with access to your workspace.</p>
      <p>If you have any questions, please contact our support team at ${EMAIL_CONFIG.supportEmail}</p>
      <br>
      <p>Best regards,<br>BizOS Team</p>
    `;
    
    GmailApp.sendEmail(email, subject, "", { htmlBody: body, name: "BizOS" });
    console.log('Business verification notice sent to:', email);
  } catch (error) {
    console.error('Email send error:', error);
  }
}

function updateUserVerification(email, businessId) {
  const userSheet = getOrCreateUserSheet();
  const data = userSheet.getDataRange().getValues();
  const headers = data[0];
  
  const emailCol = headers.indexOf('Email');
  const verifiedCol = headers.indexOf('Is_Verified');
  const businessIdCol = headers.indexOf('Business_ID');
  
  const rowIndex = data.findIndex(row => row[emailCol] === email.toLowerCase());
  
  if (rowIndex !== -1) {
    userSheet.getRange(rowIndex + 1, verifiedCol + 1).setValue('YES');
    if (businessIdCol !== -1) {
      userSheet.getRange(rowIndex + 1, businessIdCol + 1).setValue(businessId);
    }
  }
}

function deleteBusinessAccount(businessId, email) {
  try {
    const businessSheet = getOrCreateBusinessSheet();
    const data = businessSheet.getDataRange().getValues();
    const headers = data[0];
    
    const bizIdCol = headers.indexOf('Business_ID');
    const activeCol = headers.indexOf('Is_Active');
    
    const rowIndex = data.findIndex(r => r[bizIdCol] === businessId);
    if (rowIndex !== -1 && activeCol !== -1) {
      businessSheet.getRange(rowIndex + 1, activeCol + 1).setValue('NO');
    }
    
    // Notify admin
    GmailApp.sendEmail(EMAIL_CONFIG.supportEmail, `Account Deletion Request - ${businessId}`, `Business ${businessId} requested deletion. Owner: ${email}`);
    
    return { success: true };
  } catch(error) {
    return { success: false, message: error.message };
  }
}
