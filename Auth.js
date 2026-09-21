// ============================================================
// 📁 Auth.gs - FIXED VERSION
// ============================================================

/**
 * ✅ AUTHENTICATE USER - WITH PROPER CLIENT HANDLING
 */
function authenticateUser(email, password, businessName) {
  try {
    console.log('🔐 authenticateUser called:', { email, businessName });
    
    const userSheet = getOrCreateUserSheet();
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    const emailCol = headers.indexOf('Email');
    const passwordCol = headers.indexOf('Password_Hash');
    const nameCol = headers.indexOf('Name');
    const roleCol = headers.indexOf('Role');
    const businessIdCol = headers.indexOf('Business_ID');
    const businessNameCol = headers.indexOf('Business_Name');
    const isVerifiedCol = headers.indexOf('Is_Verified');
    const isDemoCol = headers.indexOf('Is_Demo');
    const isClientCol = headers.indexOf('Is_Client');
    
    if (emailCol === -1) {
      console.error('❌ Email column not found!');
      return { success: false, message: 'System error: Email column missing' };
    }
    
    // ===== FIND USER =====
    // A staff member may belong to more than one business with the same email.
    // When a business name is supplied, select the matching business row instead
    // of assuming one email can only belong to one workspace.
    let userRow = null;
    let rowIndex = -1;
    const normalizedEmail = String(email || '').trim().toLowerCase();
    const normalizedBusinessName = String(businessName || '').trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][emailCol] || '').trim().toLowerCase();
      const rowBusinessName = businessNameCol !== -1
        ? String(data[i][businessNameCol] || '').trim().toLowerCase()
        : '';

      if (rowEmail !== normalizedEmail) continue;
      if (normalizedBusinessName && rowBusinessName !== normalizedBusinessName) continue;

      userRow = data[i];
      rowIndex = i;
      break;
    }

    if (!userRow) {
      console.log('❌ User not found:', email);
      return { success: false, message: "Email not found. Please sign up first." };
    }
    
    console.log('✅ User found:', email);
    
    // ===== DETERMINE IF CLIENT USER =====
    const isClientUser = isClientCol !== -1 && userRow[isClientCol] === 'YES';
    console.log('🆔 Is client user:', isClientUser);
    
    // ===== BUSINESS NAME VALIDATION (Only for main app users) =====
    if (!isClientUser && businessName && businessNameCol !== -1) {
      const userBusinessName = userRow[businessNameCol];
      console.log('🏢 Business name check:', { provided: businessName, stored: userBusinessName, isClient: isClientUser });
      
      if (!userBusinessName || userBusinessName.trim() === '') {
        console.log('❌ Business name not set in user record');
        return { success: false, message: "Account setup incomplete. Please contact support." };
      }
      
      if (userBusinessName.toLowerCase() !== businessName.toLowerCase()) {
        console.log('❌ Business name mismatch');
        return { success: false, message: "This email is not registered under this business name. Please check your business name or contact your business owner." };
      }
      
      console.log('✅ Business name matched');
    } else if (isClientUser) {
      console.log('ℹ️ Skipping business name validation for client user');
    }
    
    // ===== PASSWORD VALIDATION =====
    const isDemo = isDemoCol !== -1 && userRow[isDemoCol] === 'YES';
    let isPasswordValid = false;
    
    if (isDemo) {
      console.log('✅ Demo account - password check skipped');
      isPasswordValid = true;
    } else if (passwordCol !== -1) {
      const hashedPassword = hashPassword(password);
      isPasswordValid = userRow[passwordCol] === hashedPassword;
      
      if (isPasswordValid) {
        console.log('✅ Password is correct');
      } else {
        console.log('❌ Password is incorrect');
      }
    } else {
      console.log('❌ Password column not found!');
      return { success: false, message: 'System error: Password column missing' };
    }
    
    if (!isPasswordValid) {
      return { success: false, message: "Incorrect password" };
    }
    
    // ===== VERIFICATION CHECK (Skip for demo and client users) =====
    if (!isDemo && !isClientUser) {
      const isVerified = isVerifiedCol !== -1 && userRow[isVerifiedCol] === 'YES';
      if (!isVerified) {
        console.log('⚠️ User not verified');
        return { success: false, message: "Your account is pending verification. Please check your email for verification instructions." };
      }
      console.log('✅ User is verified');
    }
    
    // ===== UPDATE LAST LOGIN =====
    const lastLoginCol = headers.indexOf('Last_Login');
    if (lastLoginCol !== -1 && rowIndex !== -1) {
      userSheet.getRange(rowIndex + 1, lastLoginCol + 1).setValue(new Date().toISOString());
      console.log('📅 Updated last login');
    }
    
    // ===== CREATE SESSION =====
    const sessionId = createSession(email.toLowerCase(), businessId);
    console.log('🔑 Session created:', sessionId.substring(0, 20) + '...');
    
    // ===== GET BUSINESS INFO =====
    const businessId = businessIdCol !== -1 ? userRow[businessIdCol] : '';
    const businessSheet = getOrCreateBusinessSheet();
    const businessData = businessSheet.getDataRange().getValues();
    const businessHeaders = businessData[0];
    
    const bizIdCol = businessHeaders.indexOf('Business_ID');
    const tierCol = businessHeaders.indexOf('Subscription_Tier');
    const workspaceIdCol = businessHeaders.indexOf('Workspace_ID');
    
    let subscriptionTier = 'free';
    let workspaceId = null;
    let businessRowFound = null;
    
    if (businessId && bizIdCol !== -1) {
      businessRowFound = businessData.find(row => row[bizIdCol] === businessId);
      if (businessRowFound) {
        subscriptionTier = tierCol !== -1 ? businessRowFound[tierCol] || 'free' : 'free';
        workspaceId = workspaceIdCol !== -1 ? businessRowFound[workspaceIdCol] || null : null;
        console.log('📊 Business found:', { businessId, tier: subscriptionTier });
      }
    }
    
    // ===== BUILD USER RESPONSE =====
    const user = {
      email: userRow[emailCol],
      name: nameCol !== -1 ? userRow[nameCol] || email : email,
      role: roleCol !== -1 ? userRow[roleCol] || 'owner' : 'owner',
      businessId: businessId || '',
      businessName: businessNameCol !== -1 ? userRow[businessNameCol] || businessName || '' : businessName || '',
      subscriptionTier: subscriptionTier,
      workspaceId: workspaceId,
      isDemo: isDemo,
      isClient: isClientUser
    };
    
    console.log('✅ ========================================');
    console.log('✅ LOGIN SUCCESSFUL');
    console.log('✅ ========================================');
    console.log('📧 Email:', user.email);
    console.log('👤 Name:', user.name);
    console.log('👨‍💼 Role:', user.role);
    console.log('🏢 Business:', user.businessName);
    console.log('🆔 Is Client:', user.isClient);
    console.log('✅ ========================================');
    
    return { 
      success: true, 
      user: user,
      sessionId: sessionId,
      message: "Login successful" 
    };
    
  } catch (error) {
    console.error('❌ authenticateUser error:', error);
    console.error('❌ Stack:', error.stack);
    return { success: false, message: "Login failed: " + error.message };
  }
}

/**
 * Hash password using SHA-256
 */
function hashPassword(password) {
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password)
  );
}

/**
 * Create session and cache it
 */
function createSession(email, businessId) {
  const sessionId = Utilities.getUuid();
  const cache = CacheService.getScriptCache();
  const sessionData = {
    email: email,
    businessId: businessId || '',
    createdAt: new Date().toISOString()
  };
  // Store for 24 hours (86400 seconds)
  cache.put(sessionId, JSON.stringify(sessionData), 86400);
  console.log("📝 Session created for:", email, "Session ID:", sessionId.substring(0, 20) + "...");
  return sessionId;
}

/**
 * Validate session from cache
 */
function getSessionData_(sessionId) {
  if (!sessionId) return null;

  const cache = CacheService.getScriptCache();
  const sessionData = cache.get(sessionId);
  if (!sessionData) return null;

  try {
    return JSON.parse(sessionData);
  } catch (error) {
    console.error('❌ Session data parse error:', error);
    return null;
  }
}

/**
 * Validate session from cache
 */
function validateSession(sessionId) {
  if (!sessionId) {
    console.log('⚠️ No sessionId provided');
    return null;
  }
  
  const session = getSessionData_(sessionId);
  if (!session) {
    console.log('⚠️ No session data found for:', sessionId.substring(0, 20) + '...');
    return null;
  }

  console.log('✅ Session validated for:', session.email);
  return session.email;
}

/**
 * Get user from session with full details
 */
function getUserFromSession(sessionId) {
  console.log("👤 getUserFromSession called with sessionId:", sessionId ? sessionId.substring(0, 20) + "..." : "null");
  
  const email = validateSession(sessionId);
  if (!email) {
    console.log("⚠️ Session validation failed");
    return null;
  }
  
  console.log("✅ Validated email:", email);
  
  const userSheet = getOrCreateUserSheet();
  const data = userSheet.getDataRange().getValues();
  const headers = data[0];
  
  const emailCol = headers.indexOf('Email');
  const nameCol = headers.indexOf('Name');
  const roleCol = headers.indexOf('Role');
  const businessIdCol = headers.indexOf('Business_ID');
  const businessNameCol = headers.indexOf('Business_Name');
  const isDemoCol = headers.indexOf('Is_Demo');
  const isClientCol = headers.indexOf('Is_Client');
  
  if (emailCol === -1) {
    console.error("❌ Email column not found!");
    return null;
  }
  
  const session = getSessionData_(sessionId);
  const sessionBusinessId = session && session.businessId ? String(session.businessId) : '';

  const matchingRows = data.map(function(row, index) {
    return { row: row, index: index };
  }).filter(function(item) {
    const rowEmail = String(item.row[emailCol] || '').trim().toLowerCase();
    const rowBusinessId = businessIdCol !== -1 ? String(item.row[businessIdCol] || '') : '';
    return rowEmail === String(email || '').trim().toLowerCase() &&
      (!sessionBusinessId || rowBusinessId === sessionBusinessId);
  });

  if (!matchingRows.length) {
    console.log("⚠️ User row not found for email:", email);
    return null;
  }
  
  const userRow = matchingRows[0].row;

  // Get business info
  const businessSheet = getOrCreateBusinessSheet();
  const businessData = businessSheet.getDataRange().getValues();
  const businessHeaders = businessData[0];
  const bizIdCol = businessHeaders.indexOf('Business_ID');
  const tierCol = businessHeaders.indexOf('Subscription_Tier');
  const workspaceIdCol = businessHeaders.indexOf('Workspace_ID');
  
  const businessId = businessIdCol !== -1 ? userRow[businessIdCol] : '';
  const businessRow = businessData.find(row => row[bizIdCol] === businessId);
  
  const userRole = roleCol !== -1 ? userRow[roleCol] || 'staff' : 'staff';
  const isDemo = isDemoCol !== -1 && userRow[isDemoCol] === 'YES';
  const isClient = isClientCol !== -1 && userRow[isClientCol] === 'YES';
  
  // ===== GET ACCESSIBLE MODULES =====
  let accessibleModules = [];
  const isOwnerOrAdmin = userRole === 'owner' || userRole === 'admin';
  const subscriptionTier = businessRow ? (tierCol !== -1 ? businessRow[tierCol] : 'free') : 'free';
  
  if (isOwnerOrAdmin) {
    if (isDemo) {
      accessibleModules = ['Finance', 'Ecommerce', 'Sales', 'CRM', 'HR', 'Logistics', 'Tax', 'Agro', 'Productivity', 'POS', 'Attendance', 'Warehouse'];
    } else if (subscriptionTier === 'free' || subscriptionTier === 'starter') {
      accessibleModules = ['Ecommerce'];
    } else {
      accessibleModules = ['Finance', 'Ecommerce', 'Sales', 'CRM', 'HR', 'Logistics', 'Tax', 'Agro', 'Productivity', 'POS', 'Attendance', 'Warehouse'];
    }
  } else {
    accessibleModules = getStaffModules(email, businessId);
    if (accessibleModules.length === 0) {
      accessibleModules = [];
    }
  }
  
  const result = {
    email: userRow[emailCol],
    name: nameCol !== -1 ? userRow[nameCol] : email,
    role: userRole,
    businessId: businessId,
    businessName: businessNameCol !== -1 ? userRow[businessNameCol] : '',
    subscriptionTier: subscriptionTier,
    workspaceId: businessRow ? (workspaceIdCol !== -1 ? businessRow[workspaceIdCol] : null) : null,
    isDemo: isDemo,
    isClient: isClient,
    accessibleModules: accessibleModules
  };
  
  console.log("✅ User loaded:", result.email, "Role:", result.role, "Modules:", result.accessibleModules.length);
  return result;
}

/**
 * Logout - remove session
 */
function logout(sessionId) {
  if (sessionId) {
    const cache = CacheService.getScriptCache();
    cache.remove(sessionId);
    console.log('🚪 Logged out session:', sessionId.substring(0, 20) + '...');
  }
  return { success: true, message: "Logged out successfully" };
}

/**
 * Reset password flow
 */
function resetPassword(email) {
  try {
    const userSheet = getOrCreateUserSheet();
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    const isDemoCol = headers.indexOf('Is_Demo');
    
    const userRow = data.find(row => row[emailCol] && row[emailCol].toLowerCase() === email.toLowerCase());
    
    if (!userRow) {
      return { success: false, message: "Email not found" };
    }
    
    if (isDemoCol !== -1 && userRow[isDemoCol] === 'YES') {
      return { success: false, message: "Demo accounts do not support password reset. Please use demo credentials." };
    }
    
    const resetToken = Utilities.getUuid();
    let resetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Password_Resets');
    
    if (!resetSheet) {
      resetSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Password_Resets');
      resetSheet.appendRow(['Email', 'Token', 'Created_At', 'Used']);
    }
    
    resetSheet.appendRow([email.toLowerCase(), resetToken, new Date().toISOString(), false]);
    
    // Send email with reset link
    sendPasswordResetEmail(email, resetToken);
    
    return { 
      success: true, 
      message: "Password reset instructions have been sent to " + email 
    };
    
  } catch (error) {
    console.error('❌ Reset password error:', error);
    return { success: false, message: "Failed to send reset link" };
  }
}

/**
 * Confirm password reset with token
 */
function confirmPasswordReset(token, newPassword) {
  try {
    let resetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Password_Resets');
    if (!resetSheet) {
      return { success: false, message: "Invalid reset request" };
    }
    
    const data = resetSheet.getDataRange().getValues();
    const headers = data[0];
    const tokenCol = headers.indexOf('Token');
    const emailCol = headers.indexOf('Email');
    const usedCol = headers.indexOf('Used');
    const createdAtCol = headers.indexOf('Created_At');
    
    const resetRow = data.find(row => row[tokenCol] === token && row[usedCol] !== true);
    
    if (!resetRow) {
      return { success: false, message: "Invalid or expired reset token" };
    }
    
    // Check if token is expired (1 hour)
    const createdAt = new Date(resetRow[createdAtCol]);
    const now = new Date();
    const hoursDiff = (now - createdAt) / (1000 * 60 * 60);
    
    if (hoursDiff > 1) {
      return { success: false, message: "Reset token has expired. Please request a new one." };
    }
    
    const email = resetRow[emailCol];
    const rowIndex = data.findIndex(row => row[tokenCol] === token);
    
    // Mark token as used
    resetSheet.getRange(rowIndex + 1, usedCol + 1).setValue(true);
    
    // Update user password
    const userSheet = getOrCreateUserSheet();
    const userData = userSheet.getDataRange().getValues();
    const userHeaders = userData[0];
    const userEmailCol = userHeaders.indexOf('Email');
    const passwordCol = userHeaders.indexOf('Password_Hash');
    
    const userRowIndex = userData.findIndex(row => row[userEmailCol] === email);
    
    if (userRowIndex !== -1 && passwordCol !== -1) {
      const hashedPassword = hashPassword(newPassword);
      userSheet.getRange(userRowIndex + 1, passwordCol + 1).setValue(hashedPassword);
      return { success: true, message: "Password reset successful. You can now login." };
    }
    
    return { success: false, message: "User not found" };
    
  } catch (error) {
    console.error('❌ Confirm reset error:', error);
    return { success: false, message: error.message };
  }
}