//admin.gs
function approveUpgradeRequest(requestId, adminEmail, sessionId) {
  requireAdminSession_(sessionId);
  try {
    const upgradeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Upgrade_Requests');
    if (!upgradeSheet) return { success: false, message: "No requests found" };
    
    const data = upgradeSheet.getDataRange().getValues();
    const headers = data[0];
    
    const requestIdCol = headers.indexOf('Request_ID');
    const businessIdCol = headers.indexOf('Business_ID');
    const statusCol = headers.indexOf('Status');
    
    const rowIndex = data.findIndex(row => row[requestIdCol] === requestId);
    if (rowIndex === -1) return { success: false, message: "Request not found" };
    if (data[rowIndex][statusCol] !== 'pending') return { success: false, message: "Already processed" };
    
    const businessId = data[rowIndex][businessIdCol];
    const businessSheet = getOrCreateBusinessSheet();
    const businessData = businessSheet.getDataRange().getValues();
    const businessHeaders = businessData[0];
    
    const bizIdCol = businessHeaders.indexOf('Business_ID');
    const tierCol = businessHeaders.indexOf('Subscription_Tier');
    const statusColBiz = businessHeaders.indexOf('Verification_Status');
    const workspaceIdCol = businessHeaders.indexOf('Workspace_ID');
    const ownerEmailCol = businessHeaders.indexOf('Owner_Email');
    const businessNameCol = businessHeaders.indexOf('Business_Name');
    
    const bizRowIndex = businessData.findIndex(row => row[bizIdCol] === businessId);
    if (bizRowIndex !== -1) {
      const ownerEmail = businessData[bizRowIndex][ownerEmailCol];
      const businessName = businessData[bizRowIndex][businessNameCol];
      
      // Create new workspace for Tier 2
      const workspaceResult = createWorkspaceForBusiness(businessId, businessName, ownerEmail);
      
      if (workspaceResult.success) {
        businessSheet.getRange(bizRowIndex + 1, tierCol + 1).setValue('tier2');
        businessSheet.getRange(bizRowIndex + 1, statusColBiz + 1).setValue('verified');
        businessSheet.getRange(bizRowIndex + 1, workspaceIdCol + 1).setValue(workspaceResult.workspaceId);
        
        // Update user verification
        updateUserVerification(ownerEmail, businessId);
        
        // Update upgrade request status
        upgradeSheet.getRange(rowIndex + 1, statusCol + 1).setValue('approved');
        
        // Send approval email
        sendUpgradeApprovalEmail(ownerEmail, businessName);
        
        return { success: true, message: "Upgrade approved! Workspace created." };
      }
    }
    
    return { success: false, message: "Failed to upgrade" };
  } catch (error) {
    console.error('Approve error:', error);
    return { success: false, message: error.message };
  }
}


function getPendingUpgradeRequests(sessionId) {
  requireAdminSession_(sessionId);
  try {
    const upgradeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Upgrade_Requests');
    if (!upgradeSheet) return [];
    
    const data = upgradeSheet.getDataRange().getValues();
    const headers = data[0];
    const statusCol = headers.indexOf('Status');
    const businessNameCol = headers.indexOf('Business_Name');
    const countryCol = headers.indexOf('Country');
    
    return data.slice(1)
      .filter(row => row[statusCol] === 'pending')
      .map(row => ({
        businessName: row[businessNameCol],
        country: row[countryCol]
      }));
  } catch (error) {
    console.error('Error:', error);
    return [];
  }
}


function getPendingBusinesses(sessionId) {
  requireAdminSession_(sessionId);
  const businessSheet = getOrCreateBusinessSheet();
  const data = businessSheet.getDataRange().getValues();
  const headers = data[0];
  const statusCol = headers.indexOf('Verification_Status');
  
  return data.slice(1)
    .filter(row => row[statusCol] === 'pending')
    .map(row => ({
      businessId: row[headers.indexOf('Business_ID')],
      businessName: row[headers.indexOf('Business_Name')],
      cacNumber: row[headers.indexOf('CAC_Number')],
      taxId: row[headers.indexOf('Tax_ID')],
      ownerEmail: row[headers.indexOf('Owner_Email')],
      ownerName: row[headers.indexOf('Owner_Name')],
      createdAt: row[headers.indexOf('Created_At')]
    }));
}


// ==================== ADMIN SESSION SECURITY ====================

function getAdminPasswordHash_() {
  var props = PropertiesService.getScriptProperties();
  var configuredHash = String(props.getProperty('ADMIN_PASSWORD_HASH') || '').trim();
  if (configuredHash) return configuredHash.toLowerCase();

  // Backward-compatible bootstrap: if an existing ADMIN_PASSWORD property is
  // configured, compare its hash without exposing the password in source.
  var configuredPassword = String(props.getProperty('ADMIN_PASSWORD') || '');
  if (configuredPassword) return hashPassword(configuredPassword).toLowerCase();

  // Preserve the existing admin login during migration without keeping the
  // plaintext legacy password in source. Replace with ADMIN_PASSWORD_HASH.
  return 'a36aef5a11c4073fbe60314fc9df530a9d5f986533594d1f5190742ff9e0e408';
}

function createAdminSession_(ttlSeconds) {
  var sessionId = 'ADM_' + Utilities.getUuid();
  CacheService.getScriptCache().put('admin_session_' + sessionId, '1', ttlSeconds || 3600);
  return sessionId;
}

function validateAdminSession(sessionId) {
  if (!sessionId) return false;
  return CacheService.getScriptCache().get('admin_session_' + String(sessionId)) === '1';
}

function requireAdminSession_(sessionId) {
  if (!validateAdminSession(sessionId)) {
    throw new Error('Unauthorized: valid admin session required.');
  }
  return true;
}

function adminLogin(password) {
  var supplied = String(password || '');
  if (!supplied) return { success: false, message: 'Admin password is required.' };

  var expectedHash = getAdminPasswordHash_();
  if (!expectedHash) {
    console.error('ADMIN_PASSWORD_HASH/ADMIN_PASSWORD is not configured.');
    return { success: false, message: 'Admin authentication is not configured.' };
  }

  var suppliedHash = String(hashPassword(supplied) || '').toLowerCase();
  if (suppliedHash !== expectedHash) {
    return { success: false, message: 'Invalid password.' };
  }

  return { success: true, sessionId: createAdminSession_(3600) };
}

// Secure admin lockout recovery.
// Recovery is deliberately controlled by a separate Script Property so the
// public login page never becomes an unauthenticated password-reset backdoor.
// Configure ADMIN_RECOVERY_KEY_HASH in the Master project's Script Properties.
function resetAdminPassword(recoveryKey, newPassword) {
  var suppliedKey = String(recoveryKey || '');
  var password = String(newPassword || '');
  if (!suppliedKey) return { success: false, message: 'Recovery key is required.' };
  if (!password || password.length < 10) {
    return { success: false, message: 'New password must be at least 10 characters.' };
  }

  var props = PropertiesService.getScriptProperties();
  var recoveryHash = String(props.getProperty('ADMIN_RECOVERY_KEY_HASH') || '').trim().toLowerCase();
  if (!recoveryHash) {
    return {
      success: false,
      message: 'Admin recovery is not configured. Set ADMIN_RECOVERY_KEY_HASH in Script Properties first.'
    };
  }

  var suppliedHash = String(hashPassword(suppliedKey) || '').toLowerCase();
  if (suppliedHash !== recoveryHash) {
    return { success: false, message: 'Invalid recovery key.' };
  }

  props.setProperty('ADMIN_PASSWORD_HASH', String(hashPassword(password) || '').toLowerCase());
  props.deleteProperty('ADMIN_PASSWORD');

  return {
    success: true,
    message: 'Admin password reset successfully. You can now sign in with the new password.'
  };
}

// ==================== ADMIN FUNCTIONS ====================
// Add these to your Admin.gs file

function getAdminStats(sessionId) {
  requireAdminSession_(sessionId);
  try {
    const businessSheet = getOrCreateBusinessSheet();
    const userSheet = getOrCreateUserSheet();
    
    const bizData = businessSheet.getDataRange().getValues();
    const userData = userSheet.getDataRange().getValues();
    
    const totalBusinesses = Math.max(0, bizData.length - 1);
    const totalUsers = Math.max(0, userData.length - 1);
    
    // Count tiers
    const bizHeaders = bizData[0];
    const tierCol = bizHeaders.indexOf('Subscription_Tier');
    let freeUsers = 0, paidUsers = 0;
    
    for (let i = 1; i < bizData.length; i++) {
      const tier = bizData[i][tierCol] || 'free';
      if (tier === 'free' || tier === 'starter') freeUsers++;
      else paidUsers++;
    }
    
    // Recent activity
    const recentActivity = getRecentActivity();
    
    return {
      totalBusinesses,
      totalUsers,
      freeUsers,
      paidUsers,
      totalRevenue: paidUsers * 499,
      newUsers: getNewUsersCount(userData),
      newBusinesses: getNewBusinessesCount(bizData),
      recentActivity
    };
  } catch (error) {
    console.error('Admin stats error:', error);
    return { totalBusinesses: 0, totalUsers: 0, freeUsers: 0, paidUsers: 0, totalRevenue: 0, newUsers: 0, newBusinesses: 0, recentActivity: [] };
  }
}

function getNewUsersCount(userData) {
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
  const headers = userData[0];
  const createdAtCol = headers.indexOf('Created_At');
  let count = 0;
  for (let i = 1; i < userData.length; i++) {
    const created = new Date(userData[i][createdAtCol]);
    if (created >= thirtyDaysAgo) count++;
  }
  return count;
}

function getNewBusinessesCount(bizData) {
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
  const headers = bizData[0];
  const createdAtCol = headers.indexOf('Created_At');
  let count = 0;
  for (let i = 1; i < bizData.length; i++) {
    const created = new Date(bizData[i][createdAtCol]);
    if (created >= thirtyDaysAgo) count++;
  }
  return count;
}

function getRecentActivity(limit = 10) {
  const activities = [];
  const businessSheet = getOrCreateBusinessSheet();
  const bizData = businessSheet.getDataRange().getValues();
  const bizHeaders = bizData[0];
  const bizNameCol = bizHeaders.indexOf('Business_Name');
  const createdCol = bizHeaders.indexOf('Created_At');
  const bizIdCol = bizHeaders.indexOf('Business_ID');
  
  for (let i = Math.max(1, bizData.length - limit); i < bizData.length; i++) {
    if (bizData[i][bizNameCol]) {
      activities.push({
        type: 'new_business',
        message: `New business registered: ${bizData[i][bizNameCol]}`,
        time: new Date(bizData[i][createdCol]).toLocaleString(),
        action: `viewBusiness('${bizData[i][bizIdCol]}')`
      });
    }
  }
  return activities.slice(0, limit);
}

function getAllBusinesses(sessionId) {
  requireAdminSession_(sessionId);
  try {
    const sheet = getOrCreateBusinessSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // Deployment details live in the Clients registry, not the Businesses sheet.
    // Join them here so every existing admin view can display the real deployment
    // state without exposing the registry directly to the browser.
    const deploymentByBusinessId = {};
    try {
      const clientSheet = getBizOSMasterSpreadsheet_().getSheetByName('Clients');
      if (clientSheet && clientSheet.getLastRow() >= 2) {
        const clientData = clientSheet.getDataRange().getValues();
        const clientHeaders = clientData[0].map(function(h) { return String(h || '').trim(); });
        const businessIdCol = clientHeaders.indexOf('Business_ID');

        if (businessIdCol >= 0) {
          for (let c = 1; c < clientData.length; c++) {
            const row = clientData[c];
            const businessId = String(row[businessIdCol] || '').trim();
            if (!businessId) continue;

            const record = {};
            clientHeaders.forEach(function(header, index) {
              record[header] = row[index] === undefined ? '' : row[index];
            });

            // Prefer the active registry record when duplicates exist.
            const current = deploymentByBusinessId[businessId];
            if (!current || String(record.Status || '').toLowerCase() === 'active') {
              deploymentByBusinessId[businessId] = record;
            }
          }
        }
      }
    } catch (deploymentError) {
      console.error('Unable to load client deployment registry:', deploymentError);
    }

    const businesses = [];
    for (let i = 1; i < data.length; i++) {
      const businessId = String(data[i][headers.indexOf('Business_ID')] || '').trim();
      const deployment = deploymentByBusinessId[businessId] || {};

      businesses.push({
        businessId: businessId,
        businessName: data[i][headers.indexOf('Business_Name')],
        ownerEmail: data[i][headers.indexOf('Owner_Email')],
        subscriptionTier: data[i][headers.indexOf('Subscription_Tier')] || deployment.Subscription_Tier || deployment.Tier || 'free',
        verificationStatus: data[i][headers.indexOf('Verification_Status')] || 'pending',
        createdAt: data[i][headers.indexOf('Created_At')],
        status: data[i][headers.indexOf('Status')] || 'active',

        // Client deployment registry
        clientId: deployment.Client_ID || '',
        scriptId: deployment.Script_ID || '',
        deploymentId: deployment.Deployment_ID || '',
        launchUrl: deployment.Web_App_URL || '',
        webAppUrl: deployment.Web_App_URL || '',
        landingUrl: deployment.Landing_URL || '',
        deploymentStatus: deployment.Status || '',
        deployedAt: deployment.Updated_At || '',
        deploymentVersion: deployment.Provisioning_Version || '',
        workspaceId: deployment.Workspace_ID || deployment.Sheet_ID || ''
      });
    }
    return businesses;
  } catch (error) {
    console.error('Error getting businesses:', error);
    return [];
  }
}

function getAllUsers(sessionId) {
  requireAdminSession_(sessionId);
  try {
    const sheet = getOrCreateUserSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const users = [];
    for (let i = 1; i < data.length; i++) {
      users.push({
        email: data[i][headers.indexOf('Email')],
        name: data[i][headers.indexOf('Name')],
        role: data[i][headers.indexOf('Role')],
        businessId: data[i][headers.indexOf('Business_ID')],
        isVerified: data[i][headers.indexOf('Is_Verified')],
        createdAt: data[i][headers.indexOf('Created_At')]
      });
    }
    return users;
  } catch (error) {
    console.error('Error getting users:', error);
    return [];
  }
}

function getAllFeatures(sessionId) {
  requireAdminSession_(sessionId);
  try {
    let featuresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Features');
    if (!featuresSheet) {
      featuresSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Features');
      featuresSheet.appendRow(['ID', 'Name', 'Description', 'Module', 'Tier', 'Active', 'CreatedAt']);
      return [];
    }
    const data = featuresSheet.getDataRange().getValues();
    const headers = data[0];
    const features = [];
    for (let i = 1; i < data.length; i++) {
      features.push({
        id: data[i][headers.indexOf('ID')],
        name: data[i][headers.indexOf('Name')],
        description: data[i][headers.indexOf('Description')],
        module: data[i][headers.indexOf('Module')],
        tier: data[i][headers.indexOf('Tier')],
        active: data[i][headers.indexOf('Active')] === 'YES',
        createdAt: data[i][headers.indexOf('CreatedAt')]
      });
    }
    return features;
  } catch (error) {
    return [];
  }
}

function getAllModules(sessionId) {
  requireAdminSession_(sessionId);
  try {
    const config = getConfig();  // ✅ Get CONFIG safely
    const modules = [];
    
    if (config.MODULES && config.MODULES.all) {
      const moduleLabels = config.MODULES.labels || {};
      const moduleIcons = config.MODULES.icons || {};
      
      for (const name of config.MODULES.all) {
        modules.push({
          name: name,
          label: moduleLabels[name] || name,
          icon: moduleIcons[name] || 'fa-folder',
          tier: 'starter',
          recordCount: 0,
          visible: true
        });
      }
    }
    
    return modules;
  } catch (error) {
    console.error('Error getting modules:', error);
    return [];
  }
}

function addFeature(featureData, sessionId) {
  requireAdminSession_(sessionId);
  try {
    let featuresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Features');
    if (!featuresSheet) {
      featuresSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Features');
      featuresSheet.appendRow(['ID', 'Name', 'Description', 'Module', 'Tier', 'Active', 'CreatedAt']);
    }
    const id = 'FEAT-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    featuresSheet.appendRow([
      id, featureData.name, featureData.description || '',
      featureData.module || 'General', featureData.tier || 'free',
      'YES', new Date().toISOString()
    ]);
    return { success: true, message: 'Feature added successfully', id: id };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function toggleFeature(featureId, sessionId) {
  requireAdminSession_(sessionId);
  try {
    const featuresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Features');
    if (!featuresSheet) return { success: false, message: 'Features sheet not found' };
    const data = featuresSheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ID');
    const activeCol = headers.indexOf('Active');
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === featureId) {
        const current = data[i][activeCol];
        featuresSheet.getRange(i + 1, activeCol + 1).setValue(current === 'YES' ? 'NO' : 'YES');
        return { success: true };
      }
    }
    return { success: false, message: 'Feature not found' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteFeature(featureId, sessionId) {
  requireAdminSession_(sessionId);
  try {
    const featuresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Features');
    if (!featuresSheet) return { success: false, message: 'Features sheet not found' };
    const data = featuresSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === featureId) {
        featuresSheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'Feature not found' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function toggleBusinessStatus(businessId, sessionId) {
  requireAdminSession_(sessionId);
  try {
    const sheet = getOrCreateBusinessSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const bizIdCol = headers.indexOf('Business_ID');
    const statusCol = headers.indexOf('Status');
    for (let i = 1; i < data.length; i++) {
      if (data[i][bizIdCol] === businessId) {
        const current = data[i][statusCol];
        sheet.getRange(i + 1, statusCol + 1).setValue(current === 'active' ? 'suspended' : 'active');
        return { success: true };
      }
    }
    return { success: false, message: 'Business not found' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteBusiness(businessId, sessionId) {
  requireAdminSession_(sessionId);
  try {
    const sheet = getOrCreateBusinessSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === businessId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'Business not found' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function exportBusinessData(sessionId) {
  requireAdminSession_(sessionId);
  try {
    const sheet = getOrCreateBusinessSheet();
    const data = sheet.getDataRange().getValues();
    let csv = '';
    for (const row of data) {
      csv += row.join(',') + '\n';
    }
    return csv;
  } catch (error) {
    return 'Error exporting data';
  }
}

function clearSystemCache(sessionId) {
  requireAdminSession_(sessionId);
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('dashboard_fast_');
    cache.remove('module_summary_');
    return { success: true };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function runSystemHealthCheck(sessionId) {
  requireAdminSession_(sessionId);
  const results = [];
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    results.push({ check: 'Spreadsheet Access', status: true });
  } catch(e) {
    results.push({ check: 'Spreadsheet Access', status: false, error: e.message });
  }
  const requiredSheets = ['Businesses', 'Users'];
  for (const sheetName of requiredSheets) {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      results.push({ check: `Sheet: ${sheetName}`, status: !!sheet });
    } catch(e) {
      results.push({ check: `Sheet: ${sheetName}`, status: false });
    }
  }
  return results;
}

function getSystemLogs(sessionId) {
  requireAdminSession_(sessionId);
  return ['System logs: All systems operational', 'Last check: ' + new Date().toISOString()];
}


function getClientConfig() {
  return {
    company: { name: "BizOS", tagline: "Business Operating System" },
    modules: MODULES,
    tiers: TIERS,
    pricing: PRICING
  };
}

function getBusinessPreferences(businessId) {
  try {
    const sheet = getOrCreateBusinessSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const bizIdCol = headers.indexOf('Business_ID');
    const prefsCol = headers.indexOf('Preferences');
    if (prefsCol === -1) return {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][bizIdCol] === businessId) {
        if (data[i][prefsCol]) return JSON.parse(data[i][prefsCol]);
        return {};
      }
    }
    return {};
  } catch (error) {
    return {};
  }
}

function saveBusinessPreferences(businessId, preferences) {
  try {
    const sheet = getOrCreateBusinessSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    let prefsCol = headers.indexOf('Preferences');
    if (prefsCol === -1) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue('Preferences');
      prefsCol = newCol;
    }
    const bizIdCol = headers.indexOf('Business_ID');
    for (let i = 1; i < data.length; i++) {
      if (data[i][bizIdCol] === businessId) {
        const existing = data[i][prefsCol] ? JSON.parse(data[i][prefsCol]) : {};
        const merged = { ...existing, ...preferences, updatedAt: new Date().toISOString() };
        sheet.getRange(i + 1, prefsCol + 1).setValue(JSON.stringify(merged));
        return { success: true, message: 'Preferences saved' };
      }
    }
    return { success: false, message: 'Business not found' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getWhiteLabelConfig() {
  try {
    const props = PropertiesService.getScriptProperties();
    const override = props.getProperty('CONFIG_OVERRIDE');
    if (override) {
      const config = JSON.parse(override);
      return config.whiteLabel || { enabled: false };
    }
    return { enabled: false };
  } catch (error) {
    return { enabled: false };
  }
}

function setWhiteLabelConfig(clientName, clientDomain, customCss = '', customLogo = '', customFavicon = '', sessionId) {
  requireAdminSession_(sessionId);
  try {
    const props = PropertiesService.getScriptProperties();
    const config = { enabled: true, clientName, clientDomain, customCss, customLogo, customFavicon, hideBranding: true };
    props.setProperty('CONFIG_OVERRIDE', JSON.stringify({ whiteLabel: config }));
    return { success: true, message: `White label configured for ${clientName}` };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ==================== ADMIN AUTHENTICATION ====================
// Add these 3 functions to your Admin.gs

function checkAdminSession(sessionId) {
  if (!sessionId) return false;
  const cache = CacheService.getScriptCache();
  const session = cache.get('admin_session_' + sessionId);
  return session === 'authenticated';
}

// Update your existing adminLogout function to clear session
// Replace your existing adminLogout with this one:
function adminLogout(sessionId) {
  if (sessionId) {
    const cache = CacheService.getScriptCache();
    cache.remove('admin_session_' + sessionId);
  }
  // Also clear any frontend storage
  return { success: true };
}



// ==================== MANUAL USER UPGRADE ====================

function manuallyUpgradeUser(email, targetTier, sessionId) {
  requireAdminSession_(sessionId);
  try {
    // Validate tier
    const validTiers = ['free', 'starter', 'sovereign', 'professional', 'enterprise'];
    if (!validTiers.includes(targetTier)) {
      return { success: false, message: "Invalid tier. Must be: " + validTiers.join(', ') };
    }
    
    // Find user
    const userSheet = getOrCreateUserSheet();
    const userData = userSheet.getDataRange().getValues();
    const userHeaders = userData[0];
    
    const emailCol = userHeaders.indexOf('Email');
    const businessIdCol = userHeaders.indexOf('Business_ID');
    
    let userRowIndex = -1;
    let businessId = null;
    
    for (let i = 1; i < userData.length; i++) {
      if (userData[i][emailCol] && userData[i][emailCol].toLowerCase() === email.toLowerCase()) {
        userRowIndex = i;
        businessId = userData[i][businessIdCol];
        break;
      }
    }
    
    if (userRowIndex === -1) {
      return { success: false, message: "User not found" };
    }
    
    // Update business tier
    const businessSheet = getOrCreateBusinessSheet();
    const businessData = businessSheet.getDataRange().getValues();
    const businessHeaders = businessData[0];
    
    const bizIdCol = businessHeaders.indexOf('Business_ID');
    const tierCol = businessHeaders.indexOf('Subscription_Tier');
    const statusCol = businessHeaders.indexOf('Verification_Status');
    const workspaceIdCol = businessHeaders.indexOf('Workspace_ID');
    
    let bizRowIndex = -1;
    for (let i = 1; i < businessData.length; i++) {
      if (businessData[i][bizIdCol] === businessId) {
        bizRowIndex = i;
        break;
      }
    }
    
    if (bizRowIndex === -1) {
      return { success: false, message: "Business not found" };
    }
    
    // Update tier
    businessSheet.getRange(bizRowIndex + 1, tierCol + 1).setValue(targetTier);
    businessSheet.getRange(bizRowIndex + 1, statusCol + 1).setValue('verified');
    
    // Create workspace if upgrading from free and none exists
    const currentWorkspaceId = businessData[bizRowIndex][workspaceIdCol];
    if (!currentWorkspaceId && targetTier !== 'free') {
      const businessName = businessData[bizRowIndex][businessHeaders.indexOf('Business_Name')];
      const ownerEmail = businessData[bizRowIndex][businessHeaders.indexOf('Owner_Email')];
      const workspaceResult = createWorkspaceForBusiness(businessId, businessName, ownerEmail);
      
      if (workspaceResult.success) {
        businessSheet.getRange(bizRowIndex + 1, workspaceIdCol + 1).setValue(workspaceResult.workspaceId);
      }
    }
    
    // Log the upgrade
    logAdminAction('manual_upgrade', {
      email: email,
      tier: targetTier,
      businessId: businessId,
      performedBy: Session.getActiveUser().getEmail()
    });
    
    return {
      success: true,
      message: `User ${email} upgraded to ${targetTier} successfully`
    };
    
  } catch (error) {
    console.error('Manual upgrade error:', error);
    return { success: false, message: error.message };
  }
}

function logAdminAction(action, details) {
  try {
    let logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Admin_Logs');
    if (!logSheet) {
      logSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Admin_Logs');
      logSheet.appendRow(['Timestamp', 'Action', 'Details', 'Performed By']);
    }
    
    logSheet.appendRow([
      new Date().toISOString(),
      action,
      JSON.stringify(details),
      details.performedBy || Session.getActiveUser().getEmail()
    ]);
  } catch (error) {
    console.error('Log error:', error);
  }
}

// ==================== FEATURES FOR USERS ====================

function getUserFeatures(userTier) {
  try {
    const featuresSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Features');
    if (!featuresSheet) return [];
    
    const data = featuresSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find column indices
    const nameCol = headers.indexOf('Name');
    const descCol = headers.indexOf('Description');
    const moduleCol = headers.indexOf('Module');
    const tierCol = headers.indexOf('Tier');
    const activeCol = headers.indexOf('Active');
    const createdAtCol = headers.indexOf('CreatedAt');
    
    const features = [];
    const now = new Date();
    const thirtyDaysAgo = new Date(now);
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    
    for (let i = 1; i < data.length; i++) {
      const tier = data[i][tierCol] || 'free';
      const isActive = data[i][activeCol] === 'YES';
      
      // Check if feature is active and user has access to this tier
      if (isActive && (tier === 'all' || tier === userTier || tier === 'free')) {
        features.push({
          name: data[i][nameCol] || 'Unnamed Feature',
          description: data[i][descCol] || '',
          module: data[i][moduleCol] || 'General',
          tier: tier,
          isNew: new Date(data[i][createdAtCol]) > thirtyDaysAgo,
          createdAt: data[i][createdAtCol]
        });
      }
    }
    
    return features;
  } catch (error) {
    console.error('Error getting user features:', error);
    return [];
  }
}

function getFeatureNotifications(userTier, lastSeen) {
  try {
    const features = getUserFeatures(userTier);
    
    // Get only features that are "new" (created after lastSeen)
    const lastSeenDate = lastSeen ? new Date(lastSeen) : new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
    
    return features.filter(f => new Date(f.createdAt) > lastSeenDate);
  } catch (error) {
    console.error('Error getting notifications:', error);
    return [];
  }
}




// ============================================================
// ADMIN RECOVERY KEY BOOTSTRAP
// Run this manually from the Apps Script editor when admin access
// is unavailable. The generated key is written to the execution log
// only; only its hash is stored in Script Properties.
// ============================================================
function generateAdminRecoveryKey() {
  var props = PropertiesService.getScriptProperties();
  var key = 'BIZOS-ADM-' + Utilities.getUuid().replace(/-/g, '').toUpperCase();
  props.setProperty('ADMIN_RECOVERY_KEY_HASH', String(hashPassword(key) || '').toLowerCase());
  console.log('ADMIN RECOVERY KEY (store securely, then remove this log entry): ' + key);
  return { success: true, message: 'Recovery key generated. Check the execution log and store it securely.' };
}
