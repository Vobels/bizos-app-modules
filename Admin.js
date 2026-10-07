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
    // Admin and provisioning must read the same canonical master spreadsheet.
    // Using SpreadsheetApp.getActiveSpreadsheet() here can point the web-app
    // execution at a different bound spreadsheet, which makes real Clients
    // records appear as if no paid deployments exist.
    const masterSpreadsheet = getBizOSMasterSpreadsheet_();
    const sheet = masterSpreadsheet.getSheetByName('Businesses') || getOrCreateBusinessSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // Deployment details live in the Clients registry, not the Businesses sheet.
    // Join them here so every existing admin view can display the real deployment
    // state without exposing the registry directly to the browser.
    const deploymentByBusinessId = {};
    const deploymentByEmail = {};
    try {
      const clientSheet = getBizOSMasterSpreadsheet_().getSheetByName('Clients');
      if (clientSheet && clientSheet.getLastRow() >= 2) {
        const clientData = clientSheet.getDataRange().getValues();
        const clientHeaders = clientData[0].map(function(h) { return String(h || '').trim(); });
        const businessIdCol = clientHeaders.indexOf('Business_ID');
        const emailCol = clientHeaders.indexOf('Email');

        for (let c = 1; c < clientData.length; c++) {
          const row = clientData[c];
          const businessId = businessIdCol >= 0 ? String(row[businessIdCol] || '').trim() : '';
          const email = emailCol >= 0 ? String(row[emailCol] || '').trim().toLowerCase() : '';
          if (!businessId && !email) continue;

          const record = {};
          clientHeaders.forEach(function(header, index) {
            record[header] = row[index] === undefined ? '' : row[index];
          });

          // Prefer the active registry record when duplicates exist.
          if (businessId) {
            const currentById = deploymentByBusinessId[businessId];
            if (!currentById || String(record.Status || '').toLowerCase() === 'active') {
              deploymentByBusinessId[businessId] = record;
            }
          }
          // Email is a safe secondary join for legacy/orphaned records where
          // Business_ID was not persisted consistently during provisioning.
          if (email) {
            const currentByEmail = deploymentByEmail[email];
            if (!currentByEmail || String(record.Status || '').toLowerCase() === 'active') {
              deploymentByEmail[email] = record;
            }
          }
        }
      }
    } catch (deploymentError) {
      console.error('Unable to load client deployment registry:', deploymentError);
    }

    const businesses = [];
    const businessByEmail = {};
    const businessById = {};
    const paidCustomerByEmail = {};

    for (let i = 1; i < data.length; i++) {
      const businessId = String(data[i][headers.indexOf('Business_ID')] || '').trim();
      const ownerEmail = String(data[i][headers.indexOf('Owner_Email')] || '').trim().toLowerCase();
      const deployment = deploymentByBusinessId[businessId] || deploymentByEmail[ownerEmail] || {};
      const record = {
        businessId: businessId,
        businessName: data[i][headers.indexOf('Business_Name')],
        ownerEmail: data[i][headers.indexOf('Owner_Email')],
        subscriptionTier: data[i][headers.indexOf('Subscription_Tier')] || deployment.Subscription_Tier || deployment.Tier || 'free',
        verificationStatus: data[i][headers.indexOf('Verification_Status')] || 'pending',
        createdAt: data[i][headers.indexOf('Created_At')],
        status: data[i][headers.indexOf('Status')] || 'active',
        clientId: deployment.Client_ID || '',
        scriptId: deployment.Script_ID || '',
        deploymentId: deployment.Deployment_ID || '',
        launchUrl: deployment.Web_App_URL || '',
        webAppUrl: deployment.Web_App_URL || '',
        landingUrl: deployment.Landing_URL || '',
        deploymentStatus: deployment.Status || '',
        deployedAt: deployment.Updated_At || '',
        deploymentVersion: deployment.Provisioning_Version || '',
        workspaceId: deployment.Workspace_ID || deployment.Sheet_ID || '',
        upgradeRequestId: '',
        paymentStatus: '',
        deploymentAction: 'none'
      };
      businesses.push(record);
      if (businessId) businessById[businessId] = record;
      if (ownerEmail) businessByEmail[ownerEmail] = record;
    }

    // A successful payment is the authoritative signal that a customer has
    // paid. Provisioning may still be queued or failed, so do not hide that
    // customer merely because the Businesses/Clients registries have not yet
    // been updated.
    try {
      const upgradeSheet = masterSpreadsheet.getSheetByName('Upgrade_Requests');
      const paymentSheet = masterSpreadsheet.getSheetByName('Payments');
      if (upgradeSheet && paymentSheet) {
        const upgradeData = upgradeSheet.getDataRange().getValues();
        const uh = upgradeData[0] || [];
        const paymentData = paymentSheet.getDataRange().getValues();
        const ph = paymentData[0] || [];
        const u = {
          requestId: uh.indexOf('Request_ID'),
          email: uh.indexOf('Email'),
          workspaceEmail: uh.indexOf('Workspace_Email'),
          name: uh.indexOf('Name'),
          businessName: uh.indexOf('Business_Name'),
          tier: uh.indexOf('Tier'),
          status: uh.indexOf('Status'),
          createdAt: uh.indexOf('Created_At')
        };
        const p = {
          requestId: ph.indexOf('Request_ID'),
          status: ph.indexOf('Status'),
          reference: ph.indexOf('Transaction_Ref')
        };
        const successfulPayments = {};
        if (p.requestId >= 0 && p.status >= 0) {
          for (let pi = 1; pi < paymentData.length; pi++) {
            if (String(paymentData[pi][p.status] || '').trim().toLowerCase() === 'success') {
              const rid = String(paymentData[pi][p.requestId] || '').trim();
              if (rid) successfulPayments[rid] = p.reference >= 0 ? String(paymentData[pi][p.reference] || '') : '';
            }
          }
        }

        if (u.requestId >= 0) {
          for (let ui = 1; ui < upgradeData.length; ui++) {
            const requestId = String(upgradeData[ui][u.requestId] || '').trim();
            if (!requestId || !successfulPayments.hasOwnProperty(requestId)) continue;

            const email = String(u.email >= 0 ? upgradeData[ui][u.email] || '' : '').trim().toLowerCase();
            const workspaceEmail = String(u.workspaceEmail >= 0 ? upgradeData[ui][u.workspaceEmail] || '' : '').trim().toLowerCase();
            const existing = businessByEmail[email] || businessByEmail[workspaceEmail];
            const requestStatus = String(u.status >= 0 ? upgradeData[ui][u.status] || '' : '').trim().toLowerCase();

            if (existing) {
              existing.paymentStatus = 'success';
              const existingPriority = ({'provisioned':5,'active':5,'provisioning':4,'payment_confirmed':3,'provisioning_failed':2,'pending_payment':1,'pending':1})[String(existing.deploymentStatus || '').toLowerCase()] || 0;
              const requestPriority = ({'provisioned':5,'active':5,'provisioning':4,'payment_confirmed':3,'provisioning_failed':2,'pending_payment':1,'pending':1})[requestStatus] || 0;
              if (!existing.upgradeRequestId || requestPriority >= existingPriority) {
                existing.upgradeRequestId = requestId;
                if (!existing.clientId && !existing.scriptId) {
                  existing.deploymentStatus = requestStatus === 'provisioned' || requestStatus === 'active' ? 'Paid - deployment metadata missing' : requestStatus || 'payment_confirmed';
                  existing.deploymentAction = 'provision';
                }
              }
              continue;
            }

            // No Businesses row yet: still surface the paid customer.
            // Multiple successful upgrade requests can belong to the same
            // customer. Keep ONE deployment card and prefer the request that
            // is furthest through the provisioning lifecycle.
            const customerKey = email || workspaceEmail;
            const statusPriority = {
              'provisioned': 5,
              'active': 5,
              'provisioning': 4,
              'payment_confirmed': 3,
              'provisioning_failed': 2,
              'pending_payment': 1,
              'pending': 1
            };
            // A paid customer may have no Businesses row yet, but provisioning
            // can already have created a complete Clients registry record. Reuse
            // that deployment metadata instead of rendering a blank synthetic card.
            const paidDeployment = deploymentByEmail[email] || deploymentByEmail[workspaceEmail] || {};
            const candidate = {
              businessId: String(paidDeployment.Business_ID || '').trim(),
              businessName: (u.businessName >= 0 ? String(upgradeData[ui][u.businessName] || '').trim() : '') || (u.name >= 0 ? String(upgradeData[ui][u.name] || '').trim() : '') || String(paidDeployment.Client_Name || '').trim() || 'Paid Customer',
              ownerEmail: email || workspaceEmail || String(paidDeployment.Email || '').trim().toLowerCase(),
              subscriptionTier: u.tier >= 0 ? upgradeData[ui][u.tier] || paidDeployment.Tier || 'sovereign' : paidDeployment.Tier || 'sovereign',
              verificationStatus: 'paid',
              createdAt: u.createdAt >= 0 ? upgradeData[ui][u.createdAt] : (paidDeployment.Created_At || ''),
              status: 'paid',
              clientId: paidDeployment.Client_ID || '',
              scriptId: paidDeployment.Script_ID || '',
              deploymentId: paidDeployment.Deployment_ID || '',
              launchUrl: paidDeployment.Web_App_URL || '',
              webAppUrl: paidDeployment.Web_App_URL || '',
              landingUrl: paidDeployment.Landing_URL || '',
              deploymentStatus: paidDeployment.Status || requestStatus || 'payment_confirmed',
              deployedAt: paidDeployment.Updated_At || paidDeployment.Created_At || '',
              deploymentVersion: paidDeployment.Provisioning_Version || '',
              workspaceId: paidDeployment.Workspace_ID || paidDeployment.Sheet_ID || '',
              upgradeRequestId: requestId,
              paymentStatus: 'success',
              deploymentAction: (paidDeployment.Client_ID && paidDeployment.Script_ID && paidDeployment.Deployment_ID && paidDeployment.Web_App_URL) ? 'redeploy' : 'provision'
            };

            if (customerKey && paidCustomerByEmail[customerKey]) {
              const currentPaid = paidCustomerByEmail[customerKey];
              const currentPriority = statusPriority[String(currentPaid.deploymentStatus || '').toLowerCase()] || 0;
              const candidatePriority = statusPriority[requestStatus] || 0;

              // Keep the more advanced request. If both have the same status,
              // keep the newer request so the card always points to the latest
              // confirmed request for that customer.
              const candidateCreated = candidate.createdAt ? new Date(candidate.createdAt).getTime() : 0;
              const currentCreated = currentPaid.createdAt ? new Date(currentPaid.createdAt).getTime() : 0;
              if (candidatePriority > currentPriority ||
                  (candidatePriority === currentPriority && candidateCreated >= currentCreated)) {
                Object.assign(currentPaid, candidate);
              }
            } else {
              businesses.push(candidate);
              if (customerKey) paidCustomerByEmail[customerKey] = candidate;
            }
          }
        }
      }
    } catch (paymentJoinError) {
      console.error('Unable to join successful payments into admin deployment list:', paymentJoinError);
    }

    return businesses;
  } catch (error) {
    console.error('Error getting businesses:', error);
    return [];
  }
}

function adminProvisionPaidRequest(requestId, sessionId) {
  requireAdminSession_(sessionId);
  if (!requestId) return {success:false,code:'REQUEST_ID_REQUIRED',message:'Upgrade request ID is required.'};

  try {
    var request = getUpgradeRequest(requestId);
    if (!request) return {success:false,code:'REQUEST_NOT_FOUND',message:'Upgrade request not found.'};

    var ss = getBizOSMasterSpreadsheet_();
    var paymentSheet = ss.getSheetByName('Payments');
    if (!paymentSheet) return {success:false,code:'PAYMENTS_SHEET_MISSING',message:'Payments sheet not found.'};

    var pdata = paymentSheet.getDataRange().getValues();
    var ph = pdata[0] || [];
    var reqCol = ph.indexOf('Request_ID');
    var statusCol = ph.indexOf('Status');
    if (reqCol < 0 || statusCol < 0) return {success:false,code:'PAYMENTS_SCHEMA_INVALID',message:'Payments sheet is missing required columns.'};

    var paid = false;
    for (var i = 1; i < pdata.length; i++) {
      if (String(pdata[i][reqCol] || '').trim() === String(requestId).trim() &&
          String(pdata[i][statusCol] || '').trim().toLowerCase() === 'success') {
        paid = true;
        break;
      }
    }
    if (!paid) return {success:false,code:'PAYMENT_NOT_SUCCESSFUL',message:'No successful payment is recorded for this upgrade request.'};

    // If the request says "provisioned" but there is no client deployment
    // record, treat it as a recoverable paid request. This covers historical
    // browser/provisioning failures without charging the customer again.
    var requestStatus = String(request.status || '').trim().toLowerCase();
    if (requestStatus === 'provisioned' || requestStatus === 'active') {
      var existingRecoveryClient = getExistingActiveClientDeployment_(request.workspaceEmail || request.email, request.businessName);
      if (!existingRecoveryClient) {
        var recoverySheet = ss.getSheetByName('Upgrade_Requests');
        var recoveryData = recoverySheet.getDataRange().getValues();
        var recoveryHeaders = recoveryData[0] || [];
        var recoveryReqCol = recoveryHeaders.indexOf('Request_ID');
        var recoveryStatusCol = recoveryHeaders.indexOf('Status');
        var recoveryUpdatedCol = recoveryHeaders.indexOf('Updated_At');
        for (var ri = 1; ri < recoveryData.length; ri++) {
          if (String(recoveryData[ri][recoveryReqCol] || '').trim() === String(requestId).trim()) {
            if (recoveryStatusCol >= 0) recoverySheet.getRange(ri + 1, recoveryStatusCol + 1).setValue('payment_confirmed');
            if (recoveryUpdatedCol >= 0) recoverySheet.getRange(ri + 1, recoveryUpdatedCol + 1).setValue(new Date().toISOString());
            break;
          }
        }
      }
    }

    // A confirmed payment can legitimately have an older request status
    // (for example pending/pending_payment) when browser provisioning failed.
    // Reconcile this exact request before calling the normal idempotent
    // provisioning bridge. This never charges the customer again.
    var upgradeSheet = ss.getSheetByName('Upgrade_Requests');
    if (!upgradeSheet) return {success:false,code:'UPGRADE_SHEET_MISSING',message:'Upgrade request sheet not found.'};

    var udata = upgradeSheet.getDataRange().getValues();
    var uh = udata[0] || [];
    var uReq = uh.indexOf('Request_ID');
    var uStatus = uh.indexOf('Status');
    var uUpdated = uh.indexOf('Updated_At');
    if (uReq < 0 || uStatus < 0) {
      return {success:false,code:'UPGRADE_SCHEMA_INVALID',message:'Upgrade request sheet is missing required columns.'};
    }

    for (var ui = 1; ui < udata.length; ui++) {
      if (String(udata[ui][uReq] || '').trim() !== String(requestId).trim()) continue;
      var currentStatus = String(udata[ui][uStatus] || '').trim().toLowerCase();
      if (currentStatus === 'pending' || currentStatus === 'pending_payment' || currentStatus === 'payment_pending') {
        upgradeSheet.getRange(ui + 1, uStatus + 1).setValue('payment_confirmed');
        if (uUpdated >= 0) upgradeSheet.getRange(ui + 1, uUpdated + 1).setValue(new Date().toISOString());
      }
      break;
    }

    return provisionConfirmedUpgradeRequest(requestId);
  } catch (error) {
    console.error('adminProvisionPaidRequest error:', error);
    return {success:false,code:'ADMIN_PROVISION_ERROR',message:error && error.message ? error.message : 'Could not continue client provisioning.'};
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
    let featuresSheet = getBizOSMasterSpreadsheet_().getSheetByName('Features');
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
    let featuresSheet = getBizOSMasterSpreadsheet_().getSheetByName('Features');
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
    const featuresSheet = getBizOSMasterSpreadsheet_().getSheetByName('Features');
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
    const featuresSheet = getBizOSMasterSpreadsheet_().getSheetByName('Features');
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
  return validateAdminSession(sessionId);
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
    const featuresSheet = getBizOSMasterSpreadsheet_().getSheetByName('Features');
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


function getAdminClientRegistryRecord_(identifier) {
  var wanted = String(identifier || '').trim();
  if (!wanted) return null;

  var sheet = getBizOSMasterSpreadsheet_().getSheetByName('Clients');
  if (!sheet || sheet.getLastRow() < 2) return null;

  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function(h) { return String(h || '').trim(); });
  var clientIdCol = headers.indexOf('Client_ID');
  var businessIdCol = headers.indexOf('Business_ID');

  if (clientIdCol < 0 && businessIdCol < 0) return null;

  for (var i = 1; i < values.length; i++) {
    var clientId = clientIdCol >= 0 ? String(values[i][clientIdCol] || '').trim() : '';
    var businessId = businessIdCol >= 0 ? String(values[i][businessIdCol] || '').trim() : '';
    if (clientId !== wanted && businessId !== wanted) continue;

    var record = {};
    headers.forEach(function(h, j) { record[h] = values[i][j]; });
    return record;
  }

  return null;
}

function getAdminBusinessPeople_(businessId) {
  var wanted = String(businessId || '').trim();
  if (!wanted) return [];

  var sheet = getOrCreateUserSheet();
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  var headers = values[0].map(function(h) { return String(h || '').trim(); });
  var businessIdCol = headers.indexOf('Business_ID');
  var emailCol = headers.indexOf('Email');
  if (businessIdCol < 0 || emailCol < 0) return [];

  var people = [];
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][businessIdCol] || '').trim() !== wanted) continue;

    var person = {};
    headers.forEach(function(h, j) { person[h] = values[i][j]; });
    people.push({
      email: String(person.Email || ''),
      name: String(person.Name || ''),
      role: String(person.Role || 'staff'),
      status: String(person.Invitation_Status || (String(person.Is_Verified || '').toUpperCase() === 'YES' ? 'active' : 'pending')),
      verified: String(person.Is_Verified || '').toUpperCase() === 'YES',
      assignedModules: String(person.Assigned_Modules || '').split(',').map(function(x) {
        return String(x || '').trim();
      }).filter(Boolean),
      createdAt: person.Created_At || ''
    });
  }

  return people;
}

function getAdminBusinessModules_(businessId, tier) {
  var config = getConfig();
  var key = String(tier || 'starter').trim().toLowerCase();
  var access = config.MODULES && config.MODULES.access
    ? (config.MODULES.access[key] || config.MODULES.access.starter || [])
    : [];
  var all = config.MODULES && Array.isArray(config.MODULES.all) ? config.MODULES.all : [];
  var icons = config.MODULES && config.MODULES.icons ? config.MODULES.icons : {};
  var labels = config.MODULES && config.MODULES.labels ? config.MODULES.labels : {};
  var enabled = {};
  access.forEach(function(name) { enabled[String(name)] = true; });

  return all.map(function(name) {
    return {
      name: name,
      label: labels[name] || name,
      icon: icons[name] || 'fa-folder',
      enabled: !!enabled[name],
      businessId: String(businessId || ''),
      tier: key
    };
  });
}

function getClientManagementOverview(clientId, sessionId) {
  requireAdminSession_(sessionId);
  try {
    var wanted = String(clientId || '').trim();
    var client = getAdminClientRegistryRecord_(wanted);
    if (!client) return {success:false, code:'CLIENT_NOT_FOUND', message:'No client was found for this client or business ID.'};

    var businessId = String(client.Business_ID || '').trim();
    var tier = String(client.Tier || 'starter').trim().toLowerCase();
    var workspace = {success:false};
    var billing = {maintenance:[],features:[],quotes:[],services:[],activity:[]};
    var payments = [];

    // Business is the primary admin grouping. People, modules and deployment
    // metadata are all returned together so the UI does not need to guess
    // relationships from separate lists.
    var people = getAdminBusinessPeople_(businessId);
    var modules = getAdminBusinessModules_(businessId, tier);

    var featureProfile = {enabledFeatureIds:[]};
    var featureProfileRaw = String(client.Feature_Profile_JSON || '').trim();
    if (featureProfileRaw) {
      try {
        featureProfile = normalizeClientFeatureProfile_(JSON.parse(featureProfileRaw) || {});
      } catch (profileError) {
        featureProfile = {enabledFeatureIds:[]};
      }
    }

    try {
      var masterPayments = getBizOSMasterSpreadsheet_().getSheetByName('Payments');
      if (masterPayments && masterPayments.getLastRow() >= 2) {
        var pv = masterPayments.getDataRange().getValues();
        var ph = pv[0].map(function(h){return String(h || '').trim();});
        var pBiz = ph.indexOf('Business_ID');
        var pReq = ph.indexOf('Request_ID');
        var pStatus = ph.indexOf('Status');
        var pAmount = ph.indexOf('Amount');
        var pCurrency = ph.indexOf('Currency');
        var pRef = ph.indexOf('Transaction_Ref');
        var pCreated = ph.indexOf('Created_At');
        for (var pi = 1; pi < pv.length; pi++) {
          var rowBusiness = pBiz >= 0 ? String(pv[pi][pBiz] || '').trim() : '';
          if (businessId && rowBusiness && rowBusiness !== businessId) continue;
          var payment = {};
          ph.forEach(function(h,j){payment[h]=pv[pi][j];});
          payments.push(payment);
        }
      }
    } catch (paymentError) {
      console.error('Client management payment lookup error:', paymentError);
    }

    try {
      var sheetId = String(client.Sheet_ID || client.Workspace_ID || '').trim();
      if (sheetId) {
        var ws = SpreadsheetApp.openById(sheetId);
        workspace = {
          success:true,
          sheetId:sheetId,
          url:ws.getUrl(),
          sheets:ws.getSheets().map(function(s){return s.getName();})
        };

        function rows(name) {
          var s=ws.getSheetByName(name);
          if(!s||s.getLastRow()<2)return[];
          var v=s.getDataRange().getValues(),h=v[0]||[];
          return v.slice(1).map(function(r){
            var o={};
            h.forEach(function(k,j){o[String(k||'')]=r[j];});
            return o;
          }).slice(-50).reverse();
        }

        billing.maintenance=rows('Client_Maintenance_Requests');
        billing.features=rows('Client_Feature_Requests');
        billing.quotes=rows('Client_Quotes');
        billing.services=rows('Client_Service_History');
        billing.activity=rows('Client_Activity');
      }
    } catch(workspaceError) {
      workspace={success:false,message:'Client workspace could not be read by the admin service.'};
    }

    return {
      success:true,
      client:{
        clientId:String(client.Client_ID || ''),
        email:String(client.Email || ''),
        clientName:String(client.Client_Name || ''),
        businessId:businessId,
        tier:tier,
        status:String(client.Status || ''),
        scriptId:String(client.Script_ID || ''),
        deploymentId:String(client.Deployment_ID || ''),
        webAppUrl:String(client.Web_App_URL || ''),
        landingUrl:String(client.Landing_URL || ''),
        createdAt:client.Created_At || '',
        updatedAt:client.Updated_At || '',
        provisioningVersion:String(client.Provisioning_Version || ''),
        recoveryBackupId:String(client.Recovery_Backup_ID || ''),
        recoveryBackupUrl:String(client.Recovery_Backup_URL || ''),
        ownershipStatus:String(client.Workspace_Ownership_Status || '')
      },
      people:people,
      modules:modules,
      featureProfile:featureProfile,
      workspace:workspace,
      billing:billing,
      payments:payments
    };
  } catch(error) {
    console.error('getClientManagementOverview error:',error);
    return {success:false,message:'We could not load the client management overview.'};
  }
}

function createClientAdminRecoveryBackup(clientId, sessionId) {
  requireAdminSession_(sessionId);
  try {
    var client=getActiveClientByIdForAdmin_(clientId);
    if(!client)return{success:false,message:'Active client deployment not found.'};
    var ws=SpreadsheetApp.openById(String(client.sheetId||client.Sheet_ID||''));
    var result=createClientRecoveryBackup_(ws,String(client.clientName||client.Client_Name||''),String(client.businessId||client.Business_ID||''));
    if(!result||!result.success)return result||{success:false,message:'Backup could not be created.'};
    return result;
  }catch(error){console.error('createClientAdminRecoveryBackup error:',error);return{success:false,message:'BizOS could not create the client backup.'};}
}

function archiveClientDeployment(clientId, sessionId) {
  requireAdminSession_(sessionId);
  try {
    var wanted=String(clientId||'').trim(),sheet=getBizOSMasterSpreadsheet_(),s=sheet.getSheetByName('Clients');
    if(!s)return{success:false,message:'Client registry is unavailable.'};
    var v=s.getDataRange().getValues(),h=v[0].map(function(x){return String(x||'').trim();}),idCol=h.indexOf('Client_ID'),statusCol=h.indexOf('Status'),updatedCol=h.indexOf('Updated_At');
    for(var i=1;i<v.length;i++)if(String(v[i][idCol]||'').trim()===wanted){s.getRange(i+1,statusCol+1).setValue('archived');if(updatedCol>=0)s.getRange(i+1,updatedCol+1).setValue(new Date().toISOString());return{success:true,message:'Client deployment archived.'};}
    return{success:false,message:'Client was not found.'};
  }catch(error){console.error('archiveClientDeployment error:',error);return{success:false,message:'Client could not be archived.'};}
}

function restoreArchivedClientDeployment(clientId, sessionId) {
  requireAdminSession_(sessionId);
  try {
    var wanted=String(clientId||'').trim(),sheet=getBizOSMasterSpreadsheet_(),s=sheet.getSheetByName('Clients');
    if(!s)return{success:false,message:'Client registry is unavailable.'};
    var v=s.getDataRange().getValues(),h=v[0].map(function(x){return String(x||'').trim();}),idCol=h.indexOf('Client_ID'),statusCol=h.indexOf('Status'),updatedCol=h.indexOf('Updated_At');
    for(var i=1;i<v.length;i++)if(String(v[i][idCol]||'').trim()===wanted){var current=String(v[i][statusCol]||'').toLowerCase();if(current!=='archived')return{success:false,message:'Only archived clients can be restored.'};s.getRange(i+1,statusCol+1).setValue('active');if(updatedCol>=0)s.getRange(i+1,updatedCol+1).setValue(new Date().toISOString());return{success:true,message:'Client deployment restored.'};}
    return{success:false,message:'Client was not found.'};
  }catch(error){console.error('restoreArchivedClientDeployment error:',error);return{success:false,message:'Client could not be restored.'};}
}

function getActiveClientByIdForAdmin_(clientId) {
  var wanted=String(clientId||'').trim();if(!wanted)return null;
  var s=getBizOSMasterSpreadsheet_().getSheetByName('Clients');if(!s)return null;
  var v=s.getDataRange().getValues();if(v.length<2)return null;
  var h=v[0].map(function(x){return String(x||'').trim();}),idCol=h.indexOf('Client_ID'),statusCol=h.indexOf('Status');if(idCol<0)return null;
  for(var i=1;i<v.length;i++){if(String(v[i][idCol]||'').trim()!==wanted)continue;var o={};h.forEach(function(k,j){o[k]=v[i][j];});if(statusCol<0||String(v[i][statusCol]||'').toLowerCase()==='active')return o;}
  return null;
}


function listClientRecoveryBackups(clientId, sessionId) {
  requireAdminSession_(sessionId);
  try {
    var client=getActiveClientByIdForAdmin_(clientId);
    if(!client)return{success:false,message:'Active client deployment not found.'};
    var rootIt=DriveApp.getFoldersByName('BizOS_Client_Recovery');
    if(!rootIt.hasNext())return{success:true,backups:[]};
    var root=rootIt.next(),safeName=String(client.Client_Name||'Client').replace(/[^a-zA-Z0-9]/g,'_')+'_'+String(client.Business_ID||'').replace(/[^a-zA-Z0-9]/g,'_'),folders=root.getFoldersByName(safeName);
    if(!folders.hasNext())return{success:true,backups:[]};
    var folder=folders.next(),files=folder.getFiles(),out=[];
    while(files.hasNext()){var file=files.next();out.push({id:file.getId(),name:file.getName(),createdAt:file.getDateCreated().toISOString(),url:file.getUrl(),mimeType:file.getMimeType()});}
    out.sort(function(a,b){return new Date(b.createdAt)-new Date(a.createdAt);});
    return{success:true,backups:out};
  }catch(error){console.error('listClientRecoveryBackups error:',error);return{success:false,message:'Client recovery backups could not be loaded.'};}
}

function restoreClientWorkspaceFromRecoveryBackup(clientId, backupId, sessionId) {
  requireAdminSession_(sessionId);
  try {
    var client=getActiveClientByIdForAdmin_(clientId);
    if(!client)return{success:false,message:'Active client deployment not found.'};
    var targetId=String(client.sheetId||client.Sheet_ID||''),sourceId=String(backupId||'').trim();
    if(!targetId||!sourceId)return{success:false,message:'Client workspace and backup are required.'};
    if(targetId===sourceId)return{success:false,message:'The selected backup is already the live workspace.'};
    var target=SpreadsheetApp.openById(targetId),source=SpreadsheetApp.openById(sourceId);
    var safety=createClientRecoveryBackup_(target,String(client.clientName||client.Client_Name||''),String(client.businessId||client.Business_ID||''));
    if(!safety||!safety.success)return{success:false,message:'A safety backup could not be created, so the restore was not started.'};
    var sourceSheets=source.getSheets(),sourceNames={};
    sourceSheets.forEach(function(sh){sourceNames[sh.getName()]=true;});
    var targetSheets=target.getSheets();
    targetSheets.forEach(function(sh){
      var name=sh.getName();
      if(!sourceNames[name]){if(target.getSheets().length>1)target.deleteSheet(sh);}
    });
    sourceSheets.forEach(function(src){
      var name=src.getName(),existing=target.getSheetByName(name);
      if(existing){
        var temp=src.copyTo(target);
        target.deleteSheet(existing);
        temp.setName(name);
      }else{
        var copy=src.copyTo(target);copy.setName(name);
      }
    });
    return{success:true,message:'Client workspace restored from the selected recovery backup.',safetyBackup:safety,restoredFrom:sourceId};
  }catch(error){console.error('restoreClientWorkspaceFromRecoveryBackup error:',error);return{success:false,message:'The client workspace could not be restored. The current workspace was not intentionally replaced unless the restore completed.'};}
}

function createClientQuote(clientId,data,sessionId){
  requireAdminSession_(sessionId);
  try{
    var client=getActiveClientByIdForAdmin_(clientId);if(!client)return{success:false,message:'Active client deployment not found.'};
    data=data||{};var title=String(data.title||'').trim(),details=String(data.details||'').trim(),amount=Number(data.amount||0),currency=String(data.currency||'NGN').toUpperCase();
    if(!title)return{success:false,message:'Quote title is required.'};
    if(!amount||amount<=0)return{success:false,message:'Quote amount must be greater than zero.'};
    if(details.length>4000)return{success:false,message:'Quote details are too long.'};
    var ws=SpreadsheetApp.openById(String(client.Sheet_ID||client.sheetId||'')),s=ws.getSheetByName('Client_Quotes');
    if(!s){s=ws.insertSheet('Client_Quotes');s.appendRow(['Quote_ID','Title','Details','Amount','Currency','Status','Created_At','Updated_At','Payment_Ref','Payment_Status','Paid_At']);}
    var h=s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(function(x){return String(x||'');});
    ['Payment_Ref','Payment_Status','Paid_At'].forEach(function(k){if(h.indexOf(k)<0)s.getRange(1,s.getLastColumn()+1).setValue(k);});
    var id='QUO-'+Utilities.getUuid().replace(/-/g,'').slice(0,10).toUpperCase(),now=new Date().toISOString();
    s.appendRow([id,title,details,amount,currency,'pending',now,now,'','pending','']);
    return{success:true,quoteId:id,message:'Quote created for the client.'};
  }catch(error){console.error('createClientQuote error:',error);return{success:false,message:'Quote could not be created.'};}
}
function updateClientQuoteStatus(clientId,quoteId,status,sessionId){
  requireAdminSession_(sessionId);
  try{
    var client=getActiveClientByIdForAdmin_(clientId);if(!client)return{success:false,message:'Active client deployment not found.'};
    var allowed=['pending','approved','paid','cancelled','expired'];status=String(status||'').toLowerCase();if(allowed.indexOf(status)<0)return{success:false,message:'Invalid quote status.'};
    var ws=SpreadsheetApp.openById(String(client.Sheet_ID||client.sheetId||'')),s=ws.getSheetByName('Client_Quotes');if(!s)return{success:false,message:'No quotes found.'};
    var v=s.getDataRange().getValues(),h=v[0].map(function(x){return String(x||'');}),idc=h.indexOf('Quote_ID'),sc=h.indexOf('Status'),uc=h.indexOf('Updated_At');for(var i=1;i<v.length;i++)if(String(v[i][idc]||'')===String(quoteId||'')){if(sc>=0)s.getRange(i+1,sc+1).setValue(status);if(uc>=0)s.getRange(i+1,uc+1).setValue(new Date().toISOString());return{success:true,message:'Quote status updated.'};}
    return{success:false,message:'Quote not found.'};
  }catch(error){console.error('updateClientQuoteStatus error:',error);return{success:false,message:'Quote status could not be updated.'};}
}
