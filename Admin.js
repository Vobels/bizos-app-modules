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
        businessSheet.getRange(bizRowIndex + 1, tierCol + 1).setValue('sovereign');
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
    // The Admin dashboard and Business Directory must use the same canonical
    // business view. That view also surfaces confirmed paid customers whose
    // Businesses row has not yet been created.
    const masterSpreadsheet = getBizOSMasterSpreadsheet_();
    const userSheet = masterSpreadsheet.getSheetByName('Users') || masterSpreadsheet.getSheetByName('User');
    const userData = userSheet ? userSheet.getDataRange().getValues() : [[]];
    const businesses = getAllBusinesses(sessionId);

    const totalBusinesses = businesses.length;
    const totalUsers = Math.max(0, userData.length - 1);

    // Count tiers from the canonical business view so paid deployment records
    // remain visible even when provisioning has not populated Businesses yet.
    let freeUsers = 0, paidUsers = 0;
    businesses.forEach(function(business) {
      const tier = String(business.subscriptionTier || 'free').trim().toLowerCase();
      if (tier === 'free' || tier === 'starter') freeUsers++;
      else paidUsers++;
    });

    // Keep recent/new-business activity tied to the canonical Businesses
    // registry; it should not manufacture activity for payment-only records.
    const businessSheet = masterSpreadsheet.getSheetByName('Businesses');
    const bizData = businessSheet ? businessSheet.getDataRange().getValues() : [[]];
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
        status: data[i][headers.indexOf('Subscription_Status')] || data[i][headers.indexOf('Status')] || 'active',
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
    const businessNameById = {};

    try {
      const businessSheet = getBizOSMasterSpreadsheet_().getSheetByName('Businesses');
      if (businessSheet && businessSheet.getLastRow() >= 2) {
        const bv = businessSheet.getDataRange().getValues();
        const bh = bv[0].map(function(h){ return String(h || '').trim(); });
        const idCol = bh.indexOf('Business_ID');
        const nameCol = bh.indexOf('Business_Name');
        if (idCol >= 0 && nameCol >= 0) {
          for (let i = 1; i < bv.length; i++) {
            var bid = String(bv[i][idCol] || '').trim();
            if (bid) businessNameById[bid] = String(bv[i][nameCol] || '');
          }
        }
      }
    } catch (businessError) {
      console.error('Unable to load business names for admin users:', businessError);
    }

    const users = [];
    const emailCol = headers.indexOf('Email');
    const nameCol = headers.indexOf('Name');
    const roleCol = headers.indexOf('Role');
    const businessIdCol = headers.indexOf('Business_ID');
    const verifiedCol = headers.indexOf('Is_Verified');
    const createdCol = headers.indexOf('Created_At');

    for (let i = 1; i < data.length; i++) {
      const businessId = businessIdCol >= 0 ? String(data[i][businessIdCol] || '').trim() : '';
      users.push({
        email: emailCol >= 0 ? data[i][emailCol] : '',
        name: nameCol >= 0 ? data[i][nameCol] : '',
        role: roleCol >= 0 ? data[i][roleCol] : '',
        businessId: businessId,
        businessName: businessNameById[businessId] || '',
        isVerified: verifiedCol >= 0 ? data[i][verifiedCol] : '',
        createdAt: createdCol >= 0 ? data[i][createdCol] : ''
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

function getAdminModuleManagementData(sessionId) {
  requireAdminSession_(sessionId);
  try {
    var config = getConfig();
    var all = config.MODULES && Array.isArray(config.MODULES.all) ? config.MODULES.all : [];
    var labels = config.MODULES && config.MODULES.labels ? config.MODULES.labels : {};
    var icons = config.MODULES && config.MODULES.icons ? config.MODULES.icons : {};
    var access = config.MODULES && config.MODULES.access ? config.MODULES.access : {};

    // Module Management needs the effective entitlement/connection state,
    // not a fragile copy of one Businesses column. Keep this resolution local
    // so other Admin business views retain their existing contract.
    var businesses = getAllBusinesses(sessionId).map(function(b) {
      var rawTier = String(
        b.subscriptionTier ||
        b.Subscription_Tier ||
        b.tier ||
        b.Tier ||
        ''
      ).trim().toLowerCase();

      var tier = normalizeSubscriptionTier(rawTier);
      if (['tier2','professional','enterprise','paid','premium','sovereign'].indexOf(rawTier) >= 0) {
        tier = 'sovereign';
      }

      var rawStatus = String(
        b.status ||
        b.subscriptionStatus ||
        b.Subscription_Status ||
        b.deploymentStatus ||
        ''
      ).trim().toLowerCase();

      var isInactive = ['archived','suspended','inactive','disabled'].indexOf(rawStatus) >= 0;
      var activeFlag = String(
        b.isActive ||
        b.Is_Active ||
        ''
      ).trim().toLowerCase();

      if (activeFlag === 'no' || activeFlag === 'false' || activeFlag === '0') {
        isInactive = true;
      }

      return {
        businessId: String(b.businessId || b.Business_ID || '').trim(),
        businessName: String(b.businessName || b.Business_Name || '').trim(),
        ownerEmail: String(b.ownerEmail || b.Email || '').trim().toLowerCase(),
        tier: tier,
        status: rawStatus || 'active',
        inactive: isInactive,
        workspaceId: String(
          b.workspaceId ||
          b.Workspace_ID ||
          b.sheetId ||
          b.Sheet_ID ||
          ''
        ).trim(),
        // Preserve deployment/payment evidence through this local projection.
        // getAllBusinesses() can legitimately supply a paid customer through
        // the Clients/Payments merge even when the Businesses row has an old
        // or missing tier value.
        clientId: String(b.clientId || b.Client_ID || '').trim(),
        deploymentId: String(b.deploymentId || b.Deployment_ID || '').trim(),
        paymentStatus: String(b.paymentStatus || b.Payment_Status || '').trim().toLowerCase(),
        deploymentAction: String(b.deploymentAction || '').trim().toLowerCase()
      };
    }).filter(function(b) {
      return !b.inactive && !!(b.businessId || b.businessName || b.ownerEmail);
    });

    return all.map(function(name) {
      var connected = businesses.filter(function(b) {
        // Resolve the effective module entitlement from the same canonical tier
        // rules used everywhere else in BizOS. Do not depend on a raw access
        // object lookup alone: paid/deployed customers can arrive through the
        // merged Clients/Payments view with legacy tier labels.
        var allowed = access[b.tier] || [];
        // CONFIG.MODULES.access is the canonical module entitlement map.
        // Do not call a helper that is not part of the Admin runtime.

        // A confirmed paid deployment is an operational sovereign workspace.
        // This fallback protects historical records whose Businesses row still
        // carries a legacy/free tier even though payment and deployment exist.
        var hasConfirmedPaidDeployment =
          String(b.paymentStatus || '').toLowerCase() === 'success' ||
          !!String(b.deploymentId || '').trim() ||
          !!String(b.clientId || '').trim();

        if (hasConfirmedPaidDeployment &&
            normalizeSubscriptionTier(b.tier) === 'free') {
          allowed = access.sovereign || [];
        }

        return allowed.indexOf(name) >= 0;
      });

      var recordCount = 0;
      var recordCountAvailable = true;
      var connectedDetails = connected.slice(0, 50).map(function(b) {
        var count = null;
        if (b.workspaceId) {
          try {
            var workspace = SpreadsheetApp.openById(b.workspaceId);
            // Use the canonical module definition shared by the module runtime.
            // This prevents Admin record counts from drifting from the actual
            // workspace sheet names used by BizOS modules.
            var definition = getBizOSModuleDefinition_(name);
            var sheet = definition ? workspace.getSheetByName(definition.sheet) : null;
            count = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
            recordCount += count;
          } catch (workspaceError) {
            recordCountAvailable = false;
          }
        } else {
          recordCountAvailable = false;
        }
        b.recordCount = count;
        return b;
      });

      return {
        name: name,
        label: labels[name] || name,
        icon: icons[name] || 'fa-folder',
        tier: 'mixed',
        recordCount: recordCountAvailable ? recordCount : null,
        recordCountAvailable: recordCountAvailable,
        visible: true,
        businessCount: connected.length,
        businesses: connectedDetails
      };
    });
  } catch(error) {
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
    modules: CONFIG.MODULES,
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
    const rawTargetTier = String(targetTier || 'free').trim().toLowerCase();
    const targetTierNormalized = normalizeSubscriptionTier(rawTargetTier);
    const validLegacyInputs = ['free', 'starter', 'sovereign', 'tier2', 'professional', 'enterprise'];
    if (!validLegacyInputs.includes(rawTargetTier)) {
      return { success: false, message: "Invalid tier. Use Free or Sovereign." };
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
    businessSheet.getRange(bizRowIndex + 1, tierCol + 1).setValue(targetTierNormalized);
    businessSheet.getRange(bizRowIndex + 1, statusCol + 1).setValue('verified');
    
    // Create workspace if upgrading from free and none exists
    const currentWorkspaceId = businessData[bizRowIndex][workspaceIdCol];
    if (!currentWorkspaceId && targetTierNormalized !== 'free') {
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
      tier: targetTierNormalized,
      businessId: businessId,
      performedBy: Session.getActiveUser().getEmail()
    });
    
    return {
      success: true,
      message: `User ${email} updated to ${targetTierNormalized} successfully`
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

function auditAdminBusinessRelationships(sessionId) {
  requireAdminSession_(sessionId);
  var report = {
    success:true,
    checkedAt:new Date().toISOString(),
    summary:{businesses:0,users:0,clients:0,staffModuleRows:0,issues:0},
    issues:[],
    safeRepairsAvailable:[]
  };

  function addIssue(type, key, message, details) {
    report.issues.push({type:type,key:key,message:message,details:details||{}});
  }

  try {
    var master=getBizOSMasterSpreadsheet_();
    var businesses=master.getSheetByName('Businesses');
    var users=master.getSheetByName('Users') || master.getSheetByName('User');
    var clients=master.getSheetByName('Clients');
    var staffModules=master.getSheetByName('Staff_Modules');

    var businessMap={}, businessNames={};
    if (businesses && businesses.getLastRow()>=2) {
      var bv=businesses.getDataRange().getValues(), bh=bv[0].map(function(x){return String(x||'').trim();});
      var bid=bh.indexOf('Business_ID'), bn=bh.indexOf('Business_Name');
      report.summary.businesses=bv.length-1;
      if(bid<0) addIssue('schema','Businesses','Business_ID column is missing.');
      for(var bi=1;bi<bv.length;bi++){
        var id=bid>=0?String(bv[bi][bid]||'').trim():'';
        if(!id){addIssue('business_missing_id','row_'+(bi+1),'Business has no Business_ID.');continue;}
        if(businessMap[id]) addIssue('business_duplicate_id',id,'Duplicate Business_ID.',{rows:[businessMap[id].row,bi+1]});
        businessMap[id]={row:bi+1,status:String(bv[bi][bh.indexOf('Status')]||'').trim()};
        businessNames[id]=bn>=0?String(bv[bi][bn]||'').trim():'';
      }
    } else addIssue('schema','Businesses','Businesses sheet is missing.');

    if(users && users.getLastRow()>=2){
      var uv=users.getDataRange().getValues(), uh=uv[0].map(function(x){return String(x||'').trim();});
      var ub=uh.indexOf('Business_ID'), ue=uh.indexOf('Email'), un=uh.indexOf('Business_Name'), userKeys={};
      report.summary.users=uv.length-1;
      if(ub<0) addIssue('schema','Users','Business_ID column is missing.');
      for(var ui=1;ui<uv.length;ui++){
        var userBusiness=ub>=0?String(uv[ui][ub]||'').trim():'', email=ue>=0?String(uv[ui][ue]||'').trim().toLowerCase():'';
        var key=email+'|'+userBusiness;
        if(email&&userKeys[key]) addIssue('user_duplicate_membership',key,'Duplicate user membership for the same business.',{rows:[userKeys[key],ui+1]});
        if(email) userKeys[key]=ui+1;
        if(!userBusiness) addIssue('user_orphaned',email||('row_'+(ui+1)),'User/staff has no Business_ID.',{row:ui+1});
        else if(!businessMap[userBusiness]) addIssue('user_unknown_business',email||('row_'+(ui+1)),'User/staff references a Business_ID that does not exist.',{businessId:userBusiness,row:ui+1});
        else if(un>=0 && String(uv[ui][un]||'').trim()!==businessNames[userBusiness])
          addIssue('user_business_name_mismatch',email||('row_'+(ui+1)),'User Business_Name does not match the canonical business name.',{businessId:userBusiness,row:ui+1});
      }
    } else addIssue('schema','Users','Users sheet is missing or empty.');

    var activeClientByBusiness={};
    if(clients && clients.getLastRow()>=2){
      var cv=clients.getDataRange().getValues(), ch=cv[0].map(function(x){return String(x||'').trim();});
      var cb=ch.indexOf('Business_ID'), cc=ch.indexOf('Client_ID'), cs=ch.indexOf('Status'), ce=ch.indexOf('Email'), clientIds={};
      report.summary.clients=cv.length-1;
      if(cb<0) addIssue('schema','Clients','Business_ID column is missing.');
      for(var ci=1;ci<cv.length;ci++){
        var cBusiness=cb>=0?String(cv[ci][cb]||'').trim():'', clientId=cc>=0?String(cv[ci][cc]||'').trim():'', cStatus=cs>=0?String(cv[ci][cs]||'').trim().toLowerCase():'';
        if(!clientId) addIssue('client_missing_id','row_'+(ci+1),'Client registry row has no Client_ID.',{row:ci+1});
        else if(clientIds[clientId]) addIssue('client_duplicate_id',clientId,'Duplicate Client_ID.',{rows:[clientIds[clientId],ci+1]});
        else clientIds[clientId]=ci+1;
        if(!cBusiness) addIssue('client_orphaned',clientId||('row_'+(ci+1)),'Client deployment has no Business_ID.',{row:ci+1});
        else if(!businessMap[cBusiness]) addIssue('client_unknown_business',clientId||('row_'+(ci+1)),'Client deployment references a Business_ID that does not exist.',{businessId:cBusiness,row:ci+1});
        if(cBusiness && cStatus==='active'){
          if(activeClientByBusiness[cBusiness]) addIssue('multiple_active_clients',cBusiness,'More than one active client deployment is linked to this business.',{rows:[activeClientByBusiness[cBusiness],ci+1]});
          else activeClientByBusiness[cBusiness]=ci+1;
        }
      }
    } else addIssue('schema','Clients','Clients registry is missing or empty.');

    if(staffModules && staffModules.getLastRow()>=2){
      var sv=staffModules.getDataRange().getValues(), sh=sv[0].map(function(x){return String(x||'').trim();});
      var sb=sh.indexOf('Business_ID'), se=sh.indexOf('Staff_Email'), sm=sh.indexOf('Module_Access');
      report.summary.staffModuleRows=sv.length-1;
      for(var si=1;si<sv.length;si++){
        var sBiz=sb>=0?String(sv[si][sb]||'').trim():'', sEmail=se>=0?String(sv[si][se]||'').trim().toLowerCase():'', module=sm>=0?String(sv[si][sm]||'').trim():'';
        if(!sBiz) addIssue('staff_module_orphaned',sEmail||('row_'+(si+1)),'Staff module assignment has no Business_ID.',{row:si+1,module:module});
        else if(!businessMap[sBiz]) addIssue('staff_module_unknown_business',sEmail||('row_'+(si+1)),'Staff module assignment references an unknown Business_ID.',{businessId:sBiz,row:si+1,module:module});
        if(!sEmail) addIssue('staff_module_missing_staff',sBiz||('row_'+(si+1)),'Staff module assignment has no Staff_Email.',{row:si+1});
      }
    }

    report.summary.issues=report.issues.length;
    report.safeRepairsAvailable=report.issues.filter(function(x){
      return x.type==='user_business_name_mismatch';
    }).map(function(x){return x.key;});
    return report;
  } catch(error) {
    console.error('auditAdminBusinessRelationships error:',error);
    return {success:false,message:'Business relationship audit could not be completed.',error:String(error&&error.message||error)};
  }
}

function repairAdminBusinessRelationships(sessionId) {
  requireAdminSession_(sessionId);
  var audit=auditAdminBusinessRelationships(sessionId);
  if(!audit.success)return audit;

  var repaired=0, master=getBizOSMasterSpreadsheet_(), users=master.getSheetByName('Users')||master.getSheetByName('User');
  if(users && users.getLastRow()>=2){
    var v=users.getDataRange().getValues(), h=v[0].map(function(x){return String(x||'').trim();});
    var b=h.indexOf('Business_ID'), n=h.indexOf('Business_Name'), businesses=master.getSheetByName('Businesses');
    if(b>=0&&n>=0&&businesses&&businesses.getLastRow()>=2){
      var bv=businesses.getDataRange().getValues(),bh=bv[0].map(function(x){return String(x||'').trim();}),bi=bh.indexOf('Business_ID'),bn=bh.indexOf('Business_Name'),names={};
      for(var i=1;i<bv.length;i++){var id=bi>=0?String(bv[i][bi]||'').trim():'';if(id)names[id]=bn>=0?String(bv[i][bn]||'').trim():'';}
      for(var ui=1;ui<v.length;ui++){var uid=String(v[ui][b]||'').trim();if(uid&&Object.prototype.hasOwnProperty.call(names,uid)&&String(v[ui][n]||'').trim()!==names[uid]){users.getRange(ui+1,n+1).setValue(names[uid]);repaired++;}}
    }
  }

  return {
    success:true,
    repaired:repaired,
    remainingIssues:audit.summary.issues-repaired,
    message:repaired ? 'Safe business-name relationships were repaired. Orphaned or ambiguous records were not guessed or reassigned.' : 'No safe relationship repairs were required.',
    audit:audit
  };
}

function getAdminBusinessPeople_(businessId) {
  var wanted = String(businessId || '').trim();
  if (!wanted) return [];

  var master = getBizOSMasterSpreadsheet_();
  var sheet = master.getSheetByName('Users') || master.getSheetByName('User') || getOrCreateUserSheet();
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
  var key = normalizeSubscriptionTier(tier);
  var access = config.MODULES && config.MODULES.access
    ? (config.MODULES.access[key] || config.MODULES.access.free || [])
    : [];
  var all = config.MODULES && Array.isArray(config.MODULES.all) ? config.MODULES.all : [];
  var icons = config.MODULES && config.MODULES.icons ? config.MODULES.icons : {};
  var labels = config.MODULES && config.MODULES.labels ? config.MODULES.labels : {};
  var enabled = {};
  access.forEach(function(name) { enabled[String(name)] = true; });

  // This helper is for entitlement display only. Do not open the client
  // workspace here: Admin has no canonical cross-business usage aggregator,
  // and opening 12 module sheets makes the management page unnecessarily slow.
  return all.map(function(name) {
    return {
      name: name,
      label: labels[name] || name,
      icon: icons[name] || 'fa-folder',
      enabled: !!enabled[name],
      recordCount: null,
      recordCountAvailable: false,
      hasRecords: false,
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

    // Businesses is the canonical grouping for the Admin Businesses page.
    // Free businesses may not have a Clients deployment-registry row, so the
    // business drill-down must still work for them.
    if (!client) {
      var masterForBusiness = getBizOSMasterSpreadsheet_();
      var businessesSheet = masterForBusiness.getSheetByName('Businesses');
      if (businessesSheet && businessesSheet.getLastRow() >= 2) {
        var bv = businessesSheet.getDataRange().getValues();
        var bh = bv[0].map(function(h) { return String(h || '').trim(); });
        var bidCol = bh.indexOf('Business_ID');
        for (var bi = 1; bi < bv.length; bi++) {
          if (bidCol >= 0 && String(bv[bi][bidCol] || '').trim() === wanted) {
            var fallback = {};
            bh.forEach(function(h, j) { fallback[h] = bv[bi][j]; });
            client = {
              Client_ID: '',
              Business_ID: String(fallback.Business_ID || wanted),
              Email: String(fallback.Owner_Email || ''),
              Client_Name: String(fallback.Business_Name || ''),
              Tier: String(fallback.Subscription_Tier || 'free'),
              Status: String(fallback.Status || 'active'),
              Created_At: fallback.Created_At || '',
              Updated_At: fallback.Last_Active || '',
              Sheet_ID: String(fallback.Workspace_ID || '')
            };
            break;
          }
        }
      }
    }
    if (!client) return {success:false, code:'CLIENT_NOT_FOUND', message:'No business was found for this business ID.'};

    var businessId = String(client.Business_ID || '').trim();
    var tier = normalizeSubscriptionTier(client.Tier);
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
      var master = getBizOSMasterSpreadsheet_();
      var masterPayments = master.getSheetByName('Payments');
      var masterUpgrades = master.getSheetByName('Upgrade_Requests');
      var requestOwners = {};

      if (masterUpgrades && masterUpgrades.getLastRow() >= 2) {
        var uv=masterUpgrades.getDataRange().getValues(), uh=uv[0].map(function(h){return String(h||'').trim();});
        var ur=uh.indexOf('Request_ID'), ue=uh.indexOf('Email'), uw=uh.indexOf('Workspace_Email'), ub=uh.indexOf('Business_Name');
        for(var ui=1;ui<uv.length;ui++){
          var rid=ur>=0?String(uv[ui][ur]||'').trim():'';
          if(!rid)continue;
          requestOwners[rid]={
            email:ue>=0?String(uv[ui][ue]||'').trim().toLowerCase():'',
            workspaceEmail:uw>=0?String(uv[ui][uw]||'').trim().toLowerCase():'',
            businessName:ub>=0?String(uv[ui][ub]||'').trim():''
          };
        }
      }

      if (masterPayments && masterPayments.getLastRow() >= 2) {
        var pv = masterPayments.getDataRange().getValues();
        var ph = pv[0].map(function(h){return String(h || '').trim();});
        var pReq = ph.indexOf('Request_ID');
        for (var pi = 1; pi < pv.length; pi++) {
          var requestId = pReq >= 0 ? String(pv[pi][pReq] || '').trim() : '';
          var owner = requestOwners[requestId] || {};
          var ownerMatches = !requestId ||
            String(owner.email || '').toLowerCase() === String(client.Email || '').trim().toLowerCase() ||
            String(owner.workspaceEmail || '').toLowerCase() === String(client.Email || '').trim().toLowerCase();

          if (!ownerMatches) continue;

          var payment = {};
          ph.forEach(function(h,j){payment[h]=pv[pi][j];});
          payment.Business_ID = businessId;
          payment.Client_ID = String(client.Client_ID || '');
          payment.Business_Name = owner.businessName || String(client.Client_Name || '');
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
          var lastRow=s.getLastRow(), lastCol=s.getLastColumn();
          var header=s.getRange(1,1,1,lastCol).getValues()[0]||[];
          var startRow=Math.max(2,lastRow-49);
          var v=s.getRange(startRow,1,lastRow-startRow+1,lastCol).getValues();
          return v.reverse().map(function(r){
            var o={};
            header.forEach(function(k,j){o[String(k||'')]=r[j];});
            return o;
          });
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
        ownershipStatus:String(client.Workspace_Ownership_Status || ''),
        maintenancePlan:String(client.Maintenance_Plan || 'none'),
        maintenanceStatus:String(client.Maintenance_Status || 'inactive'),
        maintenanceStartAt:String(client.Maintenance_Start_At || ''),
        maintenanceNextBackupAt:String(client.Maintenance_Next_Backup_At || ''),
        maintenanceLastBackupAt:String(client.Maintenance_Last_Backup_At || ''),
        maintenanceLastBackupStatus:String(client.Maintenance_Last_Backup_Status || ''),
        maintenanceLastBackupError:String(client.Maintenance_Last_Backup_Error || '')
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

function setClientMaintenancePlan(clientId, planId, status, sessionId) {
  requireAdminSession_(sessionId);
  try {
    var wanted=String(clientId||'').trim(), plan=String(planId||'none').trim().toLowerCase(), state=String(status||'inactive').trim().toLowerCase();
    var allowedPlans=['none','monthly'];
    if(allowedPlans.indexOf(plan)<0)return{success:false,message:'Unsupported maintenance plan.'};
    if(['active','inactive'].indexOf(state)<0)return{success:false,message:'Invalid maintenance status.'};
    if(plan==='none')state='inactive';
    var sheet=getBizOSMasterSpreadsheet_().getSheetByName('Clients');
    if(!sheet)return{success:false,message:'Client registry is unavailable.'};
    var v=sheet.getDataRange().getValues(),h=v[0].map(function(x){return String(x||'').trim();});
    var idc=h.indexOf('Client_ID');if(idc<0)return{success:false,message:'Client registry is missing Client_ID.'};
    var pc=h.indexOf('Maintenance_Plan'),sc=h.indexOf('Maintenance_Status'),startc=h.indexOf('Maintenance_Start_At'),nextc=h.indexOf('Maintenance_Next_Backup_At');
    if(pc<0||sc<0||startc<0||nextc<0)return{success:false,message:'Maintenance registry columns are unavailable.'};
    for(var i=1;i<v.length;i++){
      if(String(v[i][idc]||'').trim()!==wanted)continue;
      var now=new Date(),start=state==='active'?now:'';
      var next=state==='active'?(new Date(now.getTime()+30*24*60*60*1000)):''; 
      if(plan==='monthly'&&state==='active'){
        // First scheduled backup is due after the plan starts; manual backups remain available immediately.
        sheet.getRange(i+1,pc+1).setValue(plan);
        sheet.getRange(i+1,sc+1).setValue(state);
        sheet.getRange(i+1,startc+1).setValue(start.toISOString());
        sheet.getRange(i+1,nextc+1).setValue(next.toISOString());
      }else{
        sheet.getRange(i+1,pc+1).setValue('none');
        sheet.getRange(i+1,sc+1).setValue('inactive');
        sheet.getRange(i+1,startc+1).setValue('');
        sheet.getRange(i+1,nextc+1).setValue('');
      }
      return{success:true,plan:plan,status:state,nextBackupAt:next?next.toISOString():'',message:state==='active'?'Monthly recovery backups enabled.':'Recurring recovery backups disabled.'};
    }
    return{success:false,message:'Client was not found.'};
  }catch(error){console.error('setClientMaintenancePlan error:',error);return{success:false,message:'Maintenance backup settings could not be updated.'};}
}

function runScheduledClientMaintenanceBackups() {
  var lock=LockService.getScriptLock();
  if(!lock.tryLock(5000))return{success:false,message:'A maintenance backup run is already in progress.'};
  var results=[];
  try{
    var sheet=getBizOSMasterSpreadsheet_().getSheetByName('Clients');
    if(!sheet||sheet.getLastRow()<2)return{success:true,processed:0,backups:0,skipped:0};
    var v=sheet.getDataRange().getValues(),h=v[0].map(function(x){return String(x||'').trim();});
    var idx={};['Client_ID','Status','Sheet_ID','Workspace_ID','Client_Name','Business_ID','Maintenance_Plan','Maintenance_Status','Maintenance_Next_Backup_At','Maintenance_Last_Backup_At','Maintenance_Last_Backup_Status','Maintenance_Last_Backup_Error'].forEach(function(k){idx[k]=h.indexOf(k);});
    var now=new Date(),processed=0,backups=0,skipped=0;
    for(var i=1;i<v.length;i++){
      var status=idx.Status>=0?String(v[i][idx.Status]||'').trim().toLowerCase():'';
      var plan=idx.Maintenance_Plan>=0?String(v[i][idx.Maintenance_Plan]||'').trim().toLowerCase():'';
      var active=idx.Maintenance_Status>=0?String(v[i][idx.Maintenance_Status]||'').trim().toLowerCase():'';
      if(status!=='active'||plan!=='monthly'||active!=='active'){skipped++;continue;}
      var nextRaw=idx.Maintenance_Next_Backup_At>=0?String(v[i][idx.Maintenance_Next_Backup_At]||'').trim():'';
      var next=nextRaw?new Date(nextRaw):now;
      if(isNaN(next.getTime())||next>now){skipped++;continue;}
      processed++;
      var client={};h.forEach(function(k,j){client[k]=v[i][j];});
      try{
        var sheetId=String(client.Sheet_ID||client.Workspace_ID||'').trim();
        if(!sheetId)throw new Error('Client workspace ID is missing.');
        var ws=SpreadsheetApp.openById(sheetId);
        var backup=createClientRecoveryBackup_(ws,String(client.Client_Name||''),String(client.Business_ID||''));
        if(!backup||!backup.success)throw new Error(String(backup&&backup.message||'Recovery backup failed.'));
        updateClientRecoveryMetadata_(client,backup);
        var nextDate=new Date(now.getTime()+30*24*60*60*1000);
        if(idx.Maintenance_Last_Backup_At>=0)sheet.getRange(i+1,idx.Maintenance_Last_Backup_At+1).setValue(now.toISOString());
        if(idx.Maintenance_Last_Backup_Status>=0)sheet.getRange(i+1,idx.Maintenance_Last_Backup_Status+1).setValue('success');
        if(idx.Maintenance_Last_Backup_Error>=0)sheet.getRange(i+1,idx.Maintenance_Last_Backup_Error+1).setValue('');
        if(idx.Maintenance_Next_Backup_At>=0)sheet.getRange(i+1,idx.Maintenance_Next_Backup_At+1).setValue(nextDate.toISOString());
        logClientRecoveryActivity_(client,'BACKUP_CREATED',{status:'created',backupId:backup.backupId,source:'scheduled_maintenance',message:'Scheduled monthly recovery backup created.'});
        results.push({clientId:String(client.Client_ID||''),success:true,backupId:String(backup.backupId||''),nextBackupAt:nextDate.toISOString()});backups++;
      }catch(error){
        if(idx.Maintenance_Last_Backup_Status>=0)sheet.getRange(i+1,idx.Maintenance_Last_Backup_Status+1).setValue('failed');
        if(idx.Maintenance_Last_Backup_Error>=0)sheet.getRange(i+1,idx.Maintenance_Last_Backup_Error+1).setValue(String(error&&error.message||error).slice(0,1000));
        results.push({clientId:String(client.Client_ID||''),success:false,message:String(error&&error.message||error)});
      }
    }
    return{success:true,processed:processed,backups:backups,skipped:skipped,results:results};
  }finally{lock.releaseLock();}
}

function installClientMaintenanceBackupTrigger() {
  var existing=ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()==='runScheduledClientMaintenanceBackups';});
  existing.slice(1).forEach(function(t){ScriptApp.deleteTrigger(t);});
  if(existing.length===0)ScriptApp.newTrigger('runScheduledClientMaintenanceBackups').timeBased().everyDays(1).atHour(2).create();
  return{success:true,message:'Daily client maintenance backup trigger is installed.',triggerCount:1};
}

function createClientAdminRecoveryBackup(clientId, sessionId) {
  requireAdminSession_(sessionId);
  try {
    var client=getActiveClientByIdForAdmin_(clientId);
    if(!client)return{success:false,message:'Active client deployment not found.'};
    var ws=SpreadsheetApp.openById(String(client.sheetId||client.Sheet_ID||''));
    var result=createClientRecoveryBackup_(ws,String(client.clientName||client.Client_Name||''),String(client.businessId||client.Business_ID||''));
    if(!result||!result.success)return result||{success:false,message:'Backup could not be created.'};
    updateClientRecoveryMetadata_(client,result);
    logClientRecoveryActivity_(client,'BACKUP_CREATED',{status:'created',backupId:result.backupId,message:'Client recovery backup created.'});
    return result;
  }catch(error){console.error('createClientAdminRecoveryBackup error:',error);return{success:false,message:'BizOS could not create the client backup.'};}
}

function archiveClientDeployment(clientId, sessionId) {
  requireAdminSession_(sessionId);
  try {
    var wanted=String(clientId||'').trim();
    var master=getBizOSMasterSpreadsheet_();
    var clients=master.getSheetByName('Clients');
    if(!clients)return{success:false,message:'Client registry is unavailable.'};

    var v=clients.getDataRange().getValues(),h=v[0].map(function(x){return String(x||'').trim();});
    var idCol=h.indexOf('Client_ID'),statusCol=h.indexOf('Status'),businessIdCol=h.indexOf('Business_ID'),updatedCol=h.indexOf('Updated_At');
    if(idCol<0||statusCol<0)return{success:false,message:'Client registry schema is incomplete.'};

    var businessId='';
    var clientRow=-1;
    for(var i=1;i<v.length;i++){
      if(String(v[i][idCol]||'').trim()!==wanted)continue;
      businessId=businessIdCol>=0?String(v[i][businessIdCol]||'').trim():'';
      clientRow=i+1;
      break;
    }
    if(clientRow<0)return{success:false,message:'Client was not found.'};

    clients.getRange(clientRow,statusCol+1).setValue('archived');
    if(updatedCol>=0)clients.getRange(clientRow,updatedCol+1).setValue(new Date().toISOString());

    // Keep the business lifecycle aligned with the client deployment lifecycle.
    // Historical users, workspace data, payments and deployment metadata remain
    // intact; only the active state is changed.
    if(businessId){
      var businesses=master.getSheetByName('Businesses');
      if(businesses && businesses.getLastRow()>=2){
        var bv=businesses.getDataRange().getValues(),bh=bv[0].map(function(x){return String(x||'').trim();});
        var bb=bh.indexOf('Business_ID'), bs=bh.indexOf('Status');
        if(bb>=0&&bs>=0){
          for(var bi=1;bi<bv.length;bi++){
            if(String(bv[bi][bb]||'').trim()===businessId){
              businesses.getRange(bi+1,bs+1).setValue('archived');
              break;
            }
          }
        }
      }
    }

    logClientRecoveryActivity_(getClientByIdForAdminAnyStatus_(wanted),'ACCOUNT_ARCHIVED',{status:'archived',message:'Client and business archived. Historical records were preserved.'});
    return{success:true,message:'Client and business archived. Historical records were preserved.'};
  }catch(error){
    console.error('archiveClientDeployment error:',error);
    return{success:false,message:'Client could not be archived.'};
  }
}

function restoreArchivedClientDeployment(clientId, sessionId) {
  requireAdminSession_(sessionId);
  try {
    var wanted=String(clientId||'').trim();
    var master=getBizOSMasterSpreadsheet_(),clients=master.getSheetByName('Clients');
    if(!clients)return{success:false,message:'Client registry is unavailable.'};

    var v=clients.getDataRange().getValues(),h=v[0].map(function(x){return String(x||'').trim();});
    var idCol=h.indexOf('Client_ID'),statusCol=h.indexOf('Status'),businessIdCol=h.indexOf('Business_ID'),updatedCol=h.indexOf('Updated_At');
    if(idCol<0||statusCol<0)return{success:false,message:'Client registry schema is incomplete.'};

    var businessId='',clientRow=-1;
    for(var i=1;i<v.length;i++){
      if(String(v[i][idCol]||'').trim()!==wanted)continue;
      businessId=businessIdCol>=0?String(v[i][businessIdCol]||'').trim():'';
      clientRow=i+1;
      break;
    }
    if(clientRow<0)return{success:false,message:'Client was not found.'};
    if(String(v[clientRow-1][statusCol]||'').toLowerCase()!=='archived')return{success:false,message:'Only archived clients can be restored.'};

    clients.getRange(clientRow,statusCol+1).setValue('active');
    if(updatedCol>=0)clients.getRange(clientRow,updatedCol+1).setValue(new Date().toISOString());

    if(businessId){
      var businesses=master.getSheetByName('Businesses');
      if(businesses && businesses.getLastRow()>=2){
        var bv=businesses.getDataRange().getValues(),bh=bv[0].map(function(x){return String(x||'').trim();});
        var bb=bh.indexOf('Business_ID'), bs=bh.indexOf('Status');
        if(bb>=0&&bs>=0){
          for(var bi=1;bi<bv.length;bi++){
            if(String(bv[bi][bb]||'').trim()===businessId){
              businesses.getRange(bi+1,bs+1).setValue('active');
              break;
            }
          }
        }
      }
    }

    logClientRecoveryActivity_(getClientByIdForAdminAnyStatus_(wanted),'ACCOUNT_RESTORED',{status:'active',message:'Client and business restored.'});
    return{success:true,message:'Client and business restored.'};
  }catch(error){
    console.error('restoreArchivedClientDeployment error:',error);
    return{success:false,message:'Client could not be restored.'};
  }
}

function getActiveClientByIdForAdmin_(clientId) {
  var client = getClientByIdForAdminAnyStatus_(clientId);
  if (!client) return null;
  var status = String(client.Status || '').trim().toLowerCase();
  return (!status || status === 'active') ? client : null;
}

function getClientByIdForAdminAnyStatus_(clientId) {
  var wanted=String(clientId||'').trim();if(!wanted)return null;
  var s=getBizOSMasterSpreadsheet_().getSheetByName('Clients');if(!s)return null;
  var v=s.getDataRange().getValues();if(v.length<2)return null;
  var h=v[0].map(function(x){return String(x||'').trim();}),idCol=h.indexOf('Client_ID');if(idCol<0)return null;
  for(var i=1;i<v.length;i++){
    if(String(v[i][idCol]||'').trim()!==wanted)continue;
    var o={};h.forEach(function(k,j){o[k]=v[i][j];});
    return o;
  }
  return null;
}

function updateClientRecoveryMetadata_(client, backup) {
  try {
    if(!client || !backup || !backup.success)return;
    var clientId=String(client.Client_ID||client.clientId||'').trim();
    if(!clientId)return;
    var sheet=getBizOSMasterSpreadsheet_().getSheetByName('Clients');
    if(!sheet)return;
    var v=sheet.getDataRange().getValues(),h=v[0].map(function(x){return String(x||'').trim();});
    var idCol=h.indexOf('Client_ID'),backupIdCol=h.indexOf('Recovery_Backup_ID'),backupUrlCol=h.indexOf('Recovery_Backup_URL'),updatedCol=h.indexOf('Updated_At');
    if(idCol<0)return;
    for(var i=1;i<v.length;i++){
      if(String(v[i][idCol]||'').trim()!==clientId)continue;
      if(backupIdCol>=0)sheet.getRange(i+1,backupIdCol+1).setValue(String(backup.backupId||''));
      if(backupUrlCol>=0)sheet.getRange(i+1,backupUrlCol+1).setValue(String(backup.backupUrl||''));
      if(updatedCol>=0)sheet.getRange(i+1,updatedCol+1).setValue(new Date().toISOString());
      return;
    }
  }catch(error){console.warn('Client recovery metadata update warning:',error&&error.message?error.message:error);}
}

function logClientRecoveryActivity_(client, type, details) {
  try {
    var clientId=String(client && (client.Client_ID || client.clientId) || '');
    var businessId=String(client && (client.Business_ID || client.businessId) || '');
    var sheetId=String(client && (client.Sheet_ID || client.Workspace_ID || client.sheetId) || '');
    var masterDetails=Object.assign({},details||{},{
      clientId:clientId,businessId:businessId,type:type
    });
    logAdminAction('client_recovery_'+String(type||'event').toLowerCase(),masterDetails);

    // Reuse the existing client activity sheet when it is present. We map only
    // columns that already exist, so this does not create a second activity schema.
    if(!sheetId)return;
    var ss=SpreadsheetApp.openById(sheetId),s=ss.getSheetByName('Client_Activity');
    if(!s || s.getLastColumn()<1)return;
    var h=s.getRange(1,1,1,s.getLastColumn()).getValues()[0].map(function(x){return String(x||'').trim();});
    var now=new Date().toISOString();
    var row=new Array(h.length).fill('');
    var values={
      Activity_ID:'ACT_REC_'+Utilities.getUuid().replace(/-/g,'').slice(0,16).toUpperCase(),
      Activity_Type:type,
      Type:type,
      Action:type,
      Event:type,
      Description:String(details && details.message || ('Recovery event: '+type)),
      Details:JSON.stringify(masterDetails),
      Business_ID:businessId,
      Client_ID:clientId,
      Created_At:now,
      Timestamp:now,
      Updated_At:now,
      Status:String(details && details.status || '')
    };
    h.forEach(function(header,i){if(Object.prototype.hasOwnProperty.call(values,header))row[i]=values[header];});
    if(row.some(function(v){return v!=='';}))s.appendRow(row);
  }catch(error){
    console.warn('Client recovery activity logging warning:',error && error.message ? error.message : error);
  }
}

function auditClientRecoveryBackup(clientId, backupId, sessionId) {
  requireAdminSession_(sessionId);
  try {
    var client=getClientByIdForAdminAnyStatus_(clientId);
    if(!client)return{success:false,code:'CLIENT_NOT_FOUND',message:'Client was not found.'};
    var sourceId=String(backupId||'').trim();
    if(!sourceId)return{success:false,code:'RECOVERY_BACKUP_REQUIRED',message:'A recovery backup is required.'};

    var safeName=String(client.Client_Name||'Client').replace(/[^a-zA-Z0-9]/g,'_');
    var businessId=String(client.Business_ID||'').trim();
    var expectedFolderName=safeName+'_'+businessId.replace(/[^a-zA-Z0-9]/g,'_');
    var expectedPrefix='BizOS Recovery - '+safeName+' - ';
    var file;
    try{file=DriveApp.getFileById(sourceId);}catch(e){return{success:false,code:'RECOVERY_BACKUP_NOT_FOUND',message:'The selected recovery backup could not be found.'};}

    var parentMatches=false,parents=file.getParents();
    while(parents.hasNext()){if(String(parents.next().getName()||'')===expectedFolderName){parentMatches=true;break;}}
    var report={
      success:true,healthy:true,checkedAt:new Date().toISOString(),clientId:String(client.Client_ID||''),
      businessId:businessId,backupId:sourceId,backupName:String(file.getName()||''),createdAt:file.getDateCreated().toISOString(),
      checks:{scope:false,type:false,sheets:false,modules:{},businessCenter:{},staff:false,activity:false},
      warnings:[]
    };
    if(!parentMatches || String(file.getName()||'').indexOf(expectedPrefix)!==0){
      report.healthy=false;report.warnings.push('Backup is outside the expected business recovery folder or has an unexpected name.');
      return report;
    }
    report.checks.scope=true;
    if(String(file.getMimeType()||'')!=='application/vnd.google-apps.spreadsheet'){
      report.healthy=false;report.warnings.push('Recovery object is not a Google Spreadsheet snapshot.');
      return report;
    }
    report.checks.type=true;

    var ss=SpreadsheetApp.openById(sourceId),sheetMap={
      Finance:'Financial_Data',Ecommerce:'Ecommerce_Data',Sales:'Sales_Data',CRM:'CRM_Data',
      HR:'HR_Data',Logistics:'Logistics_Data',Tax:'Tax_Data',Agro:'Agro_Data',
      Productivity:'Productivity_Data',POS:'POS_Data',Attendance:'Attendance_Data',Warehouse:'Warehouse_Data'
    };
    var allNames=ss.getSheets().map(function(s){return s.getName();}),nameSet={};
    allNames.forEach(function(n){nameSet[n]=true;});
    Object.keys(sheetMap).forEach(function(module){
      var name=sheetMap[module],present=!!nameSet[name];
      report.checks.modules[module]={sheet:name,present:present,rows:present?Math.max(0,ss.getSheetByName(name).getLastRow()-1):0};
      if(!present){report.healthy=false;report.warnings.push(module+' workspace sheet is missing: '+name);}
    });
    report.checks.sheets=Object.keys(report.checks.modules).every(function(k){return report.checks.modules[k].present;});

    var bc=['BusinessCenter_Products','BusinessCenter_Customers','BusinessCenter_Sales','BusinessCenter_Sale_Items','BusinessCenter_Stock_Movements','BusinessCenter_Customer_Ledger','BusinessCenter_Finance_Transactions'];
    bc.forEach(function(name){
      var present=!!nameSet[name];
      report.checks.businessCenter[name]={present:present,rows:present?Math.max(0,ss.getSheetByName(name).getLastRow()-1):0};
    });
    report.checks.businessCenter.expected=bc.filter(function(name){return nameSet[name];}).length;
    report.checks.businessCenter.present=bc.filter(function(name){return nameSet[name];}).length;
    // Business Center sheets are only required when this workspace actually has
    // Business Center data; an empty workspace may legitimately have none.
    if(report.checks.businessCenter.present>0 && report.checks.businessCenter.present<bc.length){
      report.healthy=false;report.warnings.push('Business Center snapshot is incomplete: some existing Business Center sheets are missing.');
    }

    ['Assigned_Modules','User_Data','Client_Activity','Client_Notification_Preferences','Client_Migration_Index'].forEach(function(name){
      if(nameSet[name])report.checks.staff=true;
    });
    report.checks.activity=!!nameSet.Client_Activity;
    report.checks.staff=!!(nameSet.Assigned_Modules||nameSet.User_Data);
    if(!report.checks.activity)report.warnings.push('Client_Activity is not present in the recovery snapshot.');
    if(!report.checks.staff)report.warnings.push('No client staff/access sheet was found in the recovery snapshot.');

    report.healthy=!!report.healthy;
    return report;
  }catch(error){
    console.error('auditClientRecoveryBackup error:',error);
    return{success:false,code:'RECOVERY_AUDIT_FAILED',message:'The recovery backup could not be audited.',error:String(error&&error.message||error)};
  }
}

function verifyRestoredClientWorkspace_(client, workspace) {
  var sheetMap={
    Finance:'Financial_Data',Ecommerce:'Ecommerce_Data',Sales:'Sales_Data',CRM:'CRM_Data',
    HR:'HR_Data',Logistics:'Logistics_Data',Tax:'Tax_Data',Agro:'Agro_Data',
    Productivity:'Productivity_Data',POS:'POS_Data',Attendance:'Attendance_Data',Warehouse:'Warehouse_Data'
  };
  var names={};workspace.getSheets().forEach(function(s){names[s.getName()]=true;});
  var missing=[];
  Object.keys(sheetMap).forEach(function(k){if(!names[sheetMap[k]])missing.push(sheetMap[k]);});
  var result={success:missing.length===0,checkedAt:new Date().toISOString(),missingSheets:missing,sheetCount:workspace.getSheets().length};
  if(!names.Client_Info)result.success=false;
  result.clientInfoPresent=!!names.Client_Info;
  result.businessId=String(client.Business_ID||'');
  return result;
}

function listClientRecoveryBackups(clientId, sessionId) {
  requireAdminSession_(sessionId);
  try {
    var client=getClientByIdForAdminAnyStatus_(clientId);
    if(!client)return{success:false,message:'Client not found.'};
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
  var client=null;
  try {
    client=getClientByIdForAdminAnyStatus_(clientId);
    if(!client)return{success:false,message:'Client not found.'};
    var targetId=String(client.Sheet_ID||client.Workspace_ID||'').trim(),sourceId=String(backupId||'').trim();
    if(!targetId||!sourceId)return{success:false,message:'Client workspace and backup are required.'};
    if(targetId===sourceId)return{success:false,message:'The selected backup is already the live workspace.'};

    var target=SpreadsheetApp.openById(targetId),sourceFile;
    try{sourceFile=DriveApp.getFileById(sourceId);}catch(e){return{success:false,code:'RECOVERY_BACKUP_NOT_FOUND',message:'The selected recovery backup could not be found.'};}

    var safeName=String(client.Client_Name||'Client').replace(/[^a-zA-Z0-9]/g,'_');
    var businessId=String(client.Business_ID||'').trim();
    var expectedFolderName=safeName+'_'+businessId.replace(/[^a-zA-Z0-9]/g,'_');
    var expectedPrefix='BizOS Recovery - '+safeName+' - ';
    var parentMatches=false,parents=sourceFile.getParents();
    while(parents.hasNext()){if(String(parents.next().getName()||'')===expectedFolderName){parentMatches=true;break;}}
    if(!parentMatches||String(sourceFile.getName()||'').indexOf(expectedPrefix)!==0)return{success:false,code:'RECOVERY_BACKUP_SCOPE_MISMATCH',message:'The selected backup does not belong to this business recovery archive.'};
    if(String(sourceFile.getMimeType()||'')!=='application/vnd.google-apps.spreadsheet')return{success:false,code:'RECOVERY_BACKUP_INVALID_TYPE',message:'The selected recovery backup is not a BizOS workspace snapshot.'};

    var health=auditClientRecoveryBackup(clientId,sourceId,sessionId);
    if(!health.success||!health.healthy)return{success:false,code:'RECOVERY_BACKUP_UNHEALTHY',message:'The selected recovery snapshot did not pass integrity checks. The live workspace was not changed.',audit:health};

    var safety=createClientRecoveryBackup_(target,String(client.Client_Name||''),businessId);
    if(!safety||!safety.success)return{success:false,message:'A safety backup could not be created, so the restore was not started.'};
    updateClientRecoveryMetadata_(client,safety);
    logClientRecoveryActivity_(client,'BACKUP_RESTORE_STARTED',{status:'started',backupId:sourceId,safetyBackupId:safety.backupId,message:'Client workspace restore started.'});

    var source=SpreadsheetApp.openById(sourceId);
    source.getSheets().forEach(function(src){
      var name=src.getName(),dst=target.getSheetByName(name);
      if(!dst)dst=target.insertSheet(name);
      dst.clear({contentsOnly:false});
      var range=src.getDataRange();
      if(range.getNumRows()&&range.getNumColumns())range.copyTo(dst.getRange(1,1),{contentsOnly:false});
    });

    var verification=verifyRestoredClientWorkspace_(client,target);
    if(!verification.success){
      logClientRecoveryActivity_(client,'BACKUP_RESTORE_FAILED',{status:'failed',backupId:sourceId,safetyBackupId:safety.backupId,message:'Post-restore verification failed.',verification:verification});
      return{success:false,code:'RESTORE_VERIFICATION_FAILED',message:'The workspace was restored, but post-restore verification found missing required components. The safety backup was preserved.',safetyBackup:safety,verification:verification};
    }
    logClientRecoveryActivity_(client,'BACKUP_RESTORED',{status:'restored',backupId:sourceId,safetyBackupId:safety.backupId,message:'Client workspace restored and verified.'});
    return{success:true,message:'Client workspace restored and verified.',safetyBackup:safety,restoredFrom:sourceId,verification:verification};
  }catch(error){
    console.error('restoreClientWorkspaceFromRecoveryBackup error:',error);
    try{logClientRecoveryActivity_(client,'BACKUP_RESTORE_FAILED',{status:'failed',backupId:String(backupId||''),message:error.message||String(error)});}catch(ignore){}
    return{success:false,message:'The client workspace could not be restored. The safety backup remains available.'};
  }
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
