
// ==================== STAFF SECURITY + INVITATION HELPERS ====================

function ensureStaffInvitationColumns_(sheet) {
  const required = ['Invitation_Expires_At'];
  const lastCol = sheet.getLastColumn();
  const headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  required.forEach(function(name) {
    if (headers.indexOf(name) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(name);
    }
  });
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function staffActor_(sessionId, businessId, allowedRoles) {
  const user = getUserFromSession(sessionId);
  if (!user) return { ok: false, message: 'Unauthorized. Please login again.' };
  if (String(user.businessId || '') !== String(businessId || '')) {
    return { ok: false, message: 'You are not authorized for this business.' };
  }
  const role = String(user.role || '').toLowerCase();
  if (allowedRoles.indexOf(role) === -1) {
    return { ok: false, message: 'You do not have permission to perform this action.' };
  }
  return { ok: true, user: user };
}

function normalizeStaffModules_(modules) {
  const valid = getAvailableModulesForStaff();
  const input = Array.isArray(modules) ? modules : String(modules || '').split(',');
  const seen = {};
  return input.map(function(m) { return String(m || '').trim(); })
    .filter(function(m) {
      if (!m || valid.indexOf(m) === -1 || seen[m]) return false;
      seen[m] = true;
      return true;
    });
}

function syncStaffModules_(email, businessId, modules, assignedBy) {
  const normalized = normalizeStaffModules_(modules);
  const sheet = createStaffModulesSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || [];
  const emailCol = headers.indexOf('Staff_Email');
  const bizIdCol = headers.indexOf('Business_ID');
  const moduleCol = headers.indexOf('Module_Access');

  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][emailCol] || '').toLowerCase() === String(email || '').toLowerCase() &&
        String(values[i][bizIdCol] || '') === String(businessId || '')) {
      sheet.deleteRow(i + 1);
    }
  }

  normalized.forEach(function(moduleName) {
    sheet.appendRow([
      String(email || '').toLowerCase(),
      businessId,
      moduleName,
      assignedBy || 'system',
      new Date().toISOString()
    ]);
  });

  return normalized;
}

// ==================== STAFF INVITATION SYSTEM ====================

function inviteStaffMember(businessId, email, name, role, invitedByEmail, assignedModules) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return { success: false, message: 'Another staff change is in progress. Please try again.' };

    email = String(email || '').trim().toLowerCase();
    invitedByEmail = String(invitedByEmail || '').trim().toLowerCase();
    role = String(role || 'staff').trim().toLowerCase();

    const userSheet = getOrCreateUserSheet();
    const headers = ensureStaffInvitationColumns_(userSheet);
    const data = userSheet.getDataRange().getValues();
    const emailCol = headers.indexOf('Email');
    const roleCol = headers.indexOf('Role');
    const businessIdCol = headers.indexOf('Business_ID');
    const inviterRow = data.find(function(row) {
      return String(row[emailCol] || '').toLowerCase() === invitedByEmail &&
             String(row[businessIdCol] || '') === String(businessId);
    });

    if (!inviterRow || (String(inviterRow[roleCol] || '').toLowerCase() !== 'owner' &&
        String(inviterRow[roleCol] || '').toLowerCase() !== 'admin')) {
      return { success: false, message: 'Only business owners or admins can invite staff members.' };
    }
    if (!email || !email.includes('@')) return { success: false, message: 'Enter a valid staff email.' };
    if (role !== 'staff' && role !== 'admin') return { success: false, message: 'Invalid staff role.' };

    const modules = role === 'admin' ? getAvailableModulesForStaff() : normalizeStaffModules_(assignedModules);
    if (!modules.length) return { success: false, message: 'Please assign at least one valid module for the staff member.' };

    const invitationCode = generateInvitationCode();
    const expiresAt = new Date(Date.now() + 48 * 60 * 60 * 1000).toISOString();
    const existingIndex = data.findIndex(function(row) {
      return String(row[emailCol] || '').toLowerCase() === email &&
             String(row[businessIdCol] || '') === String(businessId);
    });

    const invitedByCol = headers.indexOf('Invited_By');
    const statusCol = headers.indexOf('Invitation_Status');
    const codeCol = headers.indexOf('Invitation_Code');
    const assignedCol = headers.indexOf('Assigned_Modules');
    const expiresCol = headers.indexOf('Invitation_Expires_At');
    const now = new Date().toISOString();

    if (existingIndex >= 1) {
      userSheet.getRange(existingIndex + 1, invitedByCol + 1).setValue(invitedByEmail);
      userSheet.getRange(existingIndex + 1, statusCol + 1).setValue('pending');
      userSheet.getRange(existingIndex + 1, codeCol + 1).setValue(invitationCode);
      userSheet.getRange(existingIndex + 1, assignedCol + 1).setValue(modules.join(','));
      userSheet.getRange(existingIndex + 1, expiresCol + 1).setValue(expiresAt);
    } else {
      const row = new Array(headers.length).fill('');
      row[emailCol] = email;
      row[headers.indexOf('Name')] = name || '';
      row[roleCol] = role;
      row[businessIdCol] = businessId;
      row[headers.indexOf('Is_Verified')] = 'YES';
      row[invitedByCol] = invitedByEmail;
      row[statusCol] = 'pending';
      row[codeCol] = invitationCode;
      row[headers.indexOf('Created_At')] = now;
      row[headers.indexOf('Is_Client')] = 'NO';
      row[assignedCol] = modules.join(',');
      row[expiresCol] = expiresAt;
      userSheet.appendRow(row);
    }

    // Assigned_Modules is the canonical assignment. Staff_Modules is kept
    // synchronized for legacy callers during the transition.
    syncStaffModules_(email, businessId, modules, invitedByEmail);
    sendInvitationEmail(email, name, invitationCode, role);

    return {
      success: true,
      message: 'Invitation sent to ' + email,
      invitationCode: invitationCode,
      expiresAt: expiresAt
    };
  } catch (error) {
    console.error('Staff invitation error:', error);
    return { success: false, message: error.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function acceptInvitation(invitationCode, password, name, phone) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return { success: false, message: 'Another invitation is being processed. Please try again.' };

    const userSheet = getOrCreateUserSheet();
    const headers = ensureStaffInvitationColumns_(userSheet);
    const data = userSheet.getDataRange().getValues();
    const codeCol = headers.indexOf('Invitation_Code');
    const emailCol = headers.indexOf('Email');
    const statusCol = headers.indexOf('Invitation_Status');
    const businessIdCol = headers.indexOf('Business_ID');
    const nameCol = headers.indexOf('Name');
    const phoneCol = headers.indexOf('Phone');
    const passwordCol = headers.indexOf('Password_Hash');
    const roleCol = headers.indexOf('Role');
    const expiresCol = headers.indexOf('Invitation_Expires_At');

    if ([codeCol,emailCol,statusCol,businessIdCol,passwordCol,roleCol].some(function(col){ return col === -1; })) {
      return { success: false, message: 'Staff invitation schema is incomplete.' };
    }

    invitationCode = String(invitationCode || '').trim();
    const rowIndex = data.findIndex(function(row) { return String(row[codeCol] || '') === invitationCode; });
    if (rowIndex < 1) return { success: false, message: 'Invalid invitation code.' };

    const userRow = data[rowIndex];
    if (String(userRow[statusCol] || '').toLowerCase() !== 'pending') {
      return { success: false, message: 'This invitation has already been accepted or is no longer active.' };
    }

    if (expiresCol !== -1 && userRow[expiresCol]) {
      const expiresAt = new Date(userRow[expiresCol]).getTime();
      if (!Number.isFinite(expiresAt) || Date.now() > expiresAt) {
        return { success: false, message: 'This invitation has expired. Please ask the business owner to resend it.' };
      }
    }

    password = String(password || '');
    if (password.length < Number(CONFIG.AUTH.passwordMinLength || 6)) {
      return { success: false, message: 'Password must be at least ' + CONFIG.AUTH.passwordMinLength + ' characters.' };
    }

    const businessId = String(userRow[businessIdCol] || '');
    const businessSheet = getOrCreateBusinessSheet();
    const businessData = businessSheet.getDataRange().getValues();
    const businessHeaders = businessData[0] || [];
    const bizIdIndex = businessHeaders.indexOf('Business_ID');
    const workspaceIndex = businessHeaders.indexOf('Workspace_ID');
    const businessRow = businessData.find(function(row) { return String(row[bizIdIndex] || '') === businessId; });

    if (!businessRow || bizIdIndex < 0 || workspaceIndex < 0 || !businessRow[workspaceIndex]) {
      return { success: false, message: 'The business workspace is not ready yet. Please contact the business owner.' };
    }

    const workspace = SpreadsheetApp.openById(String(businessRow[workspaceIndex]));
    const workspaceUsersSheet = workspace.getSheetByName('Workspace_Users');
    const email = String(userRow[emailCol] || '').toLowerCase();
    const displayName = String(name || userRow[nameCol] || email);
    const role = String(userRow[roleCol] || 'staff').toLowerCase();

    if (workspaceUsersSheet) {
      const wValues = workspaceUsersSheet.getDataRange().getValues();
      const wHeaders = wValues[0] || [];
      const wEmailCol = wHeaders.indexOf('Email');
      const existingWorkspaceUser = wEmailCol >= 0 && wValues.slice(1).some(function(row) {
        return String(row[wEmailCol] || '').toLowerCase() === email;
      });
      if (!existingWorkspaceUser) {
        workspaceUsersSheet.appendRow([email, displayName, role, businessId, new Date().toISOString(), 'active']);
      }
    }
    workspace.addViewer(email);

    // Only mark the invitation accepted after workspace provisioning succeeds.
    userSheet.getRange(rowIndex + 1, passwordCol + 1).setValue(hashPassword(password));
    userSheet.getRange(rowIndex + 1, statusCol + 1).setValue('accepted');
    userSheet.getRange(rowIndex + 1, codeCol + 1).setValue('');
    if (expiresCol !== -1) userSheet.getRange(rowIndex + 1, expiresCol + 1).setValue('');
    if (name && nameCol !== -1) userSheet.getRange(rowIndex + 1, nameCol + 1).setValue(name);
    if (phone && phoneCol !== -1) userSheet.getRange(rowIndex + 1, phoneCol + 1).setValue(phone);

    return { success: true, message: 'Invitation accepted! You can now login.' };
  } catch (error) {
    console.error('Accept invitation error:', error);
    return { success: false, message: error.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function resendStaffInvitation(businessId, email, sessionId) {
  try {
    const actor = staffActor_(sessionId, businessId, ['owner', 'admin']);
    if (!actor.ok) return { success: false, message: actor.message };

    const userSheet = getOrCreateUserSheet();
    const headers = ensureStaffInvitationColumns_(userSheet);
    const data = userSheet.getDataRange().getValues();
    const emailCol = headers.indexOf('Email');
    const bizCol = headers.indexOf('Business_ID');
    const statusCol = headers.indexOf('Invitation_Status');
    const codeCol = headers.indexOf('Invitation_Code');
    const expiresCol = headers.indexOf('Invitation_Expires_At');
    const nameCol = headers.indexOf('Name');
    const roleCol = headers.indexOf('Role');
    const assignedCol = headers.indexOf('Assigned_Modules');

    const rowIndex = data.findIndex(function(row) {
      return String(row[emailCol] || '').toLowerCase() === String(email || '').toLowerCase() &&
             String(row[bizCol] || '') === String(businessId);
    });
    if (rowIndex < 1) return { success: false, message: 'Pending staff invitation not found.' };
    if (String(data[rowIndex][statusCol] || '').toLowerCase() === 'accepted') {
      return { success: false, message: 'This staff member has already accepted the invitation.' };
    }

    const code = generateInvitationCode();
    const expiresAt = new Date(Date.now() + 48 * 60 * 60 * 1000).toISOString();
    userSheet.getRange(rowIndex + 1, statusCol + 1).setValue('pending');
    userSheet.getRange(rowIndex + 1, codeCol + 1).setValue(code);
    userSheet.getRange(rowIndex + 1, expiresCol + 1).setValue(expiresAt);

    const role = String(data[rowIndex][roleCol] || 'staff').toLowerCase();
    const name = String(data[rowIndex][nameCol] || '');
    sendInvitationEmail(email, name, code, role);
    return { success: true, message: 'Invitation resent to ' + email, expiresAt: expiresAt };
  } catch (error) {
    console.error('Resend staff invitation error:', error);
    return { success: false, message: error.message };
  }
}

function removeTeamMember(businessId, email, sessionId) {
  try {
    const actor = staffActor_(sessionId, businessId, ['owner', 'admin']);
    if (!actor.ok) return { success: false, message: actor.message };

    const userSheet = getOrCreateUserSheet();
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    const businessIdCol = headers.indexOf('Business_ID');
    const roleCol = headers.indexOf('Role');
    const targetEmail = String(email || '').toLowerCase();
    const rowIndex = data.findIndex(function(r) {
      return String(r[emailCol] || '').toLowerCase() === targetEmail &&
             String(r[businessIdCol] || '') === String(businessId);
    });
    if (rowIndex < 1) return { success: false, message: 'User not found.' };
    if (String(data[rowIndex][roleCol] || '').toLowerCase() === 'owner') {
      return { success: false, message: 'Cannot remove business owner.' };
    }

    const assigned = getStaffModules(targetEmail, businessId);
    userSheet.deleteRow(rowIndex + 1);
    try { syncStaffModules_(targetEmail, businessId, [], actor.user.email); } catch (e) { console.error('Staff module cleanup error:', e); }
    return { success: true };
  } catch(error) {
    return { success: false, message: error.message };
  }
}

function getBusinessStaff(businessId, sessionId) {
  const actor = staffActor_(sessionId, businessId, ['owner', 'admin']);
  if (!actor.ok) return { success: false, message: actor.message };

  const userSheet = getOrCreateUserSheet();
  const data = userSheet.getDataRange().getValues();
  const headers = data[0];
  const businessIdCol = headers.indexOf('Business_ID');

  return data.slice(1)
    .filter(function(row) { return String(row[businessIdCol] || '') === String(businessId); })
    .map(function(row) {
      return {
        email: row[headers.indexOf('Email')],
        name: row[headers.indexOf('Name')],
        role: row[headers.indexOf('Role')],
        isVerified: row[headers.indexOf('Is_Verified')],
        invitationStatus: row[headers.indexOf('Invitation_Status')],
        createdAt: row[headers.indexOf('Created_At')]
      };
    });
}

// ==================== STAFF MODULE ASSIGNMENT ====================

function createStaffModulesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Staff_Modules');
  
  if (!sheet) {
    sheet = ss.insertSheet('Staff_Modules');
    sheet.appendRow(['Staff_Email', 'Business_ID', 'Module_Access', 'Assigned_By', 'Assigned_At']);
    
    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f4f6');
    
    console.log('Staff_Modules sheet created');
  }
  
  return sheet;
}

function getStaffModules(email, businessId) {
  try {
    const userSheet = getOrCreateUserSheet();
    const data = userSheet.getDataRange().getValues();
    const headers = data[0] || [];
    const emailCol = headers.indexOf('Email');
    const bizIdCol = headers.indexOf('Business_ID');
    const assignedCol = headers.indexOf('Assigned_Modules');
    if (emailCol < 0 || bizIdCol < 0) return [];

    const row = data.find(function(r) {
      return String(r[emailCol] || '').toLowerCase() === String(email || '').toLowerCase() &&
             String(r[bizIdCol] || '') === String(businessId || '');
    });
    if (!row) return [];

    // Assigned_Modules is now the single application-level source of truth.
    if (assignedCol >= 0) return normalizeStaffModules_(row[assignedCol]);

    // Legacy fallback for existing installations that predate Assigned_Modules.
    const legacy = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Staff_Modules');
    if (!legacy) return [];
    const values = legacy.getDataRange().getValues();
    const h = values[0] || [];
    const e = h.indexOf('Staff_Email'), b = h.indexOf('Business_ID'), m = h.indexOf('Module_Access');
    return values.slice(1).filter(function(r) {
      return String(r[e] || '').toLowerCase() === String(email || '').toLowerCase() &&
             String(r[b] || '') === String(businessId || '');
    }).map(function(r) { return r[m]; });
}

function assignModuleToStaff(email, businessId, moduleName, assignedBy, sessionId) {
  try {
    const actor = staffActor_(sessionId, businessId, ['owner', 'admin']);
    if (!actor.ok) return { success: false, message: actor.message };
    const normalized = normalizeStaffModules_([moduleName]);
    if (!normalized.length) return { success: false, message: 'Invalid module.' };

    const userSheet = getOrCreateUserSheet();
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    const bizCol = headers.indexOf('Business_ID');
    const assignedCol = headers.indexOf('Assigned_Modules');
    const rowIndex = data.findIndex(function(row) {
      return String(row[emailCol] || '').toLowerCase() === String(email || '').toLowerCase() &&
             String(row[bizCol] || '') === String(businessId);
    });
    if (rowIndex < 1) return { success: false, message: 'Staff member not found.' };

    const current = getStaffModules(email, businessId);
    if (current.indexOf(normalized[0]) === -1) current.push(normalized[0]);
    if (assignedCol >= 0) userSheet.getRange(rowIndex + 1, assignedCol + 1).setValue(normalizeStaffModules_(current).join(','));
    syncStaffModules_(email, businessId, current, actor.user.email);
    return { success: true, message: 'Module ' + normalized[0] + ' assigned to ' + email };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function removeModuleFromStaff(email, businessId, moduleName, sessionId) {
  try {
    const actor = staffActor_(sessionId, businessId, ['owner', 'admin']);
    if (!actor.ok) return { success: false, message: actor.message };

    const current = getStaffModules(email, businessId);
    const remaining = current.filter(function(m) { return m !== moduleName; });
    if (remaining.length === current.length) return { success: false, message: 'Assignment not found.' };
    if (!remaining.length) return { success: false, message: 'A staff member must retain at least one module.' };

    const userSheet = getOrCreateUserSheet();
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    const bizCol = headers.indexOf('Business_ID');
    const assignedCol = headers.indexOf('Assigned_Modules');
    const rowIndex = data.findIndex(function(row) {
      return String(row[emailCol] || '').toLowerCase() === String(email || '').toLowerCase() &&
             String(row[bizCol] || '') === String(businessId);
    });
    if (rowIndex < 1) return { success: false, message: 'Staff member not found.' };
    if (assignedCol >= 0) userSheet.getRange(rowIndex + 1, assignedCol + 1).setValue(remaining.join(','));
    syncStaffModules_(email, businessId, remaining, actor.user.email);
    return { success: true, message: 'Module ' + moduleName + ' removed from ' + email };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getStaffDashboard(sessionId) {
  const user = getUserFromSession(sessionId);
  if (!user) return { success: false, message: 'User not found' };
  
  const userRole = user.role || 'staff';
  const isOwnerOrAdmin = userRole === 'owner' || userRole === 'admin';
  
  if (isOwnerOrAdmin) {
    return { success: false, message: 'Use full dashboard' };
  }
  
  const accessibleModules = user.accessibleModules || ['Finance'];
  const dashboardData = {
    user: {
      name: user.name,
      email: user.email,
      role: user.role
    },
    modules: [],
    records: {}
  };
  
  const workspace = getWorkspaceFile(sessionId);
  for (const moduleName of accessibleModules) {
    const module = MODULES[moduleName];
    if (!module) continue;
    
    const sheet = workspace.getSheetByName(module.sheet);
    if (!sheet) continue;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const records = data.slice(1).map(row => {
      const record = {};
      headers.forEach((h, i) => { record[h] = row[i]; });
      return record;
    });
    
    dashboardData.modules.push({
      name: moduleName,
      label: module.label,
      icon: module.icon,
      recordCount: records.length,
      recentRecords: records.slice(0, 5)
    });
    
    dashboardData.records[moduleName] = records;
  }
  
  return { success: true, dashboard: dashboardData };
}



function getAvailableModulesForStaff() {
  return ['Finance', 'Ecommerce', 'Sales', 'CRM', 'HR', 'Logistics', 'Tax', 'Agro', 'Productivity', 'POS', 'Attendance', 'Warehouse'];
}
