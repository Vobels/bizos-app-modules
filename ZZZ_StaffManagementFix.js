// ============================================================
// BizOS — Staff Management / Invitation Backend
// This is intentionally isolated from the legacy Staff.js flow.
// ============================================================

function zzzStaffEnsureColumns_(sheet, requiredHeaders) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  requiredHeaders.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }
  });
  return headers;
}

function zzzStaffCol_(headers, name) {
  return headers.indexOf(name);
}

function zzzStaffEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function zzzStaffAuthorizeManager_(businessId, sessionId) {
  const user = getUserFromSession(sessionId);
  if (!user) return { ok: false, message: 'Your session has expired. Please login again.' };

  const role = String(user.role || '').toLowerCase();
  if (role !== 'owner' && role !== 'admin') {
    return { ok: false, message: 'Only business owners and administrators can manage staff.' };
  }
  if (String(user.businessId || '') !== String(businessId || '')) {
    return { ok: false, message: 'You are not authorized to manage this business.' };
  }
  return { ok: true, user: user };
}

function zzzStaffWriteInvitation_(sheet, headers, rowNumber, values) {
  Object.keys(values).forEach(function(key) {
    const col = zzzStaffCol_(headers, key);
    if (col !== -1) sheet.getRange(rowNumber, col + 1).setValue(values[key]);
  });
}

function inviteStaffMemberV2(businessId, email, name, role, assignedModules, sessionId) {
  try {
    const auth = zzzStaffAuthorizeManager_(businessId, sessionId);
    if (!auth.ok) return { success: false, message: auth.message };

    email = zzzStaffEmail_(email);
    name = String(name || '').trim();
    role = String(role || 'staff').trim().toLowerCase();
    assignedModules = Array.isArray(assignedModules) ? assignedModules : [];

    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return { success: false, message: 'Please enter a valid email address.' };
    }
    if (!name) return { success: false, message: 'Please enter the team member name.' };
    if (role !== 'staff' && role !== 'admin') {
      return { success: false, message: 'Invalid staff role.' };
    }
    if (role === 'staff' && assignedModules.length === 0) {
      return { success: false, message: 'Select at least one module for this staff member.' };
    }

    const allowedModules = getAvailableModulesForStaff();
    assignedModules = assignedModules.filter(function(module) {
      return allowedModules.indexOf(module) !== -1;
    });
    if (role === 'staff' && assignedModules.length === 0) {
      return { success: false, message: 'Select at least one valid module.' };
    }

    const sheet = getOrCreateUserSheet();
    const headers = zzzStaffEnsureColumns_(sheet, [
      'Email', 'Password_Hash', 'Name', 'Phone', 'Role', 'Business_ID',
      'Business_Name', 'Is_Demo', 'Invited_By', 'Invitation_Status',
      'Invitation_Code', 'Invitation_Expires_At', 'Created_At', 'Last_Login',
      'Is_Verified', 'Is_Client', 'Assigned_Modules'
    ]);

    const data = sheet.getDataRange().getValues();
    const emailCol = zzzStaffCol_(headers, 'Email');
    const businessCol = zzzStaffCol_(headers, 'Business_ID');
    const statusCol = zzzStaffCol_(headers, 'Invitation_Status');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (zzzStaffEmail_(data[i][emailCol]) === email && String(data[i][businessCol] || '') === String(businessId || '')) {
        rowIndex = i;
        break;
      }
    }

    // Do not silently overwrite an existing active member.
    if (rowIndex !== -1) {
      const existingStatus = String(data[rowIndex][statusCol] || '').toLowerCase();
      if (existingStatus === 'accepted' || existingStatus === 'active' || existingStatus === '') {
        return { success: false, message: 'This email is already a team member.' };
      }
      if (existingStatus === 'pending') {
        return { success: false, message: 'An invitation is already pending for this email.' };
      }
    } else {
      // Reject an email already belonging to another business.
      for (let i = 1; i < data.length; i++) {
        if (zzzStaffEmail_(data[i][emailCol]) === email && String(data[i][businessCol] || '') !== String(businessId || '')) {
          return { success: false, message: 'This email is already registered with another business.' };
        }
      }
    }

    const invitationCode = Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase();
    const expiresAt = new Date();
    expiresAt.setDate(expiresAt.getDate() + 7);
    const modulesString = assignedModules.join(',');
    const businessName = auth.user.businessName || auth.user.Business_Name || '';
    const invitedBy = zzzStaffEmail_(auth.user.email);

    if (rowIndex === -1) {
      const row = new Array(headers.length).fill('');
      function put(header, value) {
        const col = zzzStaffCol_(headers, header);
        if (col !== -1) row[col] = value;
      }
      put('Email', email);
      put('Name', name);
      put('Role', role);
      put('Business_ID', businessId);
      put('Business_Name', businessName);
      put('Is_Demo', 'NO');
      put('Invited_By', invitedBy);
      put('Invitation_Status', 'pending');
      put('Invitation_Code', invitationCode);
      put('Invitation_Expires_At', expiresAt.toISOString());
      put('Created_At', new Date().toISOString());
      put('Is_Verified', 'NO');
      put('Is_Client', 'NO');
      put('Assigned_Modules', modulesString);
      sheet.appendRow(row);
    } else {
      zzzStaffWriteInvitation_(sheet, headers, rowIndex + 1, {
        Name: name,
        Role: role,
        Business_Name: businessName,
        Invited_By: invitedBy,
        Invitation_Status: 'pending',
        Invitation_Code: invitationCode,
        Invitation_Expires_At: expiresAt.toISOString(),
        Assigned_Modules: modulesString,
        Password_Hash: '',
        Is_Verified: 'NO'
      });
    }

    // Keep the dedicated module sheet in sync with the invitation.
    const moduleSheet = createStaffModulesSheet();
    const moduleData = moduleSheet.getDataRange().getValues();
    for (let i = moduleData.length - 1; i >= 1; i--) {
      if (zzzStaffEmail_(moduleData[i][0]) === email && String(moduleData[i][1]) === String(businessId)) {
        moduleSheet.deleteRow(i + 1);
      }
    }
    assignedModules.forEach(function(moduleName) {
      moduleSheet.appendRow([email, businessId, moduleName, invitedBy, new Date().toISOString()]);
    });

    const sent = sendInvitationEmail(email, name, invitationCode, role);
    if (!sent) {
      return { success: false, message: 'The invitation was created, but the email could not be sent. Please try again.' };
    }

    return { success: true, message: 'Invitation sent to ' + email };
  } catch (error) {
    console.error('inviteStaffMemberV2 error:', error);
    return { success: false, message: error.message || 'Failed to send invitation.' };
  }
}

function getBusinessStaffV2(businessId, sessionId) {
  try {
    const auth = zzzStaffAuthorizeManager_(businessId, sessionId);
    if (!auth.ok) return { success: false, message: auth.message, members: [] };

    const sheet = getOrCreateUserSheet();
    const headers = zzzStaffEnsureColumns_(sheet, [
      'Email', 'Name', 'Role', 'Business_ID', 'Invitation_Status',
      'Invitation_Expires_At', 'Is_Verified', 'Created_At', 'Assigned_Modules'
    ]);
    const data = sheet.getDataRange().getValues();
    const businessCol = zzzStaffCol_(headers, 'Business_ID');

    const members = data.slice(1)
      .filter(function(row) { return String(row[businessCol] || '') === String(businessId || ''); })
      .map(function(row) {
        const get = function(header) {
          const col = zzzStaffCol_(headers, header);
          return col === -1 ? '' : row[col];
        };
        const rawStatus = String(get('Invitation_Status') || '').toLowerCase();
        const verified = String(get('Is_Verified') || '').toUpperCase() === 'YES';
        const status = rawStatus === 'pending' ? 'pending' : (verified || rawStatus === 'accepted' || rawStatus === 'active' ? 'active' : rawStatus || 'active');
        return {
          email: get('Email'),
          name: get('Name'),
          role: get('Role') || 'staff',
          status: status,
          invitationExpiresAt: get('Invitation_Expires_At'),
          isVerified: verified,
          createdAt: get('Created_At'),
          assignedModules: String(get('Assigned_Modules') || '').split(',').map(function(x){ return x.trim(); }).filter(Boolean)
        };
      });

    return { success: true, members: members };
  } catch (error) {
    console.error('getBusinessStaffV2 error:', error);
    return { success: false, message: error.message || 'Failed to load staff.', members: [] };
  }
}

function removeTeamMemberV2(businessId, email, sessionId) {
  try {
    const auth = zzzStaffAuthorizeManager_(businessId, sessionId);
    if (!auth.ok) return { success: false, message: auth.message };

    email = zzzStaffEmail_(email);
    const sheet = getOrCreateUserSheet();
    const headers = sheet.getDataRange().getValues()[0];
    const data = sheet.getDataRange().getValues();
    const emailCol = zzzStaffCol_(headers, 'Email');
    const businessCol = zzzStaffCol_(headers, 'Business_ID');
    const roleCol = zzzStaffCol_(headers, 'Role');

    for (let i = 1; i < data.length; i++) {
      if (zzzStaffEmail_(data[i][emailCol]) === email && String(data[i][businessCol] || '') === String(businessId || '')) {
        if (String(data[i][roleCol] || '').toLowerCase() === 'owner') {
          return { success: false, message: 'The business owner cannot be removed.' };
        }
        sheet.deleteRow(i + 1);

        const moduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Staff_Modules');
        if (moduleSheet) {
          const moduleData = moduleSheet.getDataRange().getValues();
          for (let j = moduleData.length - 1; j >= 1; j--) {
            if (zzzStaffEmail_(moduleData[j][0]) === email && String(moduleData[j][1]) === String(businessId || '')) {
              moduleSheet.deleteRow(j + 1);
            }
          }
        }
        return { success: true, message: 'Team member removed.' };
      }
    }
    return { success: false, message: 'Team member not found.' };
  } catch (error) {
    console.error('removeTeamMemberV2 error:', error);
    return { success: false, message: error.message || 'Failed to remove team member.' };
  }
}

function resendStaffInvitationV2(businessId, email, sessionId) {
  try {
    const auth = zzzStaffAuthorizeManager_(businessId, sessionId);
    if (!auth.ok) return { success: false, message: auth.message };

    email = zzzStaffEmail_(email);
    const sheet = getOrCreateUserSheet();
    const headers = sheet.getDataRange().getValues()[0];
    const data = sheet.getDataRange().getValues();
    const emailCol = zzzStaffCol_(headers, 'Email');
    const businessCol = zzzStaffCol_(headers, 'Business_ID');
    const statusCol = zzzStaffCol_(headers, 'Invitation_Status');
    const nameCol = zzzStaffCol_(headers, 'Name');
    const roleCol = zzzStaffCol_(headers, 'Role');

    for (let i = 1; i < data.length; i++) {
      if (zzzStaffEmail_(data[i][emailCol]) === email && String(data[i][businessCol] || '') === String(businessId || '')) {
        if (String(data[i][statusCol] || '').toLowerCase() !== 'pending') {
          return { success: false, message: 'This team member does not have a pending invitation.' };
        }
        const code = Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase();
        const expiresAt = new Date();
        expiresAt.setDate(expiresAt.getDate() + 7);
        zzzStaffWriteInvitation_(sheet, headers, i + 1, {
          Invitation_Code: code,
          Invitation_Expires_At: expiresAt.toISOString()
        });
        const sent = sendInvitationEmail(email, data[i][nameCol], code, data[i][roleCol] || 'staff');
        return sent
          ? { success: true, message: 'Invitation resent to ' + email }
          : { success: false, message: 'The invitation could not be emailed.' };
      }
    }
    return { success: false, message: 'Pending invitation not found.' };
  } catch (error) {
    console.error('resendStaffInvitationV2 error:', error);
    return { success: false, message: error.message || 'Failed to resend invitation.' };
  }
}

function acceptInvitationV2(invitationCode, password, name, phone) {
  try {
    invitationCode = String(invitationCode || '').trim().toUpperCase();
    if (!invitationCode || !password) return { success: false, message: 'Invitation code and password are required.' };
    if (String(password).length < 6) return { success: false, message: 'Password must be at least 6 characters.' };

    const sheet = getOrCreateUserSheet();
    const headers = zzzStaffEnsureColumns_(sheet, [
      'Email', 'Password_Hash', 'Name', 'Phone', 'Role', 'Business_ID',
      'Invitation_Status', 'Invitation_Code', 'Invitation_Expires_At', 'Is_Verified'
    ]);
    const data = sheet.getDataRange().getValues();
    const codeCol = zzzStaffCol_(headers, 'Invitation_Code');
    const statusCol = zzzStaffCol_(headers, 'Invitation_Status');

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][codeCol] || '').trim().toUpperCase() === invitationCode) {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex === -1) return { success: false, message: 'Invalid invitation code.' };

    const row = data[rowIndex];
    if (String(row[statusCol] || '').toLowerCase() !== 'pending') {
      return { success: false, message: 'This invitation has already been accepted or is no longer active.' };
    }

    const expiryCol = zzzStaffCol_(headers, 'Invitation_Expires_At');
    if (expiryCol !== -1 && row[expiryCol]) {
      const expiry = new Date(row[expiryCol]);
      if (!isNaN(expiry.getTime()) && expiry.getTime() < Date.now()) {
        return { success: false, message: 'This invitation has expired. Ask the business owner to resend it.' };
      }
    }

    const email = zzzStaffEmail_(row[zzzStaffCol_(headers, 'Email')]);
    const businessId = row[zzzStaffCol_(headers, 'Business_ID')];
    const finalName = String(name || row[zzzStaffCol_(headers, 'Name')] || email).trim();
    const passwordHash = hashPassword(password);

    zzzStaffWriteInvitation_(sheet, headers, rowIndex + 1, {
      Password_Hash: passwordHash,
      Name: finalName,
      Phone: String(phone || '').trim(),
      Invitation_Status: 'accepted',
      Invitation_Code: '',
      Invitation_Expires_At: '',
      Is_Verified: 'YES'
    });

    // Add the accepted member to the business workspace once.
    const businessSheet = getOrCreateBusinessSheet();
    const businessData = businessSheet.getDataRange().getValues();
    const businessHeaders = businessData[0];
    const bizIdCol = zzzStaffCol_(businessHeaders, 'Business_ID');
    const workspaceIdCol = zzzStaffCol_(businessHeaders, 'Workspace_ID');
    const businessRow = businessData.find(function(r){ return bizIdCol !== -1 && String(r[bizIdCol] || '') === String(businessId || ''); });
    const workspaceId = businessRow && workspaceIdCol !== -1 ? businessRow[workspaceIdCol] : '';

    if (workspaceId) {
      const workspace = SpreadsheetApp.openById(workspaceId);
      let usersSheet = workspace.getSheetByName('Workspace_Users');
      if (!usersSheet) {
        usersSheet = workspace.insertSheet('Workspace_Users');
        usersSheet.appendRow(['Email', 'Name', 'Role', 'Business_ID', 'Added_At', 'Status']);
      }
      const workspaceData = usersSheet.getDataRange().getValues();
      const alreadyThere = workspaceData.slice(1).some(function(r){ return zzzStaffEmail_(r[0]) === email; });
      if (!alreadyThere) {
        usersSheet.appendRow([email, finalName, row[zzzStaffCol_(headers, 'Role')] || 'staff', businessId, new Date().toISOString(), 'active']);
      }
      try { workspace.addViewer(email); } catch (permissionError) { console.warn('Workspace viewer permission:', permissionError.message); }
    }

    return { success: true, message: 'Invitation accepted! You can now login.' };
  } catch (error) {
    console.error('acceptInvitationV2 error:', error);
    return { success: false, message: error.message || 'Failed to accept invitation.' };
  }
}
