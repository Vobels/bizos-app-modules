// ==================== STAFF INVITATION SYSTEM ====================

function inviteStaffMember(businessId, email, name, role, invitedByEmail, assignedModules) {
  try {
    // Verify inviter
    const userSheet = getOrCreateUserSheet();
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    const emailCol = headers.indexOf('Email');
    const roleCol = headers.indexOf('Role');
    const businessIdCol = headers.indexOf('Business_ID');
    
    const inviterRow = data.find(row => 
      row[emailCol] === invitedByEmail.toLowerCase() && 
      row[businessIdCol] === businessId
    );
    
    if (!inviterRow || (inviterRow[roleCol] !== 'owner' && inviterRow[roleCol] !== 'admin')) {
      return { success: false, message: "Only business owners can invite staff members" };
    }
    
    // For staff role, assignedModules is required
    if (role === 'staff' && !assignedModules) {
      return { success: false, message: "Please assign at least one module for staff member" };
    }
    
    // For staff, modules should be comma-separated string
    const modulesStr = Array.isArray(assignedModules) ? assignedModules.join(',') : assignedModules || '';
    
    const invitationCode = generateInvitationCode();
    
    // Check if user already exists
    const existingUser = data.find(row => row[emailCol] === email.toLowerCase());
    
    if (existingUser) {
      const rowIndex = data.findIndex(row => row[emailCol] === email.toLowerCase());
      const invitedByCol = headers.indexOf('Invited_By');
      const statusCol = headers.indexOf('Invitation_Status');
      const codeCol = headers.indexOf('Invitation_Code');
      const assignedCol = headers.indexOf('Assigned_Modules');
      
      if (invitedByCol !== -1) userSheet.getRange(rowIndex + 1, invitedByCol + 1).setValue(invitedByEmail);
      if (statusCol !== -1) userSheet.getRange(rowIndex + 1, statusCol + 1).setValue('pending');
      if (codeCol !== -1) userSheet.getRange(rowIndex + 1, codeCol + 1).setValue(invitationCode);
      if (assignedCol !== -1) userSheet.getRange(rowIndex + 1, assignedCol + 1).setValue(modulesStr);
    } else {
      const newUser = [
        email.toLowerCase(),
        '',
        name,
        '',
        role,
        businessId,
        '',
        'NO',
        invitedByEmail,
        'pending',
        invitationCode,
        new Date().toISOString(),
        '',
        'YES',
        'NO',
        modulesStr
      ];
      userSheet.appendRow(newUser);
    }
    
    sendInvitationEmail(email, name, invitationCode, role);
    
    return { 
      success: true, 
      message: `Invitation sent to ${email}`,
      invitationCode: invitationCode
    };
    
  } catch (error) {
    console.error('Staff invitation error:', error);
    return { success: false, message: error.message };
  }
}

function acceptInvitation(invitationCode, password, name, phone) {
  try {
    const userSheet = getOrCreateUserSheet();
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    const codeCol = headers.indexOf('Invitation_Code');
    const emailCol = headers.indexOf('Email');
    const statusCol = headers.indexOf('Invitation_Status');
    const businessIdCol = headers.indexOf('Business_ID');
    const nameCol = headers.indexOf('Name');
    const phoneCol = headers.indexOf('Phone');
    const passwordCol = headers.indexOf('Password_Hash');
    
    const userRow = data.find(row => row[codeCol] === invitationCode);
    
    if (!userRow) {
      return { success: false, message: "Invalid invitation code" };
    }
    
    const rowIndex = data.findIndex(row => row[codeCol] === invitationCode);
    const email = userRow[emailCol];
    
    // Update user
    const hashedPassword = hashPassword(password);
    userSheet.getRange(rowIndex + 1, passwordCol + 1).setValue(hashedPassword);
    userSheet.getRange(rowIndex + 1, statusCol + 1).setValue('accepted');
    if (name && nameCol !== -1) userSheet.getRange(rowIndex + 1, nameCol + 1).setValue(name);
    if (phone && phoneCol !== -1) userSheet.getRange(rowIndex + 1, phoneCol + 1).setValue(phone);
    
    // Add to workspace
    const businessId = userRow[businessIdCol];
    const businessSheet = getOrCreateBusinessSheet();
    const businessData = businessSheet.getDataRange().getValues();
    const businessHeaders = businessData[0];
    
    const businessRow = businessData.find(row => row[businessHeaders.indexOf('Business_ID')] === businessId);
    const workspaceId = businessRow[businessHeaders.indexOf('Workspace_ID')];
    
    if (workspaceId) {
      const workspace = SpreadsheetApp.openById(workspaceId);
      const workspaceUsersSheet = workspace.getSheetByName('Workspace_Users');
      if (workspaceUsersSheet) {
        workspaceUsersSheet.appendRow([
          email,
          name || userRow[nameCol],
          userRow[roleCol],
          businessId,
          new Date().toISOString(),
          'active'
        ]);
      }
      workspace.addViewer(email);
    }
    
    return { 
      success: true, 
      message: "Invitation accepted! You can now login."
    };
    
  } catch (error) {
    console.error('Accept invitation error:', error);
    return { success: false, message: error.message };
  }
}

function removeTeamMember(businessId, email) {
  try {
    const userSheet = getOrCreateUserSheet();
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    const emailCol = headers.indexOf('Email');
    const businessIdCol = headers.indexOf('Business_ID');
    const roleCol = headers.indexOf('Role');
    
    // Don't allow removing owner
    const rowIndex = data.findIndex(r => r[emailCol] === email.toLowerCase() && r[businessIdCol] === businessId);
    if (rowIndex === -1) return { success: false, message: "User not found" };
    
    if (data[rowIndex][roleCol] === 'owner') {
      return { success: false, message: "Cannot remove business owner" };
    }
    
    userSheet.deleteRow(rowIndex + 1);
    return { success: true };
  } catch(error) {
    return { success: false, message: error.message };
  }
}

function getBusinessStaff(businessId) {
  const userSheet = getOrCreateUserSheet();
  const data = userSheet.getDataRange().getValues();
  const headers = data[0];
  const businessIdCol = headers.indexOf('Business_ID');
  
  return data.slice(1)
    .filter(row => row[businessIdCol] === businessId)
    .map(row => ({
      email: row[headers.indexOf('Email')],
      name: row[headers.indexOf('Name')],
      role: row[headers.indexOf('Role')],
      isVerified: row[headers.indexOf('Is_Verified')],
      createdAt: row[headers.indexOf('Created_At')]
    }));
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
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Staff_Modules');
    if (!sheet) {
      createStaffModulesSheet();
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    const headers = data[0];
    const emailCol = headers.indexOf('Staff_Email');
    const bizIdCol = headers.indexOf('Business_ID');
    const moduleCol = headers.indexOf('Module_Access');
    
    const modules = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol] === email.toLowerCase() && data[i][bizIdCol] === businessId) {
        const moduleName = data[i][moduleCol];
        if (moduleName) modules.push(moduleName);
      }
    }
    
    return modules;
  } catch (error) {
    console.error('Error getting staff modules:', error);
    return [];
  }
}

function assignModuleToStaff(email, businessId, moduleName, assignedBy) {
  try {
    const sheet = createStaffModulesSheet();
    sheet.appendRow([
      email.toLowerCase(),
      businessId,
      moduleName,
      assignedBy || 'system',
      new Date().toISOString()
    ]);
    return { success: true, message: `Module ${moduleName} assigned to ${email}` };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function removeModuleFromStaff(email, businessId, moduleName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Staff_Modules');
    if (!sheet) return { success: false, message: 'Staff_Modules sheet not found' };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Staff_Email');
    const bizIdCol = headers.indexOf('Business_ID');
    const moduleCol = headers.indexOf('Module_Access');
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][emailCol] === email.toLowerCase() && 
          data[i][bizIdCol] === businessId && 
          data[i][moduleCol] === moduleName) {
        sheet.deleteRow(i + 1);
        return { success: true, message: `Module ${moduleName} removed from ${email}` };
      }
    }
    return { success: false, message: 'Assignment not found' };
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
