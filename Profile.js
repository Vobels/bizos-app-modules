//profile.gs
function updateUserProfile(businessId, name, phone, sessionId) {  // ADD sessionId
  try {
    // Verify session first
    const user = getUserFromSession(sessionId);
    if (!user) return { success: false, message: "Unauthorized" };
    
    const userSheet = getOrCreateUserSheet();
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    const businessIdCol = headers.indexOf('Business_ID');
    const nameCol = headers.indexOf('Name');
    const phoneCol = headers.indexOf('Phone');
    
    const rowIndex = data.findIndex(r => r[businessIdCol] === businessId);
    
    if (rowIndex === -1) return { success: false, message: "User not found" };
    
    if (nameCol !== -1) userSheet.getRange(rowIndex + 1, nameCol + 1).setValue(name);
    if (phoneCol !== -1) userSheet.getRange(rowIndex + 1, phoneCol + 1).setValue(phone);
    
    return { success: true };
  } catch(error) {
    return { success: false, message: error.message };
  }
}

function changeUserPassword(email, currentPassword, newPassword, sessionId) {  // ADD sessionId
  try {
    // Verify session
    const user = getUserFromSession(sessionId);
    if (!user) return { success: false, message: "Unauthorized" };
    
    const userSheet = getOrCreateUserSheet();
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    const emailCol = headers.indexOf('Email');
    const passwordCol = headers.indexOf('Password_Hash');
    const isDemoCol = headers.indexOf('Is_Demo');
    
    const rowIndex = data.findIndex(r => r[emailCol] === email.toLowerCase());
    if (rowIndex === -1) return { success: false, message: "User not found" };
    
    if (data[rowIndex][isDemoCol] === 'YES') {
      return { success: false, message: "Demo accounts cannot change password" };
    }
    
    const hashedCurrent = hashPassword(currentPassword);
    if (data[rowIndex][passwordCol] !== hashedCurrent) {
      return { success: false, message: "Current password is incorrect" };
    }
    
    const hashedNew = hashPassword(newPassword);
    userSheet.getRange(rowIndex + 1, passwordCol + 1).setValue(hashedNew);
    
    return { success: true };
  } catch(error) {
    return { success: false, message: error.message };
  }
}

function uploadProfilePhoto(email, imageData, sessionId) {  // ADD sessionId
  try {
    // Verify session
    const user = getUserFromSession(sessionId);
    if (!user) return { success: false, message: "Unauthorized" };
    
    // Rest of your existing code...
    let folder;
    const folders = DriveApp.getFoldersByName('BizOS_ProfilePhotos');
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder('BizOS_ProfilePhotos');
    }
    
    const oldFiles = folder.getFilesByName(`${email}_profile.jpg`);
    while (oldFiles.hasNext()) {
      oldFiles.next().setTrashed(true);
    }
    
    const base64Data = imageData.split(',')[1];
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', `${email}_profile.jpg`);
    const file = folder.createFile(blob);
    
    const userSheet = getOrCreateUserSheet();
    const data = userSheet.getDataRange().getValues();
    const headers = data[0];
    
    const emailCol = headers.indexOf('Email');
    let photoCol = headers.indexOf('Profile_Photo_ID');
    
    if (photoCol === -1) {
      userSheet.getRange(1, userSheet.getLastColumn() + 1).setValue('Profile_Photo_ID');
      photoCol = userSheet.getLastColumn() - 1;
    }
    
    const rowIndex = data.findIndex(r => r[emailCol] === email.toLowerCase());
    if (rowIndex !== -1) {
      userSheet.getRange(rowIndex + 1, photoCol + 1).setValue(file.getId());
    }
    
    return { success: true };
  } catch(error) {
    return { success: false, message: error.message };
  }
}

function requestDataExport(email, businessId, sessionId) {  // ADD sessionId
  try {
    // Verify session
    const user = getUserFromSession(sessionId);
    if (!user) return { success: false, message: "Unauthorized" };
    
    const subject = `BizOS Data Export Request - ${businessId}`;
    const body = `Data export requested for ${email} on ${new Date().toISOString()}. Processing will begin shortly.`;
    GmailApp.sendEmail(EMAIL_CONFIG.supportEmail, subject, body);
    return { success: true };
  } catch(error) {
    return { success: false };
  }
}