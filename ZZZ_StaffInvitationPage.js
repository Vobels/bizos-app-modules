// ============================================================
// BizOS — Dedicated Staff Invitation Page Support
// Public invitation details are only exposed when the invitation
// code itself is supplied. No password is returned.
// ============================================================

function getInvitationDetailsV2(invitationCode) {
  try {
    invitationCode = String(invitationCode || '').trim().toUpperCase();
    if (!invitationCode) return { success: false, message: 'Invitation code is missing.' };

    const sheet = getOrCreateUserSheet();
    const headers = zzzStaffEnsureColumns_(sheet, [
      'Email', 'Name', 'Role', 'Business_ID', 'Business_Name',
      'Invitation_Status', 'Invitation_Code', 'Invitation_Expires_At'
    ]);
    const data = sheet.getDataRange().getValues();
    const codeCol = zzzStaffCol_(headers, 'Invitation_Code');

    let row = null;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][codeCol] || '').trim().toUpperCase() === invitationCode) {
        row = data[i];
        break;
      }
    }

    if (!row) return { success: false, message: 'This invitation link is invalid.' };

    const get = function(header) {
      const col = zzzStaffCol_(headers, header);
      return col === -1 ? '' : row[col];
    };

    if (String(get('Invitation_Status') || '').trim().toLowerCase() !== 'pending') {
      return { success: false, message: 'This invitation has already been accepted or is no longer active.' };
    }

    const expiresAt = get('Invitation_Expires_At');
    if (expiresAt) {
      const expiry = new Date(expiresAt);
      if (!isNaN(expiry.getTime()) && expiry.getTime() < Date.now()) {
        return { success: false, message: 'This invitation has expired. Ask the business owner to resend it.' };
      }
    }

    return {
      success: true,
      invitation: {
        name: String(get('Name') || '').trim(),
        email: String(get('Email') || '').trim().toLowerCase(),
        role: String(get('Role') || 'staff').trim().toLowerCase(),
        businessName: String(get('Business_Name') || '').trim(),
        expiresAt: expiresAt || ''
      }
    };
  } catch (error) {
    console.error('getInvitationDetailsV2 error:', error);
    return { success: false, message: error.message || 'Unable to load this invitation.' };
  }
}
