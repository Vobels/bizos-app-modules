// ============================================================
// 📁 ClientProvision.gs - Client Provisioning System (COMPLETE)
// ============================================================

function getClientDashboardHtml() {
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Dashboard</title>
  <style>
    * { margin:0; padding:0; box-sizing:border-box; }
    body { font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif; background:#f5f7fa; padding:20px; }
    .header { display:flex; justify-content:space-between; align-items:center; background:white; padding:20px; border-radius:12px; margin-bottom:20px; box-shadow:0 2px 8px rgba(0,0,0,0.05); }
    .header h1 { font-size:20px; color:#1f2937; }
    .logout-btn { padding:8px 16px; background:#ef4444; color:white; border:none; border-radius:8px; cursor:pointer; font-size:14px; }
    .kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:16px; margin-bottom:20px; }
    .kpi-card { background:white; padding:20px; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,0.05); }
    .kpi-card .label { font-size:13px; color:#6b7280; margin-bottom:6px; }
    .kpi-card .value { font-size:24px; font-weight:700; color:#1f2937; }
    .modules { display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:12px; }
    .module-card { background:white; padding:16px; border-radius:10px; text-align:center; box-shadow:0 2px 8px rgba(0,0,0,0.05); }
    .loading { text-align:center; padding:60px; color:#6b7280; }
  </style>
</head>
<body>
  <div class="header">
    <h1><?= clientName ?></h1>
    <button class="logout-btn" onclick="doLogout()">Logout</button>
  </div>
  <div id="content" class="loading">Loading dashboard...</div>
  <script>
    const SESSION_ID = <?!= JSON.stringify(sessionId) ?>;
    function doLogout() {
      google.script.run.proxyLogout(SESSION_ID);
      localStorage.removeItem('bizos_client_session');
      window.location.href = window.location.href.split('?')[0];
    }
    google.script.run
      .withSuccessHandler(function(result) {
        const content = document.getElementById('content');
        if (!result || !result.success) {
          content.innerHTML = '<p style="text-align:center;padding:40px;color:#ef4444;">Session expired. <a href="' + window.location.href.split('?')[0] + '">Login again</a></p>';
          return;
        }
        const d = result.dashboard;
        let html = '<div class="kpis">';
        html += '<div class="kpi-card"><div class="label">Revenue</div><div class="value">$' + (d.kpis.revenue||0).toLocaleString() + '</div></div>';
        html += '<div class="kpi-card"><div class="label">Expenses</div><div class="value">$' + (d.kpis.expenses||0).toLocaleString() + '</div></div>';
        html += '<div class="kpi-card"><div class="label">Profit</div><div class="value">$' + (d.kpis.profit||0).toLocaleString() + '</div></div>';
        html += '<div class="kpi-card"><div class="label">Margin</div><div class="value">' + (d.kpis.margin||0) + '%</div></div>';
        html += '</div><div class="modules">';
        (d.modules||[]).forEach(function(m) {
          html += '<div class="module-card"><div>' + m.label + '</div><div style="font-size:12px;color:#9ca3af;">' + m.count + ' records</div></div>';
        });
        html += '</div>';
        document.getElementById('content').innerHTML = html;
      })
      .withFailureHandler(function() {
        document.getElementById('content').innerHTML = '<p style="text-align:center;padding:40px;color:#ef4444;">Failed to load dashboard.</p>';
      })
      .proxyGetDashboard(SESSION_ID);
  </script>
</body>
</html>`;
}

function createClientSheetInClientDrive(email, clientId, businessName, options) {
  try {
    options = options || {};
    const sheetName = businessName + ' - BizOS Data';
    const newSheet = SpreadsheetApp.create(sheetName);
    const sheetId = newSheet.getId();
    const sheetUrl = newSheet.getUrl();
    
    const modules = {
      'Financial_Data': ['Date', 'Type', 'Category', 'Subcategory', 'Amount', 'Description', 'Status', 'Created_By', 'Created_At'],
      'Ecommerce_Data': ['Order_ID', 'Date', 'Customer_Email', 'Product', 'Quantity', 'Price', 'Status', 'Created_At'],
      'Sales_Data': ['Deal_ID', 'Date', 'Customer_Name', 'Product', 'Amount', 'Stage', 'Created_At'],
      'CRM_Data': ['Contact_ID', 'Name', 'Email', 'Phone', 'Company', 'Status', 'Created_At'],
      'HR_Data': ['Employee_ID', 'Name', 'Position', 'Department', 'Salary', 'Status', 'Created_At'],
      'Logistics_Data': ['Shipment_ID', 'Order_ID', 'Status', 'Location', 'Carrier', 'Created_At'],
      'Tax_Data': ['Tax_ID', 'Date', 'Type', 'Amount', 'Status', 'Created_At'],
      'Agro_Data': ['Record_ID', 'Date', 'Type', 'Activity', 'Quantity', 'Cost', 'Revenue', 'Created_At'],
      'Productivity_Data': ['Record_ID', 'Staff', 'Date', 'Task', 'Hours', 'Created_At'],
      'POS_Data': ['Transaction_ID', 'Date', 'Customer', 'Product', 'Quantity', 'Total', 'Created_At'],
      'Attendance_Data': ['Record_ID', 'Staff', 'Date', 'Clock_In', 'Clock_Out', 'Status', 'Created_At'],
      'Warehouse_Data': ['Item_ID', 'Item_Name', 'SKU', 'Quantity', 'Location', 'Status', 'Created_At']
    };
    
    Object.entries(modules).forEach(([name, headers]) => {
      const sheet = newSheet.insertSheet(name);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f4f6');
    });
    
    const infoSheet = newSheet.insertSheet('Client_Info');
    infoSheet.appendRow(['Property', 'Value']);
    infoSheet.appendRow(['Client_ID', clientId]);
    infoSheet.appendRow(['Business_ID', options.businessId || clientId]);
    infoSheet.appendRow(['Business_Name', businessName]);
    infoSheet.appendRow(['Owner_Email', String(email || '').toLowerCase()]);
    infoSheet.appendRow(['Created_At', new Date().toISOString()]);
    infoSheet.appendRow(['Tier', options.tier || 'sovereign']);
    infoSheet.appendRow(['Status', options.status || 'active']);
    infoSheet.appendRow(['Primary_Color', options.primaryColor || '#2E7D32']);
    infoSheet.appendRow(['Logo_Url', options.logoUrl || '']);
    infoSheet.appendRow(['Provisioning_Version', options.provisioningVersion || '1.0']);
    
    newSheet.addEditor(email);
    const folder = getOrCreateClientFolder(email, businessName);
    const file = DriveApp.getFileById(sheetId);
    file.moveTo(folder);

    // The spreadsheet is the client's live workspace. Keep BizOS as an
    // explicit editor/recovery operator, but transfer ownership when Google
    // Drive permits it. For consumer accounts where ownership transfer
    // requires client consent, create a pending-owner request instead.
    var ownership = transferClientWorkspaceOwnership_(file, email);

    return {
      success:true,
      sheetId:sheetId,
      sheetUrl:sheetUrl,
      folderId:folder.getId(),
      ownershipStatus:ownership.status,
      ownerEmail:ownership.ownerEmail || email,
      ownershipMessage:ownership.message || ''
    };
  } catch (error) {
    console.error('Create sheet error:', error);
    return { success:false, message:error.message };
  }
}

function transferClientWorkspaceOwnership_(file, clientEmail) {
  var email = String(clientEmail || '').trim().toLowerCase();
  if (!file || !email) return {status:'ownership_not_requested',message:'Client email is required.'};

  try {
    // setOwner automatically keeps the previous owner as an editor when the
    // transfer is supported (for example, Workspace accounts in the same org).
    file.addEditor(email);
    file.setOwner(email);
    return {
      status:'client_owned',
      ownerEmail:email,
      message:'Client workspace ownership transferred successfully. BizOS remains an editor for recovery and support.'
    };
  } catch (directTransferError) {
    console.warn('Direct workspace ownership transfer unavailable for '+email+':', directTransferError);

    // Consumer-account ownership transfers require the prospective owner to
    // accept the transfer. The advanced Drive service is already enabled in
    // appsscript.json, so initiate that handoff without blocking provisioning.
    try {
      Drive.Permissions.create(
        {
          role:'writer',
          type:'user',
          emailAddress:email,
          pendingOwner:true
        },
        file.getId(),
        {
          sendNotificationEmail:true
        }
      );
      return {
        status:'ownership_pending_client_acceptance',
        ownerEmail:email,
        message:'Client workspace created and shared. Google Drive requires the client to accept the ownership transfer from their email before ownership changes.'
      };
    } catch (pendingTransferError) {
      console.error('Client workspace ownership handoff failed:', pendingTransferError);
      return {
        status:'client_editor_only',
        ownerEmail:email,
        message:'Client workspace is shared with the client, but Google Drive did not allow automatic ownership transfer. BizOS remains the current owner and recovery backup remains enabled.'
      };
    }
  }
}

function getOrCreateClientFolder(email, businessName) {
  try {
    const folders = DriveApp.getFoldersByName('BizOS_Client_Workspaces');
    let folder;
    if (folders.hasNext()) folder = folders.next();
    else folder = DriveApp.createFolder('BizOS_Client_Workspaces');
    const cleanName = businessName.replace(/[^a-zA-Z0-9]/g, '_');
    const subfolderName = cleanName + '_' + email.split('@')[0];
    const subFolders = folder.getFoldersByName(subfolderName);
    if (subFolders.hasNext()) return subFolders.next();
    return folder.createFolder(subfolderName);
  } catch (error) {
    console.error('Folder error:', error);
    return DriveApp.createFolder('BizOS_' + businessName.replace(/[^a-zA-Z0-9]/g, '_'));
  }
}

function sendClientWelcomeEmail(email, businessName, landingUrl, clientId) {
  try {
    const subject = `🎉 Welcome to ${businessName}'s BizOS Workspace!`;
    const body = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <h1 style="color: #5D2A86;">Welcome to BizOS!</h1>
        <p>Your business <strong>${businessName}</strong> has been successfully set up.</p>
        <h2>Your Branded Landing Page</h2>
        <p>Access your business portal at:</p>
        <p><strong><a href="${landingUrl}" target="_blank">${landingUrl}</a></strong></p>
        <h2>Your Login Credentials</h2>
        <p><strong>Email:</strong> ${email}</p>
        <p><strong>Password:</strong> (The one you used during signup)</p>
        <h2>What's Next?</h2>
        <ul>
          <li>Visit your branded landing page</li>
          <li>Login with your email and password</li>
          <li>Access all 12 business modules</li>
          <li>Your data is stored securely in your own Google Sheet</li>
        </ul>
        <hr>
        <p style="color: #666; font-size: 12px;">This is an automated message from BizOS.</p>
      </div>`;
    MailApp.sendEmail({to:email,subject:subject,htmlBody:body});
    console.log('📧 Welcome email sent to:', email);
  } catch (error) {
    console.error('Welcome email error:', error);
  }
}
