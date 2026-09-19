// ============================================================
// BusinessCenter.js - Part 1 foundation
// ============================================================

const BUSINESS_CENTER_PRODUCT_HEADERS = [
  'Product_ID','Name','SKU','Barcode','Image_File_ID','Cost_Price',
  'Selling_Price','Stock_Qty','Minimum_Stock','Unit','Category',
  'Supplier','Location','Created_By','Created_At'
];

const BUSINESS_CENTER_CUSTOMER_HEADERS = [
  'Customer_ID','Name','Phone','Email','Address','Balance',
  'Created_By','Created_At'
];

function businessCenterAccess_(sessionId, write) {
  const user = getUserFromSession(sessionId);
  if (!user) return { ok: false, message: 'Session expired. Please login again.' };
  if (!userHasAccess('Ecommerce', sessionId)) {
    return { ok: false, message: 'Business Center is not available for this account.' };
  }
  if (write) {
    const role = user.role || 'staff';
    if (role !== 'owner' && role !== 'admin') {
      return { ok: false, message: 'You have view-only access to Business Center.' };
    }
    if (user.isDemo === 'YES') {
      return { ok: false, message: 'Demo accounts are view-only.' };
    }
  }
  return { ok: true, user: user };
}

function getBusinessCenterSheets_(workspace) {
  let products = workspace.getSheetByName('BusinessCenter_Products');
  if (!products) {
    products = workspace.insertSheet('BusinessCenter_Products');
    products.getRange(1, 1, 1, BUSINESS_CENTER_PRODUCT_HEADERS.length).setValues([BUSINESS_CENTER_PRODUCT_HEADERS]);
  }
  let customers = workspace.getSheetByName('BusinessCenter_Customers');
  if (!customers) {
    customers = workspace.insertSheet('BusinessCenter_Customers');
    customers.getRange(1, 1, 1, BUSINESS_CENTER_CUSTOMER_HEADERS.length).setValues([BUSINESS_CENTER_CUSTOMER_HEADERS]);
  }
  return { products: products, customers: customers };
}

function businessCenterRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1).map(function(row) {
    const item = {};
    headers.forEach(function(header, i) { item[header] = row[i]; });
    return item;
  });
}

function businessCenterImageData_(fileId) {
  if (!fileId) return '';
  try {
    const file = DriveApp.getFileById(String(fileId));
    const thumb = file.getThumbnail();
    if (!thumb) return '';
    return 'data:' + (thumb.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(thumb.getBytes());
  } catch (e) {
    return '';
  }
}

function getBusinessCenterData(sessionId) {
  try {
    const access = businessCenterAccess_(sessionId, false);
    if (!access.ok) return { success: false, message: access.message };
    const workspace = getWorkspaceFile(sessionId);
    const sheets = getBusinessCenterSheets_(workspace);
    const products = businessCenterRows_(sheets.products);
    const customers = businessCenterRows_(sheets.customers);
    products.forEach(function(product) { product.Image_Data = businessCenterImageData_(product.Image_File_ID); });
    const lowStock = products.filter(function(p) {
      return Number(p.Stock_Qty || 0) <= Number(p.Minimum_Stock || 0);
    }).length;
    return {
      success: true, products: products, customers: customers,
      summary: {
        products: products.length,
        customers: customers.length,
        stockUnits: products.reduce(function(total, p) { return total + (Number(p.Stock_Qty) || 0); }, 0),
        lowStock: lowStock
      }
    };
  } catch (error) {
    console.error('Business Center load error:', error);
    return { success: false, message: error.message };
  }
}

function businessCenterImageFolder_(workspace) {
  const workspaceFile = DriveApp.getFileById(workspace.getId());
  let parent = null;
  const parents = workspaceFile.getParents();
  if (parents.hasNext()) parent = parents.next();
  if (!parent) parent = DriveApp.getRootFolder();
  const folders = parent.getFoldersByName('BizOS Business Center Product Images');
  return folders.hasNext() ? folders.next() : parent.createFolder('BizOS Business Center Product Images');
}

function saveBusinessCenterProduct(data, sessionId) {
  try {
    const access = businessCenterAccess_(sessionId, true);
    if (!access.ok) return { success: false, message: access.message };
    data = data || {};
    if (!String(data.Name || '').trim()) return { success: false, message: 'Product name is required.' };

    const workspace = getWorkspaceFile(sessionId);
    const sheets = getBusinessCenterSheets_(workspace);
    const productId = 'PRD-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    let imageFileId = '';

    if (data.Image_Data) {
      const match = String(data.Image_Data).match(/^data:([^;]+);base64,(.+)$/);
      if (!match) return { success: false, message: 'Invalid product image.' };
      const bytes = Utilities.base64Decode(match[2]);
      if (bytes.length > 800000) return { success: false, message: 'Product image is too large. Please use an image under 800 KB.' };
      const safeName = String(data.Name).replace(/[^a-z0-9-_ ]/gi, '').trim().replace(/\s+/g, '-').substring(0, 60) || productId;
      const blob = Utilities.newBlob(bytes, match[1], safeName + '-' + productId);
      imageFileId = businessCenterImageFolder_(workspace).createFile(blob).getId();
    }

    const user = access.user;
    const row = BUSINESS_CENTER_PRODUCT_HEADERS.map(function(header) {
      if (header === 'Product_ID') return productId;
      if (header === 'Image_File_ID') return imageFileId;
      if (header === 'Created_By') return user.email || user.name || '';
      if (header === 'Created_At') return new Date().toISOString();
      const value = data[header];
      return value === undefined || value === null ? '' : value;
    });
    sheets.products.appendRow(row);
    return { success: true, productId: productId, message: 'Product added successfully.' };
  } catch (error) {
    console.error('Business Center product save error:', error);
    return { success: false, message: error.message };
  }
}

function saveBusinessCenterCustomer(data, sessionId) {
  try {
    const access = businessCenterAccess_(sessionId, true);
    if (!access.ok) return { success: false, message: access.message };
    data = data || {};
    if (!String(data.Name || '').trim()) return { success: false, message: 'Customer name is required.' };

    const workspace = getWorkspaceFile(sessionId);
    const sheets = getBusinessCenterSheets_(workspace);
    const customerId = 'CUS-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const user = access.user;
    const row = BUSINESS_CENTER_CUSTOMER_HEADERS.map(function(header) {
      if (header === 'Customer_ID') return customerId;
      if (header === 'Balance') return Number(data.Balance || 0);
      if (header === 'Created_By') return user.email || user.name || '';
      if (header === 'Created_At') return new Date().toISOString();
      const value = data[header];
      return value === undefined || value === null ? '' : value;
    });
    sheets.customers.appendRow(row);
    return { success: true, customerId: customerId, message: 'Customer added successfully.' };
  } catch (error) {
    console.error('Business Center customer save error:', error);
    return { success: false, message: error.message };
  }
}
