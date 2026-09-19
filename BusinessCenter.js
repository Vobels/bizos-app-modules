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


// ==================== BUSINESS CENTER PART 2: POS ====================

const BUSINESS_CENTER_SALE_HEADERS = [
  'Sale_ID','Date','Customer_ID','Customer_Name','Subtotal','Discount',
  'Total','Amount_Paid','Balance_Due','Payment_Method','Status','Created_By','Created_At'
];
const BUSINESS_CENTER_SALE_ITEM_HEADERS = [
  'Sale_ID','Product_ID','Product_Name','Barcode','Quantity','Unit_Price','Line_Total','Created_At'
];

function getBusinessCenterPOSSheets_(workspace) {
  const sheets = getBusinessCenterSheets_(workspace);
  let sales = workspace.getSheetByName('BusinessCenter_Sales');
  if (!sales) {
    sales = workspace.insertSheet('BusinessCenter_Sales');
    sales.getRange(1, 1, 1, BUSINESS_CENTER_SALE_HEADERS.length).setValues([BUSINESS_CENTER_SALE_HEADERS]);
  }
  let items = workspace.getSheetByName('BusinessCenter_Sale_Items');
  if (!items) {
    items = workspace.insertSheet('BusinessCenter_Sale_Items');
    items.getRange(1, 1, 1, BUSINESS_CENTER_SALE_ITEM_HEADERS.length).setValues([BUSINESS_CENTER_SALE_ITEM_HEADERS]);
  }
  return { products: sheets.products, customers: sheets.customers, sales: sales, items: items };
}

function findBusinessCenterProduct_(productsSheet, code) {
  const q = String(code || '').trim().toLowerCase();
  if (!q || productsSheet.getLastRow() < 2) return null;
  const values = productsSheet.getDataRange().getValues();
  const headers = values[0];
  const nameCol = headers.indexOf('Name');
  const skuCol = headers.indexOf('SKU');
  const barcodeCol = headers.indexOf('Barcode');
  const idCol = headers.indexOf('Product_ID');
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if ([row[idCol], row[nameCol], row[skuCol], row[barcodeCol]].some(function(v) {
      return String(v || '').trim().toLowerCase() === q;
    })) {
      const item = {};
      headers.forEach(function(h, j) { item[h] = row[j]; });
      return item;
    }
  }
  return null;
}

function lookupBusinessCenterProduct(code, sessionId) {
  try {
    const access = businessCenterAccess_(sessionId, false);
    if (!access.ok) return { success: false, message: access.message };
    const workspace = getWorkspaceFile(sessionId);
    const sheets = getBusinessCenterSheets_(workspace);
    const product = findBusinessCenterProduct_(sheets.products, code);
    return product ? { success: true, product: product } : { success: false, message: 'Product not found.' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function completeBusinessCenterSale(data, sessionId) {
  try {
    const access = businessCenterAccess_(sessionId, true);
    if (!access.ok) return { success: false, message: access.message };
    data = data || {};
    const cart = Array.isArray(data.items) ? data.items : [];
    if (!cart.length) return { success: false, message: 'Add at least one product to the sale.' };

    const workspace = getWorkspaceFile(sessionId);
    const sheets = getBusinessCenterPOSSheets_(workspace);
    const productValues = sheets.products.getDataRange().getValues();
    const headers = productValues[0];
    const idCol = headers.indexOf('Product_ID');
    const stockCol = headers.indexOf('Stock_Qty');
    const priceCol = headers.indexOf('Selling_Price');

    const resolved = [];
    let subtotal = 0;

    cart.forEach(function(line) {
      const productId = String(line.Product_ID || '');
      const qty = Number(line.Quantity || 0);
      if (!productId || qty <= 0 || !Number.isFinite(qty)) throw new Error('Invalid sale item.');
      let rowIndex = -1;
      for (let i = 1; i < productValues.length; i++) {
        if (String(productValues[i][idCol]) === productId) { rowIndex = i; break; }
      }
      if (rowIndex < 0) throw new Error('A product in the cart no longer exists.');
      const stock = Number(productValues[rowIndex][stockCol] || 0);
      const price = Number(productValues[rowIndex][priceCol] || 0);
      if (qty > stock) throw new Error('Insufficient stock for ' + productValues[rowIndex][headers.indexOf('Name')] + '. Available: ' + stock);
      const lineTotal = price * qty;
      subtotal += lineTotal;
      resolved.push({ rowIndex: rowIndex, product: productValues[rowIndex], qty: qty, price: price, lineTotal: lineTotal });
    });

    const discount = Math.max(0, Number(data.Discount || 0));
    const total = Math.max(0, subtotal - discount);
    const amountPaid = Math.max(0, Number(data.Amount_Paid || 0));
    if (amountPaid > total) throw new Error('Amount paid cannot be greater than the sale total.');
    const balanceDue = total - amountPaid;
    const saleId = 'SALE-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const now = new Date().toISOString();
    const user = access.user;

    sheets.sales.appendRow(BUSINESS_CENTER_SALE_HEADERS.map(function(h) {
      return h === 'Sale_ID' ? saleId :
        h === 'Date' ? (data.Date || now.split('T')[0]) :
        h === 'Customer_ID' ? (data.Customer_ID || '') :
        h === 'Customer_Name' ? (data.Customer_Name || '') :
        h === 'Subtotal' ? subtotal :
        h === 'Discount' ? discount :
        h === 'Total' ? total :
        h === 'Amount_Paid' ? amountPaid :
        h === 'Balance_Due' ? balanceDue :
        h === 'Payment_Method' ? (data.Payment_Method || 'Cash') :
        h === 'Status' ? (balanceDue > 0 ? 'Credit' : 'Paid') :
        h === 'Created_By' ? (user.email || user.name || '') :
        h === 'Created_At' ? now : '';
    }));

    resolved.forEach(function(line) {
      const name = line.product[headers.indexOf('Name')];
      const barcode = line.product[headers.indexOf('Barcode')];
      sheets.items.appendRow(BUSINESS_CENTER_SALE_ITEM_HEADERS.map(function(h) {
        return h === 'Sale_ID' ? saleId : h === 'Product_ID' ? line.product[idCol] :
          h === 'Product_Name' ? name : h === 'Barcode' ? barcode :
          h === 'Quantity' ? line.qty : h === 'Unit_Price' ? line.price :
          h === 'Line_Total' ? line.lineTotal : h === 'Created_At' ? now : '';
      }));
      sheets.products.getRange(line.rowIndex + 1, stockCol + 1).setValue(Number(line.product[stockCol] || 0) - line.qty);
    });

    // Finance and customer balances are intentionally NOT written here yet.
    // Part 4 will add the accounting/customer-ledger orchestration after POS is validated.
    return {
      success: true, saleId: saleId, subtotal: subtotal, discount: discount,
      total: total, amountPaid: amountPaid, balanceDue: balanceDue,
      message: 'Sale completed successfully.'
    };
  } catch (error) {
    console.error('Business Center sale error:', error);
    return { success: false, message: error.message };
  }
}
