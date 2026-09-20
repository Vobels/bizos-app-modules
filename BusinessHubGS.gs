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



// ==================== BUSINESS CENTER PART 1: SAFE EXISTING-DATA CONNECTION ====================
// Read-only adapter: lets Business Center see what existing Ecommerce/CRM modules contain
// without changing, duplicating, or rewriting those module records.

function getBusinessCenterConnectedData(sessionId) {
  try {
    const access = businessCenterAccess_(sessionId, false);
    if (!access.ok) return { success: false, message: access.message };

    const result = {
      success: true,
      sources: {
        ecommerce: { connected: false, count: 0, records: [] },
        crm: { connected: false, count: 0, records: [] }
      }
    };

    try {
      const ecommerce = getModuleData('Ecommerce', sessionId);
      if (Array.isArray(ecommerce)) {
        result.sources.ecommerce.connected = true;
        result.sources.ecommerce.count = ecommerce.length;
        result.sources.ecommerce.records = ecommerce.slice(0, 100);
      }
    } catch (e) {
      console.log('Business Center Ecommerce connection skipped:', e.message);
    }

    try {
      const crm = getModuleData('CRM', sessionId);
      if (Array.isArray(crm)) {
        result.sources.crm.connected = true;
        result.sources.crm.count = crm.length;
        result.sources.crm.records = crm.slice(0, 100);
      }
    } catch (e) {
      console.log('Business Center CRM connection skipped:', e.message);
    }

    return result;
  } catch (error) {
    console.error('Business Center connected-data error:', error);
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
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return { success: false, message: 'Another sale is being processed. Please try again.' };

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
    const nameCol = headers.indexOf('Name');
    const stockCol = headers.indexOf('Stock_Qty');
    const priceCol = headers.indexOf('Selling_Price');

    const customerId = String(data.Customer_ID || '').trim();
    const customerName = String(data.Customer_Name || '').trim();
    let customerRow = -1;
    let customerBalance = 0;

    if (customerId) {
      const customerValues = sheets.customers.getDataRange().getValues();
      const customerHeaders = customerValues[0];
      const customerIdCol = customerHeaders.indexOf('Customer_ID');
      const customerNameCol = customerHeaders.indexOf('Name');
      const customerBalanceCol = customerHeaders.indexOf('Balance');
      for (let i = 1; i < customerValues.length; i++) {
        if (String(customerValues[i][customerIdCol] || '') === customerId) {
          customerRow = i;
          customerBalance = Number(customerValues[i][customerBalanceCol] || 0);
          if (!customerName) data.Customer_Name = customerValues[i][customerNameCol] || '';
          break;
        }
      }
      if (customerRow < 1) return { success: false, message: 'The selected customer could not be found. Please select the customer again.' };
    }

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
      if (rowIndex < 0) throw new Error('A product in the sale no longer exists.');
      const stock = Number(productValues[rowIndex][stockCol] || 0);
      const price = Number(productValues[rowIndex][priceCol] || 0);
      if (qty > stock) throw new Error('Insufficient stock for ' + productValues[rowIndex][nameCol] + '. Available: ' + stock);
      const lineTotal = price * qty;
      subtotal += lineTotal;
      resolved.push({ rowIndex: rowIndex, product: productValues[rowIndex], qty: qty, price: price, lineTotal: lineTotal });
    });

    const discount = Math.max(0, Number(data.Discount || 0));
    if (!Number.isFinite(discount) || discount > subtotal) throw new Error('Discount cannot be greater than the sale subtotal.');
    const total = Math.max(0, subtotal - discount);
    const amountPaid = Math.max(0, Number(data.Amount_Paid || 0));
    if (!Number.isFinite(amountPaid) || amountPaid > total) throw new Error('Amount received cannot be greater than the sale total.');
    const balanceDue = total - amountPaid;

    if (balanceDue > 0 && !customerId) {
      return { success: false, message: 'Please select a customer before completing a sale with an outstanding balance.' };
    }

    const saleId = 'SALE-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const now = new Date().toISOString();
    const user = access.user;

    sheets.sales.appendRow(BUSINESS_CENTER_SALE_HEADERS.map(function(h) {
      return h === 'Sale_ID' ? saleId :
        h === 'Date' ? (data.Date || now.split('T')[0]) :
        h === 'Customer_ID' ? customerId :
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
      const name = line.product[nameCol];
      const barcode = line.product[headers.indexOf('Barcode')];
      sheets.items.appendRow(BUSINESS_CENTER_SALE_ITEM_HEADERS.map(function(h) {
        return h === 'Sale_ID' ? saleId : h === 'Product_ID' ? line.product[idCol] :
          h === 'Product_Name' ? name : h === 'Barcode' ? barcode :
          h === 'Quantity' ? line.qty : h === 'Unit_Price' ? line.price :
          h === 'Line_Total' ? line.lineTotal : h === 'Created_At' ? now : '';
      }));
      sheets.products.getRange(line.rowIndex + 1, stockCol + 1).setValue(Number(line.product[stockCol] || 0) - line.qty);
    });

    if (customerId && balanceDue > 0) {
      const customerHeaders = sheets.customers.getDataRange().getValues()[0];
      const customerBalanceCol = customerHeaders.indexOf('Balance');
      sheets.customers.getRange(customerRow + 1, customerBalanceCol + 1).setValue(customerBalance + balanceDue);
    }

    return {
      success: true, saleId: saleId, subtotal: subtotal, discount: discount,
      total: total, amountPaid: amountPaid, balanceDue: balanceDue,
      customerId: customerId, message: 'Sale completed successfully.'
    };
  } catch (error) {
    console.error('Business Center sale error:', error);
    return { success: false, message: error.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


// ==================== BUSINESS CENTER PART 3: SALES HISTORY + RECEIPTS ====================

function getBusinessCenterSales(payload, sessionId) {
  try {
    const access = businessCenterAccess_(sessionId, false);
    if (!access.ok) return { success: false, message: access.message };
    const workspace = getWorkspaceFile(sessionId);
    const sheets = getBusinessCenterPOSSheets_(workspace);
    const rows = businessCenterRows_(sheets.sales).reverse();
    const limit = Math.min(Math.max(Number((payload || {}).limit || 100), 1), 500);
    return { success: true, sales: rows.slice(0, limit) };
  } catch (error) {
    console.error('Business Center sales history error:', error);
    return { success: false, message: error.message };
  }
}

function getBusinessCenterSaleDetails(saleId, sessionId) {
  try {
    const access = businessCenterAccess_(sessionId, false);
    if (!access.ok) return { success: false, message: access.message };
    saleId = String(saleId || '').trim();
    if (!saleId) return { success: false, message: 'Sale ID is required.' };

    const workspace = getWorkspaceFile(sessionId);
    const sheets = getBusinessCenterPOSSheets_(workspace);
    const sales = businessCenterRows_(sheets.sales);
    const sale = sales.find(function(row) { return String(row.Sale_ID || '') === saleId; });
    if (!sale) return { success: false, message: 'Sale not found.' };

    const items = businessCenterRows_(sheets.items).filter(function(row) {
      return String(row.Sale_ID || '') === saleId;
    });

    return { success: true, sale: sale, items: items };
  } catch (error) {
    console.error('Business Center sale details error:', error);
    return { success: false, message: error.message };
  }
}


// ==================== BUSINESS CENTER PART 3: INVENTORY ====================

const BUSINESS_CENTER_STOCK_HEADERS = [
  'Movement_ID','Date','Product_ID','Product_Name','Type','Quantity',
  'Before_Qty','After_Qty','Reason','Created_By','Created_At'
];

function getBusinessCenterInventorySheets_(workspace) {
  const sheets = getBusinessCenterSheets_(workspace);
  let stock = workspace.getSheetByName('BusinessCenter_Stock_Movements');
  if (!stock) {
    stock = workspace.insertSheet('BusinessCenter_Stock_Movements');
    stock.getRange(1, 1, 1, BUSINESS_CENTER_STOCK_HEADERS.length).setValues([BUSINESS_CENTER_STOCK_HEADERS]);
  }
  return { products: sheets.products, customers: sheets.customers, stock: stock };
}

function updateBusinessCenterProduct(data, sessionId) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return { success: false, message: 'Another inventory update is in progress. Please try again.' };
    const access = businessCenterAccess_(sessionId, true);
    if (!access.ok) return { success: false, message: access.message };
    data = data || {};
    const productId = String(data.Product_ID || '').trim();
    if (!productId) return { success: false, message: 'Product ID is required.' };

    const workspace = getWorkspaceFile(sessionId);
    const sheets = getBusinessCenterInventorySheets_(workspace);
    const values = sheets.products.getDataRange().getValues();
    const headers = values[0];
    const idCol = headers.indexOf('Product_ID');
    if (idCol < 0) return { success: false, message: 'Product ID column is missing.' };

    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idCol] || '') === productId) { rowIndex = i; break; }
    }
    if (rowIndex < 0) return { success: false, message: 'Product not found.' };

    const allowed = ['Name','SKU','Barcode','Cost_Price','Selling_Price','Minimum_Stock','Unit','Category','Supplier','Location'];
    allowed.forEach(function(field) {
      if (Object.prototype.hasOwnProperty.call(data, field)) {
        const col = headers.indexOf(field);
        if (col >= 0) {
          let value = data[field];
          if (['Cost_Price','Selling_Price','Minimum_Stock'].includes(field)) {
            value = Number(value || 0);
            if (!Number.isFinite(value) || value < 0) throw new Error(field + ' must be a valid non-negative number.');
          }
          if (field === 'Name' && !String(value || '').trim()) throw new Error('Product name is required.');
          sheets.products.getRange(rowIndex + 1, col + 1).setValue(value);
        }
      }
    });
    return { success: true, productId: productId, message: 'Product details updated.' };
  } catch (error) {
    console.error('Business Center product update error:', error);
    return { success: false, message: error.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function adjustBusinessCenterStock(data, sessionId) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return { success: false, message: 'Another inventory update is in progress. Please try again.' };
    const access = businessCenterAccess_(sessionId, true);
    if (!access.ok) return { success: false, message: access.message };
    data = data || {};
    const productId = String(data.Product_ID || '').trim();
    const type = String(data.Type || '').trim();
    const quantity = Number(data.Quantity);
    const reason = String(data.Reason || '').trim();

    if (!productId) return { success: false, message: 'Product ID is required.' };
    if (!['stock_in','stock_out','set'].includes(type)) return { success: false, message: 'Choose a valid stock action.' };
    if (!Number.isFinite(quantity) || quantity < 0 || (type !== 'set' && quantity === 0)) {
      return { success: false, message: 'Enter a valid stock quantity.' };
    }
    if (!reason) return { success: false, message: 'Please enter a reason for the stock change.' };

    const workspace = getWorkspaceFile(sessionId);
    const sheets = getBusinessCenterInventorySheets_(workspace);
    const values = sheets.products.getDataRange().getValues();
    const headers = values[0];
    const idCol = headers.indexOf('Product_ID');
    const nameCol = headers.indexOf('Name');
    const stockCol = headers.indexOf('Stock_Qty');
    if (idCol < 0 || stockCol < 0) return { success: false, message: 'Inventory columns are missing.' };

    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idCol] || '') === productId) { rowIndex = i; break; }
    }
    if (rowIndex < 0) return { success: false, message: 'Product not found.' };

    const before = Number(values[rowIndex][stockCol] || 0);
    let after = before;
    if (type === 'stock_in') after = before + quantity;
    if (type === 'stock_out') after = before - quantity;
    if (type === 'set') after = quantity;
    if (after < 0) return { success: false, message: 'Stock cannot go below zero. Available: ' + before };

    sheets.products.getRange(rowIndex + 1, stockCol + 1).setValue(after);

    const now = new Date().toISOString();
    const user = access.user;
    const movementId = 'MOV-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    sheets.stock.appendRow(BUSINESS_CENTER_STOCK_HEADERS.map(function(h) {
      return h === 'Movement_ID' ? movementId :
        h === 'Date' ? now.split('T')[0] :
        h === 'Product_ID' ? productId :
        h === 'Product_Name' ? values[rowIndex][nameCol] :
        h === 'Type' ? type :
        h === 'Quantity' ? quantity :
        h === 'Before_Qty' ? before :
        h === 'After_Qty' ? after :
        h === 'Reason' ? reason :
        h === 'Created_By' ? (user.email || user.name || '') :
        h === 'Created_At' ? now : '';
    }));

    return { success: true, productId: productId, beforeQty: before, afterQty: after, movementId: movementId, message: 'Stock updated successfully.' };
  } catch (error) {
    console.error('Business Center stock adjustment error:', error);
    return { success: false, message: error.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function getBusinessCenterStockMovements(sessionId, productId) {
  try {
    const access = businessCenterAccess_(sessionId, false);
    if (!access.ok) return { success: false, message: access.message };
    const workspace = getWorkspaceFile(sessionId);
    const sheets = getBusinessCenterInventorySheets_(workspace);
    const rows = businessCenterRows_(sheets.stock);
    const id = String(productId || '').trim();
    const filtered = id ? rows.filter(function(row) { return String(row.Product_ID || '') === id; }) : rows;
    filtered.reverse();
    return { success: true, movements: filtered.slice(0, 100) };
  } catch (error) {
    return { success: false, message: error.message };
  }
}
