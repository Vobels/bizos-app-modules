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

  const role = String(user.role || 'staff').toLowerCase();
  const accessibleModules = Array.isArray(user.accessibleModules) ? user.accessibleModules : [];
  const hasEcommerceAccess = accessibleModules.indexOf('Ecommerce') !== -1;

  // Business Center is tied to the Ecommerce module. The server-side
  // module entitlement remains authoritative for both viewing and writing.
  if (!hasEcommerceAccess) {
    return { ok: false, message: 'Business Center is not available for this account.' };
  }

  if (write) {
    // Demo accounts remain read-only regardless of role/module entitlement.
    if (user.isDemo === 'YES') {
      return { ok: false, message: 'Demo accounts are view-only.' };
    }

    // Owners/admins retain full Business Center write access.
    // Staff can write when Ecommerce/Business Center is assigned to them.
    if (role !== 'owner' && role !== 'admin' && role !== 'staff') {
      return { ok: false, message: 'You have view-only access to Business Center.' };
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

function getBusinessCenterProductImages(productIds, sessionId) {
  try {
    const access = businessCenterAccess_(sessionId, false);
    if (!access.ok) return { success: false, message: access.message };
    const ids = Array.isArray(productIds) ? productIds.slice(0, 30).map(function(id) { return String(id || '').trim(); }).filter(Boolean) : [];
    if (!ids.length) return { success: true, images: {} };

    const workspace = getWorkspaceFile(sessionId);
    const sheets = getBusinessCenterSheets_(workspace);
    const products = businessCenterRows_(sheets.products);
    const wanted = {};
    ids.forEach(function(id) { wanted[id] = true; });

    const images = {};
    products.forEach(function(product) {
      if (!wanted[String(product.Product_ID || '')]) return;
      const data = businessCenterImageData_(product.Image_File_ID);
      if (data) images[String(product.Product_ID)] = data;
    });
    return { success: true, images: images };
  } catch (error) {
    console.error('Business Center product image error:', error);
    return { success: false, message: error.message };
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

    const customerBalances = getBusinessCenterCustomerBalanceMap_(workspace);
    customers.forEach(function(customer) {
      const id = String(customer.Customer_ID || '').trim();
      customer.Balance = Number(customerBalances[id] || 0);
    });
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
    const existingProducts = businessCenterRows_(sheets.products);
    if (['free', 'starter'].indexOf(String(access.user.subscriptionTier || '').toLowerCase()) !== -1 && existingProducts.length >= 10) {
      return { success: false, message: 'Free Business Center accounts can contain up to 10 products. Upgrade to Sovereign for unlimited products.' };
    }
    const productId = 'PRD-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const barcode = String(data.Barcode || '').trim() || generateBusinessCenterBarcode_(existingProducts);
    if (existingProducts.some(function(p) { return String(p.Barcode || '').trim() === barcode; })) {
      return { success: false, message: 'That barcode is already used by another product.' };
    }
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
      if (header === 'Barcode') return barcode;
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


function generateBusinessCenterBarcode_(existing) {
  const used = {};
  (existing || []).forEach(function(p) {
    const b = String(p.Barcode || '').trim();
    if (b) used[b] = true;
  });
  let code = '';
  do {
    code = '20' + String(Date.now()).slice(-8) + String(Math.floor(Math.random() * 100)).padStart(2, '0');
  } while (used[code]);
  return code;
}

function saveBusinessCenterProductsBulk(data, sessionId) {
  try {
    const access = businessCenterAccess_(sessionId, true);
    if (!access.ok) return { success: false, message: access.message };
    const rows = Array.isArray(data) ? data : [];
    if (!rows.length) return { success: false, message: 'No products were supplied.' };

    const workspace = getWorkspaceFile(sessionId);
    const sheets = getBusinessCenterSheets_(workspace);
    const existing = businessCenterRows_(sheets.products);
    const isFree = String(access.user.subscriptionTier || '').toLowerCase() === 'free';
    if (isFree && existing.length >= 10) return { success: false, message: 'Free Business Center accounts can contain up to 10 products. Upgrade to Sovereign for unlimited products.' };
    const importRows = isFree ? rows.slice(0, Math.max(0, 10 - existing.length)) : rows.slice(0, 1000);
    const existingBarcodes = {};
    existing.forEach(function(p) {
      const b = String(p.Barcode || '').trim();
      if (b) existingBarcodes[b] = true;
    });

    const user = access.user;
    const now = new Date().toISOString();
    const values = [];
    const errors = [];
    const seen = {};
    importRows.forEach(function(item, index) {
      item = item || {};
      const name = String(item.Name || '').trim();
      const selling = Number(item.Selling_Price || 0);
      const stock = Number(item.Stock_Qty || 0);
      if (!name) { errors.push('Row ' + (index + 2) + ': product name is required.'); return; }
      if (!Number.isFinite(selling) || selling < 0) { errors.push('Row ' + (index + 2) + ': selling price is invalid.'); return; }
      if (!Number.isFinite(stock) || stock < 0) { errors.push('Row ' + (index + 2) + ': stock is invalid.'); return; }

      let barcode = String(item.Barcode || '').trim();
      if (barcode) {
        if (existingBarcodes[barcode] || seen[barcode]) {
          errors.push('Row ' + (index + 2) + ': barcode ' + barcode + ' already exists.');
          return;
        }
      } else {
        barcode = generateBusinessCenterBarcode_(existing.concat(values.map(function(v) {
          return { Barcode: v[3] };
        })));
      }
      seen[barcode] = true;

      const productId = 'PRD-' + Utilities.getUuid().substring(0, 8).toUpperCase();
      values.push(BUSINESS_CENTER_PRODUCT_HEADERS.map(function(header) {
        if (header === 'Product_ID') return productId;
        if (header === 'Barcode') return barcode;
        if (header === 'Created_By') return user.email || user.name || '';
        if (header === 'Created_At') return now;
        if (header === 'Stock_Qty') return stock;
        if (header === 'Minimum_Stock') return Math.max(0, Number(item.Minimum_Stock || 0));
        if (header === 'Cost_Price') return Math.max(0, Number(item.Cost_Price || 0));
        if (header === 'Selling_Price') return selling;
        return item[header] === undefined || item[header] === null ? '' : item[header];
      }));
    });

    if (values.length) {
      sheets.products.getRange(sheets.products.getLastRow() + 1, 1, values.length, BUSINESS_CENTER_PRODUCT_HEADERS.length).setValues(values);
    }
    return {
      success: true,
      imported: values.length,
      skipped: errors.length,
      errors: errors.slice(0, 20),
      message: values.length + ' product' + (values.length === 1 ? '' : 's') + ' added.'
    };
  } catch (error) {
    console.error('Business Center bulk product import error:', error);
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
    let customerRow = -1;
    let resolvedCustomerName = '';
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
          customerBalance = getBusinessCenterCustomerBalance_(workspace, customerId);
          resolvedCustomerName = String(customerValues[i][customerNameCol] || '').trim();
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

    // Keep the original stock/balance values so a later write failure can be rolled back.
    const originalStocks = resolved.map(function(line) {
      return {
        rowIndex: line.rowIndex,
        stock: Number(line.product[stockCol] || 0)
      };
    });
    const customerHeaders = sheets.customers.getDataRange().getValues()[0];
    const customerBalanceCol = customerHeaders.indexOf('Balance');

    try {
      sheets.sales.appendRow(BUSINESS_CENTER_SALE_HEADERS.map(function(h) {
        return h === 'Sale_ID' ? saleId :
          h === 'Date' ? (data.Date || now.split('T')[0]) :
          h === 'Customer_ID' ? customerId :
          h === 'Customer_Name' ? (resolvedCustomerName || String(data.Customer_Name || '').trim()) :
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

        const beforeQty = Number(line.product[stockCol] || 0);
        const afterQty = beforeQty - line.qty;
        sheets.products.getRange(line.rowIndex + 1, stockCol + 1).setValue(afterQty);

        // Every completed sale creates an inventory audit entry so stock
        // reductions from POS are visible alongside manual stock adjustments.
        const stockMovementSheet = getBusinessCenterInventorySheets_(workspace).stock;
        stockMovementSheet.appendRow([
          'MOV-' + Utilities.getUuid().substring(0, 8).toUpperCase(),
          data.Date || now.split('T')[0],
          line.product[idCol],
          name,
          'sale',
          line.qty,
          beforeQty,
          afterQty,
          'POS sale ' + saleId,
          user.email || user.name || '',
          now
        ]);
      });

      if (customerId && balanceDue > 0) {
        sheets.customers.getRange(customerRow + 1, customerBalanceCol + 1).setValue(customerBalance + balanceDue);
      }
    } catch (writeError) {
      // Google Sheets writes are not transactional. Remove any records created
      // by this sale and restore stock/customer balance before reporting failure.
      try {
        const deleteRowsByValue = function(sheet, headerName, matchValue) {
          const values = sheet.getDataRange().getValues();
          if (!values.length) return;
          const col = values[0].indexOf(headerName);
          if (col < 0) return;
          for (let i = values.length - 1; i >= 1; i--) {
            if (String(values[i][col] || '') === String(matchValue)) sheet.deleteRow(i + 1);
          }
        };

        deleteRowsByValue(sheets.sales, 'Sale_ID', saleId);
        deleteRowsByValue(sheets.items, 'Sale_ID', saleId);

        const stockMovementSheet = workspace.getSheetByName('BusinessCenter_Stock_Movements');
        if (stockMovementSheet) {
          const values = stockMovementSheet.getDataRange().getValues();
          const reasonCol = values.length ? values[0].indexOf('Reason') : -1;
          if (reasonCol >= 0) {
            for (let i = values.length - 1; i >= 1; i--) {
              if (String(values[i][reasonCol] || '') === 'POS sale ' + saleId) {
                stockMovementSheet.deleteRow(i + 1);
              }
            }
          }
        }

        originalStocks.forEach(function(original) {
          sheets.products.getRange(original.rowIndex + 1, stockCol + 1).setValue(original.stock);
        });

        if (customerId && balanceDue > 0) {
          sheets.customers.getRange(customerRow + 1, customerBalanceCol + 1).setValue(customerBalance);
        }
      } catch (rollbackError) {
        console.error('Business Center sale rollback error:', rollbackError);
        throw new Error('Sale could not be completed safely, and automatic rollback also failed. Please check the sale and inventory records before retrying.');
      }
      throw writeError;
    }

    // Post the completed Business Center sale to the same financial ledger
    // used by the main dashboard. The source Sale_ID makes retries safe.
    try {
      const financialResult = autoSyncToFinancial('BusinessCenter', {
        Sale_ID: saleId,
        Date: data.Date || now.split('T')[0],
        Customer_Name: resolvedCustomerName || String(data.Customer_Name || '').trim(),
        Total: total,
        Amount_Paid: amountPaid,
        Payment_Method: data.Payment_Method || 'Cash',
        Status: balanceDue > 0 ? 'Credit' : 'Paid',
        Created_By: user.email || user.name || ''
      }, workspace.getId(), sessionId);
      if (!financialResult.success) {
        console.warn('Business Center financial sync warning:', financialResult.message);
      }
    } catch (financialError) {
      console.warn('Business Center financial sync error:', financialError.message);
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



// ==================== BUSINESS CENTER PART 4: CUSTOMER LEDGER ====================
function getBusinessCenterCustomerBalanceMap_(workspace) {
  const sheets = getBusinessCenterPOSSheets_(workspace);
  const ledgerSheet = getBusinessCenterLedgerSheet_(workspace);
  const balances = {};

  businessCenterRows_(sheets.sales).forEach(function(sale) {
    const id = String(sale.Customer_ID || '').trim();
    if (!id) return;
    const due = Number(sale.Balance_Due || 0);
    if (due > 0) balances[id] = (balances[id] || 0) + due;
  });

  businessCenterRows_(ledgerSheet).forEach(function(row) {
    const id = String(row.Customer_ID || '').trim();
    if (!id) return;
    balances[id] = (balances[id] || 0) - Number(row.Credit || 0);
  });

  Object.keys(balances).forEach(function(id) {
    balances[id] = Math.max(0, balances[id]);
  });
  return balances;
}

function getBusinessCenterCustomerBalance_(workspace, customerId) {
  const balances = getBusinessCenterCustomerBalanceMap_(workspace);
  return Number(balances[String(customerId || '').trim()] || 0);
}


const BUSINESS_CENTER_LEDGER_HEADERS = [
  'Entry_ID','Date','Customer_ID','Customer_Name','Type','Reference_ID',
  'Description','Debit','Credit','Balance','Created_By','Created_At'
];

function getBusinessCenterLedgerSheet_(workspace) {
  let sheet = workspace.getSheetByName('BusinessCenter_Customer_Ledger');
  if (!sheet) {
    sheet = workspace.insertSheet('BusinessCenter_Customer_Ledger');
    sheet.getRange(1, 1, 1, BUSINESS_CENTER_LEDGER_HEADERS.length).setValues([BUSINESS_CENTER_LEDGER_HEADERS]);
  }
  return sheet;
}

function getBusinessCenterCustomerLedger(customerId, sessionId) {
  try {
    const access = businessCenterAccess_(sessionId, false);
    if (!access.ok) return { success: false, message: access.message };
    customerId = String(customerId || '').trim();
    if (!customerId) return { success: false, message: 'Customer is required.' };

    const workspace = getWorkspaceFile(sessionId);
    const sheets = getBusinessCenterPOSSheets_(workspace);
    const ledgerSheet = getBusinessCenterLedgerSheet_(workspace);
    const customers = businessCenterRows_(sheets.customers);
    const customer = customers.find(function(row) {
      return String(row.Customer_ID || '') === customerId;
    });
    if (!customer) return { success: false, message: 'Customer not found.' };
    customer.Balance = getBusinessCenterCustomerBalance_(workspace, customerId);

    const entries = [];

    businessCenterRows_(sheets.sales).forEach(function(sale) {
      if (String(sale.Customer_ID || '') !== customerId) return;
      const balanceDue = Number(sale.Balance_Due || 0);
      if (balanceDue <= 0) return;
      entries.push({
        Entry_ID: 'SALE-' + String(sale.Sale_ID || ''),
        Date: sale.Date || sale.Created_At || '',
        Customer_ID: customerId,
        Customer_Name: sale.Customer_Name || customer.Name || '',
        Type: 'Sale',
        Reference_ID: sale.Sale_ID || '',
        Description: 'Credit sale',
        Debit: balanceDue,
        Credit: 0,
        Created_By: sale.Created_By || '',
        Created_At: sale.Created_At || ''
      });
    });

    businessCenterRows_(ledgerSheet).forEach(function(row) {
      if (String(row.Customer_ID || '') !== customerId) return;
      entries.push({
        Entry_ID: row.Entry_ID || '',
        Date: row.Date || row.Created_At || '',
        Customer_ID: customerId,
        Customer_Name: row.Customer_Name || customer.Name || '',
        Type: row.Type || 'Payment',
        Reference_ID: row.Reference_ID || '',
        Description: row.Description || '',
        Debit: Number(row.Debit || 0),
        Credit: Number(row.Credit || 0),
        Created_By: row.Created_By || '',
        Created_At: row.Created_At || ''
      });
    });

    entries.sort(function(a, b) {
      return new Date(a.Created_At || a.Date || 0).getTime() - new Date(b.Created_At || b.Date || 0).getTime();
    });

    let running = 0;
    entries.forEach(function(entry) {
      running += Number(entry.Debit || 0) - Number(entry.Credit || 0);
      entry.Balance = running;
    });

    return {
      success: true,
      customer: customer,
      entries: entries.reverse(),
      summary: {
        totalDebit: entries.reduce(function(t, e) { return t + Number(e.Debit || 0); }, 0),
        totalCredit: entries.reduce(function(t, e) { return t + Number(e.Credit || 0); }, 0),
        balance: getBusinessCenterCustomerBalance_(workspace, customerId)
      }
    };
  } catch (error) {
    console.error('Business Center customer ledger error:', error);
    return { success: false, message: error.message };
  }
}

function recordBusinessCenterCustomerPayment(data, sessionId) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return { success: false, message: 'Another customer payment is being processed. Please try again.' };

    const access = businessCenterAccess_(sessionId, true);
    if (!access.ok) return { success: false, message: access.message };
    data = data || {};
    const customerId = String(data.Customer_ID || '').trim();
    const amount = Number(data.Amount || 0);
    if (!customerId) return { success: false, message: 'Customer is required.' };
    if (!Number.isFinite(amount) || amount <= 0) return { success: false, message: 'Enter a valid payment amount.' };

    const workspace = getWorkspaceFile(sessionId);
    const sheets = getBusinessCenterPOSSheets_(workspace);
    const customers = sheets.customers;
    const values = customers.getDataRange().getValues();
    const headers = values[0] || [];
    const idCol = headers.indexOf('Customer_ID');
    const balanceCol = headers.indexOf('Balance');
    const nameCol = headers.indexOf('Name');
    if (idCol < 0 || balanceCol < 0) throw new Error('Customer schema is missing required fields.');

    let rowIndex = -1;
    let customerName = '';
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idCol] || '') === customerId) {
        rowIndex = i;
        customerName = String(values[i][nameCol] || '');
        break;
      }
    }
    if (rowIndex < 1) return { success: false, message: 'Customer not found.' };

    // Sales and ledger entries are the source of truth. The customer Balance
    // column is a cached value kept in sync with that derived balance.
    const currentBalance = getBusinessCenterCustomerBalance_(workspace, customerId);
    if (currentBalance <= 0) return { success: false, message: 'This customer has no outstanding balance.' };
    if (amount > currentBalance) return { success: false, message: 'Payment cannot be greater than the outstanding balance.' };

    const now = new Date().toISOString();
    const entryId = 'PAY-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const ledger = getBusinessCenterLedgerSheet_(workspace);

    try {
      ledger.appendRow([
        entryId,
        data.Date || now.split('T')[0],
        customerId,
        customerName,
        'Payment',
        entryId,
        String(data.Description || 'Customer payment'),
        0,
        amount,
        currentBalance - amount,
        access.user.email || access.user.name || '',
        now
      ]);

      customers.getRange(rowIndex + 1, balanceCol + 1).setValue(currentBalance - amount);
    } catch (writeError) {
      try {
        const valuesAfter = ledger.getDataRange().getValues();
        const entryCol = valuesAfter.length ? valuesAfter[0].indexOf('Entry_ID') : -1;
        if (entryCol >= 0) {
          for (let i = valuesAfter.length - 1; i >= 1; i--) {
            if (String(valuesAfter[i][entryCol] || '') === entryId) {
              ledger.deleteRow(i + 1);
              break;
            }
          }
        }
        customers.getRange(rowIndex + 1, balanceCol + 1).setValue(currentBalance);
      } catch (rollbackError) {
        console.error('Business Center customer payment rollback error:', rollbackError);
        throw new Error('Payment could not be recorded safely. Please check the customer ledger before retrying.');
      }
      throw writeError;
    }

    return {
      success: true,
      entryId: entryId,
      customerId: customerId,
      amount: amount,
      balance: currentBalance - amount,
      message: 'Customer payment recorded successfully.'
    };
  } catch (error) {
    console.error('Business Center customer payment error:', error);
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
    const limit = Math.min(Math.max(Number((payload || {}).limit || 100), 1), 500);
    const sheet = sheets.sales;
    if (!sheet || sheet.getLastRow() < 2) return { success: true, sales: [] };
    const lastRow = sheet.getLastRow();
    const headerCount = sheet.getLastColumn();
    const startRow = Math.max(2, lastRow - limit + 1);
    const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, headerCount).getValues();
    const headers = sheet.getRange(1, 1, 1, headerCount).getValues()[0];
    const rows = values.reverse().map(function(row) {
      const item = {};
      headers.forEach(function(header, i) { item[header] = row[i]; });
      return item;
    });
    return { success: true, sales: rows };
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

    // Barcodes must remain unique across the catalog, including when editing.
    // An empty barcode is allowed because BizOS can use the product's internal ID.
    if (Object.prototype.hasOwnProperty.call(data, 'Barcode')) {
      const barcode = String(data.Barcode || '').trim();
      if (barcode) {
        const barcodeCol = headers.indexOf('Barcode');
        if (barcodeCol >= 0) {
          for (let i = 1; i < values.length; i++) {
            if (i === rowIndex) continue;
            if (String(values[i][barcodeCol] || '').trim().toLowerCase() === barcode.toLowerCase()) {
              return { success: false, message: 'This barcode is already assigned to another product.' };
            }
          }
        }
      }
    }

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

const BUSINESS_CENTER_FINANCE_HEADERS=['Transaction_ID','Date','Type','Category','Reference_ID','Description','Customer_ID','Amount','Payment_Method','Created_By','Created_At'];
function getBusinessCenterFinanceSheet_(workspace){let s=workspace.getSheetByName('BusinessCenter_Finance_Transactions');if(!s){s=workspace.insertSheet('BusinessCenter_Finance_Transactions');s.getRange(1,1,1,BUSINESS_CENTER_FINANCE_HEADERS.length).setValues([BUSINESS_CENTER_FINANCE_HEADERS]);}return s;}
function getBusinessCenterFinanceData(sessionId){try{const access=businessCenterAccess_(sessionId,false);if(!access.ok)return{success:false,message:access.message};const workspace=getWorkspaceFile(sessionId),sheets=getBusinessCenterPOSSheets_(workspace),finance=getBusinessCenterFinanceSheet_(workspace),tx=[];
businessCenterRows_(sheets.sales).forEach(function(sale){const total=Number(sale.Total||0),paid=Number(sale.Amount_Paid||0);if(total>0)tx.push({Transaction_ID:'SALE-'+String(sale.Sale_ID||''),Date:sale.Date||sale.Created_At||'',Type:'Income',Category:'Sales Revenue',Reference_ID:sale.Sale_ID||'',Description:'Business Center sale',Customer_ID:sale.Customer_ID||'',Amount:total,Payment_Method:sale.Payment_Method||'',Created_By:sale.Created_By||'',Created_At:sale.Created_At||''});if(paid>0)tx.push({Transaction_ID:'CASH-'+String(sale.Sale_ID||''),Date:sale.Date||sale.Created_At||'',Type:'Cash In',Category:'Sales Collection',Reference_ID:sale.Sale_ID||'',Description:'Payment received for sale',Customer_ID:sale.Customer_ID||'',Amount:paid,Payment_Method:sale.Payment_Method||'',Created_By:sale.Created_By||'',Created_At:sale.Created_At||''});});
businessCenterRows_(getBusinessCenterLedgerSheet_(workspace)).forEach(function(row){const amount=Number(row.Credit||0);if(amount>0)tx.push({Transaction_ID:'PAY-'+String(row.Entry_ID||''),Date:row.Date||row.Created_At||'',Type:'Cash In',Category:'Customer Payment',Reference_ID:row.Reference_ID||row.Entry_ID||'',Description:row.Description||'Customer payment',Customer_ID:row.Customer_ID||'',Amount:amount,Payment_Method:'Customer Ledger',Created_By:row.Created_By||'',Created_At:row.Created_At||''});});
businessCenterRows_(finance).forEach(function(row){tx.push({Transaction_ID:row.Transaction_ID||'',Date:row.Date||row.Created_At||'',Type:row.Type||'Expense',Category:row.Category||'Expense',Reference_ID:row.Reference_ID||'',Description:row.Description||'',Customer_ID:row.Customer_ID||'',Amount:Number(row.Amount||0),Payment_Method:row.Payment_Method||'',Created_By:row.Created_By||'',Created_At:row.Created_At||''});});
tx.sort(function(a,b){return new Date(b.Created_At||b.Date||0).getTime()-new Date(a.Created_At||a.Date||0).getTime();});const revenue=tx.filter(function(t){return t.Type==='Income';}).reduce(function(s,t){return s+Number(t.Amount||0);},0),cashIn=tx.filter(function(t){return t.Type==='Cash In';}).reduce(function(s,t){return s+Number(t.Amount||0);},0),expenses=tx.filter(function(t){return t.Type==='Expense';}).reduce(function(s,t){return s+Number(t.Amount||0);},0),customerBalances=getBusinessCenterCustomerBalanceMap_(workspace),receivables=Object.keys(customerBalances).reduce(function(s,id){return s+Number(customerBalances[id]||0);},0);return{success:true,transactions:tx.slice(0,500),summary:{revenue:revenue,cashIn:cashIn,expenses:expenses,receivables:receivables,netCashFlow:cashIn-expenses}};}catch(error){console.error('Business Center finance load error:',error);return{success:false,message:error.message};}}
function recordBusinessCenterExpense(data,sessionId){try{const access=businessCenterAccess_(sessionId,true);if(!access.ok)return{success:false,message:access.message};data=data||{};const amount=Number(data.Amount||0),category=String(data.Category||'').trim(),description=String(data.Description||'').trim(),paymentMethod=String(data.Payment_Method||'Cash').trim();if(!Number.isFinite(amount)||amount<=0)return{success:false,message:'Enter a valid expense amount.'};if(!category)return{success:false,message:'Expense category is required.'};if(!description)return{success:false,message:'Expense description is required.'};const sheet=getBusinessCenterFinanceSheet_(getWorkspaceFile(sessionId)),now=new Date().toISOString(),id='EXP-'+Utilities.getUuid().substring(0,8).toUpperCase(),user=access.user;sheet.appendRow(BUSINESS_CENTER_FINANCE_HEADERS.map(function(h){if(h==='Transaction_ID')return id;if(h==='Date')return data.Date||now.split('T')[0];if(h==='Type')return'Expense';if(h==='Category')return category;if(h==='Reference_ID')return id;if(h==='Description')return description;if(h==='Amount')return amount;if(h==='Payment_Method')return paymentMethod;if(h==='Created_By')return user.email||user.name||'';if(h==='Created_At')return now;return'';}));

    // Keep Business Center expenses visible to the main business dashboard.
    // The Business Center transaction ID becomes the stable source identity,
    // so retries/reconciliation cannot create a second main-ledger expense.
    try {
      const financialResult = autoSyncToFinancial('BusinessCenterExpense', {
        Transaction_ID: id,
        Date: data.Date || now.split('T')[0],
        Total: amount,
        Category: category,
        Description: description,
        Payment_Method: paymentMethod,
        Created_By: user.email || user.name || ''
      }, getWorkspaceFile(sessionId).getId(), sessionId);
      if (!financialResult.success) {
        console.warn('Business Center expense financial sync warning:', financialResult.message);
      }
    } catch (financialError) {
      console.warn('Business Center expense financial sync error:', financialError.message);
    }

    return{success:true,transactionId:id,message:'Expense recorded successfully.'};}catch(error){console.error('Business Center expense error:',error);return{success:false,message:error.message};}}
