/**
 * Historical financial reconciliation for operational modules.
 *
 * SAFE MODE:
 * - auditHistoricalFinancialReconciliation(workspaceId) performs no writes to Financial_Data.
 * - applyHistoricalFinancialReconciliation(workspaceId) creates only source-linked events
 *   whose source transaction has a stable ID and is not already linked.
 * - Unlinked manual Finance rows are never matched or modified.
 */

function auditHistoricalFinancialReconciliation(workspaceId) {
  return reconcileHistoricalFinancialEvents_(workspaceId, false);
}

function applyHistoricalFinancialReconciliation(workspaceId) {
  return reconcileHistoricalFinancialEvents_(workspaceId, true);
}

function reconcileCurrentWorkspaceHistoricalFinancials(applyChanges) {
  return reconcileHistoricalFinancialEvents_(SpreadsheetApp.getActiveSpreadsheet().getId(), !!applyChanges);
}

function reconcileHistoricalFinancialEvents_(workspaceId, applyChanges) {
  var startedAt = new Date().toISOString();
  var result = {
    success: false,
    mode: applyChanges ? 'APPLY' : 'AUDIT',
    workspaceId: String(workspaceId || ''),
    startedAt: startedAt,
    summary: {
      scanned: 0,
      alreadyLinked: 0,
      added: 0,
      skipped: 0,
      staleSourceLinks: 0,
      missingStableId: 0,
      ambiguous: 0,
      errors: 0
    },
    modules: {},
    entries: [],
    message: ''
  };

  try {
    if (!workspaceId) throw new Error('workspaceId is required.');

    var workspace = SpreadsheetApp.openById(workspaceId);
    var financeSheet = workspace.getSheetByName('Financial_Data');
    if (!financeSheet) throw new Error('Financial_Data sheet not found.');

    var financeValues = financeSheet.getDataRange().getValues();
    if (!financeValues.length) throw new Error('Financial_Data has no header row.');

    var financeHeaders = financeValues[0].map(function(h) { return String(h || '').trim(); });
    var requiredFinanceHeaders = ['Transaction_ID', 'Date', 'Type', 'Amount'];
    var missingFinanceHeaders = requiredFinanceHeaders.filter(function(h) {
      return financeHeaders.indexOf(h) < 0;
    });
    if (missingFinanceHeaders.length) {
      throw new Error('Financial_Data is missing required headers: ' + missingFinanceHeaders.join(', '));
    }

    // Existing explicit source links are authoritative. We do not infer a link
    // from amount, date, description, customer name, or any other similarity.
    var existingLinks = {};
    var financeTransactionIds = {};
    var fSource = financeHeaders.indexOf('Module_Source');
    var fRecord = financeHeaders.indexOf('Module_Record_ID');
    var fType = financeHeaders.indexOf('Type');
    var fTxn = financeHeaders.indexOf('Transaction_ID');

    for (var fi = 1; fi < financeValues.length; fi++) {
      var frow = financeValues[fi];
      var txnId = String(frow[fTxn] || '').trim();
      if (txnId) financeTransactionIds[txnId] = true;

      if (fSource >= 0 && fRecord >= 0 && fType >= 0) {
        var source = String(frow[fSource] || '').trim();
        var recordId = String(frow[fRecord] || '').trim();
        var type = String(frow[fType] || '').trim().toLowerCase();
        if (source && recordId && type) {
          existingLinks[source + '|' + recordId + '|' + type] = txnId || true;
        }
      }
    }

    var targets = [
      { module: 'POS', sheet: 'POS_Data', id: 'Transaction_ID' },
      { module: 'Ecommerce', sheet: 'Ecommerce_Data', id: 'Order_ID' },
      { module: 'Sales', sheet: 'Sales_Data', id: 'Deal_ID' },
      { module: 'Agro', sheet: 'Agro_Data', id: 'Record_ID' },
      { module: 'BusinessCenter', sheet: 'BusinessCenter_Sales', id: 'Sale_ID' }
    ];

    targets.forEach(function(target) {
      var moduleResult = {
        scanned: 0,
        alreadyLinked: 0,
        added: 0,
        skipped: 0,
        staleSourceLinks: 0,
        missingStableId: 0,
        ambiguous: 0,
        errors: 0
      };
      result.modules[target.module] = moduleResult;

      var sheet = workspace.getSheetByName(target.sheet);
      if (!sheet || sheet.getLastRow() <= 1) return;

      var values = sheet.getDataRange().getValues();
      var headers = values[0].map(function(h) { return String(h || '').trim(); });
      var sourceIdCol = headers.indexOf(target.id);
      if (sourceIdCol < 0) {
        moduleResult.errors++;
        result.summary.errors++;
        result.entries.push({
          module: target.module,
          decision: 'ERROR',
          reason: 'Missing stable source ID header: ' + target.id
        });
        return;
      }

      var sourceFinancialLinkCol = headers.indexOf('Financial_Link_ID');

      values.slice(1).forEach(function(row, offset) {
        moduleResult.scanned++;
        result.summary.scanned++;

        var sourceRow = offset + 2;
        var recordId = String(row[sourceIdCol] || '').trim();

        if (!recordId) {
          moduleResult.missingStableId++;
          result.summary.missingStableId++;
          moduleResult.skipped++;
          result.summary.skipped++;
          result.entries.push({
            module: target.module,
            sourceRow: sourceRow,
            decision: 'SKIP',
            reason: 'No stable source transaction ID'
          });
          return;
        }

        var desired = historicalFinancialEntries_(target.module, headers, row, recordId);
        if (!desired.length) {
          moduleResult.skipped++;
          result.summary.skipped++;
          result.entries.push({
            module: target.module,
            sourceRow: sourceRow,
            sourceRecordId: recordId,
            decision: 'SKIP',
            reason: 'Source record has no qualifying financial event'
          });
          return;
        }

        desired.forEach(function(entry) {
          var linkKey = target.module + '|' + recordId + '|' + String(entry.Type).toLowerCase();

          if (existingLinks[linkKey]) {
            moduleResult.alreadyLinked++;
            result.summary.alreadyLinked++;
            result.entries.push({
              module: target.module,
              sourceRow: sourceRow,
              sourceRecordId: recordId,
              financialType: entry.Type,
              amount: entry.Amount,
              decision: 'SKIP',
              reason: 'Already linked in Financial_Data',
              existingTransactionId: existingLinks[linkKey] === true ? '' : existingLinks[linkKey]
            });
            return;
          }

          // If the source explicitly claims a Financial_Link_ID, but that
          // transaction is not present in Financial_Data, do not guess whether
          // another Finance row is the intended transaction.
          var explicitLink = sourceFinancialLinkCol >= 0
            ? String(row[sourceFinancialLinkCol] || '').trim()
            : '';
          if (explicitLink) {
            if (financeTransactionIds[explicitLink]) {
              moduleResult.alreadyLinked++;
              result.summary.alreadyLinked++;
              result.entries.push({
                module: target.module,
                sourceRow: sourceRow,
                sourceRecordId: recordId,
                financialType: entry.Type,
                amount: entry.Amount,
                decision: 'SKIP',
                reason: 'Source Financial_Link_ID already points to an existing Finance transaction',
                existingTransactionId: explicitLink
              });
            } else {
              moduleResult.staleSourceLinks++;
              result.summary.staleSourceLinks++;
              moduleResult.skipped++;
              result.summary.skipped++;
              result.entries.push({
                module: target.module,
                sourceRow: sourceRow,
                sourceRecordId: recordId,
                financialType: entry.Type,
                amount: entry.Amount,
                decision: 'SKIP',
                reason: 'Source Financial_Link_ID exists but target Finance transaction is missing; no inference performed',
                existingTransactionId: explicitLink
              });
            }
            return;
          }

          var entryResult = {
            module: target.module,
            sourceRow: sourceRow,
            sourceRecordId: recordId,
            financialType: entry.Type,
            amount: entry.Amount,
            decision: applyChanges ? 'ADD' : 'WOULD_ADD',
            reason: 'Stable source transaction ID exists and no explicit Finance link was found'
          };

          if (applyChanges) {
            try {
              var transactionId = generateTransactionId();
              entry.Transaction_ID = transactionId;

              var rowData = financeHeaders.map(function(header) {
                return entry[header] !== undefined ? entry[header] : '';
              });
              financeSheet.appendRow(rowData);

              existingLinks[linkKey] = transactionId;
              financeTransactionIds[transactionId] = true;

              moduleResult.added++;
              result.summary.added++;
              entryResult.newTransactionId = transactionId;
            } catch (appendError) {
              moduleResult.errors++;
              result.summary.errors++;
              entryResult.decision = 'ERROR';
              entryResult.reason = appendError.message || String(appendError);
            }
          } else {
            moduleResult.skipped++;
            result.summary.skipped++;
          }

          result.entries.push(entryResult);
        });
      });
    });

    result.success = result.summary.errors === 0;
    result.completedAt = new Date().toISOString();
    result.message = applyChanges
      ? 'Historical financial reconciliation completed. Only explicit source-linked events were added.'
      : 'Historical financial reconciliation audit completed. No Financial_Data records were changed.';

    if (applyChanges && result.summary.added > 0) {
      try {
        recalculateDashboardMetrics(workspaceId);
      } catch (metricError) {
        result.metricRecalculationWarning = metricError.message || String(metricError);
      }
    }

    return result;
  } catch (error) {
    result.success = false;
    result.completedAt = new Date().toISOString();
    result.message = error.message || String(error);
    result.summary.errors++;
    return result;
  }
}

function historicalFinancialEntries_(moduleName, headers, row, recordId) {
  var value = function(name) {
    var i = headers.indexOf(name);
    return i >= 0 ? row[i] : '';
  };
  var number = function(name) {
    var n = parseFloat(value(name));
    return isNaN(n) ? 0 : n;
  };
  var date = value('Date') || new Date().toISOString().split('T')[0];
  var createdBy = value('Created_By') || '';
  var createdAt = value('Created_At') || new Date().toISOString();
  var entries = [];

  if (moduleName === 'POS') {
    var posAmount = number('Total_Amount');
    if (posAmount > 0) {
      entries.push({
        Date: date,
        Type: 'revenue',
        Category: 'pos',
        Subcategory: 'point_of_sale',
        Module_Source: 'POS',
        Module_Record_ID: recordId,
        Description: 'POS sale ' + recordId + ': ' + String(value('Product_Name') || ''),
        Amount: posAmount,
        Payment_Method: value('Payment_Method') || 'completed',
        Reference: recordId,
        Status: value('Status') || 'completed',
        Created_By: createdBy,
        Created_At: createdAt
      });
    }
  } else if (moduleName === 'Ecommerce') {
    var ecommerceAmount = number('Total_Amount');
    if (ecommerceAmount > 0) {
      entries.push({
        Date: date,
        Type: 'revenue',
        Category: 'ecommerce',
        Subcategory: 'product_sales',
        Module_Source: 'Ecommerce',
        Module_Record_ID: recordId,
        Description: 'Order ' + recordId + ': ' + String(value('Product_Name') || ''),
        Amount: ecommerceAmount,
        Payment_Method: String(value('Payment_Status') || '').toLowerCase() === 'paid' ? 'completed' : 'pending',
        Reference: recordId,
        Status: value('Payment_Status') || 'pending',
        Created_By: createdBy,
        Created_At: createdAt
      });

      var shipping = number('Shipping_Cost');
      if (shipping > 0) {
        entries.push({
          Date: date,
          Type: 'expense',
          Category: 'shipping',
          Subcategory: 'fulfillment',
          Module_Source: 'Ecommerce',
          Module_Record_ID: recordId,
          Description: 'Shipping for order ' + recordId,
          Amount: -Math.abs(shipping),
          Payment_Method: 'completed',
          Reference: recordId,
          Status: 'completed',
          Created_By: createdBy,
          Created_At: createdAt
        });
      }
    }
  } else if (moduleName === 'Sales') {
    if (String(value('Stage') || '').toLowerCase() === 'closed_won' && number('Total_Amount') > 0) {
      entries.push({
        Date: date,
        Type: 'revenue',
        Category: 'sales',
        Subcategory: value('Product_Service') || 'product',
        Module_Source: 'Sales',
        Module_Record_ID: recordId,
        Description: 'Sale: ' + String(value('Product_Service') || '') + ' to ' + String(value('Customer_Name') || ''),
        Amount: number('Total_Amount'),
        Payment_Method: value('Payment_Method') || 'pending',
        Reference: recordId,
        Status: 'completed',
        Created_By: createdBy,
        Created_At: createdAt
      });
    }
  } else if (moduleName === 'Agro') {
    if (String(value('Activity') || '').toLowerCase() === 'sale' && number('Revenue') > 0) {
      entries.push({
        Date: date,
        Type: 'revenue',
        Category: 'agro',
        Subcategory: String(value('Type') || '').toLowerCase() === 'crop' ? 'crop_sales' : 'livestock_sales',
        Module_Source: 'Agro',
        Module_Record_ID: recordId,
        Description: 'Sold ' + String(value('Quantity') || '') + ' ' + String(value('Crop_Name') || value('Livestock_Type') || ''),
        Amount: number('Revenue'),
        Payment_Method: 'completed',
        Reference: recordId,
        Status: 'completed',
        Created_By: createdBy,
        Created_At: createdAt
      });
    } else if (number('Total_Cost') > 0) {
      entries.push({
        Date: date,
        Type: 'expense',
        Category: 'agro',
        Subcategory: String(value('Type') || '').toLowerCase() === 'crop' ? 'farming_cost' : 'animal_husbandry',
        Module_Source: 'Agro',
        Module_Record_ID: recordId,
        Description: String(value('Activity') || '') + ': ' + String(value('Quantity') || '') + ' ' + String(value('Crop_Name') || value('Livestock_Type') || ''),
        Amount: -Math.abs(number('Total_Cost')),
        Payment_Method: 'completed',
        Reference: recordId,
        Status: 'completed',
        Created_By: createdBy,
        Created_At: createdAt
      });
    }
  } else if (moduleName === 'BusinessCenter') {
    var bcTotal = number('Total');
    if (bcTotal > 0) {
      entries.push({
        Date: date,
        Type: 'revenue',
        Category: 'business_center',
        Subcategory: 'point_of_sale',
        Module_Source: 'BusinessCenter',
        Module_Record_ID: recordId,
        Description: 'Business Center sale ' + recordId,
        Amount: bcTotal,
        Payment_Method: value('Payment_Method') || 'pending',
        Reference: recordId,
        Status: value('Status') || 'completed',
        Created_By: createdBy,
        Created_At: createdAt
      });
    }
  }

  return entries;
}
