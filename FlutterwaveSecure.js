// ============================================================
// FlutterwaveSecure.js
// Flutterwave Standard checkout + server-side verification.
// Secret credentials are read only from Script Properties.
// ============================================================

function getFlutterwaveSecretKey_() {
  var key = PropertiesService.getScriptProperties().getProperty('FLUTTERWAVE_SECRET_KEY');
  if (!key) throw new Error('FLUTTERWAVE_SECRET_KEY is not configured in Script Properties.');
  return String(key).trim();
}

function getFlutterwaveCheckoutDetails(requestId) {
  try {
    if (!requestId) return {success:false,code:'REQUEST_ID_REQUIRED',message:'Request ID is required.'};

    var request = getUpgradeRequest(requestId);
    if (!request) return {success:false,code:'REQUEST_NOT_FOUND',message:'Upgrade request not found.'};

    var status = String(request.status || '').toLowerCase();
    if (status !== 'pending_payment') {
      return {success:false,code:'INVALID_PAYMENT_STATE',message:'This upgrade request is not awaiting payment.'};
    }

    var amount = Number(request.amount || 0);
    if (!amount || amount <= 0) {
      return {success:false,code:'INVALID_PAYMENT_AMOUNT',message:'Upgrade request has an invalid payment amount.'};
    }

    var currency = String(request.currency || 'NGN').toUpperCase();
    var txRef = 'BIZOS-' + String(requestId).replace(/[^A-Za-z0-9-]/g,'-') + '-' + Utilities.getUuid().replace(/-/g,'').substring(0,10);
    var publicUrl = getPublicBizOSUrl_();
    var redirectUrl = publicUrl + '/payment?requestId=' + encodeURIComponent(requestId) + '&provider=flutterwave';

    var payload = {
      tx_ref:txRef,
      amount:amount,
      currency:currency,
      redirect_url:redirectUrl,
      customer:{
        email:String(request.email || '').trim(),
        name:String(request.name || request.businessName || 'BizOS Customer').trim()
      },
      customizations:{
        title:'BizOS Sovereign Upgrade',
        description:'BizOS one-time Sovereign workspace upgrade',
        logo:publicUrl + '/favicon.ico'
      },
      meta:{requestId:requestId,businessName:String(request.businessName || '')}
    };

    var response = UrlFetchApp.fetch('https://api.flutterwave.com/v3/payments', {
      method:'post',
      contentType:'application/json',
      headers:{Authorization:'Bearer ' + getFlutterwaveSecretKey_()},
      payload:JSON.stringify(payload),
      muteHttpExceptions:true
    });

    var result = JSON.parse(response.getContentText() || '{}');
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300 || result.status !== 'success' || !result.data || !result.data.link) {
      console.error('Flutterwave checkout creation failed:', response.getContentText());
      return {success:false,code:'FLUTTERWAVE_CHECKOUT_FAILED',message:result.message || 'Unable to prepare Flutterwave payment.'};
    }

    saveFlutterwaveReference_(requestId,txRef);

    return {
      success:true,
      requestId:requestId,
      email:String(request.email || ''),
      businessName:String(request.businessName || ''),
      amount:amount,
      displayAmount:amount,
      currency:currency,
      txRef:txRef,
      paymentUrl:String(result.data.link)
    };
  } catch (error) {
    console.error('getFlutterwaveCheckoutDetails error:', error);
    return {success:false,code:'FLUTTERWAVE_CHECKOUT_CONFIG_ERROR',message:error.message || 'Unable to prepare Flutterwave payment.'};
  }
}

function saveFlutterwaveReference_(requestId, txRef) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payments');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  if (!data.length) return;
  var headers = data[0];
  var requestIdCol = headers.indexOf('Request_ID');
  var refCol = headers.indexOf('Transaction_Ref');
  if (requestIdCol === -1 || refCol === -1) return;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][requestIdCol]) === String(requestId)) {
      sheet.getRange(i + 1, refCol + 1).setValue(txRef);
      return;
    }
  }
}

function verifyFlutterwavePaymentAndProvisionSecure(transactionId, requestId, txRef) {
  try {
    if (!transactionId || !requestId) {
      return {success:false,code:'INVALID_PAYMENT_COMPLETION',message:'Flutterwave transaction ID and request ID are required.'};
    }

    var request = getUpgradeRequest(requestId);
    if (!request) return {success:false,code:'REQUEST_NOT_FOUND',message:'Upgrade request not found.'};

    var response = UrlFetchApp.fetch(
      'https://api.flutterwave.com/v3/transactions/' + encodeURIComponent(transactionId) + '/verify',
      {
        method:'get',
        headers:{Authorization:'Bearer ' + getFlutterwaveSecretKey_(),'Content-Type':'application/json'},
        muteHttpExceptions:true
      }
    );

    var result = JSON.parse(response.getContentText() || '{}');
    var tx = result.data || {};
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300 || result.status !== 'success' || String(tx.status || '').toLowerCase() !== 'successful') {
      return {success:false,code:'PAYMENT_VERIFICATION_FAILED',message:'Flutterwave did not confirm this transaction.'};
    }

    var expectedAmount = Number(request.amount || 0);
    var paidAmount = Number(tx.amount || 0);
    var expectedCurrency = String(request.currency || '').toUpperCase();
    var paidCurrency = String(tx.currency || '').toUpperCase();
    var expectedEmail = String(request.email || '').trim().toLowerCase();
    var paidEmail = String((tx.customer && tx.customer.email) || '').trim().toLowerCase();
    var expectedRef = String(txRef || request.flutterwaveReference || '').trim();
    var paidRef = String(tx.tx_ref || '').trim();

    if (expectedRef && paidRef && expectedRef !== paidRef) {
      return {success:false,code:'PAYMENT_REFERENCE_MISMATCH',message:'Payment reference does not match this upgrade request.'};
    }
    if (paidAmount < expectedAmount) {
      return {success:false,code:'PAYMENT_AMOUNT_MISMATCH',message:'Payment amount does not match the upgrade amount.'};
    }
    if (expectedCurrency && paidCurrency !== expectedCurrency) {
      return {success:false,code:'PAYMENT_CURRENCY_MISMATCH',message:'Payment currency does not match the upgrade currency.'};
    }
    if (expectedEmail && paidEmail && paidEmail !== expectedEmail) {
      return {success:false,code:'PAYMENT_EMAIL_MISMATCH',message:'Payment email does not match the upgrade request.'};
    }

    markPaymentConfirmed_(requestId,paidRef || expectedRef,String(transactionId));

    var provisioning = provisionConfirmedUpgradeRequest(requestId);
    if (!provisioning || !provisioning.success) {
      return {
        success:false,
        code:'PROVISIONING_FAILED_AFTER_PAYMENT',
        paymentVerified:true,
        message:'Payment was verified, but your BizOS workspace could not be finished yet. You can continue setup without paying again.',
        provisioning:provisioning
      };
    }

    return {
      success:true,
      paymentVerified:true,
      idempotent:!!provisioning.idempotent,
      requestId:requestId,
      reference:paidRef || expectedRef,
      transactionId:String(transactionId),
      clientId:provisioning.clientId,
      webAppUrl:provisioning.webAppUrl,
      landingUrl:provisioning.landingUrl,
      deploymentId:provisioning.deploymentId,
      message:provisioning.message || 'Your BizOS workspace is ready.'
    };
  } catch (error) {
    console.error('verifyFlutterwavePaymentAndProvisionSecure error:', error);
    return {success:false,code:'PAYMENT_PROVISIONING_ERROR',message:error.message || 'Payment completion failed.'};
  }
}


function getUpgradePaymentDetails(requestId) {
  try {
    if (!requestId) return {success:false,code:'REQUEST_ID_REQUIRED',message:'Request ID is required.'};
    var request=getUpgradeRequest(requestId);
    if (!request) return {success:false,code:'REQUEST_NOT_FOUND',message:'Upgrade request not found.'};
    var status=String(request.status||'').toLowerCase();
    if (status !== 'pending_payment') return {success:false,code:'INVALID_PAYMENT_STATE',message:'This upgrade request is not awaiting payment.'};
    var amount=Number(request.amount||0);
    if (!amount || amount<=0) return {success:false,code:'INVALID_PAYMENT_AMOUNT',message:'Upgrade request has an invalid payment amount.'};
    return {success:true,requestId:requestId,email:String(request.email||''),businessName:String(request.businessName||''),amount:amount,displayAmount:amount,currency:String(request.currency||'NGN').toUpperCase(),tier:String(request.tier||'sovereign')};
  } catch(error) {
    console.error('getUpgradePaymentDetails error:',error);
    return {success:false,code:'PAYMENT_DETAILS_ERROR',message:error.message||'Unable to load payment details.'};
  }
}


function createFlutterwaveCheckout(requestId) {
  try {
    var request=getUpgradeRequest(requestId);
    if (!request) return {success:false,code:'REQUEST_NOT_FOUND',message:'Upgrade request not found.'};
    if (String(request.status||'').toLowerCase() !== 'pending_payment') return {success:false,code:'INVALID_PAYMENT_STATE',message:'This upgrade request is not awaiting payment.'};
    var amount=Number(request.amount||0);
    if (!amount || amount<=0) return {success:false,code:'INVALID_PAYMENT_AMOUNT',message:'Upgrade request has an invalid payment amount.'};
    var currency=String(request.currency||'NGN').toUpperCase();
    var txRef='BIZOS-' + String(requestId).replace(/[^A-Za-z0-9._=-]/g,'-') + '-FLW-' + Utilities.getUuid().replace(/-/g,'').substring(0,8);
    var callbackBase=String(ScriptApp.getService().getUrl()||'').replace(/\/$/,'');
    if (!callbackBase) throw new Error('The BizOS web app URL is not available for Flutterwave callback handling.');
    var redirectUrl=callbackBase+'?page=flutterwave-callback&requestId='+encodeURIComponent(requestId);
    var payload={tx_ref:txRef,amount:amount,currency:currency,redirect_url:redirectUrl,customer:{email:String(request.email||''),name:String(request.name||request.email||'')},customizations:{title:'BizOS Upgrade',description:'Sovereign workspace upgrade for '+String(request.businessName||'your business')},meta:{requestId:requestId,businessName:String(request.businessName||''),tier:String(request.tier||'sovereign')}};
    var response=UrlFetchApp.fetch('https://api.flutterwave.com/v3/payments',{method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+getFlutterwaveSecretKey_()},payload:JSON.stringify(payload),muteHttpExceptions:true});
    var result=JSON.parse(response.getContentText()||'{}');
    if (String(result.status||'').toLowerCase() !== 'success' || !result.data || !result.data.link) { console.error('Flutterwave initialization failed:',response.getContentText()); return {success:false,code:'FLUTTERWAVE_INIT_FAILED',message:'Flutterwave could not prepare the secure checkout.'}; }
    saveFlutterwaveReference_(requestId,txRef);
    return {success:true,requestId:requestId,reference:txRef,checkoutUrl:String(result.data.link)};
  } catch(error) {
    console.error('createFlutterwaveCheckout error:',error);
    return {success:false,code:'FLUTTERWAVE_CHECKOUT_ERROR',message:error.message||'Unable to prepare Flutterwave checkout.'};
  }
}


function getStoredFlutterwaveReference_(requestId) {
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payments');
  if (!sheet) return '';
  var data=sheet.getDataRange().getValues(); if (!data.length) return '';
  var headers=data[0],requestIdCol=headers.indexOf('Request_ID'),refCol=headers.indexOf('Transaction_Ref');
  if (requestIdCol===-1||refCol===-1) return '';
  for (var i=1;i<data.length;i++) if(String(data[i][requestIdCol])===String(requestId)) return String(data[i][refCol]||'').trim();
  return '';
}


function handleFlutterwaveCallback(requestId,transactionId,status) {
  try {
    if (!requestId || !transactionId) return {success:false,code:'INVALID_FLUTTERWAVE_CALLBACK',message:'Flutterwave callback is missing required payment information.'};
    var request=getUpgradeRequest(requestId);
    if (!request) return {success:false,code:'REQUEST_NOT_FOUND',message:'Upgrade request not found.'};
    if (String(status||'').toLowerCase() !== 'successful') return {success:false,code:'FLUTTERWAVE_PAYMENT_NOT_SUCCESSFUL',message:'Flutterwave did not report a successful payment.'};
    var response=UrlFetchApp.fetch('https://api.flutterwave.com/v3/transactions/'+encodeURIComponent(transactionId)+'/verify',{method:'get',headers:{Authorization:'Bearer '+getFlutterwaveSecretKey_(),'Content-Type':'application/json'},muteHttpExceptions:true});
    var result=JSON.parse(response.getContentText()||'{}'),tx=result.data||{};
    if (String(result.status||'').toLowerCase()!=='success' || String(tx.status||'').toLowerCase()!=='successful') return {success:false,code:'FLUTTERWAVE_VERIFICATION_FAILED',message:'Flutterwave could not confirm this transaction.'};
    var storedReference=getStoredFlutterwaveReference_(requestId),paidReference=String(tx.tx_ref||'').trim();
    if (!storedReference || paidReference!==storedReference) return {success:false,code:'PAYMENT_REFERENCE_MISMATCH',message:'This Flutterwave transaction does not match the BizOS upgrade request.'};
    var expectedAmount=Number(request.amount||0),paidAmount=Number(tx.amount||0),expectedCurrency=String(request.currency||'').toUpperCase(),paidCurrency=String(tx.currency||'').toUpperCase(),expectedEmail=String(request.email||'').trim().toLowerCase(),paidEmail=String((tx.customer&&tx.customer.email)||'').trim().toLowerCase();
    if (paidAmount!==expectedAmount) return {success:false,code:'PAYMENT_AMOUNT_MISMATCH',message:'Payment amount does not match the upgrade amount.'};
    if (expectedCurrency && paidCurrency!==expectedCurrency) return {success:false,code:'PAYMENT_CURRENCY_MISMATCH',message:'Payment currency does not match the upgrade currency.'};
    if (expectedEmail && paidEmail && paidEmail!==expectedEmail) return {success:false,code:'PAYMENT_EMAIL_MISMATCH',message:'Payment email does not match the upgrade request.'};
    markPaymentConfirmed_(requestId,paidReference,tx);
    var provisioning=provisionConfirmedUpgradeRequest(requestId);
    if (!provisioning || !provisioning.success) return {success:false,code:'PROVISIONING_FAILED_AFTER_PAYMENT',paymentVerified:true,message:'Payment was verified, but your BizOS workspace could not be finished yet. You can continue setup without paying again.',provisioning:provisioning};
    return {success:true,paymentVerified:true,idempotent:!!provisioning.idempotent,requestId:requestId,clientId:provisioning.clientId,webAppUrl:provisioning.webAppUrl,landingUrl:provisioning.landingUrl};
  } catch(error) {
    console.error('handleFlutterwaveCallback error:',error);
    return {success:false,code:'FLUTTERWAVE_CALLBACK_ERROR',message:error.message||'Flutterwave payment completion failed.'};
  }
}

