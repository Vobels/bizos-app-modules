// ============================================================
// FlutterwaveSecure.js
// Secure Flutterwave checkout + server-side verification.
// Secrets are read only from Apps Script Script Properties.
// ============================================================

function getFlutterwaveEnabledCurrencies_() {
  try {
    var raw=PropertiesService.getScriptProperties().getProperty('FLUTTERWAVE_ENABLED_CURRENCIES');
    if(raw){
      var list=String(raw).split(',').map(function(x){return String(x||'').trim().toUpperCase();}).filter(Boolean);
      if(list.length)return list;
    }
  }catch(e){}
  return ['NGN','USD'];
}

function getFlutterwaveSecretKey_() {
  var key = PropertiesService.getScriptProperties().getProperty('FLUTTERWAVE_SECRET_KEY');
  if (!key) throw new Error('FLUTTERWAVE_SECRET_KEY is not configured in Script Properties.');
  return String(key).trim();
}

function createFlutterwaveCheckout(requestId, sessionId, accessToken) {
  try {
    var user=sessionId ? validateUpgradeSession(sessionId) : null;
    if (!user && accessToken) {
      var access=validatePaymentAccessToken_(accessToken, requestId);
      if (access) user={email:access.email};
    }
    if (!user) return {success:false,code:'UNAUTHORIZED',message:'Your payment-page access has expired. Please return to BizOS and reopen the payment request.'};
    var request=getUpgradeRequest(requestId);
    if (!request) return {success:false,code:'REQUEST_NOT_FOUND',message:'Upgrade request not found.'};
    if (String(request.email||'').trim().toLowerCase() !== String(user.email||'').trim().toLowerCase()) return {success:false,code:'REQUEST_ACCESS_DENIED',message:'This upgrade request does not belong to the current BizOS account.'};
    if (String(request.status||'').toLowerCase() !== 'pending_payment') return {success:false,code:'INVALID_PAYMENT_STATE',message:'This upgrade request is not awaiting payment.'};
    var amount=Number(request.amount||0);
    if (!amount || amount<=0) return {success:false,code:'INVALID_PAYMENT_AMOUNT',message:'Upgrade request has an invalid payment amount.'};
    var currency=String(request.currency||'NGN').toUpperCase();
    var txRef='BIZOS-' + String(requestId).replace(/[^A-Za-z0-9._=-]/g,'-') + '-FLW-' + Utilities.getUuid().replace(/-/g,'').substring(0,8);
    var callbackBase=String(ScriptApp.getService().getUrl()||'').replace(/\/$/,'');
    if (!callbackBase) throw new Error('The Apps Script web app URL is not available.');
    var redirectUrl=callbackBase+'/?page=flutterwave-callback&requestId='+encodeURIComponent(requestId);
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

function saveFlutterwaveReference_(requestId,reference) {
  var sheet=getBizOSMasterSpreadsheet_().getSheetByName('Payments');
  if (!sheet) return;
  var data=sheet.getDataRange().getValues(); if (!data.length) return;
  var headers=data[0],requestIdCol=headers.indexOf('Request_ID'),refCol=headers.indexOf('Transaction_Ref');
  if (requestIdCol===-1||refCol===-1) return;
  for (var i=1;i<data.length;i++) if(String(data[i][requestIdCol])===String(requestId)){sheet.getRange(i+1,refCol+1).setValue(reference);return;}
}

function getStoredFlutterwaveReference_(requestId) {
  var sheet=getBizOSMasterSpreadsheet_().getSheetByName('Payments');
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


/**
 * Resolve the payment providers that are actually appropriate for a BizOS
 * upgrade request. The customer should only see routes that BizOS can safely
 * attempt for the selected currency.
 *
 * Paystack currently supports NGN and USD for Nigerian businesses. USD is
 * intentionally feature-flagged until the merchant account has USD settlement
 * enabled. Flutterwave is the international hosted-checkout route for the
 * currencies supported by its card checkout.
 */
function getAvailablePaymentMethods(requestId, sessionId, accessToken) {
  try {
    if (!requestId) return {success:false,code:'REQUEST_ID_REQUIRED',message:'Upgrade request ID is required.'};

    var user = sessionId ? validateUpgradeSession(sessionId) : null;
    if (!user && accessToken) {
      var access = validatePaymentAccessToken_(accessToken, requestId);
      if (access) user = {email:access.email};
    }
    if (!user) return {success:false,code:'UNAUTHORIZED',message:'Your payment-page access has expired. Please return to BizOS and reopen the payment request.'};

    var request = getUpgradeRequest(requestId);
    if (!request) return {success:false,code:'REQUEST_NOT_FOUND',message:'Upgrade request not found.'};
    if (String(request.email || '').trim().toLowerCase() !== String(user.email || '').trim().toLowerCase()) {
      return {success:false,code:'REQUEST_ACCESS_DENIED',message:'This upgrade request does not belong to the current BizOS account.'};
    }
    if (String(request.status || '').toLowerCase() !== 'pending_payment') {
      return {success:false,code:'INVALID_PAYMENT_STATE',message:'This upgrade request is no longer awaiting payment.'};
    }

    var currency = String(request.currency || '').trim().toUpperCase();
    if (!currency) return {success:false,code:'CURRENCY_REQUIRED',message:'A payment currency is required before checkout can continue.'};
    if (!getPaymentCurrencyConfig_(currency)) return {success:false,code:'CURRENCY_UNAVAILABLE',message:'This currency is no longer available for BizOS checkout.'};

    var methods = [];
    var paystackConfigured = false;
    var flutterwaveConfigured = false;
    try { paystackConfigured = !!PropertiesService.getScriptProperties().getProperty('PAYSTACK_PUBLIC_KEY') && !!PropertiesService.getScriptProperties().getProperty('PAYSTACK_SECRET_KEY'); } catch (e) {}
    try { flutterwaveConfigured = !!PropertiesService.getScriptProperties().getProperty('FLUTTERWAVE_SECRET_KEY'); } catch (e) {}

    // Paystack's Nigeria account supports NGN and USD. USD is opt-in here so
    // enabling the code path does not accidentally promise USD settlement before
    // the merchant has completed Paystack's USD setup.
    var paystackUsdEnabled = false;
    try {
      var configured = PropertiesService.getScriptProperties().getProperty('PAYSTACK_USD_ENABLED');
      paystackUsdEnabled = String(configured || '').trim().toLowerCase() === 'true';
    } catch (e) {}

    if (paystackConfigured && (currency === 'NGN' || (currency === 'USD' && paystackUsdEnabled))) {
      methods.push({id:'paystack',name:'Paystack',available:true,description:currency === 'NGN' ? 'Secure card and supported local payment options.' : 'Secure international card checkout.'});
    }

    // Flutterwave capabilities are centralized in BizOS payment config.
    // Account-level approval/settings can still restrict a currency, so the
    // actual checkout remains the final server-side authority.
    var flutterwaveCurrencies = getFlutterwaveEnabledCurrencies_();
    if (flutterwaveConfigured && flutterwaveCurrencies.indexOf(currency) !== -1) {
      methods.push({id:'flutterwave',name:'Flutterwave',available:true,description:'Secure international card and supported payment options.'});
    }

    return {
      success:true,
      requestId:String(request.requestId || requestId),
      currency:currency,
      amount:Number(request.amount || 0),
      country:String(request.country || ''),
      methods:methods,
      hasPaymentRoute:methods.length > 0,
      message:methods.length ? '' : 'There is currently no available payment method for this currency. Please choose another supported currency or contact BizOS support.'
    };
  } catch (error) {
    console.error('getAvailablePaymentMethods error:',error);
    return {success:false,code:'PAYMENT_METHODS_ERROR',message:'We could not determine the available payment methods right now. Please try again.'};
  }
}
