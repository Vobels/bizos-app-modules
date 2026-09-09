// ============================================================
// ClientDeploymentV5.js - CLIENT DEPLOYMENT UX + AUTH HARDENING
// ============================================================

function generateClientCodeSafelyV5(settings) {
  var base = generateClientCodeSafelyV4(settings);
  if (!base || !base.files || !base.files.length) {
    throw new Error('Client deployment package is empty.');
  }

  base.files = base.files.map(function (file) {
    if (!file) return file;

    if (file.name === 'Code' && typeof file.source === 'string') {
      file.source += buildAuthoritativeClientAuthOverride_(settings);
      return file;
    }

    var isLanding = file.name === 'client-landing' || file.name === 'client-landing.html';
    var isDashboard = file.name === 'client-dashboard' || file.name === 'client-dashboard.html';
    if (!isLanding && !isDashboard) return file;

    var html = typeof file.source === 'string' ? file.source :
               (typeof file.content === 'string' ? file.content : '');
    if (!html) return file;

    html = addClientLoadingUX_(html, file.name);

    if (typeof file.source === 'string' || file.source === undefined) file.source = html;
    if (typeof file.content === 'string') file.content = html;
    return file;
  });

  return base;
}

// The generated V2 package contains an older inline authenticateClient().
// This later declaration is intentionally the authoritative implementation:
// credentials are checked by the master API's clientLogin action, which reads
// the master Clients registry. The generated deployment then creates only its
// own local session, bound to its immutable CLIENT_CONFIG.
function buildAuthoritativeClientAuthOverride_(settings) {
  return `

// ============================================================
// AUTHORITATIVE CLIENT AUTH OVERRIDE - V5
// ============================================================
function authenticateClient(email,password){
  try{
    email=String(email||'').trim().toLowerCase();
    password=String(password||'');
    if(!email||!password)return{success:false,message:'Email and password are required',code:'INVALID_INPUT'};

    var response=UrlFetchApp.fetch(CLIENT_CONFIG.masterApiUrl,{
      method:'post',
      contentType:'application/json',
      muteHttpExceptions:true,
      payload:JSON.stringify({
        action:'clientLogin',
        payload:{
          clientId:CLIENT_CONFIG.clientId,
          email:email,
          password:password
        }
      })
    });

    var httpCode=response.getResponseCode();
    var body=response.getContentText();
    if(httpCode<200||httpCode>=300)return{success:false,message:'Authentication service unavailable.',code:'AUTH_SERVICE'};

    var result=JSON.parse(body||'{}');
    if(!result||!result.success)return result||{success:false,message:'Invalid email or password',code:'AUTH_FAILED'};

    if(!result.user||String(result.user.clientId||'')!==String(CLIENT_CONFIG.clientId)){
      return{success:false,message:'This account is not authorized for this deployment.',code:'CLIENT_BINDING_MISMATCH'};
    }
    if(result.user.businessId&&String(result.user.businessId)!==String(CLIENT_CONFIG.businessId)){
      return{success:false,message:'This account is not assigned to this business.',code:'BUSINESS_MISMATCH'};
    }
    if(result.clientName&&String(result.clientName)!==String(CLIENT_CONFIG.clientName)){
      return{success:false,message:'Client deployment identity mismatch.',code:'CLIENT_NAME_MISMATCH'};
    }

    var u=result.user||{};
    var now=Date.now();
    var sid='CLIENT_SESS_'+Utilities.getUuid();
    var session={
      email:email,
      clientId:String(CLIENT_CONFIG.clientId),
      sheetId:String(CLIENT_CONFIG.sheetId),
      businessId:String(CLIENT_CONFIG.businessId),
      businessName:String(CLIENT_CONFIG.clientName),
      created:now,
      expiresAt:now+21600000,
      isClient:true,
      user:u
    };
    CacheService.getScriptCache().put(sid,JSON.stringify(session),21600);

    return{
      success:true,
      sessionId:sid,
      user:{
        email:u.email||email,
        name:u.name||email.split('@')[0],
        role:u.role||'owner',
        clientId:String(CLIENT_CONFIG.clientId),
        businessId:String(CLIENT_CONFIG.businessId)
      },
      clientName:CLIENT_CONFIG.clientName,
      primaryColor:CLIENT_CONFIG.primaryColor,
      logoUrl:CLIENT_CONFIG.logoUrl,
      code:'SUCCESS'
    };
  }catch(error){
    console.error('Authoritative client authentication error:',error);
    return{success:false,message:error&&error.message?error.message:'Login failed',code:'LOGIN_ERROR'};
  }
}
`;
}

function addClientLoadingUX_(html, fileName) {
  var css = '<style id="bizos-loading-ux">' +
    '#bizosLoader{position:fixed;inset:0;z-index:99999;display:none;align-items:center;justify-content:center;padding:20px;box-sizing:border-box;background:rgba(255,255,255,.94);backdrop-filter:blur(3px);-webkit-backdrop-filter:blur(3px)}' +
    '#bizosLoader.show{display:flex}' +
    '.bizosLoaderBox{width:min(360px,92vw);text-align:center;background:#fff;border:1px solid #e5e7eb;border-radius:16px;padding:24px;box-shadow:0 12px 40px rgba(0,0,0,.12);box-sizing:border-box}' +
    '.bizosSpinner{width:34px;height:34px;border:3px solid #e5e7eb;border-top-color:<?= primaryColor ?>;border-radius:50%;animation:bizosSpin .8s linear infinite;margin:0 auto 14px}' +
    '.bizosLoaderTitle{font-weight:700;color:#17202a;margin-bottom:6px}' +
    '.bizosLoaderText{font-size:13px;color:#667085;line-height:1.45}' +
    '.bizosLoaderError{display:none;margin-top:14px;color:#b42318;font-size:13px;line-height:1.45}' +
    '@keyframes bizosSpin{to{transform:rotate(360deg)}}' +
    'html,body{max-width:100%;overflow-x:hidden}' +
    '@media(max-width:600px){.wrap{padding:14px!important}.bar{padding:14px!important}.modalbox{max-height:86vh!important;width:94%!important}.tablewrap{-webkit-overflow-scrolling:touch;max-width:100%}}' +
    '</style>';

  var overlay = '<div id="bizosLoader" aria-live="polite" aria-busy="false">' +
    '<div class="bizosLoaderBox"><div class="bizosSpinner"></div>' +
    '<div id="bizosLoaderTitle" class="bizosLoaderTitle">Please wait</div>' +
    '<div id="bizosLoaderText" class="bizosLoaderText">BizOS is processing your request...</div>' +
    '<div id="bizosLoaderError" class="bizosLoaderError"></div></div></div>';

  var js = '<script id="bizos-loading-ux-js">(function(){' +
    'var timer=null,hideTimer=null,active=false,observer=null;' +
    'function el(id){return document.getElementById(id);}' +
    'function stopObserver(){if(observer){observer.disconnect();observer=null;}}' +
    'window.bizosShowLoader=function(title,text,timeout){var o=el("bizosLoader");if(!o)return;el("bizosLoaderTitle").textContent=title||"Please wait";el("bizosLoaderText").textContent=text||"BizOS is processing your request...";el("bizosLoaderError").style.display="none";el("bizosLoaderError").textContent="";o.classList.add("show");o.setAttribute("aria-busy","true");active=true;clearTimeout(timer);clearTimeout(hideTimer);stopObserver();if(window.MutationObserver){observer=new MutationObserver(function(){if(!active)return;clearTimeout(hideTimer);hideTimer=setTimeout(function(){if(active)window.bizosHideLoader();},350);});observer.observe(document.body,{subtree:true,childList:true,characterData:true,attributes:true});}timer=setTimeout(function(){if(!active)return;el("bizosLoaderTitle").textContent="This is taking longer than expected";el("bizosLoaderText").textContent="The server may be busy or your connection may be slow.";el("bizosLoaderError").textContent="You can wait a little longer or refresh and try again.";el("bizosLoaderError").style.display="block";},Number(timeout||20000));};' +
    'window.bizosHideLoader=function(){var o=el("bizosLoader");if(!o)return;active=false;clearTimeout(timer);clearTimeout(hideTimer);stopObserver();o.classList.remove("show");o.setAttribute("aria-busy","false");};' +
    'window.addEventListener("error",function(){if(active){el("bizosLoaderTitle").textContent="Something went wrong";el("bizosLoaderText").textContent="BizOS could not complete the request.";el("bizosLoaderError").style.display="block";el("bizosLoaderError").textContent="Please try again.";}});' +
    'document.addEventListener("DOMContentLoaded",function(){var buttons=document.querySelectorAll("button");Array.prototype.forEach.call(buttons,function(b){b.addEventListener("click",function(){var text=(b.textContent||"").trim().toLowerCase(),onclick=b.getAttribute("onclick")||"";if(text==="×"||text.indexOf("dashboard")>=0||text==="logout"||text.indexOf("add record")>=0||onclick.indexOf("openAdd")>=0)return;if(b.disabled)return;var serverAction=onclick.indexOf("google.script.run")>=0||onclick.indexOf("openModule")>=0||onclick.indexOf("load")>=0||onclick.indexOf("save")>=0||onclick.indexOf("delete")>=0||onclick.indexOf("login")>=0;if(!serverAction)return;window.bizosShowLoader(text.indexOf("sign")>=0?"Signing in":"Working","Please wait while BizOS completes this request...",20000);},true);});});' +
    'window.addEventListener("pageshow",function(){window.bizosHideLoader();});' +
  '})();</script>';

  html = html.replace(/Loading\.\.\./g, '<span class="bizosInlineLoading"><span class="bizosSpinner" style="display:inline-block;width:16px;height:16px;border-width:2px;margin:0 7px 0 0;vertical-align:-3px"></span>Loading...</span>');

  if (html.indexOf('id="bizosLoadingUXMarker"') === -1) {
    html = html.replace('</head>', '<meta id="bizosLoadingUXMarker" name="bizos-loading-ux" content="v5">' + css + '</head>');
    html = html.replace('</body>', overlay + js + '</body>');
  }
  return html;
}
