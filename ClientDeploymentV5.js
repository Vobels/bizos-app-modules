// ============================================================
// ClientDeploymentV5.js - CLIENT DEPLOYMENT UX HARDENING
// ============================================================
// Wraps the tested V4 generator without changing its server/data
// logic. Adds visible loading states, mobile-safe scrolling and
// a timeout/failure state so the client UI never appears frozen.
// ============================================================

function generateClientCodeSafelyV5(settings) {
  var base = generateClientCodeSafelyV4(settings);
  if (!base || !base.files || !base.files.length) {
    throw new Error('Client deployment package is empty.');
  }

  base.files = base.files.map(function (file) {
    if (!file || typeof file.content !== 'string') return file;
    if (file.name === 'client-landing.html' || file.name === 'client-dashboard.html') {
      file.content = addClientLoadingUX_(file.content, file.name);
    }
    return file;
  });

  return base;
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
    '<div class="bizosLoaderBox">' +
    '<div class="bizosSpinner"></div>' +
    '<div id="bizosLoaderTitle" class="bizosLoaderTitle">Please wait</div>' +
    '<div id="bizosLoaderText" class="bizosLoaderText">BizOS is processing your request...</div>' +
    '<div id="bizosLoaderError" class="bizosLoaderError"></div>' +
    '</div></div>';

  var js = '<script id="bizos-loading-ux-js">(function(){' +
    'var timer=null,hideTimer=null,active=false,observer=null;' +
    'function el(id){return document.getElementById(id);}' +
    'function stopObserver(){if(observer){observer.disconnect();observer=null;}}' +
    'window.bizosShowLoader=function(title,text,timeout){' +
      'var o=el("bizosLoader");if(!o)return;' +
      'el("bizosLoaderTitle").textContent=title||"Please wait";' +
      'el("bizosLoaderText").textContent=text||"BizOS is processing your request...";' +
      'el("bizosLoaderError").style.display="none";el("bizosLoaderError").textContent="";' +
      'o.classList.add("show");o.setAttribute("aria-busy","true");active=true;' +
      'clearTimeout(timer);clearTimeout(hideTimer);stopObserver();' +
      'if(window.MutationObserver){observer=new MutationObserver(function(){if(!active)return;clearTimeout(hideTimer);hideTimer=setTimeout(function(){if(active)window.bizosHideLoader();},350);});observer.observe(document.body,{subtree:true,childList:true,characterData:true,attributes:true});}' +
      'timer=setTimeout(function(){' +
        'if(!active)return;' +
        'el("bizosLoaderTitle").textContent="This is taking longer than expected";' +
        'el("bizosLoaderText").textContent="The server may be busy or your connection may be slow.";' +
        'el("bizosLoaderError").textContent="You can wait a little longer or refresh and try again.";' +
        'el("bizosLoaderError").style.display="block";' +
      '},Number(timeout||20000));' +
    '};' +
    'window.bizosHideLoader=function(){var o=el("bizosLoader");if(!o)return;active=false;clearTimeout(timer);clearTimeout(hideTimer);stopObserver();o.classList.remove("show");o.setAttribute("aria-busy","false");};' +
    'window.addEventListener("error",function(){if(active){el("bizosLoaderTitle").textContent="Something went wrong";el("bizosLoaderText").textContent="BizOS could not complete the request.";el("bizosLoaderError").style.display="block";el("bizosLoaderError").textContent="Please try again.";}});' +
    'document.addEventListener("DOMContentLoaded",function(){' +
      'var buttons=document.querySelectorAll("button");' +
      'Array.prototype.forEach.call(buttons,function(b){' +
        'b.addEventListener("click",function(){' +
          'var text=(b.textContent||"").trim().toLowerCase(),onclick=b.getAttribute("onclick")||"";' +
          'if(text==="×"||text.indexOf("dashboard")>=0||text==="logout"||text.indexOf("add record")>=0||onclick.indexOf("openAdd")>=0)return;' +
          'if(b.disabled)return;' +
          'var serverAction=onclick.indexOf("google.script.run")>=0||onclick.indexOf("openModule")>=0||onclick.indexOf("load")>=0||onclick.indexOf("save")>=0||onclick.indexOf("delete")>=0||onclick.indexOf("login")>=0;' +
          'if(!serverAction)return;' +
          'window.bizosShowLoader(text.indexOf("sign")>=0?"Signing in":"Working","Please wait while BizOS completes this request...",20000);' +
        '},true);' +
      '});' +
    '});' +
    'window.addEventListener("pageshow",function(){window.bizosHideLoader();});' +
  '})();</script>';

  // Turn existing plain loading labels into a visible spinner.
  html = html.replace(/Loading\.\.\./g, '<span class="bizosInlineLoading"><span class="bizosSpinner" style="display:inline-block;width:16px;height:16px;border-width:2px;margin:0 7px 0 0;vertical-align:-3px"></span>Loading...</span>');

  if (html.indexOf('id="bizosLoadingUXMarker"') === -1) {
    html = html.replace('</head>', '<meta id="bizosLoadingUXMarker" name="bizos-loading-ux" content="v5">' + css + '</head>');
    html = html.replace('</body>', overlay + js + '</body>');
  }
  return html;
}
