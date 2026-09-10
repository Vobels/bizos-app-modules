// ============================================================
// ZZZ_ClientDashboardMobileFix.js
// Targeted mobile fix for generated paid-client dashboard.
// Does not change the landing page or global modal controller.
// ============================================================

(function(){
  var baseGenerate = generateClientCodeSafelyV7;

  generateClientCodeSafelyV7 = function(settings){
    var result = baseGenerate(settings);
    if (!result || !result.files) return result;

    result.files = result.files.map(function(file){
      if (!file || (file.name !== 'client-dashboard' && file.name !== 'client-dashboard.html')) return file;

      var html = typeof file.source === 'string' ? file.source :
                 (typeof file.content === 'string' ? file.content : '');

      var mobileCss = `
<style id="bizos-client-mobile-scroll-fix">
html{width:100%;min-height:100%;}
body{width:100%;min-width:0;margin:0;position:static!important;height:auto!important;min-height:100vh;overflow-y:auto!important;overflow-x:hidden;}
.content{min-width:0;max-width:100%;}
@media(max-width:767px){
  html,body{width:100%;min-width:0;height:auto!important;min-height:100%;overflow-y:auto!important;position:static!important;}
  body{display:block!important;}
  .content{width:100%!important;max-width:100%!important;min-width:0!important;margin-left:0!important;}
  main.content{min-height:100dvh!important;height:auto!important;}
  section{max-width:100%!important;min-width:0!important;}
  #home,#apps,#profile,#settings,#module,#team{width:100%!important;min-width:0!important;max-width:100%!important;box-sizing:border-box;}
  #home>div,#apps>div,#profile>div,#settings>div,#module>div,#team>div{min-width:0;max-width:100%;}
}
</style>`;

      var mobileJs = `
<script id="bizos-client-mobile-scroll-fix-js">
(function(){
  function resetDashboardScrollV7_(){
    try{
      document.documentElement.style.position='static';
      document.documentElement.style.height='auto';
      document.documentElement.style.minHeight='100%';
      document.documentElement.style.overflowY='auto';
      document.body.style.position='static';
      document.body.style.height='auto';
      document.body.style.minHeight='100vh';
      document.body.style.overflowY='auto';
      document.body.style.overflowX='hidden';
      document.body.classList.remove('modal-open');
    }catch(e){}
  }

  function wrapNavV7_(name){
    var original=window[name];
    if(typeof original!=='function'||original.__bizosMobileWrapped)return;
    var wrapped=function(){
      resetDashboardScrollV7_();
      var result=original.apply(this,arguments);
      window.setTimeout(resetDashboardScrollV7_,0);
      window.setTimeout(resetDashboardScrollV7_,80);
      window.setTimeout(resetDashboardScrollV7_,300);
      return result;
    };
    wrapped.__bizosMobileWrapped=true;
    window[name]=wrapped;
  }

  resetDashboardScrollV7_();
  ['showHome','showApps','showProfile','showSettings','showTeam','openModule'].forEach(wrapNavV7_);
  window.addEventListener('resize',resetDashboardScrollV7_,{passive:true});
  window.addEventListener('orientationchange',function(){window.setTimeout(resetDashboardScrollV7_,100);});
})();
</script>`;

      if (html.indexOf('bizos-client-mobile-scroll-fix') === -1) {
        html = html.replace('</head>', mobileCss + '</head>');
      }
      if (html.indexOf('bizos-client-mobile-scroll-fix-js') === -1) {
        html = html.replace('</body></html>', mobileJs + '</body></html>');
      }

      file.source = html;
      delete file.content;
      return file;
    });

    return result;
  };
})();
