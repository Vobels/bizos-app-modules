// ============================================================
// ClientDeploymentV3.js - FINAL CLIENT PACKAGE ADAPTER
// ============================================================
// Reuses the isolated V2 server/UI package and fixes browser-side
// CRUD invocation so every one of the 12 modules supports create,
// read, update and delete without touching the free workspace.
// ============================================================
function generateClientCodeSafelyV3(settings){
  var pkg=generateClientCodeSafelyV2(settings);
  if(!pkg||!pkg.files)throw new Error('Client deployment package is empty.');
  pkg.files=pkg.files.map(function(f){
    if(f.name!=='client-dashboard')return f;
    var s=f.source;
    var old="var fn=EDITROW?'updateClientModuleRecord':'createClientModuleRecord';var args=EDITROW?[CURRENT,EDITROW,record,SID]:[CURRENT,record,SID];google.script.run.withSuccessHandler(function(r){if(!r||!r.success){alert(r&&r.message||'Save failed');return}closeModal();openModule(CURRENT)}).withFailureHandler(function(){alert('Save failed')})[fn].apply(google.script.run,args)";
    var replacement="var runner=google.script.run.withSuccessHandler(function(r){if(!r||!r.success){alert(r&&r.message||'Save failed');return}closeModal();openModule(CURRENT)}).withFailureHandler(function(){alert('Save failed')});if(EDITROW)runner.updateClientModuleRecord(CURRENT,EDITROW,record,SID);else runner.createClientModuleRecord(CURRENT,record,SID)";
    if(s.indexOf(old)<0)throw new Error('Client dashboard CRUD handler not found.');
    f.source=s.replace(old,replacement);
    return f;
  });
  return pkg;
}
