// ============================================================
// ClientDeploymentV3.js - CLIENT PACKAGE COMPATIBILITY ADAPTER
// ============================================================
// Reuses V2 and normalizes CRUD only when an older V2 package
// still contains the legacy dynamic invocation. Current V2 already
// uses explicit CRUD calls, so V3 leaves it unchanged.
// ============================================================
function generateClientCodeSafelyV3(settings){
  var pkg=generateClientCodeSafelyV2(settings);
  if(!pkg||!pkg.files)throw new Error('Client deployment package is empty.');

  pkg.files=pkg.files.map(function(f){
    if(!f)return f;
    var isDashboard=f.name==='client-dashboard'||f.name==='client-dashboard.html';
    if(!isDashboard)return f;

    var s=String(f.source||f.content||'');
    if(!s)throw new Error('Client dashboard source is empty.');

    // Current V2 already contains the explicit, Apps Script-compatible
    // CRUD calls. Do not attempt to rewrite it again.
    var hasExplicit=s.indexOf('updateClientModuleRecord(CURRENT,EDITROW,record,SID)')>=0 &&
      s.indexOf('createClientModuleRecord(CURRENT,record,SID)')>=0;
    if(hasExplicit){
      if(f.source!==undefined)f.source=s;else f.content=s;
      return f;
    }

    // Compatibility path for older V2 output.
    var old="var fn=EDITROW?'updateClientModuleRecord':'createClientModuleRecord';var args=EDITROW?[CURRENT,EDITROW,record,SID]:[CURRENT,record,SID];google.script.run.withSuccessHandler(function(r){if(!r||!r.success){alert(r&&r.message||'Save failed');return}closeModal();openModule(CURRENT)}).withFailureHandler(function(){alert('Save failed')})[fn].apply(google.script.run,args)";
    var replacement="var runner=google.script.run.withSuccessHandler(function(r){if(!r||!r.success){alert(r&&r.message||'Save failed');return}closeModal();openModule(CURRENT)}).withFailureHandler(function(){alert('Save failed')});if(EDITROW)runner.updateClientModuleRecord(CURRENT,EDITROW,record,SID);else runner.createClientModuleRecord(CURRENT,record,SID)";
    if(s.indexOf(old)>=0)s=s.replace(old,replacement);

    // If neither form exists, fail clearly instead of silently producing a
    // deployment whose save button cannot work.
    var finalExplicit=s.indexOf('updateClientModuleRecord(CURRENT,EDITROW,record,SID)')>=0 &&
      s.indexOf('createClientModuleRecord(CURRENT,record,SID)')>=0;
    if(!finalExplicit)throw new Error('Client dashboard CRUD handler not found.');

    if(f.source!==undefined)f.source=s;else f.content=s;
    return f;
  });

  return pkg;
}
