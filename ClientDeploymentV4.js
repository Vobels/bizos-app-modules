// ============================================================
// ClientDeploymentV4.js - MODULE 7-12 ISOLATION + PERSISTENCE TEST SUITE
// ============================================================
// Extends the isolated V3 client package with a production-style
// verification harness for:
//   Tax, Agro, Productivity, POS, Attendance, Warehouse.
//
// Each module is tested for:
//   1. Create
//   2. Read
//   3. Update
//   4. Delete
//   5. Persistence across a fresh session
//   6. Isolation from other modules / alternate workspace IDs
//
// The dashboard is also checked to ensure its workspaceId is the
// deployment's configured client workspace and that smoke-test data
// does not appear in financial/recent dashboard data.
//
// IMPORTANT:
// This suite never trusts a browser-supplied sheet/workspace ID.
// True two-deployment cross-client verification still requires the
// same suite to be run independently against Client A and Client B.
// ============================================================

function generateClientCodeSafelyV4(settings){
  var pkg=generateClientCodeSafelyV3(settings);
  if(!pkg||!pkg.files)throw new Error('Client deployment package is empty.');

  var target=null;
  pkg.files.forEach(function(f){
    if(f.name==='Code')target=f;
  });
  if(!target)throw new Error('Generated client Code file not found.');

  target.source += `

// ============================================================
// V4 TEST HELPERS
// ============================================================
function v4NewTestSession_(sid){
  var current=getSession_(sid);
  var fresh='CLIENT_TEST_SESS_'+Utilities.getUuid();
  var now=Date.now();
  CacheService.getScriptCache().put(fresh,JSON.stringify({
    email:current.email,
    clientId:current.clientId,
    sheetId:current.sheetId,
    businessId:current.businessId,
    businessName:current.businessName,
    user:current.user,
    created:now,
    expiresAt:now+21600000
  }),21600);
  return fresh;
}

function v4Assert_(condition,message){
  if(!condition)throw new Error(message);
}

function v4RecordExists_(moduleName,row,sid){
  var r=getClientModuleData(moduleName,sid);
  if(!r||!r.success)throw new Error(r&&r.message||'Read failed');
  return r.records.some(function(x){return Number(x._row)===Number(row);});
}

function v4GetRecord_(moduleName,row,sid){
  var r=getClientModuleData(moduleName,sid);
  if(!r||!r.success)throw new Error(r&&r.message||'Read failed');
  return r.records.filter(function(x){return Number(x._row)===Number(row);})[0]||null;
}

function v4ContainsMarker_(value,marker){
  if(value===null||value===undefined)return false;
  return String(value).indexOf(marker)>=0;
}

function v4BuildRecord_(headers,marker){
  var record={};
  headers.forEach(function(h){
    if(h)record[String(h)]=marker;
  });
  return record;
}

// ============================================================
// SINGLE MODULE TEST
// ============================================================
function v4TestModule_(name,sid){
  var started=Date.now();
  var item={
    module:name,
    create:false,
    read:false,
    update:false,
    persistence:false,
    isolation:false,
    delete:false,
    cleanup:false,
    pass:false,
    row:null,
    marker:'',
    error:'',
    durationMs:0
  };

  var row=null;
  var testSid=null;
  var marker='V4_'+name.toUpperCase()+'_'+Utilities.getUuid();
  item.marker=marker;

  try{
    getSession_(sid);
    var m=module_(name);
    var ss=getWorkspace_();
    var configuredWorkspaceId=String(CLIENT_CONFIG.sheetId);

    // The only workspace used by the test must be the immutable deployment config.
    v4Assert_(String(ss.getId())===configuredWorkspaceId,
      'Workspace resolution mismatch for '+name+'.');

    var sh=ss.getSheetByName(m.sheet);
    v4Assert_(!!sh,'Module sheet missing: '+m.sheet);

    var headers=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0]
      .filter(function(h){return h!==''&&h!==null&&h!==undefined;});
    v4Assert_(headers.length>0,'Module has no headers: '+name);

    var record=v4BuildRecord_(headers,marker);
    var created=createClientModuleRecord(name,record,sid);
    v4Assert_(created&&created.success,created&&created.message||'Create failed');
    item.create=true;
    row=created.record&&created.record._row;
    item.row=row;
    v4Assert_(row,'Create did not return a row number.');

    // Immediate read.
    var createdRecord=v4GetRecord_(name,row,sid);
    v4Assert_(!!createdRecord,'Created record cannot be read.');
    item.read=true;

    // Verify the test marker exists in the created record.
    var markerFound=false;
    Object.keys(createdRecord).forEach(function(k){
      if(k!=='_row'&&v4ContainsMarker_(createdRecord[k],marker))markerFound=true;
    });
    v4Assert_(markerFound,'Created record marker was not persisted.');

    // Update one real module column.
    var updateHeader=headers[0];
    var updateValue=marker+'_UPDATED';
    var patch={};
    patch[updateHeader]=updateValue;
    var updated=updateClientModuleRecord(name,row,patch,sid);
    v4Assert_(updated&&updated.success,updated&&updated.message||'Update failed');
    item.update=true;

    // Read after update and verify actual stored value.
    var updatedRecord=v4GetRecord_(name,row,sid);
    v4Assert_(!!updatedRecord,'Updated record cannot be read.');
    v4Assert_(String(updatedRecord[updateHeader])===String(updateValue),
      'Updated value was not persisted.');

    // Fresh session persistence test. This proves data is not merely held
    // in the original request/session/cache.
    testSid=v4NewTestSession_(sid);
    var freshRecord=v4GetRecord_(name,row,testSid);
    v4Assert_(!!freshRecord,'Record is not visible from a fresh session.');
    v4Assert_(String(freshRecord[updateHeader])===String(updateValue),
      'Updated value did not persist across a fresh session.');
    item.persistence=true;

    // Isolation test 1: deliberately pass an alternate workspace ID as an
    // extra argument to CRUD/read calls. The generated API must ignore it
    // because workspace resolution is server-side and comes from CONFIG.
    var alternateWorkspaceId='ATTACKER_WORKSPACE_'+Utilities.getUuid();
    var isolatedRead=getClientModuleData(name,testSid,alternateWorkspaceId);
    v4Assert_(isolatedRead&&isolatedRead.success,
      'Isolation probe read failed unexpectedly.');
    var isolatedRecord=v4GetRecord_(name,row,testSid);
    v4Assert_(!!isolatedRecord,'Isolation probe lost client record.');

    // Isolation test 2: another module must not expose this module's test
    // marker. This catches accidental broad-sheet reads/module routing.
    var otherModules=Object.keys(MODULES).filter(function(n){return n!==name;});
    var leaked=false;
    otherModules.forEach(function(other){
      var otherRead=getClientModuleData(other,testSid);
      if(!otherRead||!otherRead.success)return;
      otherRead.records.forEach(function(r){
        Object.keys(r).forEach(function(k){
          if(k!=='_row'&&v4ContainsMarker_(r[k],marker))leaked=true;
        });
      });
    });
    v4Assert_(!leaked,'Test marker leaked into another module.');

    // Dashboard isolation: dashboard must report this exact workspace and
    // must not expose the module smoke marker in financial/recent data.
    var dash=getClientDashboardData(testSid);
    v4Assert_(dash&&dash.success,'Dashboard read failed during isolation test.');
    v4Assert_(dash.dashboard&&String(dash.dashboard.workspaceId)===configuredWorkspaceId,
      'Dashboard workspace does not match CLIENT_CONFIG.');
    var recent=dash.dashboard.recentTransactions||[];
    recent.forEach(function(r){
      Object.keys(r).forEach(function(k){
        if(v4ContainsMarker_(r[k],marker))leaked=true;
      });
    });
    v4Assert_(!leaked,'Smoke-test marker leaked into dashboard data.');
    item.isolation=true;

    // Delete and verify the row is actually gone.
    var deleted=deleteClientModuleRecord(name,row,testSid);
    v4Assert_(deleted&&deleted.success,deleted&&deleted.message||'Delete failed');
    item.delete=true;
    v4Assert_(!v4RecordExists_(name,row,testSid),
      'Deleted record is still readable.');

    item.cleanup=true;
    item.pass=item.create&&item.read&&item.update&&item.persistence&&
      item.isolation&&item.delete&&item.cleanup;
  }catch(e){
    item.error=e&&e.message?e.message:String(e);

    // Best-effort cleanup if a test failed after creating a row.
    if(row){
      try{deleteClientModuleRecord(name,row,testSid||sid);}catch(ignore1){}
      try{if(testSid&&testSid!==sid)CacheService.getScriptCache().remove(String(testSid));}catch(ignore2){}
    }
  }

  item.durationMs=Date.now()-started;
  return item;
}

// ============================================================
// MODULES 7-12 COMPLETE SUITE
// ============================================================
function runModules7To12SmokeTest(sid){
  var names=['Tax','Agro','Productivity','POS','Attendance','Warehouse'];
  var results=[];
  var started=Date.now();

  try{
    getSession_(sid);
  }catch(e){
    return {
      success:false,
      testVersion:'V4.1',
      error:e&&e.message?e.message:String(e),
      modules:[],
      runAt:new Date().toISOString(),
      durationMs:Date.now()-started
    };
  }

  names.forEach(function(name){
    results.push(v4TestModule_(name,sid));
  });

  return {
    success:results.length===names.length&&results.every(function(x){return x.pass;}),
    testVersion:'V4.1',
    scope:names,
    checks:[
      'create',
      'read',
      'update',
      'persistence',
      'isolation',
      'delete',
      'cleanup',
      'dashboard_workspace_isolation'
    ],
    modules:results,
    passedModules:results.filter(function(x){return x.pass;}).map(function(x){return x.module;}),
    failedModules:results.filter(function(x){return !x.pass;}).map(function(x){return x.module;}),
    runAt:new Date().toISOString(),
    durationMs:Date.now()-started
  };
}

// Alias with an explicit name for operators/administrators.
function runModules7To12IsolationPersistenceTest(sid){
  return runModules7To12SmokeTest(sid);
}
`;

  return pkg;
}
