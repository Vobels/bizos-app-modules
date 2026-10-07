// ============================================================
// ProvisioningPreflight.js
// =================================================
// Builds and validates the V12 client package without creating a
// workspace or Apps Script deployment.
// ============================================================
function runProvisioningPreflight(){
  var settings={clientId:'CLIENT_PREFLIGHT_0001',clientName:'BizOS Preflight Test',primaryColor:'#2E7D32',logoUrl:'',email:'preflight@example.invalid',sheetId:'WORKSPACE_PREFLIGHT_0001',businessId:'BUSINESS_PREFLIGHT_0001',masterApiUrl:'https://example.invalid/master-api'};
  try{
    var pkg=generateClientCodeSafelyV12(settings);
    var result=validateClientDeploymentPackageV12_(pkg,settings);
    var cf=pkg&&pkg.files?pkg.files.find(function(f){return f&&f.name==='Code';}):null;
    var shell=pkg&&pkg.files?pkg.files.find(function(f){return f&&f.name==='ClientShellV12';}):null;
    var code=cf&&typeof cf.source==='string'?cf.source:'';
    var shellCode=shell&&typeof shell.source==='string'?shell.source:'';
    var checks={
      module_registry:code.indexOf('var MODULES=')!==-1,
      provenance:code.indexOf('CLIENT_V12_PACKAGE_VERSION_')!==-1&&code.indexOf('CLIENT_V12_PACKAGE_PROVENANCE_')!==-1,
      session_bridge:code.indexOf('function v12PersistSession_')!==-1&&code.indexOf('function v12ValidateWorkspaceBinding_')!==-1,
      client_records:code.indexOf('function getClientV12ModuleSmokeAudit')!==-1||shellCode.indexOf('client-records-view')!==-1,
      staff_runtime:code.indexOf('function getBusinessStaff(bid,sid)')!==-1&&code.indexOf('function assignModuleToStaff(email,bid,moduleName,assignedBy,sid)')!==-1,
      internal_helpers_hidden:code.indexOf('function getOrCreateUserSheet(')===-1&&code.indexOf('function getOrCreateBusinessSheet(')===-1
    };
    var checksValid=Object.keys(checks).every(function(k){return checks[k]===true;});
    var out={success:!!(result&&result.success&&checksValid),stage:'v12_generation_and_validation',result:result,checks:checks,checksValid:checksValid,packageVersion:'12.2.1',generatedFiles:pkg&&pkg.files?pkg.files.map(function(f){return f.name;}):[],runAt:new Date().toISOString()};
    console.log('PROVISIONING PREFLIGHT V12:',JSON.stringify(out));Logger.log(JSON.stringify(out));return out;
  }catch(e){
    var fail={success:false,stage:'v12_generation',code:'PREFLIGHT_V12_FAILED',message:e&&e.message?String(e.message):String(e),stack:e&&e.stack?String(e.stack):'',runAt:new Date().toISOString()};
    console.error('PROVISIONING PREFLIGHT V12 FAILED:',JSON.stringify(fail));Logger.log(JSON.stringify(fail));return fail;
  }
}
