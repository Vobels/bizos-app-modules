// ============================================================
// ProvisioningPreflight.js
// ============================================================
// Builds and validates the V7 client package without creating a
// workspace or Apps Script deployment.
// ============================================================
function runProvisioningPreflight(){
  var settings={clientId:'CLIENT_PREFLIGHT_0001',clientName:'BizOS Preflight Test',primaryColor:'#2E7D32',logoUrl:'',email:'preflight@example.invalid',sheetId:'WORKSPACE_PREFLIGHT_0001',businessId:'BUSINESS_PREFLIGHT_0001',masterApiUrl:'https://example.invalid/master-api'};
  try{
    var pkg=generateClientCodeSafelyV7(settings),result=validateClientDeploymentPackage_(pkg,settings);
    var cf=pkg&&pkg.files?pkg.files.find(function(f){return f&&f.name==='Code';}):null,df=pkg&&pkg.files?pkg.files.find(function(f){return f&&(f.name==='client-dashboard'||f.name==='client-dashboard.html');}):null;
    var code=cf&&typeof cf.source==='string'?cf.source:'',dash=df&&typeof df.source==='string'?df.source:'';
    var teamChecks={team_server_functions:code.indexOf('function inviteClientTeamMember')!==-1&&code.indexOf('function getClientTeam')!==-1,team_authentication:code.indexOf('function authenticateClientTeamMemberV7_')!==-1,team_ui:dash.indexOf('function showTeam()')!==-1&&dash.indexOf('id="inviteTeamBtn"')!==-1,staff_profile_ui:dash.indexOf('function editTeamMemberV7')!==-1&&dash.indexOf('function loadOwnProfileV7')!==-1,master_dashboard_not_exposed:dash.indexOf('getAllUsers')===-1&&dash.indexOf('getAdminStats')===-1};
    var teamValid=Object.keys(teamChecks).every(function(k){return teamChecks[k]===true;});
    var out={success:!!(result&&result.success&&teamValid),stage:'generation_and_validation',result:result,teamChecks:teamChecks,teamValid:teamValid,generatedFiles:pkg&&pkg.files?pkg.files.map(function(f){return f.name;}):[],runAt:new Date().toISOString()};
    console.log('PROVISIONING PREFLIGHT:',JSON.stringify(out));Logger.log(JSON.stringify(out));return out;
  }catch(e){var fail={success:false,stage:'generation',code:'PREFLIGHT_FAILED',message:e&&e.message?e.message:String(e),stack:e&&e.stack?String(e.stack):'',runAt:new Date().toISOString()};console.error('PROVISIONING PREFLIGHT FAILED:',JSON.stringify(fail));Logger.log(JSON.stringify(fail));return fail;}
}
