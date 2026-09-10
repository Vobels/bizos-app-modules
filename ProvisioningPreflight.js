// ============================================================
// ProvisioningPreflight.js
// ============================================================
// Builds the complete V7 client package with synthetic identifiers
// and validates it without creating a workspace or Apps Script project.
// Safe to run before any real provisioning test.
// ============================================================

function runProvisioningPreflight() {
  var settings = {
    clientId:'CLIENT_PREFLIGHT_0001',
    clientName:'BizOS Preflight Test',
    primaryColor:'#2E7D32',
    logoUrl:'',
    email:'preflight@example.invalid',
    sheetId:'WORKSPACE_PREFLIGHT_0001',
    businessId:'BUSINESS_PREFLIGHT_0001',
    masterApiUrl:'https://example.invalid/master-api'
  };

  try {
    var pkg = generateClientCodeSafelyV7(settings);
    var result = validateClientDeploymentPackage_(pkg,settings);
    var codeFile = pkg && pkg.files ? pkg.files.find(function(f){ return f && f.name === 'Code'; }) : null;
    var dashboardFile = pkg && pkg.files ? pkg.files.find(function(f){ return f && (f.name === 'client-dashboard' || f.name === 'client-dashboard.html'); }) : null;
    var code = codeFile && typeof codeFile.source === 'string' ? codeFile.source : '';
    var dashboard = dashboardFile && typeof dashboardFile.source === 'string' ? dashboardFile.source : '';
    var teamChecks = {
      team_server_functions: code.indexOf('function inviteClientTeamMember') !== -1 && code.indexOf('function getClientTeam') !== -1,
      team_authentication: code.indexOf('function authenticateClientTeamMember') !== -1,
      team_ui: dashboard.indexOf('function showTeam()') !== -1 && dashboard.indexOf('id="inviteTeamBtn"') !== -1,
      staff_profile_ui: dashboard.indexOf('function editTeamMember') !== -1 && dashboard.indexOf('function loadOwnProfile') !== -1,
      master_dashboard_not_exposed: dashboard.indexOf('getAllUsers') === -1 && dashboard.indexOf('getAdminStats') === -1
    };
    var teamValid = Object.keys(teamChecks).every(function(k){ return teamChecks[k] === true; });
    var output = {
      success:!!(result && result.success && teamValid),
      stage:'generation_and_validation',
      result:result,
      teamChecks:teamChecks,
      teamValid:teamValid,
      generatedFiles:pkg&&pkg.files?pkg.files.map(function(f){return f.name;}):[],
      runAt:new Date().toISOString()
    };
    console.log('PROVISIONING PREFLIGHT:', JSON.stringify(output));
    Logger.log(JSON.stringify(output));
    return output;
  } catch(error) {
    var failure = {
      success:false,
      stage:'generation',
      code:'PREFLIGHT_FAILED',
      message:error&&error.message?error.message:String(error),
      stack:error&&error.stack?String(error.stack):'',
      runAt:new Date().toISOString()
    };
    console.error('PROVISIONING PREFLIGHT FAILED:', JSON.stringify(failure));
    Logger.log(JSON.stringify(failure));
    return failure;
  }
}
