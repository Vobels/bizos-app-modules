// ============================================================
// ProvisioningPreflight.js
// ============================================================
// Builds the complete V6 client package with synthetic identifiers
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
    var pkg = generateClientCodeSafelyV6(settings);
    var result = validateClientDeploymentPackage_(pkg,settings);
    var output = {
      success:!!(result && result.success),
      stage:'generation_and_validation',
      result:result,
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
