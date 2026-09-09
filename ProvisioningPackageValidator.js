// ============================================================
// ProvisioningPackageValidator.js
// ============================================================
// Structural gate between code generation and deployment.
// This does not execute the generated client. It rejects packages
// that are missing required files/functions or contain known stale
// master/client entry-point patterns.
// ============================================================

function validateClientDeploymentPackage_(pkg, expected) {
  try {
    expected = expected || {};
    if (!pkg || !Array.isArray(pkg.files) || !pkg.files.length) {
      return {success:false,code:'PACKAGE_EMPTY',message:'Generated client package is empty.'};
    }

    var byName = {};
    pkg.files.forEach(function(f){
      if (f && f.name) byName[String(f.name)] = f;
    });

    var required = ['Code','client-landing','client-dashboard','appsscript'];
    var missing = required.filter(function(name){return !byName[name];});
    if (missing.length) return {success:false,code:'PACKAGE_MISSING_FILES',message:'Generated package is missing: '+missing.join(', ')};

    var code = String(byName.Code.source || byName.Code.content || '');
    var landing = String(byName['client-landing'].source || byName['client-landing'].content || '');
    var dashboard = String(byName['client-dashboard'].source || byName['client-dashboard'].content || '');
    var manifest = String(byName.appsscript.source || byName.appsscript.content || '');

    var requiredCode = [
      'function doGet(e)',
      'function getWorkspace_()',
      'function getSession_(sid)',
      'function authenticateClient(email,password)',
      'function validateClientSession(sid)',
      'function getClientModuleSchema(moduleName,sid)',
      'function getClientModuleData(moduleName,sid)',
      'function createClientModuleRecord(moduleName,record,sid)',
      'function updateClientModuleRecord(moduleName,rowIndex,record,sid)',
      'function deleteClientModuleRecord(moduleName,rowIndex,sid)',
      'function getClientDashboardData(sid)'
    ];
    var missingCode = requiredCode.filter(function(token){return code.indexOf(token)<0;});
    if(missingCode.length) return {success:false,code:'PACKAGE_MISSING_FUNCTIONS',message:'Generated client Code is missing: '+missingCode.join(', ')};

    if (code.indexOf('getClientSettings()') >= 0) {
      return {success:false,code:'STALE_CLIENT_SETTINGS',message:'Generated client Code contains obsolete getClientSettings().'};
    }
    if (code.indexOf('createFallbackLandingPage') >= 0) {
      return {success:false,code:'FALLBACK_DEPLOYMENT_PATH',message:'Generated client package contains a forbidden fallback landing path.'};
    }

    // Module definitions live in the generated client dashboard package,
    // not necessarily in the server Code file. Validate the whole generated
    // client package so V6's dashboard module registry is correctly checked.
    var moduleSource = code + '\n' + dashboard;
    var modules = ['Finance','Ecommerce','Sales','CRM','HR','Logistics','Tax','Agro','Productivity','POS','Attendance','Warehouse'];
    var missingModules = modules.filter(function(name){return moduleSource.indexOf(name+':')<0;});
    if (missingModules.length) return {success:false,code:'PACKAGE_MODULES_MISSING',message:'Generated client package is missing module definitions: '+missingModules.join(', ')};

    if (landing.indexOf('authenticateClient') < 0) return {success:false,code:'LANDING_AUTH_MISSING',message:'Client landing page is not wired to authenticateClient.'};
    if (dashboard.indexOf('getClientDashboardData') < 0) return {success:false,code:'DASHBOARD_API_MISSING',message:'Client dashboard is not wired to the client dashboard API.'};
    if (manifest.indexOf('USER_DEPLOYING') < 0) return {success:false,code:'DEPLOYMENT_IDENTITY_INVALID',message:'Client deployment manifest is not configured to execute as the Vobels deployer.'};
    if (manifest.indexOf('ANYONE_ANONYMOUS') < 0) return {success:false,code:'DEPLOYMENT_ACCESS_INVALID',message:'Client deployment manifest is not configured for the intended anonymous web entry point.'};

    if (expected.clientId && code.indexOf('"clientId":"'+String(expected.clientId).replace(/"/g,'\\"')+'"') < 0) {
      return {success:false,code:'PACKAGE_CLIENT_ID_MISMATCH',message:'Generated package does not contain the expected client ID.'};
    }
    if (expected.sheetId && code.indexOf('"sheetId":"'+String(expected.sheetId).replace(/"/g,'\\"')+'"') < 0) {
      return {success:false,code:'PACKAGE_WORKSPACE_MISMATCH',message:'Generated package does not contain the expected workspace ID.'};
    }

    return {
      success:true,
      code:'PACKAGE_VALID',
      fileCount:pkg.files.length,
      files:pkg.files.map(function(f){return f.name;}),
      modules:modules,
      checks:['required_files','required_functions','module_definitions','no_stale_client_settings','no_fallback_deployment','landing_auth','dashboard_api','deployment_identity','deployment_access','client_binding','workspace_binding']
    };
  } catch(error) {
    return {success:false,code:'PACKAGE_VALIDATION_ERROR',message:error&&error.message?error.message:String(error)};
  }
}
