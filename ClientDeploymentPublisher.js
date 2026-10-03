// ============================================================
// ClientDeploymentPublisher.js - AUTHORITATIVE SAFE PUBLISHER
// ============================================================
// Explicit function name. Does not override legacy publishers.
// Uploads the exact validated generated package and never grants the
// customer editor access to the Apps Script project.
// ============================================================

function createAndDeployClientScriptSafe(email, clientId, code, businessName, primaryColor, logoUrl) {
  var scriptId = '';
  try {
    if (!code || !Array.isArray(code.files) || !code.files.length) {
      return {success:false,message:'Generated client package is empty.',code:'PACKAGE_EMPTY'};
    }

    var files = code.files.map(function(f){
      return {name:String(f.name),type:String(f.type),source:String(f.source || f.content || '')};
    });

    var required=['Code','ClientShellV12','appsscript'];
    var names=files.map(function(f){return f.name;});
    var missing=required.filter(function(n){return names.indexOf(n)<0;});
    if(missing.length) return {success:false,message:'Generated client package is missing: '+missing.join(', '),code:'PACKAGE_MISSING_FILES'};

    var createResponse=UrlFetchApp.fetch('https://script.googleapis.com/v1/projects',{
      method:'POST',
      headers:{'Authorization':'Bearer '+ScriptApp.getOAuthToken(),'Content-Type':'application/json'},
      payload:JSON.stringify({title:businessName+' - BizOS'}),
      muteHttpExceptions:true
    });
    if(createResponse.getResponseCode()!==200){
      return {success:false,message:'Failed to create client Apps Script project: '+createResponse.getContentText(),code:'PROJECT_CREATE_FAILED'};
    }

    var createResult=JSON.parse(createResponse.getContentText()||'{}');
    scriptId=createResult.scriptId||'';
    if(!scriptId) throw new Error('Apps Script project creation returned no script ID.');

    var updateUrl='https://script.googleapis.com/v1/projects/'+encodeURIComponent(scriptId)+'/content';
    var updateResponse=UrlFetchApp.fetch(updateUrl,{
      method:'PUT',
      headers:{'Authorization':'Bearer '+ScriptApp.getOAuthToken(),'Content-Type':'application/json'},
      payload:JSON.stringify({files:files}),
      muteHttpExceptions:true
    });
    if(updateResponse.getResponseCode()!==200){
      cleanupPublishedProjectSafe_(scriptId);
      return {success:false,message:'Failed to upload generated client package: '+updateResponse.getContentText(),code:'PACKAGE_UPLOAD_FAILED',scriptId:scriptId};
    }

    var versionUrl='https://script.googleapis.com/v1/projects/'+encodeURIComponent(scriptId)+'/versions';
    var versionResponse=UrlFetchApp.fetch(versionUrl,{
      method:'POST',
      headers:{'Authorization':'Bearer '+ScriptApp.getOAuthToken(),'Content-Type':'application/json'},
      payload:JSON.stringify({description:'BizOS client '+clientId+' - V12 '+CLIENT_V12_PACKAGE_VERSION_}),
      muteHttpExceptions:true
    });
    if(versionResponse.getResponseCode()!==200){
      cleanupPublishedProjectSafe_(scriptId);
      return {success:false,message:'Failed to create client version: '+versionResponse.getContentText(),code:'VERSION_CREATE_FAILED',scriptId:scriptId};
    }

    var versionResult=JSON.parse(versionResponse.getContentText()||'{}');
    var versionNumber=versionResult.versionNumber;
    if(!versionNumber) throw new Error('Version creation returned no version number.');

    var deployUrl='https://script.googleapis.com/v1/projects/'+encodeURIComponent(scriptId)+'/deployments';
    var deployResponse=UrlFetchApp.fetch(deployUrl,{
      method:'POST',
      headers:{'Authorization':'Bearer '+ScriptApp.getOAuthToken(),'Content-Type':'application/json'},
      payload:JSON.stringify({versionNumber:versionNumber,manifestFileName:'appsscript',description:'BizOS client web app - '+businessName+' - V12 '+CLIENT_V12_PACKAGE_VERSION_}),
      muteHttpExceptions:true
    });
    if(deployResponse.getResponseCode()!==200){
      cleanupPublishedProjectSafe_(scriptId);
      return {success:false,message:'Failed to deploy client web app: '+deployResponse.getContentText(),code:'DEPLOYMENT_FAILED',scriptId:scriptId};
    }

    var deployResult=JSON.parse(deployResponse.getContentText()||'{}');
    var deploymentId=deployResult.deploymentId||'';
    if(!deploymentId) throw new Error('Deployment returned no deployment ID.');

    var webAppUrl='https://script.google.com/macros/s/'+deploymentId+'/exec';

    // V12 publish proof: do not assume that uploading files/versioning succeeded
    // means the public deployment is actually pointing at that exact version.
    // Verify the deployment resource and the versioned project content before
    // returning success to provisioning/admin flows.
    var publishVerification=verifyPublishedV12Deployment_(scriptId,deploymentId,clientId,versionNumber,files);
    if(!publishVerification || publishVerification.success!==true){
      cleanupPublishedProjectSafe_(scriptId);
      return {
        success:false,
        message:publishVerification&&publishVerification.message?publishVerification.message:'Published V12 deployment could not be verified.',
        code:'DEPLOYMENT_VERIFICATION_FAILED',
        scriptId:scriptId,
        deploymentId:deploymentId,
        versionNumber:versionNumber,
        verification:publishVerification||null
      };
    }

    // Intentionally NO DriveApp.getFileById(scriptId).addEditor(email).
    // Customer receives only the deployed web application.
    return {
      success:true,
      scriptId:scriptId,
      deploymentId:deploymentId,
      versionNumber:versionNumber,
      webAppUrl:webAppUrl,
      editUrl:'https://script.google.com/d/'+scriptId+'/edit',
      packageVersion:String(code.version||CLIENT_V12_PACKAGE_VERSION_),
      packageRelease:String(code.release||CLIENT_V12_PACKAGE_RELEASE_),
      packageProvenance:String(code.provenance||CLIENT_V12_PACKAGE_PROVENANCE_),
      publishVerification:publishVerification,
      message:'Client V12 package deployed and verified'
    };
  } catch(error) {
    if(scriptId) cleanupPublishedProjectSafe_(scriptId);
    return {success:false,message:error&&error.message?error.message:'Client deployment failed.',code:'DEPLOYMENT_EXCEPTION',scriptId:scriptId||''};
  }
}

function cleanupPublishedProjectSafe_(scriptId) {
  try {
    if(!scriptId) return;
    UrlFetchApp.fetch('https://script.googleapis.com/v1/projects/'+encodeURIComponent(scriptId),{
      method:'DELETE',
      headers:{'Authorization':'Bearer '+ScriptApp.getOAuthToken()},
      muteHttpExceptions:true
    });
  } catch(error) { console.error('Published project cleanup failed:',error); }
}


/**
 * Verifies that the deployment ID returned by the Apps Script API points to
 * the exact version created from the generated V12 package, and that the
 * versioned project content contains the V12 runtime markers. This is a
 * publisher-side proof; it does not execute the public web app.
 */
function verifyPublishedV12Deployment_(scriptId,deploymentId,clientId,expectedVersionNumber,files){
  try{
    var auth={Authorization:'Bearer '+ScriptApp.getOAuthToken()};
    var deploymentUrl='https://script.googleapis.com/v1/projects/'+encodeURIComponent(scriptId)+'/deployments/'+encodeURIComponent(deploymentId);
    var dr=UrlFetchApp.fetch(deploymentUrl,{method:'GET',headers:auth,muteHttpExceptions:true});
    if(dr.getResponseCode()<200||dr.getResponseCode()>=300)return{success:false,stage:'deployment_lookup',message:'Unable to verify the created deployment: '+dr.getContentText()};
    var d=JSON.parse(dr.getContentText()||'{}'),cfg=d.deploymentConfig||{},actualVersion=String(cfg.versionNumber||d.versionNumber||'');
    if(actualVersion!==String(expectedVersionNumber))return{success:false,stage:'deployment_version',message:'Deployment '+deploymentId+' is not pointing to the generated version '+expectedVersionNumber+'.',expectedVersion:String(expectedVersionNumber),actualVersion:actualVersion};

    var contentUrl='https://script.googleapis.com/v1/projects/'+encodeURIComponent(scriptId)+'/content?versionNumber='+encodeURIComponent(String(expectedVersionNumber));
    var cr=UrlFetchApp.fetch(contentUrl,{method:'GET',headers:auth,muteHttpExceptions:true});
    if(cr.getResponseCode()<200||cr.getResponseCode()>=300)return{success:false,stage:'version_content_lookup',message:'Unable to verify the content of deployed V12 version: '+cr.getContentText(),versionNumber:String(expectedVersionNumber)};
    var body=JSON.parse(cr.getContentText()||'{}'),publishedFiles=Array.isArray(body.files)?body.files:[],codeFile=null;
    for(var i=0;i<publishedFiles.length;i++){if(String(publishedFiles[i].name||'')==='Code'){codeFile=publishedFiles[i];break;}}
    if(!codeFile)return{success:false,stage:'version_content',message:'The deployed version does not contain the Code file.',versionNumber:String(expectedVersionNumber)};
    var src=String(codeFile.source||'');
    var markers={
      clientId:src.indexOf('clientId:'+JSON.stringify(String(clientId)))>=0,
      packageVersion:src.indexOf('CLIENT_V12_PACKAGE_VERSION_')>=0&&src.indexOf(String(CLIENT_V12_PACKAGE_VERSION_))>=0,
      modules:src.indexOf('var MODULES=')>=0,
      v12Dashboard:src.indexOf('function getClientDashboardData')>=0,
      packageInfo:src.indexOf('function getClientPackageInfo')>=0
    };
    var missing=[];Object.keys(markers).forEach(function(k){if(!markers[k])missing.push(k);});
    if(missing.length)return{success:false,stage:'runtime_markers',message:'The deployed version is missing required V12 runtime markers: '+missing.join(', '),versionNumber:String(expectedVersionNumber),markers:markers};

    var expectedNames=(files||[]).map(function(f){return String(f.name||'');}).sort(),actualNames=publishedFiles.map(function(f){return String(f.name||'');}).sort();
    if(JSON.stringify(expectedNames)!==JSON.stringify(actualNames))return{success:false,stage:'file_set',message:'The deployed version file set differs from the generated V12 package.',versionNumber:String(expectedVersionNumber),expectedFiles:expectedNames,actualFiles:actualNames};

    return{success:true,scriptId:String(scriptId),deploymentId:String(deploymentId),versionNumber:String(expectedVersionNumber),webAppUrl:'https://script.google.com/macros/s/'+deploymentId+'/exec',markers:markers,fileCount:publishedFiles.length,verifiedAt:new Date().toISOString(),runtimeExecuted:false};
  }catch(error){return{success:false,stage:'exception',message:error&&error.message?error.message:String(error)};}
}
