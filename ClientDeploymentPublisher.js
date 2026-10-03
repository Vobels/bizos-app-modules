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
      payload:JSON.stringify({description:'BizOS client deployment '+clientId}),
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
      payload:JSON.stringify({versionNumber:versionNumber,manifestFileName:'appsscript',description:'BizOS client web app - '+businessName}),
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

    // Intentionally NO DriveApp.getFileById(scriptId).addEditor(email).
    // Customer receives only the deployed web application.
    return {success:true,scriptId:scriptId,deploymentId:deploymentId,webAppUrl:webAppUrl,editUrl:'https://script.google.com/d/'+scriptId+'/edit',message:'Client script deployed as web app'};
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
