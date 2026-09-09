// ============================================================
// SecretRotation.js - ONE-TIME MASTER SECRET ROTATION
// ============================================================
// Run rotateMasterSecretAndRemoveTestValue() once from the master
// Apps Script editor before production. It stores a random secret in
// Script Properties and replaces the exposed test literal in the
// Master Backend source while preserving the rest of the project files.
// ============================================================
function rotateMasterSecretAndRemoveTestValue(){
  var secret=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');
  PropertiesService.getScriptProperties().setProperty('MASTER_API_SECRET',secret);
  var scriptId=ScriptApp.getScriptId();
  var token=ScriptApp.getOAuthToken();
  var getUrl='https://script.googleapis.com/v1/projects/'+encodeURIComponent(scriptId)+'/content';
  var getResponse=UrlFetchApp.fetch(getUrl,{headers:{Authorization:'Bearer '+token},muteHttpExceptions:true});
  if(getResponse.getResponseCode()<200||getResponse.getResponseCode()>=300)throw new Error('Could not read current Apps Script project content: '+getResponse.getContentText());
  var current=JSON.parse(getResponse.getContentText()),files=current.files||[],changed=false;
  var updated=files.map(function(file){
    var source=String(file.source||'');
    if(source.indexOf("const MASTER_SECRET = 'TEST_SECRET_2024';")>=0){
      source=source.replace("const MASTER_SECRET = 'TEST_SECRET_2024';","const MASTER_SECRET = PropertiesService.getScriptProperties().getProperty('MASTER_API_SECRET') || '';" );
      changed=true;
    }
    return {name:file.name,type:file.type,source:source};
  });
  if(!changed)throw new Error('The exposed TEST_SECRET_2024 declaration was not found. No source change was made.');
  var putResponse=UrlFetchApp.fetch(getUrl,{method:'put',contentType:'application/json',headers:{Authorization:'Bearer '+token},muteHttpExceptions:true,payload:JSON.stringify({files:updated})});
  if(putResponse.getResponseCode()<200||putResponse.getResponseCode()>=300)throw new Error('Could not update Apps Script project content: '+putResponse.getContentText());
  return {success:true,message:'Master secret rotated and test secret removed from source. Create a new version/deployment of the master API if the production deployment does not automatically use HEAD.',secretStoredInScriptProperties:true};
}
