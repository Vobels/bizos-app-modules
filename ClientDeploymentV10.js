// ============================================================
// ClientDeploymentV10.js - CLIENT PACKAGE MANIFEST HARDENING
// ============================================================
// Keeps V9 behavior unchanged and only ensures the generated client
// manifest explicitly permits invitation emails through MailApp.
// ============================================================
function generateClientCodeSafelyV10(settings){
  var base=generateClientCodeSafelyV9(settings);
  if(!base||!base.files)throw new Error('Client deployment package is empty.');
  base.files=base.files.map(function(file){
    if(!file||file.name!=='appsscript')return file;
    var manifest={};
    try{manifest=JSON.parse(String(file.source||file.content||'{}'));}catch(e){throw new Error('Client appsscript manifest is invalid: '+e.message);}
    manifest.oauthScopes=manifest.oauthScopes||[];
    var mailScope='https://www.googleapis.com/auth/script.send_mail';
    if(manifest.oauthScopes.indexOf(mailScope)<0)manifest.oauthScopes.push(mailScope);
    file.source=JSON.stringify(manifest);
    if(typeof file.content==='string')file.content=file.source;
    return file;
  });
  return base;
}
