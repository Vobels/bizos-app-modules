// ============================================================
// ClientDeploymentV4.js - MODULE 7-12 SMOKE TEST ADAPTER
// ============================================================
// Adds a server-side smoke test to the isolated V3 client package.
// Modules tested: Tax, Agro, Productivity, POS, Attendance, Warehouse.
// The test creates temporary records, verifies read/update/delete,
// then removes them. It never accepts a client-supplied workspace ID.
// ============================================================
function generateClientCodeSafelyV4(settings){
  var pkg=generateClientCodeSafelyV3(settings);
  if(!pkg||!pkg.files)throw new Error('Client deployment package is empty.');
  var target=null;
  pkg.files.forEach(function(f){if(f.name==='Code')target=f;});
  if(!target)throw new Error('Generated client Code file not found.');
  target.source += `\n\nfunction runModules7To12SmokeTest(sid){\n  var names=['Tax','Agro','Productivity','POS','Attendance','Warehouse'];\n  var results=[];\n  getSession_(sid);\n  names.forEach(function(name){\n    var item={module:name,create:false,read:false,update:false,delete:false,error:''};\n    var row=null;\n    try{\n      var m=module_(name),ss=getWorkspace_(),sh=ss.getSheetByName(m.sheet);\n      if(!sh)throw new Error('Module sheet missing: '+m.sheet);\n      var headers=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0];\n      var record={};\n      headers.forEach(function(h){if(h)record[h]='SMOKE_TEST_'+Date.now();});\n      if(!Object.keys(record).length)throw new Error('Module has no headers.');\n      var created=createClientModuleRecord(name,record,sid);\n      if(!created.success)throw new Error(created.message||'Create failed');\n      item.create=true;\n      row=created.record&&created.record._row;\n      var read=getClientModuleData(name,sid);\n      if(!read.success)throw new Error(read.message||'Read failed');\n      item.read=read.records.some(function(r){return Number(r._row)===Number(row);});\n      if(!item.read)throw new Error('Created record was not readable.');\n      var patch={};\n      if(headers[0])patch[headers[0]]='SMOKE_UPDATED_'+Date.now();\n      var updated=updateClientModuleRecord(name,row,patch,sid);\n      if(!updated.success)throw new Error(updated.message||'Update failed');\n      item.update=true;\n      var after=getClientModuleData(name,sid);\n      var found=after.records.filter(function(r){return Number(r._row)===Number(row);})[0];\n      if(headers[0]&&(!found||String(found[headers[0]])!==String(patch[headers[0]])))throw new Error('Updated record was not persisted.');\n      var deleted=deleteClientModuleRecord(name,row,sid);\n      if(!deleted.success)throw new Error(deleted.message||'Delete failed');\n      item.delete=true;\n    }catch(e){\n      item.error=e.message||String(e);\n      if(row){try{deleteClientModuleRecord(name,row,sid);}catch(ignore){}}\n    }\n    item.pass=item.create&&item.read&&item.update&&item.delete;\n    results.push(item);\n  });\n  return{success:results.every(function(x){return x.pass;}),modules:results,runAt:new Date().toISOString()};\n}\n`;
  return pkg;
}
