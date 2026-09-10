// ============================================================
// ClientDeploymentV7.js - CLIENT TEAM + STAFF PROFILE
// ============================================================
// V7 wraps the proven V6 client package. It adds a client-scoped
// Team screen and staff authentication without exposing the master
// BizOS user/admin system.
// ============================================================

function generateClientCodeSafelyV7(settings) {
  var base = generateClientCodeSafelyV6(settings);
  if (!base || !base.files || !base.files.length) {
    throw new Error('Client deployment package is empty.');
  }

  base.files = base.files.map(function(file) {
    if (!file) return file;

    if (file.name === 'Code' && typeof file.source === 'string') {
      file.source += buildClientTeamServerV7_();
    }

    if (file.name === 'client-dashboard' || file.name === 'client-dashboard.html') {
      var html = typeof file.source === 'string' ? file.source :
                 (typeof file.content === 'string' ? file.content : '');
      html = injectClientTeamUiV7_(html, settings);
      file.source = html;
      delete file.content;
    }

    return file;
  });

  return base;
}

function buildClientTeamServerV7_() {
  return `
function getClientTeamSheetV7_(){
  var ss=SpreadsheetApp.openById(CLIENT_CONFIG.sheetId);
  var sh=ss.getSheetByName('Client_Team');
  if(!sh){
    sh=ss.insertSheet('Client_Team');
    sh.appendRow(['User_ID','Email','Name','Phone','Role','Department','Status','Password_Hash','Must_Change_Password','Invited_By','Created_At','Updated_At','Last_Login']);
  }
  return sh;
}
function getClientTeamRowsV7_(){
  var sh=getClientTeamSheetV7_(),v=sh.getDataRange().getValues();
  if(v.length<2)return[];
  var h=v[0];
  return v.slice(1).map(function(r,i){var o={row:i+2};h.forEach(function(k,j){if(k)o[String(k)]=r[j];});return o;});
}
function hashClientTeamPasswordV7_(p){return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(p||'')));}
function clientTeamManagerV7_(session){
  var role=String(session&&session.user&&session.user.role||'').toLowerCase();
  return role==='owner'||role==='admin';
}
function escapeClientTeamV7_(v){return String(v==null?'':v).replace(/[&<>\"']/g,function(m){return{'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;',"'":'&#039;'}[m];});}
function getClientTeam(sid){
  try{
    var session=validateClientSession(sid);
    if(!session)return{success:false,message:'Session expired',code:'SESSION_EXPIRED'};
    return{success:true,members:getClientTeamRowsV7_().map(function(r){return{userId:String(r.User_ID||''),email:String(r.Email||''),name:String(r.Name||''),phone:String(r.Phone||''),role:String(r.Role||'staff'),department:String(r.Department||''),status:String(r.Status||'invited'),lastLogin:r.Last_Login||''};})};
  }catch(e){return{success:false,message:e.message,code:'TEAM_ERROR'};}
}
function inviteClientTeamMember(sid,data){
  try{
    var session=validateClientSession(sid);
    if(!session||!clientTeamManagerV7_(session))return{success:false,message:'Only the business owner or an admin can invite team members.',code:'TEAM_FORBIDDEN'};
    data=data||{};
    var email=String(data.email||'').trim().toLowerCase(),name=String(data.name||'').trim();
    if(!email||!name)return{success:false,message:'Full name and email are required.',code:'INVALID_INPUT'};
    if(email===String(session.email||'').toLowerCase())return{success:false,message:'The owner already has access to this business.',code:'OWNER_ACCOUNT'};
    var role=String(data.role||'staff').toLowerCase();if(role!=='admin'&&role!=='staff')role='staff';
    var rows=getClientTeamRowsV7_();
    if(rows.some(function(r){return String(r.Email||'').toLowerCase()===email;}))return{success:false,message:'This email is already on the team.',code:'TEAM_EXISTS'};
    var tempPassword=Utilities.getUuid().replace(/-/g,'').slice(0,10)+'A1!';
    var now=new Date().toISOString(),userId='TEAM_'+Utilities.getUuid().replace(/-/g,'').slice(0,8).toUpperCase();
    getClientTeamSheetV7_().appendRow([userId,email,name,String(data.phone||'').trim(),role,String(data.department||'').trim(),'invited',hashClientTeamPasswordV7_(tempPassword),'YES',session.email||'',now,now,'']);
    var businessName=String(CLIENT_CONFIG.clientName||'BizOS');
    MailApp.sendEmail({to:email,subject:'Invitation to '+businessName+' on BizOS',htmlBody:'<p>Hello <strong>'+escapeClientTeamV7_(name)+'</strong>,</p><p>You have been invited to join <strong>'+escapeClientTeamV7_(businessName)+'</strong> on BizOS.</p><p><strong>Email:</strong> '+escapeClientTeamV7_(email)+'<br><strong>Temporary password:</strong> '+escapeClientTeamV7_(tempPassword)+'</p><p>Open your company BizOS login page and sign in. You can change your password from your profile.</p>'});
    return{success:true,message:'Invitation sent to '+email+'.',userId:userId};
  }catch(e){return{success:false,message:e.message,code:'TEAM_INVITE_ERROR'};}
}
function updateClientTeamMember(sid,userId,data){
  try{
    var session=validateClientSession(sid);
    if(!session||!clientTeamManagerV7_(session))return{success:false,message:'Only the business owner or an admin can update staff profiles.',code:'TEAM_FORBIDDEN'};
    var sh=getClientTeamSheetV7_(),rows=getClientTeamRowsV7_(),target=rows.find(function(r){return String(r.User_ID||'')===String(userId||'');});
    if(!target)return{success:false,message:'Team member not found.',code:'TEAM_NOT_FOUND'};
    data=data||{};
    var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
    function set_(key,value){var i=headers.indexOf(key);if(i>=0)sh.getRange(target.row,i+1).setValue(value);}
    if(data.name!==undefined)set_('Name',String(data.name||'').trim());
    if(data.phone!==undefined)set_('Phone',String(data.phone||'').trim());
    if(data.department!==undefined)set_('Department',String(data.department||'').trim());
    if(data.role!==undefined){var role=String(data.role).toLowerCase();set_('Role',role==='admin'?'admin':'staff');}
    if(data.status!==undefined){var status=String(data.status).toLowerCase();if(['invited','active','suspended'].indexOf(status)<0)status='active';set_('Status',status);}
    set_('Updated_At',new Date().toISOString());
    return{success:true,message:'Staff profile updated successfully.'};
  }catch(e){return{success:false,message:e.message,code:'TEAM_UPDATE_ERROR'};}
}
function removeClientTeamMember(sid,userId){
  try{
    var session=validateClientSession(sid);
    if(!session||!clientTeamManagerV7_(session))return{success:false,message:'Only the business owner or an admin can remove team members.',code:'TEAM_FORBIDDEN'};
    var sh=getClientTeamSheetV7_(),target=getClientTeamRowsV7_().find(function(r){return String(r.User_ID||'')===String(userId||'');});
    if(!target)return{success:false,message:'Team member not found.',code:'TEAM_NOT_FOUND'};
    sh.deleteRow(target.row);return{success:true,message:'Team member removed.'};
  }catch(e){return{success:false,message:e.message,code:'TEAM_REMOVE_ERROR'};}
}
function getClientOwnProfile(sid){
  try{
    var session=validateClientSession(sid);if(!session)return{success:false,message:'Session expired',code:'SESSION_EXPIRED'};
    var role=String(session.user&&session.user.role||'').toLowerCase();
    if(role==='owner')return{success:true,profile:{email:session.email,name:session.user.name||session.email,role:'owner',isOwner:true}};
    var target=getClientTeamRowsV7_().find(function(r){return String(r.Email||'').toLowerCase()===String(session.email||'').toLowerCase();});
    if(!target)return{success:false,message:'Profile not found',code:'PROFILE_NOT_FOUND'};
    return{success:true,profile:{userId:target.User_ID,email:target.Email,name:target.Name,phone:target.Phone,role:target.Role,department:target.Department,status:target.Status,isOwner:false}};
  }catch(e){return{success:false,message:e.message,code:'PROFILE_ERROR'};}
}
function updateClientOwnProfile(sid,data){
  try{
    var session=validateClientSession(sid);if(!session)return{success:false,message:'Session expired',code:'SESSION_EXPIRED'};
    if(String(session.user&&session.user.role||'').toLowerCase()==='owner')return{success:false,message:'Owner profile is managed through the main BizOS account.',code:'OWNER_PROFILE'};
    var sh=getClientTeamSheetV7_(),target=getClientTeamRowsV7_().find(function(r){return String(r.Email||'').toLowerCase()===String(session.email||'').toLowerCase();});
    if(!target)return{success:false,message:'Profile not found',code:'PROFILE_NOT_FOUND'};
    data=data||{};var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
    function set_(key,value){var i=headers.indexOf(key);if(i>=0)sh.getRange(target.row,i+1).setValue(value);}
    if(data.name!==undefined)set_('Name',String(data.name||'').trim());
    if(data.phone!==undefined)set_('Phone',String(data.phone||'').trim());
    if(data.department!==undefined)set_('Department',String(data.department||'').trim());
    if(data.password){set_('Password_Hash',hashClientTeamPasswordV7_(data.password));set_('Must_Change_Password','NO');}
    set_('Updated_At',new Date().toISOString());return{success:true,message:'Profile updated successfully.'};
  }catch(e){return{success:false,message:e.message,code:'PROFILE_UPDATE_ERROR'};}
}
function authenticateClientTeamMemberV7_(email,password){
  var target=getClientTeamRowsV7_().find(function(r){return String(r.Email||'').toLowerCase()===String(email||'').toLowerCase();});
  if(!target)return null;
  if(String(target.Status||'').toLowerCase()==='suspended')return{success:false,message:'Your team access is suspended. Please contact the business owner.',code:'TEAM_SUSPENDED'};
  if(String(target.Password_Hash||'')!==hashClientTeamPasswordV7_(password))return{success:false,message:'Incorrect password.',code:'AUTH_FAILED'};
  var now=Date.now(),sid='CLIENT_SESS_'+Utilities.getUuid(),user={email:String(target.Email),name:String(target.Name||target.Email),role:String(target.Role||'staff'),clientId:String(CLIENT_CONFIG.clientId),businessId:String(CLIENT_CONFIG.businessId)},session={email:String(target.Email),clientId:String(CLIENT_CONFIG.clientId),sheetId:String(CLIENT_CONFIG.sheetId),businessId:String(CLIENT_CONFIG.businessId),businessName:String(CLIENT_CONFIG.clientName),created:now,expiresAt:now+21600000,isClient:true,user:user};
  CacheService.getScriptCache().put(sid,JSON.stringify(session),21600);
  var sh=getClientTeamSheetV7_(),headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String),li=headers.indexOf('Last_Login'),st=headers.indexOf('Status');
  if(li>=0)sh.getRange(target.row,li+1).setValue(new Date().toISOString());
  if(st>=0&&String(target.Status||'').toLowerCase()==='invited')sh.getRange(target.row,st+1).setValue('active');
  return{success:true,sessionId:sid,clientName:CLIENT_CONFIG.clientName,primaryColor:CLIENT_CONFIG.primaryColor||'#5D2A86',logoUrl:CLIENT_CONFIG.logoUrl||'',user:user,mustChangePassword:String(target.Must_Change_Password||'').toUpperCase()==='YES'};
}
`;
}

function injectClientTeamUiV7_(html, settings) {
  if (!html) return html;
  var color = String(settings && settings.primaryColor || '#5D2A86');
  var teamSection = `<section id="team" class="hidden p-4 md:p-6 max-w-7xl mx-auto"><div class="flex items-center justify-between mb-5"><div><h2 class="text-2xl font-bold">Team</h2><p class="text-sm text-gray-500">People with access to this business.</p></div><button id="inviteTeamBtn" onclick="openInviteTeamV7()" class="px-4 py-2 rounded-lg text-white font-semibold" style="background:${color}"><i class="fas fa-user-plus mr-2"></i>Invite Team Member</button></div><div id="teamNotice" class="hidden mb-4 p-3 rounded-lg text-sm"></div><div class="card overflow-hidden"><div id="teamTable" class="overflow-x-auto">Loading team...</div></div></section>`;
  var teamModal = `<div id="teamModal" class="hidden fixed inset-0 bg-black/50 z-[60] items-center justify-center p-4"><div class="bg-white rounded-xl w-full max-w-lg p-5"><div class="flex items-center justify-between mb-4"><h3 id="teamModalTitle" class="font-bold text-lg">Invite Team Member</h3><button onclick="closeTeamModalV7()" class="text-2xl leading-none">×</button></div><div id="teamModalBody"></div><div class="flex justify-end gap-2 mt-5"><button onclick="closeTeamModalV7()" class="px-4 py-2 rounded-lg border">Cancel</button><button id="teamModalSave" onclick="saveTeamModalV7()" class="px-4 py-2 rounded-lg text-white font-semibold" style="background:${color}">Send Invitation</button></div></div></div>`;
  var teamJs = `<script>
function applyTeamPermissionsV7(){var u=JSON.parse(localStorage.getItem('bizos_client_user')||'{}'),r=String(u.role||'').toLowerCase(),b=document.getElementById('inviteTeamBtn');if(b)b.style.display=(r==='owner'||r==='admin')?'':'none';}
function showTeam(){hideAll();document.getElementById('team').classList.remove('hidden');document.getElementById('pageTitle').textContent='Team';document.getElementById('pageSubtitle').textContent='People with access to this business';closeSidebar();loadTeamV7();}
function teamEscV7(v){return String(v==null?'':v).replace(/[&<>\"']/g,function(m){return{'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;',"'":'&#039;'}[m];});}
function loadTeamV7(){run(function(r){if(!r||!r.success){document.getElementById('teamTable').innerHTML='<div class="p-6 text-red-600">'+teamEscV7(r&&r.message||'Unable to load team')+'</div>';return;}var m=r.members||[],u=JSON.parse(localStorage.getItem('bizos_client_user')||'{}'),can=['owner','admin'].indexOf(String(u.role||'').toLowerCase())>=0;if(!m.length){document.getElementById('teamTable').innerHTML='<div class="p-8 text-center text-gray-400">No team members yet. Invite someone to get started.</div>';return;}var h='<table class="min-w-full text-sm"><thead><tr class="border-b bg-gray-50"><th class="text-left p-3">Name</th><th class="text-left p-3">Email</th><th class="text-left p-3">Role</th><th class="text-left p-3">Department</th><th class="text-left p-3">Status</th>'+(can?'<th class="text-right p-3">Action</th>':'')+'</tr></thead><tbody>';m.forEach(function(x){h+='<tr class="border-b"><td class="p-3 font-medium">'+teamEscV7(x.name)+'</td><td class="p-3">'+teamEscV7(x.email)+'</td><td class="p-3">'+teamEscV7(x.role)+'</td><td class="p-3">'+teamEscV7(x.department||'—')+'</td><td class="p-3">'+teamEscV7(x.status)+'</td>'+(can?'<td class="p-3 text-right"><button class="mr-3 font-semibold" style="color:'+PRIMARY+'" onclick="editTeamMemberV7('+JSON.stringify(x.userId)+')">Edit</button><button class="text-red-600" onclick="removeTeamMemberV7('+JSON.stringify(x.userId)+')">Remove</button></td>':'')+'</tr>';});h+='</tbody></table>';document.getElementById('teamTable').innerHTML=h;}).getClientTeam(SID);}
function openTeamModalV7(){var m=document.getElementById('teamModal');m.classList.remove('hidden');m.style.display='flex';}
function closeTeamModalV7(){var m=document.getElementById('teamModal');m.classList.add('hidden');m.style.display='none';}
function openInviteTeamV7(){document.getElementById('teamModalTitle').textContent='Invite Team Member';document.getElementById('teamModalSave').textContent='Send Invitation';document.getElementById('teamModalSave').dataset.mode='invite';document.getElementById('teamModalBody').innerHTML='<div class="space-y-3"><input id="tm-name" class="w-full border rounded-lg px-3 py-2" placeholder="Full name"><input id="tm-email" type="email" class="w-full border rounded-lg px-3 py-2" placeholder="Email"><input id="tm-phone" class="w-full border rounded-lg px-3 py-2" placeholder="Phone"><input id="tm-dept" class="w-full border rounded-lg px-3 py-2" placeholder="Department"><select id="tm-role" class="w-full border rounded-lg px-3 py-2"><option value="staff">Staff</option><option value="admin">Admin</option></select></div>';openTeamModalV7();}
function editTeamMemberV7(uid){run(function(r){if(!r||!r.success)return;var m=(r.members||[]).find(function(x){return String(x.userId)===String(uid);});if(!m)return;document.getElementById('teamModalTitle').textContent='Edit Staff Profile';document.getElementById('teamModalSave').textContent='Save Changes';document.getElementById('teamModalSave').dataset.mode='edit';document.getElementById('teamModalSave').dataset.userId=m.userId;document.getElementById('teamModalBody').innerHTML='<div class="space-y-3"><input id="tm-name" class="w-full border rounded-lg px-3 py-2" value="'+teamEscV7(m.name)+'"><input id="tm-phone" class="w-full border rounded-lg px-3 py-2" value="'+teamEscV7(m.phone)+'"><input id="tm-dept" class="w-full border rounded-lg px-3 py-2" value="'+teamEscV7(m.department||'')+'"><select id="tm-role" class="w-full border rounded-lg px-3 py-2"><option value="staff" '+(m.role==='staff'?'selected':'')+'>Staff</option><option value="admin" '+(m.role==='admin'?'selected':'')+'>Admin</option></select><select id="tm-status" class="w-full border rounded-lg px-3 py-2"><option value="invited" '+(m.status==='invited'?'selected':'')+'>Invited</option><option value="active" '+(m.status==='active'?'selected':'')+'>Active</option><option value="suspended" '+(m.status==='suspended'?'selected':'')+'>Suspended</option></select></div>';openTeamModalV7();}).getClientTeam(SID);}
function saveTeamModalV7(){var b=document.getElementById('teamModalSave'),p={name:(document.getElementById('tm-name')||{}).value||'',phone:(document.getElementById('tm-phone')||{}).value||'',department:(document.getElementById('tm-dept')||{}).value||'',role:(document.getElementById('tm-role')||{}).value||'staff'};b.disabled=true;if((b.dataset.mode||'invite')==='invite'){p.email=(document.getElementById('tm-email')||{}).value||'';run(function(r){b.disabled=false;if(!r||!r.success){alert(r&&r.message||'Invitation failed');return;}closeTeamModalV7();loadTeamV7();alert(r.message);}).inviteClientTeamMember(SID,p);}else{p.status=(document.getElementById('tm-status')||{}).value||'active';run(function(r){b.disabled=false;if(!r||!r.success){alert(r&&r.message||'Update failed');return;}closeTeamModalV7();loadTeamV7();alert(r.message);}).updateClientTeamMember(SID,b.dataset.userId,p);}}
function removeTeamMemberV7(uid){if(!confirm('Remove this team member?'))return;run(function(r){if(!r||!r.success){alert(r&&r.message||'Remove failed');return;}loadTeamV7();}).removeClientTeamMember(SID,uid);}
function loadOwnProfileV7(){run(function(r){if(!r||!r.success){document.getElementById('profileBody').innerHTML='<div class="text-red-600">'+teamEscV7(r&&r.message||'Unable to load profile')+'</div>';return;}var p=r.profile||{};document.getElementById('profileBody').innerHTML='<div class="space-y-4"><input id="profile-name" class="w-full border rounded-lg px-3 py-2" value="'+teamEscV7(p.name)+'" '+(p.isOwner?'disabled':'')+'><input class="w-full border rounded-lg px-3 py-2 bg-gray-50" value="'+teamEscV7(p.email)+'" disabled><input id="profile-phone" class="w-full border rounded-lg px-3 py-2" value="'+teamEscV7(p.phone||'')+'" '+(p.isOwner?'disabled':'')+'><input id="profile-dept" class="w-full border rounded-lg px-3 py-2" value="'+teamEscV7(p.department||'')+'" '+(p.isOwner?'disabled':'')+'>'+(p.isOwner?'<div class="text-xs text-gray-500">Owner profile is managed through the main BizOS account.</div>':'<input id="profile-password" type="password" class="w-full border rounded-lg px-3 py-2" placeholder="New password"><button onclick="saveOwnProfileV7()" class="px-4 py-2 rounded-lg text-white font-semibold" style="background:'+PRIMARY+'">Save Profile</button>')+'</div>';}).getClientOwnProfile(SID);}
function saveOwnProfileV7(){var p={name:(document.getElementById('profile-name')||{}).value||'',phone:(document.getElementById('profile-phone')||{}).value||'',department:(document.getElementById('profile-dept')||{}).value||'',password:(document.getElementById('profile-password')||{}).value||''};run(function(r){if(!r||!r.success){alert(r&&r.message||'Profile update failed');return;}var u=JSON.parse(localStorage.getItem('bizos_client_user')||'{}');u.name=p.name;localStorage.setItem('bizos_client_user',JSON.stringify(u));document.getElementById('userName').textContent=p.name;loadOwnProfileV7();alert(r.message);}).updateClientOwnProfile(SID,p);}
</script>`;

  var marker = '</section>\n<div id="modal"';
  if (html.indexOf('id="team"') === -1 && html.indexOf(marker) !== -1) {
    html = html.replace(marker, teamSection + '\n<div id="modal"');
  }
  if (html.indexOf('id="teamModal"') === -1) {
    html = html.replace('<div id="modal"', teamModal + '\n<div id="modal"');
  }
  if (html.indexOf('function showTeam()') === -1) {
    html = html.replace('</script></main></body></html>', teamJs + '</script></main></body></html>');
  }
  var teamNav = '<button onclick="showTeam()" class="w-full text-left px-3 py-2.5 rounded-lg hover:bg-gray-100 flex items-center gap-3"><i class="fas fa-users w-5 text-gray-500"></i>Team</button>';
  if (html.indexOf('>Team</button>') === -1) {
    html = html.replace('<button onclick="showProfile()"', teamNav + '<button onclick="showProfile()"');
  }
  html = html.replace("function showProfile(){hideAll();document.getElementById('profile').classList.remove('hidden');document.getElementById('pageTitle').textContent='Profile';document.getElementById('pageSubtitle').textContent='Account information';var u=JSON.parse(localStorage.getItem('bizos_client_user')||'{}');document.getElementById('profileBody').innerHTML='<b>'+escapeHtml(CLIENT)+'</b><br><span class=\"text-gray-500\">'+escapeHtml(u.email||'')+'</span>';closeSidebar();}", "function showProfile(){hideAll();document.getElementById('profile').classList.remove('hidden');document.getElementById('pageTitle').textContent='Profile';document.getElementById('pageSubtitle').textContent='Account information';closeSidebar();loadOwnProfileV7();}");
  html = html.replace('(function(){var u=JSON.parse(localStorage.getItem(\'bizos_client_user\')||\'{}\');document.getElementById(\'userName\').textContent=u.name||\'Business User\';', '(function(){var u=JSON.parse(localStorage.getItem(\'bizos_client_user\')||\'{}\');document.getElementById(\'userName\').textContent=u.name||\'Business User\';');
  html = html.replace('document.getElementById(\'userEmail\').textContent=u.email||\'\';if(!SID)', 'document.getElementById(\'userEmail\').textContent=u.email||\'\';if(typeof applyTeamPermissionsV7===\'function\')applyTeamPermissionsV7();if(!SID)');
  return html;
}
