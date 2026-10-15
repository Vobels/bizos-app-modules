// BizOSMessages.js - persistent in-app conversations and admin announcements.
// Message bodies are stored in the BizOS master spreadsheet, never in Gmail.

var BIZOS_MESSAGE_HEADERS_V1_ = [
  'Message_ID','Conversation_ID','Business_ID','Scope','Sender_Email','Sender_Name',
  'Recipient_Email','Recipient_Name','Message_Body','Created_At','Read_At'
];

function ensureBizOSMessagesSheet_() {
  var ss = getBizOSMasterSpreadsheet_();
  var sheet = ss.getSheetByName('BizOS_Messages');
  if (!sheet) {
    sheet = ss.insertSheet('BizOS_Messages');
    sheet.getRange(1,1,1,BIZOS_MESSAGE_HEADERS_V1_.length).setValues([BIZOS_MESSAGE_HEADERS_V1_]);
    sheet.getRange(1,1,1,BIZOS_MESSAGE_HEADERS_V1_.length).setFontWeight('bold');
  } else {
    ensureBizOSSheetHeadersV1_(sheet, BIZOS_MESSAGE_HEADERS_V1_);
  }
  return sheet;
}

function bizOSMessageActor_(sessionId) {
  var user = getUserFromSession(sessionId);
  if (!user) return {ok:false,message:'Your session has expired. Please sign in again.'};
  return {ok:true,user:user};
}

function bizOSMessageEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function bizOSMessageRole_(user) {
  return String(user && user.role || 'staff').trim().toLowerCase();
}

function bizOSMessageUserInBusiness_(email, businessId) {
  email = bizOSMessageEmail_(email);
  businessId = String(businessId || '').trim();
  if (!email || !businessId || email === 'bizos-support') return false;
  var ss = getBizOSMasterSpreadsheet_();
  var sheet = ss.getSheetByName('Users') || ss.getSheetByName('User');
  if (!sheet || sheet.getLastRow() < 2) return false;
  var data = sheet.getDataRange().getValues();
  var headers = data[0] || [];
  var idx = {};
  headers.forEach(function(h,i){idx[String(h || '').trim().toLowerCase()] = i;});
  var emailCol = idx.email;
  var businessCol = idx.business_id;
  var statusCol = idx.status;
  for (var i=1;i<data.length;i++) {
    var row = data[i];
    if (emailCol == null || businessCol == null) continue;
    if (bizOSMessageEmail_(row[emailCol]) !== email) continue;
    if (String(row[businessCol] || '').trim() !== businessId) continue;
    if (statusCol != null && String(row[statusCol] || '').trim().toLowerCase() === 'disabled') continue;
    return true;
  }
  return false;
}

function getBizOSMessageContacts(sessionId) {
  var actor = bizOSMessageActor_(sessionId);
  if (!actor.ok) return {success:false,message:actor.message};
  try {
    var ss = getBizOSMasterSpreadsheet_();
    var sheet = ss.getSheetByName('Users') || ss.getSheetByName('User');
    if (!sheet || sheet.getLastRow() < 2) return {success:true,contacts:[]};
    var data = sheet.getDataRange().getValues(), headers = data[0] || [], idx = {};
    headers.forEach(function(h,i){idx[String(h || '').trim().toLowerCase()] = i;});
    var e=idx.email, b=idx.business_id, n=idx.name, r=idx.role, st=idx.status;
    var contacts=[], seen={};
    if(e==null||b==null)return {success:false,message:'The team directory is not configured correctly.'};
    for(var i=1;i<data.length;i++){
      var row=data[i], email=bizOSMessageEmail_(row[e]);
      if(!email||email===bizOSMessageEmail_(actor.user.email)||String(row[b]||'').trim()!==String(actor.user.businessId||'').trim())continue;
      if(st!=null&&['disabled','removed','inactive'].indexOf(String(row[st]||'').trim().toLowerCase())>=0)continue;
      if(seen[email])continue;seen[email]=true;
      contacts.push({email:email,name:String(n!=null?row[n]:'')||email,role:String(r!=null?row[r]:'staff')});
    }
    contacts.sort(function(a,b){return a.name.localeCompare(b.name);});
    return {success:true,contacts:contacts};
  } catch(error) {
    console.error('getBizOSMessageContacts error:',error);
    return {success:false,message:'Could not load your team directory.'};
  }
}

function bizOSMessageRowToObject_(row, idx) {
  return {
    id:String(row[idx.Message_ID]||''),
    conversationId:String(row[idx.Conversation_ID]||''),
    businessId:String(row[idx.Business_ID]||''),
    scope:String(row[idx.Scope]||'team'),
    senderEmail:String(row[idx.Sender_Email]||''),
    senderName:String(row[idx.Sender_Name]||''),
    recipientEmail:String(row[idx.Recipient_Email]||''),
    recipientName:String(row[idx.Recipient_Name]||''),
    message:String(row[idx.Message_Body]||''),
    createdAt:String(row[idx.Created_At]||''),
    readAt:String(row[idx.Read_At]||'')
  };
}

function getBizOSMessages(conversationId, sessionId) {
  var actor = bizOSMessageActor_(sessionId);
  if (!actor.ok) return {success:false,message:actor.message};
  try {
    var user=actor.user, email=bizOSMessageEmail_(user.email), businessId=String(user.businessId||'').trim();
    var isOwnerAdmin=['owner','admin'].indexOf(bizOSMessageRole_(user))>=0;
    var sheet=ensureBizOSMessagesSheet_(), data=sheet.getDataRange().getValues(), headers=data[0]||[], idx={};
    headers.forEach(function(h,i){idx[h]=i;});
    var wanted=String(conversationId||'').trim(), rows=[];
    for(var i=1;i<data.length;i++){
      var row=data[i], scope=String(row[idx.Scope]||'team'), rowBiz=String(row[idx.Business_ID]||'');
      if(rowBiz!==businessId)continue;
      var sender=bizOSMessageEmail_(row[idx.Sender_Email]), recipient=bizOSMessageEmail_(row[idx.Recipient_Email]);
      var canSee=(sender===email||recipient===email);
      if(scope==='support' && isOwnerAdmin)canSee=true;
      if(scope==='support'&&!isOwnerAdmin)canSee=false;
      if(!canSee)continue;
      var item=bizOSMessageRowToObject_(row,idx);
      if(wanted&&item.conversationId!==wanted)continue;
      rows.push(item);
    }
    rows.sort(function(a,b){return new Date(a.createdAt)-new Date(b.createdAt);});
    return {success:true,messages:rows.slice(-200)};
  } catch(error) {
    console.error('getBizOSMessages error:',error);
    return {success:false,message:'Could not load messages.'};
  }
}

function sendBizOSMessage(payload, sessionId) {
  var actor=bizOSMessageActor_(sessionId);
  if(!actor.ok)return {success:false,message:actor.message};
  payload=payload||{};
  var user=actor.user, sender=bizOSMessageEmail_(user.email), businessId=String(user.businessId||'').trim();
  var body=String(payload.message||'').trim(), scope=String(payload.scope||'team').trim().toLowerCase();
  var recipient=bizOSMessageEmail_(payload.recipientEmail);
  if(!body||body.length>5000)return {success:false,message:'Enter a message of up to 5,000 characters.'};
  if(!businessId||!sender)return {success:false,message:'Your account is missing required workspace details.'};
  var support=scope==='support';
  if(support){
    if(['owner','admin'].indexOf(bizOSMessageRole_(user))<0)return {success:false,message:'Only the owner or an administrator can contact BizOS Support.'};
    recipient='bizos-support';
  } else {
    scope='team';
    if(!bizOSMessageUserInBusiness_(recipient,businessId))return {success:false,message:'Choose a valid team member in your workspace.'};
    if(recipient===sender)return {success:false,message:'Choose another team member.'};
  }
  try {
    var sheet=ensureBizOSMessagesSheet_(), id='MSG_'+Utilities.getUuid().replace(/-/g,'').substring(0,20).toUpperCase();
    var conversationId=String(payload.conversationId||'').trim();
    if(conversationId){
      var data=sheet.getDataRange().getValues(), h=data[0]||[], ix={};h.forEach(function(k,i){ix[k]=i;});
      var valid=false;
      for(var i=1;i<data.length;i++){
        var row=data[i];
        if(String(row[ix.Conversation_ID]||'')!==conversationId||String(row[ix.Business_ID]||'')!==businessId||String(row[ix.Scope]||'')!==scope)continue;
        var s=bizOSMessageEmail_(row[ix.Sender_Email]), r=bizOSMessageEmail_(row[ix.Recipient_Email]);
        if(s===sender||r===sender||(support&&['owner','admin'].indexOf(bizOSMessageRole_(user))>=0)){valid=true;break;}
      }
      if(!valid)return {success:false,message:'That conversation could not be found.'};
    } else {
      conversationId='CONV_'+Utilities.getUuid().replace(/-/g,'').substring(0,20).toUpperCase();
    }
    var recipientName=support?'BizOS Support':String(payload.recipientName||recipient);
    sheet.appendRow([id,conversationId,businessId,scope,sender,String(user.name||user.fullName||user.email||sender),recipient,recipientName,body,new Date().toISOString(),'']);
    if(!support){
      createBizOSNotification_(businessId,recipient,'message','New message from '+String(user.name||user.email||sender),body.slice(0,240),sender,'messages');
    } else {
      createBizOSNotification_('ALL','admins','message','New BizOS Support message',String(user.businessName||businessId)+': '+body.slice(0,180),sender,'admin-messages');
    }
    logBizOSActivity_(businessId,sender,String(user.name||user.email||sender),'message_sent',support?'Sent a message to BizOS Support.':'Sent a team message.','messages');
    return {success:true,messageId:id,conversationId:conversationId};
  } catch(error) {
    console.error('sendBizOSMessage error:',error);
    return {success:false,message:'Your message could not be sent. Please try again.'};
  }
}

function getBizOSSupportInbox(adminSessionId) {
  requireAdminSession_(adminSessionId);
  try {
    var sheet=ensureBizOSMessagesSheet_(), data=sheet.getDataRange().getValues(), headers=data[0]||[], idx={};
    headers.forEach(function(h,i){idx[h]=i;});
    var conversations={};
    for(var i=1;i<data.length;i++){
      var item=bizOSMessageRowToObject_(data[i],idx);
      if(item.scope!=='support')continue;
      if(!conversations[item.conversationId])conversations[item.conversationId]={conversationId:item.conversationId,businessId:item.businessId,clientEmail:'',clientName:'',lastMessage:'',lastAt:'',messages:[]};
      var c=conversations[item.conversationId];
      if(item.senderEmail!=='bizos-support'){c.clientEmail=item.senderEmail;c.clientName=item.senderName;}
      c.messages.push(item);c.lastMessage=item.message;c.lastAt=item.createdAt;
    }
    var list=Object.keys(conversations).map(function(k){return conversations[k];});
    list.sort(function(a,b){return new Date(b.lastAt)-new Date(a.lastAt);});
    var unreadCount=0;
    for(var ri=1;ri<data.length;ri++){
      var message=bizOSMessageRowToObject_(data[ri],idx);
      if(message.scope==='support'&&bizOSMessageEmail_(message.recipientEmail)==='bizos-support'&&!message.readAt)unreadCount++;
    }
    return {success:true,conversations:list.slice(0,100),unreadCount:unreadCount};
  } catch(error) {
    console.error('getBizOSSupportInbox error:',error);
    return {success:false,message:'Could not load the support inbox.'};
  }
}

function sendBizOSSupportReply(adminSessionId, conversationId, message) {
  requireAdminSession_(adminSessionId);
  conversationId=String(conversationId||'').trim();message=String(message||'').trim();
  if(!conversationId||!message||message.length>5000)return {success:false,message:'Enter a reply of up to 5,000 characters.'};
  try {
    var sheet=ensureBizOSMessagesSheet_(), data=sheet.getDataRange().getValues(), headers=data[0]||[], idx={};
    headers.forEach(function(h,i){idx[h]=i;});
    var thread=null,clientEmail='',clientName='';
    for(var i=1;i<data.length;i++){
      var row=data[i], item=bizOSMessageRowToObject_(row,idx);
      if(item.conversationId!==conversationId||item.scope!=='support')continue;
      thread=item;
      if(item.senderEmail!=='bizos-support'){clientEmail=item.senderEmail;clientName=item.senderName;}
    }
    if(!thread||!clientEmail)return {success:false,message:'Support conversation not found.'};
    var id='MSG_'+Utilities.getUuid().replace(/-/g,'').substring(0,20).toUpperCase();
    var supportEmail='bizos-support';
    sheet.appendRow([id,conversationId,thread.businessId,'support',supportEmail,'BizOS Support',clientEmail,clientName,message,new Date().toISOString(),'']);
    createBizOSNotification_(thread.businessId,clientEmail,'message','Reply from BizOS Support',message.slice(0,240),supportEmail,'messages');
    var workspaceRecords=bizOSClientWorkspaceRecords_();
    var workspace=workspaceRecords.find(function(w){return w.businessId===thread.businessId;});
    if(workspace)writeClientWorkspaceNotificationV12_(workspace.sheetId,{id:'NTF_'+Utilities.getUuid().replace(/-/g,'').substring(0,20).toUpperCase(),audience:clientEmail,type:'message',title:'Reply from BizOS Support',message:message.slice(0,240),createdAt:new Date().toISOString(),createdBy:'BizOS Support',relatedView:'notifications'});
    return {success:true,messageId:id};
  }catch(error){
    console.error('sendBizOSSupportReply error:',error);
    return {success:false,message:'Reply could not be sent.'};
  }
}

function writeClientWorkspaceNotificationV12_(sheetId, notification) {
  sheetId=String(sheetId||'').trim();
  if(!sheetId)return {success:false,code:'WORKSPACE_ID_MISSING'};
  try {
    var ws=SpreadsheetApp.openById(sheetId), sheet=ws.getSheetByName('Client_Notifications');
    var headers=['Notification_ID','Audience','Type','Title','Message','Created_At','Created_By','Related_View','Read_By_JSON'];
    if(!sheet){
      sheet=ws.insertSheet('Client_Notifications');
      sheet.getRange(1,1,1,headers.length).setValues([headers]);
      sheet.getRange(1,1,1,headers.length).setFontWeight('bold');
    } else {
      ensureBizOSSheetHeadersV1_(sheet,headers);
      headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]||headers;
    }
    var data=sheet.getDataRange().getValues(), h=data[0]||[], ix={};
    h.forEach(function(v,i){ix[String(v||'')]=i;});
    var id=String(notification.id||'');
    if(id&&ix.Notification_ID!=null){
      for(var i=1;i<data.length;i++)if(String(data[i][ix.Notification_ID]||'')===id)return {success:true,duplicate:true};
    }
    var values={
      Notification_ID:id||('NTF_'+Utilities.getUuid().replace(/-/g,'').substring(0,20).toUpperCase()),
      Audience:String(notification.audience||'all'),
      Type:String(notification.type||'info'),
      Title:String(notification.title||'Workspace update'),
      Message:String(notification.message||''),
      Created_At:String(notification.createdAt||new Date().toISOString()),
      Created_By:String(notification.createdBy||'BizOS'),
      Related_View:String(notification.relatedView||'notifications'),
      Read_By_JSON:'{}'
    };
    var row=new Array(h.length).fill('');
    h.forEach(function(k,j){if(Object.prototype.hasOwnProperty.call(values,k))row[j]=values[k];});
    sheet.appendRow(row);
    return {success:true};
  } catch(error) {
    console.error('writeClientWorkspaceNotificationV12_ error for workspace '+sheetId+':',error);
    return {success:false,message:error.message};
  }
}

function bizOSClientWorkspaceRecords_() {
  var ss=getBizOSMasterSpreadsheet_(), sheet=ss.getSheetByName('Clients');
  if(!sheet||sheet.getLastRow()<2)return [];
  var data=sheet.getDataRange().getValues(), headers=data[0]||[], idx={};
  headers.forEach(function(h,i){idx[String(h||'').trim()]=i;});
  var b=idx.Business_ID, w=idx.Workspace_ID!=null?idx.Workspace_ID:idx.Sheet_ID, st=idx.Status;
  if(b==null||w==null)return [];
  var byBusiness={};
  for(var i=1;i<data.length;i++){
    var row=data[i], businessId=String(row[b]||'').trim(), sheetId=String(row[w]||'').trim();
    if(!businessId||!sheetId)continue;
    var status=st==null?'active':String(row[st]||'active').trim().toLowerCase();
    if(['archived','disabled','inactive','failed','provisioning_failed','deleted'].indexOf(status)>=0)continue;
    var current=byBusiness[businessId];
    if(!current||status==='active'||current.status!=='active')byBusiness[businessId]={businessId:businessId,sheetId:sheetId,status:status};
  }
  return Object.keys(byBusiness).map(function(k){return byBusiness[k];});
}

function broadcastBizOSNotificationToClientWorkspaces_(notification) {
  var workspaces=bizOSClientWorkspaceRecords_(), delivered=0, failed=0;
  workspaces.forEach(function(w){
    var result=writeClientWorkspaceNotificationV12_(w.sheetId,{
      id:notification.id,audience:'all',type:notification.type,title:notification.title,
      message:notification.message,createdAt:notification.createdAt,createdBy:notification.createdBy,
      relatedView:notification.relatedView
    });
    if(result&&result.success)delivered++;else failed++;
  });
  return {workspaces:workspaces.length,delivered:delivered,failed:failed};
}

function adminPublishBizOSAnnouncement(adminSessionId, payload) {
  requireAdminSession_(adminSessionId);
  payload=payload||{};
  var title=String(payload.title||'').trim(), message=String(payload.message||'').trim();
  var type=String(payload.type||'feature').trim().toLowerCase();
  var relatedView=String(payload.relatedView||'notifications').trim();
  if(!title||title.length>140||!message||message.length>5000)return {success:false,message:'Enter a title (up to 140 characters) and message (up to 5,000 characters).'};
  if(['feature','info','success','warning','security','system'].indexOf(type)<0)type='feature';
  var created=createBizOSNotification_('ALL','all',type,title,message,'BizOS','notifications');
  if(!created||!created.success)return created||{success:false,message:'Announcement could not be published.'};
  var broadcast=broadcastBizOSNotificationToClientWorkspaces_({
    id:created.notificationId,audience:'all',type:type,title:title,message:message,
    createdAt:new Date().toISOString(),createdBy:'BizOS',relatedView:relatedView
  });
  return {success:true,notificationId:created.notificationId,clientWorkspaces:broadcast.workspaces,clientWorkspacesNotified:broadcast.delivered,clientWorkspaceFailures:broadcast.failed};
}


/**
 * Return only the unread message count needed by the sidebar badge.
 * Keeps the page load light and uses the same per-user read state as the inbox.
 */
function getBizOSMessageUnreadCount(sessionId) {
  var actor = bizOSMessageActor_(sessionId);
  if (!actor.ok) return {success:false,message:actor.message};
  try {
    var email = bizOSMessageEmail_(actor.user.email);
    var businessId = String(actor.user.businessId || '').trim();
    if (!email || !businessId) return {success:true,unreadCount:0};
    var sheet = ensureBizOSMessagesSheet_();
    var data = sheet.getDataRange().getValues();
    var headers = data[0] || [], idx = {};
    headers.forEach(function(h,i){idx[String(h || '').trim()] = i;});
    if (idx.Business_ID == null || idx.Recipient_Email == null || idx.Read_At == null) {
      return {success:false,message:'Message storage is not configured correctly.'};
    }
    var unread = 0;
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[idx.Business_ID] || '').trim() !== businessId) continue;
      if (bizOSMessageEmail_(row[idx.Recipient_Email]) !== email) continue;
      if (String(row[idx.Read_At] || '').trim()) continue;
      unread++;
    }
    return {success:true,unreadCount:unread};
  } catch (error) {
    console.error('getBizOSMessageUnreadCount error:', error);
    return {success:false,message:'Could not load the message count.'};
  }
}

function markBizOSConversationRead(conversationId, sessionId) {
  var actor=bizOSMessageActor_(sessionId);
  if(!actor.ok)return {success:false,message:actor.message};
  conversationId=String(conversationId||'').trim();
  if(!conversationId)return {success:false,message:'Conversation not found.'};
  try {
    var user=actor.user,email=bizOSMessageEmail_(user.email),isOwnerAdmin=['owner','admin'].indexOf(bizOSMessageRole_(user))>=0;
    var sheet=ensureBizOSMessagesSheet_(),data=sheet.getDataRange().getValues(),headers=data[0]||[],idx={};
    headers.forEach(function(h,i){idx[h]=i;});
    var now=new Date().toISOString(),changed=0;
    for(var i=1;i<data.length;i++){
      var row=data[i];
      if(String(row[idx.Conversation_ID]||'')!==conversationId||String(row[idx.Business_ID]||'')!==String(user.businessId||''))continue;
      if(String(row[idx.Scope]||'')==='support'&&!isOwnerAdmin)continue;
      var recipient=bizOSMessageEmail_(row[idx.Recipient_Email]);
      if(recipient!==email&&!(String(row[idx.Scope]||'')==='support'&&isOwnerAdmin))continue;
      if(String(row[idx.Read_At]||''))continue;
      sheet.getRange(i+1,idx.Read_At+1).setValue(now);changed++;
    }
    return {success:true,updated:changed};
  }catch(error){console.error('markBizOSConversationRead error:',error);return{success:false,message:'Could not mark messages as read.'};}
}

function markBizOSSupportConversationRead(adminSessionId, conversationId) {
  requireAdminSession_(adminSessionId);
  conversationId=String(conversationId||'').trim();
  if(!conversationId)return {success:false,message:'Support conversation not found.'};
  try {
    var sheet=ensureBizOSMessagesSheet_(),data=sheet.getDataRange().getValues(),headers=data[0]||[],idx={};
    headers.forEach(function(h,i){idx[h]=i;});
    var now=new Date().toISOString(),changed=0;
    for(var i=1;i<data.length;i++){
      var row=data[i];
      if(String(row[idx.Conversation_ID]||'')!==conversationId||String(row[idx.Scope]||'')!=='support')continue;
      if(bizOSMessageEmail_(row[idx.Recipient_Email])!=='bizos-support'||String(row[idx.Read_At]||''))continue;
      sheet.getRange(i+1,idx.Read_At+1).setValue(now);changed++;
    }
    return {success:true,updated:changed};
  }catch(error){console.error('markBizOSSupportConversationRead error:',error);return{success:false,message:'Could not mark support messages as read.'};}
}
