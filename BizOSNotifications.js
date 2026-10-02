// BizOSNotifications.js - persistent BizOS notification center + business activity

var BIZOS_NOTIFICATION_HEADERS_V1_ = [
  'Notification_ID','Business_ID','Audience','Type','Title','Message',
  'Created_At','Created_By','Related_View','Read_By_JSON'
];

var BIZOS_ACTIVITY_HEADERS_V1_ = [
  'Activity_ID','Business_ID','Actor_Email','Actor_Name','Event_Type',
  'Summary','Created_At','Related_View'
];

var BIZOS_NOTIFICATION_PREF_HEADERS_V1_ = [
  'Preference_ID','Business_ID','Email','Feature_Updates','System_Alerts',
  'Weekly_Reports','Security_Alerts','Marketing_Updates','Updated_At'
];

function ensureBizOSNotificationsSheet_() {
  var ss = getBizOSMasterSpreadsheet_();
  var sheet = ss.getSheetByName('BizOS_Notifications');
  if (!sheet) {
    sheet = ss.insertSheet('BizOS_Notifications');
    sheet.getRange(1,1,1,BIZOS_NOTIFICATION_HEADERS_V1_.length)
      .setValues([BIZOS_NOTIFICATION_HEADERS_V1_]);
    sheet.getRange(1,1,1,BIZOS_NOTIFICATION_HEADERS_V1_.length).setFontWeight('bold');
  } else {
    ensureBizOSSheetHeadersV1_(sheet, BIZOS_NOTIFICATION_HEADERS_V1_);
  }
  return sheet;
}

function ensureBizOSActivitySheet_() {
  var ss = getBizOSMasterSpreadsheet_();
  var sheet = ss.getSheetByName('BizOS_Activity');
  if (!sheet) {
    sheet = ss.insertSheet('BizOS_Activity');
    sheet.getRange(1,1,1,BIZOS_ACTIVITY_HEADERS_V1_.length)
      .setValues([BIZOS_ACTIVITY_HEADERS_V1_]);
    sheet.getRange(1,1,1,BIZOS_ACTIVITY_HEADERS_V1_.length).setFontWeight('bold');
  } else {
    ensureBizOSSheetHeadersV1_(sheet, BIZOS_ACTIVITY_HEADERS_V1_);
  }
  return sheet;
}

function ensureBizOSNotificationPreferencesSheet_() {
  var ss = getBizOSMasterSpreadsheet_();
  var sheet = ss.getSheetByName('BizOS_Notification_Preferences');
  if (!sheet) {
    sheet = ss.insertSheet('BizOS_Notification_Preferences');
    sheet.getRange(1,1,1,BIZOS_NOTIFICATION_PREF_HEADERS_V1_.length)
      .setValues([BIZOS_NOTIFICATION_PREF_HEADERS_V1_]);
    sheet.getRange(1,1,1,BIZOS_NOTIFICATION_PREF_HEADERS_V1_.length).setFontWeight('bold');
  } else {
    ensureBizOSSheetHeadersV1_(sheet, BIZOS_NOTIFICATION_PREF_HEADERS_V1_);
  }
  return sheet;
}

function ensureBizOSSheetHeadersV1_(sheet, requiredHeaders) {
  var lastColumn = Math.max(1, sheet.getLastColumn());
  var current = sheet.getRange(1,1,1,lastColumn).getValues()[0] || [];
  var changed = false;
  requiredHeaders.forEach(function(header) {
    if (current.indexOf(header) === -1) {
      current.push(header);
      changed = true;
    }
  });
  if (changed) {
    sheet.getRange(1,1,1,current.length).setValues([current]);
    sheet.getRange(1,1,1,current.length).setFontWeight('bold');
  }
  return current;
}

function createBizOSNotification_(businessId, audience, type, title, message, createdBy, relatedView) {
  try {
    businessId = String(businessId || '').trim();
    if (!businessId || !String(title || '').trim()) return {success:false,code:'NOTIFICATION_INPUT_INVALID'};
    var sheet = ensureBizOSNotificationsSheet_();
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] || [];
    var row = new Array(headers.length).fill('');
    var values = {
      Notification_ID:'NTF_' + Utilities.getUuid().replace(/-/g,'').substring(0,20).toUpperCase(),
      Business_ID:businessId,
      Audience:String(audience || 'all').trim(),
      Type:String(type || 'info').trim(),
      Title:String(title || '').trim(),
      Message:String(message || '').trim(),
      Created_At:new Date().toISOString(),
      Created_By:String(createdBy || '').trim(),
      Related_View:String(relatedView || '').trim(),
      Read_By_JSON:'[]'
    };
    headers.forEach(function(header,index){ if (Object.prototype.hasOwnProperty.call(values,header)) row[index]=values[header]; });
    sheet.appendRow(row);
    return {success:true,notificationId:values.Notification_ID};
  } catch (error) {
    console.error('createBizOSNotification_ error:',error);
    return {success:false,code:'NOTIFICATION_CREATE_FAILED',message:error.message};
  }
}

function logBizOSActivity_(businessId, actorEmail, actorName, eventType, summary, relatedView) {
  try {
    businessId = String(businessId || '').trim();
    if (!businessId || !String(summary || '').trim()) return {success:false,code:'ACTIVITY_INPUT_INVALID'};
    var sheet = ensureBizOSActivitySheet_();
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] || [];
    var row = new Array(headers.length).fill('');
    var values = {
      Activity_ID:'ACT_' + Utilities.getUuid().replace(/-/g,'').substring(0,20).toUpperCase(),
      Business_ID:businessId,
      Actor_Email:String(actorEmail || '').trim().toLowerCase(),
      Actor_Name:String(actorName || '').trim(),
      Event_Type:String(eventType || 'activity').trim(),
      Summary:String(summary || '').trim(),
      Created_At:new Date().toISOString(),
      Related_View:String(relatedView || '').trim()
    };
    headers.forEach(function(header,index){ if (Object.prototype.hasOwnProperty.call(values,header)) row[index]=values[header]; });
    sheet.appendRow(row);
    return {success:true,activityId:values.Activity_ID};
  } catch (error) {
    console.error('logBizOSActivity_ error:',error);
    return {success:false,code:'ACTIVITY_CREATE_FAILED',message:error.message};
  }
}

function bizOSNotificationActor_(sessionId) {
  var user = getUserFromSession(sessionId);
  if (!user) return {ok:false,message:'Session expired. Please sign in again.'};
  return {ok:true,user:user};
}

function bizOSNotificationAudienceMatches_(audience, user) {
  var a = String(audience || 'all').trim().toLowerCase();
  var email = String(user.email || '').trim().toLowerCase();
  var role = String(user.role || 'staff').trim().toLowerCase();
  if (!a || a === 'all') return true;
  if (a === email) return true;
  if (a === 'owner_admin' || a === 'owner-admin' || a === 'admins') return role === 'owner' || role === 'admin';
  if (a === 'staff') return role === 'staff';
  return false;
}

function bizOSNotificationReadSet_(raw) {
  try {
    var parsed = JSON.parse(String(raw || '[]'));
    return Array.isArray(parsed) ? parsed.map(function(v){return String(v || '').trim().toLowerCase();}).filter(Boolean) : [];
  } catch (e) {
    return [];
  }
}

function getBizOSNotifications(sessionId, limit) {
  var actor = bizOSNotificationActor_(sessionId);
  if (!actor.ok) return {success:false,message:actor.message};
  try {
    var sheet = ensureBizOSNotificationsSheet_();
    var data = sheet.getDataRange().getValues();
    var headers = data[0] || [];
    var index = {};
    headers.forEach(function(h,i){index[h]=i;});
    var max = Math.min(Math.max(parseInt(limit,10) || 50,1),100);
    var rows = [];
    for (var i=data.length-1;i>=1 && rows.length<max;i--) {
      var row=data[i];
      if (String(row[index.Business_ID] || '') !== String(actor.user.businessId || '')) continue;
      if (!bizOSNotificationAudienceMatches_(row[index.Audience], actor.user)) continue;
      var readBy=bizOSNotificationReadSet_(row[index.Read_By_JSON]);
      var email=String(actor.user.email || '').trim().toLowerCase();
      rows.push({
        id:String(row[index.Notification_ID] || ''),
        audience:String(row[index.Audience] || 'all'),
        type:String(row[index.Type] || 'info'),
        title:String(row[index.Title] || ''),
        message:String(row[index.Message] || ''),
        createdAt:String(row[index.Created_At] || ''),
        createdBy:String(row[index.Created_By] || ''),
        relatedView:String(row[index.Related_View] || ''),
        isRead:readBy.indexOf(email)!==-1
      });
    }
    var unreadCount=rows.filter(function(n){return !n.isRead;}).length;
    return {success:true,notifications:rows,unreadCount:unreadCount};
  } catch (error) {
    console.error('getBizOSNotifications error:',error);
    return {success:false,message:'Could not load notifications.'};
  }
}

function markBizOSNotificationRead(sessionId, notificationId) {
  var actor = bizOSNotificationActor_(sessionId);
  if (!actor.ok) return {success:false,message:actor.message};
  return updateBizOSNotificationReadState_(actor.user, notificationId, false);
}

function markAllBizOSNotificationsRead(sessionId) {
  var actor = bizOSNotificationActor_(sessionId);
  if (!actor.ok) return {success:false,message:actor.message};
  try {
    var sheet=ensureBizOSNotificationsSheet_();
    var data=sheet.getDataRange().getValues();
    var headers=data[0]||[];
    var idx={};
    headers.forEach(function(h,i){idx[h]=i;});
    var email=String(actor.user.email || '').trim().toLowerCase();
    var changed=0;
    for(var i=1;i<data.length;i++){
      if(String(data[i][idx.Business_ID]||'')!==String(actor.user.businessId||''))continue;
      if(!bizOSNotificationAudienceMatches_(data[i][idx.Audience],actor.user))continue;
      var readBy=bizOSNotificationReadSet_(data[i][idx.Read_By_JSON]);
      if(readBy.indexOf(email)===-1){
        readBy.push(email);
        sheet.getRange(i+1,idx.Read_By_JSON+1).setValue(JSON.stringify(readBy));
        changed++;
      }
    }
    return {success:true,updated:changed};
  }catch(error){
    console.error('markAllBizOSNotificationsRead error:',error);
    return {success:false,message:'Could not update notifications.'};
  }
}

function updateBizOSNotificationReadState_(user, notificationId, markUnread) {
  try {
    var sheet=ensureBizOSNotificationsSheet_();
    var data=sheet.getDataRange().getValues();
    var headers=data[0]||[];
    var idx={};
    headers.forEach(function(h,i){idx[h]=i;});
    var email=String(user.email || '').trim().toLowerCase();
    for(var i=1;i<data.length;i++){
      if(String(data[i][idx.Notification_ID]||'')!==String(notificationId||''))continue;
      if(String(data[i][idx.Business_ID]||'')!==String(user.businessId||''))return {success:false,message:'Notification not found.'};
      if(!bizOSNotificationAudienceMatches_(data[i][idx.Audience],user))return {success:false,message:'Notification not found.'};
      var readBy=bizOSNotificationReadSet_(data[i][idx.Read_By_JSON]);
      var pos=readBy.indexOf(email);
      if(markUnread && pos!==-1) readBy.splice(pos,1);
      if(!markUnread && pos===-1) readBy.push(email);
      sheet.getRange(i+1,idx.Read_By_JSON+1).setValue(JSON.stringify(readBy));
      return {success:true};
    }
    return {success:false,message:'Notification not found.'};
  }catch(error){
    console.error('updateBizOSNotificationReadState_ error:',error);
    return {success:false,message:'Could not update notification.'};
  }
}

function getBizOSActivity(sessionId, limit) {
  var actor = bizOSNotificationActor_(sessionId);
  if (!actor.ok) return {success:false,message:actor.message};
  try {
    var sheet=ensureBizOSActivitySheet_();
    var data=sheet.getDataRange().getValues();
    var headers=data[0]||[];
    var idx={};
    headers.forEach(function(h,i){idx[h]=i;});
    var max=Math.min(Math.max(parseInt(limit,10)||50,1),100);
    var activity=[];
    for(var i=data.length-1;i>=1 && activity.length<max;i--){
      var row=data[i];
      if(String(row[idx.Business_ID]||'')!==String(actor.user.businessId||''))continue;
      activity.push({
        id:String(row[idx.Activity_ID]||''),
        actorEmail:String(row[idx.Actor_Email]||''),
        actorName:String(row[idx.Actor_Name]||row[idx.Actor_Email]||''),
        eventType:String(row[idx.Event_Type]||'activity'),
        summary:String(row[idx.Summary]||''),
        createdAt:String(row[idx.Created_At]||''),
        relatedView:String(row[idx.Related_View]||'')
      });
    }
    return {success:true,activity:activity};
  }catch(error){
    console.error('getBizOSActivity error:',error);
    return {success:false,message:'Could not load activity.'};
  }
}

function getBizOSNotificationPreferences(sessionId) {
  var actor=bizOSNotificationActor_(sessionId);
  if(!actor.ok)return {success:false,message:actor.message};
  try{
    var sheet=ensureBizOSNotificationPreferencesSheet_();
    var data=sheet.getDataRange().getValues();
    var headers=data[0]||[];
    var idx={};headers.forEach(function(h,i){idx[h]=i;});
    var email=String(actor.user.email||'').trim().toLowerCase();
    for(var i=1;i<data.length;i++){
      if(String(data[i][idx.Business_ID]||'')===String(actor.user.businessId||'')&&String(data[i][idx.Email]||'').trim().toLowerCase()===email){
        return {success:true,preferences:{
          features:data[i][idx.Feature_Updates]!==false,
          system:data[i][idx.System_Alerts]!==false,
          weekly:data[i][idx.Weekly_Reports]===true,
          security:data[i][idx.Security_Alerts]!==false,
          marketing:data[i][idx.Marketing_Updates]===true
        }};
      }
    }
    return {success:true,preferences:{features:true,system:true,weekly:false,security:true,marketing:false}};
  }catch(error){
    console.error('getBizOSNotificationPreferences error:',error);
    return {success:false,message:'Could not load notification preferences.'};
  }
}

function saveBizOSNotificationPreferences(sessionId, preferences) {
  var actor=bizOSNotificationActor_(sessionId);
  if(!actor.ok)return {success:false,message:actor.message};
  try{
    var sheet=ensureBizOSNotificationPreferencesSheet_();
    var data=sheet.getDataRange().getValues();
    var headers=data[0]||[];
    var idx={};headers.forEach(function(h,i){idx[h]=i;});
    var email=String(actor.user.email||'').trim().toLowerCase();
    var rowIndex=-1;
    for(var i=1;i<data.length;i++){
      if(String(data[i][idx.Business_ID]||'')===String(actor.user.businessId||'')&&String(data[i][idx.Email]||'').trim().toLowerCase()===email){rowIndex=i;break;}
    }
    var p=preferences||{};
    var values={
      Preference_ID:'NP_'+Utilities.getUuid().replace(/-/g,'').substring(0,20).toUpperCase(),
      Business_ID:String(actor.user.businessId||''),
      Email:email,
      Feature_Updates:p.features!==false,
      System_Alerts:p.system!==false,
      Weekly_Reports:p.weekly===true,
      Security_Alerts:p.security!==false,
      Marketing_Updates:p.marketing===true,
      Updated_At:new Date().toISOString()
    };
    if(rowIndex===-1){
      var newRow=new Array(headers.length).fill('');
      headers.forEach(function(h,j){if(Object.prototype.hasOwnProperty.call(values,h))newRow[j]=values[h];});
      sheet.appendRow(newRow);
    }else{
      headers.forEach(function(h,j){if(Object.prototype.hasOwnProperty.call(values,h))sheet.getRange(rowIndex+1,j+1).setValue(values[h]);});
    }
    return {success:true,preferences:{features:values.Feature_Updates,system:values.System_Alerts,weekly:values.Weekly_Reports,security:values.Security_Alerts,marketing:values.Marketing_Updates}};
  }catch(error){
    console.error('saveBizOSNotificationPreferences error:',error);
    return {success:false,message:'Could not save notification preferences.'};
  }
}
