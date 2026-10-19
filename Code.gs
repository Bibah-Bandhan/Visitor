/************************************************************
 * BIBAH BANDHAN — VISITOR REGISTER (2-SHEET DATABASE)
 * Visitors = one row per person; Visits = one row per visit.
 ************************************************************/
const APP_CONFIG = {
  APP_NAME: 'Bibah Bandhan Visitor Register',
  SPREADSHEET_ID: '1HMo8H-UUYSwDD6NoYFmTWpRFaecUtabdjBGcsuejLeY',
  PHOTO_FOLDER_ID: '1QjxrHEuMnZgQA-UaKTpoHH8rropHaUIF',
  ADMIN_USERNAME: 'admin',
  ADMIN_PASSWORD: '2594', // CHANGE before deployment
  SESSION_MINUTES: 3,
  MAX_PHOTO_BYTES: 550 * 1024
};
const VISITOR_HEADERS = ['VisitorID','CreatedAt','VisitorName','Address','Mobile','PhotoURL','PhotoFileID','UpdatedAt'];
const VISIT_HEADERS = ['VisitID','VisitorID','VisitDateTime','Purpose','Visitor2Name','Visitor2Mobile','VisitNotes','UpdatedAt','CreatedBy','FollowUpDate','FollowUpStatus'];
const PROSPECT_HEADERS = ['ProspectID','CreatedAt','Name','Mobile','Address','Gender','ForWhom','Source','Notes','FollowUpDate','PhotoURL','PhotoFileID','UpdatedAt'];
const BIODATA_HEADERS = ['BiodataID','CreatedAt','UpdatedAt','Gender','FullName','DateOfBirth','Age','Height','MaritalStatus','Religion','Education','Occupation','Income','District','Mobile','Address','PartnerPreference','Notes','Status'];
const AUDIT_HEADERS = ['Timestamp','Username','Action','RecordType','RecordID','Details'];
const LEGACY_HEADERS = ['ID','CreatedAt','VisitorName','Address','Mobile','Purpose','Visitor2Name','Visitor2Mobile','PhotoURL','PhotoFileID','UpdatedAt','VisitNotes'];

function setup() {
  const ss = SpreadsheetApp.openById(APP_CONFIG.SPREADSHEET_ID);
  let people = ss.getSheetByName('Visitors');
  let visits = ss.getSheetByName('Visits');
  let prospects = ss.getSheetByName('Prospects');
  let biodata = ss.getSheetByName('Biodata');
  let audit = ss.getSheetByName('ActivityLog');
  if (people && !headersMatch_(people, VISITOR_HEADERS)) {
    // Preserve old data as a backup before creating the normalized tables.
    const name = 'Legacy_Visitors_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    people.setName(name);
    migrateLegacy_(ss, ss.getSheetByName(name));
    people = ss.getSheetByName('Visitors');
    visits = ss.getSheetByName('Visits');
  }
  if (!people) people = ss.insertSheet('Visitors');
  if (!visits) visits = ss.insertSheet('Visits');
  ensureHeaders_(people, VISITOR_HEADERS);
  ensureHeaders_(visits, VISIT_HEADERS);
  if (!prospects) prospects = ss.insertSheet('Prospects');
  ensureHeaders_(prospects, PROSPECT_HEADERS);
  if (!biodata) biodata = ss.insertSheet('Biodata');
  if (!audit) audit = ss.insertSheet('ActivityLog');
  ensureHeaders_(biodata, BIODATA_HEADERS); ensureHeaders_(audit, AUDIT_HEADERS);
  styleHeader_(people, VISITOR_HEADERS.length); styleHeader_(visits, VISIT_HEADERS.length);
  styleHeader_(prospects, PROSPECT_HEADERS.length); styleHeader_(biodata, BIODATA_HEADERS.length); styleHeader_(audit, AUDIT_HEADERS.length);
  // Fail early if the configured folder is inaccessible.
  DriveApp.getFolderById(APP_CONFIG.PHOTO_FOLDER_ID);
  return {ok:true, spreadsheetUrl:ss.getUrl(), sheets:['Visitors','Visits','Prospects','Biodata','ActivityLog']};
}
function headersMatch_(sheet, headers) {
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < headers.length) return false;
  const row = sheet.getRange(1,1,1,headers.length).getDisplayValues()[0];
  return headers.every((h,i) => String(row[i] || '').trim() === h);
}
function ensureHeaders_(sheet, headers) {
  if (!headersMatch_(sheet, headers)) sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}
function styleHeader_(sheet, count) {
  sheet.getRange(1,1,1,count).setFontWeight('bold').setBackground('#153b70').setFontColor('#ffffff');
}
function migrateLegacy_(ss, legacy) {
  const people = ss.insertSheet('Visitors');
  const visits = ss.insertSheet('Visits');
  ensureHeaders_(people, VISITOR_HEADERS); ensureHeaders_(visits, VISIT_HEADERS);
  if (!legacy || legacy.getLastRow() < 2) return;
  const values = legacy.getRange(2,1,legacy.getLastRow()-1,Math.min(legacy.getLastColumn(), LEGACY_HEADERS.length)).getValues();
  const personRows = [], visitRows = [];
  values.forEach(r => {
    if (!String(r[2] || '').trim()) return;
    const personId = Utilities.getUuid();
    const visitId = String(r[0] || Utilities.getUuid());
    const created = r[1] instanceof Date ? r[1] : (r[1] ? new Date(r[1]) : new Date());
    const photoId = String(r[9] || extractDriveId_(r[8]) || '');
    const photoUrl = photoId ? photoThumbUrl_(photoId) : String(r[8] || '');
    let mobile = ''; try { mobile = normalizeIndianMobile_(r[4]); } catch (_) { mobile = String(r[4] || ''); }
    let mobile2 = ''; try { mobile2 = normalizeIndianMobile_(r[7]); } catch (_) { mobile2 = String(r[7] || ''); }
    personRows.push([personId, created, String(r[2]||''), String(r[3]||''), mobile, photoUrl, photoId, r[10] || created]);
    visitRows.push([visitId, personId, created, String(r[5]||''), String(r[6]||''), mobile2, String(r[11]||''), r[10] || created, APP_CONFIG.ADMIN_USERNAME, '', 'Pending']);
  });
  if (personRows.length) people.getRange(2,1,personRows.length,VISITOR_HEADERS.length).setValues(personRows);
  if (visitRows.length) visits.getRange(2,1,visitRows.length,VISIT_HEADERS.length).setValues(visitRows);
  styleHeader_(people,VISITOR_HEADERS.length); styleHeader_(visits,VISIT_HEADERS.length);
}
function getSS_(){ return SpreadsheetApp.openById(APP_CONFIG.SPREADSHEET_ID); }
function getSheet_(name, headers) {
  const ss=getSS_(); let s=ss.getSheetByName(name);
  if(!s){ setup(); s=ss.getSheetByName(name); }
  if(!s) throw new Error(name+' sheet পাওয়া যায়নি। setup() চালান।');
  ensureHeaders_(s,headers); return s;
}
function doGet() {
  return HtmlService.createHtmlOutput('<!doctype html><meta name="viewport" content="width=device-width,initial-scale=1"><body style="font:16px Arial;padding:24px"><h2>Bibah Bandhan Visitor Register API</h2><p>API is running. Run setup() once from the Apps Script editor, then deploy as Web App.</p></body>').setTitle(APP_CONFIG.APP_NAME);
}
function doPost(e) {
  try {
    const p=JSON.parse((e&&e.postData&&e.postData.contents)||'{}');
    const methods={login,logout,getVisitors,keepAlive,saveVisitor,deleteVisitor,updateVisitFollowup,getProspects,saveProspect,deleteProspect,getBiodata,saveBiodata,deleteBiodata,backupNow,getActivityLog};
    if(!Object.prototype.hasOwnProperty.call(methods,p.method)) throw new Error('অজানা API method।');
    // Initialize schema on first call; setup is idempotent and preserves legacy sheet.
    if(p.method!=='login' && p.method!=='logout') setup();
    const args=Array.isArray(p.args)?p.args:[];
    const data=methods[p.method].apply(null,args);
    return ContentService.createTextOutput(JSON.stringify({ok:true,data})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err&&err.message?err.message:err)})).setMimeType(ContentService.MimeType.JSON);
  }
}
function login(username,password){
  if(String(username||'')!==APP_CONFIG.ADMIN_USERNAME || String(password||'')!==APP_CONFIG.ADMIN_PASSWORD) throw new Error('Username অথবা password ভুল।');
  const token=Utilities.getUuid()+Utilities.getUuid();
  CacheService.getScriptCache().put('session_'+token,APP_CONFIG.ADMIN_USERNAME,APP_CONFIG.SESSION_MINUTES*60);
  return {ok:true,token,username:APP_CONFIG.ADMIN_USERNAME,expiresInMinutes:APP_CONFIG.SESSION_MINUTES};
}
function requireAuth_(token){
  const key='session_'+String(token||''); const user=CacheService.getScriptCache().get(key);
  if(!user) throw new Error('Session শেষ হয়েছে। আবার লগইন করুন।');
  CacheService.getScriptCache().put(key,user,APP_CONFIG.SESSION_MINUTES*60); return user;
}
function keepAlive(token){requireAuth_(token);return {ok:true,expiresInMinutes:APP_CONFIG.SESSION_MINUTES};}
function logout(token){CacheService.getScriptCache().remove('session_'+String(token||''));return {ok:true};}
function normalizeIndianMobile_(value){
  const raw=String(value||'').trim(); if(!raw)return '';
  let d=raw.replace(/\D/g,'');
  if(d.length===12&&d.startsWith('91'))d=d.slice(2); else if(d.length===11&&d.startsWith('0'))d=d.slice(1);
  if(!/^[6-9]\d{9}$/.test(d)) throw new Error('মোবাইল নম্বরটি সঠিক ১০ সংখ্যার ভারতীয় নম্বর দিন (৬–৯ দিয়ে শুরু)।');
  return d;
}
function rowsAsObjects_(sheet, headers){
  if(sheet.getLastRow()<2)return [];
  return sheet.getRange(2,1,sheet.getLastRow()-1,headers.length).getValues().map(r=>{const o={};headers.forEach((h,i)=>o[h]=r[i]);return o;});
}
function toIso_(v){ if(!v)return ''; if(v instanceof Date)return v.toISOString(); const d=new Date(v); return isNaN(d.getTime())?String(v):d.toISOString(); }
function extractDriveId_(url){const s=String(url||'');let m=s.match(/[?&]id=([\w-]+)/)||s.match(/\/d\/([\w-]+)/);return m?m[1]:'';}
function photoThumbUrl_(id){return id?'https://drive.google.com/thumbnail?id='+encodeURIComponent(id)+'&sz=w1000':'';}
function photoUrlFor_(person){const id=String(person.PhotoFileID||extractDriveId_(person.PhotoURL)||'');return id?photoThumbUrl_(id):String(person.PhotoURL||'');}
function getVisitors(token){
  requireAuth_(token);
  const people=rowsAsObjects_(getSheet_('Visitors',VISITOR_HEADERS),VISITOR_HEADERS);
  const visits=rowsAsObjects_(getSheet_('Visits',VISIT_HEADERS),VISIT_HEADERS);
  const byId={}; people.forEach(p=>byId[String(p.VisitorID)] = p);
  const out=visits.map(v=>{const p=byId[String(v.VisitorID)]||{};return {
    id:String(v.VisitID||''), personId:String(v.VisitorID||''), createdAt:toIso_(v.VisitDateTime),
    name:String(p.VisitorName||''),address:String(p.Address||''),mobile:String(p.Mobile||''),
    purpose:String(v.Purpose||''),visitor2Name:String(v.Visitor2Name||''),visitor2Mobile:String(v.Visitor2Mobile||''),
    photoUrl:photoUrlFor_(p),photoFileId:String(p.PhotoFileID||''),updatedAt:toIso_(v.UpdatedAt||p.UpdatedAt),visitNotes:String(v.VisitNotes||''),followUpDate:toIso_(v.FollowUpDate),followUpStatus:String(v.FollowUpStatus||'Pending')
  };}).sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt));
  return {visitors:out,stats:makeStats_(out)};
}
function getProspects(token){
  requireAuth_(token);
  return rowsAsObjects_(getSheet_('Prospects',PROSPECT_HEADERS),PROSPECT_HEADERS).map(r=>({
    id:String(r.ProspectID||''),createdAt:toIso_(r.CreatedAt),name:String(r.Name||''),mobile:String(r.Mobile||''),
    address:String(r.Address||''),gender:String(r.Gender||''),forWhom:String(r.ForWhom||''),source:String(r.Source||''),
    notes:String(r.Notes||''),followUpDate:toIso_(r.FollowUpDate),photoUrl:photoUrlFor_({PhotoURL:r.PhotoURL,PhotoFileID:r.PhotoFileID}),
    photoFileId:String(r.PhotoFileID||''),updatedAt:toIso_(r.UpdatedAt)
  })).sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt));
}
function saveProspect(token,d){
  const user=requireAuth_(token); if(!d||!String(d.name||'').trim())throw new Error('Prospect-এর নাম লিখুন।');
  const name=String(d.name).trim(),mobile=normalizeIndianMobile_(d.mobile),address=String(d.address||'').trim();
  const gender=String(d.gender||''),forWhom=String(d.forWhom||''),source=String(d.source||'').trim(),notes=String(d.notes||'').trim();
  if(!['Male','Female','Other',''].includes(gender))throw new Error('ছেলে/মেয়ে নির্বাচন সঠিক নয়।');
  if(name.length>150||address.length>1000||source.length>250||notes.length>2000)throw new Error('কোনও ফিল্ডে অনুমোদিত সীমার বেশি লেখা আছে।');
  const s=getSheet_('Prospects',PROSPECT_HEADERS),rows=rowsAsObjects_(s,PROSPECT_HEADERS),id=String(d.id||'').trim(),old=id?rows.find(r=>String(r.ProspectID)===id):null;
  if(id&&!old)throw new Error('Prospect রেকর্ড পাওয়া যায়নি। তালিকা রিফ্রেশ করুন।');
  const now=new Date();let photoId=old?String(old.PhotoFileID||''):'',photoUrl=old?photoUrlFor_(old):'';
  if(d.photoData){const saved=savePhoto_(d.photoData,name);const previous=photoId;photoId=saved.id;photoUrl=saved.url;if(previous&&previous!==photoId){try{DriveApp.getFileById(previous).setTrashed(true)}catch(e){}}}else if(d.removePhoto&&photoId){try{DriveApp.getFileById(photoId).setTrashed(true)}catch(e){}photoId='';photoUrl='';}
  const obj={ProspectID:id||Utilities.getUuid(),CreatedAt:old?old.CreatedAt:now,Name:name,Mobile:mobile,Address:address,Gender:gender,ForWhom:forWhom,Source:source,Notes:notes,FollowUpDate:d.followUpDate?new Date(d.followUpDate):(old?old.FollowUpDate:''),PhotoURL:photoUrl,PhotoFileID:photoId,UpdatedAt:now};
  const values=PROSPECT_HEADERS.map(k=>obj[k]===undefined?'':obj[k]);
  if(old){const i=rows.findIndex(r=>String(r.ProspectID)===id);s.getRange(i+2,1,1,PROSPECT_HEADERS.length).setValues([values]);}else{s.appendRow(values);}
  audit_(user,old?'UPDATE':'CREATE','Prospect',obj.ProspectID,name+' / '+mobile);
  return {ok:true,id:obj.ProspectID,message:old?'Prospect আপডেট হয়েছে।':'নতুন Prospect সেভ হয়েছে।'};
}
function deleteProspect(token,id){
  const user=requireAuth_(token);if(!id)throw new Error('Prospect ID পাওয়া যায়নি।');
  const s=getSheet_('Prospects',PROSPECT_HEADERS),rows=rowsAsObjects_(s,PROSPECT_HEADERS),i=rows.findIndex(r=>String(r.ProspectID)===String(id));
  if(i<0)throw new Error('Prospect রেকর্ড পাওয়া যায়নি।');const row=rows[i];s.deleteRow(i+2);
  const photo=String(row.PhotoFileID||'');if(photo){try{DriveApp.getFileById(photo).setTrashed(true)}catch(e){}}
  audit_(user,'DELETE','Prospect',String(id),String(row.Name||''));return {ok:true,message:'Prospect রেকর্ড মুছে দেওয়া হয়েছে।'};
}

function makeStats_(arr){
  const today=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  const isToday=v=>v.createdAt&&Utilities.formatDate(new Date(v.createdAt),Session.getScriptTimeZone(),'yyyy-MM-dd')===today;
  return {total:arr.length,today:arr.filter(isToday).length,uniqueVisitors:new Set(arr.map(v=>v.personId||v.mobile||v.id)).size};
}
function savePhoto_(data,name){
  const m=String(data||'').match(/^data:image\/(jpeg|jpg|png|webp);base64,([A-Za-z0-9+/=]+)$/);
  if(!m)throw new Error('ছবির ফরম্যাট সঠিক নয়। JPG, PNG অথবা WebP ব্যবহার করুন।');
  const bytes=Utilities.base64Decode(m[2]); if(bytes.length>APP_CONFIG.MAX_PHOTO_BYTES)throw new Error('ছবির সাইজ ৫৫০ KB-এর বেশি। ছোট ছবি আপলোড করুন।');
  const mime=m[1]==='jpg'?'image/jpeg':'image/'+m[1];
  const safe=String(name||'visitor').replace(/[^\w\u0980-\u09FF-]/g,'_').substring(0,60);
  const file=getPhotoFolder_().createFile(Utilities.newBlob(bytes,mime,safe+'_'+Date.now()+'.'+(m[1]==='jpeg'?'jpg':m[1])));
  // Drive thumbnail embeds require link viewing; avoid changing old file permissions here.
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
  return {id:file.getId(),url:photoThumbUrl_(file.getId())};
}
function getPhotoFolder_(){return DriveApp.getFolderById(APP_CONFIG.PHOTO_FOLDER_ID);}
function saveVisitor(token,data){
  const user=requireAuth_(token); if(!data||!String(data.name||'').trim())throw new Error('ভিজিটরের নাম লিখুন।');
  const name=String(data.name).trim(), address=String(data.address||'').trim(), mobile=normalizeIndianMobile_(data.mobile),
    purpose=String(data.purpose||'').trim(), v2Name=String(data.visitor2Name||'').trim(), v2Mobile=normalizeIndianMobile_(data.visitor2Mobile), notes=String(data.visitNotes||'').trim();
  if(name.length>150||address.length>1000||purpose.length>1000||notes.length>2000)throw new Error('কোনও একটি ফিল্ডে অনুমোদিত সীমার বেশি লেখা আছে।');
  const ss=getSS_(), people=getSheet_('Visitors',VISITOR_HEADERS), visits=getSheet_('Visits',VISIT_HEADERS), lock=LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const now=new Date(), visitId=String(data.id||'').trim(), peopleRows=rowsAsObjects_(people,VISITOR_HEADERS), visitRows=rowsAsObjects_(visits,VISIT_HEADERS);
    let visit=null, person=null;
    if(visitId){visit=visitRows.find(v=>String(v.VisitID)===visitId);if(!visit)throw new Error('রেকর্ডটি পাওয়া যায়নি। তালিকা রিফ্রেশ করুন।');person=peopleRows.find(p=>String(p.VisitorID)===String(visit.VisitorID));if(!person)throw new Error('ভিজিটর প্রোফাইল পাওয়া যায়নি।');}
    else if(data.copyFromId){const src=visitRows.find(v=>String(v.VisitID)===String(data.copyFromId));if(!src)throw new Error('আগের ভিজিটর রেকর্ডটি পাওয়া যায়নি।');person=peopleRows.find(p=>String(p.VisitorID)===String(src.VisitorID));if(!person)throw new Error('আগের ভিজিটর প্রোফাইল পাওয়া যায়নি।');}
    let personId=person?String(person.VisitorID):Utilities.getUuid(); let photoId=person?String(person.PhotoFileID||''):''; let photoUrl=person?photoUrlFor_(person):'';
    if(data.photoData){const saved=savePhoto_(data.photoData,name);const oldPhoto=photoId;photoId=saved.id;photoUrl=saved.url;if(oldPhoto&&oldPhoto!==photoId){try{DriveApp.getFileById(oldPhoto).setTrashed(true);}catch(e){console.warn('Old photo cleanup: '+e.message);}}}
    const personRow=[personId,person?person.CreatedAt:now,name,address,mobile,photoUrl,photoId,now];
    if(person){const idx=peopleRows.findIndex(p=>String(p.VisitorID)===personId);people.getRange(idx+2,1,1,VISITOR_HEADERS.length).setValues([personRow]);}
    else people.appendRow(personRow);
    if(visitId){const idx=visitRows.findIndex(v=>String(v.VisitID)===visitId);const old=visitRows[idx];visits.getRange(idx+2,1,1,VISIT_HEADERS.length).setValues([[visitId,personId,old.VisitDateTime,purpose,v2Name,v2Mobile,notes,now,user,data.followUpDate!==undefined?(data.followUpDate?new Date(data.followUpDate):''):old.FollowUpDate,(data.followUpStatus||old.FollowUpStatus||'Pending')]]);}
    else visits.appendRow([Utilities.getUuid(),personId,now,purpose,v2Name,v2Mobile,notes,now,user,data.followUpDate?new Date(data.followUpDate):'',data.followUpStatus||'Pending']);
    audit_(user,visitId?'UPDATE':'CREATE','VisitorVisit',visitId||'new',name+' / '+mobile);
    return {ok:true,id:visitId||'created',personId,message:visitId?'রেকর্ড আপডেট হয়েছে।':'নতুন ভিজিট সেভ হয়েছে।'};
  } finally {lock.releaseLock();}
}
function deleteVisitor(token,id){
  const user=requireAuth_(token);if(!id)throw new Error('রেকর্ড ID পাওয়া যায়নি।');
  const people=getSheet_('Visitors',VISITOR_HEADERS),visits=getSheet_('Visits',VISIT_HEADERS),lock=LockService.getScriptLock();lock.waitLock(20000);
  try {
    const vs=rowsAsObjects_(visits,VISIT_HEADERS),idx=vs.findIndex(v=>String(v.VisitID)===String(id));if(idx<0)throw new Error('রেকর্ডটি পাওয়া যায়নি।');
    const personId=String(vs[idx].VisitorID);visits.deleteRow(idx+2);
    const remaining=rowsAsObjects_(visits,VISIT_HEADERS).some(v=>String(v.VisitorID)===personId);
    if(!remaining){const ps=rowsAsObjects_(people,VISITOR_HEADERS),pi=ps.findIndex(p=>String(p.VisitorID)===personId);if(pi>=0){const photo=String(ps[pi].PhotoFileID||'');people.deleteRow(pi+2);if(photo){try{DriveApp.getFileById(photo).setTrashed(true);}catch(e){}}}}
    audit_(user,'DELETE','VisitorVisit',String(id),'VisitorID: '+personId);
    return {ok:true,message:'ভিজিট রেকর্ড মুছে দেওয়া হয়েছে।'};
  } finally {lock.releaseLock();}
}


/* === New modules: Follow-up, Biodata, audit trail and Drive backup === */
function audit_(username,action,type,id,details){try{getSheet_('ActivityLog',AUDIT_HEADERS).appendRow([new Date(),String(username||'admin'),String(action||''),String(type||''),String(id||''),String(details||'').slice(0,1000)]);}catch(e){console.warn(e.message);}}
function updateVisitFollowup(token,id,date,status){const user=requireAuth_(token);const allowed=['Pending','Completed','Overdue','Not required'];status=String(status||'Pending');if(!allowed.includes(status))throw new Error('Follow-up status সঠিক নয়।');const s=getSheet_('Visits',VISIT_HEADERS),r=rowsAsObjects_(s,VISIT_HEADERS),i=r.findIndex(x=>String(x.VisitID)===String(id));if(i<0)throw new Error('ভিজিট রেকর্ড পাওয়া যায়নি।');r[i].FollowUpDate=date?new Date(date):'';r[i].FollowUpStatus=status;r[i].UpdatedAt=new Date();s.getRange(i+2,1,1,VISIT_HEADERS.length).setValues([VISIT_HEADERS.map(k=>r[i][k]===undefined?'':r[i][k])]);audit_(user,'FOLLOWUP_UPDATE','VisitorVisit',id,status);return {ok:true,message:'Follow-up আপডেট হয়েছে।'};}
function getBiodata(token){requireAuth_(token);return rowsAsObjects_(getSheet_('Biodata',BIODATA_HEADERS),BIODATA_HEADERS).map(r=>({id:String(r.BiodataID||''),createdAt:toIso_(r.CreatedAt),updatedAt:toIso_(r.UpdatedAt),gender:String(r.Gender||''),fullName:String(r.FullName||''),dob:toIso_(r.DateOfBirth),age:String(r.Age||''),height:String(r.Height||''),maritalStatus:String(r.MaritalStatus||''),religion:String(r.Religion||''),education:String(r.Education||''),occupation:String(r.Occupation||''),income:String(r.Income||''),district:String(r.District||''),mobile:String(r.Mobile||''),address:String(r.Address||''),partnerPreference:String(r.PartnerPreference||''),notes:String(r.Notes||''),status:String(r.Status||'Active')})).sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt));}
function saveBiodata(token,d){const user=requireAuth_(token);if(!d||!String(d.fullName||'').trim())throw new Error('পাত্র/পাত্রীর নাম লিখুন।');if(!['Bride','Groom','Other'].includes(String(d.gender||'')))throw new Error('প্রোফাইল নির্বাচন করুন।');const age=Number(d.age||0);if(age&&(age<18||age>100))throw new Error('বয়স ১৮ থেকে ১০০-এর মধ্যে হতে হবে।');const mobile=normalizeIndianMobile_(d.mobile),id=String(d.id||''),s=getSheet_('Biodata',BIODATA_HEADERS),rows=rowsAsObjects_(s,BIODATA_HEADERS),old=id?rows.find(r=>String(r.BiodataID)===id):null;if(id&&!old)throw new Error('Biodata record পাওয়া যায়নি।');const now=new Date(),o={BiodataID:id||Utilities.getUuid(),CreatedAt:old?old.CreatedAt:now,UpdatedAt:now,Gender:d.gender,FullName:String(d.fullName).trim(),DateOfBirth:d.dob?new Date(d.dob):'',Age:age||'',Height:String(d.height||'').slice(0,30),MaritalStatus:String(d.maritalStatus||'').slice(0,50),Religion:String(d.religion||'').slice(0,80),Education:String(d.education||'').slice(0,200),Occupation:String(d.occupation||'').slice(0,200),Income:String(d.income||'').slice(0,100),District:String(d.district||'').slice(0,150),Mobile:mobile,Address:String(d.address||'').slice(0,1000),PartnerPreference:String(d.partnerPreference||'').slice(0,1500),Notes:String(d.notes||'').slice(0,1500),Status:['Active','On Hold','Matched','Inactive'].includes(String(d.status||''))?d.status:'Active'};const vals=BIODATA_HEADERS.map(k=>o[k]===undefined?'':o[k]);if(old){const i=rows.findIndex(r=>String(r.BiodataID)===id);s.getRange(i+2,1,1,BIODATA_HEADERS.length).setValues([vals]);}else s.appendRow(vals);audit_(user,id?'UPDATE':'CREATE','Biodata',o.BiodataID,o.FullName);return {ok:true,id:o.BiodataID,message:id?'Biodata আপডেট হয়েছে।':'Biodata সেভ হয়েছে।'};}
function deleteBiodata(token,id){const user=requireAuth_(token);const s=getSheet_('Biodata',BIODATA_HEADERS),r=rowsAsObjects_(s,BIODATA_HEADERS),i=r.findIndex(x=>String(x.BiodataID)===String(id));if(i<0)throw new Error('Biodata record পাওয়া যায়নি।');const name=r[i].FullName;s.deleteRow(i+2);audit_(user,'DELETE','Biodata',id,name);return {ok:true,message:'Biodata মুছে দেওয়া হয়েছে।'};}
function getActivityLog(token){requireAuth_(token);const s=getSheet_('ActivityLog',AUDIT_HEADERS),n=s.getLastRow();if(n<2)return [];return s.getRange(Math.max(2,n-199),1,Math.min(200,n-1),AUDIT_HEADERS.length).getDisplayValues().reverse().map(r=>({timestamp:r[0],username:r[1],action:r[2],recordType:r[3],recordId:r[4],details:r[5]}));}
function backupNow(token){const user=requireAuth_(token),file=DriveApp.getFileById(APP_CONFIG.SPREADSHEET_ID),stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss'),copy=file.makeCopy('Bibah_Bandhan_Backup_'+stamp);audit_(user,'BACKUP','Spreadsheet',copy.getId(),'Spreadsheet backup created');return {ok:true,url:copy.getUrl(),name:copy.getName(),message:'Backup spreadsheet তৈরি হয়েছে।'};}
