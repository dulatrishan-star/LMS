/* AVSEC LMS Google Sheet Backend
Deploy: Extensions > Apps Script > Deploy > New deployment > Web app
Execute as: Me | Who has access: Anyone (or Anyone with link)
Paste the Web App URL into js/config.js APPS_SCRIPT_URL.
*/
function doGet(e){return json_({ok:true,message:'AVSEC LMS API running'});}
function doPost(e){
  try{
    var body = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : '{}');
    if(body.action === 'saveAll') return saveAll_(body.state || {});
    if(body.action === 'loadAll') return loadAll_();
    return json_({ok:false,error:'Unknown action'});
  }catch(err){return json_({ok:false,error:String(err)});}
}
function ss_(){return SpreadsheetApp.getActiveSpreadsheet();}
function sheet_(name){var sh=ss_().getSheetByName(name); return sh || ss_().insertSheet(name);}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
function setRows_(name, headers, rows){
  var sh=sheet_(name); sh.clearContents(); sh.getRange(1,1,1,headers.length).setValues([headers]);
  if(rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows);
  sh.setFrozenRows(1);
}
function saveAll_(state){
  sheet_('_STATE').clearContents(); sheet_('_STATE').getRange(1,1).setValue(JSON.stringify(state));
  var depts = ['Engineering','Utility','Ramp','AIS','Cabin Crew','Security'];
  setRows_('Users',['username','password','displayName','pages'],(state.users||[]).map(function(u){return [u.username,u.password,u.displayName,(u.pages||[]).join(', ')];}));
  setRows_('Live_Class',['title','link','note'],[[state.live && state.live.title || '', state.live && state.live.link || '', state.live && state.live.note || '']]);
  var videos=[]; depts.forEach(function(d){(state.videos&&state.videos[d]||[]).forEach(function(v){videos.push([d,v.title,v.link,v.date]);});});
  setRows_('Drive_Videos',['department','title','driveLink','date'],videos);
  var q=[]; depts.forEach(function(d){(state.qbank&&state.qbank[d]||[]).forEach(function(x){q.push([d,x.id,x.question,x.A,x.B,x.C,x.D,x.answer]);});});
  setRows_('QBank',['department','id','question','A','B','C','D','answer_admin_only'],q);
  var papers=[]; depts.forEach(function(d){(state.papers&&state.papers[d]||[]).forEach(function(p){papers.push([d,p.id,p.title,p.created,(p.questions||[]).join('|')]);});});
  setRows_('Papers',['department','paperId','title','created','questionIds'],papers);
  setRows_('Paper_Attempts',['attemptId','username','name','department','paperId','paperTitle','total','correct','score','date','answers_json'],(state.attempts||[]).map(function(a){return [a.id,a.user,a.name,a.dept,a.paperId,a.paperTitle,a.total,a.correct,a.score,a.date,JSON.stringify(a.answers||[])];}));
  setRows_('Video_Views',['username','name','department','videoTitle','date'],(state.videoViews||[]).map(function(v){return [v.user,v.name,v.dept,v.videoTitle,v.date];}));
  setRows_('Settings',['key','value'],Object.keys(state.settings||{}).map(function(k){return [k,state.settings[k]];}));
  return json_({ok:true,savedAt:new Date().toISOString()});
}
function loadAll_(){
  var sh=ss_().getSheetByName('_STATE');
  if(!sh || !sh.getRange(1,1).getValue()) return json_({ok:true,state:null});
  return json_({ok:true,state:JSON.parse(sh.getRange(1,1).getValue())});
}
