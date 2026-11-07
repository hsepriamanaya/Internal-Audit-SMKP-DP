// Backend: Google Sheets sebagai DB. Sheet "Data" (key,value,updated). Script Properties: ADMIN_KEY, READER_KEY, AUDITEE_KEY
function _json(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function _role(k){var p=PropertiesService.getScriptProperties();
  if(k&&k===p.getProperty('ADMIN_KEY'))return 'admin';
  if(k&&k===p.getProperty('READER_KEY'))return 'reader';
  if(k&&k===p.getProperty('AUDITEE_KEY'))return 'auditee'; return null;}
// auditee hanya boleh menulis key berakhiran /tl_auditi
function _sheet(){var ss=SpreadsheetApp.getActiveSpreadsheet();var s=ss.getSheetByName('Data');
  if(!s){s=ss.insertSheet('Data');s.appendRow(['key','value','updated']);}return s;}
function _find(s,key){var v=s.getRange(1,1,Math.max(s.getLastRow(),1),1).getValues();
  for(var i=1;i<v.length;i++) if(v[i][0]===key) return i+1; return 0;}
function doPost(e){
  var b=JSON.parse(e.postData.contents), role=_role(b.k);
  if(!role) return _json({ok:false,err:'auth'});
  var lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    var s=_sheet();
    if(b.a==='get'){var r=_find(s,b.key); return _json({ok:true,value:r?s.getRange(r,2).getValue():null,role:role});}
    if(b.a==='list'){var v=s.getDataRange().getValues().slice(1).map(function(x){return x[0];});return _json({ok:true,keys:v,role:role});}
    if(b.a==='set'){
      if(role==='reader') return _json({ok:false,err:'readonly'});
      if(role==='auditee' && !/\/tl_auditi$/.test(b.key)) return _json({ok:false,err:'forbidden'});
      // audit terkunci: hanya tindak lanjut/arsip/index yang boleh ditulis
      var pre=b.key.split('/')[0];
      if(b.key!=='index' && pre!=='snap' && !/\/tl_(auditi|verif)$/.test(b.key)){
        var ri=_find(s,'index'); if(ri){var ix=JSON.parse(s.getRange(ri,2).getValue()||'[]');
          for(var i=0;i<ix.length;i++) if(ix[i].id===pre && ix[i].locked) return _json({ok:false,err:'locked'});}}
      if(b.key==='index' && role!=='admin') return _json({ok:false,err:'forbidden'});
      var r2=_find(s,b.key), val=JSON.stringify(b.value);
      if(val.length>45000) return _json({ok:false,err:'too_big'}); // batas sel; pecah per key
      if(r2) s.getRange(r2,2,1,2).setValues([[val,new Date()]]); else s.appendRow([b.key,val,new Date()]);
      return _json({ok:true});
    }
    return _json({ok:false,err:'action'});
  } finally {lock.releaseLock();}
}
function doGet(){return _json({ok:true,svc:'smkp'});}
