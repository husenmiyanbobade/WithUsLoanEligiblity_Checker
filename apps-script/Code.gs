// WITH US Society - shared data store. Paste into script.google.com, change PASSCODE, then Deploy > New deployment > Web app
// (Execute as: Me, Who has access: Anyone). Copy the Web app URL into SYNC_URL in index.html.
var PASSCODE = 'CHANGE-THIS-PASSCODE';
var FILE_NAME = 'withus-dataset.json';

function out_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function file_(){ var it = DriveApp.getFilesByName(FILE_NAME); return it.hasNext() ? it.next() : null; }

function doGet(e){
  var f = file_();
  if(!f) return out_({ empty:true });
  var saved = JSON.parse(f.getBlob().getDataAsString());
  var since = Number((e && e.parameter && e.parameter.since) || 0);
  if(since && since === saved.updatedAt) return out_({ unchanged:true, updatedAt:saved.updatedAt });
  return out_(saved);
}

function doPost(e){
  try{
    var body = JSON.parse(e.postData.contents);
    if(body.passcode !== PASSCODE) return out_({ ok:false, error:'wrong-passcode' });
    var saved = { updatedAt: Date.now(), dataset: body.dataset };
    var txt = JSON.stringify(saved), f = file_();
    if(f) f.setContent(txt); else DriveApp.createFile(FILE_NAME, txt, MimeType.PLAIN_TEXT);
    return out_({ ok:true, updatedAt:saved.updatedAt });
  }catch(err){ return out_({ ok:false, error:String(err) }); }
}
