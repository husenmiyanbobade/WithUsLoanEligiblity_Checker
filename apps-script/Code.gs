// WITH US Society - shared data store with admin login.
// Paste into script.google.com (replace the old code), then
// Deploy > Manage deployments > pencil icon > Version: "New version" > Deploy.  (The Web app URL stays the same.)
// Execute as: Me, Who has access: Anyone.
//
// The username and password are checked HERE, never inside the web page.
// To change them later, edit the two lines below and deploy a new version.
var ADMIN_USER = 'withusadmin';
var ADMIN_PASS = 'withusadmin';
var MAX_FAILS = 5;          // wrong tries allowed per device
var MAX_FAILS_ALL = 40;     // safety net: wrong tries from ALL devices together before everyone pauses
var LOCK_MIN = 15;          // minutes the login stays paused after that
var MAX_SESSIONS = 30;      // phones that can be signed in at once (oldest is signed out first)

var FILE_NAME = 'withus-dataset.json';              // Excel dataset
var CONTRIB_FILE = 'withus-contributions.json';     // monthly contribution marks
var REPAY_FILE = 'withus-repayments.json';          // monthly loan repayment marks

function out_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function file_(name){ var it = DriveApp.getFilesByName(name); return it.hasNext() ? it.next() : null; }

/* ---------- sessions ---------- */
function sessions_(){
  var raw = PropertiesService.getScriptProperties().getProperty('sessions');
  try{ return raw ? JSON.parse(raw) : {}; }catch(e){ return {}; }
}
function saveSessions_(s){ PropertiesService.getScriptProperties().setProperty('sessions', JSON.stringify(s)); }
function validToken_(t){ return typeof t === 'string' && t.length >= 32 && !!sessions_()[t]; }

function login_(body){
  var cache = CacheService.getScriptCache();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    // The pause is per DEVICE: each phone sends its own random id, so one phone's wrong tries never block another phone.
    var dev = String(body.dev || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 64) || 'unknown';
    var kLock = 'lock_' + dev, kFail = 'fail_' + dev;
    var now = Date.now();
    var until = Number(cache.get(kLock) || 0);
    if(until > now) return out_({ ok:false, error:'locked', minutes:Math.ceil((until - now) / 60000) });
    // Safety net against someone who keeps changing the device id: a high limit on wrong tries from all devices together.
    var gUntil = Number(cache.get('lockAll') || 0);
    if(gUntil > now) return out_({ ok:false, error:'locked', minutes:Math.ceil((gUntil - now) / 60000) });

    var good = String(body.user || '').trim() === ADMIN_USER && String(body.pass || '') === ADMIN_PASS;
    if(!good){
      var g = Number(cache.get('failAll') || 0) + 1;
      if(g >= MAX_FAILS_ALL){
        cache.put('lockAll', String(now + LOCK_MIN * 60000), LOCK_MIN * 60);
        cache.remove('failAll');
      } else cache.put('failAll', String(g), LOCK_MIN * 60);
      var f = Number(cache.get(kFail) || 0) + 1;
      if(f >= MAX_FAILS){
        cache.put(kLock, String(now + LOCK_MIN * 60000), LOCK_MIN * 60);
        cache.remove(kFail);
        return out_({ ok:false, error:'locked', minutes:LOCK_MIN });
      }
      cache.put(kFail, String(f), LOCK_MIN * 60);
      return out_({ ok:false, error:'bad-login', left:MAX_FAILS - f });
    }
    cache.remove(kFail);
    var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    var s = sessions_(), ids = Object.keys(s);
    if(ids.length >= MAX_SESSIONS){
      ids.sort(function(a, b){ return s[a] - s[b]; });
      for(var i = 0; i <= ids.length - MAX_SESSIONS; i++) delete s[ids[i]];
    }
    s[token] = Date.now();
    saveSessions_(s);
    return out_({ ok:true, token:token });
  } finally { lock.releaseLock(); }
}

function logout_(body){
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    var s = sessions_();
    if(body.token && s[body.token]){ delete s[body.token]; saveSessions_(s); }
    return out_({ ok:true });
  } finally { lock.releaseLock(); }
}

/* ---------- reading (needs a signed-in session) ---------- */
function doGet(e){
  var p = (e && e.parameter) || {};
  if(!validToken_(p.token)) return out_({ ok:false, error:'login-required' });
  var since = Number(p.since || 0);
  if(p.kind === 'contrib' || p.kind === 'repay'){
    var isRepay = p.kind === 'repay';
    var cf = file_(isRepay ? REPAY_FILE : CONTRIB_FILE);
    var flag = {}; flag[isRepay ? 'repay' : 'contrib'] = true;
    if(!cf){ flag.empty = true; return out_(flag); }
    var c = JSON.parse(cf.getBlob().getDataAsString());
    if(since && since === c.updatedAt){ flag.unchanged = true; flag.updatedAt = c.updatedAt; return out_(flag); }
    flag.updatedAt = c.updatedAt; flag.months = c.months || {};
    return out_(flag);
  }
  var f = file_(FILE_NAME);
  if(!f) return out_({ empty:true });
  var saved = JSON.parse(f.getBlob().getDataAsString());
  if(since && since === saved.updatedAt) return out_({ unchanged:true, updatedAt:saved.updatedAt });
  return out_(saved);
}

/* ---------- writing ---------- */
function doPost(e){
  try{
    var body = JSON.parse(e.postData.contents);
    if(body.kind === 'login') return login_(body);
    if(body.kind === 'logout') return logout_(body);
    if(!validToken_(body.token)) return out_({ ok:false, error:'bad-session' });
    if(body.kind === 'check') return out_({ ok:true });
    if(body.kind === 'contrib') return saveMarks_(body, CONTRIB_FILE, false);
    if(body.kind === 'repay') return saveMarks_(body, REPAY_FILE, true);

    // Excel dataset upload (never write an empty dataset)
    if(!body.dataset || !body.dataset.rows || !body.dataset.rows.length) return out_({ ok:false, error:'no-dataset' });
    var saved = { updatedAt: Date.now(), dataset: body.dataset };
    var txt = JSON.stringify(saved), f = file_(FILE_NAME);
    if(f) f.setContent(txt); else DriveApp.createFile(FILE_NAME, txt, MimeType.PLAIN_TEXT);
    return out_({ ok:true, updatedAt:saved.updatedAt });
  }catch(err){ return out_({ ok:false, error:String(err) }); }
}

// body.month = 'YYYY-MM', body.marks = { key: value | null }   (null removes the mark)
//   contributions: value is true | false
//   repayments:    value is { v: true|false, d: 'YYYY-MM-DD' (optional date received) }
// Each save is merged into the stored data under a lock, so two phones marking at the same time never overwrite each other.
function saveMarks_(body, fileName, objectValues){
  if(!/^\d{4}-\d{2}$/.test(String(body.month || ''))) return out_({ ok:false, error:'bad-month' });
  var marks = body.marks || {}, keys = Object.keys(marks);
  if(!keys.length || keys.length > 400) return out_({ ok:false, error:'bad-marks' });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    var cf = file_(fileName);
    var data = cf ? JSON.parse(cf.getBlob().getDataAsString()) : { updatedAt:0, months:{} };
    if(!data.months) data.months = {};
    var m = data.months[body.month] || (data.months[body.month] = {});
    keys.forEach(function(k){
      if(k.length > 200) return;
      var v = marks[k];
      if(objectValues){
        if(v && (v.v === true || v.v === false)){
          var o = { v: v.v };
          if(typeof v.d === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v.d)) o.d = v.d;
          m[k] = o;
        } else delete m[k];
      } else {
        if(v === true || v === false) m[k] = v; else delete m[k];
      }
    });
    data.updatedAt = Date.now();
    var txt = JSON.stringify(data);
    if(cf) cf.setContent(txt); else DriveApp.createFile(fileName, txt, MimeType.PLAIN_TEXT);
    return out_({ ok:true, updatedAt:data.updatedAt, months:data.months });
  } finally { lock.releaseLock(); }
}
