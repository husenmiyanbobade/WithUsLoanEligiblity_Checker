// WITH US Society - shared data store, admin logins, and the society database (a private Google Sheet).
//
// FIRST TIME ONLY (after pasting this file and deploying):
//   1. In the script editor pick the function  setup  and press Run. Allow the permissions it asks for.
//   2. Open View > Logs (Execution log). It shows the link of the new private Sheet and a TEMPORARY password for withusadmin1.
//   3. Sign in to the app as withusadmin1 with that password. The app makes you choose your own password at once.
//   4. In the app (menu > Admin accounts) set the password for withusadmin2.
//   5. Pick the function  installNightlyBackup  and press Run once (daily backup copy of the Sheet, last 30 kept).
// Every later change: Deploy > Manage deployments > pencil icon > Version: "New version" > Deploy.  (The Web app URL stays the same.)
// Execute as: Me, Who has access: Anyone.
//
// Passwords are checked HERE and are stored only as salted hashes in the Admins tab of the Sheet, never inside the web page.
// Until setup has been run, the old single login below keeps working so nothing breaks while you switch over.
var ADMIN_USER = 'withusadmin';     // old shared login, used ONLY until setup() has been run
var ADMIN_PASS = 'withusadmin';
var MAX_FAILS = 5;          // wrong tries allowed per device
var MAX_FAILS_ALL = 40;     // safety net: wrong tries from ALL devices together before everyone pauses
var LOCK_MIN = 15;          // minutes the login stays paused after that
var MAX_SESSIONS = 30;      // phones that can be signed in at once (oldest is signed out first)
var HASH_ROUNDS = 800;      // password hashing strength
var MIN_PASS = 8;           // shortest admin password
var BACKUPS_KEEP = 30;
var SCHEMA_VER = 5;          // bump when a tab gets new columns; the Sheet is upgraded automatically on the next sign-in

var FILE_NAME = 'withus-dataset.json';              // Excel dataset
var CONTRIB_FILE = 'withus-contributions.json';     // monthly contribution marks
var REPAY_FILE = 'withus-repayments.json';          // monthly loan repayment marks
var APPS_FILE = 'withus-applications.json';         // monthly loan application list (from the Eligibility tab)

var DB_NAME = 'WITH US Society Database';
var ADMIN1 = 'withusadmin1', ADMIN2 = 'withusadmin2';

// Every tab of the database and its columns. Nothing is ever deleted from money tabs: rows are voided, not removed.
var TABS = {
  Members: ['MemberID','Name','Mobile','DOB','JoinDate','Status','OpeningDeposit','JoiningFee','PinHash','PinSalt','PinSetAt','FailedTries','LockedUntil','Source','CreatedAt','CreatedBy'],
  Contributions: ['ContribID','MemberID','Month','Amount','PaidDate','Mode','Ref','Source','EnteredBy','EnteredAt','Voided','ReceiptNo','VoidedBy','VoidedAt'],
  Loans: ['LoanID','MemberID','ApplicantName','Amount','TakenDate','DueDate','ReturnAmount','Fine','LateExtra','Status','Suspended','ClosedDate','Source','CreatedAt','LateWaived','WaiveNote'],
  Guarantors: ['LoanID','MemberID','Name','Source'],
  LoanPayments: ['PaymentID','LoanID','Date','Amount','PaidBy','ToPrincipal','ToExtra','ToLateExtra','Mode','Ref','EnteredBy','EnteredAt','Voided'],
  Applications: ['AppID','Month','MemberID','Name','Amount','Guarantors','Status','Remark','CreatedAt'],
  Ledger: ['LedgerID','Date','Pot','Type','MemberID','LoanID','Amount','Ref','Note','EnteredBy','EnteredAt','Voided','VoidedBy','VoidedAt','ReceiptNo'],
  Penalties: ['PenaltyID','MemberID','Month','Amount','Status','WaivedBy','PaidDate','CreatedAt','CreatedBy','WaiveReason','WaivedAt'],
  Suspensions: ['SuspensionID','MemberID','Reason','LoanID','StartDate','EndDate','LiftedBy'],
  Expenses: ['ExpenseID','Date','Category','Amount','Note','EnteredBy','EnteredAt','Voided'],
  Exits: ['ExitID','MemberID','Date','Refund','Note','EnteredBy','EnteredAt'],
  PaymentClaims: ['ClaimID','MemberID','Amount','Purpose','Months','LoanID','UpiRef','Status','ConfirmedBy','ClaimedAt','DecidedAt'],
  Reconciliations: ['ReconID','Date','BankBalance','CashInHand','ExpectedTotal','Difference','EnteredBy','EnteredAt'],
  MonthClose: ['Month','ClosedBy','ClosedAt'],
  Admins: ['Username','Role','Active','Salt','Hash','MustChange','LastLogin','CreatedAt'],
  Audit: ['AuditID','At','Admin','Role','Action','Target','Detail'],
  Settings: ['Key','Value','Note','UpdatedAt','UpdatedBy'],
  ClosedAccounts: ['SrNo','MemberID','Name','Deposit','LoanAmount','LoanReturn','LateFees','Net','Remarks','ImportedAt']
};
var NUM_COLS = ['Deposit','LoanAmount','LoanReturn','LateFees','Net','SrNo','Amount','OpeningDeposit','JoiningFee','ReturnAmount','Fine','LateExtra','ToPrincipal','ToExtra','ToLateExtra','Refund','BankBalance','CashInHand','ExpectedTotal','Difference','FailedTries'];
var DEFAULT_SETTINGS = [
  ['ContributionAmount', 500, 'Monthly contribution per member'],
  ['JoiningFee', 0, 'Joining fee (goes to the fine pot, not refunded)'],
  ['PenaltyAmount', 500, 'Penalty per missed month'],
  ['SuspensionMonths', 6, 'Loan suspension after an overdue loan'],
  ['OverdueDay', 15, 'Loan is overdue after this day of the return month'],
  ['SocietyUPI', '', 'UPI ID that receives contributions and repayments'],
  ['FinancialYearStartMonth', 4, 'April']
];

function out_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function file_(name){ var it = DriveApp.getFilesByName(name); return it.hasNext() ? it.next() : null; }

/* ---------- the database (Google Sheet) ---------- */
// Each read of a script property is a slow call to Google. The same few properties were read thousands of times in one request
// (every row lookup asked for the Sheet id again), which made some screens take minutes. Now each property is read once per request.
var PCACHE_ = {};
function props_(){
  var p = PropertiesService.getScriptProperties();
  return {
    getProperty: function(k){ if(Object.prototype.hasOwnProperty.call(PCACHE_, k)) return PCACHE_[k]; var v = p.getProperty(k); PCACHE_[k] = v; return v; },
    setProperty: function(k, v){ p.setProperty(k, v); PCACHE_[k] = String(v); },
    deleteProperty: function(k){ p.deleteProperty(k); PCACHE_[k] = null; }
  };
}
function sheetReady_(){ return !!props_().getProperty('sheetId'); }
// Speed: opening the spreadsheet and reading a tab are the slow parts of every request, so each is done at most once per
// request and remembered here. Any write clears the remembered rows of that tab. doPost/doGet start with resetCache_().
var BOOK_ = null, BOOK_ID_ = '', TABS_ = {}, ROWS_ = {};
var DIRTY_ = false;
function resetCache_(){ PCACHE_ = {}; BOOK_ = null; BOOK_ID_ = ''; TABS_ = {}; ROWS_ = {}; DIRTY_ = false; }
// Answers that are expensive to work out (member list, month view, profile, loan book) are kept for a few minutes. Any write to the
// Sheet changes 'dataVer', which makes every kept answer stale at once, so a screen never shows old numbers after a change.
function markDirty_(){ if(DIRTY_) return; DIRTY_ = true; props_().setProperty('dataVer', String(Date.now()) + '-' + Math.floor(Math.random() * 1e6)); }
function dataVer_(){ return props_().getProperty('dataVer') || '0'; }
function bodyKey_(b){ var o = {}; Object.keys(b).sort().forEach(function(k){ if(k !== 'token') o[k] = b[k]; }); var s = JSON.stringify(o), h = 5381; for(var i = 0; i < s.length; i++) h = ((h * 33) ^ s.charCodeAt(i)) >>> 0; return s.length + '.' + h; }
function cachedCall_(name, fn){
  var cache = CacheService.getScriptCache(), ver = dataVer_(), key = 'rc|' + name + '|' + ver + '|' + todayStr_();
  if(key.length > 240) return fn();
  var hit = null; try{ hit = cache.get(key); }catch(e){}
  if(hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  var res = fn();
  try{
    var txt = res.getContent();
    if(txt && txt.indexOf('{"ok":true') === 0 && txt.length < 95000 && dataVer_() === ver) cache.put(key, txt, 600);
  }catch(e){}
  return res;
}
function ss_(){
  var id = props_().getProperty('sheetId');
  if(!id) throw new Error('not-set-up');
  if(BOOK_ && BOOK_ID_ === id) return BOOK_;
  BOOK_ = SpreadsheetApp.openById(id); BOOK_ID_ = id; TABS_ = {}; ROWS_ = {};
  return BOOK_;
}
function tab_(name){
  var book = ss_();
  if(TABS_[name]) return TABS_[name];
  var s = book.getSheetByName(name);
  if(!s){                                   // a tab that an older version never created: make it now instead of failing
    if(!TABS[name]) throw new Error('missing-tab-' + name);
    s = book.insertSheet(name); writeHeader_(s, name);
  }
  TABS_[name] = s;
  return s;
}
function fmts_(name){ return TABS[name].map(function(h){ return NUM_COLS.indexOf(h) >= 0 ? '0.##' : '@'; }); }
function rows_(name){
  var sh = tab_(name);
  if(ROWS_[name]) return ROWS_[name].slice();
  var v = sh.getDataRange().getValues(), out = [];
  if(v.length >= 2){
    var h = v[0];
    for(var i = 1; i < v.length; i++){
      var o = { _row: i + 1 };
      for(var j = 0; j < h.length; j++) o[h[j]] = v[i][j];
      out.push(o);
    }
  }
  ROWS_[name] = out;
  return out.slice();
}
function appendMany_(name, list){
  if(!list.length) return;
  delete ROWS_[name]; markDirty_();
  var s = tab_(name), h = TABS[name], f = fmts_(name), r = s.getLastRow() + 1;
  var vals = list.map(function(o){ return h.map(function(k){ return o[k] === undefined || o[k] === null ? '' : o[k]; }); });
  var rng = s.getRange(r, 1, vals.length, h.length);
  rng.setNumberFormats(vals.map(function(){ return f; }));
  rng.setValues(vals);
}
function append_(name, o){ appendMany_(name, [o]); }
function update_(name, rowNum, o){
  delete ROWS_[name]; markDirty_();
  var s = tab_(name), h = TABS[name], rng = s.getRange(rowNum, 1, 1, h.length);
  var cur = rng.getValues()[0];
  h.forEach(function(k, j){ if(Object.prototype.hasOwnProperty.call(o, k)) cur[j] = o[k]; });
  rng.setValues([cur]);
}
function clearRows_(name){
  delete ROWS_[name]; markDirty_();
  var s = tab_(name), last = s.getLastRow();
  if(last > 1) s.getRange(2, 1, last - 1, TABS[name].length).clearContent();
}
function nowIso_(){ return Utilities.formatDate(new Date(), 'Asia/Kolkata', "yyyy-MM-dd'T'HH:mm:ss"); }
function todayStr_(){ return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd'); }
function newId_(prefix){ return prefix + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase(); }
function pad_(n, w){ var s = String(n); while(s.length < w) s = '0' + s; return s; }
function num_(x){ var n = Number(x); return isFinite(n) ? n : 0; }

function audit_(who, action, target, detail){
  try{
    if(!sheetReady_()) return;
    append_('Audit', { AuditID:newId_('A'), At:nowIso_(), Admin:who ? who.u : '', Role:who ? who.r : '', Action:action, Target:target || '', Detail:String(detail || '').slice(0, 500) });
  }catch(e){}
}

/* ---------- passwords ---------- */
function hex_(bytes){ return bytes.map(function(b){ b = b < 0 ? b + 256 : b; return ('0' + b.toString(16)).slice(-2); }).join(''); }
function hash_(pass, salt){
  var h = salt + '|' + pass;
  for(var i = 0; i < HASH_ROUNDS; i++) h = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt + i, Utilities.Charset.UTF_8));
  return h;
}
function newSalt_(){ return Utilities.getUuid().replace(/-/g, ''); }
function tempPass_(){
  var chars = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789', p = '';
  for(var i = 0; i < 10; i++) p += chars.charAt(Math.floor(Math.random() * chars.length));
  return p;
}
function passOk_(user, pass){
  pass = String(pass || '');
  if(pass.length < MIN_PASS) return 'short';
  if(pass.toLowerCase() === String(user).toLowerCase()) return 'same-as-user';
  return '';
}
function adminRow_(user){
  var u = String(user || '').trim().toLowerCase(), list = rows_('Admins');
  for(var i = 0; i < list.length; i++) if(String(list[i].Username).toLowerCase() === u) return list[i];
  return null;
}
function isTrue_(v){ return v === true || String(v).toLowerCase() === 'true'; }

/* ---------- sessions ---------- */
// token -> { t: time, u: username, r: 'admin1' | 'admin2', mc: true while a password change is still required }
function sessions_(){
  var raw = props_().getProperty('sessions');
  try{ return raw ? JSON.parse(raw) : {}; }catch(e){ return {}; }
}
function saveSessions_(s){ props_().setProperty('sessions', JSON.stringify(s)); }
function sess_(t){
  if(typeof t !== 'string' || t.length < 32) return null;
  var s = sessions_()[t];
  return (s && typeof s === 'object' && s.u) ? s : null;      // old-format sessions are simply asked to sign in again
}
function validToken_(t){ return !!sess_(t); }
function dropSessionsOf_(user){
  var s = sessions_(), n = 0;
  Object.keys(s).forEach(function(k){ if(s[k] && s[k].u === user){ delete s[k]; n++; } });
  saveSessions_(s);
  return n;
}

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

    var user = String(body.user || '').trim(), pass = String(body.pass || '');
    var who = null, row = null;
    if(sheetReady_()){
      row = adminRow_(user);
      if(row && isTrue_(row.Active) && row.Hash && row.Salt && hash_(pass, String(row.Salt)) === String(row.Hash)){
        who = { u:String(row.Username), r:String(row.Role), mc:isTrue_(row.MustChange) };
      }
    } else if(user === ADMIN_USER && pass === ADMIN_PASS){
      who = { u:ADMIN_USER, r:'admin1', mc:false };            // old login, only until setup() has been run
    }
    if(!who){
      var g = Number(cache.get('failAll') || 0) + 1;
      if(g >= MAX_FAILS_ALL){
        cache.put('lockAll', String(now + LOCK_MIN * 60000), LOCK_MIN * 60);
        cache.remove('failAll');
      } else cache.put('failAll', String(g), LOCK_MIN * 60);
      var f = Number(cache.get(kFail) || 0) + 1;
      audit_({ u:user.slice(0, 40), r:'' }, 'login-failed', user.slice(0, 40), 'try ' + f);
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
      ids.sort(function(a, b){ return (s[a].t || 0) - (s[b].t || 0); });
      for(var i = 0; i <= ids.length - MAX_SESSIONS; i++) delete s[ids[i]];
    }
    s[token] = { t:Date.now(), u:who.u, r:who.r, mc:!!who.mc };
    saveSessions_(s);
    if(row) update_('Admins', row._row, { LastLogin:nowIso_() });
    audit_(who, 'login', who.u, '');
    return out_({ ok:true, token:token, user:who.u, role:who.r, mustChange:!!who.mc, db:sheetReady_() });
  } finally { lock.releaseLock(); }
}

function logout_(body){
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    var s = sessions_();
    if(body.token && s[body.token]){ audit_(s[body.token], 'logout', s[body.token].u, ''); delete s[body.token]; saveSessions_(s); }
    return out_({ ok:true });
  } finally { lock.releaseLock(); }
}

/* ---------- admin account actions ---------- */
function need_(S, role){ return S.r === role ? null : out_({ ok:false, error:'forbidden' }); }
function needDb_(){ return sheetReady_() ? null : out_({ ok:false, error:'not-set-up' }); }

function adminList_(S){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var ses = sessions_(), counts = {};
  Object.keys(ses).forEach(function(k){ var u = ses[k] && ses[k].u; if(u) counts[u] = (counts[u] || 0) + 1; });
  return out_({ ok:true, admins: rows_('Admins').map(function(r){
    return { user:r.Username, role:r.Role, active:isTrue_(r.Active), hasPassword:!!r.Hash, mustChange:isTrue_(r.MustChange), lastLogin:r.LastLogin || '', phones:counts[r.Username] || 0 };
  }) });
}

// admin1 changes their own password (needs the old one). admin2 can never change any password.
function changePassword_(S, body, token){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var row = adminRow_(S.u);
  if(!row || hash_(String(body.old || ''), String(row.Salt)) !== String(row.Hash)) return out_({ ok:false, error:'bad-old' });
  var pe = passOk_(S.u, body.pass); if(pe) return out_({ ok:false, error:pe });
  if(String(body.pass) === String(body.old)) return out_({ ok:false, error:'same-as-old' });
  var salt = newSalt_();
  update_('Admins', row._row, { Salt:salt, Hash:hash_(String(body.pass), salt), MustChange:false });
  var ses = sessions_();
  Object.keys(ses).forEach(function(k){ if(ses[k] && ses[k].u === S.u){ ses[k].mc = false; if(k !== token) delete ses[k]; } });  // other phones sign in again with the new password
  saveSessions_(ses);
  audit_(S, 'password-changed', S.u, 'own password');
  return out_({ ok:true });
}

// admin1 sets (or resets) the password of admin2.
function setPassword_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  if(String(body.user) !== ADMIN2) return out_({ ok:false, error:'bad-target' });
  var pe = passOk_(ADMIN2, body.pass); if(pe) return out_({ ok:false, error:pe });
  var row = adminRow_(ADMIN2); if(!row) return out_({ ok:false, error:'bad-target' });
  var salt = newSalt_();
  update_('Admins', row._row, { Salt:salt, Hash:hash_(String(body.pass), salt), MustChange:false, Active:true });
  dropSessionsOf_(ADMIN2);
  audit_(S, 'password-set', ADMIN2, 'by admin1');
  return out_({ ok:true });
}

function setActive_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  if(String(body.user) !== ADMIN2) return out_({ ok:false, error:'bad-target' });
  var row = adminRow_(ADMIN2); if(!row) return out_({ ok:false, error:'bad-target' });
  var on = body.active === true;
  update_('Admins', row._row, { Active:on });
  if(!on) dropSessionsOf_(ADMIN2);
  audit_(S, on ? 'enabled' : 'disabled', ADMIN2, '');
  return out_({ ok:true, active:on });
}

function signOutAll_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  if(String(body.user) !== ADMIN2) return out_({ ok:false, error:'bad-target' });
  var n = dropSessionsOf_(ADMIN2);
  audit_(S, 'signed-out-all', ADMIN2, n + ' phones');
  return out_({ ok:true, phones:n });
}

function auditList_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var lim = Math.max(1, Math.min(200, Number(body.limit) || 50));
  var list = rows_('Audit');
  return out_({ ok:true, items: list.slice(-lim).reverse().map(function(r){ return { at:r.At, admin:r.Admin, action:r.Action, target:r.Target, detail:r.Detail }; }) });
}

/* ---------- database status, import, totals check, export ---------- */
function dbStatus_(S){
  if(!sheetReady_()) return out_({ ok:true, ready:false });
  var o = { ok:true, ready:true, counts:{} };
  ['Members','Contributions','Loans','Ledger','Audit'].forEach(function(n){ o.counts[n] = Math.max(0, tab_(n).getLastRow() - 1); });
  if(S.r === 'admin1') o.url = ss_().getUrl();
  return out_(o);
}

function monthKey_(mc){
  var y = mc.year || new Date().getFullYear();
  return y + '-' + pad_((Number(mc.month) || 0) + 1, 2);
}
function nameKey_(s){ return String(s || '').replace(/\(\s*number\s*\d+\s*\)/i, '').replace(/\s+/g, ' ').trim().toUpperCase(); }   // "NAME (Number 2)" is the same person
function readDataset_(){
  var f = file_(FILE_NAME);
  if(!f) return null;
  var saved = JSON.parse(f.getBlob().getDataAsString());
  return (saved && saved.dataset) || null;
}
function fineFigures_(ds){
  var md = ds.memberData || {}, info = md.fineInfo || null, rows = ds.rows || [];
  var sumRows = rows.reduce(function(s, r){ return s + num_(r.fine); }, 0);
  var collected = info && info.total !== null && info.total !== undefined ? num_(info.total) : sumRows;
  var spent = info && info.spent !== null && info.spent !== undefined ? num_(info.spent) : 0;
  var items = (info && info.items) || [];
  return { collected:collected, spent:spent, remaining: info && info.remaining !== null && info.remaining !== undefined ? num_(info.remaining) : collected - spent, items:items, hasInfo:!!info };
}

// Copies the Excel data already stored in the app into the Sheet (members, monthly contributions, loans, fine pot).
// Safe to repeat while only imported data exists. Refuses once real entries or member PINs exist.
function importDataset_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  if(goLiveInfo_()) return out_({ ok:false, error:'golive-locked' });
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    var ds = readDataset_();
    if(!ds || !ds.memberData || !ds.memberData.members || !ds.memberData.members.length) return out_({ ok:false, error:'no-dataset' });
    var existing = rows_('Members');
    if(existing.length){
      if(body.replace !== true) return out_({ ok:false, error:'already-imported', members:existing.length });
      var live = rows_('Ledger').some(function(r){ return String(r.EnteredBy) !== 'import'; }) || existing.some(function(m){ return !!m.PinHash || String(m.Source) !== 'import'; });
      if(live) return out_({ ok:false, error:'has-live-data' });
      ['Members','Contributions','Loans','Guarantors','Ledger'].forEach(clearRows_);
    }
    var md = ds.memberData, at = nowIso_(), today = todayStr_();
    var members = [], contribs = [], ledger = [], byName = {}, dupes = [];
    md.members.forEach(function(m, i){
      var id = 'M-' + pad_(i + 1, 4), key = nameKey_(m.name);
      if(byName[key]) dupes.push(m.name); else byName[key] = id;
      members.push({ MemberID:id, Name:String(m.name).trim(), Mobile:'', DOB:'', JoinDate:m.joinDate || '', Status:'Active', OpeningDeposit:num_(m.depositPrior), JoiningFee:0, FailedTries:0, Source:'import', CreatedAt:at, CreatedBy:'import' });
      if(num_(m.depositPrior)) ledger.push({ LedgerID:newId_('L'), Date:today, Pot:'deposit', Type:'opening', MemberID:id, Amount:num_(m.depositPrior), Note:'Opening deposit (imported)', EnteredBy:'import', EnteredAt:at });
      (m.months || []).forEach(function(c, k){
        var amt = num_(c.amount); if(!amt) return;
        var mk = monthKey_(md.monthCols[k] || c);
        contribs.push({ ContribID:newId_('C'), MemberID:id, Month:mk, Amount:amt, PaidDate:mk + '-01', Mode:'', Source:'import', EnteredBy:'import', EnteredAt:at });
        ledger.push({ LedgerID:newId_('L'), Date:mk + '-01', Pot:'deposit', Type:'contribution', MemberID:id, Amount:amt, Note:'Imported ' + mk, EnteredBy:'import', EnteredAt:at });
      });
    });
    var loans = [], guars = [], unmatched = 0, gUnmatched = 0;
    (ds.rows || []).forEach(function(r, i){
      var lid = 'L-' + pad_(r.srno || (i + 1), 4), mid = byName[nameKey_(r.applicant)] || '';
      if(!mid) unmatched++;
      loans.push({ LoanID:lid, MemberID:mid, ApplicantName:r.applicant, Amount:num_(r.takenAmt), TakenDate:r.takenDate || '', DueDate:r.repayDate || '', ReturnAmount:num_(r.repayAmt), Fine:num_(r.fine), LateExtra:0, Status:r.status || '', Suspended:r.suspended || '', Source:'import', CreatedAt:at });
      String(r.guarantor || '').split(/\s*(?:,|&|\+|\n|;|\band\b)\s*/i).forEach(function(g){
        g = String(g).trim(); if(!g || /^(none|nil|n\/a|-|not specified)$/i.test(g)) return;
        var gid = byName[nameKey_(g)] || ''; if(!gid) gUnmatched++;
        guars.push({ LoanID:lid, MemberID:gid, Name:g, Source:'import' });
      });
    });
    var fp = fineFigures_(ds);
    if(fp.collected) ledger.push({ LedgerID:newId_('L'), Date:today, Pot:'fine', Type:'opening', Amount:fp.collected, Note:'Fine pot collected so far (imported)', EnteredBy:'import', EnteredAt:at });
    var itemSum = 0;
    fp.items.forEach(function(it){
      var a = num_(it.amount); if(!a) return; itemSum += a;
      ledger.push({ LedgerID:newId_('L'), Date:today, Pot:'fine', Type:'expense', Amount:-a, Note:String(it.label || 'Expense (imported)').slice(0, 120), EnteredBy:'import', EnteredAt:at });
    });
    if(Math.abs(fp.spent - itemSum) > 0.5 && fp.spent > itemSum) ledger.push({ LedgerID:newId_('L'), Date:today, Pot:'fine', Type:'expense', Amount:-(fp.spent - itemSum), Note:'Other spending (imported total)', EnteredBy:'import', EnteredAt:at });
    appendMany_('Members', members); appendMany_('Contributions', contribs); appendMany_('Loans', loans); appendMany_('Guarantors', guars); appendMany_('Ledger', ledger);
    audit_(S, 'import', 'Excel dataset', members.length + ' members, ' + loans.length + ' loans, ' + ledger.length + ' ledger rows');
    return out_({ ok:true, members:members.length, contributions:contribs.length, loans:loans.length, guarantors:guars.length, ledger:ledger.length, duplicateNames:dupes, loansWithoutMember:unmatched, guarantorsWithoutMember:gUnmatched, check:check_() });
  } finally { lock.releaseLock(); }
}

// Compares the Excel data with what the Sheet now says. Every line must read ok.
function check_(){
  var ds = readDataset_();
  if(!ds || !ds.memberData) return { ok:false, error:'no-dataset' };
  var md = ds.memberData, rowsX = ds.rows || [];
  var all = rows_('Members'), loans = rows_('Loans'), ledger = rows_('Ledger').filter(function(r){ return !isTrue_(r.Voided); });
  var members = all.filter(function(m){ return !isClosedSrc_(m); });          // closed accounts have their own Excel sheet: they are not counted with members
  var closedS = all.filter(isClosedSrc_);
  function sum(list, f){ return list.reduce(function(s, x){ return s + num_(f(x)); }, 0); }
  var lines = [], notes = [];
  function line(label, excel, sheet){ lines.push({ label:label, excel:excel, sheet:sheet, ok:Math.abs(excel - sheet) < 0.5 }); }
  line('Members (closed accounts not counted)', md.members.length, members.length);
  var closedX = md.closedAccounts ? md.closedAccounts.length : 0;
  if(closedX || closedS.length) line('Closed accounts (Exited members)', closedX || closedS.length, closedS.length);
  // The Excel is only updated when someone updates it, while the app keeps taking entries: compare what the import wrote, and show the app's own entries on their own.
  var impL = ledger.filter(function(r){ return r.Pot === 'deposit' && String(r.EnteredBy) === 'import'; });
  var appL = ledger.filter(function(r){ return r.Pot === 'deposit' && String(r.EnteredBy) !== 'import'; });
  var depX = sum(md.members, function(m){ return m.depositCurrent; });
  line('Deposit pot (as imported from Excel)', depX, sum(impL, function(r){ return r.Amount; }));
  var appSum = sum(appL, function(r){ return r.Amount; });
  if(appL.length) notes.push({ label:'Entered in the app after the import (not in Excel)', amount:r2_(appSum), count:appL.length });
  var openX = sum(md.members, function(m){ return m.depositPrior; });
  line('Opening deposits', openX, sum(members, function(m){ return m.OpeningDeposit; }));
  var blankX = rowsX.filter(function(r){ return blankName_(r.applicant) && !num_(r.takenAmt); }).length, blankS = loans.filter(function(r){ return blankLoan_(r); }).length;
  line('Loans (count, blank rows left out)', rowsX.length - blankX, loans.length - blankS);
  if(blankS) notes.push({ label:'Blank loan rows still in the Sheet (clean-up removes them)', amount:0, count:blankS });
  line('Loans (amount taken)', sum(rowsX, function(r){ return r.takenAmt; }), sum(loans, function(r){ return r.Amount; }));
  line('Loans (amount to return)', sum(rowsX, function(r){ return r.repayAmt; }), sum(loans, function(r){ return r.ReturnAmount; }));
  line('Loans (extra / fine)', sum(rowsX, function(r){ return r.fine; }), sum(loans, function(r){ return r.Fine; }));
  var fp = fineFigures_(ds);
  line('Fine pot (remaining)', fp.remaining, sum(ledger.filter(function(r){ return r.Pot === 'fine' && String(r.EnteredBy) === 'import'; }), function(r){ return r.Amount; }));
  var fineApp = ledger.filter(function(r){ return r.Pot === 'fine' && String(r.EnteredBy) !== 'import'; });
  if(fineApp.length) notes.push({ label:'Fine pot entries made in the app after the import', amount:r2_(sum(fineApp, function(r){ return r.Amount; })), count:fineApp.length });
  // member by member: Excel current deposit against what the import wrote
  var byId = {}; impL.forEach(function(r){ byId[r.MemberID] = (byId[r.MemberID] || 0) + num_(r.Amount); });
  var diffs = [];
  md.members.forEach(function(m, i){
    var id = 'M-' + pad_(i + 1, 4), a = num_(m.depositCurrent), b = byId[id] || 0;
    if(Math.abs(a - b) >= 0.5) diffs.push({ id:id, name:m.name, excel:a, sheet:b });
  });
  return { ok: lines.every(function(l){ return l.ok; }) && !diffs.length, lines:lines, notes:notes, memberDiffs:diffs.slice(0, 25), memberDiffCount:diffs.length };
}
function totalsCheck_(S){
  var bad = needDb_(); if(bad) return bad;
  var res = check_(); audit_(S, 'totals-check', '', res.ok ? 'all ok' : 'differences found');
  res.ready = true;
  return out_(res);
}

// Everything in the Sheet as plain rows, for the Excel backup made in the app. Password and PIN hashes are left out.
function exportAll_(S){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var secret = { Admins:['Salt','Hash'], Members:['PinHash','PinSalt'] }, tabs = {};
  Object.keys(TABS).forEach(function(n){
    var h = TABS[n], hide = secret[n] || [];
    var v = tab_(n).getDataRange().getValues();
    tabs[n] = v.map(function(r, i){ return r.map(function(c, j){ return (i > 0 && hide.indexOf(h[j]) >= 0) ? '' : c; }); });
  });
  audit_(S, 'export', 'Excel backup', '');
  return out_({ ok:true, at:nowIso_(), tabs:tabs });
}

function writeHeader_(sh, n){
  var h = TABS[n], f = fmts_(n);
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
  sh.setFrozenRows(1);
  for(var c = 0; c < h.length; c++) sh.getRange(1, c + 1, Math.max(sh.getMaxRows(), 2), 1).setNumberFormat(f[c]);
}
// Adds missing tabs and new columns after an update, once per schema version.
function ensureSchema_(){
  if(!sheetReady_() || Number(props_().getProperty('schemaVer') || 0) >= SCHEMA_VER) return;
  var book = ss_();
  Object.keys(TABS).forEach(function(n){ writeHeader_(book.getSheetByName(n) || book.insertSheet(n), n); });
  props_().setProperty('schemaVer', String(SCHEMA_VER));
}

/* ---------- one-time setup and backups (run from the script editor) ---------- */
function setup(){
  resetCache_();
  var id = props_().getProperty('sheetId'), book, fresh = false;
  if(id) book = SpreadsheetApp.openById(id);
  else { book = SpreadsheetApp.create(DB_NAME); props_().setProperty('sheetId', book.getId()); fresh = true; }
  var first = true;
  Object.keys(TABS).forEach(function(n){
    var sh = book.getSheetByName(n);
    if(!sh) sh = first && fresh ? book.getSheets()[0].setName(n) : book.insertSheet(n);
    first = false;
    writeHeader_(sh, n);
  });
  props_().setProperty('schemaVer', String(SCHEMA_VER));
  var set = book.getSheetByName('Settings');
  if(set.getLastRow() < 2){
    var at = nowIso_();
    set.getRange(2, 1, DEFAULT_SETTINGS.length, 5).setValues(DEFAULT_SETTINGS.map(function(r){ return [r[0], r[1], r[2], at, 'setup']; }));
  }
  var temp = '';
  if(!adminRow_(ADMIN1)){
    temp = tempPass_();
    var salt = newSalt_(), at2 = nowIso_();
    append_('Admins', { Username:ADMIN1, Role:'admin1', Active:true, Salt:salt, Hash:hash_(temp, salt), MustChange:true, CreatedAt:at2 });
    append_('Admins', { Username:ADMIN2, Role:'admin2', Active:true, Salt:'', Hash:'', MustChange:false, CreatedAt:at2 });
    audit_({ u:'setup', r:'' }, 'setup', 'database', 'created Admins');
  }
  Logger.log('Database Sheet: ' + book.getUrl());
  if(temp) Logger.log('TEMPORARY password for ' + ADMIN1 + ': ' + temp + '   (you must choose your own at first sign in). ' + ADMIN2 + ' has no password yet: set it from the app, menu > Admin accounts.');
  else Logger.log('Tabs checked. Admin accounts already exist and were not changed.');
  return book.getUrl();
}

// Emergency only: if withusadmin1 forgets the password, the owner of this Google account runs this and reads the log.
function resetAdmin1(){
  resetCache_();
  var row = adminRow_(ADMIN1);
  if(!row) throw new Error('Run setup first');
  var temp = tempPass_(), salt = newSalt_();
  update_('Admins', row._row, { Salt:salt, Hash:hash_(temp, salt), MustChange:true, Active:true });
  dropSessionsOf_(ADMIN1);
  audit_({ u:'owner', r:'' }, 'reset-admin1', ADMIN1, 'from script editor');
  Logger.log('New TEMPORARY password for ' + ADMIN1 + ': ' + temp);
}

function nightlyBackup(){
  resetCache_();
  if(!sheetReady_()) return;
  var file = DriveApp.getFileById(props_().getProperty('sheetId'));
  var it = DriveApp.getFoldersByName('WITH US backups'), folder = it.hasNext() ? it.next() : DriveApp.createFolder('WITH US backups');
  file.makeCopy('WITH US backup ' + todayStr_(), folder);
  var list = [], files = folder.getFiles();
  while(files.hasNext()){ var f = files.next(); list.push({ f:f, t:f.getDateCreated().getTime() }); }
  list.sort(function(a, b){ return b.t - a.t; });
  list.slice(BACKUPS_KEEP).forEach(function(x){ x.f.setTrashed(true); });
}
function installNightlyBackup(){
  ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction() === 'nightlyBackup') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('nightlyBackup').timeBased().everyDays(1).atHour(2).create();
  Logger.log('Nightly backup installed (about 2 am).');
}

/* ---------- step 2: members and contributions ---------- */
var UNDO_SECONDS = 120;      // an admin may void their OWN entry within this time; later only withusadmin1 can

function setting_(key, dflt){
  var list = rows_('Settings');
  for(var i = 0; i < list.length; i++) if(String(list[i].Key) === key && list[i].Value !== '' && list[i].Value !== null) return list[i].Value;
  return dflt;
}
function contribAmount_(){ return num_(setting_('ContributionAmount', 500)) || 500; }
function trackStart_(){ var v = String(setting_('TrackingStartMonth', '2026-01')); return validMonth_(v) ? v : '2026-01'; }
function curMonth_(){ return todayStr_().slice(0, 7); }
function validMonth_(m){ return /^\d{4}-(0[1-9]|1[0-2])$/.test(String(m)); }
function validDate_(s){
  s = String(s || '');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  var d = new Date(s + 'T00:00:00Z');
  return !isNaN(d.getTime()) && d.toISOString().slice(0, 10) === s;
}
function addMonths_(m, n){
  var y = Number(m.slice(0, 4)), mo = Number(m.slice(5, 7)) - 1 + n;
  y += Math.floor(mo / 12); mo = ((mo % 12) + 12) % 12;
  return y + '-' + pad_(mo + 1, 2);
}
function monthsBetween_(a, b){
  var out = [], m = a, guard = 0;
  while(m <= b && guard++ < 600){ out.push(m); m = addMonths_(m, 1); }
  return out;
}
function startMonth_(mem){
  var ts = trackStart_(), j = String(mem.JoinDate || '').slice(0, 7);
  return (validMonth_(j) && j > ts) ? j : ts;
}
function cleanName_(s){ return String(s || '').replace(/[^A-Za-zऀ-ॿ .'\-]/g, ' ').replace(/\s+/g, ' ').trim().toUpperCase(); }
function cleanMobile_(s){
  var d = String(s || '').replace(/\D/g, '');
  if(d.length === 12 && d.indexOf('91') === 0) d = d.slice(2);
  if(d.length === 11 && d.charAt(0) === '0') d = d.slice(1);
  return /^[6-9]\d{9}$/.test(d) ? d : '';
}
function cleanRef_(s){ return String(s || '').replace(/[^A-Za-z0-9 _\-\/]/g, '').trim().slice(0, 40); }
function lev_(a, b){
  var m = a.length, n = b.length, d = [], i, j;
  for(i = 0; i <= m; i++){ d[i] = [i]; }
  for(j = 1; j <= n; j++) d[0][j] = j;
  for(i = 1; i <= m; i++) for(j = 1; j <= n; j++) d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1));
  return d[m][n];
}
function monthClosed_(month){
  var list = rows_('MonthClose');
  for(var i = 0; i < list.length; i++) if(String(list[i].Month) === month) return true;
  return false;
}
function nextReceipt_(){
  var n = Number(props_().getProperty('receiptSeq') || 0) + 1;
  props_().setProperty('receiptSeq', String(n));
  return 'R-' + pad_(n, 5);
}
function memberById_(id){
  var list = rows_('Members');
  for(var i = 0; i < list.length; i++) if(String(list[i].MemberID) === String(id)) return list[i];
  return null;
}
function nextMemberId_(){
  var max = 0;
  rows_('Members').forEach(function(m){ var n = Number(String(m.MemberID).replace(/\D/g, '')); if(n > max) max = n; });
  return 'M-' + pad_(max + 1, 4);
}

// memberId -> { 'YYYY-MM': { id, amount, date, mode, receipt, by } } for contributions that are not voided
function contribIndex_(){
  var idx = {};
  rows_('Contributions').forEach(function(c){
    if(isTrue_(c.Voided)) return;
    var m = idx[c.MemberID] || (idx[c.MemberID] = {});
    m[String(c.Month)] = { id:c.ContribID, amount:num_(c.Amount), date:String(c.PaidDate || ''), mode:c.Mode || '', receipt:c.ReceiptNo || '', by:c.EnteredBy || '', at:c.EnteredAt || '', ref:c.Ref || '' };
  });
  return idx;
}
function depositTotals_(){
  var t = {};
  rows_('Ledger').forEach(function(r){ if(r.Pot === 'deposit' && !isTrue_(r.Voided) && r.MemberID) t[r.MemberID] = (t[r.MemberID] || 0) + num_(r.Amount); });
  return t;
}
// Past months not paid (arrears) and whether the current month is still unpaid (due). Exited members owe nothing.
function standing_(mem, paid, cur){
  if(String(mem.Status) === 'Exited') return { arrears:[], due:false };
  var start = startMonth_(mem), arrears = [], due = false;
  monthsBetween_(start, cur).forEach(function(m){
    if(paid && paid[m]) return;
    if(m < cur) arrears.push(m); else due = true;
  });
  return { arrears:arrears, due:due };
}
function memberPublic_(m, st, deposit, pen, susp){
  var o = { id:m.MemberID, name:m.Name, mobile:String(m.Mobile || ''), dob:String(m.DOB || ''), joinDate:String(m.JoinDate || ''), status:m.Status || 'Active',
    deposit:deposit || 0, arrears:st.arrears, dueNow:st.due, hasPin:!!m.PinHash };
  o.penDue = pen ? pen.due : 0;
  var el = eligibility_(m, st, o.penDue, susp);
  o.canApply = el.canApply; o.canGuarantee = el.canGuarantee; o.blocks = el.reasons;
  return o;
}

/* ---------- step 3: penalties and the eligibility block ---------- */
function penaltyAmount_(){ return num_(setting_('PenaltyAmount', 500)) || 500; }
// memberId -> { due, charged, paid, waived, list:[{id, month, amount, status, paid, outstanding, ...}] } (oldest month first)
function penaltyIndex_(){
  var paidBy = {};
  rows_('Ledger').forEach(function(l){ if(l.Type === 'penalty_paid' && !isTrue_(l.Voided)) paidBy[l.Ref] = (paidBy[l.Ref] || 0) + num_(l.Amount); });
  var idx = {};
  rows_('Penalties').forEach(function(p){
    var m = idx[p.MemberID] || (idx[p.MemberID] = { due:0, charged:0, paid:0, waived:0, list:[] });
    var amt = num_(p.Amount), paid = paidBy[p.PenaltyID] || 0, waived = String(p.Status) === 'Waived';
    var out = waived ? 0 : Math.max(0, amt - paid);
    m.list.push({ id:p.PenaltyID, month:String(p.Month), amount:amt, status:waived ? 'Waived' : (out === 0 ? 'Paid' : (paid > 0 ? 'Part paid' : 'Unpaid')), paid:paid, outstanding:out, waivedBy:p.WaivedBy || '', reason:p.WaiveReason || '', createdAt:p.CreatedAt || '' });
    m.charged += amt; m.paid += paid; m.due += out; if(waived) m.waived += Math.max(0, amt - paid);
  });
  Object.keys(idx).forEach(function(k){ idx[k].list.sort(function(a, b){ return a.month.localeCompare(b.month); }); });
  return idx;
}
function suspensionIndex_(){
  var today = todayStr_(), idx = {};
  rows_('Suspensions').forEach(function(s){
    if(s.LiftedBy) return;
    var end = String(s.EndDate || ''), start = String(s.StartDate || '');
    if(end && end >= today && (!start || start <= today)) idx[s.MemberID] = { until:end, reason:s.Reason || '', loan:s.LoanID || '' };
  });
  return idx;
}
// What stops a member from applying for a loan or standing as guarantor. The reasons name exactly what is owed.
function eligibility_(mem, st, penDue, susp){
  var r = [], status = String(mem.Status || 'Active'), apply = true, guar = true;
  if(status === 'Exited'){ r.push({ type:'exited' }); apply = false; guar = false; }
  else if(status === 'Suspended'){ r.push({ type:'suspended' }); apply = false; guar = false; }
  if(st.arrears && st.arrears.length){ r.push({ type:'arrears', months:st.arrears, amount:st.arrears.length * contribAmount_() }); apply = false; guar = false; }
  if(penDue > 0){ r.push({ type:'penalty', amount:penDue }); apply = false; guar = false; }
  if(susp){ r.push({ type:'loan-suspension', until:susp.until }); apply = false; }
  return { canApply:apply, canGuarantee:guar, reasons:r };
}
// Candidates for a month's penalty: active members who have not paid that (ended) month and have no penalty for it yet.
function penaltyCandidates_(month){
  var ci = contribIndex_(), have = {};
  rows_('Penalties').forEach(function(p){ have[p.MemberID + '|' + p.Month] = 1; });
  var out = [];
  rows_('Members').forEach(function(m){
    if(String(m.Status) === 'Exited') return;
    if(month < startMonth_(m)) return;
    if(ci[m.MemberID] && ci[m.MemberID][month]) return;
    if(have[m.MemberID + '|' + month]) return;
    out.push({ id:m.MemberID, name:m.Name });
  });
  out.sort(function(a, b){ return a.name < b.name ? -1 : 1; });
  return out;
}
function endedMonths_(){
  var cur = curMonth_(), ts = trackStart_(), last = addMonths_(cur, -1);
  return ts <= last ? monthsBetween_(ts, last) : [];
}

function penaltyPreview_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var amt = penaltyAmount_(), months = endedMonths_(), out = [];
  months.forEach(function(m){ var c = penaltyCandidates_(m); out.push({ month:m, count:c.length, closed:monthClosed_(m) }); });
  var res = { ok:true, amount:amt, month:curMonth_(), months:out.reverse() };
  if(body && body.month){
    if(months.indexOf(String(body.month)) < 0) return out_({ ok:false, error:'bad-month' });
    res.pick = String(body.month); res.candidates = penaltyCandidates_(res.pick);
  }
  var pi = penaltyIndex_(), names = {}, owed = [];
  rows_('Members').forEach(function(m){ names[m.MemberID] = m; });
  Object.keys(pi).forEach(function(id){ if(pi[id].due > 0 && names[id]) owed.push({ id:id, name:names[id].Name, due:pi[id].due, count:pi[id].list.filter(function(x){ return x.outstanding > 0; }).length }); });
  owed.sort(function(a, b){ return b.due - a.due || (a.name < b.name ? -1 : 1); });
  res.owed = owed; res.owedTotal = owed.reduce(function(s2, x){ return s2 + x.due; }, 0);
  return out_(res);
}

function penaltyPost_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var month = String(body.month || '');
    if(endedMonths_().indexOf(month) < 0) return out_({ ok:false, error:'bad-month' });
    var cands = penaltyCandidates_(month), want = null;
    if(Array.isArray(body.ids)){ want = {}; body.ids.forEach(function(i){ want[String(i)] = 1; }); cands = cands.filter(function(c){ return want[c.id]; }); }
    if(!cands.length) return out_({ ok:false, error:'nothing-to-post' });
    var amt = penaltyAmount_(), at = nowIso_();
    appendMany_('Penalties', cands.map(function(c){ return { PenaltyID:newId_('P'), MemberID:c.id, Month:month, Amount:amt, Status:'Charged', CreatedAt:at, CreatedBy:S.u }; }));
    audit_(S, 'penalties-posted', month, cands.length + ' member(s) x ' + amt);
    return out_({ ok:true, posted:cands.length, amount:amt, total:cands.length * amt, month:month });
  } finally { lock.releaseLock(); }
}

// Pay a member's penalties, oldest first. Partial payment is allowed. The money goes to the fine pot.
function penaltyPay_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var mem = memberById_(body.memberId); if(!mem) return out_({ ok:false, error:'no-member' });
    var today = todayStr_(), date = body.date ? String(body.date) : today;
    if(!validDate_(date) || date > today) return out_({ ok:false, error:'bad-date' });
    var mode = ['Bank','UPI'].indexOf(body.mode) >= 0 ? body.mode : '';
    if(!mode) return out_({ ok:false, error:'bad-mode' });
    var amount = Math.round(num_(body.amount) * 100) / 100, pi = penaltyIndex_()[mem.MemberID];
    var open = pi ? pi.list.filter(function(x){ return x.outstanding > 0; }) : [];
    if(!open.length) return out_({ ok:false, error:'nothing-owed' });
    if(!(amount > 0)) return out_({ ok:false, error:'bad-amount' });
    if(amount > pi.due) return out_({ ok:false, error:'too-much', due:pi.due });
    if(monthClosed_(date.slice(0, 7))) return out_({ ok:false, error:'month-closed', month:date.slice(0, 7) });
    var no = nextReceipt_(), at = nowIso_(), left = amount, ls = [], months = [], ref = cleanRef_(body.ref);
    open.forEach(function(p){
      if(left <= 0) return;
      var part = Math.min(left, p.outstanding); left -= part; months.push(p.month);
      ls.push({ LedgerID:newId_('L'), Date:date, Pot:'fine', Type:'penalty_paid', MemberID:mem.MemberID, Amount:part, Ref:p.id, Note:'Penalty ' + p.month + (ref ? ' · ' + ref : '') + ' · ' + mode, EnteredBy:S.u, EnteredAt:at, ReceiptNo:no });
    });
    appendMany_('Ledger', ls);
    audit_(S, 'penalty-paid', mem.MemberID, no + ' · ' + amount + ' · ' + months.join(', ') + ' · ' + mode);
    var r = receiptObj_(no, mem, months, amount, date, mode, ref, S.u); r.kind = 'penalty';
    return out_({ ok:true, receipt:r, due:pi.due - amount });
  } finally { lock.releaseLock(); }
}

// Void a penalty payment receipt. withusadmin1 any time; the admin who made it, only within UNDO_SECONDS.
function penaltyPayVoid_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var key = String(body.receipt || ''); if(!key) return out_({ ok:false, error:'bad-request' });
    var hit = rows_('Ledger').filter(function(l){ return l.Type === 'penalty_paid' && !isTrue_(l.Voided) && String(l.ReceiptNo) === key; });
    if(!hit.length) return out_({ ok:false, error:'not-found' });
    var nowMs = Date.now();
    for(var i = 0; i < hit.length; i++){
      if(monthClosed_(String(hit[i].Date).slice(0, 7))) return out_({ ok:false, error:'month-closed', month:String(hit[i].Date).slice(0, 7) });
      if(S.r !== 'admin1'){
        var t = Date.parse(String(hit[i].EnteredAt) + '+05:30');
        if(String(hit[i].EnteredBy) !== S.u || !isFinite(t) || nowMs - t > UNDO_SECONDS * 1000) return out_({ ok:false, error:'forbidden' });
      }
    }
    var at = nowIso_();
    hit.forEach(function(l){ update_('Ledger', l._row, { Voided:true, VoidedBy:S.u, VoidedAt:at }); });
    audit_(S, 'penalty-payment-voided', hit[0].MemberID, key + ' · ' + String(body.reason || '').slice(0, 100));
    return out_({ ok:true, voided:hit.length });
  } finally { lock.releaseLock(); }
}

// Waive what is still owed on one penalty. withusadmin1 only, with a reason. Money already paid stays paid.
function penaltyWaive_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var reason = String(body.reason || '').replace(/\s+/g, ' ').trim().slice(0, 120);
    if(reason.length < 3) return out_({ ok:false, error:'reason-needed' });
    var p = rows_('Penalties').filter(function(x){ return String(x.PenaltyID) === String(body.penaltyId); })[0];
    if(!p) return out_({ ok:false, error:'not-found' });
    if(String(p.Status) === 'Waived') return out_({ ok:true, unchanged:true });
    update_('Penalties', p._row, { Status:'Waived', WaivedBy:S.u, WaiveReason:reason, WaivedAt:nowIso_() });
    audit_(S, 'penalty-waived', p.MemberID, p.PenaltyID + ' · ' + p.Month + ' · ' + reason);
    return out_({ ok:true });
  } finally { lock.releaseLock(); }
}


/* ---------- step 4: loans on the ledger ---------- */
// Loan sizes: extra charge and months to repay, number of guarantors. 5000 is an old size that is no longer given.
var LOAN_TIERS_ = { 5000:{ extra:200, months:2, g:0 }, 10000:{ extra:500, months:3, g:0 }, 25000:{ extra:1000, months:4, g:1 }, 50000:{ extra:2000, months:5, g:1 }, 100000:{ extra:4000, months:5, g:2 } };
var SUSPEND_MONTHS = 6;
function lastDay_(month){ var y = parseInt(month.slice(0, 4), 10), m = parseInt(month.slice(5, 7), 10); return month + '-' + pad_(new Date(y, m, 0).getDate(), 2); }
function addDays_(d, n){ var t = new Date(Date.UTC(+d.slice(0, 4), +d.slice(5, 7) - 1, +d.slice(8, 10)) + n * 86400000); return t.getUTCFullYear() + '-' + ('0' + (t.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + t.getUTCDate()).slice(-2); }
function dayDiff_(a, b){ return Math.round((Date.UTC(+b.slice(0, 4), +b.slice(5, 7) - 1, +b.slice(8, 10)) - Date.UTC(+a.slice(0, 4), +a.slice(5, 7) - 1, +a.slice(8, 10))) / 86400000); }
function loanClosedText_(l){ return /complete|closed|paid/i.test(String(l.Status || '')) || !!l.ClosedDate; }

// paid parts per loan, from LoanPayments that are not voided
function loanPaidIndex_(){
  var idx = {};
  rows_('LoanPayments').forEach(function(p){
    if(isTrue_(p.Voided)) return;
    var o = idx[p.LoanID] || (idx[p.LoanID] = { p:0, e:0, l:0, n:0 });
    o.p += num_(p.ToPrincipal); o.e += num_(p.ToExtra); o.l += num_(p.ToLateExtra); o.n++;
  });
  return idx;
}
// All the figures of one loan as of today. Status is worked out from dates, never stored.
function loanFig_(l, pi, today){
  var amt = num_(l.Amount), tier = LOAN_TIERS_[amt] || { extra:0 }, extra = num_(l.Fine) > 0 ? num_(l.Fine) : tier.extra;
  var pd = pi[l.LoanID] || { p:0, e:0, l:0, n:0 }, closed = loanClosedText_(l);
  var due = String(l.DueDate || ''), dm = due.slice(0, 7), od = dm ? dm + '-16' : '', lastD = dm ? lastDay_(dm) : '';
  var def = dm ? addMonths_(dm, 1) + '-01' : '';
  var legacyPaid = closed && pd.n === 0;
  var status = closed ? 'Closed' : (def && today >= def ? 'Defaulted' : (od && today >= od ? 'Overdue' : 'Active'));
  var late = isTrue_(l.LateWaived) ? 0 : (num_(l.LateExtra) > 0 ? num_(l.LateExtra) : ((status === 'Overdue' || status === 'Defaulted') ? extra : 0));
  var total = amt + extra + late, paid = legacyPaid ? num_(l.ReturnAmount) : pd.p + pd.e + pd.l;
  var oP = Math.max(0, amt - pd.p), oE = Math.max(0, extra - pd.e), oL = Math.max(0, late - pd.l);
  var out = closed ? 0 : oP + oE + oL;
  return { id:l.LoanID, memberId:l.MemberID || '', name:l.ApplicantName || '', amount:amt, extra:extra, late:late, lateWaived:isTrue_(l.LateWaived), total:total, paid:paid, outstanding:out, oP:closed ? 0 : oP, oE:closed ? 0 : oE, oL:closed ? 0 : oL,
    taken:String(l.TakenDate || ''), due:due, closedDate:String(l.ClosedDate || ''), status:status, overdueFrom:od, lastDay:lastD, daysLeft:(status === 'Overdue' && lastD) ? Math.max(0, dayDiff_(today, lastD)) : 0, defaultFrom:def, source:String(l.Source || ''), row:l._row };
}
function guarantorsOf_(){
  var g = {};
  rows_('Guarantors').forEach(function(x){ (g[x.LoanID] || (g[x.LoanID] = [])).push({ id:x.MemberID || '', name:x.Name || '' }); });
  return g;
}
function principalOut_(figs){ var t = 0; figs.forEach(function(f){ if(f.status !== 'Closed') t += f.oP; }); return t; }
function depositPot_(){ var d = depositTotals_(), t = 0; Object.keys(d).forEach(function(k){ t += d[k]; }); return t; }

// Once a day (first request of the day): late extra is fixed and the 6-month loan suspension is recorded for every overdue loan.
// Writes one column for several rows in a single call. values = { sheetRowNumber: newValue }.
function setColumn_(name, col, values){
  var sh = tab_(name), last = sh.getLastRow(), ci = TABS[name].indexOf(col) + 1;
  if(last < 2 || ci < 1) return 0;
  var rng = sh.getRange(2, ci, last - 1, 1), v = rng.getValues(), n = 0;
  Object.keys(values).forEach(function(r){ var i = Number(r) - 2; if(i >= 0 && i < v.length){ v[i][0] = values[r]; n++; } });
  rng.setValues(v); delete ROWS_[name]; markDirty_();
  return n;
}
function loanSweep_(S, force){
  var today = todayStr_(), pr = props_();
  if(!force && pr.getProperty('sweepDate') === today) return 0;
  var pi = loanPaidIndex_(), sus = rows_('Suspensions'), have = {}, n = 0, at = nowIso_(), news = [], lateSet = {};
  sus.forEach(function(x){ if(x.LoanID) have[x.LoanID] = 1; });
  rows_('Loans').forEach(function(l){
    var f = loanFig_(l, pi, today);
    if(f.status !== 'Overdue' && f.status !== 'Defaulted') return;
    if(!(num_(l.LateExtra) > 0) && !isTrue_(l.LateWaived) && f.extra > 0){ lateSet[l._row] = f.extra; n++; }
    if(!have[f.id] && f.memberId){ news.push({ SuspensionID:newId_('S'), MemberID:f.memberId, Reason:'Overdue loan', LoanID:f.id, StartDate:f.overdueFrom, EndDate:addMonths_(f.overdueFrom.slice(0, 7), SUSPEND_MONTHS) + '-' + f.overdueFrom.slice(8, 10) }); n++; }
  });
  if(n) setColumn_('Loans', 'LateExtra', lateSet);      // one write for all loans, not one per loan
  if(news.length) appendMany_('Suspensions', news);
  pr.setProperty('sweepDate', today);
  if(n) audit_(S || { u:'system', r:'' }, 'overdue-rules-applied', 'loans', n + ' change(s)');
  return n;
}

function loanView_(f, gs, extraLists){
  var o = f; o.guarantors = gs || [];
  return o;
}

function loansList_(S, body){
  var bad = needDb_(); if(bad) return bad;
  loanSweep_(S, false);
  var today = todayStr_(), pi = loanPaidIndex_(), gi = guarantorsOf_(), figs = [], closedN = 0;
  rows_('Loans').forEach(function(l){
    if(!num_(l.Amount)) return;
    var f = loanFig_(l, pi, today);
    if(f.status === 'Closed'){ closedN++; return; }
    loanView_(f, gi[f.id]); figs.push(f);
  });
  var order = { Defaulted:0, Overdue:1, Active:2 };
  figs.sort(function(a, b){ return (order[a.status] - order[b.status]) || a.due.localeCompare(b.due) || (a.name < b.name ? -1 : 1); });
  var pot = depositPot_(), out = principalOut_(figs), counts = { Active:0, Overdue:0, Defaulted:0 };
  figs.forEach(function(f){ counts[f.status]++; });
  return out_({ ok:true, today:today, loans:figs, counts:counts, closedCount:closedN, principalOut:out, depositPot:pot, available:pot - out,
    toCollect:figs.reduce(function(s2, f){ return s2 + f.outstanding; }, 0) });
}

// Loans by repayment month for the Monthly > Loan repayment screen.
// Returns a summary for every month that has loans, and the full list (with payments) for body.month (default: this month).
function loansDue_(S, body){
  var bad = needDb_(); if(bad) return bad;
  loanSweep_(S, false);
  var month = String(body.month || curMonth_()); if(!validMonth_(month)) return out_({ ok:false, error:'bad-months' });
  var today = todayStr_(), pi = loanPaidIndex_(), gi = guarantorsOf_(), summary = {}, list = [];
  rows_('Loans').forEach(function(l){
    if(!num_(l.Amount)) return;
    var f = loanFig_(l, pi, today), m = String(f.due || '').slice(0, 7); if(!m) return;
    var sm = summary[m] || (summary[m] = { total:0, received:0, toCollect:0 });
    sm.total++; if(f.status === 'Closed') sm.received++; else sm.toCollect += f.outstanding;
    if(m === month){ f.guarantors = (gi[f.id] || []).map(function(g){ return g.name; }); list.push(f); }
  });
  var pays = {};
  rows_('LoanPayments').forEach(function(p){
    if(isTrue_(p.Voided)) return;
    (pays[p.LoanID] || (pays[p.LoanID] = [])).push({ receipt:String(p.Ref || ''), date:String(p.Date || ''), amount:num_(p.Amount), paidBy:p.PaidBy || '', mode:p.Mode || '', by:p.EnteredBy || '', at:String(p.EnteredAt || '') });
  });
  list.forEach(function(f){ f.payments = (pays[f.id] || []).sort(function(a, b){ return a.at < b.at ? -1 : 1; }); });
  list.sort(function(a, b){ return (a.status === 'Closed') - (b.status === 'Closed') || (a.name < b.name ? -1 : 1); });
  return out_({ ok:true, month:month, today:today, summary:summary, loans:list });
}

function loanDetail_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var l = rows_('Loans').filter(function(x){ return String(x.LoanID) === String(body.loanId); })[0];
  if(!l) return out_({ ok:false, error:'not-found' });
  var today = todayStr_(), f = loanFig_(l, loanPaidIndex_(), today), gs = guarantorsOf_()[f.id] || [];
  var pays = rows_('LoanPayments').filter(function(p){ return p.LoanID === f.id && !isTrue_(p.Voided); }).map(function(p){
    return { receipt:p.Ref && /^R-/.test(String(p.Ref)) ? p.Ref : (p.PaymentID || ''), paymentId:p.PaymentID, date:String(p.Date || ''), amount:num_(p.Amount), by:p.EnteredBy || '', at:p.EnteredAt || '', paidBy:p.PaidBy || '', mode:p.Mode || '', toP:num_(p.ToPrincipal), toE:num_(p.ToExtra), toL:num_(p.ToLateExtra) };
  });
  pays.sort(function(a, b){ return String(b.at).localeCompare(String(a.at)); });
  var sus = rows_('Suspensions').filter(function(x){ return x.LoanID === f.id; }).map(function(x){ return { id:x.SuspensionID, start:String(x.StartDate || ''), end:String(x.EndDate || ''), lifted:x.LiftedBy || '' }; });
  var dep = depositTotals_();
  gs = gs.map(function(g){ return { id:g.id, name:g.name, deposit:g.id ? (dep[g.id] || 0) : 0 }; });
  var res = { ok:true, loan:f, guarantors:gs, payments:pays.slice(0, 30), suspensions:sus, borrowerDeposit:f.memberId ? (dep[f.memberId] || 0) : 0 };
  return out_(res);
}

// A payment on a loan, from the borrower or a guarantor. Goes to principal first, then extra, then late extra.
function loanPay_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var l = rows_('Loans').filter(function(x){ return String(x.LoanID) === String(body.loanId); })[0];
    if(!l) return out_({ ok:false, error:'not-found' });
    var today = todayStr_(), date = body.date ? String(body.date) : today;
    if(!validDate_(date) || date > today) return out_({ ok:false, error:'bad-date' });
    var mode = ['Bank','UPI'].indexOf(body.mode) >= 0 ? body.mode : '';
    if(!mode) return out_({ ok:false, error:'bad-mode' });
    var paidBy = body.paidBy === 'guarantor' ? 'guarantor' : 'borrower';
    var f = loanFig_(l, loanPaidIndex_(), today);
    if(f.status === 'Closed') return out_({ ok:false, error:'loan-closed' });
    var amount = Math.round(num_(body.amount) * 100) / 100;
    if(!(amount > 0)) return out_({ ok:false, error:'bad-amount' });
    if(amount > f.outstanding + 0.001) return out_({ ok:false, error:'too-much', due:f.outstanding });
    if(monthClosed_(date.slice(0, 7))) return out_({ ok:false, error:'month-closed', month:date.slice(0, 7) });
    var left = amount, toP = Math.min(left, f.oP); left -= toP; var toE = Math.min(left, f.oE); left -= toE; var toL = Math.min(left, f.oL);
    var no = nextReceipt_(), at = nowIso_(), ref = cleanRef_(body.ref), pid = newId_('LP');
    appendMany_('LoanPayments', [{ PaymentID:pid, LoanID:f.id, Date:date, Amount:amount, PaidBy:paidBy, ToPrincipal:toP, ToExtra:toE, ToLateExtra:toL, Mode:mode, Ref:no, EnteredBy:S.u, EnteredAt:at, Voided:'' }]);
    var ls = [];
    if(toE > 0) ls.push({ LedgerID:newId_('L'), Date:date, Pot:'fine', Type:'loan_extra', MemberID:f.memberId, LoanID:f.id, Amount:toE, Ref:pid, Note:'Loan extra ' + f.id + ' · ' + mode + (ref ? ' · ' + ref : ''), EnteredBy:S.u, EnteredAt:at, ReceiptNo:no });
    if(toL > 0) ls.push({ LedgerID:newId_('L'), Date:date, Pot:'fine', Type:'loan_late_extra', MemberID:f.memberId, LoanID:f.id, Amount:toL, Ref:pid, Note:'Late extra ' + f.id + ' · ' + mode + (ref ? ' · ' + ref : ''), EnteredBy:S.u, EnteredAt:at, ReceiptNo:no });
    if(ls.length) appendMany_('Ledger', ls);
    var upd = {}; if(f.late > 0 && !(num_(l.LateExtra) > 0)) upd.LateExtra = f.late;
    var closes = amount >= f.outstanding - 0.001;
    if(closes){ upd.Status = 'Complete'; upd.ClosedDate = date; }
    if(Object.keys(upd).length) update_('Loans', l._row, upd);
    audit_(S, 'loan-payment', f.id, no + ' · ' + amount + ' · ' + paidBy + ' · ' + mode + (closes ? ' · closed' : ''));
    return out_({ ok:true, receipt:{ no:no, memberId:f.memberId, name:f.name, months:[f.id], amount:amount, date:date, mode:mode, ref:ref, by:S.u, kind:'loan', loanId:f.id, left:Math.max(0, f.outstanding - amount), closed:closes, split:{ p:toP, e:toE, l:toL } } });
  } finally { lock.releaseLock(); }
}

// Void a loan payment (by receipt). withusadmin1 any time; the admin who made it, only within UNDO_SECONDS.
function loanPayVoid_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var key = String(body.receipt || ''); if(!key) return out_({ ok:false, error:'bad-request' });
    var pays = rows_('LoanPayments').filter(function(p){ return !isTrue_(p.Voided) && String(p.Ref) === key; });
    if(!pays.length) return out_({ ok:false, error:'not-found' });
    var nowMs = Date.now(), at = nowIso_();
    for(var i = 0; i < pays.length; i++){
      if(monthClosed_(String(pays[i].Date).slice(0, 7))) return out_({ ok:false, error:'month-closed', month:String(pays[i].Date).slice(0, 7) });
      if(S.r !== 'admin1'){
        var t = Date.parse(String(pays[i].EnteredAt) + '+05:30');
        if(String(pays[i].EnteredBy) !== S.u || !isFinite(t) || nowMs - t > UNDO_SECONDS * 1000) return out_({ ok:false, error:'forbidden' });
      }
    }
    var loanId = pays[0].LoanID;
    pays.forEach(function(p){ update_('LoanPayments', p._row, { Voided:true }); });
    rows_('Ledger').forEach(function(l){ if(!isTrue_(l.Voided) && String(l.ReceiptNo) === key) update_('Ledger', l._row, { Voided:true, VoidedBy:S.u, VoidedAt:at }); });
    var loan = rows_('Loans').filter(function(x){ return String(x.LoanID) === String(loanId); })[0];
    if(loan && loanClosedText_(loan) && String(loan.Source) !== 'import'){ update_('Loans', loan._row, { Status:'In Process', ClosedDate:'' }); }
    else if(loan && String(loan.Source) === 'import' && loanClosedText_(loan) && loan.ClosedDate){ update_('Loans', loan._row, { Status:'In Process', ClosedDate:'' }); }
    audit_(S, 'loan-payment-voided', loanId, key + ' · ' + String(body.reason || '').slice(0, 100));
    return out_({ ok:true, voided:pays.length });
  } finally { lock.releaseLock(); }
}

// admin1 only: take the late extra off one loan, with a reason.
function loanWaiveLate_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var reason = String(body.reason || '').replace(/\s+/g, ' ').trim().slice(0, 120);
    if(reason.length < 3) return out_({ ok:false, error:'reason-needed' });
    var l = rows_('Loans').filter(function(x){ return String(x.LoanID) === String(body.loanId); })[0];
    if(!l) return out_({ ok:false, error:'not-found' });
    var f = loanFig_(l, loanPaidIndex_(), todayStr_());
    if(f.status === 'Closed') return out_({ ok:false, error:'loan-closed' });
    var pd = loanPaidIndex_()[f.id]; if(pd && pd.l > 0) return out_({ ok:false, error:'late-already-paid' });
    update_('Loans', l._row, { LateWaived:true, LateExtra:0, WaiveNote:reason });
    audit_(S, 'late-extra-waived', f.id, reason);
    return out_({ ok:true });
  } finally { lock.releaseLock(); }
}

// admin1 only: end a loan suspension early, with a reason.
function suspensionLift_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var reason = String(body.reason || '').replace(/\s+/g, ' ').trim().slice(0, 120);
  if(reason.length < 3) return out_({ ok:false, error:'reason-needed' });
  var s2 = rows_('Suspensions').filter(function(x){ return String(x.SuspensionID) === String(body.id); })[0];
  if(!s2) return out_({ ok:false, error:'not-found' });
  update_('Suspensions', s2._row, { LiftedBy:S.u });
  audit_(S, 'suspension-lifted', s2.MemberID, (s2.LoanID || '') + ' · ' + reason);
  return out_({ ok:true });
}

// The settlement plan for a Defaulted loan: borrower's deposit first, then the guarantors equally (a short one is covered by the rest).
function settlePlan_(f, gs, dep){
  var owed = f.outstanding, plan = [], left = owed;
  if(f.memberId){ var b = Math.min(left, Math.max(0, dep[f.memberId] || 0)); if(b > 0){ plan.push({ id:f.memberId, name:f.name, role:'borrower', amount:b }); left -= b; } }
  var gm = gs.filter(function(g){ return g.id && g.id !== f.memberId; }), room = {};
  gm.forEach(function(g){ room[g.id] = Math.max(0, dep[g.id] || 0); });
  var take = {}; gm.forEach(function(g){ take[g.id] = 0; });
  var active = gm.slice();
  while(left > 0.001 && active.length){
    var share = Math.floor(left / active.length * 100) / 100 || left, next = [];
    var progressed = false;
    active.forEach(function(g){
      var can = room[g.id] - take[g.id], x = Math.min(share, can, left);
      if(x > 0){ take[g.id] += x; left -= x; progressed = true; }
      if(room[g.id] - take[g.id] > 0.001) next.push(g);
    });
    if(!progressed) break; active = next;
  }
  gm.forEach(function(g){ if(take[g.id] > 0) plan.push({ id:g.id, name:g.name, role:'guarantor', amount:Math.round(take[g.id] * 100) / 100 }); });
  var covered = plan.reduce(function(s2, p){ return s2 + p.amount; }, 0);
  return { plan:plan, owed:owed, covered:covered, short:Math.max(0, Math.round((owed - covered) * 100) / 100) };
}
function settlePreview_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var l = rows_('Loans').filter(function(x){ return String(x.LoanID) === String(body.loanId); })[0];
  if(!l) return out_({ ok:false, error:'not-found' });
  var f = loanFig_(l, loanPaidIndex_(), todayStr_());
  if(f.status !== 'Defaulted') return out_({ ok:false, error:'not-defaulted' });
  var r = settlePlan_(f, guarantorsOf_()[f.id] || [], depositTotals_());
  r.ok = true; r.loan = f; return out_(r);
}
function settleConfirm_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var l = rows_('Loans').filter(function(x){ return String(x.LoanID) === String(body.loanId); })[0];
    if(!l) return out_({ ok:false, error:'not-found' });
    var today = todayStr_(), f = loanFig_(l, loanPaidIndex_(), today);
    if(f.status !== 'Defaulted') return out_({ ok:false, error:'not-defaulted' });
    if(monthClosed_(today.slice(0, 7))) return out_({ ok:false, error:'month-closed', month:today.slice(0, 7) });
    var r = settlePlan_(f, guarantorsOf_()[f.id] || [], depositTotals_());
    if(!r.plan.length) return out_({ ok:false, error:'nothing-to-take' });
    var no = nextReceipt_(), at = nowIso_(), left = { p:f.oP, e:f.oE, l:f.oL }, lp = [], ls = [], total = 0;
    r.plan.forEach(function(p){
      var a = p.amount, toP = Math.min(a, left.p); left.p -= toP; var toE = Math.min(a - toP, left.e); left.e -= toE; var toL = Math.min(a - toP - toE, left.l); left.l -= toL;
      var pid = newId_('LP'); total += a;
      lp.push({ PaymentID:pid, LoanID:f.id, Date:today, Amount:a, PaidBy:'deposit:' + p.id, ToPrincipal:toP, ToExtra:toE, ToLateExtra:toL, Mode:'Deposit', Ref:no, EnteredBy:S.u, EnteredAt:at, Voided:'' });
      ls.push({ LedgerID:newId_('L'), Date:today, Pot:'deposit', Type:'settlement', MemberID:p.id, LoanID:f.id, Amount:-a, Ref:pid, Note:p.role === 'guarantor' ? 'Deducted as guarantor for ' + f.name + ' (' + f.id + ')' : 'Deducted from own deposit for ' + f.id, EnteredBy:S.u, EnteredAt:at, ReceiptNo:no });
      if(toE > 0) ls.push({ LedgerID:newId_('L'), Date:today, Pot:'fine', Type:'loan_extra', MemberID:p.id, LoanID:f.id, Amount:toE, Ref:pid, Note:'Loan extra ' + f.id + ' (settlement)', EnteredBy:S.u, EnteredAt:at, ReceiptNo:no });
      if(toL > 0) ls.push({ LedgerID:newId_('L'), Date:today, Pot:'fine', Type:'loan_late_extra', MemberID:p.id, LoanID:f.id, Amount:toL, Ref:pid, Note:'Late extra ' + f.id + ' (settlement)', EnteredBy:S.u, EnteredAt:at, ReceiptNo:no });
    });
    appendMany_('LoanPayments', lp); appendMany_('Ledger', ls);
    var upd = {}; if(f.late > 0 && !(num_(l.LateExtra) > 0)) upd.LateExtra = f.late;
    var closes = r.short <= 0.001; if(closes){ upd.Status = 'Complete'; upd.ClosedDate = today; }
    if(Object.keys(upd).length) update_('Loans', l._row, upd);
    audit_(S, 'loan-settled', f.id, no + ' · ' + total + ' from ' + r.plan.map(function(p){ return p.id + ':' + p.amount; }).join(', ') + (closes ? ' · closed' : ' · short ' + r.short));
    return out_({ ok:true, receipt:no, taken:total, short:r.short, closed:closes, plan:r.plan });
  } finally { lock.releaseLock(); }
}

// Give a new loan (disbursement). Checks the same rules as the Eligibility screen.
function loanTierFor_(amount){ var t = LOAN_TIERS_[amount]; return (t && amount >= 10000) ? t : null; }
function loanDisburse_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var mem = memberById_(body.memberId); if(!mem) return out_({ ok:false, error:'no-member' });
    var amount = Math.round(num_(body.amount)), tier = loanTierFor_(amount); if(!tier) return out_({ ok:false, error:'bad-amount' });
    var today = todayStr_(), date = body.date ? String(body.date) : today;
    if(!validDate_(date) || date > today) return out_({ ok:false, error:'bad-date' });
    if(monthClosed_(date.slice(0, 7))) return out_({ ok:false, error:'month-closed', month:date.slice(0, 7) });
    var cur = curMonth_(), ci = contribIndex_(), pi = penaltyIndex_(), si = suspensionIndex_();
    var el = eligibility_(mem, standing_(mem, ci[mem.MemberID], cur), (pi[mem.MemberID] || { due:0 }).due, si[mem.MemberID]);
    if(!el.canApply) return out_({ ok:false, error:'not-eligible', reasons:el.reasons });
    var figs = rows_('Loans').map(function(l){ return loanFig_(l, loanPaidIndex_(), today); }), gi = guarantorsOf_();
    for(var i = 0; i < figs.length; i++) if(figs[i].memberId === mem.MemberID && figs[i].status !== 'Closed') return out_({ ok:false, error:'open-loan', loanId:figs[i].id });
    var gids = Array.isArray(body.guarantors) ? body.guarantors.map(String) : [];
    gids = gids.filter(function(x, k){ return x && gids.indexOf(x) === k; });
    if(gids.length !== tier.g) return out_({ ok:false, error:'bad-guarantors', need:tier.g });
    var names = [];
    for(var k = 0; k < gids.length; k++){
      var gm = memberById_(gids[k]); if(!gm || gm.MemberID === mem.MemberID) return out_({ ok:false, error:'bad-guarantors', need:tier.g });
      var ge = eligibility_(gm, standing_(gm, ci[gm.MemberID], cur), (pi[gm.MemberID] || { due:0 }).due, si[gm.MemberID]);
      if(!ge.canGuarantee) return out_({ ok:false, error:'guarantor-blocked', member:gm.MemberID, name:gm.Name, reasons:ge.reasons });
      for(var j = 0; j < figs.length; j++) if(figs[j].status !== 'Closed' && (gi[figs[j].id] || []).some(function(g){ return g.id === gm.MemberID; })) return out_({ ok:false, error:'guarantor-busy', member:gm.MemberID, name:gm.Name, loanId:figs[j].id });
      names.push(gm.Name);
    }
    var pot = depositPot_(), out = principalOut_(figs);
    if(amount > pot - out + 0.001) return out_({ ok:false, error:'no-funds', available:pot - out });
    var n = 0; rows_('Loans').forEach(function(l){ var m = /^L-(\d+)$/.exec(String(l.LoanID)); if(m) n = Math.max(n, parseInt(m[1], 10)); });
    var lid = 'L-' + pad_(n + 1, 4), due = addMonths_(date.slice(0, 7), tier.months) + '-' + date.slice(8, 10);
    appendMany_('Loans', [{ LoanID:lid, MemberID:mem.MemberID, ApplicantName:mem.Name, Amount:amount, TakenDate:date, DueDate:due, ReturnAmount:0, Fine:tier.extra, LateExtra:0, Status:'In Process', Suspended:'', ClosedDate:'', Source:'app', CreatedAt:nowIso_() }]);
    if(gids.length) appendMany_('Guarantors', gids.map(function(g, q){ return { LoanID:lid, MemberID:g, Name:names[q], Source:'app' }; }));
    audit_(S, 'loan-disbursed', lid, mem.MemberID + ' · ' + amount + ' · guarantors ' + (gids.join(', ') || 'none'));
    return out_({ ok:true, loanId:lid, due:due, extra:tier.extra, total:amount + tier.extra });
  } finally { lock.releaseLock(); }
}


/* ---------- step 5: fine pot and expenses ---------- */
var EXPENSE_CATS_ = ['Bank charges','Stationery','Printing','Meeting','Travel','Donation','Other'];
function expenseCats_(){
  var v = String(setting_('ExpenseCategories', '')).split(',').map(function(x){ return x.trim(); }).filter(Boolean);
  return v.length ? v : EXPENSE_CATS_;
}
function fyStartMonth_(){ var m = Math.round(num_(setting_('FinancialYearStartMonth', 4))); return (m >= 1 && m <= 12) ? m : 4; }
// the financial year that contains a month 'YYYY-MM': { start:'YYYY-MM', end:'YYYY-MM', label:'2026-27' }
function fyOf_(month){
  var y = parseInt(month.slice(0, 4), 10), m = parseInt(month.slice(5, 7), 10), sm = fyStartMonth_(), sy = m >= sm ? y : y - 1;
  var start = sy + '-' + pad_(sm, 2), end = addMonths_(start, 11);
  return { start:start, end:end, label:sm === 1 ? String(sy) : sy + '-' + String(sy + 1).slice(2) };
}
var FINE_IN_TYPES_ = ['loan_extra','loan_late_extra','penalty_paid','joining_fee','opening'];
function fineRows_(){
  var members = {}; rows_('Members').forEach(function(m){ members[m.MemberID] = m.Name; });
  var list = rows_('Ledger').filter(function(l){ return l.Pot === 'fine' && !isTrue_(l.Voided); });
  list.sort(function(a, b){ return String(a.Date).localeCompare(String(b.Date)) || (a._row - b._row); });
  var bal = 0;
  return list.map(function(l){
    var amt = num_(l.Amount); bal += amt;
    var isExp = l.Type === 'expense';
    var cat = '', note = String(l.Note || '');
    if(isExp && /^[^·]+ · /.test(note)){ cat = note.split(' · ')[0]; }
    return { id:l.LedgerID, date:String(l.Date || ''), type:l.Type, amount:amt, balance:Math.round(bal * 100) / 100, note:note, category:isExp ? (cat || (String(l.EnteredBy) === 'import' ? 'Imported' : 'Other')) : '', member:l.MemberID ? (members[l.MemberID] || '') : '', memberId:l.MemberID || '', by:l.EnteredBy || '', ref:l.Ref || '', expenseId:isExp && /^X-/.test(String(l.Ref)) ? l.Ref : '', receipt:l.ReceiptNo || '' };
  });
}
function finePot_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var all = fineRows_(), today = todayStr_(), cur = today.slice(0, 7), fy = fyOf_(cur);
  var balance = all.length ? all[all.length - 1].balance : 0;
  function bucket(rows){
    var inn = {}, outt = {}, ti = 0, to = 0;
    rows.forEach(function(r){
      if(r.amount >= 0 && r.type !== 'expense'){ inn[r.type] = (inn[r.type] || 0) + r.amount; ti += r.amount; }
      else { var k = r.type === 'expense' ? r.category : r.type; outt[k] = (outt[k] || 0) - r.amount; to -= r.amount; }
    });
    return { inn:inn, out:outt, totalIn:ti, totalOut:to };
  }
  var inFy = all.filter(function(r){ var m = r.date.slice(0, 7); return m >= fy.start && m <= fy.end; });
  var months = {}, order = [];
  all.forEach(function(r){ var m = r.date.slice(0, 7); if(!months[m]){ months[m] = []; order.push(m); } months[m].push(r); });
  order = order.filter(function(m, i){ return order.indexOf(m) === i; }).sort().reverse();
  var monthRows = order.slice(0, 18).map(function(m){ var b = bucket(months[m]), last = months[m][months[m].length - 1]; return { month:m, inn:b.inn, out:b.out, totalIn:b.totalIn, totalOut:b.totalOut, closing:last.balance }; });
  var limit = Math.min(200, Math.max(10, Math.round(num_(body && body.limit)) || 60)), recent = all.slice().reverse().slice(0, limit);
  return out_({ ok:true, balance:balance, today:today, fy:{ label:fy.label, start:fy.start, end:fy.end }, thisFy:bucket(inFy), allTime:bucket(all), months:monthRows, rows:recent, count:all.length, categories:expenseCats_() });
}

function expenseAdd_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var today = todayStr_(), date = body.date ? String(body.date) : today;
    if(!validDate_(date) || date > today) return out_({ ok:false, error:'bad-date' });
    var cat = String(body.category || '').trim(); if(expenseCats_().indexOf(cat) < 0) return out_({ ok:false, error:'bad-category' });
    var amount = Math.round(num_(body.amount) * 100) / 100; if(!(amount > 0)) return out_({ ok:false, error:'bad-amount' });
    var note = String(body.note || '').replace(/\s+/g, ' ').trim().slice(0, 100);
    if(monthClosed_(date.slice(0, 7))) return out_({ ok:false, error:'month-closed', month:date.slice(0, 7) });
    var all = fineRows_(), bal = all.length ? all[all.length - 1].balance : 0;
    if(amount > bal + 0.001) return out_({ ok:false, error:'exceeds-fine-pot', balance:bal });
    var id = newId_('X'), at = nowIso_();
    appendMany_('Expenses', [{ ExpenseID:id, Date:date, Category:cat, Amount:amount, Note:note, EnteredBy:S.u, EnteredAt:at, Voided:'' }]);
    appendMany_('Ledger', [{ LedgerID:newId_('L'), Date:date, Pot:'fine', Type:'expense', Amount:-amount, Ref:id, Note:cat + ' · ' + (note || '-'), EnteredBy:S.u, EnteredAt:at }]);
    audit_(S, 'expense-added', id, cat + ' · ' + amount + (note ? ' · ' + note : ''));
    return out_({ ok:true, id:id, balance:Math.round((bal - amount) * 100) / 100 });
  } finally { lock.releaseLock(); }
}

function expenseVoid_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var reason = String(body.reason || '').replace(/\s+/g, ' ').trim().slice(0, 120);
    if(reason.length < 3) return out_({ ok:false, error:'reason-needed' });
    var led = rows_('Ledger').filter(function(l){ return l.Pot === 'fine' && l.Type === 'expense' && !isTrue_(l.Voided) && String(l.LedgerID) === String(body.id); })[0];
    if(!led) return out_({ ok:false, error:'not-found' });
    if(monthClosed_(String(led.Date).slice(0, 7))) return out_({ ok:false, error:'month-closed', month:String(led.Date).slice(0, 7) });
    var at = nowIso_();
    update_('Ledger', led._row, { Voided:true, VoidedBy:S.u, VoidedAt:at });
    if(/^X-/.test(String(led.Ref))){
      var ex = rows_('Expenses').filter(function(x){ return String(x.ExpenseID) === String(led.Ref); })[0];
      if(ex) update_('Expenses', ex._row, { Voided:true });
    }
    audit_(S, 'expense-voided', led.Ref || led.LedgerID, num_(led.Amount) + ' · ' + String(led.Note).slice(0, 60) + ' · ' + reason);
    return out_({ ok:true });
  } finally { lock.releaseLock(); }
}

/* ---------- step 6: exit settlement, bank check, month close, reports ---------- */
function r2_(x){ return Math.round(num_(x) * 100) / 100; }
function exitPlan_(mem){
  var today = todayStr_(), cur = curMonth_(), pi = loanPaidIndex_(), gi = guarantorsOf_(), blocks = [];
  rows_('Loans').forEach(function(l){
    if(!num_(l.Amount)) return;
    var f = loanFig_(l, pi, today); if(f.status === 'Closed') return;
    if(f.memberId === mem.MemberID) blocks.push({ type:'open-loan', loanId:f.id, amount:f.outstanding });
    else if((gi[f.id] || []).some(function(g){ return g.id === mem.MemberID; })) blocks.push({ type:'guarantor', loanId:f.id, borrower:f.name });
  });
  if(String(mem.Status) === 'Exited') blocks.push({ type:'exited' });
  var gross = r2_(depositTotals_()[mem.MemberID] || 0), amt = contribAmount_();
  var ci = contribIndex_()[mem.MemberID] || {}, arrears = [];
  if(String(mem.Status) !== 'Exited'){
    var st = standing_(mem, ci, cur); arrears = st.arrears.slice(); if(st.due) arrears.push(cur);
  }
  var pen = penaltyIndex_()[mem.MemberID] || { due:0, list:[] };
  var penList = pen.list.filter(function(x){ return x.outstanding > 0; });
  var penDue = r2_(pen.due), arrAmt = r2_(arrears.length * amt);
  var penTake = Math.min(penDue, gross), arrTake = Math.min(arrAmt, Math.max(0, gross - penTake));
  var refund = r2_(Math.max(0, gross - penTake - arrTake));
  return { blocks:blocks, gross:gross, penalties:penList, penaltyDue:penDue, arrearsMonths:arrears, arrearsAmount:arrAmt, penTake:r2_(penTake), arrTake:r2_(arrTake), refund:refund, shortfall:r2_((penDue + arrAmt) - penTake - arrTake) };
}
function exitPreview_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var mem = memberById_(body.id); if(!mem) return out_({ ok:false, error:'no-member' });
  var p = exitPlan_(mem); p.ok = true; p.member = { id:mem.MemberID, name:mem.Name, status:mem.Status, joinDate:String(mem.JoinDate || '') };
  p.note = 'Joining fee, paid penalties and loan charges are not refunded.';
  return out_(p);
}
function exitConfirm_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var mem = memberById_(body.id); if(!mem) return out_({ ok:false, error:'no-member' });
    var today = todayStr_(), p = exitPlan_(mem);
    if(p.blocks.length) return out_({ ok:false, error:'blocked', blocks:p.blocks });
    if(body.confirmRefund !== undefined && Math.abs(num_(body.confirmRefund) - p.refund) > 0.009) return out_({ ok:false, error:'changed', refund:p.refund });
    if(monthClosed_(today.slice(0, 7))) return out_({ ok:false, error:'month-closed', month:today.slice(0, 7) });
    var no = nextReceipt_(), at = nowIso_(), ls = [], eid = newId_('E');
    var note = String(body.note || '').replace(/\s+/g, ' ').trim().slice(0, 100);
    if(p.refund > 0) ls.push({ LedgerID:newId_('L'), Date:today, Pot:'deposit', Type:'exit_refund', MemberID:mem.MemberID, Amount:-p.refund, Ref:eid, Note:'Exit refund ' + mem.Name, EnteredBy:S.u, EnteredAt:at, ReceiptNo:no });
    // unpaid penalties are cleared out of the deposit (oldest first) and become fine-pot income
    var left = p.penTake;
    p.penalties.forEach(function(x){
      if(left <= 0) return;
      var a = r2_(Math.min(x.outstanding, left)); left = r2_(left - a);
      ls.push({ LedgerID:newId_('L'), Date:today, Pot:'fine', Type:'penalty_paid', MemberID:mem.MemberID, Amount:a, Ref:x.id, Note:'Penalty ' + x.month + ' · deducted at exit', EnteredBy:S.u, EnteredAt:at, ReceiptNo:no });
    });
    if(p.penTake > 0) ls.push({ LedgerID:newId_('L'), Date:today, Pot:'deposit', Type:'exit_deduction', MemberID:mem.MemberID, Amount:-p.penTake, Ref:eid, Note:'Penalties deducted at exit', EnteredBy:S.u, EnteredAt:at, ReceiptNo:no });
    if(ls.length) appendMany_('Ledger', ls);
    appendMany_('Exits', [{ ExitID:eid, MemberID:mem.MemberID, Date:today, Refund:p.refund, Note:('deposit ' + p.gross + ' · penalties ' + p.penTake + ' · unpaid months ' + p.arrearsMonths.length + ' (' + p.arrTake + ')' + (note ? ' · ' + note : '')).slice(0, 200), EnteredBy:S.u, EnteredAt:at }]);
    update_('Members', mem._row, { Status:'Exited' });
    audit_(S, 'member-exited', mem.MemberID, mem.Name + ' · refund ' + p.refund + ' · ' + no);
    return out_({ ok:true, exitId:eid, receipt:no, refund:p.refund, deposit:p.gross, penalties:p.penTake, unpaidMonths:p.arrearsMonths, unpaidAmount:p.arrTake, shortfall:p.shortfall, name:mem.Name, date:today });
  } finally { lock.releaseLock(); }
}

// Money that should physically be held: deposit pot + fine pot - loan principal still out.
function expectedHeld_(){
  var today = todayStr_(), pi = loanPaidIndex_(), figs = [];
  rows_('Loans').forEach(function(l){ if(num_(l.Amount)) figs.push(loanFig_(l, pi, today)); });
  var dep = depositPot_(), fine = 0;
  rows_('Ledger').forEach(function(l){ if(l.Pot === 'fine' && !isTrue_(l.Voided)) fine += num_(l.Amount); });
  var out = principalOut_(figs);
  return { deposit:r2_(dep), fine:r2_(fine), principalOut:r2_(out), expected:r2_(dep + fine - out) };
}
function closeView_(S){
  var bad = needDb_(); if(bad) return bad;
  var ex = expectedHeld_(), checks = rows_('Reconciliations').map(function(r){ return { id:r.ReconID, date:String(r.Date || ''), bank:num_(r.BankBalance), cash:num_(r.CashInHand), expected:num_(r.ExpectedTotal), diff:num_(r.Difference), by:r.EnteredBy || '', at:r.EnteredAt || '' }; });
  checks.reverse();
  var closedMap = {}; rows_('MonthClose').forEach(function(m){ closedMap[String(m.Month)] = { by:m.ClosedBy || '', at:m.ClosedAt || '' }; });
  var cur = curMonth_(), ts = trackStart_(), months = [], lastCheck = checks.length ? checks[0].date : '';
  var latestDate = ''; checks.forEach(function(c){ if(c.date > latestDate) latestDate = c.date; });
  monthsBetween_(ts, addMonths_(cur, -1)).forEach(function(m){
    var c = closedMap[m];
    var pend = c ? 0 : penaltyCandidates_(m).length;
    months.push({ month:m, closed:!!c, by:c ? c.by : '', at:c ? c.at : '', penaltiesPending:pend, canClose:!c && latestDate >= m + '-01' && pend === 0 });
  });
  months.reverse();
  return out_({ ok:true, today:todayStr_(), expected:ex, checks:checks.slice(0, 24), months:months.slice(0, 36), lastCheck:lastCheck, canAct:S.r === 'admin1' });
}
function bankCheck_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var bank = r2_(body.bank), cash = r2_(body.cash);
  if(!(num_(body.bank) >= 0) || !(num_(body.cash) >= 0) || body.bank === '' || body.cash === '' || body.bank === null || body.cash === null) return out_({ ok:false, error:'bad-amount' });
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var ex = expectedHeld_(), id = newId_('RC'), today = todayStr_(), diff = r2_(bank + cash - ex.expected);
    appendMany_('Reconciliations', [{ ReconID:id, Date:today, BankBalance:bank, CashInHand:cash, ExpectedTotal:ex.expected, Difference:diff, EnteredBy:S.u, EnteredAt:nowIso_() }]);
    audit_(S, 'bank-check', id, 'bank ' + bank + ' + cash ' + cash + ' vs expected ' + ex.expected + ' = ' + (diff >= 0 ? '+' : '') + diff);
    return out_({ ok:true, id:id, expected:ex.expected, diff:diff, balanced:Math.abs(diff) < 0.005 });
  } finally { lock.releaseLock(); }
}
function monthClose_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var m = String(body.month || ''); if(!validMonth_(m)) return out_({ ok:false, error:'bad-month' });
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    if(m >= curMonth_()) return out_({ ok:false, error:'month-not-ended' });
    if(monthClosed_(m)) return out_({ ok:false, error:'already-closed' });
    var ok = rows_('Reconciliations').some(function(r){ return String(r.Date) >= m + '-01'; });
    if(!ok) return out_({ ok:false, error:'check-first' });
    var pend = penaltyCandidates_(m).length; if(pend) return out_({ ok:false, error:'penalties-pending', count:pend });
    appendMany_('MonthClose', [{ Month:m, ClosedBy:S.u, ClosedAt:nowIso_() }]);
    audit_(S, 'month-closed', m, '');
    return out_({ ok:true, month:m });
  } finally { lock.releaseLock(); }
}
function dashboard_(S){
  var bad = needDb_(); if(bad) return bad;
  loanSweep_(S, false);
  var today = todayStr_(), cur = today.slice(0, 7), pi = loanPaidIndex_(), figs = [];
  rows_('Loans').forEach(function(l){ if(num_(l.Amount)) figs.push(loanFig_(l, pi, today)); });
  var open = figs.filter(function(f){ return f.status !== 'Closed'; }), ex = expectedHeld_();
  var ci = contribIndex_(), defaulters = 0;
  rows_('Members').forEach(function(m){ if(String(m.Status) !== 'Exited' && standing_(m, ci[m.MemberID], cur).arrears.length) defaulters++; });
  return out_({ ok:true, depositPot:ex.deposit, finePot:ex.fine, loansOut:ex.principalOut, available:r2_(ex.deposit - ex.principalOut), defaulters:defaulters,
    overdueLoans:open.filter(function(f){ return f.status === 'Overdue' || f.status === 'Defaulted'; }).length, dueThisMonth:open.filter(function(f){ return f.due.slice(0, 7) === cur; }).length });
}

/* reports: read-only, both admins. Every report returns { title, cols:[labels], rows:[[...]], foot:[...] | null, numeric:[col indexes] } */
var REPORT_TYPES_ = ['collection','arrears','penalties','statement','members','loanbook','loanpay','guarantors','settlements','finepot','expenses','depositpot','bankchecks','fysummary','audit'];
function inRange_(d, from, to){ d = String(d || '').slice(0, 10); return (!from || d >= from) && (!to || d <= to); }
function sumCol_(rows, i){ return r2_(rows.reduce(function(s2, r){ return s2 + num_(r[i]); }, 0)); }
function report_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var type = String(body.type || ''); if(REPORT_TYPES_.indexOf(type) < 0) return out_({ ok:false, error:'bad-report' });
  var from = validDate_(body.from) ? String(body.from) : '', to = validDate_(body.to) ? String(body.to) : '';
  var today = todayStr_(), cur = today.slice(0, 7), names = {}, cols = [], rows = [], foot = null, numeric = [], title = '', sub = '';
  rows_('Members').forEach(function(m){ names[m.MemberID] = m; });
  var nm = function(id){ return names[id] ? names[id].Name : (id || ''); };
  if(type === 'collection'){
    var mo = validMonth_(String(body.month || '')) ? String(body.month) : cur; title = 'Monthly collection ' + mo; sub = 'Contribution due per member for the month';
    cols = ['Member ID','Name','Status','Paid','Amount','Date','Mode','Receipt']; numeric = [4];
    var ci = contribIndex_();
    rows_('Members').forEach(function(m){
      if(String(m.Status) === 'Exited' && !(ci[m.MemberID] && ci[m.MemberID][mo])) return;
      if(mo < startMonth_(m)) return;
      var c = ci[m.MemberID] && ci[m.MemberID][mo];
      rows.push([m.MemberID, m.Name, m.Status || 'Active', c ? 'Yes' : 'No', c ? c.amount : 0, c ? c.date : '', c ? c.mode : '', c ? c.receipt : '']);
    });
    foot = ['', 'Total', '', rows.filter(function(r){ return r[3] === 'Yes'; }).length + ' paid / ' + rows.length, sumCol_(rows, 4), '', '', ''];
  } else if(type === 'arrears'){
    title = 'Arrears as of ' + today; sub = 'Members with unpaid past months'; cols = ['Member ID','Name','Months','Count','Amount']; numeric = [3, 4];
    var ci2 = contribIndex_(), amt = contribAmount_();
    rows_('Members').forEach(function(m){
      if(String(m.Status) === 'Exited') return; var st = standing_(m, ci2[m.MemberID], cur);
      if(st.arrears.length) rows.push([m.MemberID, m.Name, st.arrears.join(', '), st.arrears.length, st.arrears.length * amt]);
    });
    rows.sort(function(a, b){ return b[3] - a[3]; });
    foot = ['', 'Total', '', sumCol_(rows, 3), sumCol_(rows, 4)];
  } else if(type === 'penalties'){
    title = 'Penalties'; sub = (from || to) ? 'For months ' + (from || '…').slice(0, 7) + ' to ' + (to || '…').slice(0, 7) : 'All months'; cols = ['Member ID','Name','Month','Charged','Paid','Outstanding','Status']; numeric = [3, 4, 5];
    var pi2 = penaltyIndex_();
    Object.keys(pi2).forEach(function(id){ pi2[id].list.forEach(function(p){
      if((from && p.month < from.slice(0, 7)) || (to && p.month > to.slice(0, 7))) return;
      rows.push([id, nm(id), p.month, p.amount, p.paid, p.outstanding, p.status]);
    }); });
    rows.sort(function(a, b){ return a[2].localeCompare(b[2]) || (a[1] < b[1] ? -1 : 1); });
    foot = ['', 'Total', '', sumCol_(rows, 3), sumCol_(rows, 4), sumCol_(rows, 5), ''];
  } else if(type === 'statement'){
    var mem = memberById_(body.memberId); if(!mem) return out_({ ok:false, error:'no-member' });
    title = 'Member statement · ' + mem.Name + ' (' + mem.MemberID + ')'; sub = 'Deposit and fine-pot movements'; cols = ['Date','Pot','Type','Amount','Deposit balance','Receipt','Note']; numeric = [3, 4];
    var bal = 0, list = rows_('Ledger').filter(function(l){ return l.MemberID === mem.MemberID && !isTrue_(l.Voided); });
    list.sort(function(a, b){ return String(a.Date).localeCompare(String(b.Date)) || (a._row - b._row); });
    list.forEach(function(l){ var a = num_(l.Amount); if(l.Pot === 'deposit') bal += a; if(inRange_(l.Date, from, to)) rows.push([String(l.Date || ''), l.Pot, l.Type, a, l.Pot === 'deposit' ? r2_(bal) : '', l.ReceiptNo || '', String(l.Note || '')]); });
    foot = ['', '', 'Deposit balance', '', r2_(bal), '', ''];
  } else if(type === 'members'){
    title = 'Member list'; sub = 'As of ' + today; cols = ['Member ID','Name','Mobile','Status','Joined','Deposit','Penalty due','Arrears (months)']; numeric = [5, 6, 7];
    var ci3 = contribIndex_(), di = depositTotals_(), pi3 = penaltyIndex_();
    rows_('Members').forEach(function(m){
      var st = standing_(m, ci3[m.MemberID], cur);
      rows.push([m.MemberID, m.Name, String(m.Mobile || ''), m.Status || 'Active', String(m.JoinDate || ''), r2_(di[m.MemberID] || 0), (pi3[m.MemberID] || { due:0 }).due, st.arrears.length]);
    });
    foot = ['', 'Total (' + rows.length + ')', '', '', '', sumCol_(rows, 5), sumCol_(rows, 6), sumCol_(rows, 7)];
  } else if(type === 'loanbook'){
    var stt = String(body.status || ''); title = 'Loan book' + (stt ? ' · ' + stt : ''); sub = 'As of ' + today; cols = ['Loan','Borrower','Taken','Due','Principal','Extra','Late extra','Paid','Outstanding','Status','Guarantors']; numeric = [4, 5, 6, 7, 8];
    var lpi = loanPaidIndex_(), gi = guarantorsOf_();
    rows_('Loans').forEach(function(l){
      if(!num_(l.Amount)) return; var f = loanFig_(l, lpi, today);
      if(stt === 'open' ? f.status === 'Closed' : (stt && f.status !== stt)) return;
      if(!inRange_(f.taken, from, to) && (from || to)) return;
      rows.push([f.id, f.name, f.taken, f.due, f.amount, f.extra, f.late, f.paid, f.outstanding, f.status, (gi[f.id] || []).map(function(g){ return g.name; }).join(', ')]);
    });
    rows.sort(function(a, b){ return String(b[2]).localeCompare(String(a[2])); });
    foot = ['', 'Total (' + rows.length + ')', '', '', sumCol_(rows, 4), sumCol_(rows, 5), sumCol_(rows, 6), sumCol_(rows, 7), sumCol_(rows, 8), '', ''];
  } else if(type === 'loanpay'){
    title = 'Loan payments'; sub = (from || to) ? (from || '…') + ' to ' + (to || '…') : 'All dates'; cols = ['Date','Loan','Borrower','Paid by','Amount','Principal','Extra','Late extra','Mode','Receipt']; numeric = [4, 5, 6, 7];
    var hasPay = {}, lname = {}; rows_('Loans').forEach(function(l){ lname[l.LoanID] = String(l.ApplicantName || '') || nm(l.MemberID); });
    rows_('LoanPayments').forEach(function(p){
      if(isTrue_(p.Voided)) return; hasPay[p.LoanID] = 1;
      if(!inRange_(p.Date, from, to)) return;
      var pb = String(p.PaidBy || ''); if(/^deposit:/.test(pb)) pb = 'Deposit of ' + nm(pb.slice(8));
      rows.push([String(p.Date || ''), p.LoanID, lname[p.LoanID] || '', pb, num_(p.Amount), num_(p.ToPrincipal), num_(p.ToExtra), num_(p.ToLateExtra), p.Mode || '', /^R-/.test(String(p.Ref)) ? p.Ref : '']);
    });
    // loans closed in the Excel history carry only the total returned, so they are listed once, on the closing (or due) date
    rows_('Loans').forEach(function(l){
      if(hasPay[l.LoanID] || !loanClosedText_(l)) return;
      var ret = num_(l.ReturnAmount), d0 = String(l.ClosedDate || l.DueDate || '').slice(0, 10); if(!(ret > 0) || !inRange_(d0, from, to)) return;
      var amt0 = num_(l.Amount), pr0 = Math.min(amt0, ret), ex0 = Math.max(0, ret - pr0);
      rows.push([d0, l.LoanID, String(l.ApplicantName || '') || nm(l.MemberID), 'Excel', ret, pr0, ex0, 0, '', 'Excel history']);
    });
    rows.sort(function(a, b){ return String(a[0]).localeCompare(String(b[0])); });
    foot = ['', 'Total (' + rows.length + ')', '', '', sumCol_(rows, 4), sumCol_(rows, 5), sumCol_(rows, 6), sumCol_(rows, 7), '', ''];
  } else if(type === 'guarantors'){
    title = 'Guarantor exposure'; sub = 'Open loans and their guarantors'; cols = ['Guarantor','Loan','Borrower','Outstanding','Guarantor deposit','Due','Status']; numeric = [3, 4];
    var lpi2 = loanPaidIndex_(), gi2 = guarantorsOf_(), dep2 = depositTotals_();
    rows_('Loans').forEach(function(l){
      if(!num_(l.Amount)) return; var f = loanFig_(l, lpi2, today); if(f.status === 'Closed') return;
      (gi2[f.id] || []).forEach(function(g){ rows.push([g.name, f.id, f.name, f.outstanding, r2_(dep2[g.id] || 0), f.due, f.status]); });
    });
    rows.sort(function(a, b){ return a[0] < b[0] ? -1 : 1; });
    foot = ['Total', '', '', sumCol_(rows, 3), '', '', ''];
  } else if(type === 'settlements'){
    title = 'Settlements'; sub = 'Member exits and defaulted-loan deductions'; cols = ['Date','Kind','Member','Loan','Amount','Receipt / note']; numeric = [4];
    rows_('Exits').forEach(function(x){ if(inRange_(x.Date, from, to)) rows.push([String(x.Date || ''), 'Exit refund', nm(x.MemberID), '', num_(x.Refund), String(x.Note || '')]); });
    rows_('Ledger').forEach(function(l){ if(l.Type === 'settlement' && !isTrue_(l.Voided) && inRange_(l.Date, from, to)) rows.push([String(l.Date || ''), 'Loan settlement', nm(l.MemberID), l.LoanID || '', -num_(l.Amount), l.ReceiptNo || '']); });
    rows.sort(function(a, b){ return String(a[0]).localeCompare(String(b[0])); });
    foot = ['', 'Total', '', '', sumCol_(rows, 4), ''];
  } else if(type === 'finepot'){
    title = 'Fine pot statement'; sub = (from || to) ? (from || '…') + ' to ' + (to || '…') : 'All dates'; cols = ['Date','Type','Member','Amount','Balance','Note']; numeric = [3, 4];
    fineRows_().forEach(function(r){ if(inRange_(r.date, from, to)) rows.push([r.date, r.type, r.member, r.amount, r.balance, r.note]); });
    foot = ['', 'Net', '', sumCol_(rows, 3), rows.length ? rows[rows.length - 1][4] : 0, ''];
  } else if(type === 'expenses'){
    title = 'Expenses'; sub = (from || to) ? (from || '…') + ' to ' + (to || '…') : 'All dates'; cols = ['Date','Category','Amount','Note','Entered by']; numeric = [2];
    fineRows_().forEach(function(r){ if(r.type === 'expense' && inRange_(r.date, from, to)) rows.push([r.date, r.category, -r.amount, r.note.replace(/^[^·]+ · /, ''), r.by]); });
    var bycat = {}; rows.forEach(function(r){ bycat[r[1]] = (bycat[r[1]] || 0) + r[2]; });
    rows.sort(function(a, b){ return String(a[0]).localeCompare(String(b[0])); });
    foot = ['', 'Total', sumCol_(rows, 2), Object.keys(bycat).sort().map(function(k){ return k + ' ' + r2_(bycat[k]); }).join(' · '), ''];
  } else if(type === 'depositpot'){
    title = 'Deposit pot statement'; sub = (from || to) ? (from || '…') + ' to ' + (to || '…') : 'All dates'; cols = ['Date','Type','Member','Amount','Balance','Note']; numeric = [3, 4];
    var dl = rows_('Ledger').filter(function(l){ return l.Pot === 'deposit' && !isTrue_(l.Voided); }); dl.sort(function(a, b){ return String(a.Date).localeCompare(String(b.Date)) || (a._row - b._row); });
    var bal2 = 0; dl.forEach(function(l){ bal2 += num_(l.Amount); if(inRange_(l.Date, from, to)) rows.push([String(l.Date || ''), l.Type, nm(l.MemberID), num_(l.Amount), r2_(bal2), String(l.Note || '')]); });
    foot = ['', 'Net', '', sumCol_(rows, 3), r2_(bal2), ''];
  } else if(type === 'bankchecks'){
    title = 'Bank and cash checks'; cols = ['Date','Bank','Cash','Actual total','Expected','Difference','Entered by']; numeric = [1, 2, 3, 4, 5];
    rows_('Reconciliations').forEach(function(r){ if(inRange_(r.Date, from, to)) rows.push([String(r.Date || ''), num_(r.BankBalance), num_(r.CashInHand), r2_(num_(r.BankBalance) + num_(r.CashInHand)), num_(r.ExpectedTotal), num_(r.Difference), r.EnteredBy || '']); });
    rows.reverse();
  } else if(type === 'fysummary'){
    var fy = fyOf_(validMonth_(String(body.month || '')) ? String(body.month) : cur); title = 'Financial year summary ' + fy.label; sub = fy.start + ' to ' + fy.end; cols = ['Item','Amount']; numeric = [1];
    var tot = {}; rows_('Ledger').forEach(function(l){ if(isTrue_(l.Voided)) return; var m = String(l.Date).slice(0, 7); if(m < fy.start || m > fy.end) return; var k = l.Pot + '|' + l.Type; tot[k] = (tot[k] || 0) + num_(l.Amount); });
    var g = function(k){ return r2_(tot[k] || 0); };
    var lent = 0, back = 0;
    rows_('Loans').forEach(function(l){ if(num_(l.Amount) && inRange_(l.TakenDate, fy.start + '-01', fy.end + '-31')) lent += num_(l.Amount); });
    rows_('LoanPayments').forEach(function(p){ if(!isTrue_(p.Voided) && inRange_(p.Date, fy.start + '-01', fy.end + '-31')) back += num_(p.ToPrincipal); });
    rows = [['Contributions collected', g('deposit|contribution')], ['Opening deposits', g('deposit|opening')], ['Exit refunds paid', -g('deposit|exit_refund')], ['Deducted for defaulted loans', -g('deposit|settlement')], ['Joining fees', g('fine|joining_fee')], ['Penalties paid', g('fine|penalty_paid')], ['Loan extra', g('fine|loan_extra')], ['Late extra', g('fine|loan_late_extra')], ['Expenses', g('fine|expense')], ['Loans given (principal)', r2_(lent)], ['Principal returned', r2_(back)]];
    var ex2 = expectedHeld_(); rows.push(['Deposit pot now', ex2.deposit], ['Fine pot now', ex2.fine], ['Loan principal out now', ex2.principalOut]);
  } else if(type === 'audit'){
    title = 'Audit log'; sub = (from || to) ? (from || '…') + ' to ' + (to || '…') : 'Latest entries'; cols = ['When','Admin','Action','Target','Detail'];
    var who = String(body.admin || ''), q = String(body.q || '').toLowerCase();
    rows_('Audit').forEach(function(a){
      if(!inRange_(a.At, from, to)) return; if(who && String(a.Admin) !== who) return;
      var line = [String(a.At || '').replace('T', ' '), a.Admin || '', a.Action || '', a.Target || '', a.Detail || '']; if(q && line.join(' ').toLowerCase().indexOf(q) < 0) return;
      rows.push(line);
    });
    rows.reverse();
  }
  var cap = 3000, total = rows.length; if(rows.length > cap) rows = rows.slice(0, cap);
  return out_({ ok:true, type:type, title:title, sub:sub, cols:cols, rows:rows, foot:foot, numeric:numeric, total:total, truncated:total > cap, at:nowIso_() });
}




/* ---------- go live (owner only): checklist, then lock the Excel import for good ---------- */
function goLiveInfo_(){ var v = props_().getProperty('goLive'); try{ return v ? JSON.parse(v) : null; }catch(e){ return null; } }
function goLiveView_(S){ var bad = needDb_(); if(bad) return bad; return out_(goLiveState_()); }
function goLiveState_(){
  var gl = goLiveInfo_(), today = todayStr_();
  if(gl) return { ok:true, live:true, date:gl.date, at:gl.at, by:gl.by, canGo:false, checks:[] };
  var checks = [], nowMs = Date.now();
  function ago(iso){ var t = Date.parse(String(iso) + '+05:30'); return isFinite(t) ? (nowMs - t) / 3600000 : 1e9; }
  var xc = null; try{ xc = JSON.parse(props_().getProperty('lastExcelCheck') || 'null'); }catch(e){}
  checks.push({ key:'excel', ok:!!(xc && xc.clean && ago(xc.at) <= 36), at:xc ? xc.at : '', note:xc ? (xc.clean ? 'clean' : 'differences') : 'never' });
  var rec = rows_('Reconciliations').filter(function(r){ return String(r.Date) >= addDays_(today, -3); }).sort(function(a, b){ return String(a.Date) < String(b.Date) ? 1 : -1; })[0];
  checks.push({ key:'bank', ok:!!(rec && Math.abs(num_(rec.Difference)) < 1), at:rec ? String(rec.Date) : '', note:rec ? (Math.abs(num_(rec.Difference)) < 1 ? 'no difference' : 'difference') : 'none in last 3 days' });
  var bk = rows_('Audit').filter(function(a){ return a.Action === 'export'; }).sort(function(a, b){ return String(a.At) < String(b.At) ? 1 : -1; })[0];
  checks.push({ key:'backup', ok:!!(bk && ago(bk.At) <= 24), at:bk ? String(bk.At) : '', note:bk ? '' : 'never' });
  var ca = rows_('ClosedAccounts').length;
  checks.push({ key:'closed', ok:ca > 0, at:'', note:String(ca) });
  var cuPaid = {}; rows_('LoanPayments').forEach(function(p){ cuPaid[p.LoanID] = 1; });
  var cuNo = rows_('Loans').filter(function(l){ return blankLoan_(l, cuPaid); }).length;
  checks.push({ key:'blank', ok:cuNo === 0, at:'', note:String(cuNo) });
  return { ok:true, live:false, canGo:checks.every(function(c){ return c.ok; }), checks:checks, today:today };
}
function goLiveConfirm_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  if(String(body.confirm || '') !== 'GO LIVE') return out_({ ok:false, error:'bad-request' });
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    if(goLiveInfo_()) return out_({ ok:false, error:'already-live' });
    var v = goLiveState_(); if(!v.canGo) return out_({ ok:false, error:'not-ready', checks:v.checks });
    var info = { date:todayStr_(), at:nowIso_(), by:S.u };
    props_().setProperty('goLive', JSON.stringify(info)); markDirty_();
    audit_(S, 'go-live', info.date, 'Excel import and clean-up locked');
    return out_({ ok:true, date:info.date, at:info.at });
  } finally { lock.releaseLock(); }
}

/* ---------- check the app against the latest Excel (dry run and go-live) ---------- */
// body: members [{name, prior, cur, months:{'YYYY-MM':amount}}], loans [{name, taken, amount, repay, fine, status}], totals {deposit, fine, expense}
function excelCheck_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var xm = Array.isArray(body.members) ? body.members.slice(0, 500) : [], xl = Array.isArray(body.loans) ? body.loans.slice(0, 4000) : [], tot = body.totals || {};
  var today = todayStr_(), members = rows_('Members').filter(function(m){ return !isClosedSrc_(m); });
  var ledger = rows_('Ledger').filter(function(r){ return !isTrue_(r.Voided); }), dep = depositTotals_(), ci = contribIndex_();
  function sumL(pot){ return ledger.filter(function(r){ return r.Pot === pot; }).reduce(function(a, r){ return a + num_(r.Amount); }, 0); }
  var depApp = sumL('deposit'), fineApp = sumL('fine'), expApp = -ledger.filter(function(r){ return r.Pot === 'fine' && r.Type === 'expense'; }).reduce(function(a, r){ return a + num_(r.Amount); }, 0);
  var lines = [];
  function line(label, ex, ap){ if(ex === null || ex === undefined || ex === '') return; ex = num_(ex); lines.push({ label:label, excel:ex, app:r2_(ap), ok:Math.abs(ex - ap) < 0.5 }); }
  var live = members.filter(function(m){ return String(m.Status) !== 'Exited'; });
  line('Members (live)', xm.length, live.length);
  line('Fine pot (remaining)', tot.fine, fineApp);
  line('Total expense', tot.expense, expApp);
  // members
  var byNorm = {}, byTok = {}; members.forEach(function(m){ byNorm[normName_(m.Name)] = m; byTok[sortedTokens_(m.Name)] = m; });
  var seenId = {}, diffs = [], missApp = [], maxD = 0, diffN = 0, appOnlyN = 0, appOnlyAmt = 0;
  xm.forEach(function(x){
    var raw = String(x.name || '').trim(); if(!raw) return;
    var m = byNorm[normName_(raw)] || byTok[sortedTokens_(raw)];
    if(!m){ missApp.push(raw); return; }
    seenId[m.MemberID] = 1;
    var ex = num_(x.cur), ap = dep[m.MemberID] || 0, md = [], months = x.months || {}, paid = ci[m.MemberID] || {};
    Object.keys(paid).forEach(function(k){ var c = paid[k]; if(String(c.by) !== 'import' && !(num_(months[k]) > 0)){ ap -= c.amount; appOnlyN++; appOnlyAmt += c.amount; } });   // paid in the app, not yet in the Excel: expected, not a difference
    Object.keys(months).sort().forEach(function(k){
      if(!validMonth_(k) || k > today.slice(0, 7)) return;
      var e = num_(months[k]) > 0, a = !!paid[k], st = startMonth_(m);
      if(a && !e && String(paid[k].by) !== 'import') return;
      if(e !== a && k >= st) md.push({ m:k, excel:e, app:a });
    });
    if(Math.abs(ex - ap) >= 0.5 || md.length){ diffN++; if(diffs.length < 60) diffs.push({ id:m.MemberID, name:m.Name, excel:ex, app:r2_(ap), months:md.slice(0, 12), status:m.Status }); }
  });
  var missExcel = live.filter(function(m){ return !seenId[m.MemberID]; }).map(function(m){ return m.Name; });
  if(tot.deposit !== null && tot.deposit !== undefined && tot.deposit !== '') lines.splice(1, 0, { label:'Deposit pot (app entries not yet in Excel left out)', excel:num_(tot.deposit), app:r2_(depApp - appOnlyAmt), ok:Math.abs(num_(tot.deposit) - (depApp - appOnlyAmt)) < 0.5 });
  // loans
  var pi = loanPaidIndex_(), key = function(n, d, a){ return normName_(n) + '|' + String(d || '').slice(0, 10) + '|' + Math.round(num_(a)); };
  var appL = {}, exL = {};
  rows_('Loans').forEach(function(l){ if(blankName_(l.ApplicantName) && !num_(l.Amount)) return; var k = key(l.ApplicantName, l.TakenDate, l.Amount); (appL[k] || (appL[k] = [])).push(l); });
  var onlyEx = [], stDiff = [];
  xl.forEach(function(x){
    if(!String(x.name || '').trim()) return;
    var k = key(x.name, x.taken, x.amount), list = appL[k];
    if(!list || !list.length){ onlyEx.push({ name:String(x.name).trim(), taken:String(x.taken || ''), amount:num_(x.amount) }); return; }
    var l = list.shift(), f = loanFig_(l, pi, today), exClosed = /complete|closed|paid/i.test(String(x.status || '')), apClosed = f.status === 'Closed';
    if(exClosed !== apClosed) stDiff.push({ id:f.id, name:f.name, taken:String(l.TakenDate || ''), amount:num_(l.Amount), excel:String(x.status || ''), app:f.status });
  });
  var onlyApp = []; Object.keys(appL).forEach(function(k){ appL[k].forEach(function(l){ onlyApp.push({ id:l.LoanID, name:l.ApplicantName, taken:String(l.TakenDate || ''), amount:num_(l.Amount) }); }); });
  var n = xl.filter(function(x){ return String(x.name || '').trim(); }).length;
  line('Loans (count)', n, rows_('Loans').filter(function(l){ return !(blankName_(l.ApplicantName) && !num_(l.Amount)); }).length);
  var ok = lines.every(function(l){ return l.ok; }) && !diffN && !missApp.length && !missExcel.length && !onlyEx.length && !onlyApp.length && !stDiff.length;
  try{ props_().setProperty('lastExcelCheck', JSON.stringify({ at:nowIso_(), clean:ok })); }catch(e){}
  audit_(S, 'excel-check', '', ok ? 'clean' : (diffN + ' member, ' + onlyEx.length + ' + ' + onlyApp.length + ' loan, ' + stDiff.length + ' status differences'));
  return out_({ ok:true, clean:ok, at:nowIso_(), lines:lines, appOnly:{ count:appOnlyN, amount:r2_(appOnlyAmt) }, memberDiffCount:diffN, memberDiffs:diffs, missingInApp:missApp.slice(0, 40), missingInExcel:missExcel.slice(0, 40),
    loanOnlyExcel:onlyEx.slice(0, 40), loanOnlyApp:onlyApp.slice(0, 40), loanStatusDiffs:stDiff.slice(0, 60), counts:{ onlyEx:onlyEx.length, onlyApp:onlyApp.length, stDiff:stDiff.length, missApp:missApp.length, missExcel:missExcel.length } });
}

/* ---------- contributions from a pasted WhatsApp list ---------- */
// entries: [{ name, amount }] already read from the list by the app (paid lines only). Returns how each name matches.
function waPlan_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var month = String(body.month || curMonth_()); if(!validMonth_(month)) return out_({ ok:false, error:'bad-months' });
  var entries = Array.isArray(body.entries) ? body.entries.slice(0, 400) : [], amt = contribAmount_(), ci = contribIndex_();
  var pool = rows_('Members').filter(function(m){ return String(m.Status) !== 'Exited'; }).map(function(m){ return { id:m.MemberID, name:m.Name, norm:normName_(m.Name), tok:sortedTokens_(m.Name), m:m }; });
  var used = {}, items = [];
  entries.forEach(function(e){
    var raw = String(e && e.name || '').trim().slice(0, 80), n = normName_(raw), tok = sortedTokens_(raw), a = Number(e && e.amount) || 0;
    var it = { raw:raw, amount:a, status:'none', id:'', name:'', candidates:[] };
    if(n.length < 3){ it.status = 'skip'; items.push(it); return; }
    var hit = pool.filter(function(p){ return p.norm === n; });
    if(!hit.length) hit = pool.filter(function(p){ return p.tok === tok; });
    if(hit.length === 1){ it.id = hit[0].id; it.name = hit[0].name; it.status = 'matched'; }
    else if(hit.length > 1){ it.status = 'maybe'; it.candidates = hit.slice(0, 4).map(function(p){ return { id:p.id, name:p.name }; }); }
    else {
      var cands = [];
      pool.forEach(function(p){
        var d = lev_(n, p.norm), why = '';
        if(d <= 2) why = 'spelling'; else if(n.length > 8 && (p.norm.indexOf(n) === 0 || n.indexOf(p.norm) === 0)) why = 'starts';
        if(why) cands.push({ id:p.id, name:p.name, dist:d });
      });
      cands.sort(function(x, y){ return x.dist - y.dist; });
      if(cands.length){ it.status = 'maybe'; it.candidates = cands.slice(0, 4).map(function(c){ return { id:c.id, name:c.name }; }); }
    }
    items.push(it);
  });
  items.forEach(function(it){
    var ids = it.id ? [it.id] : it.candidates.map(function(c){ return c.id; });
    it.candidates.forEach(function(c){ c.paid = !!((ci[c.id] || {})[month]); });
    if(it.status === 'matched'){
      var mem = memberById_(it.id), p = (ci[it.id] || {})[month];
      if(p) it.status = 'already';
      else if(startMonth_(mem) > month) it.status = 'early';
      else if(used[it.id]) it.status = 'duplicate';
      else { used[it.id] = 1; if(it.amount && it.amount !== amt) it.status = 'amount'; }
    }
  });
  var cnt = {}; items.forEach(function(it){ cnt[it.status] = (cnt[it.status] || 0) + 1; });
  return out_({ ok:true, month:month, amount:amt, items:items, counts:cnt, closed:monthClosed_(month) });
}
// ids: members to mark paid for one month, with one mode and date. One receipt per member.
function waApply_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var month = String(body.month || ''); if(!validMonth_(month)) return out_({ ok:false, error:'bad-months' });
  var mode = ['Bank','UPI'].indexOf(body.mode) >= 0 ? body.mode : ''; if(!mode) return out_({ ok:false, error:'bad-mode' });
  var today = todayStr_(), date = body.date ? String(body.date) : today;
  if(!validDate_(date) || date > today) return out_({ ok:false, error:'bad-date' });
  if(monthClosed_(date.slice(0, 7))) return out_({ ok:false, error:'month-closed', month:date.slice(0, 7) });
  var ids = Array.isArray(body.ids) ? body.ids.map(String).slice(0, 250) : []; if(!ids.length) return out_({ ok:false, error:'bad-request' });
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var ci = contribIndex_(), done = [], skipped = [], seen = {}, cur = curMonth_();
    if(month > addMonths_(cur, 12)) return out_({ ok:false, error:'too-far', month:month });
    ids.forEach(function(id){
      if(seen[id]) return; seen[id] = 1;
      var mem = memberById_(id);
      if(!mem){ skipped.push({ id:id, error:'no-member' }); return; }
      if(String(mem.Status) === 'Exited'){ skipped.push({ id:id, name:mem.Name, error:'member-exited' }); return; }
      if(startMonth_(mem) > month){ skipped.push({ id:id, name:mem.Name, error:'before-join' }); return; }
      if((ci[id] || {})[month]){ skipped.push({ id:id, name:mem.Name, error:'already-paid' }); return; }
      var rec = writeContribs_(S, mem, [month], date, mode, 'WhatsApp list', 'whatsapp');
      done.push({ id:id, name:mem.Name, receipt:rec.no });
    });
    audit_(S, 'contribution-whatsapp-list', month, done.length + ' recorded, ' + skipped.length + ' skipped · ' + mode);
    return out_({ ok:true, done:done, skipped:skipped });
  } finally { lock.releaseLock(); }
}

/* ---------- clean-up import: closed accounts, blank loan rows, name matches (admin1) ---------- */
function isClosedSrc_(m){ return String((m && m.Source) || '').indexOf('import-closed') === 0; }   // 'import-closed' or 'import-closed:<Sr>'
function blankName_(s){ var n = normName_(s); return !n || n === 'UNNAMED APPLICANT'; }          // the Excel reader names empty rows "Unnamed applicant"
function blankLoan_(l, paid){ return blankName_(l.ApplicantName) && !num_(l.Amount) && !num_(l.ReturnAmount) && !(paid && paid[l.LoanID]); }
function normName_(s){ return String(s || '').toUpperCase().replace(/\([^)]*\)/g, ' ').replace(/[^A-Zऀ-ॿ ]+/g, ' ').replace(/\s+/g, ' ').trim(); }
function sortedTokens_(s){ return normName_(s).split(' ').filter(Boolean).sort().join(' '); }
// What would change. closedNames = names from the Excel "Closed Account Details" sheet.
function closedClean_(raw){ return String(raw || '').replace(/\([^)]*\)/g, ' ').replace(/[\t ]+/g, ' ').replace(/\s+/g, ' ').trim().toUpperCase(); }
function closedRows_(list){
  var out = [], seen = {};
  (Array.isArray(list) ? list.slice(0, 300) : []).forEach(function(x){
    var sr = Math.round(num_(x && x.sr)), name = closedClean_(x && x.name);
    if(!(sr > 0) || name.length < 3 || seen[sr]) return; seen[sr] = 1;
    out.push({ sr:sr, name:name, raw:String(x.name).trim().slice(0, 80), deposit:num_(x.deposit), loan:num_(x.loan), lra:num_(x.lra), late:num_(x.late), net:num_(x.net), remarks:String(x.remarks || '').trim().slice(0, 150) });
  });
  return out;
}
// Each row of the Excel "Closed Account Details" sheet is its own closed account (a person can join, leave and join again).
// A row already imported (same Sr. No.) is skipped. Loans are linked: live member first; a name with exactly one closed account links to
// it; a name with several closed accounts (and no live member) is left for you to choose.
function cleanupPlan_(closed){
  var rowsIn = closedRows_(closed), members = rows_('Members'), have = {};
  rows_('ClosedAccounts').forEach(function(c){ have[String(c.SrNo)] = c.MemberID; });
  // closed accounts added by an earlier build (Source 'import-closed' or 'import-closed:<Sr>'): adopt them instead of adding twins
  var usedId = {}; Object.keys(have).forEach(function(k){ usedId[have[k]] = 1; });
  var oldClosed = members.filter(function(m){ return isClosedSrc_(m) && !usedId[m.MemberID]; });
  var adopt = [];
  rowsIn.forEach(function(r){
    if(have[String(r.sr)]) return;
    var hit = null;
    oldClosed.forEach(function(m){ if(hit) return; var mm = /^import-closed:(\d+)/.exec(String(m.Source)); if(mm && mm[1] === String(r.sr) && !usedId[m.MemberID]) hit = m; });
    if(!hit) oldClosed.forEach(function(m){ if(!hit && !usedId[m.MemberID] && normName_(m.Name) === normName_(r.name)) hit = m; });
    if(hit){ usedId[hit.MemberID] = 1; have[String(r.sr)] = hit.MemberID; adopt.push({ sr:r.sr, id:hit.MemberID }); }
  });
  var live = {}, liveAny = [];
  members.forEach(function(m){ if(!isClosedSrc_(m)){ live[normName_(m.Name)] = m; liveAny.push(m); } });
  var fresh = [], already = [];
  rowsIn.forEach(function(r){ if(have[String(r.sr)]) already.push({ name:r.name, sr:r.sr, id:have[String(r.sr)] }); else fresh.push(r); });
  var closedBy = {};                                         // normalised name -> [{ sr, id }]
  rowsIn.forEach(function(r){ var k = normName_(r.name); (closedBy[k] || (closedBy[k] = [])).push({ sr:r.sr, name:r.name, id:'closed:' + r.sr, loan:r.loan }); });
  var paid = {}; rows_('LoanPayments').forEach(function(p){ if(!isTrue_(p.Voided)) paid[p.LoanID] = 1; });
  var loans = rows_('Loans'), blanks = 0, linkClosed = 0, unmatched = {}, byKey = {}, byId = {};
  loans.forEach(function(l){
    byId[l.LoanID] = l;
    if(blankLoan_(l, paid)){ blanks++; return; }
    var nm = normName_(l.ApplicantName);
    var k = nm + '|' + String(l.TakenDate || '').slice(0, 10) + '|' + Math.round(num_(l.Amount)); byKey[k] = (byKey[k] || 0) + 1;
    if(blankName_(l.ApplicantName)) return;
    var mid = String(l.MemberID || '');
    if(mid && memberById_(mid)) return;
    if(live[nm]) return;
    var cb = closedBy[nm];
    if(cb && cb.length === 1){ linkClosed++; return; }
    var u = unmatched[nm] || (unmatched[nm] = { name:String(l.ApplicantName).trim(), norm:nm, count:0, candidates:[] }); u.count++;
  });
  // loans in the Excel that the Sheet does not have yet, and loans whose name was corrected in the Excel
  var ds = null; try{ ds = readDataset_(); }catch(e){}
  var newLoans = [], renamed = [], exRows = (ds && ds.rows) || [];
  exRows.forEach(function(r){
    if(blankName_(r.applicant) && !num_(r.takenAmt)) return;
    var lid = 'L-' + pad_(r.srno, 4), cur = byId[lid], taken = String(r.takenDate || '').slice(0, 10);
    var k = normName_(r.applicant) + '|' + taken + '|' + Math.round(num_(r.takenAmt));
    if(byKey[k] > 0){ byKey[k]--; return; }
    if(cur && !blankLoan_(cur, paid) && String(cur.TakenDate || '').slice(0, 10) === taken && Math.round(num_(cur.Amount)) === Math.round(num_(r.takenAmt)) && normName_(cur.ApplicantName) !== normName_(r.applicant)){
      renamed.push({ id:lid, from:String(cur.ApplicantName), to:String(r.applicant).trim(), row:r.srno }); return;
    }
    newLoans.push({ srno:r.srno, name:String(r.applicant).trim(), taken:taken, amount:num_(r.takenAmt), repay:num_(r.repayAmt), fine:num_(r.fine) });
  });
  var pool = liveAny.map(function(m){ return { id:m.MemberID, name:m.Name, norm:normName_(m.Name), status:m.Status }; });
  rowsIn.forEach(function(r){ pool.push({ id:'closed:' + r.sr, name:r.name + ' (closed account Sr ' + r.sr + ')', norm:normName_(r.name), status:'Exited' }); });
  var list = Object.keys(unmatched).map(function(k){
    var u = unmatched[k], tok = sortedTokens_(u.name);
    pool.forEach(function(p){
      var d = lev_(u.norm, p.norm), why = '';
      if(p.norm === u.norm) why = 'same name, several closed accounts';
      else if(sortedTokens_(p.name.replace(/ \(closed.*$/, '')) === tok) why = 'same words';
      else if(d <= 2) why = 'spelling';
      else if(u.norm.indexOf(p.norm) >= 0 && p.norm.length > 10) why = 'contains';
      if(why) u.candidates.push({ id:p.id, name:p.name, status:p.status, why:why, dist:d });
    });
    u.candidates.sort(function(a, b){ return a.dist - b.dist; }); u.candidates = u.candidates.slice(0, 5);
    return u;
  });
  list.sort(function(a, b){ return b.count - a.count || (a.name < b.name ? -1 : 1); });
  return { newMembers:fresh, alreadyMembers:already, adopt:adopt, blankLoanRows:blanks, loansLinkedToClosed:linkClosed, unmatched:list, newLoans:newLoans, renamedLoans:renamed };
}
function cleanupPreview_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var p = cleanupPlan_(body.closed); p.ok = true; p.adoptCount = p.adopt.length; delete p.adopt; return out_(p);
}
function cleanupApply_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  if(goLiveInfo_()) return out_({ ok:false, error:'golive-locked' });
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var plan = cleanupPlan_(body.closed), at = nowIso_(), ids = {};        // 'closed:<sr>' -> MemberID
    rows_('ClosedAccounts').forEach(function(c){ ids['closed:' + c.SrNo] = c.MemberID; });
    var closedAll = closedRows_(body.closed || []), addC = [], addM = [];
    plan.adopt.forEach(function(x){
      ids['closed:' + x.sr] = x.id;
      closedAll.forEach(function(r){ if(String(r.sr) === String(x.sr)) addC.push({ SrNo:r.sr, MemberID:x.id, Name:r.name, Deposit:r.deposit, LoanAmount:r.loan, LoanReturn:r.lra, LateFees:r.late, Net:r.net, Remarks:r.remarks, ImportedAt:at }); });
    });
    var maxN = 0; rows_('Members').forEach(function(m){ var n = Number(String(m.MemberID).replace(/\D/g, '')); if(n > maxN) maxN = n; });
    plan.newMembers.forEach(function(f){
      maxN++; var id = 'M-' + pad_(maxN, 4); ids['closed:' + f.sr] = id;
      addM.push({ MemberID:id, Name:f.name, Mobile:'', DOB:'', JoinDate:'', Status:'Exited', OpeningDeposit:0, JoiningFee:0, FailedTries:0, Source:'import-closed', CreatedAt:at, CreatedBy:S.u });
      addC.push({ SrNo:f.sr, MemberID:id, Name:f.name, Deposit:f.deposit, LoanAmount:f.loan, LoanReturn:f.lra, LateFees:f.late, Net:f.net, Remarks:f.remarks, ImportedAt:at });
    });
    if(addM.length) appendMany_('Members', addM);
    if(addC.length) appendMany_('ClosedAccounts', addC);
    var live = {}, closedBy = {};
    rows_('Members').forEach(function(m){ if(!isClosedSrc_(m)) live[normName_(m.Name)] = m; });
    closedAll.forEach(function(r){ var k = normName_(r.name); (closedBy[k] || (closedBy[k] = [])).push(ids['closed:' + r.sr]); });
    var approved = {}; (Array.isArray(body.links) ? body.links : []).forEach(function(x){ if(x && x.name && x.to) approved[normName_(x.name)] = String(x.to); });
    var ren = {}; plan.renamedLoans.forEach(function(x){ ren[x.id] = x.to; });
    var paid = {}; rows_('LoanPayments').forEach(function(p){ if(!isTrue_(p.Voided)) paid[p.LoanID] = 1; });
    var loans = rows_('Loans'), linked = 0, linkedApproved = 0, removed = 0, renamedN = 0;
    loans.forEach(function(l){
      if(blankLoan_(l, paid)){ if(body.removeBlank !== false) removed++; return; }
      var upd = {}, nameNow = l.ApplicantName;
      if(ren[l.LoanID]){ upd.ApplicantName = ren[l.LoanID]; nameNow = ren[l.LoanID]; renamedN++; l.MemberID = ''; upd.MemberID = ''; }
      var nm = normName_(nameNow), mid = String(l.MemberID || '');
      if(!blankName_(nameNow) && !(mid && memberById_(mid))){
        var to = approved[nm], m = live[nm], cb = closedBy[nm];
        if(to){ to = ids[to] || to; if(memberById_(to)){ upd.MemberID = to; linkedApproved++; } }
        else if(m){ upd.MemberID = m.MemberID; linked++; }
        else if(cb && cb.length === 1 && cb[0]){ upd.MemberID = cb[0]; linked++; }
      }
      if(Object.keys(upd).length) update_('Loans', l._row, upd);
    });
    if(removed){
      var fresh = rows_('Loans').filter(function(l){ return !blankLoan_(l, paid); });
      clearRows_('Loans'); appendMany_('Loans', fresh.map(function(l){ var o = {}; TABS.Loans.forEach(function(k){ o[k] = l[k]; }); return o; }));
    }
    // loans that are in the Excel but not in the Sheet yet (added after the blank rows are gone, so their numbers are free)
    var addedL = 0;
    if(plan.newLoans.length){
      var haveIds = {}; rows_('Loans').forEach(function(l){ haveIds[l.LoanID] = 1; });
      var nextN = 0; Object.keys(haveIds).forEach(function(k){ var n = Number(String(k).replace(/\D/g, '')); if(n > nextN) nextN = n; });
      var byName = {}; rows_('Members').forEach(function(m){ var k = nameKey_(m.Name); if(!byName[k] || isClosedSrc_(byName[k])) byName[k] = m; });
      var dsx = readDataset_(), exByNo = {}; ((dsx && dsx.rows) || []).forEach(function(r){ exByNo[r.srno] = r; });
      var nl = [], ng = [];
      plan.newLoans.forEach(function(x){
        var r = exByNo[x.srno]; if(!r) return;
        var lid = 'L-' + pad_(x.srno, 4); if(haveIds[lid]){ nextN++; lid = 'L-' + pad_(nextN, 4); } haveIds[lid] = 1;
        var m = byName[nameKey_(r.applicant)];
        nl.push({ LoanID:lid, MemberID:m ? m.MemberID : '', ApplicantName:r.applicant, Amount:num_(r.takenAmt), TakenDate:r.takenDate || '', DueDate:r.repayDate || '', ReturnAmount:num_(r.repayAmt), Fine:num_(r.fine), LateExtra:0, Status:r.status || '', Suspended:r.suspended || '', Source:'import', CreatedAt:at });
        String(r.guarantor || '').split(/\s*(?:,|&|\+|\n|;|\band\b)\s*/i).forEach(function(g){
          g = String(g).trim(); if(!g || /^(none|nil|n\/a|-|not specified)$/i.test(g)) return;
          var gm = byName[nameKey_(g)]; ng.push({ LoanID:lid, MemberID:gm ? gm.MemberID : '', Name:g, Source:'import' });
        });
      });
      if(nl.length){ appendMany_('Loans', nl); if(ng.length) appendMany_('Guarantors', ng); addedL = nl.length; }
    }
    audit_(S, 'cleanup-import', 'closed accounts', addM.length + ' closed accounts added, ' + plan.adopt.length + ' adopted, ' + linked + ' loans linked, ' + linkedApproved + ' approved links, ' + removed + ' blank loan rows removed, ' + addedL + ' loans added, ' + renamedN + ' names corrected');
    return out_({ ok:true, added:addM.length, adopted:plan.adopt.length, linked:linked, linkedApproved:linkedApproved, removedBlank:removed, addedLoans:addedL, renamed:renamedN });
  } finally { lock.releaseLock(); }
}

/* ---------- settings screen ---------- */
var SETTING_EDIT_ = { ContributionAmount:'int', JoiningFee:'int0', PenaltyAmount:'int0', SocietyUPI:'upi', ExpenseCategories:'list' };
var SETTING_SHOW_ = ['ContributionAmount','JoiningFee','PenaltyAmount','SocietyUPI','ExpenseCategories','FinancialYearStartMonth','TrackingStartMonth','SuspensionMonths','OverdueDay'];
function settingsGet_(S){
  var bad = needDb_(); if(bad) return bad;
  var defaults = { ContributionAmount:500, JoiningFee:0, PenaltyAmount:500, SocietyUPI:'', ExpenseCategories:EXPENSE_CATS_.join(', '), FinancialYearStartMonth:4, TrackingStartMonth:'2026-01', SuspensionMonths:6, OverdueDay:15 };
  var list = SETTING_SHOW_.map(function(k){ return { key:k, value:setting_(k, defaults[k]), editable:!!SETTING_EDIT_[k] }; });
  return out_({ ok:true, settings:list, canEdit:S.r === 'admin1' });
}
function settingsSet_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var key = String(body.key || ''), kind = SETTING_EDIT_[key]; if(!kind) return out_({ ok:false, error:'bad-key' });
  var v = body.value, val;
  if(kind === 'int' || kind === 'int0'){ val = Math.round(num_(v)); if(!(kind === 'int' ? val >= 100 : val >= 0) || val > 1000000) return out_({ ok:false, error:'bad-value' }); }
  else if(kind === 'upi'){ val = String(v || '').trim(); if(val && !/^[A-Za-z0-9._\-]{2,}@[A-Za-z0-9]{2,}$/.test(val)) return out_({ ok:false, error:'bad-value' }); }
  else { val = String(v || '').split(',').map(function(x){ return x.trim().replace(/[^\w &.\-]/g, ''); }).filter(Boolean); if(!val.length || val.length > 20) return out_({ ok:false, error:'bad-value' }); val = val.join(', '); }
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var cur = rows_('Settings').filter(function(r){ return String(r.Key) === key; })[0], at = nowIso_(), old = cur ? cur.Value : '';
    if(cur) update_('Settings', cur._row, { Value:val, UpdatedAt:at, UpdatedBy:S.u });
    else appendMany_('Settings', [{ Key:key, Value:val, Note:'', UpdatedAt:at, UpdatedBy:S.u }]);
    audit_(S, 'setting-changed', key, old + ' -> ' + val);
    return out_({ ok:true, key:key, value:val });
  } finally { lock.releaseLock(); }
}

function membersList_(S){
  var bad = needDb_(); if(bad) return bad;
  loanSweep_(S, false);
  var cur = curMonth_(), ci = contribIndex_(), di = depositTotals_(), pi = penaltyIndex_(), si = suspensionIndex_();
  return out_({ ok:true, month:cur, amount:contribAmount_(), today:todayStr_(), trackStart:trackStart_(), joiningFee:num_(setting_('JoiningFee', 0)), penaltyAmount:penaltyAmount_(), members: rows_('Members').map(function(m){
    return memberPublic_(m, standing_(m, ci[m.MemberID], cur), di[m.MemberID], pi[m.MemberID], si[m.MemberID]);
  }) });
}

function receiptObj_(no, mem, months, amount, date, mode, ref, by){
  return { no:no, memberId:mem.MemberID, name:mem.Name, months:months, amount:amount, date:date, mode:mode, ref:ref, by:by };
}
// Writes the contribution and ledger rows. The caller has already checked everything and holds the lock.
function writeContribs_(S, mem, months, date, mode, ref, source){
  var amt = contribAmount_(), no = nextReceipt_(), at = nowIso_(), cs = [], ls = [];
  months.forEach(function(mo){
    var cid = newId_('C');
    cs.push({ ContribID:cid, MemberID:mem.MemberID, Month:mo, Amount:amt, PaidDate:date, Mode:mode, Ref:ref, Source:source || 'app', EnteredBy:S.u, EnteredAt:at, ReceiptNo:no });
    ls.push({ LedgerID:newId_('L'), Date:date, Pot:'deposit', Type:'contribution', MemberID:mem.MemberID, Amount:amt, Ref:cid, Note:'Receipt ' + no + ' ' + mo, EnteredBy:S.u, EnteredAt:at });
  });
  appendMany_('Contributions', cs); appendMany_('Ledger', ls);
  audit_(S, 'contribution', mem.MemberID, no + ' · ' + months.join(', ') + ' · ' + (amt * months.length) + ' · ' + mode);
  return receiptObj_(no, mem, months.slice(), amt * months.length, date, mode, ref, S.u);
}

function contribAdd_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var mem = memberById_(body.memberId);
    if(!mem) return out_({ ok:false, error:'no-member' });
    if(String(mem.Status) === 'Exited') return out_({ ok:false, error:'member-exited' });
    var months = Array.isArray(body.months) ? body.months.map(String) : [];
    var seen = {}; months = months.filter(function(m){ if(seen[m]) return false; seen[m] = 1; return true; }).sort();
    if(!months.length || months.length > 12 || !months.every(validMonth_)) return out_({ ok:false, error:'bad-months' });
    var cur = curMonth_(), today = todayStr_(), date = body.date ? String(body.date) : today;
    if(!validDate_(date) || date > today) return out_({ ok:false, error:'bad-date' });
    var mode = ['Bank','UPI'].indexOf(body.mode) >= 0 ? body.mode : '';
    if(!mode) return out_({ ok:false, error:'bad-mode' });
    var jm = String(mem.JoinDate || '').slice(0, 7);
    for(var i = 0; i < months.length; i++){
      if(validMonth_(jm) && months[i] < jm) return out_({ ok:false, error:'before-join', month:months[i] });
      if(months[i] > addMonths_(cur, 12)) return out_({ ok:false, error:'too-far', month:months[i] });
    }
    if(monthClosed_(date.slice(0, 7))) return out_({ ok:false, error:'month-closed', month:date.slice(0, 7) });
    var paid = contribIndex_()[mem.MemberID] || {}, dupes = months.filter(function(m){ return !!paid[m]; });
    if(dupes.length) return out_({ ok:false, error:'already-paid', months:dupes });
    var rec = writeContribs_(S, mem, months, date, mode, cleanRef_(body.ref), 'app');
    return out_({ ok:true, receipt:rec });
  } finally { lock.releaseLock(); }
}

// Void a receipt (all its months) or one contribution. withusadmin1 any time; the admin who made it, only within UNDO_SECONDS.
function contribVoid_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var key = String(body.receipt || ''), cid = String(body.contribId || '');
    if(!key && !cid) return out_({ ok:false, error:'bad-request' });
    var hit = rows_('Contributions').filter(function(c){ return !isTrue_(c.Voided) && (key ? String(c.ReceiptNo) === key : String(c.ContribID) === cid); });
    if(!hit.length) return out_({ ok:false, error:'not-found' });
    var nowMs = Date.now();
    for(var i = 0; i < hit.length; i++){
      var c = hit[i];
      if(monthClosed_(String(c.PaidDate).slice(0, 7))) return out_({ ok:false, error:'month-closed', month:String(c.PaidDate).slice(0, 7) });
      if(S.r !== 'admin1'){
        var t = Date.parse(String(c.EnteredAt) + '+05:30');
        if(String(c.EnteredBy) !== S.u || !isFinite(t) || nowMs - t > UNDO_SECONDS * 1000) return out_({ ok:false, error:'forbidden' });
      }
    }
    var ids = {}, at = nowIso_();
    hit.forEach(function(c){ ids[c.ContribID] = 1; update_('Contributions', c._row, { Voided:true, VoidedBy:S.u, VoidedAt:at }); });
    rows_('Ledger').forEach(function(l){ if(ids[l.Ref] && !isTrue_(l.Voided)) update_('Ledger', l._row, { Voided:true, VoidedBy:S.u, VoidedAt:at }); });
    audit_(S, 'contribution-voided', hit[0].MemberID, (key || cid) + ' · ' + hit.length + ' month(s) · ' + String(body.reason || '').slice(0, 100));
    return out_({ ok:true, voided:hit.length });
  } finally { lock.releaseLock(); }
}

function checkMemberFields_(body, selfId){
  var f = {}, all = rows_('Members');
  if(body.name !== undefined){
    f.Name = cleanName_(body.name);
    if(f.Name.length < 3 || f.Name.length > 80) return { error:'bad-name' };
  }
  if(body.mobile !== undefined){
    f.Mobile = cleanMobile_(body.mobile);
    if(!f.Mobile) return { error:'bad-mobile' };
    for(var i = 0; i < all.length; i++) if(String(all[i].Mobile) === f.Mobile && String(all[i].MemberID) !== String(selfId || '')) return { error:'mobile-exists', member:{ id:all[i].MemberID, name:all[i].Name } };
  }
  if(body.dob !== undefined){
    var d = String(body.dob || '');
    if(d && (!validDate_(d) || d > todayStr_() || d < '1900-01-01')) return { error:'bad-dob' };
    f.DOB = d;
  }
  if(body.joinDate !== undefined){
    var j = String(body.joinDate || '');
    if(!validDate_(j) || j > todayStr_() || j < '1990-01-01') return { error:'bad-join' };
    f.JoinDate = j;
  }
  return { fields:f };
}

function memberAdd_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var chk = checkMemberFields_({ name:body.name, mobile:body.mobile, dob:body.dob || '', joinDate:body.joinDate || todayStr_() }, '');
    if(chk.error) return out_({ ok:false, error:chk.error, member:chk.member });
    var f = chk.fields, fee = body.joiningFee === undefined || body.joiningFee === '' ? num_(setting_('JoiningFee', 0)) : num_(body.joiningFee);
    if(fee < 0 || fee > 100000) return out_({ ok:false, error:'bad-fee' });
    var first = body.first === 'now', mode = body.mode;
    if(first && ['Bank','UPI'].indexOf(mode) < 0) return out_({ ok:false, error:'bad-mode' });
    if(!body.confirmDup){
      var key = nameKey_(f.Name), sim = [];
      rows_('Members').forEach(function(m){ var k = nameKey_(m.Name); if(k === key || (Math.abs(k.length - key.length) <= 2 && lev_(k, key) <= 2)) sim.push({ id:m.MemberID, name:m.Name }); });
      if(sim.length) return out_({ ok:false, error:'similar-name', similar:sim.slice(0, 3) });
    }
    var jm = f.JoinDate.slice(0, 7);
    if(monthClosed_(todayStr_().slice(0, 7)) && (fee > 0 || first)) return out_({ ok:false, error:'month-closed', month:todayStr_().slice(0, 7) });
    var id = nextMemberId_(), at = nowIso_();
    append_('Members', { MemberID:id, Name:f.Name, Mobile:f.Mobile, DOB:f.DOB, JoinDate:f.JoinDate, Status:'Active', OpeningDeposit:0, JoiningFee:fee, FailedTries:0, Source:'app', CreatedAt:at, CreatedBy:S.u });
    if(fee > 0) append_('Ledger', { LedgerID:newId_('L'), Date:todayStr_(), Pot:'fine', Type:'joining_fee', MemberID:id, Amount:fee, Ref:id, Note:'Joining fee', EnteredBy:S.u, EnteredAt:at });
    audit_(S, 'member-added', id, f.Name);
    var rec = null;
    if(first) rec = writeContribs_(S, memberById_(id), [jm], todayStr_(), mode, cleanRef_(body.ref), 'app');
    return out_({ ok:true, member:{ id:id, name:f.Name }, receipt:rec });
  } finally { lock.releaseLock(); }
}

function memberEdit_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var mem = memberById_(body.id); if(!mem) return out_({ ok:false, error:'no-member' });
    var req = {};
    ['name','mobile','dob','joinDate'].forEach(function(k){ if(body[k] !== undefined) req[k] = body[k]; });
    var chk = checkMemberFields_(req, mem.MemberID);
    if(chk.error) return out_({ ok:false, error:chk.error, member:chk.member });
    var changes = [], upd = {};
    Object.keys(chk.fields).forEach(function(k){ if(String(mem[k] || '') !== String(chk.fields[k])){ upd[k] = chk.fields[k]; changes.push(k + ': ' + (mem[k] || '-') + ' > ' + (chk.fields[k] || '-')); } });
    if(!changes.length) return out_({ ok:true, unchanged:true });
    update_('Members', mem._row, upd);
    audit_(S, 'member-edited', mem.MemberID, changes.join('; '));
    return out_({ ok:true });
  } finally { lock.releaseLock(); }
}

// Manual suspension (separate from loan suspension). withusadmin1 only. Exit is handled later.
function memberStatus_(S, body){
  var bad = need_(S, 'admin1') || needDb_(); if(bad) return bad;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var mem = memberById_(body.id); if(!mem) return out_({ ok:false, error:'no-member' });
    var to = body.status === 'Suspended' ? 'Suspended' : body.status === 'Active' ? 'Active' : '';
    if(!to || String(mem.Status) === 'Exited') return out_({ ok:false, error:'bad-status' });
    if(String(mem.Status) === to) return out_({ ok:true, unchanged:true });
    update_('Members', mem._row, { Status:to });
    audit_(S, to === 'Suspended' ? 'member-suspended' : 'member-reactivated', mem.MemberID, mem.Name);
    return out_({ ok:true, status:to });
  } finally { lock.releaseLock(); }
}

function monthView_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var month = String(body.month || curMonth_()); if(!validMonth_(month)) return out_({ ok:false, error:'bad-months' });
  var ci = contribIndex_(), amt = contribAmount_(), rowsOut = [], expected = 0, paidN = 0, collected = 0;
  rows_('Members').forEach(function(m){
    var p = (ci[m.MemberID] || {})[month], elig = String(m.Status) !== 'Exited' && startMonth_(m) <= month;
    if(!elig && !p) return;
    if(elig) expected++;
    if(p){ paidN++; collected += p.amount; }
    rowsOut.push({ id:m.MemberID, name:m.Name, status:m.Status || 'Active', paid:!!p, amount:p ? p.amount : amt, date:p ? p.date : '', mode:p ? p.mode : '', receipt:p ? p.receipt : '', by:p ? p.by : '', at:p ? p.at : '' });
  });
  rowsOut.sort(function(a, b){ return (a.paid === b.paid) ? (a.name < b.name ? -1 : 1) : (a.paid ? 1 : -1); });
  return out_({ ok:true, month:month, amount:amt, rows:rowsOut, totals:{ expected:expected, paid:paidN, collected:collected, outstanding:Math.max(0, expected - paidN) * amt } });
}

// Paid / expected / collected for every tracked month, in one light call (for the Monthly tab's "All months" list).
function monthsSummary_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var ci = contribIndex_(), start = trackStart_(), cur = curMonth_(), list = rows_('Members'), res = {}, g = 0;
  for(var m = start; m <= cur && g < 120; m = addMonths_(m, 1), g++){
    var exp = 0, paid = 0, col = 0;
    list.forEach(function(x){
      var p = (ci[x.MemberID] || {})[m], elig = String(x.Status) !== 'Exited' && startMonth_(x) <= m;
      if(elig) exp++;
      if(p){ paid++; col += p.amount; }
    });
    res[m] = { expected:exp, paid:paid, collected:col };
  }
  return out_({ ok:true, months:res });
}

function memberProfile_(S, body){
  var bad = needDb_(); if(bad) return bad;
  var mem = memberById_(body.id); if(!mem) return out_({ ok:false, error:'no-member' });
  var cur = curMonth_(), ci = contribIndex_()[mem.MemberID] || {}, di = depositTotals_();
  var groups = {}, order = [];
  rows_('Contributions').forEach(function(c){
    if(c.MemberID !== mem.MemberID || isTrue_(c.Voided)) return;
    var k = String(c.ReceiptNo || c.ContribID);
    if(!groups[k]){ groups[k] = { receipt:c.ReceiptNo || '', months:[], amount:0, date:String(c.PaidDate || ''), mode:c.Mode || '', ref:c.Ref || '', by:c.EnteredBy || '', at:c.EnteredAt || '', source:c.Source || '', contribId:c.ContribID }; order.push(k); }
    groups[k].months.push(String(c.Month)); groups[k].amount += num_(c.Amount);
  });
  var recs = order.map(function(k){ groups[k].months.sort(); return groups[k]; });
  recs.sort(function(a, b){ return (b.months[b.months.length - 1] || '').localeCompare(a.months[a.months.length - 1] || '') || String(b.at).localeCompare(String(a.at)); });
  var loansById = {}, loans = [];
  var lpi = loanPaidIndex_(), lgi = guarantorsOf_(), tdy = todayStr_();
  rows_('Loans').forEach(function(l){
    loansById[l.LoanID] = l;
    if(l.MemberID !== mem.MemberID || !num_(l.Amount)) return;
    var f = loanFig_(l, lpi, tdy);
    loans.push({ id:f.id, amount:f.amount, taken:f.taken, due:f.due, returnAmount:num_(l.ReturnAmount), status:f.status, extra:f.extra, late:f.late, paid:f.paid, outstanding:f.outstanding, total:f.total, daysLeft:f.daysLeft, closedDate:f.closedDate, guarantors:(lgi[f.id] || []).map(function(g){ return g.name; }) });
  });
  loans.sort(function(a, b){ return b.taken.localeCompare(a.taken); });
  var gf = rows_('Guarantors').filter(function(g){ return g.MemberID === mem.MemberID; }).map(function(g){
    var l = loansById[g.LoanID] || {}; return { loan:g.LoanID, borrower:l.ApplicantName || '', amount:num_(l.Amount), due:String(l.DueDate || ''), status:l.Status || '' };
  });
  gf.sort(function(a, b){ return b.due.localeCompare(a.due); });
  var pen = penaltyIndex_()[mem.MemberID] || { due:0, charged:0, paid:0, waived:0, list:[] };
  var pgroups = {}, porder = [];
  rows_('Ledger').forEach(function(l){
    if(l.Type !== 'penalty_paid' || isTrue_(l.Voided) || l.MemberID !== mem.MemberID) return;
    var k = String(l.ReceiptNo || l.LedgerID);
    if(!pgroups[k]){ pgroups[k] = { receipt:l.ReceiptNo || '', amount:0, date:String(l.Date || ''), by:l.EnteredBy || '', at:l.EnteredAt || '', months:[] }; porder.push(k); }
    pgroups[k].amount += num_(l.Amount); pgroups[k].months.push(String(l.Note).slice(8, 15));
  });
  var ppay = porder.map(function(k){ return pgroups[k]; }).sort(function(a, b){ return String(b.at).localeCompare(String(a.at)); });
  return out_({ ok:true, member:memberPublic_(mem, standing_(mem, ci, cur), di[mem.MemberID], pen, suspensionIndex_()[mem.MemberID]), penalties:pen.list, penaltyDue:pen.due, penaltyPayments:ppay.slice(0, 20), penaltyAmount:penaltyAmount_(), amount:contribAmount_(), month:cur,
    paidMonths:Object.keys(ci).sort(), receipts:recs.slice(0, 40), loans:loans.slice(0, 20), loanCount:loans.length, guarantorFor:gf.slice(0, 20), guarantorCount:gf.length });
}

// Owner-only, from the script editor, for the testing period only: empties every data tab (not Admins, Settings, Audit),
// so the Excel import can be run again. Needs the exact phrase.  Example: wipeAllData('YES-WIPE-ALL-DATA')
function wipeAllData(phrase){
  resetCache_();
  if(phrase !== 'YES-WIPE-ALL-DATA') throw new Error("Type the exact phrase: wipeAllData('YES-WIPE-ALL-DATA')");
  ['Members','Contributions','Loans','Guarantors','LoanPayments','Applications','Ledger','Penalties','Suspensions','Expenses','Exits','PaymentClaims','Reconciliations','MonthClose'].forEach(clearRows_);
  props_().deleteProperty('receiptSeq');
  audit_({ u:'owner', r:'' }, 'wipe-all-data', 'database', 'from script editor');
  Logger.log('All data tabs emptied. Sign in as withusadmin1 and import the Excel data again.');
}

/* ---------- reading (needs a signed-in session) ---------- */
function doGet(e){
  resetCache_();
  var p = (e && e.parameter) || {};
  var S = sess_(p.token);
  if(!S) return out_({ ok:false, error:'login-required' });
  if(S.mc) return out_({ ok:false, error:'must-change' });
  var since = Number(p.since || 0);
  // Fast path: the version of each shared file is remembered in a script property, so "nothing changed" never opens Drive.
  var vk = 'ver_' + (p.kind === 'contrib' || p.kind === 'repay' || p.kind === 'apply' ? p.kind : 'dataset');
  if(since && String(since) === String(props_().getProperty(vk) || '')){
    var fast = {}; if(vk !== 'ver_dataset') fast[p.kind] = true; fast.unchanged = true; fast.updatedAt = since; return out_(fast);
  }
  if(p.kind === 'contrib' || p.kind === 'repay' || p.kind === 'apply'){
    var cf = file_(p.kind === 'repay' ? REPAY_FILE : p.kind === 'apply' ? APPS_FILE : CONTRIB_FILE);
    var flag = {}; flag[p.kind] = true;
    if(!cf){ flag.empty = true; return out_(flag); }
    var c = JSON.parse(cf.getBlob().getDataAsString());
    props_().setProperty(vk, String(c.updatedAt || ''));
    if(since && since === c.updatedAt){ flag.unchanged = true; flag.updatedAt = c.updatedAt; return out_(flag); }
    flag.updatedAt = c.updatedAt; flag.months = c.months || {};
    return out_(flag);
  }
  var f = file_(FILE_NAME);
  if(!f) return out_({ empty:true });
  var saved = JSON.parse(f.getBlob().getDataAsString());
  props_().setProperty(vk, String(saved.updatedAt || ''));
  if(since && since === saved.updatedAt) return out_({ unchanged:true, updatedAt:saved.updatedAt });
  return out_(saved);
}

/* ---------- writing ---------- */
function doPost(e){
  resetCache_();
  try{
    var body = JSON.parse(e.postData.contents);
    if(body.kind === 'login') return login_(body);
    if(body.kind === 'logout') return logout_(body);
    var S = sess_(body.token);
    if(!S) return out_({ ok:false, error:'bad-session' });
    if(body.kind === 'check') return out_({ ok:true, user:S.u, role:S.r, mustChange:!!S.mc, db:sheetReady_() });
    if(S.mc && body.kind !== 'changePassword') return out_({ ok:false, error:'must-change' });   // a temporary password must be replaced first

    ensureSchema_();

    // members and contributions
    if(body.kind === 'members') return cachedCall_('members', function(){ return membersList_(S); });
    if(body.kind === 'memberProfile') return cachedCall_('mp|' + String(body.id), function(){ return memberProfile_(S, body); });
    if(body.kind === 'memberAdd') return memberAdd_(S, body);
    if(body.kind === 'memberEdit') return memberEdit_(S, body);
    if(body.kind === 'memberStatus') return memberStatus_(S, body);
    if(body.kind === 'monthView') return cachedCall_('mv|' + String(body.month), function(){ return monthView_(S, body); });
    if(body.kind === 'monthsSummary') return cachedCall_('ms', function(){ return monthsSummary_(S, body); });
    if(body.kind === 'penaltyPreview') return cachedCall_('pp|' + S.r + '|' + bodyKey_(body), function(){ return penaltyPreview_(S, body); });
    if(body.kind === 'penaltyPost') return penaltyPost_(S, body);
    if(body.kind === 'penaltyPay') return penaltyPay_(S, body);
    if(body.kind === 'penaltyPayVoid') return penaltyPayVoid_(S, body);
    if(body.kind === 'penaltyWaive') return penaltyWaive_(S, body);
    if(body.kind === 'finePot') return cachedCall_('fp|' + String(body.limit || ''), function(){ return finePot_(S, body); });
    if(body.kind === 'expenseAdd') return expenseAdd_(S, body);
    if(body.kind === 'expenseVoid') return expenseVoid_(S, body);
    if(body.kind === 'loansList') return cachedCall_('ll', function(){ return loansList_(S, body); });
    if(body.kind === 'loansDue') return cachedCall_('ldue|' + String(body.month || ''), function(){ return loansDue_(S, body); });
    if(body.kind === 'loanDetail') return cachedCall_('ld|' + String(body.loanId), function(){ return loanDetail_(S, body); });
    if(body.kind === 'loanPay') return loanPay_(S, body);
    if(body.kind === 'loanPayVoid') return loanPayVoid_(S, body);
    if(body.kind === 'loanWaiveLate') return loanWaiveLate_(S, body);
    if(body.kind === 'suspensionLift') return suspensionLift_(S, body);
    if(body.kind === 'settlePreview') return settlePreview_(S, body);
    if(body.kind === 'settleConfirm') return settleConfirm_(S, body);
    if(body.kind === 'exitPreview') return exitPreview_(S, body);
    if(body.kind === 'exitConfirm') return exitConfirm_(S, body);
    if(body.kind === 'closeView') return cachedCall_('cv|' + S.r, function(){ return closeView_(S); });
    if(body.kind === 'bankCheck') return bankCheck_(S, body);
    if(body.kind === 'monthClose') return monthClose_(S, body);
    if(body.kind === 'dashboard') return cachedCall_('db', function(){ return dashboard_(S); });
    if(body.kind === 'report') return cachedCall_('rp|' + S.r + '|' + bodyKey_(body), function(){ return report_(S, body); });
    if(body.kind === 'excelCheck') return excelCheck_(S, body);
    if(body.kind === 'goLiveView') return goLiveView_(S);
    if(body.kind === 'goLiveConfirm') return goLiveConfirm_(S, body);
    if(body.kind === 'waPlan') return waPlan_(S, body);
    if(body.kind === 'waApply') return waApply_(S, body);
    if(body.kind === 'cleanupPreview') return cleanupPreview_(S, body);
    if(body.kind === 'cleanupApply') return cleanupApply_(S, body);
    if(body.kind === 'settingsGet') return cachedCall_('sg|' + S.r, function(){ return settingsGet_(S); });
    if(body.kind === 'settingsSet') return settingsSet_(S, body);
    if(body.kind === 'loanDisburse') return loanDisburse_(S, body);
    if(body.kind === 'contribAdd') return contribAdd_(S, body);
    if(body.kind === 'contribVoid') return contribVoid_(S, body);

    // admin accounts and the database
    if(body.kind === 'changePassword') return changePassword_(S, body, body.token);
    if(body.kind === 'adminList') return adminList_(S);
    if(body.kind === 'setPassword') return setPassword_(S, body);
    if(body.kind === 'setActive') return setActive_(S, body);
    if(body.kind === 'signOutAll') return signOutAll_(S, body);
    if(body.kind === 'auditList') return cachedCall_('al|' + S.r + '|' + bodyKey_(body), function(){ return auditList_(S, body); });
    if(body.kind === 'dbStatus') return dbStatus_(S);
    if(body.kind === 'importDataset') return importDataset_(S, body);
    if(body.kind === 'totalsCheck') return totalsCheck_(S);
    if(body.kind === 'exportAll') return exportAll_(S);

    // existing app data (unchanged)
    if(body.kind === 'contrib') return saveMarks_(body, CONTRIB_FILE, false);
    if(body.kind === 'repay') return saveMarks_(body, REPAY_FILE, true);
    if(body.kind === 'apply') return saveMarks_(body, APPS_FILE, 'apply');

    // Excel dataset upload (never write an empty dataset)
    if(!body.dataset || !body.dataset.rows || !body.dataset.rows.length) return out_({ ok:false, error:'no-dataset' });
    var saved = { updatedAt: Date.now(), dataset: body.dataset };
    props_().setProperty('ver_dataset', String(saved.updatedAt));
    var txt = JSON.stringify(saved), f = file_(FILE_NAME);
    if(f) f.setContent(txt); else DriveApp.createFile(FILE_NAME, txt, MimeType.PLAIN_TEXT);
    return out_({ ok:true, updatedAt:saved.updatedAt });
  }catch(err){ return out_({ ok:false, error:String(err) }); }
}

// body.month = 'YYYY-MM', body.marks = { key: value | null }   (null removes the mark)
//   contributions: value is true | false
//   repayments:    value is { v: true|false, d: 'YYYY-MM-DD' (optional date received) }
//   applications:  value is an entry { n:name, a:amount, g:[guarantors], gb:[true if that guarantor is rejected], st:'a'|'r', rm:remark, ts:time }
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
      if(objectValues === 'apply'){
        var ao = cleanApp_(v);
        if(ao) m[k] = ao; else delete m[k];
      } else if(objectValues){
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
    props_().setProperty('ver_' + (fileName === REPAY_FILE ? 'repay' : fileName === APPS_FILE ? 'apply' : 'contrib'), String(data.updatedAt));
    var txt = JSON.stringify(data);
    if(cf) cf.setContent(txt); else DriveApp.createFile(fileName, txt, MimeType.PLAIN_TEXT);
    return out_({ ok:true, updatedAt:data.updatedAt, months:data.months });
  } finally { lock.releaseLock(); }
}

// Keeps only the known fields of a loan application entry (and limits their size). Returns null for anything else.
function cleanApp_(v){
  if(!v || typeof v !== 'object' || typeof v.n !== 'string' || !v.n) return null;
  var amt = Number(v.a);
  if([10000, 25000, 50000, 100000].indexOf(amt) < 0) return null;
  var g = [], gb = [];
  (Array.isArray(v.g) ? v.g : []).slice(0, 2).forEach(function(x, i){
    g.push(String(x || '').slice(0, 80));
    gb.push(!!(Array.isArray(v.gb) && v.gb[i]));
  });
  var out = { n: v.n.slice(0, 80), a: amt, g: g, gb: gb, st: v.st === 'r' ? 'r' : 'a', rm: String(v.rm || '').slice(0, 400), ts: Number(v.ts) || Date.now() };
  if(typeof v.ds === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v.ds)) out.ds = v.ds;   // disbursed loan date
  return out;
}
