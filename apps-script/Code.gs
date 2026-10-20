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

var FILE_NAME = 'withus-dataset.json';              // Excel dataset
var CONTRIB_FILE = 'withus-contributions.json';     // monthly contribution marks
var REPAY_FILE = 'withus-repayments.json';          // monthly loan repayment marks
var APPS_FILE = 'withus-applications.json';         // monthly loan application list (from the Eligibility tab)

var DB_NAME = 'WITH US Society Database';
var ADMIN1 = 'withusadmin1', ADMIN2 = 'withusadmin2';

// Every tab of the database and its columns. Nothing is ever deleted from money tabs: rows are voided, not removed.
var TABS = {
  Members: ['MemberID','Name','Mobile','DOB','JoinDate','Status','OpeningDeposit','JoiningFee','PinHash','PinSalt','PinSetAt','FailedTries','LockedUntil','Source','CreatedAt','CreatedBy'],
  Contributions: ['ContribID','MemberID','Month','Amount','PaidDate','Mode','Ref','Source','EnteredBy','EnteredAt','Voided'],
  Loans: ['LoanID','MemberID','ApplicantName','Amount','TakenDate','DueDate','ReturnAmount','Fine','LateExtra','Status','Suspended','ClosedDate','Source','CreatedAt'],
  Guarantors: ['LoanID','MemberID','Name','Source'],
  LoanPayments: ['PaymentID','LoanID','Date','Amount','PaidBy','ToPrincipal','ToExtra','ToLateExtra','Mode','Ref','EnteredBy','EnteredAt','Voided'],
  Applications: ['AppID','Month','MemberID','Name','Amount','Guarantors','Status','Remark','CreatedAt'],
  Ledger: ['LedgerID','Date','Pot','Type','MemberID','LoanID','Amount','Ref','Note','EnteredBy','EnteredAt','Voided','VoidedBy','VoidedAt'],
  Penalties: ['PenaltyID','MemberID','Month','Amount','Status','WaivedBy','PaidDate','CreatedAt'],
  Suspensions: ['SuspensionID','MemberID','Reason','LoanID','StartDate','EndDate','LiftedBy'],
  Expenses: ['ExpenseID','Date','Category','Amount','Note','EnteredBy','EnteredAt','Voided'],
  Exits: ['ExitID','MemberID','Date','Refund','Note','EnteredBy','EnteredAt'],
  PaymentClaims: ['ClaimID','MemberID','Amount','Purpose','Months','LoanID','UpiRef','Status','ConfirmedBy','ClaimedAt','DecidedAt'],
  Reconciliations: ['ReconID','Date','BankBalance','CashInHand','ExpectedTotal','Difference','EnteredBy','EnteredAt'],
  MonthClose: ['Month','ClosedBy','ClosedAt'],
  Admins: ['Username','Role','Active','Salt','Hash','MustChange','LastLogin','CreatedAt'],
  Audit: ['AuditID','At','Admin','Role','Action','Target','Detail'],
  Settings: ['Key','Value','Note','UpdatedAt','UpdatedBy']
};
var NUM_COLS = ['Amount','OpeningDeposit','JoiningFee','ReturnAmount','Fine','LateExtra','ToPrincipal','ToExtra','ToLateExtra','Refund','BankBalance','CashInHand','ExpectedTotal','Difference','FailedTries'];
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
function props_(){ return PropertiesService.getScriptProperties(); }
function sheetReady_(){ return !!props_().getProperty('sheetId'); }
function ss_(){
  var id = props_().getProperty('sheetId');
  if(!id) throw new Error('not-set-up');
  return SpreadsheetApp.openById(id);
}
function tab_(name){
  var s = ss_().getSheetByName(name);
  if(!s) throw new Error('missing-tab-' + name);
  return s;
}
function fmts_(name){ return TABS[name].map(function(h){ return NUM_COLS.indexOf(h) >= 0 ? '0.##' : '@'; }); }
function rows_(name){
  var v = tab_(name).getDataRange().getValues();
  if(v.length < 2) return [];
  var h = v[0], out = [];
  for(var i = 1; i < v.length; i++){
    var o = { _row: i + 1 };
    for(var j = 0; j < h.length; j++) o[h[j]] = v[i][j];
    out.push(o);
  }
  return out;
}
function appendMany_(name, list){
  if(!list.length) return;
  var s = tab_(name), h = TABS[name], f = fmts_(name), r = s.getLastRow() + 1;
  var vals = list.map(function(o){ return h.map(function(k){ return o[k] === undefined || o[k] === null ? '' : o[k]; }); });
  var rng = s.getRange(r, 1, vals.length, h.length);
  rng.setNumberFormats(vals.map(function(){ return f; }));
  rng.setValues(vals);
}
function append_(name, o){ appendMany_(name, [o]); }
function update_(name, rowNum, o){
  var s = tab_(name), h = TABS[name], rng = s.getRange(rowNum, 1, 1, h.length);
  var cur = rng.getValues()[0];
  h.forEach(function(k, j){ if(Object.prototype.hasOwnProperty.call(o, k)) cur[j] = o[k]; });
  rng.setValues([cur]);
}
function clearRows_(name){
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
  var members = rows_('Members'), loans = rows_('Loans'), ledger = rows_('Ledger').filter(function(r){ return !isTrue_(r.Voided); });
  function sum(list, f){ return list.reduce(function(s, x){ return s + num_(f(x)); }, 0); }
  var lines = [];
  function line(label, excel, sheet){ lines.push({ label:label, excel:excel, sheet:sheet, ok:Math.abs(excel - sheet) < 0.5 }); }
  line('Members', md.members.length, members.length);
  var depX = sum(md.members, function(m){ return m.depositCurrent; });
  var depS = sum(ledger.filter(function(r){ return r.Pot === 'deposit'; }), function(r){ return r.Amount; });
  line('Deposit pot (total of member deposits)', depX, depS);
  var openX = sum(md.members, function(m){ return m.depositPrior; });
  line('Opening deposits', openX, sum(members, function(m){ return m.OpeningDeposit; }));
  line('Loans (count)', rowsX.length, loans.length);
  line('Loans (amount taken)', sum(rowsX, function(r){ return r.takenAmt; }), sum(loans, function(r){ return r.Amount; }));
  line('Loans (amount to return)', sum(rowsX, function(r){ return r.repayAmt; }), sum(loans, function(r){ return r.ReturnAmount; }));
  line('Loans (extra / fine)', sum(rowsX, function(r){ return r.fine; }), sum(loans, function(r){ return r.Fine; }));
  var fp = fineFigures_(ds);
  line('Fine pot (remaining)', fp.remaining, sum(ledger.filter(function(r){ return r.Pot === 'fine'; }), function(r){ return r.Amount; }));
  // member by member: Excel current deposit against the ledger
  var byId = {}; ledger.forEach(function(r){ if(r.Pot === 'deposit') byId[r.MemberID] = (byId[r.MemberID] || 0) + num_(r.Amount); });
  var diffs = [];
  md.members.forEach(function(m, i){
    var id = 'M-' + pad_(i + 1, 4), a = num_(m.depositCurrent), b = byId[id] || 0;
    if(Math.abs(a - b) >= 0.5) diffs.push({ id:id, name:m.name, excel:a, sheet:b });
  });
  return { ok: lines.every(function(l){ return l.ok; }) && !diffs.length, lines:lines, memberDiffs:diffs.slice(0, 25), memberDiffCount:diffs.length };
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

/* ---------- one-time setup and backups (run from the script editor) ---------- */
function setup(){
  var id = props_().getProperty('sheetId'), book, fresh = false;
  if(id) book = SpreadsheetApp.openById(id);
  else { book = SpreadsheetApp.create(DB_NAME); props_().setProperty('sheetId', book.getId()); fresh = true; }
  var first = true;
  Object.keys(TABS).forEach(function(n){
    var sh = book.getSheetByName(n);
    if(!sh){
      sh = first && fresh ? book.getSheets()[0].setName(n) : book.insertSheet(n);
    }
    first = false;
    var h = TABS[n];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    sh.setFrozenRows(1);
    var f = fmts_(n);
    for(var c = 0; c < h.length; c++) sh.getRange(1, c + 1, Math.max(sh.getMaxRows(), 2), 1).setNumberFormat(f[c]);
  });
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
  var row = adminRow_(ADMIN1);
  if(!row) throw new Error('Run setup first');
  var temp = tempPass_(), salt = newSalt_();
  update_('Admins', row._row, { Salt:salt, Hash:hash_(temp, salt), MustChange:true, Active:true });
  dropSessionsOf_(ADMIN1);
  audit_({ u:'owner', r:'' }, 'reset-admin1', ADMIN1, 'from script editor');
  Logger.log('New TEMPORARY password for ' + ADMIN1 + ': ' + temp);
}

function nightlyBackup(){
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

/* ---------- reading (needs a signed-in session) ---------- */
function doGet(e){
  var p = (e && e.parameter) || {};
  var S = sess_(p.token);
  if(!S) return out_({ ok:false, error:'login-required' });
  if(S.mc) return out_({ ok:false, error:'must-change' });
  var since = Number(p.since || 0);
  if(p.kind === 'contrib' || p.kind === 'repay' || p.kind === 'apply'){
    var cf = file_(p.kind === 'repay' ? REPAY_FILE : p.kind === 'apply' ? APPS_FILE : CONTRIB_FILE);
    var flag = {}; flag[p.kind] = true;
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
    var S = sess_(body.token);
    if(!S) return out_({ ok:false, error:'bad-session' });
    if(body.kind === 'check') return out_({ ok:true, user:S.u, role:S.r, mustChange:!!S.mc, db:sheetReady_() });
    if(S.mc && body.kind !== 'changePassword') return out_({ ok:false, error:'must-change' });   // a temporary password must be replaced first

    // admin accounts and the database
    if(body.kind === 'changePassword') return changePassword_(S, body, body.token);
    if(body.kind === 'adminList') return adminList_(S);
    if(body.kind === 'setPassword') return setPassword_(S, body);
    if(body.kind === 'setActive') return setActive_(S, body);
    if(body.kind === 'signOutAll') return signOutAll_(S, body);
    if(body.kind === 'auditList') return auditList_(S, body);
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
