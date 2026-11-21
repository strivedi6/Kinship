/**
 * KINSHIP - Google Apps Script backend
 * 1. Open the "Kinship" Google Sheet > Extensions > Apps Script, and paste this whole file.
 * 2. Run the function  setup  once (builds the tabs, links the Media folder, starts the daily birthday reminder).
 * 3. Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
 * 4. Paste the Web app URL into index.html (CONFIG.API_URL).
 * Notifications: paste the OneSignal App ID and REST API key in the Settings tab.
 */

// Drive folder "Kinship / Media". Each person gets their own subfolder inside it.
const MEDIA_FOLDER_ID = '1Nr_Q34Kr27e9RV5KPkL0VA6Ye0Ty8UnG';

const SCHEMA = {
  Members: ['id','name','nick','gender','role','birth','passed','country','city','tz','fatherId','motherId','spouseId','anniv','avatar','about','favFood','favColor','hobby','askMe','by','ts','folder','pinHash','salt','fails','lockUntil','pinLen','viewOnly','photo'],
  Posts: ['id','by','text','photo','ts','tags'],
  Comments: ['id','postId','by','text','ts','tags'],
  Reactions: ['id','postId','by','emoji','ts'],
  Events: ['id','by','title','emoji','date','at','yearly','note','ts','travel','dinner','venue','with','invited','rsvp','kind','dep','arr','fromCode','fromCity','fromTz','fromTerm','toCode','toCity','toTz','toTerm','flightNo'],
  Habits: ['id','by','title','emoji','points','who','days','active','ts'],
  Logs: ['id','habitId','memberId','date','points','by','ts'],
  Rewards: ['id','by','title','emoji','cost','who','active','ts'],
  Redemptions: ['id','rewardId','memberId','title','cost','status','by','decidedBy','ts'],
  Answers: ['id','week','memberId','text','ts'],
  Wishes: ['id','memberId','title','emoji','note','cost','by','ts'],
  Messages: ['id','by','to','text','ts','tags'],
  Sessions: ['token','memberId','ts'],
  Settings: ['key','value']
};
const SECRET = ['folder','pinHash','salt','fails','lockUntil','pinLen'];
const DATA_TABLES = ['Members','Posts','Comments','Reactions','Events','Habits','Logs','Rewards','Redemptions','Answers','Wishes','Messages'];
const WRITES = ['join','login','put','del','upload','settings','setPin','signout','link','cleanup','rsvp'];

/* ---------------- one-time setup ---------------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SCHEMA).forEach(function (name) {
    let s = ss.getSheetByName(name) || ss.insertSheet(name);
    const head = SCHEMA[name];
    s.getRange(1, 1, s.getMaxRows(), head.length).setNumberFormat('@');
    s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    s.setFrozenRows(1);
  });
  const first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0) ss.deleteSheet(first);
  if (!setting('appName')) setSetting('appName', 'Kinship');
  if (!setting('familyCode')) setSetting('familyCode', randomCode());
  if (!setting('photoFolder')) setSetting('photoFolder', MEDIA_FOLDER_ID);
  ['oneSignalAppId', 'oneSignalKey', 'appUrl'].forEach(function (k) { if (!findBy('Settings', k, 'key')) setSetting(k, ''); });
  DriveApp.getFolderById(setting('photoFolder')).getName();   // checks the Media folder can be reached
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyDigest') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyDigest').timeBased().everyDays(1).atHour(8).create();
  ensureTriggers();
  const msg = 'Ready!\n\nYour family code is: ' + setting('familyCode') + '\n\nYou can change it in the Settings tab or inside the app.\nNow deploy this script as a Web app.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}
function onOpen() { SpreadsheetApp.getUi().createMenu('Family App').addItem('Set up / repair tabs', 'setup').addToUi(); }
function randomCode() { const a = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; let s = ''; for (let i = 0; i < 8; i++) s += a.charAt(Math.floor(Math.random() * a.length)); return s.slice(0, 4) + '-' + s.slice(4); }

/* ---------------- web entry ---------------- */
function doGet() { return out({ ok: true, app: 'family' }); }
function doPost(e) {
  let q;
  try { q = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'Bad request' }); }
  const lock = LockService.getScriptLock(), write = WRITES.indexOf(q.action) >= 0;
  try {
    if (write) lock.waitLock(25000);
    return out({ ok: true, data: route(q) });
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  } finally {
    if (write) try { lock.releaseLock(); } catch (x) {}
  }
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function route(q) {
  switch (q.action) {
    case 'boot': checkCode(q.code); return { appName: setting('appName'), members: readAll('Members').map(function (m) { return { id: m.id, name: m.name, nick: m.nick, avatar: m.avatar, claimed: !!m.pinHash, pinLen: m.pinLen || '' }; }) };
    case 'join': return join(q);
    case 'login': return login(q);
    case 'msgs': return msgsPoll(q);
  }
  const me = auth(q.token);
  if (me.viewOnly === '1' && ['link', 'put', 'del', 'upload', 'setPin', 'settings', 'sigsend', 'rsvp'].indexOf(q.action) >= 0) throw new Error('This account is view only');
  switch (q.action) {
    case 'sync':
      if (me.role === 'admin' && /^https:\/\//.test(String(q.appUrl || '')) && setting('appUrl') !== q.appUrl) setSetting('appUrl', q.appUrl);
      return syncAll(me);
    case 'link': return link(me, q);
    case 'put': return put(me, q.table, q.row || {});
    case 'del': return del(me, q.table, q.id);
    case 'cleanup': if (me.role !== 'admin') throw new Error('Only an admin can do this'); try { ensureTriggers(); } catch (e) {} return runCleanup();
    case 'upload': return upload(me, q.dataUrl);
    case 'img': return img(q.id);
    case 'setPin': setPin(me, q.pin); return {};
    case 'settings': return saveSettings(me, q);
    case 'read': return saveReads(me, q.map);
    case 'rsvp': return rsvp(me, q);
    case 'sigsend': return sigSend(me, q);
    case 'sigpoll': return sigPoll(me, q);
    case 'signout': delRow('Sessions', q.token, 'token'); CacheService.getScriptCache().remove('s:' + q.token); return {};
  }
  throw new Error('Unknown action');
}

/* ---------------- sheet helpers ---------------- */
function sh(name) {   // a tab added in a later version is created the first time it is needed
  const ss = SpreadsheetApp.getActive(); let s = ss.getSheetByName(name);
  if (!s) { if (!SCHEMA[name]) throw new Error('Unknown tab'); s = ss.insertSheet(name); s.getRange(1, 1, s.getMaxRows(), SCHEMA[name].length).setNumberFormat('@'); s.getRange(1, 1, 1, SCHEMA[name].length).setValues([SCHEMA[name]]).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function str(v) { if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd'); return v === null || v === undefined ? '' : String(v); }
let MEMO = {};   // each tab is read once per request
function readAll(name) {
  if (MEMO[name]) return MEMO[name];
  const s0 = sh(name), v = s0.getDataRange().getValues(), head = SCHEMA[name], rows = [];
  if (v[0].length < head.length || String(v[0][head.length - 1]) !== head[head.length - 1]) s0.getRange(1, 1, 1, head.length).setNumberFormat('@').setValues([head]);   // new columns added by an update
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === '') continue;
    const o = { _row: i + 1 };
    for (let j = 0; j < head.length; j++) o[head[j]] = str(v[i][j]);
    rows.push(o);
  }
  MEMO[name] = rows;
  return rows;
}
function dropRow(name, row) { sh(name).deleteRow(row); delete MEMO[name]; }
function ageOf(birth) { const p = String(birth).split('-').map(Number), n = new Date(); let a = n.getFullYear() - p[0]; if (n.getMonth() + 1 < p[1] || (n.getMonth() + 1 === p[1] && n.getDate() < p[2])) a--; return a; }
function roleFor(birth, fallback) { return /^\d{4}-\d{2}-\d{2}$/.test(String(birth || '')) ? (ageOf(birth) < 18 ? 'kid' : 'adult') : (fallback === 'kid' ? 'kid' : 'adult'); }
function safe(v, max) { let s = str(v).slice(0, max || 3000); if (/^[=+\-@]/.test(s)) s = "'" + s; return s; }
function writeRow(name, obj, rowNum) {
  const s = sh(name), head = SCHEMA[name];
  const vals = head.map(function (h) { return safe(obj[h], h === 'avatar' ? 48000 : 3000); });
  const r = rowNum || s.getLastRow() + 1;
  s.getRange(r, 1, 1, head.length).setNumberFormat('@').setValues([vals]);
  delete MEMO[name];
}
function findBy(name, id, key) { key = key || 'id'; const all = readAll(name); for (let i = 0; i < all.length; i++) if (all[i][key] === id) return all[i]; return null; }
function delRow(name, id, key) { const r = findBy(name, id, key); if (r) dropRow(name, r._row); }
function setting(k) { const r = findBy('Settings', k, 'key'); return r ? r.value : ''; }
function setSetting(k, v) { const r = findBy('Settings', k, 'key'); writeRow('Settings', { key: k, value: v }, r ? r._row : 0); }
function now() { return String(Date.now()); }
function newId() { return Utilities.getUuid().replace(/-/g, '').slice(0, 14); }

/* ---------------- sign-in ---------------- */
function checkCode(code) { const c = setting('familyCode'); if (!c || String(code || '').trim().toUpperCase() !== c.toUpperCase()) throw new Error('That family code is not right'); }
function hash(pin, salt) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin, Utilities.Charset.UTF_8)); }
function checkPin(pin) { if (!/^[0-9]{4,8}$/.test(String(pin || ''))) throw new Error('PIN must be 4 to 8 digits'); }
function session(memberId) { const t = Utilities.getUuid() + Utilities.getUuid(); writeRow('Sessions', { token: t, memberId: memberId, ts: now() }); return { token: t, id: memberId, data: syncAll(findBy('Members', memberId)) }; }
/** Sessions are remembered for a few minutes so chat checks do not have to read the sheet every time. */
function sessionOf(token) {
  if (!token) return null;
  const c = CacheService.getScriptCache(), k = 's:' + token; let v = c.get(k);
  if (!v) { const s = findBy('Sessions', token, 'token'); if (!s) return null; v = s.memberId + '|' + s.ts; c.put(k, v, 600); }
  const p = v.split('|'); return { memberId: p[0], ts: Number(p[1]) };
}
function sessionHours() {
  const c = CacheService.getScriptCache(); let v = c.get('sessionHours');
  if (v === null) { v = String(Number(setting('sessionHours')) || 0); c.put('sessionHours', v, 600); }
  return Number(v) || 0;
}
function expired(sess) { const h = sessionHours(); return h > 0 && Date.now() - sess.ts > h * 3600000; }
function auth(token) {
  const sess = sessionOf(token);
  if (sess && expired(sess)) { delRow('Sessions', token, 'token'); CacheService.getScriptCache().remove('s:' + token); throw new Error('Session expired, please sign in again'); }
  const m = sess && findBy('Members', sess.memberId);
  if (!m) throw new Error('Session expired, please sign in again');
  return m;
}
function join(q) {
  checkCode(q.code); checkPin(q.pin);
  const salt = Utilities.getUuid();
  if (q.claimId) {
    const m = findBy('Members', q.claimId);
    if (!m) throw new Error('Profile not found');
    if (m.pinHash) throw new Error('This profile already has a PIN');
    m.salt = salt; m.pinHash = hash(q.pin, salt); m.pinLen = String(String(q.pin).length);
    writeRow('Members', m, m._row);
    return session(m.id);
  }
  const name = String(q.name || '').trim();
  if (!name) throw new Error('Please enter your name');
  // Same name again: a repeated tap signs in to the profile just made; otherwise point them to their existing name.
  const same = readAll('Members').filter(function (x) { return x.name.trim().toLowerCase() === name.toLowerCase(); });
  for (let i = 0; i < same.length; i++) if (same[i].pinHash && hash(q.pin, same[i].salt) === same[i].pinHash) return session(same[i].id);
  if (same.length) throw new Error(name + ' is already in the family. Go back and tap your name, or add a middle name if you are someone else.');
  const firstEver = readAll('Members').filter(function (m) { return m.pinHash; }).length === 0;
  const birth = /^\d{4}-\d{2}-\d{2}$/.test(String(q.birth || '')) ? q.birth : '';
  const m = { id: newId(), name: name, birth: birth, role: firstEver ? 'admin' : roleFor(birth, q.role), ts: now(), salt: salt, pinHash: hash(q.pin, salt), pinLen: String(String(q.pin).length) };
  m.by = m.id;
  writeRow('Members', m);
  if (!firstEver) push(readAll('Members').filter(function (x) { return x.role === 'admin'; }).map(function (x) { return x.id; }), name + ' joined the family', 'Say hello and help them find their place in the tree.', 'profile:' + m.id);
  return session(m.id);
}
function login(q) {
  checkCode(q.code);
  const m = findBy('Members', q.id);
  if (!m || !m.pinHash) throw new Error('Profile not found');
  if (Number(m.lockUntil || 0) > Date.now()) throw new Error('Too many tries. Please wait 15 minutes.');
  if (hash(String(q.pin || ''), m.salt) !== m.pinHash) {
    m.fails = String(Number(m.fails || 0) + 1);
    if (Number(m.fails) >= 5) { m.fails = '0'; m.lockUntil = String(Date.now() + 15 * 60000); }
    writeRow('Members', m, m._row);
    throw new Error('Wrong PIN');
  }
  const len = String(String(q.pin).length);
  if (m.fails !== '0' || m.lockUntil || m.pinLen !== len) { m.fails = '0'; m.lockUntil = ''; m.pinLen = len; writeRow('Members', m, m._row); }
  return session(m.id);
}
function setPin(me, pin) { checkPin(pin); me.salt = Utilities.getUuid(); me.pinHash = hash(pin, me.salt); me.pinLen = String(String(pin).length); writeRow('Members', me, me._row); }

/* ---------------- data ---------------- */
function pub(name, r) { const o = {}; SCHEMA[name].forEach(function (h) { if (SECRET.indexOf(h) < 0) o[h] = r[h]; }); if (name === 'Members') o.claimed = !!r.pinHash; return o; }
/* the time each person last read each chat, so unread counts follow them to any device */
function readsOf(me) { try { return JSON.parse(setting('read:' + me.id) || 'null'); } catch (e) { return null; } }
function saveReads(me, map) {
  const cur = readsOf(me) || {}, ok = {};
  Object.keys(map || {}).slice(0, 80).forEach(function (k) { if (/^[\w-]{1,40}$/.test(k)) { const v = Number(map[k]) || 0; ok[k] = Math.max(Number(cur[k]) || 0, v); } });
  Object.keys(cur).forEach(function (k) { if (!(k in ok)) ok[k] = cur[k]; });
  setSetting('read:' + me.id, JSON.stringify(ok));
  return { reads: ok };
}
function syncAll(me) {
  const d = { me: me.id, settings: { appName: setting('appName'), oneSignalAppId: setting('oneSignalAppId'), sessionHours: String(Number(setting('sessionHours')) || 0), celebrations: setting('celebrations') || '' } };
  if (me.role !== 'kid') d.settings.familyCode = setting('familyCode');
  DATA_TABLES.forEach(function (t) { if (t !== 'Messages') d[t.toLowerCase()] = readAll(t).map(function (r) { return pub(t, r); }); });
  d.messages = messagesFor(me, 0).slice(-600);
  d.reads = readsOf(me);
  return d;
}
/** Chat check. With wait=1 it holds the line for up to 14 seconds and answers the moment a new message is written,
    so messages arrive almost at once without reading the sheet on every check. */
function msgsPoll(q) {
  const sess = sessionOf(q.token);
  if (!sess || expired(sess)) throw new Error('Session expired, please sign in again');
  const c = CacheService.getScriptCache(), since = Number(q.since) || 0, until = Date.now() + (q.wait ? 14000 : 0);
  let latest;
  for (;;) {
    latest = c.get('msgTs');
    if (latest === null) { delete MEMO.Messages; latest = String(readAll('Messages').reduce(function (x, m) { return Math.max(x, Number(m.ts) || 0); }, 0)); c.put('msgTs', latest, 21600); }
    if (Number(latest) > since || Date.now() >= until) break;
    Utilities.sleep(350);
  }
  const sigs = takeSigs(sess.memberId);
  if (Number(latest) <= since) return { messages: [], ts: Number(latest), sigs: sigs };
  delete MEMO.Messages;
  return { messages: messagesFor({ id: sess.memberId }, since - 2000), ts: Number(latest), sigs: sigs };
}
/** Chat: the family group plus this person's own one-to-one messages. Nobody else's are ever sent. */
function messagesFor(me, since) {
  return readAll('Messages').filter(function (m) { return (m.to === 'all' || m.by === me.id || m.to === me.id) && Number(m.ts) > since; }).map(function (m) { return pub('Messages', m); });
}
function isParent(p, kidId) { const k = findBy('Members', kidId); return !!k && (k.fatherId === p.id || k.motherId === p.id); }
function cleanEvent(r, old) {
  const ids = {}; readAll('Members').forEach(function (m) { ids[m.id] = 1; });
  const pick = function (v) { const seen = {}; return String(v || '').split(',').filter(function (x) { if (!ids[x] || seen[x]) return false; seen[x] = 1; return true; }); };
  r.invited = pick(r.invited).join(','); r.with = pick(r.with).join(',');
  r.travel = /^\d{4}-\d{2}-\d{2}$/.test(r.travel || '') ? r.travel : '';
  r.dinner = isNaN(Date.parse(r.dinner || '')) ? '' : String(r.dinner);
  r.venue = String(r.venue || '').slice(0, 120);
  r.kind = r.kind === 'flight' ? 'flight' : '';
  ['dep', 'arr'].forEach(function (k) { r[k] = r.kind && !isNaN(Date.parse(r[k] || '')) ? String(r[k]) : ''; });
  ['fromTz', 'toTz'].forEach(function (k) { r[k] = /^[A-Za-z_]+(\/[A-Za-z_+\-0-9]+)*$/.test(r[k] || '') ? String(r[k]) : ''; });
  ['fromCode', 'toCode'].forEach(function (k) { r[k] = String(r[k] || '').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 4); });
  ['fromCity', 'toCity', 'fromTerm', 'toTerm', 'flightNo'].forEach(function (k) { r[k] = String(r[k] || '').slice(0, 40); });
  // Answers belong to the people who gave them: keep them, drop those of anyone no longer invited.
  let a = {}; try { a = JSON.parse(old && old.rsvp || '{}') || {}; } catch (e) {}
  const keep = {}; r.invited.split(',').forEach(function (id) { if (a[id]) keep[id] = a[id]; });
  r.rsvp = JSON.stringify(keep);
}
function rsvp(me, q) {
  const ev = findBy('Events', String(q.id || '')); if (!ev) throw new Error('This date was removed');
  const who = String(q.memberId || me.id), ans = String(q.answer || '');
  if (['yes', 'maybe', 'no'].indexOf(ans) < 0) throw new Error('Choose Going, Maybe or Can\'t');
  if (!mayActFor(me, who)) throw new Error('Not allowed');
  if (String(ev.invited || '').split(',').indexOf(who) < 0) throw new Error('This invite is not for them');
  let a = {}; try { a = JSON.parse(ev.rsvp || '{}') || {}; } catch (e) {}
  a[who] = ans; ev.rsvp = JSON.stringify(a);
  writeRow('Events', ev, ev._row);
  const m = findBy('Members', who) || me;
  if (ev.by !== me.id) push([ev.by], shortName(m) + (ans === 'yes' ? ' is going' : ans === 'maybe' ? ' might come' : " can't come"), ev.title + ' · ' + niceDate(ev.date), 'cal');
  return { rsvp: ev.rsvp };
}
function isGrandparent(g, kidId) {
  const k = findBy('Members', kidId); if (!k) return false;
  return [k.fatherId, k.motherId].some(function (pid) { return pid && isParent(g, pid); });
}
function mayActFor(me, memberId) { return me.id === memberId || me.role === 'admin' || (me.role !== 'kid' && (isParent(me, memberId) || isGrandparent(me, memberId))); }
function adultOnly(me) { if (me.role === 'kid') throw new Error('Ask a grown-up to do this'); }
function ownerOnly(me, old) { if (old && old.by !== me.id && me.role !== 'admin') throw new Error('Only the person who added this can change it'); }

function put(me, table, row) {
  if (DATA_TABLES.indexOf(table) < 0) throw new Error('Unknown table');
  const head = SCHEMA[table], clean = {};
  head.forEach(function (h) { if (SECRET.indexOf(h) < 0 && row[h] !== undefined) clean[h] = str(row[h]); });
  if (!clean.id) clean.id = newId();
  if (!/^[A-Za-z0-9_\-]{1,80}$/.test(clean.id)) throw new Error('Bad id');
  const old = findBy(table, clean.id);
  const r = Object.assign({}, old || {}, clean);
  if (!old) { r.ts = now(); if (table !== 'Members' || !r.by) r.by = me.id; } else { r.by = old.by; r.ts = old.ts; }

  switch (table) {
    case 'Members':
      if (!String(r.name || '').trim()) throw new Error('Name is needed');
      if (!old) { adultOnly(me); r.by = me.id; r.role = roleFor(r.birth, r.role); }
      else {
        // Only the person, their parents, or the family admin may edit a profile.
        if (!(old.id === me.id || me.role === 'admin' || old.fatherId === me.id || old.motherId === me.id)) throw new Error('Only this person, their parents or the family admin can edit this profile');
        // Role comes from the date of birth (under 18 = kid). Only an admin can give or remove the admin role.
        if (me.role === 'admin' && clean.role === 'admin') r.role = 'admin';
        else if (old.role === 'admin' && !(me.role === 'admin' && clean.role && clean.role !== 'admin')) r.role = 'admin';
        else r.role = roleFor(r.birth, old.role);
        if (old.id === me.id && old.role === 'admin' && r.role !== 'admin') {
          const others = readAll('Members').filter(function (m) { return m.role === 'admin' && m.id !== me.id; });
          if (!others.length) r.role = 'admin';
        }
        SECRET.forEach(function (h) { r[h] = old[h]; });
      }
      // Only the family admin can make an account view only, or lift it.
      if (me.role !== 'admin') r.viewOnly = old ? (old.viewOnly || '') : '';
      if (r.avatar && !/^data:image\/(jpeg|png|webp);base64,/.test(r.avatar)) r.avatar = '';
      break;
    case 'Posts': case 'Comments':
      ownerOnly(me, old); r.by = old ? old.by : me.id;
      if (table === 'Posts') r.photo = String(r.photo || '').split(',').filter(String).slice(0, 5).join(',');
      r.tags = mentionIds(r.text, me.id);
      break;
    case 'Reactions':
      r.by = me.id; r.id = r.postId + '_' + me.id; break;
    case 'Answers':
      r.memberId = me.id; r.id = r.week + '_' + me.id; break;
    case 'Events': case 'Habits': case 'Rewards':
      adultOnly(me); ownerOnly(me, old);
      if (table === 'Events') cleanEvent(r, old);
      break;
    case 'Logs':
      if (!mayActFor(me, r.memberId)) throw new Error('Not allowed');
      r.id = r.habitId + '_' + r.memberId + '_' + r.date; r.by = me.id; break;
    case 'Wishes':
      if (!old) { if (!mayActFor(me, r.memberId)) throw new Error('Not allowed'); if (me.role === 'kid') r.cost = ''; }
      else { if (!mayActFor(me, old.memberId)) throw new Error('Not allowed'); r.memberId = old.memberId; if (me.role === 'kid') r.cost = old.cost; }   // only a grown-up sets the stars
      if (!String(r.title || '').trim()) throw new Error('What is the wish?');
      if (r.cost && !/^[1-9][0-9]{0,5}$/.test(r.cost)) r.cost = '';
      break;
    case 'Messages':
      if (old) throw new Error('Messages cannot be changed');
      r.by = me.id; r.text = String(r.text || '').trim().slice(0, 2000);
      if (!r.text) throw new Error('Write a message first');
      r.tags = r.to === 'all' ? mentionIds(r.text, me.id) : '';
      if (r.to !== 'all' && !findBy('Members', r.to)) throw new Error('Person not found');
      break;
    case 'Redemptions':
      if (!old) {
        if (!mayActFor(me, r.memberId)) throw new Error('Not allowed'); if (me.role === 'kid') r.status = 'requested';
        const rid = String(r.rewardId || ''), src = rid.indexOf('wish:') === 0 ? findBy('Wishes', rid.slice(5)) : findBy('Rewards', rid);
        if (src) { if (!src.cost) throw new Error('A grown-up needs to set the stars first'); r.cost = src.cost; r.title = src.title; if (src.memberId) r.memberId = src.memberId; }   // the price always comes from the sheet
      }
      else {
        if (me.role === 'kid' || !mayActFor(me, old.memberId)) throw new Error('A parent needs to approve this');
        const st = r.status; Object.assign(r, old); r.status = st; r.decidedBy = me.id;
      }
      if (['requested', 'approved', 'declined'].indexOf(r.status) < 0) r.status = 'requested';
      break;
  }
  const target = (r.id !== clean.id) ? findBy(table, r.id) : old;
  writeRow(table, r, target ? target._row : 0);
  if (old && table === 'Posts') trashPhotos(String(old.photo || '').split(',').filter(function (p) { return p && String(r.photo || '').split(',').indexOf(p) < 0; }));
  if (old && table === 'Members' && old.photo && r.photo !== old.photo) trashPhotos([old.photo]);
  if (table === 'Messages') CacheService.getScriptCache().put('msgTs', String(r.ts), 21600);
  announce(me, table, r, target);
  return pub(table, r);
}

function del(me, table, id) {
  if (DATA_TABLES.indexOf(table) < 0) throw new Error('Unknown table');
  const old = findBy(table, id); if (!old) return {};
  if (table === 'Members') { if (me.role !== 'admin' || id === me.id) throw new Error('Only an admin can remove people'); }
  else if (table === 'Logs' || table === 'Redemptions' || table === 'Wishes') { if (!mayActFor(me, old.memberId)) throw new Error('Not allowed'); }
  else if (table === 'Reactions') { if (old.by !== me.id) throw new Error('Not allowed'); }
  else if (table === 'Answers') { if (old.memberId !== me.id && me.role !== 'admin') throw new Error('Not allowed'); }
  else ownerOnly(me, old);
  dropRow(table, old._row);
  if (table === 'Posts') dropPostData(old);
  if (table === 'Members') dropMemberData(old);
  if (table === 'Members') readAll('Sessions').filter(function (s) { return s.memberId === id; }).reverse().forEach(function (s) { sh('Sessions').deleteRow(s._row); }); delete MEMO.Sessions;
  return {};
}


/* ---------------- tidy up: nothing is left behind when something is deleted ---------------- */
function purgeRows(name, test) {
  const gone = readAll(name).filter(test); if (!gone.length) return 0;
  const s = sh(name); gone.slice().reverse().forEach(function (r) { s.deleteRow(r._row); }); delete MEMO[name];
  return gone.length;
}
function trashPhotos(ids) { let n = 0; (ids || []).forEach(function (id) { if (!id) return; try { DriveApp.getFileById(id).setTrashed(true); n++; } catch (e) {} }); return n; }
function dropPostData(p) {
  trashPhotos(String(p.photo || '').split(','));
  purgeRows('Comments', function (c) { return c.postId === p.id; });
  purgeRows('Reactions', function (x) { return x.postId === p.id; });
}
function dropMemberData(m) {
  readAll('Posts').filter(function (p) { return p.by === m.id; }).forEach(function (p) { dropPostData(p); });
  purgeRows('Posts', function (p) { return p.by === m.id; });
  purgeRows('Comments', function (c) { return c.by === m.id; }); purgeRows('Reactions', function (x) { return x.by === m.id; });
  purgeRows('Logs', function (l) { return l.memberId === m.id; }); purgeRows('Redemptions', function (x) { return x.memberId === m.id; });
  purgeRows('Answers', function (a) { return a.memberId === m.id; }); purgeRows('Wishes', function (w) { return w.memberId === m.id; });
  purgeRows('Messages', function (x) { return x.by === m.id || x.to === m.id; });
  readAll('Events').forEach(function (e) {   // take them off other people's travel lists and invites
    const f = function (v) { return String(v || '').split(',').filter(function (x) { return x && x !== m.id; }).join(','); };
    if (String(e.with).indexOf(m.id) < 0 && String(e.invited).indexOf(m.id) < 0) return;
    e.with = f(e.with); e.invited = f(e.invited); cleanEvent(e, e); writeRow('Events', e, e._row);
  });
  purgeRows('Settings', function (s) { return s.key === 'read:' + m.id; });
  if (m.photo) trashPhotos([m.photo]);
  if (m.folder) { try { DriveApp.getFolderById(m.folder).setTrashed(true); } catch (e) {} }
}
/* finds anything that no longer belongs to something and removes it (rows in the sheet, files in Drive; Drive items go to the bin for 30 days) */
function runCleanup() {
  const members = {}, posts = {}, keep = {}, folders = {}, out = {};
  readAll('Members').forEach(function (m) { members[m.id] = 1; if (m.photo) keep[m.photo] = 1; if (m.folder) folders[m.folder] = 1; });
  if (!Object.keys(members).length) throw new Error('No people found, nothing was removed');
  readAll('Posts').forEach(function (p) { posts[p.id] = 1; String(p.photo || '').split(',').forEach(function (x) { if (x) keep[x] = 1; }); });
  out.comments = purgeRows('Comments', function (c) { return !posts[c.postId] || !members[c.by]; });
  out.reactions = purgeRows('Reactions', function (x) { return !posts[x.postId] || !members[x.by]; });
  out.posts = purgeRows('Posts', function (p) { return !members[p.by]; });
  out.logs = purgeRows('Logs', function (l) { return !members[l.memberId]; });
  out.redemptions = purgeRows('Redemptions', function (x) { return !members[x.memberId]; });
  out.answers = purgeRows('Answers', function (a) { return !members[a.memberId]; });
  out.wishes = purgeRows('Wishes', function (w) { return !members[w.memberId]; });
  out.messages = purgeRows('Messages', function (x) { return !members[x.by] || (x.to !== 'all' && !members[x.to]); });
  out.sessions = purgeRows('Sessions', function (s) { return !members[s.memberId]; });
  out.settings = purgeRows('Settings', function (s) { return /^read:/.test(s.key) && !members[s.key.slice(5)]; });
  out.photos = 0; out.folders = 0;
  try {
    const root = DriveApp.getFolderById(setting('photoFolder') || MEDIA_FOLDER_ID), cutoff = Date.now() - 6 * 3600 * 1000, it = root.getFolders();
    while (it.hasNext()) {
      const f = it.next();
      if (!folders[f.getId()]) { out.folders++; f.setTrashed(true); continue; }
      const fi = f.getFiles();
      while (fi.hasNext()) { const x = fi.next(); if (!keep[x.getId()] && x.getDateCreated().getTime() < cutoff) { x.setTrashed(true); out.photos++; } }
    }
  } catch (e) { out.driveError = String(e.message || e); }
  return out;
}

/* ---------------- photos (private, one Drive folder per person) ---------------- */
function memberFolder(me) {
  if (me.folder) { try { const f = DriveApp.getFolderById(me.folder); if (!f.isTrashed()) return f; } catch (e) {} }
  const f = DriveApp.getFolderById(setting('photoFolder') || MEDIA_FOLDER_ID).createFolder(me.name + ' (' + me.id + ')');
  me.folder = f.getId();
  writeRow('Members', me, me._row);
  return f;
}
function upload(me, dataUrl) {
  const m = /^data:(image\/(?:jpeg|png|webp));base64,(.+)$/.exec(String(dataUrl || ''));
  if (!m) throw new Error('That is not a photo');
  if (m[2].length > 4000000) throw new Error('Photo is too large');
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], stamp + '.jpg');
  return { id: memberFolder(me).createFile(blob).getId() };
}
function img(id) {
  // A post photo or someone's profile photo.
  if (!readAll('Posts').some(function (p) { return String(p.photo || '').split(',').indexOf(id) >= 0; }) &&
      !readAll('Members').some(function (m) { return m.photo === id; })) throw new Error('Not found');
  const b = DriveApp.getFileById(id).getBlob();
  return { dataUrl: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()) };
}


/* ---------------- calls: short-lived signals between two phones ----------------
   A call is set up by passing small messages (ring, accept, offer, answer, ice, end) through the cache.
   They live for 90 seconds, are never written to the sheet, and the sound and picture go straight between the phones. */
function sigLocked(fn) { const lock = LockService.getScriptLock(); lock.waitLock(5000); try { return fn(); } finally { try { lock.releaseLock(); } catch (x) {} } }
function takeSigs(id) {
  const c = CacheService.getScriptCache(), k = 'sg:' + id;
  if (c.get(k) === null) return [];
  return sigLocked(function () { const v = c.get(k); c.remove(k); try { return v ? JSON.parse(v) : []; } catch (e) { return []; } });
}
function sigSend(me, q) {
  const to = String(q.to || ''), peer = findBy('Members', to);
  if (!peer || !peer.pinHash || to === me.id) throw new Error('Cannot call this person');
  const msgs = (Array.isArray(q.msgs) ? q.msgs : []).slice(0, 12).map(function (m) { return { t: String(m.t || '').slice(0, 12), id: String(m.id || '').slice(0, 40), video: m.video ? 1 : 0, sdp: m.sdp ? String(m.sdp).slice(0, 12000) : undefined, c: m.c ? String(m.c).slice(0, 1500) : undefined, from: me.id, ts: Date.now() }; });
  if (!msgs.length) return {};
  const c = CacheService.getScriptCache(), k = 'sg:' + to;
  sigLocked(function () { let cur = []; try { cur = JSON.parse(c.get(k) || '[]'); } catch (e) {} cur = cur.concat(msgs).slice(-60); c.put(k, JSON.stringify(cur), 90); c.put('msgTs', String(Date.now()), 21600); });
  const ring = msgs.filter(function (m) { return m.t === 'ring'; })[0];
  if (ring) push([to], shortName(me) + ' is calling', ring.video ? 'Video call on Kinship' : 'Voice call on Kinship', 'chat:' + me.id, { ttl: 60, priority: 10 });
  return {};
}
function sigPoll(me, q) {
  const until = Date.now() + Math.min(10000, Math.max(0, Number(q.wait) || 0)), c = CacheService.getScriptCache();
  for (;;) { const s = takeSigs(me.id); if (s.length || Date.now() >= until) return { sigs: s }; Utilities.sleep(300); }
}

/* ---------------- notifications (OneSignal) ---------------- */
/** title = the first line people see (always specific, never the app name); go = what the notification opens. */
function push(ids, title, body, go, extra) {
  try {
    const app = setting('oneSignalAppId'), key = setting('oneSignalKey');
    ids = (ids || []).filter(function (x, i, a) { return x && a.indexOf(x) === i; });
    if (!app || !key || !ids.length) return;
    const payload = { app_id: app, target_channel: 'push', include_aliases: { external_id: ids }, headings: { en: String(title).slice(0, 64) }, contents: { en: String(body || '').slice(0, 170) || ' ' }, data: { go: go || 'home' } };
    if (extra) Object.keys(extra).forEach(function (k) { payload[k] = extra[k]; });
    const url = setting('appUrl'); if (url) payload.url = url.split('?')[0] + '?go=' + encodeURIComponent(go || 'home');
    const res = UrlFetchApp.fetch('https://api.onesignal.com/notifications', { method: 'post', contentType: 'application/json', headers: { Authorization: 'Key ' + key }, payload: JSON.stringify(payload), muteHttpExceptions: true });
    if (res.getResponseCode() >= 300) console.log('OneSignal: ' + res.getContentText());
  } catch (e) { console.log('push failed: ' + e); }
}
// Family members tagged in a text with @Full Name (only real members, never yourself). Stored as a comma list in the tags column.
function mentionIds(text, selfId) {
  const t = String(text || '').toLowerCase(); if (t.indexOf('@') < 0) return '';
  return readAll('Members').filter(function (m) { return m.id !== selfId && m.name && t.indexOf('@' + String(m.name).toLowerCase()) >= 0; }).map(function (m) { return m.id; }).join(',');
}
function everyoneBut(id) { return readAll('Members').filter(function (m) { return m.pinHash && m.id !== id; }).map(function (m) { return m.id; }); }
function shortName(m) { return m.nick || String(m.name || '').split(' ')[0]; }
function niceDate(d) { const p = String(d).split('-'); return p.length === 3 ? Number(p[2]) + ' ' + ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][Number(p[1]) - 1] : String(d); }
function announce(me, table, r, old) {
  const who = shortName(me);
  const tagged = String(r.tags || '').split(',').filter(function (x) { return x && x !== me.id; }), notTagged = function (ids) { return ids.filter(function (x) { return tagged.indexOf(x) < 0; }); };
  if (tagged.length && !old && (table === 'Posts' || table === 'Comments' || (table === 'Messages' && r.to === 'all')))
    push(tagged, who + ' mentioned you ' + (table === 'Posts' ? 'in a post' : table === 'Comments' ? 'in a comment' : 'in the family chat'), r.text, table === 'Messages' ? 'chat:all' : 'post:' + (table === 'Posts' ? r.id : r.postId));
  if (table === 'Posts' && !old) push(notTagged(everyoneBut(me.id)), who + (r.photo ? (String(r.photo).split(',').length > 1 ? ' shared ' + String(r.photo).split(',').length + ' photos' : ' shared a photo') : ' shared an update'), r.text || 'Tap to see it.', 'post:' + r.id);
  else if ((table === 'Comments' || table === 'Reactions') && !old) {
    const p = findBy('Posts', r.postId);
    if (p && p.by !== me.id && tagged.indexOf(p.by) < 0) push([p.by], who + (table === 'Comments' ? ' commented on your post' : ' loved your post'), table === 'Comments' ? r.text : (p.text || 'Your photo'), 'post:' + p.id);
  }
  else if (table === 'Events' && !old) {
    const inv = String(r.invited || '').split(',').filter(function (x) { return x && x !== me.id; });
    if (inv.length) push(inv, who + ' invited you: ' + r.title, niceDate(r.date) + ' · tap to answer', 'cal');
    const rest = everyoneBut(me.id).filter(function (x) { return inv.indexOf(x) < 0; });
    if (rest.length) push(rest, 'New family date: ' + r.title, niceDate(r.date) + ', added by ' + who, 'cal');
  }
  else if (table === 'Events' && old) {
    const was = String(old.invited || '').split(','), fresh = String(r.invited || '').split(',').filter(function (x) { return x && x !== me.id && was.indexOf(x) < 0; });
    if (fresh.length) push(fresh, who + ' invited you: ' + r.title, niceDate(r.date) + ' · tap to answer', 'cal');
  }
  else if (table === 'Messages' && !old) push(r.to === 'all' ? notTagged(everyoneBut(me.id)) : [r.to], r.to === 'all' ? who + ' wrote in the family chat' : who + ' sent you a message', r.text, r.to === 'all' ? 'chat:all' : 'chat:' + me.id);
  else if (table === 'Wishes') {
    const k = findBy('Members', r.memberId) || {};
    if (!old && me.id === r.memberId) push([k.fatherId, k.motherId], shortName(k) + ' added a wish: ' + r.title, 'Set how many stars it needs.', 'kids:wishes:' + r.memberId);
    else if (r.cost && (!old || old.cost !== r.cost) && r.memberId !== me.id) push([r.memberId], 'Your wish needs ' + r.cost + ' stars ⭐', r.title, 'kids:wishes:' + r.memberId);
  }
  else if (table === 'Redemptions') {
    const kid = findBy('Members', r.memberId) || {}, seg = String(r.rewardId).indexOf('wish:') === 0 ? 'wishes' : 'rewards';
    if (!old && r.status === 'requested') push([kid.fatherId, kid.motherId], shortName(kid) + ' is asking for: ' + r.title, r.cost + ' stars. Tap to approve.', 'kids:' + seg + ':' + r.memberId);
    else if (old && old.status !== r.status && r.memberId !== me.id) push([r.memberId], r.status === 'approved' ? 'Yes! You got: ' + r.title + ' 🎉' : 'Not this time: ' + r.title, r.status === 'approved' ? who + ' approved it.' : 'Your stars are back.', 'kids:' + seg + ':' + r.memberId);
  }
}
/** Runs every morning (trigger created by setup): birthdays and family dates happening today. */
function dailyDigest() {
  const tz = Session.getScriptTimeZone(), md = Utilities.formatDate(new Date(), tz, 'MM-dd'), ymd = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), yr = Number(ymd.slice(0, 4));
  const tz0 = tz, ms = readAll('Members'), seen = {};
  ms.filter(function (m) { return m.birth && !m.passed && m.birth.slice(5) === md; })
    .forEach(function (m) { push(everyoneBut(m.id), "It's " + shortName(m) + "'s birthday today", 'Tap to send your wishes.', 'home'); push([m.id], 'Happy birthday, ' + shortName(m) + '!', 'With love from the whole family.', 'home'); });
  ms.forEach(function (m) {
    const sp = ms.filter(function (x) { return x.id === m.spouseId || x.spouseId === m.id; })[0], an = m.anniv || (sp && sp.anniv);
    if (!sp || !an || an.slice(5) !== md || m.passed || sp.passed) return;
    const key = [m.id, sp.id].sort().join(); if (seen[key]) return; seen[key] = 1;
    const yrs = yr - Number(an.slice(0, 4));
    push(everyoneBut(''), "It's " + shortName(m) + ' & ' + shortName(sp) + "'s anniversary today", yrs > 0 ? yrs + (yrs === 1 ? ' year' : ' years') + ' together.' : 'Send your wishes.', 'home');
  });
  readAll('Events').filter(function (e) { return e.date && (e.yearly === '1' ? e.date.slice(5) === md : e.date === ymd); })
    .forEach(function (e) {
      if (e.kind === 'flight' && e.dep && e.arr) {
        if (Date.parse(e.dep) < Date.now()) return;   // already in the air: a notice now would be late
        const f = function (iso, tz) { try { return Utilities.formatDate(new Date(iso), tz || tz0, 'h:mm a'); } catch (x) { return ''; } };
        push(everyoneBut(''), 'Flight today: ' + e.title, f(e.dep, e.fromTz) + ' → ' + f(e.arr, e.toTz) + (e.flightNo ? ' · ' + e.flightNo : ''), 'cal');
      } else push(everyoneBut(''), 'Today: ' + e.title, e.note || 'Tap to see the details.', 'cal');
    });
  ensureTriggers();
  try { const lock = LockService.getScriptLock(); lock.waitLock(20000); try { runCleanup(); } finally { lock.releaseLock(); } } catch (e) {}
}

/* ---------------- relatives: link two people already in the family ---------------- */
function linkPlan(ms, aId, rel, bId) {
  const A = ms.filter(function (m) { return m.id === aId; })[0], B = ms.filter(function (m) { return m.id === bId; })[0];
  if (!A || !B || aId === bId) throw new Error('Choose a person');
  const slot = function (p) { return p.gender === 'female' ? 'motherId' : 'fatherId'; };
  if (rel === 'father' || rel === 'mother') {
    if (B.fatherId === aId || B.motherId === aId) throw new Error('That would make them each other\'s parent');
    return [{ id: aId, field: rel + 'Id', value: bId }];
  }
  if (rel === 'spouse') return [{ id: aId, field: 'spouseId', value: bId }];
  if (rel === 'child') {
    if (A.fatherId === bId || A.motherId === bId) throw new Error('That would make them each other\'s parent');
    if (!A.gender) throw new Error('Set male or female in the profile first, so the tree knows father or mother');
    const out = [{ id: bId, field: slot(A), value: aId }];
    const sp = ms.filter(function (m) { return m.id === A.spouseId || m.spouseId === aId; })[0];
    if (sp && sp.gender && slot(sp) !== slot(A) && !B[slot(sp)]) out.push({ id: bId, field: slot(sp), value: sp.id });
    return out;
  }
  if (rel === 'sibling') {
    const out = [];
    ['fatherId', 'motherId'].forEach(function (f) { if (A[f] && !B[f]) out.push({ id: bId, field: f, value: A[f] }); else if (B[f] && !A[f]) out.push({ id: aId, field: f, value: B[f] }); });
    if (out.length) return out;
    if ((A.fatherId && A.fatherId === B.fatherId) || (A.motherId && A.motherId === B.motherId)) throw new Error('They are already linked as brother or sister');
    if (!A.fatherId && !A.motherId && !B.fatherId && !B.motherId) throw new Error('Add a father or mother to one of them first. Brothers and sisters are linked through a shared parent.');
    throw new Error('They have different parents in the tree. Check the parents first.');
  }
  throw new Error('Unknown relationship');
}
function link(me, q) {
  const aId = q.a || me.id, ms = readAll('Members');
  if (!(aId === me.id || me.role === 'admin' || isParent(me, aId))) throw new Error('You can add relatives for yourself or your children');
  const plan = linkPlan(ms, aId, String(q.rel), String(q.b));
  plan.forEach(function (c) {
    const t = ms.filter(function (m) { return m.id === c.id; })[0];
    const mayEdit = t.id === me.id || me.role === 'admin' || t.fatherId === me.id || t.motherId === me.id;
    // Someone else's profile may only gain a missing parent, and only if they are a child or not signed up yet.
    const mayFill = me.role !== 'kid' && (c.field === 'fatherId' || c.field === 'motherId') && !t[c.field] && (t.role === 'kid' || !t.pinHash);
    if (!mayEdit && !mayFill) throw new Error('Ask ' + shortName(t) + ' to add this from their own profile, or ask the family admin');
  });
  plan.forEach(function (c) { const t = findBy('Members', c.id); t[c.field] = c.value; writeRow('Members', t, t._row); });
  if (q.b !== me.id) push([String(q.b)], shortName(me) + ' added you as family in the tree', 'Tap to see how everyone is connected.', 'tree');
  return {};
}

/* ---------------- admin settings ---------------- */
function saveSettings(me, q) {
  if (me.role !== 'admin') throw new Error('Only an admin can change settings');
  if (q.appName) setSetting('appName', String(q.appName).slice(0, 30));
  if (q.celebrations !== undefined) setSetting('celebrations', String(q.celebrations).replace(/[^a-z,]/g, '').slice(0, 400));
  if (q.familyCode) { const c = String(q.familyCode).trim(); if (c.length < 6) throw new Error('Family code needs at least 6 characters'); setSetting('familyCode', c); }
  if (q.sessionHours !== undefined) { const h = Math.max(0, Math.min(8760, Math.round(Number(q.sessionHours) || 0))); setSetting('sessionHours', String(h)); CacheService.getScriptCache().put('sessionHours', String(h), 600); }
  return {};
}


/* The evening reminder: people on a flight in the next 24 hours get the times, so an early-morning flight is not missed. */
function flightEve() {
  const now = Date.now();
  readAll('Events').filter(function (e) { const t = Date.parse(e.dep || ''); return e.kind === 'flight' && e.with && t > now + 36e5 && t <= now + 864e5; })
    .forEach(function (e) {
      const f = function (iso, tz) { try { return Utilities.formatDate(new Date(iso), tz || Session.getScriptTimeZone(), 'h:mm a'); } catch (x) { return ''; } };
      push(String(e.with).split(','), 'Flight coming up: ' + e.title, f(e.dep, e.fromTz) + ' → ' + f(e.arr, e.toTz) + (e.flightNo ? ' · ' + e.flightNo : '') + '. Check your documents and bags.', 'cal');
    });
}
function ensureTriggers() {
  const have = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  if (have.indexOf('flightEve') < 0) ScriptApp.newTrigger('flightEve').timeBased().everyDays(1).atHour(18).create();
  if (have.indexOf('dailyDigest') < 0) ScriptApp.newTrigger('dailyDigest').timeBased().everyDays(1).atHour(8).create();
}
