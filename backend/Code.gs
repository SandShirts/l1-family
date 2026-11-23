/* L1 Family App — backend (Google Apps Script, runs as enrique@latoisonlaw.com) */
var APP_URL = 'https://sandshirts.github.io/l1-family/';           // set after hosting
var TZ = 'America/New_York';
var SHEETS = {
  Members:   ['id','name','color','key','topic','created'],
  Messages:  ['id','ts','from','to','text'],
  Menu:      ['date','meal','cook'],
  Grocery:   ['id','text','done','by','ts'],
  Events:    ['id','title','date','start','end','allDay','who','notes','by','gcalId','ts'],
  Locations: ['member','lat','lon','acc','batt','ts','src'],
  Photo:     ['chunk']
};

/* ---------- one-time setup (run from the editor) ---------- */
function setup() {
  var s = book_();
  var m = rows_('Members');
  if (!m.length) {
    [['Enrique','#1f4fbf'],['Stephanie','#c2185b'],['Cole','#2e7d32'],['Ethan','#e65100']]
      .forEach(function (p) { addMember_(p[0], p[1]); });
  }
  rows_('Members').forEach(function (x) { Logger.log(x.name + ' | ' + appUrl_() + 'setup.html?k=' + x.key); });
  CalendarApp.getDefaultCalendar().getName();   // forces calendar permission
  UrlFetchApp.fetch('https://ntfy.sh/v1/health', {muteHttpExceptions: true}); // forces fetch permission
  Logger.log('Data sheet: ' + s.getUrl());
}
function setAppUrl(u){ PropertiesService.getScriptProperties().setProperty('APP_URL', u); }
function appUrl_(){ return PropertiesService.getScriptProperties().getProperty('APP_URL') || APP_URL; }

/* ---------- storage ---------- */
function book_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('SSID'), s;
  if (id) { try { s = SpreadsheetApp.openById(id); } catch (e) { s = null; } }
  if (!s) {
    s = SpreadsheetApp.create('L1 Family App Data');
    p.setProperty('SSID', s.getId());
    Object.keys(SHEETS).forEach(function (n, i) {
      var sh = i === 0 ? s.getSheets()[0].setName(n) : s.insertSheet(n);
      sh.getRange(1, 1, 1, SHEETS[n].length).setValues([SHEETS[n]]).setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.getRange('A:Z').setNumberFormat('@');
    });
  }
  return s;
}
var _book;
function sh_(n) {
  _book = _book || book_();
  var sh = _book.getSheetByName(n);
  if (!sh) { sh = _book.insertSheet(n); sh.getRange(1, 1, 1, SHEETS[n].length).setValues([SHEETS[n]]).setFontWeight('bold'); sh.setFrozenRows(1); sh.getRange('A:Z').setNumberFormat('@'); }
  return sh;
}
var _rc = {};
function rows_(n) {
  if (_rc[n]) return _rc[n];
  var c = CacheService.getScriptCache(), hit = c.get('r_' + n);
  if (hit) return (_rc[n] = JSON.parse(hit));
  var v = sh_(n).getDataRange().getValues(), h = v.shift();
  var out = v.filter(function (r) { return r.join('') !== ''; }).map(function (r) {
    var o = {}; h.forEach(function (k, i) { o[k] = r[i] === '' ? '' : String(r[i]); }); return o;
  });
  var js = JSON.stringify(out); if (js.length < 95000) c.put('r_' + n, js, 21600);
  return (_rc[n] = out);
}
function dirty_(n) { delete _rc[n]; CacheService.getScriptCache().remove('r_' + n); }
function append_(n, o) { dirty_(n); sh_(n).appendRow(SHEETS[n].map(function (k) { return o[k] == null ? '' : String(o[k]); })); }
function findRow_(n, key, val) {
  var col = SHEETS[n].indexOf(key), v = sh_(n).getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][col]) === String(val)) return i + 1;
  return -1;
}
function update_(n, key, val, o) {
  dirty_(n);
  var r = findRow_(n, key, val);
  if (r < 0) { append_(n, o); return; }
  sh_(n).getRange(r, 1, 1, SHEETS[n].length).setValues([SHEETS[n].map(function (k) { return o[k] == null ? '' : String(o[k]); })]);
}
function remove_(n, key, val) { dirty_(n); var r = findRow_(n, key, val); if (r > 0) sh_(n).deleteRow(r); }
function uid_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function token_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
function now_() { return String(Date.now()); }

function addMember_(name, color) {
  var palette = ['#1f4fbf','#c2185b','#2e7d32','#e65100','#6a1b9a','#00838f','#ad1457','#5d4037'];
  var m = {id: uid_(), name: name, color: color || palette[rows_('Members').length % palette.length],
           key: token_(), topic: 'l1fam-' + token_().slice(0, 20), created: now_()};
  append_('Members', m);
  return m;
}

/* ---------- web endpoints ---------- */
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function who_(k) { if (!k) return null; var m = rows_('Members').filter(function (x) { return x.key === k; }); return m[0] || null; }

function doGet(e) {
  var p = e.parameter || {};
  if (p.action === 'ping') return out_({ok: true});
  var me = who_(p.k);
  if (!me) return out_({error: 'badkey'});
  if (p.action === 'photo') return out_({data: photoGet_()});
  if (p.action === 'setup') return out_({ok: true, me: pub_(me, true), api: ScriptApp.getService().getUrl(), app: appUrl_()});
  return out_(state_(me));
}

function doPost(e) {
  var p = e.parameter || {};
  try {
    if (p.ot === '1') {                         // OwnTracks location report
      var m = who_(p.k);
      if (m) {
        var b = JSON.parse(e.postData.contents || '{}');
        if (b._type === 'location' && b.lat != null) saveLoc_(m.id, b.lat, b.lon, b.acc, b.batt, b.tst ? b.tst * 1000 : Date.now(), 'tracker');
      }
      return out_([]);
    }
    var body = JSON.parse(e.postData.contents || '{}');
    var me = who_(body.k);
    if (!me) return out_({error: 'badkey'});
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    var res;
    try { res = act_(me, body) || {}; } finally { lock.releaseLock(); }
    res.state = state_(me);
    return out_(res);
  } catch (err) {
    return out_({error: String(err && err.message || err)});
  }
}

function pub_(m, mine) {
  var o = {id: m.id, name: m.name, color: m.color};
  if (mine) { o.key = m.key; o.topic = m.topic; }
  return o;
}

function state_(me) {
  var members = rows_('Members');
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var past = Utilities.formatDate(new Date(Date.now() - 45 * 864e5), TZ, 'yyyy-MM-dd');
  var msgs = rows_('Messages').filter(function (x) { return x.to === 'all' || x.to === me.id || x.from === me.id; });
  var locs = {};
  rows_('Locations').forEach(function (l) { locs[l.member] = l; });
  return {
    me: pub_(me, true),
    members: members.map(function (m) { return pub_(m, false); }),
    messages: msgs.slice(-400),
    menu: rows_('Menu').filter(function (x) { return x.date >= past; }),
    grocery: rows_('Grocery'),
    events: rows_('Events').filter(function (x) { return x.date >= past; }),
    locations: locs,
    today: today,
    photoVer: PropertiesService.getScriptProperties().getProperty('PHOTOVER') || '',
    now: Date.now()
  };
}

function saveLoc_(member, lat, lon, acc, batt, ts, src) {
  update_('Locations', 'member', member, {member: member, lat: lat, lon: lon, acc: acc || '', batt: batt == null ? '' : batt, ts: ts, src: src});
}

/* ---------- actions ---------- */
function act_(me, b) {
  var members = rows_('Members');
  var others = members.filter(function (m) { return m.id !== me.id; });
  switch (b.action) {
    case 'send':
      var text = String(b.text || '').slice(0, 4000).trim(); if (!text) return {};
      append_('Messages', {id: uid_(), ts: now_(), from: me.id, to: b.to || 'all', text: text});
      var rec = b.to && b.to !== 'all' ? others.filter(function (m) { return m.id === b.to; }) : others;
      push_(rec, me.name + (b.to && b.to !== 'all' ? '' : ' (Family)'), text, 'speech_balloon');
      return {ok: true};
    case 'menuSet':
      update_('Menu', 'date', b.date, {date: b.date, meal: b.meal || '', cook: b.cook || ''});
      return {ok: true};
    case 'grocAdd':
      String(b.text || '').split(/\n|,/).map(function (t) { return t.trim(); }).filter(String).forEach(function (t) {
        append_('Grocery', {id: uid_(), text: t, done: '', by: me.id, ts: now_()});
      });
      return {ok: true};
    case 'grocToggle':
      var g = rows_('Grocery').filter(function (x) { return x.id === b.id; })[0];
      if (g) { g.done = g.done ? '' : '1'; update_('Grocery', 'id', g.id, g); }
      return {ok: true};
    case 'grocDelete':
      remove_('Grocery', 'id', b.id); return {ok: true};
    case 'grocClear':
      rows_('Grocery').filter(function (x) { return x.done; }).forEach(function (x) { remove_('Grocery', 'id', x.id); });
      return {ok: true};
    case 'evSave':  return evSave_(me, b, members);
    case 'evDelete':
      var ev = rows_('Events').filter(function (x) { return x.id === b.id; })[0];
      if (ev) {
        if (ev.gcalId) { try { var ce = CalendarApp.getDefaultCalendar().getEventById(ev.gcalId); if (ce) ce.deleteEvent(); } catch (x) {} }
        remove_('Events', 'id', ev.id);
      }
      return {ok: true, msg: 'Deleted from the app and from your calendar.'};
    case 'loc':
      saveLoc_(me.id, b.lat, b.lon, b.acc, '', Date.now(), 'app'); return {ok: true};
    case 'addMember':
      var nm = String(b.name || '').trim(); if (!nm) return {error: 'Name needed'};
      var nmem = addMember_(nm, b.color);
      return {ok: true, newMember: pub_(nmem, true), setupLink: appUrl_() + 'setup.html?k=' + nmem.key};
    case 'memberEdit':
      var mm = members.filter(function (x) { return x.id === b.id; })[0];
      if (mm) { if (b.name) mm.name = String(b.name).trim(); if (b.color) mm.color = b.color; update_('Members', 'id', mm.id, mm); }
      return {ok: true};
    case 'setupLink':
      var t = members.filter(function (x) { return x.id === b.id; })[0];
      return t ? {ok: true, setupLink: appUrl_() + 'setup.html?k=' + t.key, name: t.name} : {error: 'not found'};
    case 'photoSet':
      var data = String(b.data || '');
      if (data.indexOf('data:image/') !== 0 || data.length > 600000) return {error: 'Photo is not valid or too big'};
      var sh = sh_('Photo'); if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
      var chunks = []; for (var i = 0; i < data.length; i += 40000) chunks.push([data.slice(i, i + 40000)]);
      sh.getRange(2, 1, chunks.length, 1).setValues(chunks);
      PropertiesService.getScriptProperties().setProperty('PHOTOVER', now_());
      return {ok: true, msg: 'Family photo updated.'};
    case 'testAlert':
      push_([me], 'L1 Family', 'Alerts are working on this phone.', 'white_check_mark');
      return {ok: true};
  }
  return {error: 'unknown action'};
}

function evSave_(me, b, members) {
  var title = String(b.title || '').trim(); if (!title) return {error: 'Title needed'};
  var e = b.id ? rows_('Events').filter(function (x) { return x.id === b.id; })[0] : null;
  var isNew = !e;
  e = e || {id: uid_(), by: me.id, gcalId: ''};
  e.title = title; e.date = b.date; e.allDay = b.allDay ? '1' : '';
  e.start = b.allDay ? '' : (b.start || '09:00'); e.end = b.allDay ? '' : (b.end || '');
  e.who = (b.who || []).join(','); e.notes = b.notes || ''; e.ts = now_();
  if (!e.allDay && (!e.end || e.end <= e.start)) e.end = addMin_(e.start, 60);

  // Copy to the enrique@latoisonlaw.com calendar (one way: app -> Google)
  var names = members.filter(function (m) { return e.who.split(',').indexOf(m.id) >= 0; }).map(function (m) { return m.name; });
  var gTitle = 'Family-' + title + (names.length && names.length < members.length ? ' (' + names.join(', ') + ')' : '');
  var desc = (e.notes ? e.notes + '\n\n' : '') + 'Added in the L1 family app by ' + me.name;
  var d = e.date.split('-').map(Number);
  var cal = CalendarApp.getDefaultCalendar(), ce = null;
  if (e.gcalId) { try { ce = cal.getEventById(e.gcalId); } catch (x) { ce = null; } }
  if (ce) {
    ce.setTitle(gTitle); ce.setDescription(desc);
    if (e.allDay) ce.setAllDayDate(new Date(d[0], d[1] - 1, d[2]));
    else ce.setTime(dt_(d, e.start), dt_(d, e.end));
  } else {
    ce = e.allDay ? cal.createAllDayEvent(gTitle, new Date(d[0], d[1] - 1, d[2]), {description: desc})
                  : cal.createEvent(gTitle, dt_(d, e.start), dt_(d, e.end), {description: desc});
    e.gcalId = ce.getId();
  }
  update_('Events', 'id', e.id, e);
  if (isNew) {
    var when = Utilities.formatDate(new Date(d[0], d[1] - 1, d[2]), TZ, 'EEE MMM d') + (e.allDay ? '' : ' ' + ampm_(e.start));
    push_(members.filter(function (m) { return m.id !== me.id; }), me.name + ' added to the calendar', title + ' — ' + when, 'calendar');
  }
  return {ok: true, msg: isNew ? 'Saved — also added to the enrique@latoisonlaw.com calendar.' : 'Updated — the calendar copy was updated too.'};
}
function dt_(d, hm) { var t = hm.split(':').map(Number); return new Date(d[0], d[1] - 1, d[2], t[0], t[1]); }
function addMin_(hm, n) { var t = hm.split(':').map(Number), m = Math.min(t[0] * 60 + t[1] + n, 23 * 60 + 59); return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + m % 60).slice(-2); }
function ampm_(hm) { var t = hm.split(':').map(Number); return ((t[0] + 11) % 12 + 1) + ':' + ('0' + t[1]).slice(-2) + (t[0] < 12 ? ' AM' : ' PM'); }

function photoGet_() {
  var sh = sh_('Photo'); if (sh.getLastRow() < 2) return '';
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; }).join('');
}

/* ---------- phone alerts through ntfy ---------- */
function push_(list, title, body, tag) {
  var reqs = list.filter(function (m) { return m.topic; }).map(function (m) {
    return {url: 'https://ntfy.sh/' + m.topic, method: 'post', payload: Utilities.newBlob(body).getBytes(), muteHttpExceptions: true,
            headers: {'Title': title, 'Tags': tag || 'house', 'Click': appUrl_()}};
  });
  if (reqs.length) { try { UrlFetchApp.fetchAll(reqs); } catch (e) {} }
}

/* ---------- "tracking went quiet" alert (time trigger every 15 minutes) ---------- */
var WATCH = ['Cole', 'Ethan'];
var PARENTS = ['Enrique', 'Stephanie'];
var QUIET_HOURS = 2;
// checkQuiet(): see live script — alerts parents via ntfy when a watched phone stops reporting for QUIET_HOURS, and again when it comes back.
