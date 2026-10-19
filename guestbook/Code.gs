// Guestbook backend for yoomimi.github.io
// ---------------------------------------
// Runs as a Google Apps Script web app bound to a Google Sheet.
// The sheet (first tab) gets these columns, filled in by this script:
//   A id | B time | C name | D message | E reply | F hidden
//
// GET  …/exec            -> JSON list of the latest visible notes
// POST …/exec  {name, message, website}            -> sign the guestbook
// POST …/exec  {action:"check",  key}              -> is this the owner key?
// POST …/exec  {action:"reply",  key, id, reply}   -> set / clear the owner reply
// POST …/exec  {action:"delete", key, id}          -> hide a note
//
// Set the owner key once: Project Settings > Script Properties > OWNER_KEY.
// Every new note is also mailed to the script owner (the account the web app
// runs as). To send it somewhere else, add a Script Property NOTIFY_EMAIL.

var MAX_NAME = 20;
var MAX_MESSAGE = 100;
var MAX_REPLY = 200;
var LIST_LIMIT = 10;          // notes returned to the page
var BURST_LIMIT = 20;         // notes allowed per BURST_WINDOW across all visitors
var BURST_WINDOW = 10 * 60;   // seconds
var KEY_TRIES = 5;            // wrong owner keys before a lockout
var KEY_LOCK = 60 * 60;       // seconds

function sheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function clean_(s, max) {
  s = String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
  return Array.from(s).slice(0, max).join('');
}

function hasLink_(s) {
  return /https?:\/\/|www\./i.test(s);
}

function rows_() {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 1) return [];
  return sh.getRange(1, 1, last, 6).getValues();
}

// ---------- GET: list ----------
function doGet() {
  var out = [];
  var rows = rows_();
  for (var i = rows.length - 1; i >= 0 && out.length < LIST_LIMIT; i--) {
    var r = rows[i];
    if (!r[0] || r[5]) continue;  // no id or hidden
    out.push({
      id: String(r[0]),
      time: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
      name: String(r[2] || ''),
      message: String(r[3] || ''),
      reply: String(r[4] || '')
    });
  }
  return json_(out);
}

// ---------- POST ----------
function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad' }); }
  if (!body || typeof body !== 'object') return json_({ ok: false, error: 'bad' });

  if (body.action) return owner_(body);
  return sign_(body);
}

function sign_(body) {
  if (body.website) return json_({ ok: true, id: '' });  // honeypot: pretend it worked

  var name = clean_(body.name, MAX_NAME);
  var message = clean_(body.message, MAX_MESSAGE);
  if (!message) return json_({ ok: false, error: 'empty' });
  if (hasLink_(name + ' ' + message)) return json_({ ok: false, error: 'link' });

  // crude flood control shared by all visitors
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get('burst') || 0);
  if (n >= BURST_LIMIT) return json_({ ok: false, error: 'busy' });
  cache.put('burst', String(n + 1), BURST_WINDOW);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return json_({ ok: false, error: 'busy' });
  try {
    var id = Utilities.getUuid();
    sheet_().appendRow([id, new Date(), name, message, '', '']);
    notify_(name, message);
    return json_({ ok: true, id: id });
  } finally {
    lock.releaseLock();
  }
}

// ---------- email notification ----------
function notify_(name, message) {
  try {
    var to = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL')
      || Session.getEffectiveUser().getEmail();
    if (!to) return;
    var who = name || 'Anonymous';
    MailApp.sendEmail({
      to: to,
      subject: '[Guestbook] ' + who + ': ' + message.slice(0, 40) + (message.length > 40 ? '…' : ''),
      body: who + ' wrote:\n\n' + message + '\n\n' +
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') + '\n' +
        'Sheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl() + '\n' +
        'Reply or delete: https://yoomimi.github.io/#admin'
    });
  } catch (err) {
    // a mail failure must never block the note itself, but leave a trace
    // in Apps Script > Executions so it can be diagnosed
    console.error('notify_ failed: ' + err);
  }
}

// ---------- one-time setup ----------
// Run this once from the editor (select "setup" > Run) after pasting new
// code. It touches every service the web app uses, so the permission
// prompt asks for all of them at once (sheet + mail + email address).
// Then redeploy a new version.
function setup() {
  var sh = sheet_();
  Logger.log('Sheet OK: ' + SpreadsheetApp.getActiveSpreadsheet().getName() + ' / ' + sh.getName() + ', rows=' + sh.getLastRow());
  Logger.log('Mail OK: remaining daily quota=' + MailApp.getRemainingDailyQuota());
  Logger.log('Notify to: ' + (PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL') || Session.getEffectiveUser().getEmail()));
  Logger.log('OWNER_KEY set: ' + !!PropertiesService.getScriptProperties().getProperty('OWNER_KEY'));
}

// Sends a test mail so you can confirm delivery.
function testNotify() {
  var to = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL')
    || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({
    to: to,
    subject: '[Guestbook] test mail',
    body: 'Guestbook notifications are working. Remaining daily quota: ' + MailApp.getRemainingDailyQuota()
  });
  Logger.log('Sent test mail to ' + to);
}

// ---------- owner actions ----------
function owner_(body) {
  var cache = CacheService.getScriptCache();
  if (cache.get('keylock')) return json_({ ok: false, error: 'locked' });

  var real = PropertiesService.getScriptProperties().getProperty('OWNER_KEY');
  if (!real || String(body.key || '') !== real) {
    var tries = Number(cache.get('keytries') || 0) + 1;
    cache.put('keytries', String(tries), KEY_LOCK);
    if (tries >= KEY_TRIES) cache.put('keylock', '1', KEY_LOCK);
    return json_({ ok: false, error: 'key' });
  }
  cache.remove('keytries');

  if (body.action === 'check') return json_({ ok: true });

  var sh = sheet_();
  var rows = rows_();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(body.id)) continue;
    var rowNum = i + 1;
    if (body.action === 'reply') {
      sh.getRange(rowNum, 5).setValue(clean_(body.reply, MAX_REPLY));
      return json_({ ok: true });
    }
    if (body.action === 'delete') {
      sh.getRange(rowNum, 6).setValue('x');
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'action' });
  }
  return json_({ ok: false, error: 'missing' });
}
