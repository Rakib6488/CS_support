/**
 * WhatsApp Chat Portal v2 - Apps Script backend
 * Open your Google Sheet > Extensions > Apps Script and paste this file (container-bound script).
 *
 * Script Properties (Project Settings > Script properties):
 *   BRIDGE_URL     = https://your-node-server.com   (no trailing slash)
 *   BRIDGE_KEY     = a long random secret (must match BRIDGE_KEY on the Node server)
 *   MAIL_KEY       = a long random secret (must match MAIL_KEY in alimail/.env)
 *   ADMIN_USER     = admin username (e.g. admin)
 *   ADMIN_PASSWORD = admin password
 *   ADMIN_NAME     = (optional) display name of the admin
 *   ALIMAIL_URL    = (optional) public HTTPS address of Alimail, e.g. https://alimail.onrender.com
 *
 * First time: run setup() once from the editor (creates the sheets and Drive folder, grants permissions).
 * Set the project timezone in Project Settings (e.g. Asia/Dhaka); "today" in the stats depends on it.
 * Deploy: Deploy > New deployment > Web app | Execute as: Me | Who has access: Anyone
 * After changing code: Deploy > Manage deployments > Edit > Version: New version.
 * Admin dashboard: <web app URL>?page=dashboard
 */

const MSG_SHEET = 'Messages';
const CHAT_SHEET = 'Chats';
const AGENT_SHEET = 'Agents';
const STATS_SHEET = 'AgentStats';

const MSG_HEADERS = ['Time', 'ChatId', 'Direction', 'Agent', 'AgentName', 'ChatName', 'IsGroup',
  'Sender', 'Body', 'MsgId', 'MediaType', 'MediaName', 'FileId', 'Ack', 'Quoted'];
const C = { TIME: 0, CHAT: 1, DIR: 2, AGENT: 3, AGENT_NAME: 4, CHAT_NAME: 5, GROUP: 6, SENDER: 7,
  BODY: 8, MSGID: 9, MTYPE: 10, MNAME: 11, FILE: 12, ACK: 13, QUOTED: 14 };
const CHAT_HEADERS = ['ChatId', 'Name', 'LastMessage', 'LastTime', 'Unread', 'IsGroup', 'LastAgent', 'Status', 'AssignedTo', 'AssignedName'];
const QR_SHEET = 'QuickReplies';
const QR_HEADERS = ['Shortcut', 'Text', 'Created'];

const CALL_SHEET = 'Calls';
const CALL_HEADERS = ['CallId', 'Direction', 'ChatId', 'Number', 'Name', 'Type', 'IsGroup', 'StartTime',
  'TakenTime', 'EndTime', 'DurationSec', 'Status', 'Agent', 'AgentName', 'Outcome', 'Notes'];
const CL = { ID: 0, DIR: 1, CHAT: 2, NUM: 3, NAME: 4, TYPE: 5, GROUP: 6, START: 7, TAKEN: 8, END: 9,
  DUR: 10, STATUS: 11, AGENT: 12, AGENT_NAME: 13, OUTCOME: 14, NOTES: 15 };
const RING_MS = 45000; // how long an incoming call is shown as ringing
const AGENT_HEADERS = ['Username', 'Name', 'Role', 'Salt', 'Hash', 'Active', 'Created'];

const MAX_VIEW_MEDIA = 16 * 1024 * 1024;  // max size previewed in the portal (the bridge saves up to 15 MB)
const MEDIA_CHUNK = 2 * 1024 * 1024;      // raw bytes sent per google.script.run call
const MAX_SEND_B64 = 14 * 1024 * 1024;    // ~10 MB file

/* ---------- helpers ---------- */

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function setVer_() { PropertiesService.getScriptProperties().setProperty('VER', String(Date.now())); }

// Only individual numbers (@c.us / @lid) and groups (@g.us). Channels, broadcasts and status are excluded.
function allowedChat_(chatId) {
  return /@(c\.us|lid|g\.us)$/.test(String(chatId || ''));
}

// WhatsApp system/notification placeholders (e2e_notification, notification_template, etc.)
function isPlaceholder_(text) {
  return /^\[(e2e_notification|notification_template|notification|gp2|call_log|protocol|broadcast_notification)\]$/.test(String(text || ''));
}

function textCols_(sh, cols) {
  cols.forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
}

// Plain-text column format so a message like "=..." is not treated as a formula
function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (name === MSG_SHEET) textCols_(sh, [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15]);
    if (name === CHAT_SHEET) textCols_(sh, [1, 2, 3, 6, 7, 8, 9, 10]);
    if (name === QR_SHEET) textCols_(sh, [1, 2]);
    if (name === AGENT_SHEET) textCols_(sh, [1, 2, 3, 4, 5, 6]);
    if (name === CALL_SHEET) textCols_(sh, [1, 2, 3, 4, 5, 6, 7, 12, 13, 14, 15, 16]);
  }
  return sh;
}

function appendRow_(sh, values) {
  const row = sh.getLastRow() + 1;
  if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 500);
  sh.getRange(row, 1, 1, values.length).setValues([values]);
  return row;
}

// Adds the Status column (open/closed) to the Chats sheet if missing and sets the status of existing chats
function ensureChatCols_() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('CHAT_COLS_V3') === '1') return;
  const cs = sheet_(CHAT_SHEET, CHAT_HEADERS);
  if (String(cs.getRange(1, 8).getValue()) !== 'Status') {
    cs.getRange(1, 8).setValue('Status').setFontWeight('bold');
    cs.getRange(1, 8, cs.getMaxRows(), 1).setNumberFormat('@');
    const n = cs.getLastRow();
    if (n >= 2) {
      // Empty LastAgent means the last message came from the customer, so open. Otherwise a reply was sent, so closed.
      const vals = cs.getRange(2, 7, n - 1, 1).getValues();
      cs.getRange(2, 8, n - 1, 1).setValues(vals.map(function (r) { return [r[0] ? 'closed' : 'open']; }));
    }
  }
  if (String(cs.getRange(1, 9).getValue()) !== 'AssignedTo') {
    cs.getRange(1, 9, 1, 2).setValues([['AssignedTo', 'AssignedName']]).setFontWeight('bold');
    cs.getRange(1, 9, cs.getMaxRows(), 2).setNumberFormat('@');
  }
  p.setProperty('CHAT_COLS_V3', '1');
}

function mediaFolder_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('MEDIA_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* recreate */ }
  }
  const f = DriveApp.createFolder('WhatsApp Portal Media');
  p.setProperty('MEDIA_FOLDER_ID', f.getId());
  return f;
}

// Run once from the editor
function setup() {
  sheet_(MSG_SHEET, MSG_HEADERS);
  sheet_(CHAT_SHEET, CHAT_HEADERS);
  sheet_(AGENT_SHEET, AGENT_HEADERS);
  callSheet_();
  sheet_(QR_SHEET, QR_HEADERS);
  ensureChatCols_();
  mediaFolder_();
  Logger.log('Setup done');
}

// Run once from the editor to check that saving to Google Drive works (see the Execution log)
function testMediaSave() {
  const folder = mediaFolder_();
  const f = folder.createFile(Utilities.newBlob('test', 'text/plain', 'portal-test.txt'));
  Logger.log('OK: created ' + f.getName() + ' in the folder "' + folder.getName() + '"');
  f.setTrashed(true);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('WhatsApp portal')
    .addItem('Rebuild agent stats sheet', 'rebuildStatsSheet')
    .addItem('Clean up junk chats (channels, notifications)', 'cleanupJunk')
    .addToUi();
}

/* ---------- Web app entry points ---------- */

// Pages (one web app):
//   <exec URL>               agent inbox (chat portal)
//   <exec URL>?page=admin    admin portal (also works as ?page=dashboard)
//   <exec URL>?page=calls    call portal
// Every page includes the shared stylesheet file "Styles" with <?!= include('Styles') ?>
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function renderPage_(file, title) {
  const template = HtmlService.createTemplateFromFile(file);
  template.token = '';
  return template.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doGet(e) {
  const page = e && e.parameter && e.parameter.page;
  if (page === 'admin' || page === 'dashboard') {
    const template = HtmlService.createTemplateFromFile('Dashboard');
    template.token = e && e.parameter && e.parameter.token ? String(e.parameter.token) : '';
    return template.evaluate().setTitle('Admin portal').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'calls') return renderPage_('Index', 'WhatsApp Support');
  return renderPage_('Index', 'WhatsApp Support');
}

// The Node servers (WhatsApp bridge and Alimail) post all events here
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.type === 'mail_tag' || d.type === 'mail_categories') return tagPost_(d);
    if (String(d.type || '').indexOf('mail_') === 0) return mailPost_(d);
    if (d.key !== prop_('BRIDGE_KEY')) return json_({ ok: false, error: 'unauthorized' });
    if (d.type === 'message') {
      if (allowedChat_(d.chatId) && !isPlaceholder_(d.body)) saveMessage_(d);
    }
    else if (d.type === 'ack') updateAck_(d.id, d.ack);
    else if (d.type === 'revoke') revoke_(d.id);
    else if (d.type === 'status') saveStatus_(d);
    else if (d.type === 'call') saveCall_(d);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ---------- Storage ---------- */

function preview_(d, mtype) {
  const t = String(d.body || '');
  if (!mtype) return t;
  const kind = mtype.indexOf('image/') === 0 ? 'Photo'
    : mtype.indexOf('video/') === 0 ? 'Video'
    : mtype.indexOf('audio/') === 0 ? 'Audio' : 'File';
  return t ? kind + ': ' + t : kind;
}

function saveMessage_(d) {
  const cache = CacheService.getScriptCache();
  if (cache.get('m_' + d.id)) return;
  // Store what the agent typed, without the name line that was added for WhatsApp
  if (d.fromMe && d.agentName) d.body = stripSignature_(String(d.body || ''), d.agentName);

  let fileId = '', mtype = '', mname = '';
  if (d.media && d.media.data) {
    mtype = d.media.mimetype || 'application/octet-stream';
    mname = d.media.filename || 'file';
    let lastErr = null;
    // Try twice: Drive sometimes fails once and works on the next call
    for (let attempt = 0; attempt < 2 && !fileId; attempt++) {
      try {
        const blob = Utilities.newBlob(Utilities.base64Decode(d.media.data), mtype, mname);
        fileId = mediaFolder_().createFile(blob).getId();
      } catch (err) {
        lastErr = err;
        Utilities.sleep(700);
      }
    }
    if (!fileId) {
      // Keep the real reason in the Sheet (MediaName column), in the portal and in the execution log
      mname = 'media save failed: ' + String(lastErr).slice(0, 200);
      console.error('Drive save failed: ' + lastErr);
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (cache.get('m_' + d.id)) return;
    cache.put('m_' + d.id, '1', 21600);

    const time = new Date(d.ts * 1000);
    const dir = d.fromMe ? 'out' : 'in';

    appendRow_(sheet_(MSG_SHEET, MSG_HEADERS), [
      time, d.chatId, dir, d.agent || '', d.agentName || '', d.chatName || '', d.isGroup ? 'yes' : '',
      d.sender || '', d.body || '', d.id, mtype, mname, fileId, d.fromMe ? 1 : '', d.quoted || '',
    ]);

    ensureChatCols_();
    const cs = sheet_(CHAT_SHEET, CHAT_HEADERS);
    const found = cs.getRange('A:A').createTextFinder(d.chatId).matchEntireCell(true).findNext();
    const last = preview_(d, mtype);
    const label = d.fromMe ? (d.agentName || 'Phone') : '';
    // A customer message opens the chat; a reply (agent or phone) closes it
    const status = d.fromMe ? 'closed' : 'open';
    const isGroupChat = /@g\.us$/.test(d.chatId);
    // Group chats are shared with every agent, so they are never assigned.
    // In a 1:1 chat the first agent who replies becomes the assignee.
    const autoUser = !isGroupChat && d.fromMe && d.agent ? d.agent : '';
    const autoName = autoUser ? (d.agentName || d.agent) : '';
    if (found) {
      const r = found.getRow();
      const cur = cs.getRange(r, 2, 1, 9).getValues()[0];
      const unread = Number(cur[3]) || 0;
      cs.getRange(r, 2, 1, 7).setValues([[
        d.chatName || cur[0], last, time, d.fromMe ? 0 : unread + 1, d.isGroup ? 'yes' : '', label, status,
      ]]);
      const curUser = String(cur[7]);
      if (autoUser && !curUser) {
        cs.getRange(r, 9, 1, 2).setValues([[autoUser, autoName]]);
      } else if (!d.fromMe && !isGroupChat && (!curUser || !isOnline_(curUser))) {
        // A customer wrote in and nobody (or an agent who is not logged in) owns the chat: give it to a logged-in agent
        const pick = pickAgentForChat_();
        if (pick) cs.getRange(r, 9, 1, 2).setValues([[pick.username, pick.name]]);
      }
    } else {
      let aUser = autoUser, aName = autoName;
      if (!d.fromMe && !isGroupChat) {
        const pick = pickAgentForChat_();
        if (pick) { aUser = pick.username; aName = pick.name; }
      }
      appendRow_(cs, [d.chatId, d.chatName || d.chatId, last, time, d.fromMe ? 0 : 1, d.isGroup ? 'yes' : '', label, status, aUser, aName]);
    }
    setVer_();
  } finally {
    lock.releaseLock();
  }
}

// Finds the row of a MsgId among the last 500 messages
function findRecentRow_(id) {
  const ms = sheet_(MSG_SHEET, MSG_HEADERS);
  const last = ms.getLastRow();
  if (last < 2) return 0;
  const start = Math.max(2, last - 499);
  const ids = ms.getRange(start, C.MSGID + 1, last - start + 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0] === id) return start + i;
  }
  return 0;
}

function updateAck_(id, ack) {
  const row = findRecentRow_(id);
  if (!row) return;
  const ms = sheet_(MSG_SHEET, MSG_HEADERS);
  const cell = ms.getRange(row, C.ACK + 1);
  const cur = Number(cell.getValue()) || 0;
  if (ack < 0 || ack > cur) {
    cell.setValue(ack);
    setVer_();
  }
}

function revoke_(id) {
  const row = findRecentRow_(id);
  if (!row) return;
  sheet_(MSG_SHEET, MSG_HEADERS).getRange(row, C.BODY + 1).setValue('This message was deleted');
  setVer_();
}

function saveStatus_(d) {
  PropertiesService.getScriptProperties().setProperty(
    'STATUS', JSON.stringify({ status: d.status, qr: d.qr || null, ts: Date.now() })
  );
}

/* ---------- Auth (agents) ---------- */

function hash_(salt, pw) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function agentRows_() {
  const sh = sheet_(AGENT_SHEET, AGENT_HEADERS);
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 7).getValues().map(function (r, i) {
    return {
      row: i + 2,
      username: String(r[0]).toLowerCase(),
      name: String(r[1]),
      role: String(r[2]),
      salt: String(r[3]),
      hash: String(r[4]),
      active: String(r[5]).toLowerCase() !== 'no',
    };
  });
}

function login(username, password) {
  username = String(username || '').trim().toLowerCase();
  password = String(password || '');
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + username;
  if (Number(cache.get(failKey) || 0) >= 5) throw new Error('Too many failed attempts. Try again in 10 minutes.');

  let user = null;
  const adminU = String(prop_('ADMIN_USER') || '').toLowerCase();
  if (adminU && username === adminU && password === prop_('ADMIN_PASSWORD')) {
    user = { username: username, name: prop_('ADMIN_NAME') || 'Admin', role: 'admin' };
  } else {
    const a = agentRows_().filter(function (x) { return x.username === username; })[0];
    if (a && a.active && a.hash === hash_(a.salt, password)) {
      user = { username: a.username, name: a.name, role: a.role === 'admin' ? 'admin' : (a.role === 'caller' ? 'caller' : 'agent') };
    }
  }
  if (!user) {
    cache.put(failKey, String(Number(cache.get(failKey) || 0) + 1), 600);
    throw new Error('Incorrect username or password');
  }
  cache.remove(failKey);
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('tok_' + token, JSON.stringify(user), 21600); // 6 hours
  return { token: token, username: user.username, name: user.name, role: user.role };
}

// area: 'chat' (default) or 'calls'. Call agents (role 'caller') may only use the call portal.
function session_(token, area) {
  const cache = CacheService.getScriptCache();
  const s = cache.get('tok_' + token);
  if (!s) throw new Error('SESSION_EXPIRED');
  const user = JSON.parse(s);
  if (cache.get('revoked_' + user.username)) throw new Error('SESSION_EXPIRED');
  if (user.role === 'caller' && area !== 'calls') throw new Error('This account can only use the call portal');
  return user;
}

function adminOnly_(token) {
  const me = session_(token);
  if (me.role !== 'admin') throw new Error('Only admins can do this');
  return me;
}

function listAgents(token) {
  adminOnly_(token);
  return agentRows_().map(function (a) {
    return { username: a.username, name: a.name, role: a.role === 'admin' ? 'admin' : (a.role === 'caller' ? 'caller' : 'agent'), active: a.active };
  });
}

function saveAgent(token, a) {
  adminOnly_(token);
  const username = String(a.username || '').trim().toLowerCase();
  if (!/^[a-z0-9._-]{3,30}$/.test(username)) throw new Error('Username must be 3-30 characters: a-z, 0-9, . _ -');
  const name = String(a.name || '').trim();
  if (!name) throw new Error('Name is required');
  const role = a.role === 'admin' ? 'admin' : (a.role === 'caller' ? 'caller' : 'agent');
  const active = a.active === false ? 'no' : 'yes';
  const pw = String(a.password || '');
  const sh = sheet_(AGENT_SHEET, AGENT_HEADERS);
  const existing = agentRows_().filter(function (x) { return x.username === username; })[0];

  if (existing) {
    let salt = existing.salt, hash = existing.hash;
    if (pw) {
      if (pw.length < 6) throw new Error('Password must be at least 6 characters');
      salt = Utilities.getUuid();
      hash = hash_(salt, pw);
    }
    sh.getRange(existing.row, 2, 1, 5).setValues([[name, role, salt, hash, active]]);
  } else {
    if (pw.length < 6) throw new Error('Password must be at least 6 characters');
    const salt = Utilities.getUuid();
    appendRow_(sh, [username, name, role, salt, hash_(salt, pw), active, new Date()]);
  }
  const cache = CacheService.getScriptCache();
  if (active === 'no') cache.put('revoked_' + username, '1', 21600);
  else cache.remove('revoked_' + username);
  return true;
}

/* ---------- Portal API (google.script.run) ---------- */

// Lightweight poll: bridge status + data version. Chat data is fetched only when the version changes.
function pulse(token) {
  const me = session_(token);
  touchPresence_(me); // this agent is logged in and active
  const p = PropertiesService.getScriptProperties().getProperties();
  let st = { status: 'offline', qr: null };
  try {
    const s = JSON.parse(p.STATUS || 'null');
    if (s && Date.now() - s.ts < 150000) st = { status: s.status, qr: s.qr };
  } catch (e) { /* ignore */ }
  return { status: st.status, qr: st.qr, ver: p.VER || '0', sign: signMode_(), me: { name: me.name, role: me.role } };
}

function getChats(token) {
  const me = session_(token);
  ensureChatCols_();
  const sh = sheet_(CHAT_SHEET, CHAT_HEADERS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 10).getValues()
    .filter(function (r) { return allowedChat_(r[0]) && !isPlaceholder_(r[2]); })
    .filter(function (r) { return canSeeChat_(me, r[0], r[8]); })
    .map(function (r) {
      return {
        chatId: String(r[0]),
        name: String(r[1]),
        last: String(r[2]),
        time: r[3] instanceof Date ? r[3].getTime() : 0,
        unread: Number(r[4]) || 0,
        isGroup: String(r[5]) === 'yes',
        lastAgent: String(r[6]),
        status: String(r[7]) === 'closed' ? 'closed' : 'open',
        assignee: String(r[8] || ''),
        assigneeName: String(r[9] || ''),
      };
    })
    .sort(function (a, b) { return b.time - a.time; });
}

function getMessages(token, chatId, limit) {
  const me = session_(token);
  assertChatAccess_(me, chatId);
  limit = Math.min(Number(limit) || 100, 1000);
  const sh = sheet_(MSG_SHEET, MSG_HEADERS);
  const last = sh.getLastRow();
  if (last < 2) return { messages: [], hasMore: false };

  const ids = sh.getRange(2, C.CHAT + 1, last - 1, 1).getValues();
  const idx = [];
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === chatId) idx.push(i + 2);
  if (!idx.length) return { messages: [], hasMore: false };

  const picked = idx.slice(-limit);
  const first = picked[0];
  const block = sh.getRange(first, 1, last - first + 1, MSG_HEADERS.length).getValues();
  const out = [];
  for (let i = 0; i < block.length; i++) {
    const r = block[i];
    if (r[C.CHAT] !== chatId) continue;
    out.push({
      id: String(r[C.MSGID]),
      time: r[C.TIME] instanceof Date ? r[C.TIME].getTime() : 0,
      dir: String(r[C.DIR]),
      sender: String(r[C.SENDER]),
      agentName: String(r[C.AGENT_NAME]),
      body: String(r[C.BODY]),
      mtype: String(r[C.MTYPE]),
      mname: String(r[C.MNAME]),
      fileId: String(r[C.FILE]),
      ack: Number(r[C.ACK]) || 0,
      quoted: String(r[C.QUOTED]),
    });
  }
  return { messages: out, hasMore: idx.length > picked.length };
}

function markRead(token, chatId) {
  const me = session_(token);
  assertChatAccess_(me, chatId);
  const cs = sheet_(CHAT_SHEET, CHAT_HEADERS);
  const found = cs.getRange('A:A').createTextFinder(chatId).matchEntireCell(true).findNext();
  if (found) {
    cs.getRange(found.getRow(), 5).setValue(0);
    setVer_();
  }
  return true;
}

// Reads media from Drive in chunks (so large videos work). Only files in our media folder are allowed.
function getMedia(token, fileId, index) {
  session_(token);
  const file = DriveApp.getFileById(fileId);
  const folderId = mediaFolder_().getId();
  let ok = false;
  const parents = file.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === folderId) { ok = true; break; }
  }
  if (!ok) throw new Error('Not allowed');

  const size = file.getSize();
  if (size > MAX_VIEW_MEDIA) {
    return { tooLarge: true, size: size, name: file.getName(), url: file.getUrl() };
  }
  const blob = file.getBlob();
  const bytes = blob.getBytes();
  const total = Math.max(1, Math.ceil(bytes.length / MEDIA_CHUNK));
  const i = Math.max(0, Math.min(Number(index) || 0, total - 1));
  const part = bytes.slice(i * MEDIA_CHUNK, (i + 1) * MEDIA_CHUNK);
  return {
    mime: blob.getContentType(),
    name: file.getName(),
    size: size,
    index: i,
    total: total,
    data: Utilities.base64Encode(part),
  };
}

function bridge_(path, payload) {
  const res = UrlFetchApp.fetch(prop_('BRIDGE_URL') + path, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': prop_('BRIDGE_KEY'), 'ngrok-skip-browser-warning': '1' },
    payload: JSON.stringify(payload || {}),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code !== 200) {
    let msg = 'Send failed (bridge HTTP ' + code + ')';
    try {
      msg = JSON.parse(text).error || msg;
    } catch (e) {
      msg += ': ' + text.replace(/\s+/g, ' ').slice(0, 150);
    }
    throw new Error(msg);
  }
  try {
    return JSON.parse(text);
  } catch (e) {
    throw new Error('Unexpected response from the bridge: ' + text.replace(/\s+/g, ' ').slice(0, 150));
  }
}

/* ---------- Agent name on replies ---------- */
// All agents send from the same WhatsApp account, so people in a WhatsApp group cannot tell who replied.
// To fix that, the agent's name is added as the first line of the message (in bold): *Rakib:*
// Script property SIGN_MESSAGES: 'groups' (default) = group chats only, 'all' = every chat, 'off' = never.
function signMode_() {
  const m = String(prop_('SIGN_MESSAGES') || 'groups').toLowerCase();
  return (m === 'all' || m === 'off') ? m : 'groups';
}

function signTag_(name) {
  return '*' + String(name || '').replace(/[*_~`\r\n]/g, '').trim() + ':*';
}

function signText_(me, chatId, text) {
  const mode = signMode_();
  const isGroup = /@g\.us$/.test(chatId);
  if (mode === 'off' || (mode === 'groups' && !isGroup)) return text;
  const tag = signTag_(me.name);
  return text ? tag + '\n' + text : tag;
}

// The message we store keeps only what the agent typed (the agent name is shown separately in the portal)
function stripSignature_(body, agentName) {
  if (!agentName) return body;
  const tag = signTag_(agentName);
  if (body === tag) return '';
  if (body.indexOf(tag + '\n') === 0) return body.slice(tag.length + 1);
  return body;
}

// Replies carry service tags (required). The tags go to WpTags. The message is sent only if the tags are valid.
function sendMessage(token, chatId, text, quotedId, tags) {
  const me = session_(token);
  assertChatAccess_(me, chatId);
  text = String(text || '').trim();
  if (!text) throw new Error('Message is empty');
  const cleanTags = validateTags_(tags);
  const res = bridge_('/send', { to: chatId, text: signText_(me, chatId, text), quotedId: quotedId || '', agent: me.username, agentName: me.name });

  if (res && res.id) {
    try {
      saveTags_('WhatsApp', chatId, chatId.split('@')[0], res.id, me.username, me.name, cleanTags);
    } catch (e) {
      console.error('Tag save failed: ' + e);
    }
  }

  // Save the sent message to the Sheet right now so it shows up without waiting for the bridge to
  // report it back. The bridge reports the same message later; saveMessage_ ignores that duplicate.
  try {
    if (res && res.id) {
      let quoted = '';
      if (quotedId) {
        try {
          const qrow = findRecentRow_(quotedId);
          if (qrow) quoted = String(sheet_(MSG_SHEET, MSG_HEADERS).getRange(qrow, C.BODY + 1).getValue());
        } catch (e) { /* the quote text is optional */ }
      }
      saveMessage_({
        id: res.id, chatId: chatId, chatName: '', isGroup: /@g\.us$/.test(chatId), sender: '',
        fromMe: true, body: text, ts: Math.floor(Date.now() / 1000), quoted: quoted, media: null,
        agent: me.username, agentName: me.name,
      });
    }
  } catch (e) {
    console.error('Early save failed: ' + e);
  }
  return true;
}

function sendMedia(token, chatId, media, caption, quotedId, tags) {
  const me = session_(token);
  assertChatAccess_(me, chatId);
  if (!media || !media.data) throw new Error('No file attached');
  if (media.data.length > MAX_SEND_B64) throw new Error('File is larger than 10 MB');
  const cleanTags = validateTags_(tags);
  const res = bridge_('/send', {
    to: chatId,
    text: signText_(me, chatId, String(caption || '').trim()),
    quotedId: quotedId || '',
    media: { mimetype: String(media.mimetype || 'application/octet-stream'), filename: String(media.filename || 'file'), data: media.data },
    agent: me.username,
    agentName: me.name,
  });
  if (res && res.id) {
    try {
      saveTags_('WhatsApp', chatId, chatId.split('@')[0], res.id, me.username, me.name, cleanTags);
    } catch (e) {
      console.error('Tag save failed: ' + e);
    }
  }
  return true;
}

function logoutWhatsApp(token) {
  adminOnly_(token);
  bridge_('/logout', {});
  return true;
}

/* ---------- Agent stats ---------- */

function rangeFor_(key) {
  const to = Date.now();
  let from = 0;
  if (key === 'today') {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    from = d.getTime();
  } else if (key === '7d') from = to - 7 * 86400000;
  else if (key === '30d') from = to - 30 * 86400000;
  return [from, to];
}

function dayKey_(d) {
  const m = d.getMonth() + 1, day = d.getDate();
  return d.getFullYear() + '-' + (m < 10 ? '0' : '') + m + '-' + (day < 10 ? '0' : '') + day;
}

function computeStats_(fromMs, toMs) {
  const sh = sheet_(MSG_SHEET, MSG_HEADERS);
  const last = sh.getLastRow();
  const totals = {
    incoming: 0, answered: 0, unanswered: 0, customers: 0,
    servedChats: 0, pendingChats: 0, outgoingAgents: 0, outgoingPhone: 0,
  };
  if (last < 2) return { agents: [], totals: totals, daily: [] };

  const rows = sh.getRange(2, 1, last - 1, 5).getValues(); // Time, ChatId, Direction, Agent, AgentName
  rows.sort(function (a, b) { return a[0] - b[0]; });

  const agents = {}, daily = {};
  const pending = {};            // chat -> { first: time of the first unanswered incoming message, n: unanswered incoming count within the range }
  const custSet = {}, servedSet = {};

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!(r[0] instanceof Date) || !allowedChat_(r[1])) continue;
    const t = r[0].getTime();
    const chat = r[1];
    const inRange = t >= fromMs && t <= toMs;

    let dayO = null;
    if (inRange) {
      const dk = dayKey_(r[0]);
      dayO = daily[dk] || (daily[dk] = { day: dk, incoming: 0, replies: 0, cust: {}, replied: {} });
    }

    if (r[2] === 'in') {
      const p = pending[chat] || (pending[chat] = { first: t, n: 0 });
      if (inRange) {
        p.n++;
        totals.incoming++;
        custSet[chat] = 1;
        dayO.incoming++;
        dayO.cust[chat] = 1;
      }
    } else {
      const p = pending[chat];
      if (inRange) {
        const key = r[3] || '(phone)';
        const a = agents[key] || (agents[key] = {
          username: r[3] || '', name: r[3] ? (r[4] || r[3]) : 'Phone',
          replies: 0, chatSet: {}, rtSum: 0, rtN: 0, lastTs: 0,
        });
        a.replies++;
        a.chatSet[chat] = 1;
        a.lastTs = Math.max(a.lastTs, t);
        if (r[3]) { totals.outgoingAgents++; dayO.replies++; } else totals.outgoingPhone++;
        dayO.replied[chat] = 1;
        if (p) { a.rtSum += t - p.first; a.rtN++; }
      }
      if (p) {
        if (p.n > 0) { totals.answered += p.n; servedSet[chat] = 1; }
        delete pending[chat];
      }
    }
  }

  Object.keys(pending).forEach(function (c) {
    if (pending[c].n > 0) { totals.unanswered += pending[c].n; totals.pendingChats++; }
  });
  totals.customers = Object.keys(custSet).length;
  totals.servedChats = Object.keys(servedSet).length;

  const list = Object.keys(agents).map(function (k) {
    const a = agents[k];
    return {
      username: a.username,
      name: a.name,
      replies: a.replies,
      chats: Object.keys(a.chatSet).length,
      avgRespSec: a.rtN ? Math.round(a.rtSum / a.rtN / 1000) : null,
      lastTs: a.lastTs,
    };
  }).sort(function (a, b) { return b.replies - a.replies; });

  const dailyList = Object.keys(daily).sort().reverse().slice(0, 62).map(function (k) {
    const d = daily[k];
    return {
      day: d.day,
      incoming: d.incoming,
      customers: Object.keys(d.cust).length,
      replies: d.replies,
      chatsReplied: Object.keys(d.replied).length,
    };
  });

  return { agents: list, totals: totals, daily: dailyList };
}

function getStats(token, rangeKey) {
  const me = session_(token);
  const cache = CacheService.getScriptCache();
  const ver = prop_('VER') || '0';
  const ck = 'stats2_' + rangeKey + '_' + ver;
  let s;
  const hit = cache.get(ck);
  if (hit) {
    s = JSON.parse(hit);
  } else {
    const r = rangeFor_(rangeKey);
    s = computeStats_(r[0], r[1]);
    try { cache.put(ck, JSON.stringify(s), 120); } catch (e) { /* too big, skip */ }
  }
  if (me.role !== 'admin') {
    s = { agents: s.agents.filter(function (a) { return a.username === me.username; }), totals: null, daily: [] };
  }
  return s;
}

// Sheet menu: WhatsApp portal > Clean up junk chats
// Removes channel (@newsletter), broadcast and notification messages, then rebuilds Messages and Chats.
function cleanupJunk() {
  const ui = SpreadsheetApp.getUi();
  const ans = ui.alert(
    'Channel, broadcast and notification chats will be removed from the Messages and Chats sheets. ' +
    'Make a backup first (File > Make a copy). Continue?',
    ui.ButtonSet.YES_NO
  );
  if (ans !== ui.Button.YES) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ensureChatCols_();
    const ms = sheet_(MSG_SHEET, MSG_HEADERS);
    const cs = sheet_(CHAT_SHEET, CHAT_HEADERS);
    const mLast = ms.getLastRow();
    const cLast = cs.getLastRow();

    const oldChat = {};
    if (cLast >= 2) {
      cs.getRange(2, 1, cLast - 1, 10).getValues().forEach(function (r) {
        oldChat[r[0]] = { unread: Number(r[4]) || 0, status: String(r[7]), assignee: String(r[8]), assigneeName: String(r[9]) };
      });
    }

    const keep = [];
    let removed = 0;
    if (mLast >= 2) {
      const all = ms.getRange(2, 1, mLast - 1, MSG_HEADERS.length).getValues();
      all.forEach(function (r) {
        if (allowedChat_(r[C.CHAT]) && !isPlaceholder_(r[C.BODY])) keep.push(r);
        else removed++;
      });
      ms.getRange(2, 1, mLast - 1, MSG_HEADERS.length).clearContent();
      if (keep.length) ms.getRange(2, 1, keep.length, MSG_HEADERS.length).setValues(keep);
    }

    // Rebuild the Chats sheet from the remaining messages
    const map = {};
    keep.forEach(function (r) {
      const id = r[C.CHAT];
      const prev = map[id];
      map[id] = {
        name: r[C.CHAT_NAME] || (prev && prev.name) || id,
        last: preview_({ body: r[C.BODY] }, r[C.MTYPE]),
        time: r[C.TIME],
        group: r[C.GROUP],
        agent: r[C.DIR] === 'out' ? (r[C.AGENT_NAME] || 'Phone') : '',
      };
    });
    const ids = Object.keys(map);
    if (cLast >= 2) cs.getRange(2, 1, cLast - 1, 10).clearContent();
    if (ids.length) {
      const rows = ids.map(function (id) {
        const m = map[id];
        const o = oldChat[id] || {};
        return [id, m.name, m.last, m.time, o.unread || 0, m.group, m.agent, o.status || (m.agent ? 'closed' : 'open'), o.assignee || '', o.assigneeName || ''];
      });
      if (rows.length + 1 > cs.getMaxRows()) cs.insertRowsAfter(cs.getMaxRows(), rows.length + 1 - cs.getMaxRows() + 100);
      cs.getRange(2, 1, rows.length, 10).setValues(rows);
    }
    setVer_();
    ui.alert('Done. Removed ' + removed + ' junk messages; ' + ids.length + ' chats remain.');
  } finally {
    lock.releaseLock();
  }
}

// Sheet menu: WhatsApp portal > Rebuild agent stats sheet
function rebuildStatsSheet() {
  const ms = sheet_(MSG_SHEET, MSG_HEADERS);
  const last = ms.getLastRow();
  const tz = Session.getScriptTimeZone();
  const map = {};
  if (last >= 2) {
    ms.getRange(2, 1, last - 1, 5).getValues().forEach(function (r) {
      if (!(r[0] instanceof Date) || !allowedChat_(r[1])) return;
      const day = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd');
      const who = r[2] === 'in' ? '(customer messages)' : (r[3] ? (r[4] || r[3]) : 'Phone');
      const k = day + '|' + who;
      const o = map[k] || (map[k] = { day: day, who: who, msgs: 0, chats: {} });
      o.msgs++;
      o.chats[r[1]] = 1;
    });
  }
  const out = Object.keys(map).sort().map(function (k) {
    return [map[k].day, map[k].who, map[k].msgs, Object.keys(map[k].chats).length];
  });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(STATS_SHEET) || ss.insertSheet(STATS_SHEET);
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([['Date', 'Who', 'Messages', 'Chats']]).setFontWeight('bold');
  if (out.length) {
    sh.getRange(2, 1, out.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, out.length, 4).setValues(out);
  }
  sh.setFrozenRows(1);
}

/* ---------- Admin dashboard (page=dashboard) ---------- */

function parseDayParts_(s, endOfDay) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  if (!m) return null;
  return endOfDay
    ? new Date(+m[1], +m[2] - 1, +m[3], 23, 59, 59, 999).getTime()
    : new Date(+m[1], +m[2] - 1, +m[3], 0, 0, 0, 0).getTime();
}

// filter.range: today | 7d | 30d | all | custom (fromDate, toDate = yyyy-MM-dd)
function dashRange_(f) {
  const k = (f && f.range) || '7d';
  if (k === 'custom') {
    const from = parseDayParts_(f.fromDate, false);
    const to = parseDayParts_(f.toDate, true);
    return [from || 0, to || Date.now()];
  }
  return rangeFor_(k);
}

function typeOk_(chatId, type) {
  const isGroup = /@g\.us$/.test(String(chatId));
  if (type === 'group') return isGroup;
  if (type === 'individual') return !isGroup;
  return true;
}

function dayFromKey_(k) {
  const p = k.split('-');
  return new Date(+p[0], +p[1] - 1, +p[2]);
}

// One pass computes everything: totals, agents, days, hours, chats, pending
function computeDash_(filter) {
  filter = filter || {};
  const rng = dashRange_(filter);
  const fromMs = rng[0], toMs = rng[1];
  const type = filter.type || 'all';
  const tz = Session.getScriptTimeZone();
  const now = Date.now();
  const slaReq = Number(filter.sla);
  const slaMin = slaReq > 0 ? Math.min(Math.round(slaReq), 1440) : (Number(prop_('SLA_MINUTES')) || 10);
  const slaMs = slaMin * 60000;

  const totals = {
    incoming: 0, answered: 0, unanswered: 0, customers: 0, servedChats: 0, pendingChats: 0,
    outgoingAgents: 0, outgoingPhone: 0, avgRespSec: null, answerRate: null,
    individualIncoming: 0, groupIncoming: 0,
    slaMin: slaMin, slaN: 0, slaMet: 0, slaRate: null, overdueChats: 0,
  };
  const hourly = [];
  for (let h = 0; h < 24; h++) hourly.push({ hour: h, incoming: 0, agentReplies: 0, phoneReplies: 0 });
  const res = {
    range: { fromMs: fromMs, toMs: toMs, tz: tz },
    totals: totals, agents: [], agentDay: [], daily: [], hourly: hourly, chats: [], pending: [],
  };

  const sh = sheet_(MSG_SHEET, MSG_HEADERS);
  const last = sh.getLastRow();
  if (last < 2) return res;

  const nameMap = {}, ownerMap = {};
  const cs = sheet_(CHAT_SHEET, CHAT_HEADERS);
  const cl = cs.getLastRow();
  if (cl >= 2) {
    cs.getRange(2, 1, cl - 1, 10).getValues().forEach(function (r) {
      nameMap[r[0]] = String(r[1]);
      ownerMap[r[0]] = { u: String(r[8] || ''), n: String(r[9] || ''), closed: String(r[7]) === 'closed' };
    });
  }

  const rows = sh.getRange(2, 1, last - 1, 5).getValues(); // Time, ChatId, Direction, Agent, AgentName
  rows.sort(function (a, b) { return a[0] - b[0]; });

  const agents = {}, daily = {}, agentDay = {}, chats = {}, pending = {};
  const custSet = {}, servedSet = {};
  let rtSum = 0, rtN = 0;

  function dayObj(dk) {
    return daily[dk] || (daily[dk] = { day: dk, incoming: 0, replies: 0, phone: 0, cust: {}, replied: {} });
  }
  function chatObj(id) {
    return chats[id] || (chats[id] = { id: id, incoming: 0, replies: 0, lastTs: 0, lastAgent: '' });
  }

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!(r[0] instanceof Date)) continue;
    const chat = r[1];
    if (!allowedChat_(chat) || !typeOk_(chat, type)) continue;
    const t = r[0].getTime();
    const inRange = t >= fromMs && t <= toMs;

    if (r[2] === 'in') {
      const p = pending[chat] || (pending[chat] = { first: t, n: 0 });
      if (inRange) {
        p.n++;
        totals.incoming++;
        if (/@g\.us$/.test(chat)) totals.groupIncoming++; else totals.individualIncoming++;
        custSet[chat] = 1;
        const dO = dayObj(dayKey_(r[0]));
        dO.incoming++;
        dO.cust[chat] = 1;
        hourly[r[0].getHours()].incoming++;
        const c = chatObj(chat);
        c.incoming++;
        c.lastTs = Math.max(c.lastTs, t);
      }
    } else {
      const p = pending[chat];
      if (inRange) {
        const isAgent = !!r[3];
        const key = isAgent ? r[3] : '(phone)';
        const a = agents[key] || (agents[key] = {
          username: isAgent ? r[3] : '', name: '', replies: 0, chatSet: {}, rtSum: 0, rtN: 0, lastTs: 0, slaMet: 0,
        });
        a.name = isAgent ? (r[4] || r[3]) : 'Phone';
        a.replies++;
        a.chatSet[chat] = 1;
        a.lastTs = Math.max(a.lastTs, t);

        const dk = dayKey_(r[0]);
        const dO = dayObj(dk);
        dO.replied[chat] = 1;
        const hr = r[0].getHours();
        if (isAgent) {
          totals.outgoingAgents++;
          dO.replies++;
          hourly[hr].agentReplies++;
          const ak = dk + '|' + key;
          const ad = agentDay[ak] || (agentDay[ak] = { day: dk, username: key, name: a.name, replies: 0, chatSet: {} });
          ad.replies++;
          ad.chatSet[chat] = 1;
        } else {
          totals.outgoingPhone++;
          dO.phone++;
          hourly[hr].phoneReplies++;
        }
        const c = chatObj(chat);
        c.replies++;
        c.lastTs = Math.max(c.lastTs, t);
        c.lastAgent = a.name;
        if (p) {
          const dt = t - p.first;
          a.rtSum += dt; a.rtN++;
          if (isAgent) { rtSum += dt; rtN++; }
          totals.slaN++;
          if (dt <= slaMs) { a.slaMet++; totals.slaMet++; }
        }
      }
      if (p) {
        if (p.n > 0) { totals.answered += p.n; servedSet[chat] = 1; }
        delete pending[chat];
      }
    }
  }

  // Still waiting for a reply
  const pendingList = [];
  Object.keys(pending).forEach(function (c) {
    const p = pending[c];
    if (p.n > 0) {
      totals.unanswered += p.n;
      totals.pendingChats++;
      const waitMin = Math.round((now - p.first) / 60000);
      if (waitMin > slaMin) totals.overdueChats++;
      pendingList.push({
        id: c, name: nameMap[c] || c.split('@')[0], isGroup: /@g\.us$/.test(c),
        unanswered: p.n, sinceMs: p.first, waitMin: waitMin, overdue: waitMin > slaMin,
        assignee: (ownerMap[c] || {}).u || '', assigneeName: (ownerMap[c] || {}).n || '',
      });
    }
  });
  pendingList.sort(function (a, b) { return b.waitMin - a.waitMin; });

  totals.customers = Object.keys(custSet).length;
  totals.servedChats = Object.keys(servedSet).length;
  totals.avgRespSec = rtN ? Math.round(rtSum / rtN / 1000) : null;
  totals.slaRate = totals.slaN ? Math.round(totals.slaMet * 1000 / totals.slaN) / 10 : null;
  totals.answerRate = totals.incoming ? Math.round(totals.answered * 1000 / totals.incoming) / 10 : null;

  res.agents = Object.keys(agents).map(function (k) {
    const a = agents[k];
    return {
      username: a.username, name: a.name, replies: a.replies,
      chats: Object.keys(a.chatSet).length,
      avgRespSec: a.rtN ? Math.round(a.rtSum / a.rtN / 1000) : null,
      lastTs: a.lastTs,
      slaRate: a.rtN ? Math.round(a.slaMet * 1000 / a.rtN) / 10 : null,
    };
  }).sort(function (a, b) { return b.replies - a.replies; });

  // Day-wise (gaps are filled with zeros for the charts)
  const dkeys = Object.keys(daily).sort();
  if (dkeys.length) {
    const e = dayFromKey_(dkeys[dkeys.length - 1]);
    let s = dayFromKey_(dkeys[0]);
    if (Math.round((e - s) / 86400000) > 399) s = new Date(e.getFullYear(), e.getMonth(), e.getDate() - 399);
    for (let d = new Date(s); d <= e; d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1)) {
      const k = dayKey_(d);
      const o = daily[k];
      res.daily.push({
        day: k,
        incoming: o ? o.incoming : 0,
        customers: o ? Object.keys(o.cust).length : 0,
        replies: o ? o.replies : 0,
        phone: o ? o.phone : 0,
        chatsReplied: o ? Object.keys(o.replied).length : 0,
      });
    }
    const firstDay = res.daily[0].day;
    res.agentDay = Object.keys(agentDay).map(function (k) {
      const a = agentDay[k];
      return { day: a.day, username: a.username, name: a.name, replies: a.replies, chats: Object.keys(a.chatSet).length };
    }).filter(function (a) { return a.day >= firstDay; })
      .sort(function (a, b) { return a.day < b.day ? -1 : a.day > b.day ? 1 : (a.name < b.name ? -1 : 1); });
  }

  res.hourly = hourly;

  const pendSet = {};
  pendingList.forEach(function (p) { pendSet[p.id] = 1; });
  res.chats = Object.keys(chats).map(function (id) {
    const c = chats[id];
    return {
      id: id, name: nameMap[id] || id.split('@')[0], isGroup: /@g\.us$/.test(id),
      incoming: c.incoming, replies: c.replies,
      status: pendSet[id] ? 'Pending' : (c.incoming > 0 ? 'Served' : 'Outbound only'),
      lastAgent: c.lastAgent, lastTs: c.lastTs,
      assignee: (ownerMap[id] || {}).u || '', assigneeName: (ownerMap[id] || {}).n || '',
      chatStatus: (ownerMap[id] || {}).closed ? 'closed' : 'open',
    };
  }).sort(function (a, b) { return b.lastTs - a.lastTs; });
  res.pending = pendingList;
  return res;
}

function getDashboard(token, filter) {
  adminOnly_(token);
  const sla = Number(filter && filter.sla);
  if (sla > 0) PropertiesService.getScriptProperties().setProperty('SLA_MINUTES', String(Math.min(Math.round(sla), 1440)));
  const d = computeDash_(filter);
  d.totalChats = d.chats.length;
  d.totalPending = d.pending.length;
  d.chats = d.chats.slice(0, 150);
  d.pending = d.pending.slice(0, 100);
  return d;
}

function getDashboardUrl(token) {
  adminOnly_(token);
  return ScriptApp.getService().getUrl() + '?page=dashboard';
}

// Links for the two supported portals (admins only)
function getPortalLinks(token) {
  adminOnly_(token);
  const base = ScriptApp.getService().getUrl();
  return { chat: base, dashboard: base + '?page=admin', admin: base + '?page=admin' };
}

/* ---------- CSV export ---------- */

function csvCell_(v) {
  if (v === null || v === undefined) return '';
  let s;
  if (v instanceof Date) {
    s = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } else {
    s = String(v);
    // Prevents Excel formula injection (customer messages are untrusted)
    if (typeof v === 'string' && /^[=+\-@\t\r]/.test(s)) s = "'" + s;
  }
  if (/[",\n\r]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function csvLine_(arr) {
  return arr.map(csvCell_).join(',') + '\n';
}

function fmtTs_(ms) {
  return ms ? Utilities.formatDate(new Date(ms), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '';
}

const MSG_EXPORT_PAGE = 4000;

// kind: daily | agent | agentday | hour | chat | pending | messages
function exportCsv(token, kind, filter, page) {
  adminOnly_(token);
  filter = filter || {};
  if (kind === 'messages') return exportMessages_(filter, Number(page) || 0);

  const d = computeDash_(filter);
  const agent = filter.agent || '';
  let head, rows;

  if (kind === 'daily') {
    head = ['Date', 'Messages received', 'Customers (chats)', 'Agent replies', 'Phone replies', 'Chats replied'];
    rows = d.daily.map(function (x) { return [x.day, x.incoming, x.customers, x.replies, x.phone, x.chatsReplied]; });
  } else if (kind === 'agent') {
    head = ['Agent', 'Username', 'Replies', 'Chats replied', 'Avg first response (sec)', 'Within SLA target (%)', 'Last reply'];
    rows = d.agents
      .filter(function (a) { return !agent || a.username === agent; })
      .map(function (a) { return [a.name, a.username, a.replies, a.chats, a.avgRespSec === null ? '' : a.avgRespSec, a.slaRate === null ? '' : a.slaRate, fmtTs_(a.lastTs)]; });
  } else if (kind === 'agentday') {
    head = ['Date', 'Agent', 'Username', 'Replies', 'Chats replied'];
    rows = d.agentDay
      .filter(function (a) { return !agent || a.username === agent; })
      .map(function (a) { return [a.day, a.name, a.username, a.replies, a.chats]; });
  } else if (kind === 'hour') {
    head = ['Hour', 'Messages received', 'Agent replies', 'Phone replies'];
    rows = d.hourly.map(function (x) {
      const h = (x.hour < 10 ? '0' : '') + x.hour;
      return [h + ':00-' + h + ':59', x.incoming, x.agentReplies, x.phoneReplies];
    });
  } else if (kind === 'chat') {
    head = ['Chat ID', 'Name', 'Type', 'Messages received', 'Replies', 'Status', 'Last reply by', 'Last activity'];
    rows = d.chats.map(function (c) {
      return [c.id, c.name, c.isGroup ? 'Group' : 'Individual', c.incoming, c.replies, c.status, c.lastAgent, fmtTs_(c.lastTs)];
    });
  } else if (kind === 'pending') {
    head = ['Chat ID', 'Name', 'Type', 'Unanswered messages', 'Waiting since', 'Waiting (minutes)', 'Over SLA target'];
    rows = d.pending.map(function (p) {
      return [p.id, p.name, p.isGroup ? 'Group' : 'Individual', p.unanswered, fmtTs_(p.sinceMs), p.waitMin, p.overdue ? 'Yes' : 'No'];
    });
  } else {
    throw new Error('Unknown export type');
  }
  let csv = csvLine_(head);
  rows.forEach(function (r) { csv += csvLine_(r); });
  return { csv: csv, done: true, rows: rows.length };
}

// All messages (raw data). Fetched page by page; the client joins the pages.
function exportMessages_(filter, page) {
  const rng = dashRange_(filter);
  const type = filter.type || 'all';
  const agent = filter.agent || '';
  const sh = sheet_(MSG_SHEET, MSG_HEADERS);
  const last = sh.getLastRow();
  const header = ['Time', 'Chat ID', 'Chat name', 'Type', 'Direction', 'Agent', 'Agent username',
    'Sender', 'Message', 'Media type', 'Media file', 'Quoted', 'Delivery status'];
  let csv = page === 0 ? csvLine_(header) : '';
  const start = 2 + page * MSG_EXPORT_PAGE;
  if (last < start) return { csv: csv, done: true, rows: 0 };

  const n = Math.min(MSG_EXPORT_PAGE, last - start + 1);
  const rows = sh.getRange(start, 1, n, MSG_HEADERS.length).getValues();
  const ackName = { '-1': 'error', '0': 'pending', '1': 'sent', '2': 'delivered', '3': 'read', '4': 'played' };
  let count = 0;
  rows.forEach(function (r) {
    if (!(r[C.TIME] instanceof Date)) return;
    const t = r[C.TIME].getTime();
    if (t < rng[0] || t > rng[1]) return;
    const chat = r[C.CHAT];
    if (!allowedChat_(chat) || isPlaceholder_(r[C.BODY]) || !typeOk_(chat, type)) return;
    if (agent && r[C.AGENT] !== agent) return;
    const out = r[C.DIR] === 'out';
    csv += csvLine_([
      r[C.TIME], chat, r[C.CHAT_NAME], /@g\.us$/.test(chat) ? 'Group' : 'Individual',
      out ? 'Sent' : 'Received', out ? (r[C.AGENT_NAME] || 'Phone') : '', r[C.AGENT],
      r[C.SENDER], r[C.BODY], r[C.MTYPE], r[C.MNAME], r[C.QUOTED],
      out ? (ackName[String(r[C.ACK])] || '') : '',
    ]);
    count++;
  });
  return { csv: csv, done: start + n > last, rows: count };
}

/* ---------- Chat open/close + details ---------- */

// Manual close/reopen (a reply closes automatically; a customer message reopens automatically)
function setChatStatus(token, chatId, status) {
  const me = session_(token);
  assertChatAccess_(me, chatId);
  if (status !== 'open' && status !== 'closed') throw new Error('Invalid status');
  ensureChatCols_();
  const cs = sheet_(CHAT_SHEET, CHAT_HEADERS);
  const found = cs.getRange('A:A').createTextFinder(chatId).matchEntireCell(true).findNext();
  if (!found) throw new Error('This chat has not been saved yet. Send the first message.');
  const r = found.getRow();
  cs.getRange(r, 8).setValue(status);
  if (status === 'closed') cs.getRange(r, 5).setValue(0);
  setVer_();
  return true;
}

// Full details of one chat: message counts, replies per agent, how many times the customer wrote in, etc.
function getChatDetails(token, chatId) {
  const me = session_(token);
  assertChatAccess_(me, chatId);
  const info = {
    chatId: chatId, name: '', isGroup: /@g\.us$/.test(chatId),
    phone: /@c\.us$/.test(chatId) ? chatId.split('@')[0] : '',
    status: 'open', firstTs: 0, lastTs: 0, incoming: 0, outgoing: 0, conversations: 0,
    avgRespSec: null, agents: [], media: { image: 0, video: 0, audio: 0, file: 0 }, participants: 0,
    assignee: '', assigneeName: '', calls: { total: 0, inbound: 0, missed: 0, lastMs: 0, lastStatus: '' },
  };

  const cs = sheet_(CHAT_SHEET, CHAT_HEADERS);
  const found = cs.getRange('A:A').createTextFinder(chatId).matchEntireCell(true).findNext();
  if (found) {
    const cur = cs.getRange(found.getRow(), 2, 1, 9).getValues()[0];
    info.name = String(cur[0]);
    info.status = String(cur[6]) === 'closed' ? 'closed' : 'open';
    info.assignee = String(cur[7] || '');
    info.assigneeName = String(cur[8] || '');
  }
  info.calls = callsForChat_(chatId);

  const sh = sheet_(MSG_SHEET, MSG_HEADERS);
  const last = sh.getLastRow();
  if (last < 2) return info;
  const ids = sh.getRange(2, C.CHAT + 1, last - 1, 1).getValues();
  let first = 0;
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === chatId) { first = i + 2; break; }
  }
  if (!first) return info;

  const block = sh.getRange(first, 1, last - first + 1, MSG_HEADERS.length).getValues();
  const agents = {}, senders = {};
  let prevDir = '', pendingFirst = 0, rtSum = 0, rtN = 0;
  block.forEach(function (r) {
    if (r[C.CHAT] !== chatId || !(r[C.TIME] instanceof Date) || isPlaceholder_(r[C.BODY])) return;
    const t = r[C.TIME].getTime();
    if (!info.firstTs) info.firstTs = t;
    info.lastTs = t;
    if (r[C.CHAT_NAME]) info.name = String(r[C.CHAT_NAME]);

    const mt = String(r[C.MTYPE]);
    if (mt) {
      const k = mt.indexOf('image/') === 0 ? 'image' : mt.indexOf('video/') === 0 ? 'video' : mt.indexOf('audio/') === 0 ? 'audio' : 'file';
      info.media[k]++;
    }
    if (r[C.DIR] === 'in') {
      info.incoming++;
      if (prevDir !== 'in') { info.conversations++; pendingFirst = t; }
      if (r[C.SENDER]) senders[r[C.SENDER]] = 1;
    } else {
      info.outgoing++;
      const nm = r[C.AGENT] ? (r[C.AGENT_NAME] || r[C.AGENT]) : 'Phone';
      agents[nm] = (agents[nm] || 0) + 1;
      if (pendingFirst) { rtSum += t - pendingFirst; rtN++; pendingFirst = 0; }
    }
    prevDir = r[C.DIR];
  });

  info.avgRespSec = rtN ? Math.round(rtSum / rtN / 1000) : null;
  info.participants = info.isGroup ? Object.keys(senders).length : 0;
  info.agents = Object.keys(agents).map(function (k) { return { name: k, replies: agents[k] }; })
    .sort(function (a, b) { return b.replies - a.replies; });
  return info;
}

/* ---------- Call portal (page=calls) ---------- */
// WhatsApp Web cannot answer or place calls. The portal alerts agents to incoming calls,
// lets them take/reject them, and records history. Calls are answered in WhatsApp on the phone.


function callSheet_() { return sheet_(CALL_SHEET, CALL_HEADERS); }
function callVer_() { PropertiesService.getScriptProperties().setProperty('CALL_VER', String(Date.now())); }
function toMs_(v) { return v instanceof Date ? v.getTime() : 0; }

function readRing_() {
  try { return JSON.parse(prop_('RING') || '[]'); } catch (e) { return []; }
}
function writeRing_(list) {
  const cutoff = Date.now() - 120000;
  PropertiesService.getScriptProperties().setProperty('RING', JSON.stringify(list.filter(function (r) { return r.start > cutoff; })));
}
function removeRing_(callId) {
  writeRing_(readRing_().filter(function (r) { return r.id !== callId; }));
}

// A call that has been "ringing" for longer than RING_MS counts as missed
function effCallStatus_(status, startMs) {
  return (status === 'ringing' && Date.now() - startMs > RING_MS) ? 'missed' : status;
}

function callObj_(r) {
  const startMs = toMs_(r[CL.START]);
  return {
    id: String(r[CL.ID]),
    dir: String(r[CL.DIR]),
    chatId: String(r[CL.CHAT]),
    number: String(r[CL.NUM]),
    name: String(r[CL.NAME]),
    type: String(r[CL.TYPE]),
    isGroup: String(r[CL.GROUP]) === 'yes',
    startMs: startMs,
    takenMs: toMs_(r[CL.TAKEN]),
    endMs: toMs_(r[CL.END]),
    durSec: r[CL.DUR] === '' || r[CL.DUR] === null ? null : Number(r[CL.DUR]),
    status: effCallStatus_(String(r[CL.STATUS]), startMs),
    agent: String(r[CL.AGENT]),
    agentName: String(r[CL.AGENT_NAME]),
    outcome: String(r[CL.OUTCOME]),
    notes: String(r[CL.NOTES]),
  };
}

// Called by the bridge when WhatsApp reports an incoming call
function saveCall_(d) {
  if (!d.callId || !allowedChat_(d.chatId)) return;
  const cache = CacheService.getScriptCache();
  if (cache.get('c_' + d.callId)) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (cache.get('c_' + d.callId)) return;
    cache.put('c_' + d.callId, '1', 21600);
    const now = new Date();
    const number = /@c\.us$/.test(d.chatId) ? d.chatId.split('@')[0] : '';
    appendRow_(callSheet_(), [
      d.callId, 'in', d.chatId, number, d.name || '', d.isVideo ? 'video' : 'voice', d.isGroup ? 'yes' : '',
      now, '', '', '', 'ringing', '', '', '', '',
    ]);
    const ring = readRing_();
    ring.push({
      id: d.callId, start: now.getTime(), chatId: d.chatId, number: number,
      name: d.name || '', type: d.isVideo ? 'video' : 'voice', isGroup: !!d.isGroup,
    });
    writeRing_(ring);
    callVer_();
  } finally {
    lock.releaseLock();
  }
}

function findCallRow_(callId) {
  const sh = callSheet_();
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const start = Math.max(2, last - 999);
  const ids = sh.getRange(start, 1, last - start + 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0] === callId) return start + i;
  }
  return 0;
}

// Marks old "ringing" rows as missed in the Sheet
function expireRinging_() {
  const sh = callSheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const start = Math.max(2, last - 199);
  const vals = sh.getRange(start, 1, last - start + 1, CALL_HEADERS.length).getValues();
  let changed = false;
  vals.forEach(function (r, i) {
    if (String(r[CL.STATUS]) === 'ringing' && Date.now() - toMs_(r[CL.START]) > RING_MS) {
      sh.getRange(start + i, CL.STATUS + 1).setValue('missed');
      changed = true;
    }
  });
  if (changed) callVer_();
}

function notRingingMsg_(r, st) {
  if (st === 'taken') return 'Already taken by ' + (String(r[CL.AGENT_NAME]) || 'another agent');
  if (st === 'rejected') return 'This call was already rejected';
  if (st === 'missed') return 'This call has already ended (missed)';
  return 'This call is no longer ringing';
}

// Light poll for the call portal: bridge status, change version, and calls that are ringing now
function pulseCalls(token) {
  const me = session_(token, 'calls');
  CacheService.getScriptCache().put('cpres_' + me.username, String(Date.now()), 90); // shown as online in the admin portal
  const p = PropertiesService.getScriptProperties().getProperties();
  let status = 'offline';
  try {
    const s = JSON.parse(p.STATUS || 'null');
    if (s && Date.now() - s.ts < 150000) status = s.status;
  } catch (e) { /* ignore */ }
  const now = Date.now();
  let ring = [];
  try { ring = JSON.parse(p.RING || '[]'); } catch (e) { /* ignore */ }
  ring = ring.filter(function (r) { return now - r.start < RING_MS; }).map(function (r) {
    return {
      id: r.id, chatId: r.chatId, number: r.number, name: r.name, type: r.type,
      isGroup: r.isGroup, ageSec: Math.round((now - r.start) / 1000),
    };
  });
  return { status: status, ver: p.CALL_VER || '0', ring: ring, me: { name: me.name, role: me.role, username: me.username } };
}

function callFilterRows_(filter) {
  filter = filter || {};
  const sh = callSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, CALL_HEADERS.length).getValues();
  const from = parseDayParts_(filter.fromDate, false) || 0;
  const to = parseDayParts_(filter.toDate, true) || (Date.now() + 86400000);
  const q = String(filter.q || '').trim().toLowerCase();
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const c = callObj_(rows[i]);
    if (!c.id || c.startMs < from || c.startMs > to) continue;
    if (filter.direction && c.dir !== filter.direction) continue;
    if (filter.status && c.status !== filter.status) continue;
    if (filter.agent && c.agent !== filter.agent) continue;
    if (q && (c.name + ' ' + c.number + ' ' + c.chatId).toLowerCase().indexOf(q) < 0) continue;
    out.push(c);
  }
  out.sort(function (a, b) { return b.startMs - a.startMs; });
  return out;
}

// filter: { direction: 'in'|'out'|'', status, agent, fromDate, toDate (yyyy-MM-dd), q }
function getCalls(token, filter) {
  session_(token, 'calls');
  expireRinging_();
  const rows = callFilterRows_(filter);
  const summary = {
    total: rows.length, inbound: 0, outbound: 0, answered: 0, missed: 0, rejected: 0,
    outboundDone: 0, talkSec: 0, agents: [],
  };
  const agents = {};
  rows.forEach(function (c) {
    if (c.dir === 'in') {
      summary.inbound++;
      if (c.status === 'taken' || c.status === 'completed') summary.answered++;
      else if (c.status === 'missed') summary.missed++;
      else if (c.status === 'rejected') summary.rejected++;
    } else {
      summary.outbound++;
      if (c.status === 'completed') summary.outboundDone++;
    }
    if (c.durSec) summary.talkSec += c.durSec;
    if (c.agent) {
      const a = agents[c.agent] || (agents[c.agent] = { username: c.agent, name: c.agentName || c.agent, inbound: 0, outbound: 0, talkSec: 0 });
      if (c.dir === 'in' && (c.status === 'taken' || c.status === 'completed')) a.inbound++;
      if (c.dir === 'out') a.outbound++;
      if (c.durSec) a.talkSec += c.durSec;
    }
  });
  summary.agents = Object.keys(agents).map(function (k) { return agents[k]; })
    .sort(function (a, b) { return (b.inbound + b.outbound) - (a.inbound + a.outbound); });
  return { calls: rows.slice(0, 300), total: rows.length, summary: summary };
}

// An agent claims a ringing call (other agents stop seeing it ring). The call itself is answered in WhatsApp.
function takeCall(token, callId) {
  const me = session_(token, 'calls');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const row = findCallRow_(callId);
    if (!row) throw new Error('Call not found');
    const sh = callSheet_();
    const r = sh.getRange(row, 1, 1, CALL_HEADERS.length).getValues()[0];
    const st = effCallStatus_(String(r[CL.STATUS]), toMs_(r[CL.START]));
    if (st !== 'ringing') throw new Error(notRingingMsg_(r, st));
    sh.getRange(row, CL.TAKEN + 1).setValue(new Date());
    sh.getRange(row, CL.STATUS + 1).setValue('taken');
    sh.getRange(row, CL.AGENT + 1, 1, 2).setValues([[me.username, me.name]]);
    removeRing_(callId);
    callVer_();
    return true;
  } finally {
    lock.releaseLock();
  }
}

// Rejects a ringing call on WhatsApp (through the bridge)
function rejectCall(token, callId) {
  const me = session_(token, 'calls');
  const row = findCallRow_(callId);
  if (!row) throw new Error('Call not found');
  const sh = callSheet_();
  const r = sh.getRange(row, 1, 1, CALL_HEADERS.length).getValues()[0];
  const st = effCallStatus_(String(r[CL.STATUS]), toMs_(r[CL.START]));
  if (st !== 'ringing') throw new Error(notRingingMsg_(r, st));
  bridge_('/call/reject', { callId: callId });

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    sh.getRange(row, CL.END + 1).setValue(new Date());
    sh.getRange(row, CL.STATUS + 1).setValue('rejected');
    sh.getRange(row, CL.AGENT + 1, 1, 2).setValues([[me.username, me.name]]);
    removeRing_(callId);
    callVer_();
  } finally {
    lock.releaseLock();
  }
  return true;
}

// Finishes a taken / in-progress call (saves duration, outcome, notes) or only adds a note to any other call
function finishCall(token, callId, outcome, notes, durationSec) {
  const me = session_(token, 'calls');
  outcome = String(outcome || '').slice(0, 60);
  notes = String(notes || '').slice(0, 1000);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const row = findCallRow_(callId);
    if (!row) throw new Error('Call not found');
    const sh = callSheet_();
    const r = sh.getRange(row, 1, 1, CALL_HEADERS.length).getValues()[0];
    const st = String(r[CL.STATUS]);
    if (st === 'taken' || st === 'in-progress') {
      if (String(r[CL.AGENT]) !== me.username && me.role !== 'admin') {
        throw new Error('Only the agent who took this call can finish it');
      }
      const takenMs = toMs_(r[CL.TAKEN]) || toMs_(r[CL.START]);
      let dur = (durationSec === null || durationSec === undefined || durationSec === '')
        ? Math.round((Date.now() - takenMs) / 1000)
        : Math.round(Number(durationSec));
      if (!(dur >= 0)) dur = 0;
      dur = Math.min(dur, 6 * 3600);
      sh.getRange(row, CL.END + 1, 1, 2).setValues([[new Date(), dur]]);
      sh.getRange(row, CL.STATUS + 1).setValue('completed');
    }
    sh.getRange(row, CL.OUTCOME + 1, 1, 2).setValues([[outcome, notes]]);
    callVer_();
    return true;
  } finally {
    lock.releaseLock();
  }
}

// Outbound: the agent starts a call from WhatsApp (opened by the portal) and the portal logs it
function startOutbound(token, number, name, type) {
  const me = session_(token, 'calls');
  const digits = String(number || '').replace(/\D/g, '');
  if (digits.length < 8 || digits.length > 15) throw new Error('Enter a valid number with country code');
  const id = 'out_' + Utilities.getUuid();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const now = new Date();
    appendRow_(callSheet_(), [
      id, 'out', digits + '@c.us', digits, String(name || '').slice(0, 80), type === 'video' ? 'video' : 'voice', '',
      now, now, '', '', 'in-progress', me.username, me.name, '', '',
    ]);
    callVer_();
  } finally {
    lock.releaseLock();
  }
  return { callId: id };
}

function exportCallsCsv(token, filter) {
  session_(token, 'calls');
  expireRinging_();
  const rows = callFilterRows_(filter).slice(0, 50000);
  let csv = csvLine_(['Start time', 'Direction', 'Number', 'Name', 'Type', 'Group', 'Status', 'Agent',
    'Taken / started', 'Ended', 'Duration (sec)', 'Outcome', 'Notes']);
  rows.forEach(function (c) {
    csv += csvLine_([
      new Date(c.startMs), c.dir === 'in' ? 'Inbound' : 'Outbound', c.number || c.chatId, c.name, c.type,
      c.isGroup ? 'Yes' : '', c.status, c.agentName, c.takenMs ? new Date(c.takenMs) : '',
      c.endMs ? new Date(c.endMs) : '', c.durSec === null ? '' : c.durSec, c.outcome, c.notes,
    ]);
  });
  return { csv: csv, rows: rows.length };
}

/* ---------- Chat assignment ---------- */

// Agents a chat can be assigned to (the built-in admin + active chat agents; call agents are excluded)
function assignableAgents_() {
  const out = [];
  const adminU = String(prop_('ADMIN_USER') || '').toLowerCase();
  if (adminU) out.push({ username: adminU, name: prop_('ADMIN_NAME') || 'Admin', role: 'admin' });
  agentRows_().forEach(function (a) {
    if (a.active && a.role !== 'caller') out.push({ username: a.username, name: a.name, role: a.role === 'admin' ? 'admin' : 'agent' });
  });
  return out;
}

function getAssignableAgents(token) {
  session_(token);
  const online = onlineMap_();
  return assignableAgents_().map(function (a) {
    return { username: a.username, name: a.name, role: a.role, online: !!online[a.username] };
  });
}

// username '' = unassign. Admins can assign to anyone; agents only to themselves.
function setChatAssignee(token, chatId, username) {
  const me = session_(token);
  if (/@g\.us$/.test(chatId)) throw new Error('Group chats are shared with every agent and cannot be assigned');
  assertChatAccess_(me, chatId);
  username = String(username || '').trim().toLowerCase();
  let name = '';
  if (username) {
    if (me.role !== 'admin' && username !== me.username) throw new Error('You can only assign chats to yourself');
    const hit = assignableAgents_().filter(function (a) { return a.username === username; })[0];
    if (!hit) throw new Error('Unknown or inactive agent');
    name = hit.name;
  }
  ensureChatCols_();
  const cs = sheet_(CHAT_SHEET, CHAT_HEADERS);
  const found = cs.getRange('A:A').createTextFinder(chatId).matchEntireCell(true).findNext();
  if (!found) throw new Error('This chat has not been saved yet. Send the first message.');
  cs.getRange(found.getRow(), 9, 1, 2).setValues([[username, name]]);
  setVer_();
  return true;
}

/* ---------- Quick replies (type /shortcut in the message box) ---------- */

function getQuickReplies(token) {
  session_(token);
  const sh = sheet_(QR_SHEET, QR_HEADERS);
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 2).getValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) { return { shortcut: String(r[0]), text: String(r[1]) }; })
    .sort(function (a, b) { return a.shortcut < b.shortcut ? -1 : 1; });
}

function findQuickRow_(sh, key) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const keys = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0]) === key) return i + 2;
  }
  return 0;
}

function saveQuickReply(token, shortcut, text) {
  adminOnly_(token);
  const key = String(shortcut || '').trim().toLowerCase().replace(/^\//, '');
  if (!/^[a-z0-9_-]{1,20}$/.test(key)) throw new Error('Shortcut: 1-20 characters, letters, digits, - or _ (without the leading /)');
  text = String(text || '').trim();
  if (!text) throw new Error('Message text is required');
  if (text.length > 1000) throw new Error('Message text is too long (max 1000 characters)');
  const sh = sheet_(QR_SHEET, QR_HEADERS);
  const row = findQuickRow_(sh, key);
  if (row) sh.getRange(row, 2).setValue(text);
  else appendRow_(sh, [key, text, new Date()]);
  return true;
}

function deleteQuickReply(token, shortcut) {
  adminOnly_(token);
  const sh = sheet_(QR_SHEET, QR_HEADERS);
  const row = findQuickRow_(sh, String(shortcut || '').toLowerCase());
  if (row) sh.deleteRow(row);
  return true;
}

/* ---------- Bridge offline alert (email) ---------- */
// Run installBridgeMonitor() once from the editor and allow the email permission.
// Optional script property ALERT_EMAIL = where to send alerts (defaults to the script owner).

function checkBridge() {
  const p = PropertiesService.getScriptProperties();
  let down = true;
  let why = 'The bridge has not reported in yet.';
  try {
    const s = JSON.parse(p.getProperty('STATUS') || 'null');
    if (s) {
      const ageMin = Math.round((Date.now() - s.ts) / 60000);
      if (Date.now() - s.ts > 180000) why = 'No heartbeat from the bridge for ' + ageMin + ' minutes.';
      else if (s.status !== 'ready') why = 'The bridge is running but WhatsApp status is "' + s.status + '".';
      else down = false;
    }
  } catch (e) { /* treat as down */ }

  const state = p.getProperty('ALERT_STATE') || 'ok';
  const to = p.getProperty('ALERT_EMAIL') || Session.getEffectiveUser().getEmail();
  if (down) {
    const miss = Number(p.getProperty('ALERT_MISS') || 0) + 1;
    p.setProperty('ALERT_MISS', String(miss));
    // Two checks in a row (about 10 minutes) so a quick restart does not trigger an email
    const lastSent = Number(p.getProperty('ALERT_SENT') || 0);
    if (miss >= 2 && (state === 'ok' || Date.now() - lastSent > 3600000)) {
      MailApp.sendEmail(to, 'WhatsApp portal: bridge is DOWN',
        why + '\n\nCheck that the PC is on and that "node server.js" and ngrok are running.' +
        '\n\nThis alert repeats every hour until the bridge is back.');
      p.setProperty('ALERT_SENT', String(Date.now()));
      p.setProperty('ALERT_STATE', 'down');
    }
  } else {
    p.setProperty('ALERT_MISS', '0');
    if (state === 'down') {
      MailApp.sendEmail(to, 'WhatsApp portal: bridge is back online', 'The bridge is connected to WhatsApp again.');
      p.setProperty('ALERT_STATE', 'ok');
    }
  }
}

function installBridgeMonitor() {
  removeBridgeMonitor();
  ScriptApp.newTrigger('checkBridge').timeBased().everyMinutes(5).create();
  Logger.log('Bridge monitor installed: checks every 5 minutes.');
}

function removeBridgeMonitor() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'checkBridge') ScriptApp.deleteTrigger(t);
  });
}

/* ---------- Key helper ---------- */
// Run from the editor, then copy the key from the Execution log into the BRIDGE_KEY script property
// and into config.bat on the PC.
function generateBridgeKey() {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    Utilities.getUuid() + Utilities.getUuid() + Date.now(),
    Utilities.Charset.UTF_8
  );
  const key = bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  Logger.log('New BRIDGE_KEY: ' + key);
}

/* ---------- Who sees which chat, and automatic assignment ---------- */
// - Group chats are shown to every agent.
// - A 1:1 chat is shown only to the agent it is assigned to. Admins see everything.
//   A chat with no assignee yet is shown to everyone so that anyone can pick it up.
// - When a customer writes in, the chat goes to a logged-in agent: its own agent if that agent is logged in,
//   otherwise the logged-in agent with the fewest open chats.
// - "Logged in" means the agent's portal tab has checked in during the last 90 seconds.

function presenceKey_(username) { return 'pres_' + username; }

function onlineMap_() {
  const users = assignableAgents_().map(function (a) { return a.username; });
  if (!users.length) return {};
  const got = CacheService.getScriptCache().getAll(users.map(presenceKey_));
  const out = {};
  users.forEach(function (u) { if (got[presenceKey_(u)]) out[u] = true; });
  return out;
}

function isOnline_(username) {
  return !!CacheService.getScriptCache().get(presenceKey_(username));
}

function onlineAgents_() {
  const m = onlineMap_();
  return assignableAgents_().filter(function (a) { return m[a.username]; });
}

// Prefer logged-in agents; use logged-in admins only when no agent is logged in
function agentPool_(online) {
  const agents = online.filter(function (a) { return a.role === 'agent'; });
  return agents.length ? agents : online.filter(function (a) { return a.role === 'admin'; });
}

function chatRows_() {
  const cs = sheet_(CHAT_SHEET, CHAT_HEADERS);
  const n = cs.getLastRow();
  return n < 2 ? [] : cs.getRange(2, 1, n - 1, 10).getValues();
}

// Open 1:1 chats per agent
function chatLoads_(rows) {
  const loads = {};
  rows.forEach(function (r) {
    if (allowedChat_(r[0]) && !/@g\.us$/.test(r[0]) && String(r[7]) !== 'closed' && r[8]) {
      loads[r[8]] = (loads[r[8]] || 0) + 1;
    }
  });
  return loads;
}

function leastBusy_(pool, loads) {
  let best = null;
  pool.forEach(function (a) {
    const l = loads[a.username] || 0;
    if (!best || l < best.l || (l === best.l && a.name < best.a.name)) best = { a: a, l: l };
  });
  return best ? best.a : null;
}

function pickAgentForChat_() {
  const pool = agentPool_(onlineAgents_());
  if (!pool.length) return null;
  return leastBusy_(pool, chatLoads_(chatRows_()));
}

function canSeeChat_(me, chatId, assignee) {
  if (me.role === 'admin' || /@g\.us$/.test(String(chatId))) return true;
  const a = String(assignee || '');
  return !a || a === me.username;
}

function chatAssignee_(chatId) {
  const cs = sheet_(CHAT_SHEET, CHAT_HEADERS);
  const found = cs.getRange('A:A').createTextFinder(chatId).matchEntireCell(true).findNext();
  return found ? String(cs.getRange(found.getRow(), 9).getValue()) : '';
}

function assertChatAccess_(me, chatId) {
  if (me.role === 'admin' || /@g\.us$/.test(String(chatId))) return;
  const a = chatAssignee_(chatId);
  if (a && a !== me.username) throw new Error('This chat is assigned to another agent');
}

// Called on every portal check-in. When an agent shows up after being away, waiting chats are handed out.
function touchPresence_(me) {
  const cache = CacheService.getScriptCache();
  const key = presenceKey_(me.username);
  const wasOnline = !!cache.get(key);
  cache.put(key, String(Date.now()), 90);
  if (!wasOnline) {
    try { distributeWaitingChats_(); } catch (e) { console.error('Could not distribute waiting chats: ' + e); }
  }
}

// Open 1:1 chats that have no agent, or whose agent is not logged in, go to the least busy logged-in agent.
// Chats that already got a reply (closed) stay with their agent until the customer writes again.
function distributeWaitingChats_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) return;
  try {
    ensureChatCols_();
    const online = onlineAgents_();
    const pool = agentPool_(online);
    if (!pool.length) return;
    const onlineSet = {};
    online.forEach(function (a) { onlineSet[a.username] = true; });
    const rows = chatRows_();
    if (!rows.length) return;
    const loads = chatLoads_(rows);
    let changed = false;
    rows.forEach(function (r) {
      const id = String(r[0]);
      if (!allowedChat_(id) || /@g\.us$/.test(id)) return;
      if (String(r[7]) === 'closed') return;
      const cur = String(r[8]);
      if (cur && onlineSet[cur]) return;
      const pick = leastBusy_(pool, loads);
      if (!pick) return;
      r[8] = pick.username;
      r[9] = pick.name;
      loads[pick.username] = (loads[pick.username] || 0) + 1;
      changed = true;
    });
    if (changed) {
      sheet_(CHAT_SHEET, CHAT_HEADERS).getRange(2, 9, rows.length, 2)
        .setValues(rows.map(function (r) { return [r[8], r[9]]; }));
      setVer_();
    }
  } finally {
    lock.releaseLock();
  }
}

// Called when an agent logs out, so they stop receiving new chats right away
function signOut(token) {
  const me = session_(token);
  CacheService.getScriptCache().remove(presenceKey_(me.username));
  return true;
}

/* ---------- Admin portal: team, agent profile, settings, links ---------- */

// Every account: the built-in admin plus the Agents sheet
function teamAccounts_() {
  const out = [];
  const adminU = String(prop_('ADMIN_USER') || '').toLowerCase();
  if (adminU) out.push({ username: adminU, name: prop_('ADMIN_NAME') || 'Admin', role: 'admin', active: true, builtin: true });
  agentRows_().forEach(function (a) {
    out.push({
      username: a.username, name: a.name, active: a.active, builtin: false,
      role: a.role === 'admin' ? 'admin' : (a.role === 'caller' ? 'caller' : 'agent'),
    });
  });
  return out;
}

// Team page: accounts with live status, open chats and replies today
function listTeam(token) {
  adminOnly_(token);
  const accounts = teamAccounts_();
  const cache = CacheService.getScriptCache();
  const keys = [];
  accounts.forEach(function (a) { keys.push(presenceKey_(a.username), 'cpres_' + a.username); });
  const seen = keys.length ? cache.getAll(keys) : {};
  const loads = chatLoads_(chatRows_());

  const t0 = new Date();
  t0.setHours(0, 0, 0, 0);
  const repliesToday = {};
  const sh = sheet_(MSG_SHEET, MSG_HEADERS);
  const last = sh.getLastRow();
  if (last >= 2) {
    const start = Math.max(2, last - 3999);
    sh.getRange(start, 1, last - start + 1, 5).getValues().forEach(function (r) {
      if (r[2] === 'out' && r[3] && r[0] instanceof Date && r[0].getTime() >= t0.getTime()) {
        repliesToday[r[3]] = (repliesToday[r[3]] || 0) + 1;
      }
    });
  }
  return accounts.map(function (a) {
    return {
      username: a.username, name: a.name, role: a.role, active: a.active, builtin: a.builtin,
      online: !!(seen[presenceKey_(a.username)] || seen['cpres_' + a.username]),
      openChats: loads[a.username] || 0,
      repliesToday: repliesToday[a.username] || 0,
    };
  });
}

// Everything about one agent: period summary, replies per day, assigned chats and recent calls
function getAgentProfile(token, username, filter) {
  adminOnly_(token);
  username = String(username || '').toLowerCase();
  const member = teamAccounts_().filter(function (a) { return a.username === username; })[0];
  if (!member) throw new Error('Unknown agent');
  const d = computeDash_(filter);
  const summary = d.agents.filter(function (a) { return a.username === username; })[0] || null;
  const perDay = {};
  d.agentDay.forEach(function (x) { if (x.username === username) perDay[x.day] = { replies: x.replies, chats: x.chats }; });
  const days = d.daily.map(function (x) {
    return { day: x.day, replies: (perDay[x.day] || {}).replies || 0, chats: (perDay[x.day] || {}).chats || 0 };
  });
  const chats = chatRows_()
    .filter(function (r) { return String(r[8]) === username && allowedChat_(r[0]); })
    .map(function (r) {
      return {
        chatId: String(r[0]), name: String(r[1]), last: String(r[2]), timeMs: toMs_(r[3]),
        unread: Number(r[4]) || 0, status: String(r[7]) === 'closed' ? 'closed' : 'open',
      };
    })
    .sort(function (a, b) { return b.timeMs - a.timeMs; });
  const online = !!CacheService.getScriptCache().get(presenceKey_(username)) || !!CacheService.getScriptCache().get('cpres_' + username);
  return {
    member: member, online: online, range: d.range, summary: summary, days: days,
    chats: chats.slice(0, 200), calls: callFilterRows_({ agent: username }).slice(0, 20),
    sla: { min: d.totals.slaMin },
  };
}

function getSettings(token) {
  adminOnly_(token);
  const p = PropertiesService.getScriptProperties().getProperties();
  let bridge = { status: 'offline', ageSec: null };
  try {
    const s = JSON.parse(p.STATUS || 'null');
    if (s) {
      const age = Math.round((Date.now() - s.ts) / 1000);
      bridge = { status: age > 180 ? 'offline' : s.status, ageSec: age };
    }
  } catch (e) { /* keep offline */ }
  let owner = '';
  try { owner = Session.getEffectiveUser().getEmail(); } catch (e) { /* not available */ }
  return {
    slaMinutes: Number(p.SLA_MINUTES) || 10,
    signMode: signMode_(),
    alertEmail: p.ALERT_EMAIL || '',
    ownerEmail: owner,
    bridge: bridge,
    bridgeUrlSet: !!p.BRIDGE_URL,
    alertState: p.ALERT_STATE || 'ok',
  };
}

function saveSettings(token, s) {
  adminOnly_(token);
  s = s || {};
  const sla = Math.round(Number(s.slaMinutes));
  if (!(sla >= 1 && sla <= 1440)) throw new Error('SLA target must be between 1 and 1440 minutes');
  const mode = String(s.signMode || '');
  if (['groups', 'all', 'off'].indexOf(mode) < 0) throw new Error('Choose where the agent name is added');
  const email = String(s.alertEmail || '').trim();
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Enter a valid alert email address, or leave it empty');
  const p = PropertiesService.getScriptProperties();
  p.setProperty('SLA_MINUTES', String(sla));
  p.setProperty('SIGN_MESSAGES', mode);
  if (email) p.setProperty('ALERT_EMAIL', email); else p.deleteProperty('ALERT_EMAIL');
  return true;
}

// Links for the agent and admin portals
function getAppLinks(token) {
  const me = session_(token);
  const base = ScriptApp.getService().getUrl();
  return {
    admin: me.role === 'admin' ? base + '?page=admin' : '',
    chat: base,
  };
}

// Calls with one chat (used in the chat details panel)
function callsForChat_(chatId) {
  const out = { total: 0, inbound: 0, missed: 0, lastMs: 0, lastStatus: '' };
  const sh = callSheet_();
  const last = sh.getLastRow();
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, CALL_HEADERS.length).getValues().forEach(function (r) {
    if (String(r[CL.CHAT]) !== chatId) return;
    const c = callObj_(r);
    out.total++;
    if (c.dir === 'in') out.inbound++;
    if (c.status === 'missed') out.missed++;
    if (c.startMs > out.lastMs) { out.lastMs = c.startMs; out.lastStatus = c.status; }
  });
  return out;
}

// Service tags: list for the WhatsApp picker (see Tags.gs for the list itself)