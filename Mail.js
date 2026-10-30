/**
 * Alimail data, saved in the same Google Sheet as WhatsApp but in separate sheets:
 *   MailMessages : every mail (received and sent), one row per mail
 *   MailThreads  : one row per chat, updated when it changes
 * Script property MAIL_KEY must match MAIL_KEY in alimail/.env
 */

const MAIL_MSG_SHEET = 'MailMessages';
const MAIL_THREAD_SHEET = 'MailThreads';
const MAIL_MSG_HEADERS = ['Time', 'ThreadId', 'Direction', 'Contact', 'From', 'To', 'Subject', 'Preview', 'Agent', 'AgentName', 'MessageId'];
const MAIL_THREAD_HEADERS = ['ThreadId', 'Contact', 'Subject', 'Status', 'FirstTime', 'LastTime', 'LastDirection', 'LastAgent', 'Messages'];

// Stops a value like "=cmd" or "-text" from being read as a formula
function mailSafe_(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

// Creates the two sheets if needed. ThreadId and MessageId are kept as text, so long IDs do not lose digits.
function mailSheets_() {
  const msg = sheet_(MAIL_MSG_SHEET, MAIL_MSG_HEADERS);
  const thr = sheet_(MAIL_THREAD_SHEET, MAIL_THREAD_HEADERS);
  msg.getRange(1, 2, msg.getMaxRows(), 1).setNumberFormat('@');
  msg.getRange(1, 11, msg.getMaxRows(), 1).setNumberFormat('@');
  thr.getRange(1, 1, thr.getMaxRows(), 1).setNumberFormat('@');
  return { msg: msg, thr: thr };
}

// Called from doPost when the Alimail server sends { type: 'mail_sync', key, messages, threads }
function mailPost_(d) {
  if (d.key !== prop_('MAIL_KEY')) return json_({ ok: false, error: 'unauthorized' });
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (d.type === 'mail_sync') mailSync_(d);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function mailSync_(d) {
  const s = mailSheets_();
  const ms = s.msg;

  // Skip mail that is already in the sheet (matched by MessageId)
  const seen = {};
  const n = ms.getLastRow();
  if (n >= 2) ms.getRange(2, 11, n - 1, 1).getValues().forEach(function (r) { seen[String(r[0])] = true; });

  const rows = [];
  (d.messages || []).forEach(function (m) {
    if (!m.messageId || seen[m.messageId]) return;
    seen[m.messageId] = true;
    rows.push([
      new Date(m.time), m.threadId, m.direction, m.contact, m.from, m.to,
      m.subject, m.preview, m.agent, m.agentName, m.messageId,
    ].map(mailSafe_));
  });
  if (rows.length) {
    const start = ms.getLastRow() + 1;
    const need = start + rows.length - 1;
    if (need > ms.getMaxRows()) ms.insertRowsAfter(ms.getMaxRows(), need - ms.getMaxRows() + 100);
    ms.getRange(start, 1, rows.length, MAIL_MSG_HEADERS.length).setValues(rows);
  }

  mailUpsertThreads_(s.thr, d.threads || []);
}

// Updates the row of a chat if it exists, otherwise adds a new row
function mailUpsertThreads_(sh, list) {
  if (!list.length) return;
  const n = sh.getLastRow();
  const index = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 1).getValues().forEach(function (r, i) { index[String(r[0])] = i + 2; });

  list.forEach(function (t) {
    const vals = [
      t.threadId, t.contact, t.subject, t.status, new Date(t.firstTime), new Date(t.lastTime),
      t.lastDirection, t.lastAgent, t.messages,
    ].map(mailSafe_);
    let row = index[t.threadId];
    if (!row) {
      row = sh.getLastRow() + 1;
      index[t.threadId] = row;
    }
    if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 100);
    sh.getRange(row, 1, 1, vals.length).setValues([vals]);
  });
}

// Data for the admin portal "Mail" page. Admin only.
function getMailDashboard(token, filter) {
  adminOnly_(token);
  filter = filter || {};
  const rng = dashRange_(filter);
  const from = rng[0], to = rng[1];
  const tz = Session.getScriptTimeZone();
  const s = mailSheets_();

  const mn = s.msg.getLastRow();
  const msgs = mn >= 2 ? s.msg.getRange(2, 1, mn - 1, MAIL_MSG_HEADERS.length).getValues() : [];
  const totals = { received: 0, sent: 0, customers: 0, agentReplies: 0, threads: 0, open: 0, closed: 0 };
  const custSet = {}, daily = {}, agents = {};

  msgs.forEach(function (r) {
    if (!(r[0] instanceof Date)) return;
    const t = r[0].getTime();
    if (t < from || t > to) return;
    const day = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd');
    const d = daily[day] || (daily[day] = { day: day, received: 0, sent: 0 });
    if (r[2] === 'in') {
      totals.received++;
      d.received++;
      custSet[String(r[3])] = true;
    } else {
      totals.sent++;
      d.sent++;
      const key = String(r[8] || '(unknown)');
      const a = agents[key] || (agents[key] = { name: String(r[9] || key), replies: 0 });
      a.replies++;
      if (r[8]) totals.agentReplies++;
    }
  });
  totals.customers = Object.keys(custSet).length;

  const tn = s.thr.getLastRow();
  const thr = tn >= 2 ? s.thr.getRange(2, 1, tn - 1, MAIL_THREAD_HEADERS.length).getValues() : [];
  const threads = [];
  thr.forEach(function (r) {
    if (!(r[5] instanceof Date)) return;
    const last = r[5].getTime();
    if (last < from || last > to) return;
    const st = String(r[3]) === 'closed' ? 'closed' : 'open';
    if (st === 'open') totals.open++; else totals.closed++;
    threads.push({
      id: String(r[0]), contact: String(r[1]), subject: String(r[2]), status: st,
      lastTime: last, messages: Number(r[8]) || 0, lastAgent: String(r[7] || ''),
    });
  });
  threads.sort(function (a, b) { return b.lastTime - a.lastTime; });
  totals.threads = threads.length;

  const agentList = Object.keys(agents).map(function (k) {
    return { username: k, name: agents[k].name, replies: agents[k].replies };
  }).sort(function (a, b) { return b.replies - a.replies; });

  return {
    range: { fromMs: from, toMs: to, tz: tz },
    totals: totals,
    daily: Object.keys(daily).sort().map(function (k) { return daily[k]; }),
    agents: agentList,
    threads: threads.slice(0, 100),
  };
}