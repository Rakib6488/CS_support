/**
 * Service tags: which service(s) the agent gave in each reply.
 * WhatsApp tags -> sheet "WpTags", Mail tags -> sheet "MailTags" (kept separate on purpose).
 * One row per tag, so a reply with 3 services gives 3 rows.
 */
const TAG_SHEETS = { WhatsApp: 'WpTags', Mail: 'MailTags' };
const TAG_HEADERS = ['Time', 'ContactId', 'Contact', 'MsgRef', 'Agent', 'AgentName', 'Category'];

const TAG_CATEGORIES = [
  "Payment methods",
  "Installment amounts",
  "Repayment status/update",
  "Repayment via Mobile Recharge/Cashout/Send Money",
  "First-hand customer info (Second-hand mobile)",
  "Partial EMI payment",
  "Change installment due date",
  "Temporary unlock for repayment",
  "Mistakenly paid loan amount",
  "Unable to pay installment on time",
  "Unable to log into App",
  "OTP not received",
  "App PIN reset/error",
  "App features",
  "Phone locked before repayment date",
  "Phone locked after full repayment",
  "EMI phone flash/reset issue",
  "Watermark persists before repayment",
  "Watermark persists after full repayment",
  "Change mobile model",
  "Refund/Money back",
  "Change registered account number",
  "Permanent unlock (Developer option)",
  "Return the mobile device",
  "Factory reset the phone after clear all repayments",
  "BTRC registration",
  "Customer death case",
  "Stolen phone",
  "Guarantor unawareness & threat calls",
  "Installment policy",
  "Late fee info",
  "Updated phone prices & Re-application after an order is returned",
  "PalmPay and Telco (GP/Banglalink) offers",
  "Auto-rejection (Credit Review Policy)",
  "New loan application process",
  "Merchant account/address change",
  "Loan Cancel after approval",
  "Need payment statement",
  "Sales statement / Reconciliation",
  "Shop onboard",
  "Merchant/Shop Related",
  "Change BP account under dealer",
  "Undisclosed charges by BP",
  "BP Partner App issue",
  "Verification process delay",
  "Frequent BP policy changes",
  "General collection team conduct",
  "Smart Buy information",
  "New Smart Buy",
  "Buy Another Phone Request",
  "Smart Buy issue",
  "Non-PalmPay related queries",
  "Others info",
  "Complaint Status",
  "Escalation Request",
  "Feedback",
  "Fraud Activity",
  "CS Agent Complaint",
  "Complaint against BP",
  "Staff Misbehavior",
  "Customer Misbehavior",
  "Abusive Language",
  "Callback Request",
  "EMI/Loan Query",
  "Cash Loan Info",
  "After Installment NOC Request",
  "Incomplete Information Shared",
  "Wrong Information Shared with Customer",
  "Query Not Clear",
  "Payment Not Updated",
  "Account Verification Issue",
  "Customer Verification Issue",
  "Application Status",
  "General Inquiry"
];

function tagSheet_(channel) { return sheet_(TAG_SHEETS[channel], TAG_HEADERS); }

function channelsOf_(filter) {
  const c = (filter && filter.channel) || 'all';
  return c === 'all' ? ['WhatsApp', 'Mail'] : [c];
}

// Checks the tag list. Throws if it is empty or contains an unknown service.
function validateTags_(tags) {
  if (!Array.isArray(tags) || !tags.length) throw new Error('Select at least one service tag before sending');
  const out = [];
  tags.forEach(function (t) {
    const c = String(t);
    if (TAG_CATEGORIES.indexOf(c) < 0) throw new Error('Unknown service tag: ' + c);
    if (out.indexOf(c) < 0) out.push(c);
  });
  return out;
}

// Saves one row per service tag into WpTags or MailTags
function saveTags_(channel, contactId, contact, msgRef, agent, agentName, tags) {
  const clean = validateTags_(tags);
  const sh = tagSheet_(channel);
  const now = new Date();
  const rows = clean.map(function (cat) {
    return [now, contactId, contact, msgRef, agent || '', agentName || '', cat].map(mailSafe_);
  });
  const start = sh.getLastRow() + 1;
  const end = start + rows.length - 1;
  if (end > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), end - sh.getMaxRows() + 100);
  sh.getRange(start, 1, rows.length, TAG_HEADERS.length).setValues(rows);
}

// Reads tag rows from one or more channel sheets, filtered by date range and agent
function readTags_(channels, from, to, agent) {
  const out = [];
  channels.forEach(function (ch) {
    const sh = tagSheet_(ch);
    const n = sh.getLastRow();
    if (n < 2) return;
    sh.getRange(2, 1, n - 1, TAG_HEADERS.length).getValues().forEach(function (r) {
      if (!(r[0] instanceof Date)) return;
      const t = r[0].getTime();
      if (t < from || t > to) return;
      if (agent && String(r[4]) !== agent) return;
      out.push({
        time: t,
        channel: ch,
        contact: String(r[2] || r[1]),
        msgRef: String(r[3]),
        agent: String(r[4]),
        agentName: String(r[5]),
        category: String(r[6]),
      });
    });
  });
  return out;
}

// Web app (WhatsApp inbox): list of services for the tag picker
function getTagCategories(token) {
  session_(token);
  return TAG_CATEGORIES.slice();
}

// Called from doPost for Alimail: category list and mail tags
function tagPost_(d) {
  if (d.key !== prop_('MAIL_KEY')) return json_({ ok: false, error: 'unauthorized' });
  if (d.type === 'mail_categories') return json_({ ok: true, categories: TAG_CATEGORIES.slice() });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    saveTags_('Mail', d.contact || '', d.contact || '', d.messageId || '', d.agent || '', d.agentName || '', d.tags);
    return json_({ ok: true });
  } catch (e) {
    return json_({ ok: false, error: String(e) });
  } finally {
    lock.releaseLock();
  }
}

// Admin dashboard data for the "Service tags" page
function getTagDashboard(token, filter) {
  adminOnly_(token);
  filter = filter || {};
  const rng = dashRange_(filter);
  const tz = Session.getScriptTimeZone();
  const rows = readTags_(channelsOf_(filter), rng[0], rng[1], filter.agent || '');
  const cats = {}, chans = {}, daily = {}, msgs = {};
  rows.forEach(function (r) {
    cats[r.category] = (cats[r.category] || 0) + 1;
    chans[r.channel] = (chans[r.channel] || 0) + 1;
    const day = Utilities.formatDate(new Date(r.time), tz, 'yyyy-MM-dd');
    daily[day] = (daily[day] || 0) + 1;
    msgs[r.channel + '|' + r.msgRef] = true;
  });
  const recent = rows.slice().sort(function (a, b) { return b.time - a.time; }).slice(0, 100);
  return {
    range: { fromMs: rng[0], toMs: rng[1], tz: tz },
    total: rows.length,
    messages: Object.keys(msgs).length,
    channels: chans,
    categories: Object.keys(cats).map(function (k) { return { name: k, count: cats[k] }; })
      .sort(function (a, b) { return b.count - a.count; }),
    daily: Object.keys(daily).sort().map(function (k) { return { day: k, count: daily[k] }; }),
    recent: recent,
  };
}

// CSV download of the tag rows (same filter as the dashboard)
function exportTagsCsv(token, filter) {
  adminOnly_(token);
  filter = filter || {};
  const rng = dashRange_(filter);
  const rows = readTags_(channelsOf_(filter), rng[0], rng[1], filter.agent || '');
  let csv = csvLine_(['Time', 'Channel', 'Contact', 'Message reference', 'Agent', 'Agent name', 'Service']);
  rows.forEach(function (r) {
    csv += csvLine_([new Date(r.time), r.channel, r.contact, r.msgRef, r.agent, r.agentName, r.category]);
  });
  return { csv: csv, rows: rows.length };
}