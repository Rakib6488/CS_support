require('dotenv').config();
const express = require('express');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const { ImapFlow } = require('imapflow');
const { simpleParser } = require('mailparser');
const nodemailer = require('nodemailer');
const config = require('./config.json');

const PORT = process.env.PORT || config.port;
const APP_USER = (process.env.ALIMAIL_USER || 'admin').toLowerCase();
const APP_PASS = process.env.ALIMAIL_PASSWORD;
const SENT_FOLDER = process.env.SENT_FOLDER || config.sentFolder;
const WHATSAPP_URL = process.env.WHATSAPP_URL || config.whatsappUrl;
// 0 means all messages. Set MAIL_LIST_LIMIT for a deliberate performance cap.
const LIST_LIMIT = Number(process.env.MAIL_LIST_LIMIT || config.listLimit || 0);
const SESSION_MS = config.sessionHours * 3600 * 1000;
const INBOX = 'INBOX';
const LOGIN_WINDOW_MS = 10 * 60 * 1000;
const LOGIN_MAX_FAILURES = 5;

// Google Sheet sync (Apps Script web app, same /exec URL as the WhatsApp portal)
const GAS_URL = process.env.GAS_URL || '';
const MAIL_KEY = process.env.MAIL_KEY || '';
const AGENT_NAME = process.env.ALIMAIL_NAME || 'Alimail';

const DATA_DIR = path.join(__dirname, 'data');
const STATE_FILE = path.join(DATA_DIR, 'state.json');
fs.mkdirSync(DATA_DIR, { recursive: true });

/* ---------- State: open/closed status per chat ---------- */
function loadState() {
  try {
    const s = JSON.parse(fs.readFileSync(STATE_FILE, 'utf8'));
    return { threads: s.threads || {} };
  } catch (e) { return { threads: {} }; }
}
function saveState(s) {
  // Replace atomically so a process interruption cannot leave truncated JSON.
  const tmp = STATE_FILE + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(s, null, 2));
  fs.renameSync(tmp, STATE_FILE);
}

/* ---------- Mail transport and IMAP ---------- */
const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: Number(process.env.SMTP_PORT || 587),
  secure: Number(process.env.SMTP_PORT) === 465,
  auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS },
});

function sendMail(opts) {
  return transporter.sendMail(Object.assign({
    from: '"Alimail Support" <' + process.env.SMTP_USER + '>',
  }, opts));
}

function imapClient() {
  return new ImapFlow({
    host: process.env.IMAP_HOST,
    port: Number(process.env.IMAP_PORT || 993),
    secure: true,
    auth: { user: process.env.IMAP_USER, pass: process.env.IMAP_PASS },
    logger: false,
  });
}

async function withImap(fn) {
  const c = imapClient();
  await c.connect();
  try { return await fn(c); }
  finally { await c.logout().catch(() => {}); }
}

/* ---------- Helpers ---------- */
function addr(a) {
  return a ? { name: a.name || '', address: a.address || '' } : { name: '', address: '' };
}
function cleanSubject(s) {
  return String(s || '').replace(/^((re|fwd?|fw)\s*:\s*)+/i, '').trim() || '(no subject)';
}
// The other person in the chat: for mail we received it is the sender, for mail we sent it is the recipient
function contactOf(m) { return m.fromMe ? m.to : m.from; }

// Removes quoted old replies so the chat shows only the new text
function cleanText(t) {
  let s = String(t || '').replace(/\r/g, '');
  const cut = s.search(/\n(On [\s\S]{5,250}? wrote:|-{2,}\s*Original Message|From: [^\n]+\nSent:)/i);
  if (cut > 0) s = s.slice(0, cut);
  return s.split('\n').filter((l) => !/^\s*>/.test(l)).join('\n').trim();
}

// A chat is open when the customer wrote last. A manual choice or our reply holds until the customer writes again.
function effectiveStatus(st, lastDate, lastFromMe) {
  if (st && st.at && st.at >= lastDate) return st.status;
  return lastFromMe ? 'closed' : 'open';
}

const previewCache = new Map(); // threadId -> { lastDate, text } (filled when a chat is opened)
const agentOfMsg = new Map();   // messageId -> { agent, agentName, preview } for mail sent from Alimail
const pushedMsgs = new Set();   // messageIds already saved to the Google Sheet
const pushedThreads = new Map(); // threadId -> signature last saved to the Google Sheet
let listCache = { at: 0, data: null };
let tagCatCache = null;          // service categories, loaded once from Apps Script
function invalidateList() { listCache.at = 0; }

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
function layoutHtml(text) {
  return '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1d2330;line-height:1.6;max-width:600px">' +
    '<div style="border-bottom:3px solid #128C7E;padding-bottom:8px;margin-bottom:16px;font-weight:700;color:#128C7E">Alimail Support</div>' +
    '<div>' + escapeHtml(text).replace(/\n/g, '<br>') + '</div>' +
    '<div style="margin-top:28px;font-size:12px;color:#7b8190">This message was sent from our support desk.</div></div>';
}

function uniqueTags(tags) {
  return [...new Set(tags.map((t) => String(t)))];
}

/* ---------- Google Sheet calls (mail sync and service tags) ---------- */
// Sends a request to the Apps Script web app with the MAIL_KEY
async function gasCall(payload) {
  if (!GAS_URL || !MAIL_KEY) throw new Error('GAS_URL and MAIL_KEY must be set in .env');
  const res = await fetch(GAS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(Object.assign({ key: MAIL_KEY }, payload)),
  });
  const j = await res.json().catch(() => ({}));
  if (!j.ok) throw new Error(j.error || 'Apps Script did not accept the request');
  return j;
}

// Saves the service tags of a mail in the MailTags sheet
async function saveMailTags(tags, contact, messageId, agent) {
  return gasCall({
    type: 'mail_tag', contact: contact || '', messageId: messageId || '',
    agent: agent || '', agentName: AGENT_NAME, tags,
  });
}

// Sends new mail and changed chats to the Apps Script web app. Runs in the background; a failure never stops Alimail.
async function pushMail(messages, threads) {
  if (!GAS_URL || !MAIL_KEY || (!messages.length && !threads.length)) return;
  try {
    await gasCall({ type: 'mail_sync', messages, threads });
    messages.forEach((m) => pushedMsgs.add(m.messageId));
    threads.forEach((t) => pushedThreads.set(t.threadId, t.sig));
  } catch (e) {
    console.error('Mail sheet sync failed:', e.message);
  }
}

/* ---------- IMAP reading ---------- */
// Lists the newest mail of one folder with its Gmail thread ID
async function listBox(c, box) {
  const lock = await c.getMailboxLock(box);
  try {
    const all = await c.search({ all: true }, { uid: true });
    const uids = LIST_LIMIT > 0 ? all.slice(-LIST_LIMIT) : all;
    const out = [];
    if (!uids.length) return out;
    for await (const m of c.fetch(uids.join(','), {
      uid: true, flags: true, envelope: true, internalDate: true, threadId: true,
    }, { uid: true })) {
      const env = m.envelope || {};
      out.push({
        box,
        fromMe: box === SENT_FOLDER,
        uid: m.uid,
        date: m.internalDate ? new Date(m.internalDate).getTime() : 0,
        unseen: !(m.flags && m.flags.has('\\Seen')),
        from: addr((env.from || [])[0]),
        to: addr((env.to || [])[0]),
        subject: env.subject || '(no subject)',
        messageId: env.messageId || '',
        // Gmail thread ID groups the inbox mail and our replies. Fallback keeps each mail separate.
        threadKey: m.threadId ? String(m.threadId) : (env.messageId || box + m.uid),
      });
    }
    return out;
  } finally { lock.release(); }
}

// Full content of one chat (all mail with the same thread ID), oldest first
async function fetchThread(tid) {
  return withImap(async (c) => {
    const inbox = await listBox(c, INBOX);
    const sent = await listBox(c, SENT_FOLDER);
    const metas = inbox.concat(sent).filter((m) => m.threadKey === tid);
    const out = [];
    for (const box of [INBOX, SENT_FOLDER]) {
      const part = metas.filter((m) => m.box === box);
      if (!part.length) continue;
      const lock = await c.getMailboxLock(box);
      try {
        for (const m of part) {
          const full = await c.fetchOne(String(m.uid), { source: true, flags: true }, { uid: true });
          if (!full) continue;
          if (box === INBOX && !full.flags.has('\\Seen')) {
            await c.messageFlagsAdd(String(m.uid), ['\\Seen'], { uid: true });
          }
          const p = await simpleParser(full.source);
          out.push({
            key: box + ':' + m.uid,
            fromMe: box === SENT_FOLDER,
            date: m.date,
            from: addr(p.from && p.from.value[0]),
            to: addr(p.to && p.to.value[0]),
            subject: p.subject || '(no subject)',
            messageId: p.messageId || '',
            references: [].concat(p.references || []).filter(Boolean),
            text: cleanText(p.text),
          });
        }
      } finally { lock.release(); }
    }
    out.sort((a, b) => a.date - b.date);
    return out;
  });
}

// All chats (threads) built from Inbox and Sent
async function buildThreads() {
  if (listCache.data && Date.now() - listCache.at < 20000) return listCache.data;

  const msgs = await withImap(async (c) => {
    const inbox = await listBox(c, INBOX);
    const sent = await listBox(c, SENT_FOLDER);
    return inbox.concat(sent);
  });
  const state = loadState();

  const groups = new Map();
  msgs.forEach((m) => {
    if (!groups.has(m.threadKey)) groups.set(m.threadKey, []);
    groups.get(m.threadKey).push(m);
  });

  const list = [];
  const threadPush = [];
  groups.forEach((items, id) => {
    items.sort((a, b) => a.date - b.date);
    const last = items[items.length - 1];
    const inbound = items.filter((x) => !x.fromMe);
    const cust = inbound.length ? inbound[inbound.length - 1] : last;
    const lastFromMe = last.fromMe;
    const prev = previewCache.get(id);
    const status = effectiveStatus(state.threads[id], last.date, lastFromMe);
    const contact = contactOf(cust);
    const subject = cleanSubject(cust.subject);

    list.push({
      id,
      subject,
      contact,
      lastDate: last.date,
      lastFromMe,
      preview: prev && prev.lastDate === last.date ? prev.text : last.subject,
      unread: inbound.filter((x) => x.unseen).length,
      status,
    });

    // Save this chat to the sheet only when something changed
    const sig = last.date + '|' + status + '|' + items.length;
    if (pushedThreads.get(id) !== sig) {
      threadPush.push({
        threadId: id,
        contact: contact.address,
        subject,
        status,
        firstTime: items[0].date,
        lastTime: last.date,
        lastDirection: lastFromMe ? 'out' : 'in',
        lastAgent: lastFromMe ? ((agentOfMsg.get(last.messageId) || {}).agentName || '') : '',
        messages: items.length,
        sig,
      });
    }
  });
  list.sort((a, b) => b.lastDate - a.lastDate);

  // Save every mail that is not in the sheet yet
  const messagePush = [];
  msgs.forEach((m) => {
    if (!m.messageId || pushedMsgs.has(m.messageId)) return;
    const a = agentOfMsg.get(m.messageId) || {};
    messagePush.push({
      time: m.date,
      threadId: m.threadKey,
      direction: m.fromMe ? 'out' : 'in',
      contact: contactOf(m).address,
      from: m.from.address,
      to: m.to.address,
      subject: m.subject,
      preview: a.preview || '',
      agent: a.agent || '',
      agentName: a.agentName || '',
      messageId: m.messageId,
    });
  });
  pushMail(messagePush, threadPush); // runs in the background

  listCache = { at: Date.now(), data: list };
  return list;
}

/* ---------- App ---------- */
const sessions = new Map();
const loginFailures = new Map();
function auth(req, res, next) {
  const t = (req.headers.authorization || '').replace(/^Bearer\s+/, '');
  const s = sessions.get(t);
  if (!s || Date.now() - s.at > SESSION_MS) {
    sessions.delete(t);
    return res.status(401).json({ error: 'Please sign in again' });
  }
  req.user = s.user;
  next();
}
const wrap = (fn) => (req, res) => fn(req, res).catch((e) => res.status(500).json({ error: e.message }));

const app = express();

// Allow Alimail to be shown inside the WhatsApp portal (Google Apps Script)
app.use((req, res, next) => {
  res.setHeader(
    'Content-Security-Policy',
    "frame-ancestors 'self' https://script.google.com https://*.google.com https://*.googleusercontent.com"
  );
  next();
});

app.use(express.json({ limit: '5mb' }));
app.use(express.static(path.join(__dirname, 'public')));

app.get('/api/config', (req, res) => res.json({ whatsappUrl: WHATSAPP_URL }));

app.post('/api/login', (req, res) => {
  const { username, password } = req.body || {};
  if (!APP_PASS) return res.status(500).json({ error: 'ALIMAIL_PASSWORD is not set in .env' });
  const ip = req.ip || req.socket.remoteAddress || 'unknown';
  const now = Date.now();
  const attempt = loginFailures.get(ip);
  if (attempt && now - attempt.started < LOGIN_WINDOW_MS && attempt.count >= LOGIN_MAX_FAILURES) {
    return res.status(429).json({ error: 'Too many failed attempts. Try again later.' });
  }
  if (String(username || '').toLowerCase() !== APP_USER || password !== APP_PASS) {
    const current = loginFailures.get(ip);
    if (!current || now - current.started >= LOGIN_WINDOW_MS) {
      loginFailures.set(ip, { started: now, count: 1 });
    } else {
      current.count++;
    }
    return res.status(401).json({ error: 'Wrong username or password' });
  }
  loginFailures.delete(ip);
  const token = crypto.randomBytes(32).toString('hex');
  sessions.set(token, { user: APP_USER, at: Date.now() });
  res.json({ token, whatsappUrl: WHATSAPP_URL });
});

app.post('/api/logout', auth, (req, res) => {
  const t = (req.headers.authorization || '').replace(/^Bearer\s+/, '');
  sessions.delete(t);
  res.json({ ok: true });
});

// view: open | unread | closed | all
app.get('/api/threads', auth, wrap(async (req, res) => {
  const view = req.query.view || 'open';
  const all = await buildThreads();
  const counts = {
    open: all.filter((t) => t.status === 'open').length,
    unread: all.filter((t) => t.unread > 0).length,
    closed: all.filter((t) => t.status === 'closed').length,
    all: all.length,
  };
  let items = all;
  if (view === 'open') items = all.filter((t) => t.status === 'open');
  else if (view === 'unread') items = all.filter((t) => t.unread > 0);
  else if (view === 'closed') items = all.filter((t) => t.status === 'closed');
  res.json({ items, counts });
}));

// One chat with all its messages. Opening it marks the customer's mail as read.
app.get('/api/threads/:tid', auth, wrap(async (req, res) => {
  const tid = req.params.tid;
  const msgs = await fetchThread(tid);
  if (!msgs.length) return res.status(404).json({ error: 'Chat not found' });

  const last = msgs[msgs.length - 1];
  previewCache.set(tid, { lastDate: last.date, text: (last.text || last.subject).slice(0, 120) });
  invalidateList();

  const inbound = msgs.filter((m) => !m.fromMe);
  const cust = inbound.length ? inbound[inbound.length - 1] : last;
  const st = loadState().threads[tid];

  res.json({
    id: tid,
    subject: cleanSubject(msgs[0].subject),
    contact: contactOf(cust),
    status: effectiveStatus(st, last.date, last.fromMe),
    messages: msgs.map((m) => ({
      key: m.key, fromMe: m.fromMe, date: m.date, subject: m.subject, text: m.text,
      from: m.from, to: m.to,
    })),
  });
}));

// Service categories, loaded from the Apps Script (same list as the WhatsApp inbox)
app.get('/api/tag-categories', auth, wrap(async (req, res) => {
  if (!tagCatCache) {
    const j = await gasCall({ type: 'mail_categories' });
    tagCatCache = j.categories || [];
  }
  res.json({ categories: tagCatCache });
}));

// Reply inside a chat. The mail is threaded to the customer's last message. Service tags are required.
app.post('/api/threads/:tid/reply', auth, wrap(async (req, res) => {
  const body = req.body || {};
  const text = String(body.body || '').trim();
  if (!text) return res.status(400).json({ error: 'Message is empty' });
  if (!Array.isArray(body.tags) || !body.tags.length) return res.status(400).json({ error: 'Select at least one service tag' });
  const tags = uniqueTags(body.tags);

  const msgs = await fetchThread(req.params.tid);
  if (!msgs.length) return res.status(404).json({ error: 'Chat not found' });

  const inbound = msgs.filter((m) => !m.fromMe);
  const target = inbound.length ? inbound[inbound.length - 1] : msgs[msgs.length - 1];
  const to = contactOf(target).address;
  if (!to) return res.status(400).json({ error: 'Customer address not found' });

  const refs = target.references.slice();
  if (target.messageId) refs.push(target.messageId);

  const info = await sendMail({
    to,
    subject: String(body.subject || '').trim() || ('Re: ' + cleanSubject(target.subject)),
    text: text + '\n\n-- Alimail Support',
    html: layoutHtml(text),
    inReplyTo: target.messageId || undefined,
    references: refs.length ? refs.join(' ') : undefined,
  });
  if (info && info.messageId) {
    agentOfMsg.set(info.messageId, { agent: req.user || '', agentName: AGENT_NAME, preview: text.slice(0, 300) });
  }
  await saveMailTags(tags, to, info && info.messageId, req.user);

  // Our reply closes the chat, the same way a WhatsApp reply does
  const st = loadState();
  st.threads[req.params.tid] = { status: 'closed', at: Date.now() };
  saveState(st);
  invalidateList();
  res.json({ ok: true });
}));

// Manual open / close
app.post('/api/threads/:tid/status', auth, wrap(async (req, res) => {
  const status = (req.body || {}).status === 'closed' ? 'closed' : 'open';
  const st = loadState();
  st.threads[req.params.tid] = { status, at: Date.now() };
  saveState(st);
  invalidateList();
  res.json({ ok: true });
}));

// New mail to a customer (starts a new chat). Service tags are required.
app.post('/api/compose', auth, wrap(async (req, res) => {
  const body = req.body || {};
  if (!body.to || !body.body) return res.status(400).json({ error: 'Recipient and message are required' });
  if (!Array.isArray(body.tags) || !body.tags.length) return res.status(400).json({ error: 'Select at least one service tag' });
  const tags = uniqueTags(body.tags);

  const info = await sendMail({
    to: body.to,
    subject: String(body.subject || '').trim() || '(no subject)',
    text: body.body + '\n\n-- Alimail Support',
    html: layoutHtml(body.body),
  });
  if (info && info.messageId) {
    agentOfMsg.set(info.messageId, { agent: req.user || '', agentName: AGENT_NAME, preview: String(body.body).slice(0, 300) });
  }
  await saveMailTags(tags, body.to, info && info.messageId, req.user);
  invalidateList();
  res.json({ ok: true });
}));

app.listen(PORT, () => console.log('Alimail running on http://localhost:' + PORT));
