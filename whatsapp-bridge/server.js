// WhatsApp bridge v2
// QR login + messages/media/groups -> Apps Script (Google Sheet), agent attribution for replies
require('dotenv').config();
const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const express = require('express');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3000;
const BRIDGE_KEY = process.env.BRIDGE_KEY; // must match the BRIDGE_KEY script property
const GAS_URL = process.env.GAS_URL;       // Apps Script web app /exec URL
const MAX_MEDIA_BYTES = 15 * 1024 * 1024;  // larger media is not saved to Drive

// Only individual numbers (@c.us / @lid) and groups (@g.us). Channels (@newsletter), broadcasts and status are excluded.
const ALLOWED_CHAT = /@(c\.us|lid|g\.us)$/;
// WhatsApp system/notification message types are not saved
const SKIP_TYPES = new Set([
  'e2e_notification', 'notification_template', 'notification', 'gp2',
  'call_log', 'protocol', 'broadcast_notification',
]);

if (!BRIDGE_KEY || !GAS_URL) {
  console.error('BRIDGE_KEY and GAS_URL environment variables are required.');
  process.exit(1);
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

let state = { status: 'starting', qr: null }; // qr = raw QR string (rendered in the portal)
let inflight = 0;                              // number of portal sends currently in progress
const agentByMsg = new Map();                  // msgId -> { agent, agentName }

function gas(payload) {
  return axios.post(GAS_URL, Object.assign({ key: BRIDGE_KEY }, payload), {
    timeout: 90000,
    maxRedirects: 5,
    maxBodyLength: Infinity,
    maxContentLength: Infinity,
  }).then((res) => {
    const data = res && res.data;
    if (!data || data.ok !== true) {
      throw new Error((data && data.error) || 'Apps Script rejected the request');
    }
    return res;
  });
}

function pushStatus() {
  return gas({ type: 'status', status: state.status, qr: state.qr }).catch((e) =>
    console.error('status push failed:', e.message)
  );
}
function setState(next) {
  state = next;
  pushStatus();
}
setInterval(pushStatus, 60000); // heartbeat: lets the portal know the bridge is online

// On Render: mount a persistent disk at /data and set AUTH_PATH=/data/wwebjs_auth
const AUTH_PATH = process.env.AUTH_PATH || './.wwebjs_auth';

// Stale Chromium lock files can stop the browser from starting after a restart, so remove them
function clearChromeLocks(dir) {
  try {
    for (const name of fs.readdirSync(dir)) {
      const p = path.join(dir, name);
      let st;
      try { st = fs.lstatSync(p); } catch (e) { continue; }
      if (name === 'SingletonLock' || name === 'SingletonCookie' || name === 'SingletonSocket') {
        try { fs.unlinkSync(p); } catch (e) { /* ignore */ }
      } else if (st.isDirectory()) {
        clearChromeLocks(p);
      }
    }
  } catch (e) { /* directory does not exist yet */ }
}
clearChromeLocks(AUTH_PATH);

const client = new Client({
  authStrategy: new LocalAuth({ dataPath: AUTH_PATH }),
  puppeteer: {
    headless: process.env.HEADLESS !== 'false', // debug: set HEADLESS=false to see the Chrome window
    executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || undefined,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-gpu'],
  },
});

let reiniting = false;
function reinit() {
  if (reiniting) return;
  reiniting = true;
  setTimeout(async () => {
    try { await client.destroy(); } catch (e) { /* ignore */ }
    try { await client.initialize(); } catch (e) { console.error('initialize failed:', e); }
    reiniting = false;
  }, 3000);
}

client.on('qr', (qr) => setState({ status: 'qr', qr }));
client.on('authenticated', () => setState({ status: 'authenticating', qr: null }));
client.on('auth_failure', (m) => { console.error('auth_failure', m); setState({ status: 'disconnected', qr: null }); reinit(); });
client.on('ready', () => { console.log('WhatsApp ready'); setState({ status: 'ready', qr: null }); });
client.on('disconnected', (reason) => {
  console.log('disconnected', reason);
  setState({ status: 'disconnected', qr: null });
  reinit();
});

// Finds the agent for a message that was sent from the portal
async function resolveAgent(id) {
  const deadline = Date.now() + 25000;
  for (;;) {
    const a = agentByMsg.get(id);
    if (a) return a;
    if (inflight === 0) {
      await sleep(400);
      return agentByMsg.get(id) || null; // null = sent from the phone
    }
    if (Date.now() > deadline) return null;
    await sleep(300);
  }
}

// All incoming and outgoing messages arrive here (individual + group)
client.on('message_create', async (msg) => {
  try {
    if (msg.from === 'status@broadcast' || msg.to === 'status@broadcast') return;

    const chatId = msg.fromMe ? msg.to : msg.from;
    if (!ALLOWED_CHAT.test(chatId || '') || SKIP_TYPES.has(msg.type)) return;
    const isGroup = chatId.endsWith('@g.us');

    let chatName = chatId.split('@')[0];
    try {
      const chat = await msg.getChat();
      if (chat && chat.name) chatName = chat.name;
    } catch (e) { /* ignore */ }

    let sender = '';
    if (!msg.fromMe) {
      if (isGroup) {
        try {
          const c = await msg.getContact();
          sender = c.pushname || c.name || c.number || '';
        } catch (e) { /* ignore */ }
        if (!sender) sender = (msg.author || '').split('@')[0];
      } else {
        sender = chatName;
      }
    }

    let body = msg.body || '';
    let media = null;
    if (msg.hasMedia) {
      try {
        const m = await msg.downloadMedia();
        if (m && m.data) {
          if (m.data.length * 0.75 > MAX_MEDIA_BYTES) {
            body = (body ? body + ' ' : '') + '[file too large, not saved]';
          } else {
            const ext = ((m.mimetype || 'application/octet-stream').split('/')[1] || 'bin').split(';')[0];
            media = {
              mimetype: m.mimetype || 'application/octet-stream',
              filename: m.filename || (msg.type + '.' + ext),
              data: m.data,
            };
          }
        }
      } catch (e) {
        console.error('downloadMedia failed:', e.message);
      }
    } else if (!body) {
      body = '[' + msg.type + ']';
    }

    let quoted = '';
    if (msg.hasQuotedMsg) {
      try {
        const q = await msg.getQuotedMessage();
        quoted = q.body || '[' + q.type + ']';
      } catch (e) { /* ignore */ }
    }

    let agent = null;
    if (msg.fromMe) agent = await resolveAgent(msg.id._serialized);

    await gas({
      type: 'message',
      id: msg.id._serialized,
      chatId,
      chatName,
      isGroup,
      sender,
      fromMe: !!msg.fromMe,
      body,
      ts: msg.timestamp,
      quoted,
      media,
      agent: agent ? agent.agent : '',
      agentName: agent ? agent.agentName : '',
    });
  } catch (e) {
    console.error('push failed:', e.message);
  }
});

// Delivery ticks (sent / delivered / read)
client.on('message_ack', (msg, ack) => {
  if (!msg.fromMe || !ALLOWED_CHAT.test(msg.to || '')) return;
  gas({ type: 'ack', id: msg.id._serialized, ack }).catch((e) => console.error('ack push failed:', e.message));
});

// "Delete for everyone"
client.on('message_revoke_everyone', (after, before) => {
  const m = before || after;
  if (!m || !m.id) return;
  const cid = m.fromMe ? m.to : m.from;
  if (!ALLOWED_CHAT.test(cid || '')) return;
  gas({ type: 'revoke', id: m.id._serialized }).catch((e) => console.error('revoke push failed:', e.message));
});

// ---- Incoming WhatsApp calls ----
// WhatsApp Web can detect and reject calls, but it cannot answer them or carry call audio.
const callObjs = new Map();   // callId -> Call object (needed to reject it)
const seenCalls = new Set();

async function onIncomingCall(call) {
  try {
    const id = call && call.id;
    if (!id || seenCalls.has(id)) return;
    seenCalls.add(id);
    setTimeout(() => seenCalls.delete(id), 300000);
    if (call.fromMe || !ALLOWED_CHAT.test(call.from || '')) return;
    callObjs.set(id, call);
    setTimeout(() => callObjs.delete(id), 180000);
    let name = '';
    try {
      const c = await client.getContactById(call.from);
      name = c.pushname || c.name || '';
    } catch (e) { /* ignore */ }
    await gas({
      type: 'call', callId: id, chatId: call.from, name,
      isVideo: !!call.isVideo, isGroup: !!call.isGroup, ts: call.timestamp,
    });
  } catch (e) {
    console.error('call push failed:', e.message);
  }
}
client.on('incoming_call', onIncomingCall);
client.on('call', onIncomingCall); // older library versions use this event name

const app = express();
// Authenticate before parsing large JSON bodies so unauthenticated requests
// cannot force the process to allocate the 60 MB body limit.
app.use((req, res, next) => {
  if (req.headers['x-api-key'] !== BRIDGE_KEY) return res.status(401).json({ error: 'unauthorized' });
  next();
});
app.use(express.json({ limit: '60mb' }));

app.get('/status', (req, res) => res.json(state));

app.post('/send', async (req, res) => {
  inflight++;
  try {
    const { to, text, quotedId, media, agent, agentName } = req.body || {};
    if (!to) return res.status(400).json({ error: "'to' is required" });
    if (!text && !media) return res.status(400).json({ error: 'text or a file is required' });
    if (state.status !== 'ready') return res.status(503).json({ error: 'WhatsApp is not connected' });

    const opts = {};
    if (quotedId) opts.quotedMessageId = quotedId;

    let sent;
    if (media && media.data) {
      const mm = new MessageMedia(media.mimetype, media.data, media.filename);
      if (text) opts.caption = text;
      if (!/^(image|video|audio)\//.test(media.mimetype || '')) opts.sendMediaAsDocument = true;
      sent = await client.sendMessage(to, mm, opts);
    } else {
      sent = await client.sendMessage(to, text, opts);
    }

    const id = sent.id._serialized;
    agentByMsg.set(id, { agent: agent || '', agentName: agentName || '' });
    setTimeout(() => agentByMsg.delete(id), 180000);
    res.json({ ok: true, id });
  } catch (e) {
    res.status(500).json({ error: e.message });
  } finally {
    inflight--;
  }
});

app.post('/call/reject', async (req, res) => {
  try {
    const call = callObjs.get((req.body || {}).callId);
    if (!call) return res.status(404).json({ error: 'Call not found or already ended' });
    await call.reject();
    callObjs.delete(call.id);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/logout', async (req, res) => {
  try { await client.logout(); } catch (e) { /* ignore */ }
  setState({ status: 'starting', qr: null });
  reinit();
  res.json({ ok: true });
});

app.listen(PORT, () => console.log('Bridge listening on', PORT));
client.initialize().catch((e) => console.error('initialize failed:', e));
