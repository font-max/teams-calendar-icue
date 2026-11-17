#!/usr/bin/env node
// Teams Calendar Helper
// Signs in to Microsoft 365 (device-code flow), reads today's calendar from Microsoft Graph,
// and serves it to the iCUE "Teams Calendar" widget on http://localhost:38127.
//
// Usage:
//   node teams-calendar-helper.mjs login [--tenant <id|organizations|common>] [--client-id <guid>]
//   node teams-calendar-helper.mjs serve      (default)
//   node teams-calendar-helper.mjs key        (print the widget access key)
//   node teams-calendar-helper.mjs logout
//
// No npm dependencies. Requires Node 18+ (global fetch) and Windows (DPAPI token encryption).

import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { spawnSync, spawn } from 'node:child_process';

const PORT = 38127;
// Default public client: "Microsoft Graph Command Line Tools" (the app used by Connect-MgGraph).
// Replace with your own app registration via --client-id if your tenant blocks it.
const DEFAULT_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const DEFAULT_TENANT = 'organizations';
const SCOPES = 'offline_access User.Read Calendars.Read';
const GRAPH_REFRESH_MS = 2 * 60 * 1000;

const DATA_DIR = path.join(process.env.LOCALAPPDATA || path.join(process.env.HOME || '.', '.local'), 'TeamsCalendarHelper');
const CONFIG_FILE = path.join(DATA_DIR, 'config.json');
const TOKEN_FILE = path.join(DATA_DIR, 'token.dat');
const LOG_FILE = path.join(DATA_DIR, 'helper.log');

fs.mkdirSync(DATA_DIR, { recursive: true });

// ---------- logging ----------
function log(...args) {
  const line = `[${new Date().toISOString()}] ${args.join(' ')}`;
  console.log(line);
  try {
    if (fs.existsSync(LOG_FILE) && fs.statSync(LOG_FILE).size > 1024 * 1024) fs.renameSync(LOG_FILE, LOG_FILE + '.old');
    fs.appendFileSync(LOG_FILE, line + '\n');
  } catch { /* ignore */ }
}

// ---------- config ----------
function loadConfig() {
  let cfg = {};
  try { cfg = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8')); } catch { /* first run */ }
  let changed = false;
  if (!cfg.key) { cfg.key = crypto.randomBytes(12).toString('hex'); changed = true; }
  if (!cfg.clientId) { cfg.clientId = DEFAULT_CLIENT_ID; changed = true; }
  if (!cfg.tenant) { cfg.tenant = DEFAULT_TENANT; changed = true; }
  if (changed) saveConfig(cfg);
  return cfg;
}
function saveConfig(cfg) {
  fs.writeFileSync(CONFIG_FILE, JSON.stringify(cfg, null, 2));
}

// ---------- DPAPI token storage (Windows, current user) ----------
function dpapi(mode, input) {
  const script =
    "$ErrorActionPreference='Stop';Add-Type -AssemblyName System.Security;" +
    "$in=[Console]::In.ReadToEnd().Trim();" +
    (mode === 'protect'
      ? "$b=[Text.Encoding]::UTF8.GetBytes($in);[Convert]::ToBase64String([Security.Cryptography.ProtectedData]::Protect($b,$null,'CurrentUser'))"
      : "$b=[Convert]::FromBase64String($in);[Text.Encoding]::UTF8.GetString([Security.Cryptography.ProtectedData]::Unprotect($b,$null,'CurrentUser'))");
  const r = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], {
    input, encoding: 'utf8', windowsHide: true,
  });
  if (r.status !== 0) throw new Error('DPAPI ' + mode + ' failed: ' + (r.stderr || r.error));
  return r.stdout.trim();
}
function saveRefreshToken(rt) {
  fs.writeFileSync(TOKEN_FILE, dpapi('protect', rt));
}
function loadRefreshToken() {
  try { return dpapi('unprotect', fs.readFileSync(TOKEN_FILE, 'utf8')); } catch { return null; }
}

// ---------- OAuth ----------
const authBase = (tenant) => `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0`;

async function postForm(url, params) {
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(params),
  });
  const json = await res.json().catch(() => ({}));
  return { ok: res.ok, status: res.status, json };
}

async function login(cfg) {
  const dc = await postForm(`${authBase(cfg.tenant)}/devicecode`, { client_id: cfg.clientId, scope: SCOPES });
  if (!dc.ok) throw new Error('Device code request failed: ' + (dc.json.error_description || dc.status));
  const { device_code, user_code, verification_uri, interval = 5, expires_in = 900, message } = dc.json;
  console.log('\n' + message + '\n');
  console.log(`Code: ${user_code}`);
  try { spawn('cmd.exe', ['/c', 'start', '', verification_uri], { detached: true, stdio: 'ignore', windowsHide: true }).unref(); } catch { /* ignore */ }

  let wait = interval * 1000;
  const deadline = Date.now() + expires_in * 1000;
  while (Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, wait));
    const t = await postForm(`${authBase(cfg.tenant)}/token`, {
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      client_id: cfg.clientId,
      device_code,
    });
    if (t.ok) {
      saveRefreshToken(t.json.refresh_token);
      console.log('\nSigned in. Token stored (encrypted for your Windows user) in ' + TOKEN_FILE);
      return;
    }
    const err = t.json.error;
    if (err === 'authorization_pending') continue;
    if (err === 'slow_down') { wait += 5000; continue; }
    throw new Error(`Sign-in failed: ${err}: ${t.json.error_description || ''}`);
  }
  throw new Error('Sign-in timed out. Run login again.');
}

let accessToken = null;
let accessTokenExp = 0;
let needsLogin = false;

async function getAccessToken(cfg) {
  if (accessToken && Date.now() < accessTokenExp - 60_000) return accessToken;
  const rt = loadRefreshToken();
  if (!rt) { needsLogin = true; throw new Error('Not signed in'); }
  const t = await postForm(`${authBase(cfg.tenant)}/token`, {
    grant_type: 'refresh_token', client_id: cfg.clientId, refresh_token: rt, scope: SCOPES,
  });
  if (!t.ok) {
    if (t.json.error === 'invalid_grant' || t.json.error === 'interaction_required') needsLogin = true;
    throw new Error('Token refresh failed: ' + (t.json.error || t.status));
  }
  needsLogin = false;
  accessToken = t.json.access_token;
  accessTokenExp = Date.now() + (t.json.expires_in || 3600) * 1000;
  if (t.json.refresh_token && t.json.refresh_token !== rt) saveRefreshToken(t.json.refresh_token);
  return accessToken;
}

// ---------- Graph ----------
const MEETING_LINK_RE = /https:\/\/(?:teams\.microsoft\.com\/l\/meetup-join\/[^\s<>"]+|teams\.live\.com\/meet\/[^\s<>"]+|[\w.-]*zoom\.us\/[jw]\/[^\s<>"]+|meet\.google\.com\/[a-z0-9-]+|[\w.-]+\.webex\.com\/[^\s<>"]+)/i;

function toIsoUtc(dt) {
  // Graph returns "2026-11-17T14:00:00.0000000" when Prefer: outlook.timezone="UTC"
  if (!dt || !dt.dateTime) return null;
  const s = dt.dateTime.replace(/(\.\d{3})\d*$/, '$1');
  return /Z$|[+-]\d\d:\d\d$/.test(s) ? s : s + 'Z';
}

async function graphGet(cfg, url) {
  let token = await getAccessToken(cfg);
  let res = await fetch(url, {
    headers: {
      Authorization: 'Bearer ' + token,
      Prefer: 'outlook.timezone="UTC", outlook.body-content-type="text"',
    },
  });
  if (res.status === 401) {
    accessToken = null;
    token = await getAccessToken(cfg);
    res = await fetch(url, { headers: { Authorization: 'Bearer ' + token, Prefer: 'outlook.timezone="UTC", outlook.body-content-type="text"' } });
  }
  if (!res.ok) throw new Error(`Graph ${res.status}: ${(await res.text()).slice(0, 300)}`);
  return res.json();
}

async function fetchEvents(cfg) {
  const start = new Date(); start.setHours(0, 0, 0, 0);
  const end = new Date(start); end.setDate(end.getDate() + 2); // today + tomorrow (for "next up tomorrow")
  const select = 'id,subject,start,end,isAllDay,location,isOnlineMeeting,onlineMeeting,onlineMeetingUrl,showAs,isCancelled,organizer,responseStatus,body,sensitivity';
  let url = `https://graph.microsoft.com/v1.0/me/calendarView?startDateTime=${start.toISOString()}&endDateTime=${end.toISOString()}` +
    `&$select=${select}&$orderby=start/dateTime&$top=100`;
  const events = [];
  while (url && events.length < 300) {
    const page = await graphGet(cfg, url);
    for (const e of page.value || []) {
      const bodyLink = e.body && e.body.content ? (e.body.content.match(MEETING_LINK_RE) || [])[0] : null;
      const isPrivate = e.sensitivity === 'private' || e.sensitivity === 'confidential';
      events.push({
        id: crypto.createHash('sha1').update(e.id).digest('hex').slice(0, 16),
        subject: isPrivate ? 'Private appointment' : (e.subject || '(No title)'),
        start: toIsoUtc(e.start),
        end: toIsoUtc(e.end),
        isAllDay: !!e.isAllDay,
        location: isPrivate ? '' : (e.location && e.location.displayName) || '',
        organizer: (e.organizer && e.organizer.emailAddress && e.organizer.emailAddress.name) || '',
        joinUrl: (e.onlineMeeting && e.onlineMeeting.joinUrl) || e.onlineMeetingUrl || bodyLink || '',
        isOnline: !!e.isOnlineMeeting || !!bodyLink,
        showAs: e.showAs || 'busy',
        isCancelled: !!e.isCancelled,
        response: (e.responseStatus && e.responseStatus.response) || 'none',
      });
    }
    url = page['@odata.nextLink'] || null;
  }
  return events;
}

async function fetchMe(cfg) {
  const me = await graphGet(cfg, 'https://graph.microsoft.com/v1.0/me?$select=displayName,givenName');
  return { name: me.givenName || me.displayName || '' };
}

// ---------- cache + server ----------
const cache = { status: 'starting', events: [], user: null, updated: null, error: null, refreshSeconds: GRAPH_REFRESH_MS / 1000 };
let refreshing = null;

async function refresh(cfg) {
  if (refreshing) return refreshing;
  refreshing = (async () => {
    try {
      if (!cache.user) cache.user = await fetchMe(cfg);
      cache.events = await fetchEvents(cfg);
      cache.updated = new Date().toISOString();
      cache.status = 'ok';
      cache.error = null;
    } catch (err) {
      cache.status = needsLogin ? 'needs_login' : 'error';
      cache.error = String(err.message || err);
      log('Refresh failed:', cache.error);
    } finally {
      refreshing = null;
    }
  })();
  return refreshing;
}

function serve(cfg) {
  const allowedHosts = new Set([`localhost:${PORT}`, `127.0.0.1:${PORT}`, `[::1]:${PORT}`]);

  const handler = async (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Headers', 'X-Calendar-Key');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Cache-Control', 'no-store');
    if (!allowedHosts.has(String(req.headers.host || '').toLowerCase())) { res.writeHead(421); return res.end(); }
    if (req.method === 'OPTIONS') { res.writeHead(204); return res.end(); }

    const url = new URL(req.url, `http://localhost:${PORT}`);
    const json = (code, obj) => { res.writeHead(code, { 'Content-Type': 'application/json' }); res.end(JSON.stringify(obj)); };

    if (url.pathname === '/health') return json(200, { ok: true, signedIn: cache.status === 'ok' });

    if (url.pathname === '/api/events') {
      const key = req.headers['x-calendar-key'] || url.searchParams.get('key') || '';
      const a = Buffer.from(String(key)); const b = Buffer.from(cfg.key);
      if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return json(401, { status: 'bad_key' });
      const stale = !cache.updated || Date.now() - Date.parse(cache.updated) > GRAPH_REFRESH_MS;
      if (stale) await refresh(cfg);
      return json(200, cache);
    }
    json(404, { error: 'not found' });
  };

  const listen = (host) => {
    const srv = http.createServer(handler);
    srv.on('error', (e) => {
      if (e.code === 'EADDRINUSE') { log(`Port ${PORT} already in use on ${host} - is the helper already running?`); if (host === '127.0.0.1') process.exit(1); }
      else if (host !== '::1') log('Server error:', e.message);
    });
    srv.listen(PORT, host, () => log(`Listening on http://${host.includes(':') ? `[${host}]` : host}:${PORT}`));
  };
  listen('127.0.0.1');
  listen('::1');

  refresh(cfg);
  setInterval(() => refresh(cfg), GRAPH_REFRESH_MS);
  if (!loadRefreshToken()) log('Not signed in yet. Run: node teams-calendar-helper.mjs login');
  log(`Widget access key: ${cfg.key}`);
}

// ---------- CLI ----------
function argValue(name) {
  const i = process.argv.indexOf(name);
  return i > -1 ? process.argv[i + 1] : undefined;
}

const cmd = (process.argv[2] && !process.argv[2].startsWith('--')) ? process.argv[2] : 'serve';
const cfg = loadConfig();

try {
  if (cmd === 'login') {
    const tenant = argValue('--tenant'); const clientId = argValue('--client-id');
    if (tenant) cfg.tenant = tenant;
    if (clientId) cfg.clientId = clientId;
    saveConfig(cfg);
    await login(cfg);
    console.log(`\nWidget access key: ${cfg.key}\nPaste it into the widget's "Helper Access Key" setting in iCUE.`);
  } else if (cmd === 'key') {
    console.log(cfg.key);
  } else if (cmd === 'logout') {
    try { fs.unlinkSync(TOKEN_FILE); } catch { /* none */ }
    console.log('Signed out (token removed).');
  } else if (cmd === 'serve') {
    serve(cfg);
  } else {
    console.log('Commands: login | serve | key | logout');
  }
} catch (err) {
  console.error(String(err.message || err));
  process.exit(1);
}
