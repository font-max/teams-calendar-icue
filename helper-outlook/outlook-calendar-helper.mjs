#!/usr/bin/env node
// Teams Calendar Helper v2 (classic Outlook)
// Reads today's meetings from classic Outlook on this PC (read-outlook-calendar.ps1) and serves them
// to the iCUE "Teams Calendar" widget on http://localhost:38127. You sign in to Microsoft 365 inside
// classic Outlook (email, password, 2FA); this helper never sees your credentials.
//
// Usage:
//   node outlook-calendar-helper.mjs serve   (default)
//   node outlook-calendar-helper.mjs key     (print the widget access key)
//   node outlook-calendar-helper.mjs test    (read the calendar once and print a summary)
//
// Uses the same port and access key as v1 (helper/teams-calendar-helper.mjs). Run only one of them.

import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { execFile } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const PORT = 38127;
const REFRESH_MS = 2 * 60 * 1000;
const HERE = path.dirname(fileURLToPath(import.meta.url));
const READER = path.join(HERE, 'read-outlook-calendar.ps1');

const DATA_DIR = path.join(process.env.LOCALAPPDATA || path.join(process.env.HOME || '.', '.local'), 'TeamsCalendarHelper');
const CONFIG_FILE = path.join(DATA_DIR, 'config.json');
const LOG_FILE = path.join(DATA_DIR, 'helper-outlook.log');
fs.mkdirSync(DATA_DIR, { recursive: true });

function log(...args) {
  const line = `[${new Date().toISOString()}] ${args.join(' ')}`;
  console.log(line);
  try {
    if (fs.existsSync(LOG_FILE) && fs.statSync(LOG_FILE).size > 1024 * 1024) fs.renameSync(LOG_FILE, LOG_FILE + '.old');
    fs.appendFileSync(LOG_FILE, line + '\n');
  } catch { /* ignore */ }
}

// Shared with v1 so the widget's access key keeps working whichever helper runs.
function loadConfig() {
  let cfg = {};
  try { cfg = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8')); } catch { /* first run */ }
  if (!cfg.key) {
    cfg.key = crypto.randomBytes(12).toString('hex');
    fs.writeFileSync(CONFIG_FILE, JSON.stringify(cfg, null, 2));
  }
  return cfg;
}

function readOutlook() {
  return new Promise((resolve) => {
    execFile('powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', READER],
      { windowsHide: true, timeout: 120_000, maxBuffer: 10 * 1024 * 1024, encoding: 'utf8' },
      (err, stdout, stderr) => {
        const text = String(stdout || '').trim();
        try {
          resolve(JSON.parse(text.slice(text.indexOf('{'))));
        } catch {
          resolve({ status: 'error', message: 'Could not read Outlook: ' + String((err && err.message) || stderr || 'no output').slice(0, 300) });
        }
      });
  });
}

const cache = { status: 'starting', message: 'Reading classic Outlook...', events: [], user: null, updated: null, refreshSeconds: REFRESH_MS / 1000 };
let refreshing = null;

async function refresh() {
  if (refreshing) return refreshing;
  refreshing = (async () => {
    const r = await readOutlook();
    if (r.status === 'ok') {
      cache.events = (r.events || []).map(({ key, ...e }) => ({
        id: crypto.createHash('sha1').update(String(key)).digest('hex').slice(0, 16), ...e,
      }));
      cache.user = r.user || null;
      cache.updated = new Date().toISOString();
      cache.status = 'ok';
      cache.message = null;
    } else {
      cache.status = r.status || 'error';
      cache.message = r.message || null;
      log('Refresh:', cache.status, cache.message || '');
    }
    refreshing = null;
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

    if (url.pathname === '/health') return json(200, { ok: true, source: 'outlook', signedIn: cache.status === 'ok' });

    if (url.pathname === '/api/events') {
      const key = req.headers['x-calendar-key'] || url.searchParams.get('key') || '';
      const a = Buffer.from(String(key)); const b = Buffer.from(cfg.key);
      if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return json(401, { status: 'bad_key' });
      // Serve whatever is cached; a slow Outlook read never blocks the widget.
      if (!cache.updated || Date.now() - Date.parse(cache.updated) > REFRESH_MS) refresh();
      // While a stale-but-good cache exists, report ok so the widget keeps showing meetings.
      const body = cache.updated && cache.status !== 'ok' && cache.status !== 'needs_login'
        ? { ...cache, status: 'ok' } : cache;
      return json(200, body);
    }
    json(404, { error: 'not found' });
  };

  const listen = (host) => {
    const srv = http.createServer(handler);
    srv.on('error', (e) => {
      if (e.code === 'EADDRINUSE') { log(`Port ${PORT} already in use on ${host} - is another Teams Calendar helper (v1 or v2) running?`); if (host === '127.0.0.1') process.exit(1); }
      else if (host !== '::1') log('Server error:', e.message);
    });
    srv.listen(PORT, host, () => log(`Listening on http://${host.includes(':') ? `[${host}]` : host}:${PORT} (source: classic Outlook)`));
  };
  listen('127.0.0.1');
  listen('::1');

  refresh();
  setInterval(refresh, REFRESH_MS);
  log(`Widget access key: ${cfg.key}`);
}

const cmd = (process.argv[2] && !process.argv[2].startsWith('--')) ? process.argv[2] : 'serve';
const cfg = loadConfig();

if (cmd === 'key') {
  console.log(cfg.key);
} else if (cmd === 'test') {
  const r = await readOutlook();
  if (r.status !== 'ok') { console.log(`${r.status}: ${r.message || ''}`); process.exit(1); }
  console.log(`Signed in as: ${(r.user && r.user.name) || '(unknown)'}  -  ${r.events.length} event(s) today/tomorrow`);
  for (const e of r.events) {
    console.log(`  ${e.isAllDay ? 'all day ' : new Date(e.start).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}  ${e.subject}${e.joinUrl ? '  [join link]' : ''}`);
  }
} else if (cmd === 'serve') {
  serve(cfg);
} else {
  console.log('Commands: serve | key | test');
}
