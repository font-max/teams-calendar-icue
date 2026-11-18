#!/usr/bin/env node
// Teams Calendar Helper v2 (published ICS link)
// Downloads your published Outlook calendar (ICS link from new Outlook / Outlook on the web),
// expands repeating meetings for today and tomorrow, and serves them to the iCUE "Teams Calendar"
// widget on http://localhost:38127. Works with new Outlook; no app approval, no classic Outlook.
//
// Usage:
//   node ics-calendar-helper.mjs set-url <ics-link>   (stored encrypted for your Windows user)
//   node ics-calendar-helper.mjs test                 (download once and list today/tomorrow)
//   node ics-calendar-helper.mjs serve                (default)
//   node ics-calendar-helper.mjs key                  (print the widget access key)
//
// Uses the same port and access key as the other helpers. Run only one of them at a time.

import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';

const PORT = Number(process.env.TEAMS_CAL_PORT) || 38127; // override only for testing
const REFRESH_MS = 5 * 60 * 1000;
const DAY = 86400000;

const DATA_DIR = path.join(process.env.LOCALAPPDATA || path.join(process.env.HOME || '.', '.local'), 'TeamsCalendarHelper');
const CONFIG_FILE = path.join(DATA_DIR, 'config.json');
const URL_FILE = path.join(DATA_DIR, 'ics-url.dat');
const LOG_FILE = path.join(DATA_DIR, 'helper-ics.log');

// =====================================================================================
// ICS parsing
// =====================================================================================

function unfold(text) {
  return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\n[ \t]/g, '').split('\n');
}

function unescapeText(v) {
  return v.replace(/\\([\\;,nN])/g, (_, c) => (c === 'n' || c === 'N' ? '\n' : c));
}

function parseLine(line) {
  // NAME;PARAM=a;PARAM2="b:c":VALUE  (colons inside quoted params are allowed)
  let i = 0, inQ = false;
  for (; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') inQ = !inQ;
    else if (ch === ':' && !inQ) break;
  }
  const head = line.slice(0, i), value = line.slice(i + 1);
  const parts = head.match(/(?:[^;"]|"[^"]*")+/g) || [head];
  const name = parts[0].toUpperCase();
  const params = {};
  for (const p of parts.slice(1)) {
    const eq = p.indexOf('=');
    if (eq > 0) params[p.slice(0, eq).toUpperCase()] = p.slice(eq + 1).replace(/^"|"$/g, '');
  }
  return { name, params, value };
}

// Returns a tree: { type, props: {NAME: [{params, value}]}, children: [] }
export function parseIcs(text) {
  const root = { type: 'ROOT', props: {}, children: [] };
  const stack = [root];
  for (const raw of unfold(text)) {
    if (!raw.trim()) continue;
    const { name, params, value } = parseLine(raw);
    const top = stack[stack.length - 1];
    if (name === 'BEGIN') {
      const node = { type: value.toUpperCase(), props: {}, children: [] };
      top.children.push(node);
      stack.push(node);
    } else if (name === 'END') {
      if (stack.length > 1) stack.pop();
    } else {
      (top.props[name] = top.props[name] || []).push({ params, value });
    }
  }
  return root;
}

const prop = (node, name) => (node.props[name] && node.props[name][0]) || null;
const propVal = (node, name) => { const p = prop(node, name); return p ? p.value : ''; };

// =====================================================================================
// Time zones
// =====================================================================================
// Times are handled as "naive" milliseconds: wall-clock components stored via Date.UTC().
// A zone converts naive wall-clock -> real UTC.

const WINDOWS_TO_IANA = {
  'Dateline Standard Time': 'Etc/GMT+12', 'Hawaiian Standard Time': 'Pacific/Honolulu', 'Alaskan Standard Time': 'America/Anchorage',
  'Pacific Standard Time': 'America/Los_Angeles', 'US Mountain Standard Time': 'America/Phoenix', 'Mountain Standard Time': 'America/Denver',
  'Central Standard Time': 'America/Chicago', 'Eastern Standard Time': 'America/New_York', 'Atlantic Standard Time': 'America/Halifax',
  'Newfoundland Standard Time': 'America/St_Johns', 'E. South America Standard Time': 'America/Sao_Paulo', 'SA Pacific Standard Time': 'America/Bogota',
  'Central Standard Time (Mexico)': 'America/Mexico_City', 'Canada Central Standard Time': 'America/Regina',
  'UTC': 'Etc/UTC', 'GMT Standard Time': 'Europe/London', 'Greenwich Standard Time': 'Atlantic/Reykjavik',
  'W. Europe Standard Time': 'Europe/Berlin', 'Central Europe Standard Time': 'Europe/Budapest', 'Romance Standard Time': 'Europe/Paris',
  'Central European Standard Time': 'Europe/Warsaw', 'E. Europe Standard Time': 'Europe/Chisinau', 'FLE Standard Time': 'Europe/Kiev',
  'GTB Standard Time': 'Europe/Bucharest', 'Russian Standard Time': 'Europe/Moscow', 'Israel Standard Time': 'Asia/Jerusalem',
  'South Africa Standard Time': 'Africa/Johannesburg', 'Arabian Standard Time': 'Asia/Dubai', 'India Standard Time': 'Asia/Kolkata',
  'China Standard Time': 'Asia/Shanghai', 'Singapore Standard Time': 'Asia/Singapore', 'Tokyo Standard Time': 'Asia/Tokyo',
  'Korea Standard Time': 'Asia/Seoul', 'AUS Eastern Standard Time': 'Australia/Sydney', 'E. Australia Standard Time': 'Australia/Brisbane',
  'New Zealand Standard Time': 'Pacific/Auckland',
};

function parseOffset(s) { // "+0100" / "-0700"
  const m = /^([+-])(\d\d)(\d\d)(\d\d)?$/.exec(String(s).trim());
  if (!m) return 0;
  const v = (Number(m[2]) * 3600 + Number(m[3]) * 60 + Number(m[4] || 0)) * 1000;
  return m[1] === '-' ? -v : v;
}

const WD = { SU: 0, MO: 1, TU: 2, WE: 3, TH: 4, FR: 5, SA: 6 };

function nthWeekdayOfMonth(year, month0, weekday, n) { // n: 1..5 or -1..-5, returns day of month or null
  if (n > 0) {
    const first = new Date(Date.UTC(year, month0, 1)).getUTCDay();
    const day = 1 + ((weekday - first + 7) % 7) + (n - 1) * 7;
    return day <= daysInMonth(year, month0) ? day : null;
  }
  const dim = daysInMonth(year, month0);
  const last = new Date(Date.UTC(year, month0, dim)).getUTCDay();
  const day = dim - ((last - weekday + 7) % 7) + (n + 1) * 7;
  return day >= 1 ? day : null;
}
function daysInMonth(y, m0) { return new Date(Date.UTC(y, m0 + 1, 0)).getUTCDate(); }

// Builds a zone from a VTIMEZONE node: list of {start naive, rrule, offsetTo}
function zoneFromVtimezone(vtz) {
  const rules = vtz.children.filter((c) => c.type === 'STANDARD' || c.type === 'DAYLIGHT').map((c) => {
    const dt = parseDateValue(propVal(c, 'DTSTART'));
    return { start: dt ? dt.naive : 0, rrule: parseRrule(propVal(c, 'RRULE')), to: parseOffset(propVal(c, 'TZOFFSETTO')) };
  });
  if (!rules.length) return null;
  return (naive) => {
    const y = new Date(naive).getUTCFullYear();
    let best = null;
    for (const r of rules) {
      const cands = [];
      if (r.rrule && r.rrule.FREQ === 'YEARLY') {
        const sd = new Date(r.start);
        const months = r.rrule.BYMONTH ? r.rrule.BYMONTH.split(',').map(Number) : [sd.getUTCMonth() + 1];
        for (const yy of [y - 1, y]) {
          for (const mo of months) {
            let day = sd.getUTCDate();
            if (r.rrule.BYDAY) {
              const m = /^([+-]?\d)?(SU|MO|TU|WE|TH|FR|SA)$/.exec(r.rrule.BYDAY.split(',')[0]);
              if (m) day = nthWeekdayOfMonth(yy, mo - 1, WD[m[2]], Number(m[1] || 1));
            }
            if (day) cands.push(Date.UTC(yy, mo - 1, day, sd.getUTCHours(), sd.getUTCMinutes()));
          }
        }
      } else {
        cands.push(r.start);
      }
      for (const t of cands) if (t <= naive && t >= r.start - 366 * DAY && (!best || t > best.t)) best = { t, to: r.to };
    }
    if (!best) { // before all transitions: use the earliest rule's offset
      const first = rules.slice().sort((a, b) => a.start - b.start)[0];
      return naive - first.to;
    }
    return naive - best.to;
  };
}

function zoneFromIana(name) {
  try { new Intl.DateTimeFormat('en-US', { timeZone: name }); } catch { return null; }
  const fmt = new Intl.DateTimeFormat('en-US', {
    timeZone: name, hourCycle: 'h23', year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: 'numeric', second: 'numeric',
  });
  const offsetAt = (utc) => {
    const p = Object.fromEntries(fmt.formatToParts(new Date(utc)).map((x) => [x.type, x.value]));
    return Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour % 24, +p.minute, +p.second) - utc;
  };
  return (naive) => { let utc = naive - offsetAt(naive); utc = naive - offsetAt(utc); return utc; };
}

const localZone = (naive) => {
  const d = new Date(naive);
  return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds()).getTime();
};

function buildZones(root) {
  const zones = {};
  for (const c of root.children.flatMap((n) => [n, ...n.children])) {
    if (c.type === 'VTIMEZONE') {
      const id = propVal(c, 'TZID');
      const z = zoneFromVtimezone(c);
      if (id && z) zones[id] = z;
    }
  }
  return (tzid) => {
    if (!tzid) return localZone;
    if (zones[tzid]) return zones[tzid];
    const iana = WINDOWS_TO_IANA[tzid] || tzid;
    return (zones[tzid] = zoneFromIana(iana) || localZone);
  };
}

// =====================================================================================
// Dates and recurrence
// =====================================================================================

// "20261117", "20261117T090000", "20261117T090000Z" -> { naive, isDate, isUtc }
function parseDateValue(v) {
  const m = /^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?(Z)?)?$/.exec(String(v).trim());
  if (!m) return null;
  return {
    naive: Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)),
    isDate: !m[4],
    isUtc: !!m[7],
  };
}

function toUtc(p, getZone) {
  // p: {value, params}; returns { utc, isDate, naive }
  const d = parseDateValue(p.value);
  if (!d) return null;
  if (d.isDate) return { utc: d.naive, naive: d.naive, isDate: true };
  if (d.isUtc) return { utc: d.naive, naive: d.naive, isDate: false, isUtc: true };
  return { utc: getZone(p.params.TZID)(d.naive), naive: d.naive, isDate: false };
}

function parseRrule(v) {
  if (!v) return null;
  const r = {};
  for (const part of v.split(';')) { const [k, val] = part.split('='); if (k) r[k.toUpperCase()] = val; }
  return r.FREQ ? r : null;
}

function parseDuration(v) { // P1D, PT1H30M, P1W
  const m = /^([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/.exec(String(v).trim());
  if (!m) return null;
  const ms = ((+(m[2] || 0) * 7 + +(m[3] || 0)) * 86400 + +(m[4] || 0) * 3600 + +(m[5] || 0) * 60 + +(m[6] || 0)) * 1000;
  return m[1] === '-' ? -ms : ms;
}

// Generates naive start times of a recurring event (in its own wall clock), in order, until naive >= limitNaive.
function* expandRrule(rule, dtNaive, limitNaive) {
  const freq = rule.FREQ;
  const interval = Math.max(1, Number(rule.INTERVAL || 1));
  const count = rule.COUNT ? Number(rule.COUNT) : Infinity;
  const until = rule.UNTIL ? parseDateValue(rule.UNTIL) : null;
  const start = new Date(dtNaive);
  const h = start.getUTCHours(), mi = start.getUTCMinutes(), s = start.getUTCSeconds();
  const byDay = rule.BYDAY ? rule.BYDAY.split(',').map((x) => {
    const m = /^([+-]?\d{1,2})?(SU|MO|TU|WE|TH|FR|SA)$/.exec(x.trim());
    return m ? { n: m[1] ? Number(m[1]) : 0, wd: WD[m[2]] } : null;
  }).filter(Boolean) : null;
  const byMonthDay = rule.BYMONTHDAY ? rule.BYMONTHDAY.split(',').map(Number) : null;
  const byMonth = rule.BYMONTH ? rule.BYMONTH.split(',').map(Number) : null;
  const bySetPos = rule.BYSETPOS ? rule.BYSETPOS.split(',').map(Number) : null;
  const wkst = WD[rule.WKST] ?? 1;

  let emitted = 0;
  const at = (y, m0, d) => Date.UTC(y, m0, d, h, mi, s);
  const applySetPos = (list) => {
    if (!bySetPos) return list;
    return bySetPos.map((p) => (p > 0 ? list[p - 1] : list[list.length + p])).filter((x) => x !== undefined).sort((a, b) => a - b);
  };
  const monthDays = (y, m0) => { // candidate days within a month for MONTHLY/YEARLY
    const dim = daysInMonth(y, m0);
    let days = [];
    if (byMonthDay) days = byMonthDay.map((d) => (d > 0 ? d : dim + d + 1)).filter((d) => d >= 1 && d <= dim);
    else if (byDay) {
      for (const b of byDay) {
        if (b.n) { const d = nthWeekdayOfMonth(y, m0, b.wd, b.n); if (d) days.push(d); }
        else for (let d = 1; d <= dim; d++) if (new Date(Date.UTC(y, m0, d)).getUTCDay() === b.wd) days.push(d);
      }
    } else if (start.getUTCDate() <= dim) days = [start.getUTCDate()];
    days = [...new Set(days)].sort((a, b) => a - b);
    return applySetPos(days.map((d) => at(y, m0, d)));
  };

  let guard = 0;
  for (let period = 0; guard++ < 20000; period++) {
    let cands = [];
    if (freq === 'DAILY') {
      const t = at(start.getUTCFullYear(), start.getUTCMonth(), start.getUTCDate() + period * interval);
      const wdOk = !byDay || byDay.some((b) => b.wd === new Date(t).getUTCDay());
      const mOk = !byMonth || byMonth.includes(new Date(t).getUTCMonth() + 1);
      if (wdOk && mOk) cands = [t];
    } else if (freq === 'WEEKLY') {
      const sd = start.getUTCDay();
      const weekStart = at(start.getUTCFullYear(), start.getUTCMonth(), start.getUTCDate() - ((sd - wkst + 7) % 7) + period * 7 * interval);
      const days = byDay ? byDay.map((b) => b.wd) : [sd];
      cands = days.map((wd) => weekStart + ((wd - wkst + 7) % 7) * DAY).sort((a, b) => a - b);
      cands = applySetPos(cands);
    } else if (freq === 'MONTHLY') {
      const total = start.getUTCMonth() + period * interval;
      const y = start.getUTCFullYear() + Math.floor(total / 12), m0 = total % 12;
      if (!byMonth || byMonth.includes(m0 + 1)) cands = monthDays(y, m0);
    } else if (freq === 'YEARLY') {
      const y = start.getUTCFullYear() + period * interval;
      const months = byMonth || [start.getUTCMonth() + 1];
      cands = months.flatMap((mo) => {
        if (!byDay && !byMonthDay) {
          const d = start.getUTCDate();
          return d <= daysInMonth(y, mo - 1) ? [at(y, mo - 1, d)] : [];
        }
        return monthDays(y, mo - 1);
      });
    } else {
      return;
    }

    for (const t of cands) {
      if (t < dtNaive) continue;
      if (until && (until.isUtc ? null : t > until.naive + (until.isDate ? DAY - 1 : 0))) return;
      if (emitted >= count) return;
      if (t >= limitNaive) return;
      emitted++;
      yield t;
    }
    // stop once the period itself starts past the limit (no later candidates can fall inside it)
    if (cands.length && cands[cands.length - 1] >= limitNaive) return;
    let periodStart;
    if (freq === 'DAILY') periodStart = at(start.getUTCFullYear(), start.getUTCMonth(), start.getUTCDate() + period * interval);
    else if (freq === 'WEEKLY') periodStart = Date.UTC(start.getUTCFullYear(), start.getUTCMonth(), start.getUTCDate() + (period * interval - 1) * 7);
    else if (freq === 'MONTHLY') { const tot = start.getUTCMonth() + period * interval; periodStart = Date.UTC(start.getUTCFullYear() + Math.floor(tot / 12), tot % 12, 1); }
    else periodStart = Date.UTC(start.getUTCFullYear() + period * interval, 0, 1);
    if (periodStart >= limitNaive) return;
  }
}

// =====================================================================================
// Events for a window
// =====================================================================================

const MEETING_LINK_RE = /https:\/\/(?:teams\.microsoft\.com\/l\/meetup-join\/[^\s<>"]+|teams\.live\.com\/meet\/[^\s<>"]+|[\w.-]*zoom\.us\/[jw]\/[^\s<>"]+|meet\.google\.com\/[a-z0-9-]+|[\w.-]+\.webex\.com\/[^\s<>"]+)/i;

const iso = (ms) => new Date(ms).toISOString().replace(/\.\d{3}Z$/, 'Z');

export function eventsInWindow(icsText, winStart, winEnd) {
  const root = parseIcs(icsText);
  const cal = root.children.find((c) => c.type === 'VCALENDAR') || root;
  const getZone = buildZones(root);
  const vevents = cal.children.filter((c) => c.type === 'VEVENT');

  // Overrides of single occurrences: UID -> Map(original start utc -> VEVENT)
  const overrides = new Map();
  for (const ev of vevents) {
    const rid = prop(ev, 'RECURRENCE-ID');
    if (!rid) continue;
    const t = toUtc(rid, getZone);
    if (!t) continue;
    const uid = propVal(ev, 'UID');
    if (!overrides.has(uid)) overrides.set(uid, new Map());
    overrides.get(uid).set(t.utc, ev);
  }

  const out = [];
  const pushInstance = (ev, startUtc, endUtc, isAllDay, key) => {
    if (!(startUtc < winEnd && endUtc > winStart)) return;
    const status = propVal(ev, 'STATUS').toUpperCase();
    const cls = propVal(ev, 'CLASS').toUpperCase();
    const isPrivate = cls === 'PRIVATE' || cls === 'CONFIDENTIAL';
    const busy = (propVal(ev, 'X-MICROSOFT-CDO-BUSYSTATUS') || propVal(ev, 'TRANSP') || 'BUSY').toUpperCase();
    const desc = unescapeText(propVal(ev, 'DESCRIPTION'));
    const urlProp = propVal(ev, 'URL');
    const link = (desc.match(MEETING_LINK_RE) || [])[0] || (MEETING_LINK_RE.test(urlProp) ? urlProp : '');
    const org = prop(ev, 'ORGANIZER');
    const subj = unescapeText(propVal(ev, 'SUMMARY'));
    out.push({
      id: crypto.createHash('sha1').update(key).digest('hex').slice(0, 16),
      subject: isPrivate ? 'Private appointment' : (subj || '(No title)'),
      start: iso(startUtc),
      end: iso(endUtc),
      isAllDay,
      location: isPrivate ? '' : unescapeText(propVal(ev, 'LOCATION')),
      organizer: org ? (org.params.CN || '').replace(/^"|"$/g, '') : '',
      joinUrl: isPrivate ? '' : link,
      isOnline: !!link,
      showAs: { FREE: 'free', TRANSPARENT: 'free', TENTATIVE: 'tentative', OOF: 'oof', WORKINGELSEWHERE: 'workingElsewhere' }[busy] || 'busy',
      isCancelled: status === 'CANCELLED',
      response: 'none',
    });
  };

  // All-day events are reported as "UTC midnight of the local date" (the widget expects that).
  const allDayIso = (naive) => naive;

  for (const ev of vevents) {
    if (prop(ev, 'RECURRENCE-ID')) continue; // handled via overrides
    const dtp = prop(ev, 'DTSTART');
    if (!dtp) continue;
    const st = toUtc(dtp, getZone);
    if (!st) continue;
    const isAllDay = st.isDate || propVal(ev, 'X-MICROSOFT-CDO-ALLDAYEVENT').toUpperCase() === 'TRUE';
    const dte = prop(ev, 'DTEND');
    let durNaive;
    if (dte) { const e = parseDateValue(dte.value); durNaive = e ? e.naive - st.naive : 0; }
    else { const d = parseDuration(propVal(ev, 'DURATION')); durNaive = d != null ? d : (st.isDate ? DAY : 0); }
    if (durNaive < 0) durNaive = 0;

    const uid = propVal(ev, 'UID');
    const zone = st.isDate ? ((n) => n) : st.isUtc ? ((n) => n) : getZone(dtp.params.TZID);
    const ovr = overrides.get(uid);
    const rule = parseRrule(propVal(ev, 'RRULE'));

    if (!rule) {
      const s = isAllDay ? allDayIso(st.naive) : st.utc;
      pushInstance(ev, s, s + (isAllDay ? Math.max(DAY, durNaive) : durNaive), isAllDay, uid + '|' + s);
      continue;
    }

    const exdates = new Set();
    for (const ex of ev.props.EXDATE || []) {
      for (const v of ex.value.split(',')) {
        const t = toUtc({ value: v, params: ex.params }, getZone);
        if (t) exdates.add(isAllDay ? t.naive : t.utc);
      }
    }
    const untilD = rule.UNTIL ? parseDateValue(rule.UNTIL) : null;
    // Expand in the event's wall clock up to the window end (+1 day margin for zone offsets)
    const limitNaive = winEnd + 2 * DAY;
    for (const n of expandRrule(rule, st.naive, limitNaive)) {
      const s = isAllDay ? n : zone(n);
      if (untilD && untilD.isUtc && !isAllDay && s > untilD.naive) break;
      if (exdates.has(s)) continue;
      const o = ovr && ovr.get(s);
      if (o) continue; // emitted below from the override itself
      pushInstance(ev, s, s + (isAllDay ? Math.max(DAY, durNaive) : (zone(n + durNaive) - s)), isAllDay, uid + '|' + s);
    }
  }

  // Emit overrides (moved / edited single occurrences)
  for (const [uid, map] of overrides) {
    for (const [origUtc, ev] of map) {
      const st = toUtc(prop(ev, 'DTSTART'), getZone);
      if (!st) continue;
      const isAllDay = st.isDate;
      const dte = prop(ev, 'DTEND');
      let end;
      if (dte) { const e = toUtc(dte, getZone); end = e ? (isAllDay ? e.naive : e.utc) : null; }
      if (end == null) { const d = parseDuration(propVal(ev, 'DURATION')); end = (isAllDay ? st.naive : st.utc) + (d != null ? d : 0); }
      const s = isAllDay ? st.naive : st.utc;
      pushInstance(ev, s, end, isAllDay, uid + '|' + origUtc);
    }
  }

  return out.sort((a, b) => a.start.localeCompare(b.start));
}

// =====================================================================================
// Helper service
// =====================================================================================

function log(...args) {
  const line = `[${new Date().toISOString()}] ${args.join(' ')}`;
  console.log(line);
  try {
    if (fs.existsSync(LOG_FILE) && fs.statSync(LOG_FILE).size > 1024 * 1024) fs.renameSync(LOG_FILE, LOG_FILE + '.old');
    fs.appendFileSync(LOG_FILE, line + '\n');
  } catch { /* ignore */ }
}

function loadConfig() {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  let cfg = {};
  try { cfg = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8')); } catch { /* first run */ }
  if (!cfg.key) { cfg.key = crypto.randomBytes(12).toString('hex'); fs.writeFileSync(CONFIG_FILE, JSON.stringify(cfg, null, 2)); }
  return cfg;
}

// The ICS link is a secret (anyone with it can read the calendar), so it is stored with Windows DPAPI.
function dpapi(mode, input) {
  const script =
    "$ErrorActionPreference='Stop';Add-Type -AssemblyName System.Security;$in=[Console]::In.ReadToEnd().Trim();" +
    (mode === 'protect'
      ? "$b=[Text.Encoding]::UTF8.GetBytes($in);[Convert]::ToBase64String([Security.Cryptography.ProtectedData]::Protect($b,$null,'CurrentUser'))"
      : "$b=[Convert]::FromBase64String($in);[Text.Encoding]::UTF8.GetString([Security.Cryptography.ProtectedData]::Unprotect($b,$null,'CurrentUser'))");
  const r = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], { input, encoding: 'utf8', windowsHide: true });
  if (r.status !== 0) throw new Error('DPAPI ' + mode + ' failed');
  return r.stdout.trim();
}
function saveUrl(u) { fs.writeFileSync(URL_FILE, dpapi('protect', u)); }
function loadUrl() { try { return dpapi('unprotect', fs.readFileSync(URL_FILE, 'utf8')); } catch { return null; } }

function todayWindow() {
  const s = new Date(); s.setHours(0, 0, 0, 0);
  const e = new Date(s); e.setDate(e.getDate() + 2);
  return [s.getTime(), e.getTime()];
}

async function downloadIcs(url) {
  const res = await fetch(url.replace(/^webcal:/i, 'https:'), { headers: { Accept: 'text/calendar' }, redirect: 'follow' });
  if (!res.ok) throw new Error(`ICS download failed: HTTP ${res.status}`);
  const text = await res.text();
  if (!/BEGIN:VCALENDAR/i.test(text)) throw new Error('The link did not return a calendar (.ics). Copy the ICS link, not the HTML link.');
  return text;
}

const cache = { status: 'starting', message: 'Downloading your calendar...', events: [], user: null, updated: null, refreshSeconds: REFRESH_MS / 1000 };
let refreshing = null;

async function refresh() {
  if (refreshing) return refreshing;
  refreshing = (async () => {
    const url = loadUrl();
    if (!url) {
      cache.status = 'needs_login';
      cache.message = 'No calendar link saved yet. Run Setup on this PC to connect your Outlook calendar.';
    } else {
      try {
        const [ws, we] = todayWindow();
        cache.events = eventsInWindow(await downloadIcs(url), ws, we);
        cache.updated = new Date().toISOString();
        cache.status = 'ok';
        cache.message = null;
      } catch (err) {
        cache.status = 'error';
        cache.message = String(err.message || err);
        log('Refresh failed:', cache.message);
      }
    }
    refreshing = null;
  })();
  return refreshing;
}

// The copy Setup installs lives in %LOCALAPPDATA%\TeamsCalendarHelper\app. Only that copy may take the port over,
// so a development copy started by hand can never knock out the installed one.
const SELF = fileURLToPath(import.meta.url);
const IS_INSTALLED = SELF.toLowerCase().startsWith(path.join(DATA_DIR, 'app').toLowerCase() + path.sep);

// If another Teams Calendar helper (an older version or a stray copy) holds the port, stop it.
// Never touches unrelated programs. Returns true if something was stopped.
function takeOverPort() {
  if (!IS_INSTALLED && process.env.TEAMS_CAL_TAKEOVER !== '1') return false;
  const ps = `$c = Get-NetTCPConnection -LocalPort ${PORT} -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1; ` +
    'if ($c) { $p = Get-CimInstance Win32_Process -Filter ("ProcessId=" + $c.OwningProcess); [string]$p.ProcessId + "`t" + $p.CommandLine }';
  const r = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', ps], { encoding: 'utf8', windowsHide: true, timeout: 30_000 });
  const line = String(r.stdout || '').trim();
  const tab = line.indexOf('\t');
  if (tab < 0) return false;
  const pid = Number(line.slice(0, tab));
  const cmd = line.slice(tab + 1);
  if (!pid || pid === process.pid) return false;
  if (!/calendar-helper\.mjs/i.test(cmd) || cmd.toLowerCase().includes(SELF.toLowerCase())) {
    log(`Port ${PORT} is used by another program (PID ${pid}); not touching it: ${cmd}`);
    return false;
  }
  try {
    process.kill(pid);
    log(`Stopped an older Teams Calendar helper on port ${PORT} (PID ${pid}): ${cmd}`);
    return true;
  } catch (err) {
    log(`Could not stop PID ${pid}: ${err.message}`);
    return false;
  }
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

    if (url.pathname === '/health') return json(200, { ok: true, source: 'ics', signedIn: cache.status === 'ok' });
    if (url.pathname === '/api/events') {
      const key = req.headers['x-calendar-key'] || url.searchParams.get('key') || '';
      const a = Buffer.from(String(key)); const b = Buffer.from(cfg.key);
      if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return json(401, { status: 'bad_key' });
      if (!cache.updated || Date.now() - Date.parse(cache.updated) > REFRESH_MS) await refresh();
      // keep showing last good data during a temporary download failure
      const body = cache.updated && cache.status === 'error' ? { ...cache, status: 'ok' } : cache;
      return json(200, body);
    }
    json(404, { error: 'not found' });
  };
  const listen = (host, attempt = 0) => {
    const srv = http.createServer(handler);
    srv.on('error', (e) => {
      if (e.code === 'EADDRINUSE') {
        // At boot an older Teams Calendar helper may have grabbed the port first: the installed helper takes over.
        // (::1 just retries: the 127.0.0.1 listener does the takeover and frees both.)
        if (attempt < 3 && (attempt > 0 || host === '::1' || takeOverPort())) {
          setTimeout(() => listen(host, attempt + 1), 1500);
          return;
        }
        log(`Port ${PORT} already in use on ${host} - is another Teams Calendar helper running?`);
        if (host === '127.0.0.1') process.exit(1);
      } else if (host !== '::1') log('Server error:', e.message);
    });
    srv.listen(PORT, host, () => log(`Listening on http://${host.includes(':') ? `[${host}]` : host}:${PORT} (source: published ICS)`));
  };
  listen('127.0.0.1');
  listen('::1');
  refresh();
  setInterval(refresh, REFRESH_MS);
  log(`Widget access key: ${cfg.key}`);
}

async function main() {
  const cmd = (process.argv[2] && !process.argv[2].startsWith('--')) ? process.argv[2] : 'serve';
  const cfg = loadConfig();
  if (cmd === 'set-url') {
    const u = (process.argv[3] || '').trim();
    if (!/^(https|webcal):\/\/\S+\.ics(\?\S*)?$/i.test(u)) {
      console.error('Usage: node ics-calendar-helper.mjs set-url "<ICS link ending in .ics>"');
      process.exit(1);
    }
    saveUrl(u);
    console.log('Calendar link saved (encrypted for your Windows user). Run "test" to check it.');
  } else if (cmd === 'test') {
    const u = loadUrl();
    if (!u) { console.error('No link set. Run set-url first.'); process.exit(1); }
    const [ws, we] = todayWindow();
    const events = eventsInWindow(await downloadIcs(u), ws, we);
    console.log(`${events.length} event(s) today/tomorrow:`);
    for (const e of events) {
      const when = e.isAllDay ? 'all day  ' : new Date(e.start).toLocaleString([], { weekday: 'short', hour: '2-digit', minute: '2-digit' });
      console.log(`  ${when}  ${e.subject}${e.joinUrl ? '  [join link]' : ''}${e.isCancelled ? '  (cancelled)' : ''}`);
    }
  } else if (cmd === 'key') {
    console.log(cfg.key);
  } else if (cmd === 'serve') {
    serve(cfg);
  } else {
    console.log('Commands: set-url <link> | test | serve | key');
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  main().catch((err) => { console.error(String(err.message || err)); process.exit(1); });
}
