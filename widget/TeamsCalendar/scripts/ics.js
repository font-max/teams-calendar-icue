/* Teams Calendar - ICS reader (runs inside the widget).
   Reads an Outlook published calendar and returns the meetings in a time window,
   expanding repeating meetings, time zones, exceptions and moved occurrences.
   Exposes TeamsIcs.eventsInWindow(icsText, windowStartMs, windowEndMs). */
(function (root) {
'use strict';
var DAY = 86400000;

// Short stable id for an occurrence (FNV-1a, hex).
function hashId(s) {
  var h1 = 0x811c9dc5, h2 = 0x01000193;
  for (var i = 0; i < s.length; i++) {
    var c = s.charCodeAt(i);
    h1 = Math.imul(h1 ^ c, 16777619) >>> 0;
    h2 = Math.imul(h2 ^ c, 2246822519) >>> 0;
  }
  return ('0000000' + h1.toString(16)).slice(-8) + ('0000000' + h2.toString(16)).slice(-8);
}

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
function parseIcs(text) {
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
  try { new Intl.DateTimeFormat('en-US', { timeZone: name }); } catch (e) { return null; }
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
  const wkst = rule.WKST in WD ? WD[rule.WKST] : 1;

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

function eventsInWindow(icsText, winStart, winEnd) {
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
      id: hashId(key),
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

root.TeamsIcs = { eventsInWindow: eventsInWindow, parseIcs: parseIcs };
})(typeof window !== 'undefined' ? window : globalThis);
