/* Teams Calendar - iCUE widget runtime */

// iCUE API binding - bare assignment is intentional (see lifecycle docs).
icueEvents = {
  onDataUpdated: onIcueDataUpdated,
  onICUEInitialized: onIcueInitialized
};

// Outlook doesn't let web pages read published calendars directly (no CORS header), so the calendar is
// fetched through a tiny pass-through relay (source: /relay in the GitHub repo). It stores and logs nothing.
var RELAY_URL = 'https://teams-calendar-relay.fontoura.workers.dev/calendar';
var DEMO = /[?&]demo=1/.test(location.search);
var REFRESH_MS = 10 * 60000; // Outlook only republishes the calendar every few minutes anyway
var RETRY_MS = 2 * 60000;    // while something is wrong

// Same rule the relay applies: an Outlook published-calendar (.ics) link.
var LINK_RE = /^(?:https|webcal):\/\/outlook\.(?:office365|office|live)\.com\/owa\/calendar\/[^\/\s]+\/[^\/\s]+\/[^\/\s]+\.ics(?:[?#]\S*)?$/i;

var state = {
  payload: null,      // last good payload
  stale: false,
  problem: null,      // { title, body, kind }
  nextFetchAt: 0,     // when to download the calendar again
  fetching: false,
  language: 'en'
};

// ---------- property access ----------
function getIcueProperty(name) {
  if (typeof window !== 'undefined' && Object.prototype.hasOwnProperty.call(window, name)) {
    var v = window[name];
    if (v !== undefined && v !== null && v !== '') return v;
  }
  try {
    var v2 = Function('return typeof ' + name + ' !== "undefined" ? ' + name + ' : undefined')();
    if (v2 !== undefined && v2 !== null && v2 !== '') return v2;
  } catch (e) {}
  return undefined;
}
function prop(name, def) { var v = getIcueProperty(name); return v === undefined ? def : v; }
function boolProp(name, def) { var v = prop(name, def); return v === true || v === 'true'; }
function numProp(name, def, min, max) {
  var n = Number(prop(name, def));
  if (!isFinite(n)) n = def;
  return Math.max(min, Math.min(max, n));
}

function settings() {
  return {
    link: String(prop('calendarLink', '')).trim(),
    use24h: boolProp('use24h', false),
    hidePast: boolProp('hidePast', true),
    showAllDay: boolProp('showAllDay', true),
    showFree: boolProp('showFree', false),
    dayStart: numProp('dayStart', 8, 0, 12),
    dayEnd: numProp('dayEnd', 18, 13, 24)
  };
}

// ---------- styles ----------
function applyStyles() {
  var root = document.documentElement;
  var tc = prop('textColor', '#ffffff'), ac = prop('accentColor', '#7b83eb'), bg = prop('backgroundColor', '#000000');
  root.style.setProperty('--text-color', typeof tc === 'string' ? tc : '#ffffff');
  root.style.setProperty('--accent-color', typeof ac === 'string' ? ac : '#7b83eb');
  root.style.setProperty('--bg-color', typeof bg === 'string' ? bg : '#000000');
  root.style.setProperty('--bg-opacity', numProp('transparency', 100, 0, 100) / 100);
  root.style.setProperty('--on-accent', readableOn(typeof ac === 'string' ? ac : '#7b83eb'));
}

// Black or white, whichever contrasts more with the given color (WCAG relative luminance).
function readableOn(hex) {
  var m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})/i.exec(String(hex).trim());
  if (!m) return '#ffffff';
  var lum = [m[1], m[2], m[3]].map(function (h) {
    var c = parseInt(h, 16) / 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  });
  var L = 0.2126 * lum[0] + 0.7152 * lum[1] + 0.0722 * lum[2];
  return (L + 0.05) / 0.05 >= 1.05 / (L + 0.05) ? '#000000' : '#ffffff';
}

// ---------- storage (last known data) ----------
function storageKey() { return String(prop('uniqueId', 'teams-calendar-widget')); }
function saveCache(payload) {
  try {
    var all = JSON.parse(localStorage.getItem(storageKey()) || '{}');
    all.lastPayload = payload;
    localStorage.setItem(storageKey(), JSON.stringify(all));
  } catch (e) {}
}
function loadCache() {
  try { return JSON.parse(localStorage.getItem(storageKey()) || '{}').lastPayload || null; } catch (e) { return null; }
}

// ---------- links ----------
function openLink(url) {
  if (!url || !/^https:\/\//i.test(url)) return;
  try {
    if (window.plugins && window.plugins.Linkprovider && typeof pluginLinkprovider_initialized !== 'undefined' && pluginLinkprovider_initialized) {
      window.plugins.Linkprovider.open(url);
      return;
    }
  } catch (e) { console.warn('Link plugin failed', e); }
  window.open(url, '_blank');
}

// ---------- time helpers ----------
function fmtTime(d, s) {
  return d.toLocaleTimeString(state.language, { hour: 'numeric', minute: '2-digit', hour12: !s.use24h });
}
function fmtDuration(ms) {
  var mins = Math.max(0, Math.round(ms / 60000));
  if (mins < 60) return mins + ' min';
  var h = Math.floor(mins / 60), m = mins % 60;
  return m ? h + ' h ' + m + ' min' : h + ' h';
}
function startOfDay(d) { var x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function addDays(d, n) { var x = new Date(d); x.setDate(x.getDate() + n); return x; }

// ---------- data ----------
function normalize(payload, s) {
  var list = (payload && payload.events) || [];
  return list.map(function (e) {
    var start = new Date(e.start), end = new Date(e.end);
    if (e.isAllDay) {
      // All-day events come back as UTC midnight; treat their date as local.
      start = new Date(start.getUTCFullYear(), start.getUTCMonth(), start.getUTCDate());
      end = new Date(end.getUTCFullYear(), end.getUTCMonth(), end.getUTCDate());
    }
    return Object.assign({}, e, { startD: start, endD: end });
  }).filter(function (e) {
    if (isNaN(e.startD) || isNaN(e.endD)) return false;
    if (e.isCancelled) return false;
    if (!s.showFree && (e.showAs === 'free' || e.response === 'declined')) return false;
    if (!s.showAllDay && e.isAllDay) return false;
    return true;
  }).sort(function (a, b) { return a.startD - b.startD; });
}

function demoPayload() {
  var now = new Date();
  function at(h, m) { var d = new Date(now); d.setHours(h, m, 0, 0); return d; }
  function rel(min) { var d = new Date(now.getTime() + min * 60000); d.setSeconds(0, 0); return d; }
  var j = 'https://teams.microsoft.com/l/meetup-join/demo';
  var ev = [
    { subject: 'Company holiday planning', start: at(0, 0), end: addDays(at(0, 0), 1), isAllDay: true },
    { subject: 'Daily stand-up', start: at(9, 0), end: at(9, 15), joinUrl: j, organizer: 'Product Team' },
    { subject: 'Design review: onboarding flow', start: rel(-20), end: rel(25), joinUrl: j, organizer: 'Jordan Lee', location: 'Microsoft Teams Meeting' },
    { subject: '1:1 with manager', start: rel(40), end: rel(70), joinUrl: j, organizer: 'Sam Rivera' },
    { subject: 'Customer call - quarterly business review with the regional partner team', start: rel(90), end: rel(150), joinUrl: j, location: 'Teams', showAs: 'tentative' },
    { subject: 'Focus time', start: rel(160), end: rel(220), showAs: 'busy' },
    { subject: 'Sprint retro', start: rel(240), end: rel(285), joinUrl: j },
    { subject: 'Team sync (tomorrow)', start: addDays(at(9, 30), 1), end: addDays(at(10, 0), 1), joinUrl: j }
  ];
  return {
    status: 'ok', user: { name: 'Demo' }, updated: new Date().toISOString(),
    events: ev.map(function (e, i) {
      return {
        id: 'd' + i, subject: e.subject, start: e.isAllDay ? new Date(Date.UTC(e.start.getFullYear(), e.start.getMonth(), e.start.getDate())).toISOString() : e.start.toISOString(),
        end: e.isAllDay ? new Date(Date.UTC(e.end.getFullYear(), e.end.getMonth(), e.end.getDate())).toISOString() : e.end.toISOString(),
        isAllDay: !!e.isAllDay, location: e.location || '', organizer: e.organizer || '', joinUrl: e.joinUrl || '',
        isOnline: !!e.joinUrl, showAs: e.showAs || 'busy', isCancelled: false, response: 'accepted'
      };
    })
  };
}

function fetchWithTimeout(url, opts, ms) {
  var ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
  var t = setTimeout(function () { if (ctrl) ctrl.abort(); }, ms);
  if (ctrl) opts.signal = ctrl.signal;
  return fetch(url, opts).finally(function () { clearTimeout(t); });
}

var SETUP_STEPS =
  '<ol><li>In Outlook on the web: <b>Settings &rsaquo; Calendar &rsaquo; Shared calendars &rsaquo; Publish a calendar</b>. Choose your calendar and <b>Can view all details</b>, then <b>Publish</b>.</li>' +
  '<li>Copy the <b>ICS</b> link and paste it into this widget\'s <b>Calendar Link</b> setting in iCUE.</li></ol>';

function fail(kind, title, body) {
  var e = new Error(title); e.kind = kind; e.title = title; e.body = body; return e;
}

function refreshData() {
  var s = settings();
  if (state.fetching) return;
  state.nextFetchAt = Date.now() + RETRY_MS;

  if (DEMO) { state.payload = demoPayload(); state.stale = false; state.problem = null; render(); return; }

  if (!s.link) {
    state.problem = { kind: 'setup', title: 'Connect your calendar', body: SETUP_STEPS };
    render();
    return;
  }
  if (!LINK_RE.test(s.link)) {
    var html = /\.html?(?:[?#]|$)/i.test(s.link);
    state.problem = { kind: 'link', title: 'That isn\'t a calendar link',
      body: html ? 'This is the <b>HTML</b> link. Copy the <b>ICS</b> link from the same Outlook page instead.'
                 : 'Paste the <b>ICS</b> link from Outlook\'s <b>Publish a calendar</b> settings. It ends in <code>.ics</code>.' };
    render();
    return;
  }

  state.fetching = true;
  var link = s.link.replace(/^webcal:/i, 'https:').replace(/#.*$/, '');
  // text/plain keeps this a "simple" request (no CORS preflight); the link stays out of the URL.
  fetchWithTimeout(RELAY_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain' }, body: link, cache: 'no-store' }, 30000)
    .then(function (res) {
      return res.text().then(function (text) {
        if (res.ok) return text;
        if (res.status === 400) throw fail('link', 'That isn\'t a calendar link', 'Paste the <b>ICS</b> link from Outlook\'s <b>Publish a calendar</b> settings.');
        if (res.status === 404) throw fail('link', 'Calendar link not working', 'Outlook didn\'t return a calendar for this link. It may have been reset or unpublished. Publish the calendar again in Outlook and paste the new <b>ICS</b> link.');
        if (res.status === 429) throw fail('relay', 'Calendar service busy', 'Too many requests right now. Retrying automatically.');
        throw fail('relay', 'Calendar unavailable', 'Outlook didn\'t respond (' + escapeHtml(text.slice(0, 40) || res.status) + '). Retrying automatically.');
      });
    })
    .then(function (text) {
      var ws = startOfDay(new Date()), we = addDays(ws, 2);
      var events;
      try { events = TeamsIcs.eventsInWindow(text, ws.getTime(), we.getTime()); }
      catch (e) { throw fail('relay', 'Couldn\'t read the calendar', 'The calendar file had an unexpected format. Retrying automatically.'); }
      var data = { status: 'ok', events: events, updated: new Date().toISOString() };
      state.payload = data; state.stale = false; state.problem = null; saveCache(data);
      state.nextFetchAt = Date.now() + REFRESH_MS;
    })
    .catch(function (err) {
      state.stale = true;
      if (err && err.kind) state.problem = { kind: err.kind, title: err.title, body: err.body };
      else state.problem = { kind: 'offline', title: 'No connection', body: 'Couldn\'t reach the calendar service. Check this PC\'s internet connection. If iCUE asked about network access for this widget, choose Allow. Retrying automatically.' };
    })
    .finally(function () { state.fetching = false; render(); });
}

// ---------- rendering ----------
var $ = function (id) { return document.getElementById(id); };

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function setStatus(kind, text) {
  var dot = $('statusDot');
  dot.className = 'dot' + (kind === 'ok' ? '' : ' ' + kind);
  $('statusText').textContent = text;
}

function showStateMsg(problem) {
  $('stateMsg').classList.add('visible');
  $('main').classList.add('hidden');
  $('timeline').style.visibility = 'hidden';
  $('stateTitle').textContent = problem.title;
  // body strings are authored here (simple markup only); anything from the network is escaped before it gets here
  $('stateBody').innerHTML = problem.body;
  var needsGuide = ['setup', 'link'].indexOf(problem.kind) !== -1;
  $('guide').classList.toggle('visible', needsGuide);
}

var GUIDE_URL = 'https://github.com/font-max/teams-calendar-icue#quick-setup';
function initGuideButton() {
  var btn = $('guideBtn');
  btn.addEventListener('click', function () { openLink(GUIDE_URL); });
  btn.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') openLink(GUIDE_URL); });
}

function render() {
  var s = settings();
  var now = new Date();

  $('date').textContent = now.toLocaleDateString(state.language, { weekday: 'long', month: 'short', day: 'numeric' });
  $('clock').textContent = fmtTime(now, s);

  if (!state.payload) {
    if (state.problem) { setStatus('error', ''); showStateMsg(state.problem); }
    else { setStatus('stale', ''); showStateMsg({ title: 'Loading calendar...', body: '' }); }
    return;
  }

  $('stateMsg').classList.remove('visible');
  $('main').classList.remove('hidden');
  $('timeline').style.visibility = '';

  if (state.stale && state.problem) {
    var upd = state.payload.updated ? new Date(state.payload.updated) : null;
    setStatus(state.problem.kind === 'offline' ? 'error' : 'stale',
      state.problem.title + (upd ? ' - showing ' + fmtTime(upd, s) : ''));
  } else {
    setStatus('ok', '');
  }

  var events = normalize(state.payload, s);
  var today0 = startOfDay(now), tomorrow0 = addDays(today0, 1), dayAfter0 = addDays(today0, 2);
  var todayEvents = events.filter(function (e) { return e.startD < tomorrow0 && e.endD > today0; });
  var timed = todayEvents.filter(function (e) { return !e.isAllDay; });
  var allDay = todayEvents.filter(function (e) { return e.isAllDay; });

  var current = timed.filter(function (e) { return e.startD <= now && e.endD > now; })
    .sort(function (a, b) { return b.startD - a.startD; })[0] || null;
  var nextToday = timed.filter(function (e) { return e.startD > now; })[0] || null;
  var nextTomorrow = events.filter(function (e) { return !e.isAllDay && e.startD >= tomorrow0 && e.startD < dayAfter0; })[0] || null;

  renderHero(current, nextToday, nextTomorrow, now, s);
  renderAgenda(todayEvents, current || nextToday, allDay, now, s);
  renderTimeline(timed, current, now, s);
}

function renderHero(current, next, nextTomorrow, now, s) {
  var hero = $('hero');
  var ev = current || next;
  hero.className = 'hero';
  $('heroProgress').classList.remove('visible');

  if (!ev) {
    hero.classList.add('free');
    $('heroBadge').textContent = 'Free';
    $('heroCountdown').textContent = 'All clear';
    $('heroSubject').textContent = 'No more meetings today';
    $('heroMeta').textContent = nextTomorrow ? 'Tomorrow ' + fmtTime(nextTomorrow.startD, s) + ' - ' + nextTomorrow.subject : '';
    hero.onclick = null;
    hero.dataset.url = '';
    return;
  }

  if (current) {
    var total = current.endD - current.startD, done = now - current.startD;
    $('heroBadge').textContent = 'Now';
    $('heroCountdown').textContent = fmtDuration(current.endD - now) + ' left';
    $('heroProgress').classList.add('visible');
    $('heroProgressFill').style.width = Math.min(100, Math.max(0, (done / total) * 100)).toFixed(1) + '%';
    if (done < 5 * 60000) hero.classList.add('soon');
  } else {
    var until = next.startD - now;
    $('heroBadge').textContent = 'Up next';
    $('heroCountdown').textContent = until < 60000 ? 'Starting now' : 'in ' + fmtDuration(until);
    if (until <= 5 * 60000) hero.classList.add('soon');
  }
  if (ev.showAs === 'free') hero.classList.add('free');

  $('heroSubject').textContent = ev.subject;
  var meta = [fmtTime(ev.startD, s) + ' - ' + fmtTime(ev.endD, s)];
  if (ev.location && !/^microsoft teams/i.test(ev.location)) meta.push(ev.location);
  else if (ev.organizer) meta.push(ev.organizer);
  $('heroMeta').textContent = meta.join('  |  ');

  if (ev.joinUrl) {
    hero.classList.add('clickable');
    hero.onclick = function () { openLink(ev.joinUrl); };
  } else {
    hero.onclick = null;
  }
}

// ---------- agenda scrolling (touch drag, mouse drag, wheel) ----------
var AGENDA_IDLE_MS = 30000; // after this long without touching the list, it snaps back to the next meeting
var agendaScroll = { lastInteract: 0, programmatic: false, dragged: false };

function setAgendaScroll(list, top) {
  agendaScroll.programmatic = true;
  list.scrollTop = top;
  requestAnimationFrame(function () { agendaScroll.programmatic = false; });
}

function updateAgendaHint() {
  var list = $('agendaList');
  var box = list.getBoundingClientRect();
  var above = 0, below = 0;
  Array.prototype.forEach.call(list.querySelectorAll('.item'), function (r) {
    var rb = r.getBoundingClientRect();
    if (rb.bottom <= box.top + 2) above++;
    else if (rb.top >= box.bottom - 2) below++;
  });
  var parts = [];
  if (above) parts.push('↑ ' + above + ' earlier');
  if (below) parts.push(below + ' more ↓');
  $('agendaMore').textContent = parts.join('   ·   ');
  list.classList.toggle('fade-top', list.scrollTop > 2);
  list.classList.toggle('fade-bottom', list.scrollTop + box.height < list.scrollHeight - 2);
}

function initAgendaScrolling() {
  var list = $('agendaList');
  var drag = null, velocity = 0, lastY = 0, lastT = 0, inertia = 0;

  function touch() { agendaScroll.lastInteract = Date.now(); }
  function stopInertia() { if (inertia) { cancelAnimationFrame(inertia); inertia = 0; } }

  list.addEventListener('pointerdown', function (ev) {
    stopInertia();
    touch();
    drag = { y: ev.clientY, top: list.scrollTop };
    agendaScroll.dragged = false;
    lastY = ev.clientY; lastT = performance.now(); velocity = 0;
  });
  window.addEventListener('pointermove', function (ev) {
    if (!drag) return;
    var dy = ev.clientY - drag.y;
    if (!agendaScroll.dragged && Math.abs(dy) < 8) return;
    agendaScroll.dragged = true;
    touch();
    list.scrollTop = drag.top - dy;
    var t = performance.now();
    if (t > lastT) velocity = (ev.clientY - lastY) / (t - lastT); // px per ms
    lastY = ev.clientY; lastT = t;
  });
  function release() {
    if (!drag) return;
    drag = null;
    if (!agendaScroll.dragged || Math.abs(velocity) < 0.05) return;
    var v = velocity * 16; // px per frame
    (function step() {
      list.scrollTop -= v;
      v *= 0.92;
      touch();
      inertia = Math.abs(v) > 0.5 ? requestAnimationFrame(step) : 0;
    })();
  }
  window.addEventListener('pointerup', release);
  window.addEventListener('pointercancel', release);

  // a drag must not count as a tap on a meeting (which would open its join link)
  list.addEventListener('click', function (ev) {
    if (agendaScroll.dragged) { ev.stopPropagation(); ev.preventDefault(); agendaScroll.dragged = false; }
  }, true);

  list.addEventListener('wheel', touch, { passive: true });
  list.addEventListener('scroll', function () {
    if (!agendaScroll.programmatic) touch();
    updateAgendaHint();
  }, { passive: true });
}

function renderAgenda(todayEvents, heroEvent, allDay, now, s) {
  var list = $('agendaList');
  var prevTop = list.scrollTop;
  var userActive = Date.now() - agendaScroll.lastInteract < AGENDA_IDLE_MS;
  list.textContent = '';

  var items = allDay.concat(todayEvents.filter(function (e) {
    if (e.isAllDay) return false;
    if (heroEvent && e.id === heroEvent.id) return false;
    if (s.hidePast && e.endD <= now) return false;
    return true;
  }));

  $('agendaTitle').textContent = s.hidePast ? 'Later today' : 'Today';

  if (!items.length) {
    var empty = document.createElement('div');
    empty.className = 'agenda-empty';
    empty.textContent = heroEvent ? 'Nothing else today' : 'Enjoy the free time';
    list.appendChild(empty);
    $('agendaMore').textContent = '';
    list.classList.remove('fade-top', 'fade-bottom');
    return;
  }

  items.forEach(function (e) {
    var el = document.createElement('div');
    el.className = 'item';
    if (e.endD <= now) el.classList.add('past');
    if (e.isAllDay) el.classList.add('all-day');
    if (e.showAs === 'tentative' || e.response === 'tentativelyAccepted') el.classList.add('tentative');
    if (e.showAs === 'free') el.classList.add('free');

    var time = document.createElement('div');
    time.className = 'item-time';
    time.textContent = e.isAllDay ? 'All day' : fmtTime(e.startD, s);

    var text = document.createElement('div');
    text.className = 'item-text';
    var subj = document.createElement('div');
    subj.className = 'item-subject';
    subj.textContent = e.subject;
    var meta = document.createElement('div');
    meta.className = 'item-meta';
    if (!e.isAllDay) {
      var bits = [fmtDuration(e.endD - e.startD)];
      if (e.location && !/^microsoft teams/i.test(e.location)) bits.push(e.location);
      else if (e.organizer) bits.push(e.organizer);
      meta.textContent = bits.join('  |  ');
    }
    text.appendChild(subj);
    text.appendChild(meta);

    el.appendChild(time);
    el.appendChild(text);

    if (e.joinUrl && e.endD > now) {
      var join = document.createElement('div');
      join.className = 'item-join';
      join.textContent = 'Join';
      el.appendChild(join);
      el.classList.add('clickable');
      el.addEventListener('click', function () { openLink(e.joinUrl); });
    }
    list.appendChild(el);
  });

  if (userActive) {
    // keep the user's position while they are browsing (re-renders happen every 15 s)
    setAgendaScroll(list, prevTop);
  } else {
    // otherwise start at the next meeting that hasn't finished; earlier ones sit above it
    var first = list.querySelector('.item:not(.past):not(.all-day)') || list.querySelector('.item:not(.past)');
    setAgendaScroll(list, first ? first.offsetTop - list.firstElementChild.offsetTop : 0);
  }
  updateAgendaHint();
}

function renderTimeline(timed, current, now, s) {
  var tl = $('timeline');
  if (getComputedStyle(tl).display === 'none') return;

  var track = $('tlTrack'), labels = $('tlLabels');
  track.textContent = '';
  labels.textContent = '';

  var day0 = startOfDay(now);
  var startH = s.dayStart, endH = s.dayEnd;
  // stretch the window to include today's meetings outside the configured hours
  timed.forEach(function (e) {
    var sh = Math.max(0, (e.startD - day0) / 3600000), eh = Math.min(24, (e.endD - day0) / 3600000);
    if (sh < startH) startH = Math.floor(sh);
    if (eh > endH) endH = Math.ceil(eh);
  });
  var span = Math.max(1, endH - startH);
  var pct = function (d) { return Math.min(100, Math.max(0, (((d - day0) / 3600000) - startH) / span * 100)); };

  // lanes for overlaps (2 lanes)
  var laneEnd = [0, 0];
  var blocks = timed.map(function (e) {
    var lane = e.startD >= laneEnd[0] ? 0 : 1;
    laneEnd[lane] = Math.max(laneEnd[lane], e.endD);
    return { e: e, lane: lane };
  });
  blocks.forEach(function (b, i) {
    var overlaps = blocks.some(function (o, j) { return j !== i && o.e.startD < b.e.endD && o.e.endD > b.e.startD; });
    var el = document.createElement('div');
    el.className = 'tl-block lane-' + b.lane + (overlaps && b.lane === 0 ? ' split' : '');
    if (b.e.endD <= now) el.classList.add('past');
    if (current && b.e.id === current.id) el.classList.add('current');
    var left = pct(b.e.startD), right = pct(b.e.endD);
    el.style.left = left + '%';
    el.style.width = Math.max(0.4, right - left) + '%';
    el.textContent = b.e.subject;
    if (b.e.joinUrl && b.e.endD > now) el.addEventListener('click', function () { openLink(b.e.joinUrl); });
    track.appendChild(el);
  });

  var nowPct = pct(now);
  if (now > day0 && nowPct > 0 && nowPct < 100) {
    var line = document.createElement('div');
    line.className = 'tl-now';
    line.style.left = nowPct + '%';
    track.appendChild(line);
  }

  var step = span > 12 ? 3 : span > 8 ? 2 : 1;
  for (var h = startH; h <= endH; h += step) {
    var lab = document.createElement('span');
    var d = new Date(day0); d.setHours(h);
    lab.style.left = ((h - startH) / span * 100) + '%';
    lab.textContent = h === 24 ? (s.use24h ? '24:00' : '12 AM') :
      d.toLocaleTimeString(state.language, { hour: 'numeric', hour12: !s.use24h });
    labels.appendChild(lab);
  }
}

// ---------- lifecycle ----------
var timersStarted = false;
function startTimers() {
  if (timersStarted) return;
  timersStarted = true;
  setInterval(function () {
    var s = settings();
    if (Date.now() >= state.nextFetchAt) refreshData();
    render();
  }, 15000);
  setInterval(function () {
    var s = settings();
    $('clock').textContent = fmtTime(new Date(), s);
  }, 1000);
  if (typeof ResizeObserver !== 'undefined') {
    var raf = 0;
    new ResizeObserver(function () { cancelAnimationFrame(raf); raf = requestAnimationFrame(render); }).observe(document.body);
  }
}

function onIcueDataUpdated() {
  applyStyles();
  if (!state.payload && !DEMO) {
    var cached = loadCache();
    if (cached) { state.payload = cached; state.stale = true; }
  }
  refreshData();
  startTimers();
}

function onIcueInitialized() {
  if (typeof iCUE !== 'undefined' && iCUE.iCUELanguage) state.language = String(iCUE.iCUELanguage).replace('_', '-');
  onIcueDataUpdated();
}

var bootAttempts = 0;
function bootCheck() {
  if (typeof iCUE_initialized !== 'undefined' && iCUE_initialized) { onIcueInitialized(); return; }
  if (bootAttempts < 15) { bootAttempts++; setTimeout(bootCheck, 100); return; }
  try { state.language = navigator.language || 'en'; } catch (e) {}
  onIcueDataUpdated(); // plain browser
}
initAgendaScrolling();
initGuideButton();
render();
bootCheck();
