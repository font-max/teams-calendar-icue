// Teams Calendar relay (Cloudflare Worker)
//
// iCUE widgets cannot read an Outlook published calendar directly: Outlook does not send the
// CORS header that lets a web page read it. This relay downloads the calendar for the widget and
// passes it through unchanged with that header added.
//
// Privacy: nothing is stored or logged. The calendar link travels in the request body (never in
// the URL), the calendar is streamed straight back, and responses are marked no-store.
// Only Outlook published-calendar links are accepted, so this is not an open proxy.
//
// POST /calendar   body: the ICS link (text/plain)  ->  the calendar (text/calendar)
// GET  /           short description

const ALLOWED_HOSTS = new Set(['outlook.office365.com', 'outlook.office.com', 'outlook.live.com']);
const PATH_RE = /^\/owa\/calendar\/[^/]+\/[^/]+\/[^/]+\.ics$/i;
const MAX_BYTES = 8 * 1024 * 1024;

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Access-Control-Max-Age': '86400',
};

function reply(status, body, type = 'text/plain; charset=utf-8') {
  return new Response(body, {
    status,
    headers: { ...CORS, 'Content-Type': type, 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' },
  });
}

function checkLink(raw) {
  let u;
  try { u = new URL(String(raw).trim().replace(/^webcal:/i, 'https:')); } catch { return null; }
  if (u.protocol !== 'https:' || u.username || u.password || u.port) return null;
  if (!ALLOWED_HOSTS.has(u.hostname.toLowerCase()) || !PATH_RE.test(u.pathname)) return null;
  u.hash = '';
  return u;
}

export default {
  async fetch(request) {
    const { pathname } = new URL(request.url);

    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: CORS });

    if (pathname === '/' && request.method === 'GET') {
      return reply(200, 'Teams Calendar relay for the iCUE "Teams Calendar" widget.\n' +
        'Source and privacy notes: https://github.com/font-max/teams-calendar-icue/tree/main/relay\n');
    }

    if (pathname !== '/calendar') return reply(404, 'Not found');
    if (request.method !== 'POST') return reply(405, 'Use POST');

    const body = await request.text();
    if (body.length > 2048) return reply(400, 'bad_link');
    const link = checkLink(body);
    if (!link) return reply(400, 'bad_link');

    let upstream;
    try {
      upstream = await fetch(link.toString(), {
        headers: { Accept: 'text/calendar', 'User-Agent': 'TeamsCalendarRelay/1.0' },
        redirect: 'follow',
        cf: { cacheTtl: 0, cacheEverything: false },
      });
    } catch {
      return reply(502, 'upstream_unreachable');
    }

    // A redirect must not lead outside Outlook.
    const finalUrl = new URL(upstream.url || link.toString());
    if (!ALLOWED_HOSTS.has(finalUrl.hostname.toLowerCase())) return reply(404, 'not_found');

    if (upstream.status === 404 || upstream.status === 410 || upstream.status === 401 || upstream.status === 403) {
      return reply(404, 'not_found'); // link unpublished, reset or mistyped
    }
    if (!upstream.ok) return reply(502, 'upstream_' + upstream.status);

    const len = Number(upstream.headers.get('Content-Length') || 0);
    if (len > MAX_BYTES) return reply(413, 'too_large');

    const text = await upstream.text();
    if (text.length > MAX_BYTES) return reply(413, 'too_large');
    if (!/BEGIN:VCALENDAR/i.test(text.slice(0, 2000))) return reply(404, 'not_a_calendar');

    return reply(200, text, 'text/calendar; charset=utf-8');
  },
};
