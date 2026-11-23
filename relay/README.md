# Teams Calendar relay

A tiny Cloudflare Worker used by the iCUE **Teams Calendar** widget.

iCUE widgets can't read an Outlook published calendar (.ics) directly, because Outlook doesn't send
the CORS header that lets a web page read it. The widget sends its calendar link to this relay, which
downloads the calendar from Outlook and returns it unchanged with that header added. The widget then
reads the meetings itself, on your PC.

## What it does and doesn't do

- Accepts only Outlook published-calendar links (`outlook.office365.com`, `outlook.office.com`,
  `outlook.live.com`, path `/owa/calendar/.../*.ics`). Anything else is refused, so it is not an open proxy.
- The link is sent in the request body, never in the URL, so it doesn't appear in request logs.
- Nothing is stored, cached or logged. Workers logs/traces are disabled in `wrangler.toml`.
- The calendar is passed through as-is; the relay doesn't read or change your meetings.

## Deploy your own

```powershell
cd relay
npx wrangler login
npx wrangler deploy
```
