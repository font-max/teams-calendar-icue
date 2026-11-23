# Teams Calendar for iCUE (CORSAIR Xeneon Edge)
<img width="2000" height="1123" alt="xeneon-edge-mockup" src="https://github.com/user-attachments/assets/2e75a3d1-c025-4bed-96f1-55eafded6cc8" />

Simple connector to your iCUE devices to see your meetings in Teams.

Shows today's Microsoft Teams / Outlook meetings on the Xeneon Edge: a big "now / up next" card with a countdown, the rest of the day's agenda (swipe to scroll), and (on Large / Extra Large horizontal slots) a day timeline. Tap a meeting to open its Teams join link.

## Quick setup

About 2 minutes. You need iCUE 5.52+ and an Outlook / Microsoft 365 calendar. Nothing to install on your PC.

### 1. Get the widget
Install **Teams Calendar** from the iCUE widget marketplace, or download [`teams-calendar.icuewidget`](https://github.com/font-max/teams-calendar-icue/releases/latest/download/teams-calendar.icuewidget) and open it.

### 2. Publish your calendar in Outlook
1. Open **[Outlook on the web → Settings → Calendar → Shared calendars](https://outlook.office.com/mail/options/calendar/SharedCalendars)**.
2. Under **Publish a calendar**, choose **Calendar** and **Can view all details**, then click **Publish**.
3. Copy the **ICS** link (not the HTML one). It ends in `.ics`.

### 3. Paste the link in iCUE
Select the Teams Calendar widget and paste (Ctrl+V) into **Calendar Link**. If iCUE asks about network access, choose **Allow**. Your meetings appear within a few seconds.

### Troubleshooting
| The widget says | Do this |
|---|---|
| **Connect your calendar** | Paste your ICS link into **Calendar Link**. |
| **That isn't a calendar link** | Copy the **ICS** link from Outlook's *Publish a calendar* (not the HTML link). |
| **Calendar link not working** | The link was reset or unpublished. Publish again in Outlook and paste the new ICS link. |
| **No connection** | Check the PC's internet connection; it retries automatically. |
| Meetings are a little behind | Microsoft refreshes published calendars on its own schedule; changes can take up to ~30 min. |

- **Can't find "Publish a calendar"?** Your organization has turned calendar publishing off. Ask your IT team to allow it for your account; Teams Calendar needs it to read your meetings.
- **Treat the ICS link like a password.** Anyone who has it can read your calendar at the detail level you picked. To revoke it, click **Unpublish** (or **Reset links**) in Outlook.
- **Upgrading from v1.x (helper version)?** The helper is no longer needed. Run **Uninstall.cmd** from your old TeamsCalendarSetup folder (or the [v1.5.0 setup zip](https://github.com/font-max/teams-calendar-icue/releases/tag/v1.5.0)) to remove it.

## How it works

Outlook doesn't let web pages read a published calendar directly, and iCUE widgets are web pages. So the widget fetches it through a tiny **relay** ([source](relay/worker.js)) running on Cloudflare:

1. Every 10 minutes the widget sends your ICS link to the relay (in the request body, never in the URL).
2. The relay downloads the calendar from Outlook and passes it back **unchanged**, adding the one header iCUE needs. It accepts only Outlook published-calendar links, and it **stores and logs nothing**.
3. The widget reads the calendar **on your PC**: repeating series, moved or skipped occurrences, time zones and daylight saving. It opens join links through iCUE's Link plugin.

### Privacy
- Your calendar passes through the relay on its way from Outlook to your PC. It isn't stored, cached or logged, and Cloudflare Workers logging is turned off ([config](relay/wrangler.toml)).
- Your link is saved only in iCUE's widget settings on your PC.
- The widget uses Teams, Zoom, Google Meet and Webex links from meeting descriptions; nothing else from the description is shown. Events marked Private show as "Private appointment".
- Prefer nothing to leave your PC except the call to Outlook? Use **v1.5.0**, which reads the calendar with a local helper instead: [release](https://github.com/font-max/teams-calendar-icue/releases/tag/v1.5.0).

## Widget settings

| Setting | Default | |
|---|---|---|
| Calendar Link | (empty) | your Outlook published-calendar ICS link |
| 24-Hour Clock | off | |
| Hide Finished Meetings | on | |
| Show All-Day Events | on | |
| Show Free and Declined | off | hides cancelled and declined events, and events marked Free |
| Timeline Start / End | 8 h / 18 h | the window stretches to fit earlier or later meetings |
| Text / Accent / Background Color, Background Transparency | white / Teams purple `#7b83eb` / black / 100% | |

## Layouts

| Slot | Layout |
|---|---|
| S horizontal 840×344 | hero card + compact agenda side by side |
| M horizontal 840×696 | hero card over agenda |
| L / XL horizontal 1688×696, 2536×696 | hero + agenda side by side, day timeline with a red "now" line |
| S vertical 696×416 | hero + 1–2 agenda rows |
| M / L / XL vertical | hero + long agenda |

The agenda scrolls with a swipe and returns to your next meeting after 30 seconds.

## Permissions

| Permission | Scope | Why |
|---|---|---|
| `url` | `teams-calendar-relay.fontoura.workers.dev`, port `443` | download your published calendar through the relay |
| plugin | `widgetbuilder.linkprovider:Url:1.0` | open the setup guide and join links in the system browser (no permission needed) |

## For developers

```
teams-calendar-widget/
├── widget/TeamsCalendar/        the iCUE widget (index.html, manifest.json, scripts/, styles/)
│   └── scripts/ics.js           ICS reader: recurrence, exceptions, time zones (runs in the widget)
├── relay/                       Cloudflare Worker that adds the CORS header (worker.js, wrangler.toml)
├── helper-ics/, setup/          v1.x local helper and its Setup (kept for the v1.5.0 release)
├── tools/build-release.ps1      builds dist/teams-calendar.icuewidget
└── dist/                        release files
```

- **Build a release:** `powershell -ExecutionPolicy Bypass -File toolsuild-release.ps1` (needs `npm install -g icuewidget-cli`).
- **Deploy the relay:** `cd relay`, then `npx wrangler login` and `npx wrangler deploy`. If you deploy your own, change `RELAY_URL` in `scripts/TeamsCalendar.js` and the `url` permission in `manifest.json` to match.
- **Try the widget in a browser** with sample data: serve `widget/TeamsCalendar/` locally and open `index.html?demo=1`.
