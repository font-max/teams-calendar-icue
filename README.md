# Teams Calendar for iCUE (CORSAIR Xeneon Edge)
<img width="2000" height="1123" alt="xeneon-edge-mockup" src="https://github.com/user-attachments/assets/2e75a3d1-c025-4bed-96f1-55eafded6cc8" />

Simple connector to your iCUE devices to see your meetings in Teams.

Shows today's Microsoft Teams / Outlook meetings on the Xeneon Edge: a big "now / up next" card with a countdown, the rest of the day's agenda (swipe to scroll), and (on Large / Extra Large horizontal slots) a day timeline. Tap a meeting to open its Teams join link.

## Quick setup

About 2 minutes. You need Windows, iCUE 5.52+ and a Microsoft 365 (work or school) account.

### 1. Get the widget
Install **Teams Calendar** from the iCUE widget marketplace, or download [`teams-calendar.icuewidget`](https://github.com/font-max/teams-calendar-icue/releases/latest/download/teams-calendar.icuewidget) and open it.

### 2. Run Setup on your PC
1. Download **[TeamsCalendarSetup.zip](https://github.com/font-max/teams-calendar-icue/releases/latest/download/TeamsCalendarSetup.zip)** and extract it (right-click → **Extract All**).
2. Double-click **Setup.cmd**. If Windows shows *"Windows protected your PC"*, click **More info → Run anyway**.
3. Follow the prompts. Setup:
   - installs Node.js if you don't have it;
   - opens Outlook's calendar-sharing page. Under **Publish a calendar**, choose **Calendar** and **Can view all details**, click **Publish**, copy the **ICS** link, and paste it into Setup;
   - checks that it can read your meetings, then starts the helper and makes it start with Windows;
   - **copies your access key to the clipboard**.

### 3. Paste the key in iCUE
Select the Teams Calendar widget, paste (Ctrl+V) into **Helper Access Key**, and if iCUE asks about network access to `localhost:38127`, choose **Allow**. Your meetings appear within a minute.
<img width="599" height="283" alt="howtoicue" src="https://github.com/user-attachments/assets/e93e1de2-c9af-4251-86a0-5946ec756de2" />

### Troubleshooting
| The widget says | Do this |
|---|---|
| **Connect your calendar** | The key isn't set: run Setup and paste the key it copies. |
| **Access key not accepted** | Run Setup again (it copies the current key) and paste it. |
| **Helper not running** | Restart the PC, or run Setup again. |
| **Sign-in needed** / no calendar link | Run Setup again and paste your ICS link. |
| Meetings are a little behind | Microsoft refreshes published calendars on its own schedule; changes can take up to ~30 min. |

- **Can't find "Publish a calendar"?** Your organization has turned calendar publishing off. Ask your IT team to allow it for your account; Teams Calendar needs it to read your meetings.
- **To change your calendar link,** run Setup again.
- **To remove everything,** double-click **Uninstall.cmd**. Also click **Unpublish** in Outlook's calendar-sharing settings to stop sharing the link.
- **Privacy:** your calendar link and access key stay on your PC (the link is encrypted for your Windows user). The helper only answers requests from this PC, and only with your access key.
- **Treat the ICS link like a password.** Anyone who has it can read your calendar at the detail level you picked. To revoke it, click **Unpublish** (or **Reset links**) in Outlook, then run Setup again with the new link.

## How it works

iCUE widgets are sandboxed web pages, so a small **helper** on your PC reads the calendar for them:

1. Setup installs the helper to `%LOCALAPPDATA%\TeamsCalendarHelper\app` and adds **one** startup entry (`Teams Calendar Helper (ICS)` in your Startup folder). It removes any older Teams Calendar helper it finds in the Startup folder, registry Run keys or scheduled tasks.
2. Every 5 minutes the helper downloads your published calendar and works out today's and tomorrow's meetings: repeating series, moved or skipped occurrences, time zones and daylight saving.
3. It serves them **only on this PC** at `http://localhost:38127/api/events`. Requests need a random access key, which you paste into the widget settings. That stops other web pages in your browser from reading your calendar.
4. If an older Teams Calendar helper is holding the port when it starts, the installed helper stops it and takes over, so the right helper always wins after a restart.
5. The widget asks the helper again just after each update, and opens join links through iCUE's Link plugin (the system browser hands off to Teams).

The helper sends Teams, Zoom, Google Meet and Webex links to the widget, never the meeting description. Events marked Private show as "Private appointment". It contacts only the calendar link's host (`outlook.office365.com`). Log: `%LOCALAPPDATA%\TeamsCalendarHelper\helper-ics.log`.

## Widget settings

| Setting | Default | |
|---|---|---|
| Helper Access Key | (empty) | copied to your clipboard by Setup |
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
| `url` | `localhost`, port `38127` | read meetings from the local helper |
| plugin | `widgetbuilder.linkprovider:Url:1.0` | open the setup guide and join links in the system browser (no permission needed) |

## For developers

```
teams-calendar-widget/
├── widget/TeamsCalendar/        the iCUE widget (index.html, manifest.json, scripts/, styles/)
├── helper-ics/                  the calendar helper (ics-calendar-helper.mjs, Node.js 18+, no dependencies)
├── setup/                       Setup.cmd / Uninstall.cmd / setup.ps1, the only supported way to install the helper
├── tools/build-release.ps1      builds dist/teams-calendar.icuewidget and dist/TeamsCalendarSetup.zip
└── dist/                        release files
```

- **Build a release:** `powershell -ExecutionPolicy Bypass -File tools\build-release.ps1` (needs `npm install -g icuewidget-cli`).
- **Try the widget in a browser** with sample data: serve `widget/TeamsCalendar/` locally and open `index.html?demo=1`.
- **Helper commands** (from `helper-ics/`): `node ics-calendar-helper.mjs test` (read the saved link once), `set-url <link>`, `key`, `serve`. Install with Setup rather than starting it by hand, so there is only ever one startup entry.
- **Setup options:** `Setup.cmd -KeepLink` (reuse the saved link without asking), `-IcsUrl <link>`, `-Unattended`.
