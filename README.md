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

### Troubleshooting
| The widget says | Do this |
|---|---|
| **Connect your calendar** | The key isn't set: run Setup and paste the key it copies. |
| **Access key not accepted** | Run Setup again (it copies the current key) and paste it. |
| **Helper not running** | Restart the PC, or run Setup again. |
| **Sign-in needed** / no calendar link | Run Setup again and paste your ICS link. |
| Meetings are a little behind | Microsoft refreshes published calendars on its own schedule; changes can take up to ~30 min. |

- **Can't find "Publish a calendar"?** Your organization has turned publishing off. See [Which helper?](#which-helper) for the other options.
- **To change your calendar link,** run Setup again.
- **To remove everything,** double-click **Uninstall.cmd**. Also click **Unpublish** in Outlook's calendar-sharing settings to stop sharing the link.
- **Privacy:** your calendar link and access key stay on your PC (the link is encrypted for your Windows user). The helper only answers requests from this PC.

## For developers

```
teams-calendar-widget/
├── setup/                       one-click Setup.cmd / Uninstall.cmd (built into TeamsCalendarSetup.zip)
├── tools/build-release.ps1      builds dist/teams-calendar.icuewidget and dist/TeamsCalendarSetup.zip
├── helper/                      v1: Microsoft Graph (needs IT approval of the sign-in app)
│   ├── teams-calendar-helper.mjs
│   ├── start-hidden.vbs
│   └── install-startup.ps1
├── helper-ics/                  v2: published calendar (ICS) link, works with new Outlook
│   ├── ics-calendar-helper.mjs
│   ├── start-hidden.vbs
│   └── install-startup.ps1
├── helper-outlook/              alternative: reads classic Outlook on this PC (not used)
│   ├── outlook-calendar-helper.mjs
│   ├── read-outlook-calendar.ps1
│   ├── start-hidden.vbs
│   └── install-startup.ps1
└── widget/TeamsCalendar/        the iCUE widget (works with any helper)
    ├── index.html, manifest.json, translation.json, icon.svg
    ├── scripts/TeamsCalendar.js
    └── styles/TeamsCalendar.css
```

## How it works

iCUE widgets are sandboxed web pages, so the Microsoft sign-in happens in a small local **helper** instead:

1. The helper signs in to Microsoft 365 once using the device-code flow. Your refresh token is stored encrypted for your Windows user (DPAPI) in `%LOCALAPPDATA%\TeamsCalendarHelper\`.
2. Every 2 minutes it reads today's and tomorrow's events from Microsoft Graph (`/me/calendarView`, read-only `Calendars.Read`).
3. It serves them **only on this PC** at `http://localhost:38127/api/events`. Requests need a random access key, which you paste into the widget settings. That stops other web pages in your browser from reading your calendar.
4. The widget polls the helper and opens join links through iCUE's Link plugin (the system browser hands off to Teams).

The helper sends Teams, Zoom, Google Meet and Webex links to the widget, but never the meeting body. Events marked Private show as "Private appointment".

## Which helper?

All helpers serve the same data on the same port with the same access key, so the widget works with any of them. **Run only one at a time.**

| | v1 `helper/` (Microsoft Graph) | v2 `helper-ics/` (published ICS link) | `helper-outlook/` (classic Outlook) |
|---|---|---|---|
| Sign-in | device code in your browser | none: uses a private link | inside classic Outlook |
| IT approval | often required (depends on your organization) | not needed (publishing must be allowed) | not needed |
| Works with new Outlook only | yes | yes | no, needs classic Outlook running |
| Freshness | ~2 min | Microsoft refreshes the published copy; changes can take up to ~30 min | ~2 min |
| Teams join links | yes | only if published with "Can view all details" | yes |

## Setup: v2 (published ICS link)

1. **Publish your calendar.** In new Outlook (or Outlook on the web): **Settings → Calendar → Shared calendars → Publish a calendar**. Pick your **Calendar** and a permission level:
   - **Can view all details**: subjects, locations and Teams join links (tap to join works).
   - **Can view titles and locations**: no join links.

   Click **Publish** and copy the **ICS** link (not the HTML one). If this section is missing or greyed out, your organization has turned publishing off.
2. **Give the link to the helper.** It's stored encrypted for your Windows user:
   ```bash
   node helper-ics\ics-calendar-helper.mjs set-url "PASTE-ICS-LINK-HERE"
   ```
3. **Check it:**
   ```bash
   node helper-ics\ics-calendar-helper.mjs test
   ```
4. **Stop v1** if it's running: close its console window.
5. **Run v2 at every Windows sign-in:**
   ```bash
   powershell -ExecutionPolicy Bypass -File helper-ics\install-startup.ps1
   ```
6. **Install the widget** (step 3 of the v1 setup below) and paste the key from `node ics-calendar-helper.mjs key`.

**Treat the ICS link like a password.** Anyone who has it can read your calendar at the detail level you picked. To revoke it, go back to *Publish a calendar* and click **Unpublish** (or **Reset links**), then run `set-url` with the new link.

How it works: every 5 minutes the helper downloads the published calendar and works out today's and tomorrow's meetings: repeating series, moved or skipped occurrences, time zones and daylight saving. It sends Teams/Zoom/Meet/Webex links to the widget, never the meeting description. Log: `%LOCALAPPDATA%\TeamsCalendarHelper\helper-ics.log`.

## Alternative: classic Outlook helper

1. **Sign in to classic Outlook once.** Run `outlook.exe`. If it opens new Outlook, turn off the "New Outlook" toggle. Add your work account and finish the email, password and 2FA prompts. You can keep using new Outlook day to day.
2. **Check that the helper can read your calendar:**
   ```bash
   node helper-outlook\outlook-calendar-helper.mjs test
   ```
3. **Stop v1** if it's running: close its console window, or run `helper\install-startup.ps1 -Remove` if you set it to start with Windows.
4. **Run v2 at every Windows sign-in:**
   ```bash
   powershell -ExecutionPolicy Bypass -File helper-outlook\install-startup.ps1
   ```
5. **Install the widget** (step 3 below) and paste the key from `node outlook-calendar-helper.mjs key`.

How it works: every 2 minutes a PowerShell script asks classic Outlook for today's and tomorrow's appointments, with repeating meetings included. It gets each Teams join link from the meeting itself (falling back to the body), and sends nothing else from the body to the widget. Your password never touches the helper; Outlook handles the sign-in. If Windows ever shows "A program is trying to access e-mail address information stored in Outlook", that's Outlook's programmatic-access guard. It normally stays quiet while Windows Defender is up to date. Log: `%LOCALAPPDATA%\TeamsCalendarHelper\helper-outlook.log`.

## Setup: v1 (Microsoft Graph, after IT approval)

Requirements: iCUE 5.52+, Node.js 18+, Windows.

### 1. Sign in

```bash
node helper\teams-calendar-helper.mjs login
```

A browser opens. Enter the code that's shown, sign in with your work account, and approve the **Calendars.Read** access. The command then prints your **widget access key**.

- By default it uses Microsoft's public "Microsoft Graph Command Line Tools" app (the one `Connect-MgGraph` uses). If your organization blocks it or needs admin approval, register your own app: Entra ID → App registrations → New, then under Authentication turn on *Allow public client flows*, and add the delegated permissions `Calendars.Read` and `User.Read`. Then run:
  `node teams-calendar-helper.mjs login --client-id <your-app-id> --tenant <your-tenant-id>`
- Personal Microsoft accounts: add `--tenant consumers` (works only with your own app registration).

### 2. Run the helper in the background at every Windows sign-in

```bash
powershell -ExecutionPolicy Bypass -File helper\install-startup.ps1
```

This adds a shortcut to your Startup folder and starts the helper now. To remove it, run the same command with `-Remove`. The log is at `%LOCALAPPDATA%\TeamsCalendarHelper\helper.log`.

### 3. Install the widget

The ready-made package is [`dist/teams-calendar.icuewidget`](dist/teams-calendar.icuewidget). To build it yourself instead:

```bash
npm install -g icuewidget-cli
```

```bash
cd widget && icuewidget validate TeamsCalendar && icuewidget package TeamsCalendar
```

Open the generated `.icuewidget` file to install it in iCUE. When asked, approve the **localhost:38127** network permission. Add the widget to an Xeneon Edge slot and paste the access key into **Helper Access Key**. To show the key again, run `node teams-calendar-helper.mjs key`.

## Widget settings

| Setting | Default | |
|---|---|---|
| Helper Access Key | (empty) | from `teams-calendar-helper.mjs key` |
| 24-Hour Clock | off | |
| Hide Finished Meetings | on | |
| Show All-Day Events | on | |
| Show Free and Declined | off | hides cancelled and declined events, and events marked Free |
| Timeline Start / End | 8 h / 18 h | the window stretches to fit earlier or later meetings |
| Text / Accent / Background Color, Background Transparency | white / Teams purple `#7b83eb` / black / 100% | |

The widget has no refresh setting. It asks the helper again just after the helper's own update (every 5 min for the ICS helper, 2 min for the others), and retries every 60 s while something is wrong (no key, helper offline).

## Layouts

| Slot | Layout |
|---|---|
| S horizontal 840×344 | hero card + compact agenda side by side |
| M horizontal 840×696 | hero card over agenda |
| L / XL horizontal 1688×696, 2536×696 | hero + agenda side by side, day timeline with a red "now" line |
| S vertical 696×416 | hero + 1–2 agenda rows |
| M / L / XL vertical | hero + long agenda |

Rows that don't fit are hidden and summarized as "+N more".

## States

- **Connect your calendar**: no access key set yet.
- **Helper not running**: can't reach localhost:38127. The widget shows the last data it received, with a red dot.
- **Access key not accepted**: the key doesn't match the helper's.
- **Sign-in needed**: the token expired or was revoked. Run `login` again.
- **All clear**: no more meetings today. Shows tomorrow's first meeting.

To preview with sample data in a normal browser, open `widget/TeamsCalendar/index.html?demo=1` from a local web server.

## Permissions

| Permission | Scope | Why |
|---|---|---|
| `url` | `localhost`, port `38127` | read meetings from the local helper |
| plugin | `widgetbuilder.linkprovider:Url:1.0` | open join links in the system browser (no permission needed) |

The helper itself contacts only `login.microsoftonline.com` and `graph.microsoft.com`.

## Helper commands

`login` · `serve` (default) · `key` · `logout`
