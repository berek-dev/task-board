# Task Board

A Monday-style task board for phone and PC whose only storage is your own calendar accounts:
- **Microsoft 365:** tasks are **Microsoft To Do** tasks and events are **Outlook calendar** events.
- **Google:** tasks are **Google Tasks** and events are **Google Calendar** events.

Connect one or both. With both, the board shows everything together: each list is marked
"To Do" or "Google", Outlook and Google Calendar are separate calendar groups, and you pick which
calendar a new event goes to. Moving a task to a list in the other account moves it there.
There is no database and no server, so anything you change here shows up in Outlook / To Do /
Google, and anything you change there shows up here on the next sync (every 2 minutes, when you
reopen the app, or with ⟳).

Try it without an account: open the app with `?demo` at the end of the URL.

The interface is in English and Hebrew (עברית, right-to-left). It opens in Hebrew when your phone or
browser language is Hebrew; the English / עברית button in the top bar switches and remembers your choice.

## What it can do

- **Four screens:** Table (grouped like Monday), Board (drag cards between status columns),
  Calendar (week and month), and Overview (counts, progress per list, the next 7 days, overdue).
- **Subtasks:** a checklist of steps on any task, with "2/5" progress shown on the row and card.
- **Calendar and time blocking:** drag events to another time or day to reschedule them. Drag a task
  from the Tasks panel onto a time in the week view to create an Outlook event that blocks time for it,
  or use "📅 Block time" in the task's edit window (works on the phone too). Dragging is for PC.
- **Repeat and reminders:** tasks and events can repeat (daily, weekdays, weekly, monthly, yearly) and
  have a reminder. Outlook and To Do send the reminder notifications to your phone themselves.
  Completing a repeating task makes To Do create the next one.
- **Labels:** colored tags on tasks and events, saved as Outlook categories, so you see them in Outlook too.
- **Filters and saved views:** filter by status, priority, date and label; group the table by list,
  status, priority, due date or label; sort by date, priority, status or name. Ready-made views
  (This week, Overdue, Stuck, High priority) and your own saved views. Saved views are kept on each device.
- **Links and files:** web links on tasks, and file attachments (up to 3 MB) on tasks and events.
- **Priorities:** a tab with all your open tasks in one list. Drag ⠿ (or tap ⤒) to put them in order; the
  order is saved with each task (an Outlook To Do extension, or the Google Tasks notes line), so it syncs to
  every device. "My order" in the table's Sort by uses it too.
- **Reminders:** on the due date at a time you pick, or a one-time reminder at any date and time.
- **On a phone:** tabs at the bottom, a big + New button, full-screen editing, and on the board tap a card's
  status (or long-press the card) to move it, since dragging doesn't work well with fingers.
- **Bulk edit:** tick several rows in the table, then change status, priority, date, list or label, or delete them together.
- **Keyboard shortcuts (PC):** N new, / search, 1–4 switch screen, F filters, R sync, ← → and T in the
  calendar, Esc clear selection, Ctrl+Enter save, ? shows the list.

## How board fields are stored (Outlook)

| Board field | Microsoft To Do task | Outlook event |
|---|---|---|
| Group | The To Do list (each list is a group) | "Outlook calendar" group |
| Status | To Do's own status: Not started, In progress (= Working on it), Waiting on others (= Stuck), Completed (= Done) | Hidden extended property, not visible in Outlook |
| Priority | Importance (Low / Normal / High) | Importance (Low / Normal / High) |
| Date | Due date | Start date and time, or all-day |
| Notes | Task notes | Event description |
| Subtasks | Checklist steps | (tasks only) |
| Repeat | To Do recurrence | Outlook recurrence (a single occurrence's pattern is changed in Outlook) |
| Reminder | To Do reminder time on the due date | Outlook reminder (minutes before) |
| Labels | Categories | Categories |
| Links | Linked resources | (tasks only) |
| Files | Task attachments | Event attachments |

Outlook only has Low / Normal / High, so the board offers those three priorities.
When you edit an event, the board only sends the fields you changed, so meeting invites keep their
formatting and Teams links unless you edit the notes.

## One-time setup (about 10 minutes)

### Work account: check with IT first
Your workplace Microsoft 365 is controlled by your company's IT admin. Many companies block
employees from approving new apps themselves. If Microsoft says **"Need admin approval"** at sign-in,
send your IT admin the app's name and the three permissions it asks for:
`User.Read`, `Calendars.ReadWrite`, `Tasks.ReadWrite`. All three are "delegated": the app acts only
as you, on your own calendar and tasks. They can approve it in the Entra admin center, or register
the app themselves with the steps below. Some companies don't let employees create app
registrations at all; then IT has to do steps 1 to 4.

### Register the app
1. Go to https://entra.microsoft.com (or https://portal.azure.com → *Microsoft Entra ID*) and sign in
   with the account whose calendar you want to use.
2. **App registrations → New registration**:
   - Name: "Task Board".
   - Supported account types: *Accounts in this organizational directory only* if only people at your
     company will use it. If someone at another company (for example your spouse) will use it too, choose
     *Accounts in any organizational directory* (add *and personal Microsoft accounts* for Outlook.com).
     Each person signs in on their own device and only ever sees their own calendar and tasks;
     their company's IT may need to approve the app for their account too.
   - Redirect URI: platform **Single-page application (SPA)**, URL = where you'll host the app,
     e.g. `https://<you>.github.io/task-board/` (with the trailing slash). Then under
     Authentication add `http://localhost:8080/` too, for testing on your PC.
3. **API permissions → Add a permission → Microsoft Graph → Delegated permissions**: add
   `Calendars.ReadWrite` and `Tasks.ReadWrite` (`User.Read` is already there). If you have the rights,
   click *Grant admin consent*; otherwise that's the step for IT.
4. From **Overview**, copy the **Application (client) ID** into `MS_CLIENT_ID` in `config.js`.
   For a work-only registration, also copy the **Directory (tenant) ID** into `MS_TENANT`.
   Neither is a secret, and there is no client secret in this app.

You normally sign in once per device: Microsoft's sign-in library refreshes the session in the background.

## Run it on your PC

```
cd task-manager
python3 -m http.server 8080
```
Then open http://localhost:8080.

## Put it online (needed to use it on your phone)

It's a plain static site, so any HTTPS static host works:
- **GitHub Pages**: push this folder to a repository, Settings → Pages → deploy from the main branch.
- **Netlify / Cloudflare Pages / Vercel / Azure Static Web Apps**: drag and drop the folder, or connect the repository.

Add the resulting URL as a redirect URI in your app registration (step 2).

## Install it as an app

- **Android (Chrome)**: open the URL → menu ⋮ → *Install app* / *Add to Home screen*.
- **iPhone (Safari)**: open the URL → Share → *Add to Home Screen*.
- **PC (Chrome / Edge)**: install icon in the address bar.

## More than one person

Everyone can use the same address with the same Client IDs. Each person signs in with their own account
and only sees their own calendar and tasks. For Google, add each person's Gmail as a test user (or publish
the app); for Microsoft, use a multitenant registration (see microsoft-setup.md).

Only if several people share one device and want to switch without signing out, list them under
`PROFILES` in `config.js` (for example `PROFILES: [{ id: "barak", name: "Barak" }, { id: "dana", name: "Dana" }]`).
The app then asks "Who's using Task Board?" on each device, and each profile keeps its own sign-in,
settings and saved views there.

## Security

- The app is plain files. There is no server and no database, and no password or secret is in the code:
  the Client IDs in `config.js` are public identifiers. Your data goes straight from your browser to
  Microsoft / Google and never passes through GitHub or anyone else.
- Sign-in happens on Microsoft's and Google's own pages, and they only accept sign-ins from the web
  address you registered. The app can act only as the person who signed in, on their own calendar and tasks.
- A Content-Security-Policy in `index.html` lets the page load only its own files and Google's sign-in
  script, and connect only to Microsoft and Google.
- Sign-in tokens are kept in the browser on your device. Use "Sign out" on a shared computer.
- Turn on two-factor sign-in for the GitHub account that hosts it, so nobody else can change the code.

## Known limits

- Calendar events load from 30 days ago to 120 days ahead, from your main calendar
  (change the range in `config.js`).
- If you edit a meeting someone else organized, Outlook may refuse changes to its title or time;
  status changes always work.
- Recurring events appear as individual occurrences; a status applies to that occurrence only.
- For Google items, subtasks, labels, links, files and task repeat/reminders aren't available,
  because the Google Tasks API doesn't support them. The edit window hides those fields for Google items.
- Moving a task between a To Do list and a Google list creates it in the other account and deletes
  the original, so To Do-only details (subtasks, labels, files) don't carry over to Google.
- Google's browser sign-in lasts an hour; after that a yellow banner asks you to tap to reconnect Google,
  and your Outlook items stay on screen meanwhile.

## Files

- `index.html`, `styles.css`, `app.js`: the app screens (no build step).
- `backends.js`: reading and writing Outlook / To Do, Google, both together, and the demo.
- `core.js`: shared helpers.
- `i18n.js`: English and Hebrew interface text.
- `config.js`: Client IDs and calendar settings.
- `vendor/msal-browser.min.js`: Microsoft's sign-in library (MSAL.js 3.28.1, MIT license).
- `manifest.webmanifest`, `sw.js`, `icons/`: make it installable on phone and PC.

## Google setup (to add Google Calendar)

Tasks are Google Tasks and events are Google Calendar events. Done = the task's checkbox. The other
statuses and the priority go in a last line of the task notes (`#taskboard status=working priority=high`),
because Google Tasks has no hidden fields. Event status and priority are hidden event properties.

1. https://console.cloud.google.com/ → create a project.
2. Enable **Google Calendar API** and **Google Tasks API**.
3. OAuth consent screen: External, scopes `calendar.events` and `tasks`, add yourself (and anyone else
   who will use it, e.g. your spouse) as a test user. A company Google Workspace account may also need
   its admin to allow the app.
4. Credentials → OAuth client ID → Web application; add your hosting URL and `http://localhost:8080`
   as Authorized JavaScript origins.
5. Put the Client ID in `GOOGLE_CLIENT_ID`. Google browser sign-in lasts an hour; tap the yellow
   banner to reconnect.
