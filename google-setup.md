# Connect Task Board to your Google Calendar

This takes about 10 minutes and costs nothing. You do it once. After that, any phone or PC can open the app and sign in.

Your app's address: **https://berek-dev.github.io/task-board/**

---

## Step 1: Create a Google Cloud project

1. On your PC, go to **https://console.cloud.google.com/** and sign in with the Google account whose calendar you want to use.
2. If Google asks you to accept the terms, accept them.
3. At the top left, click the **project picker** (next to "Google Cloud"), then **New project**.
4. Enter "Task Board" as the project name and click **Create**.
5. When it's done, make sure **Task Board** is selected in the project picker.

## Step 2: Turn on the Calendar and Tasks APIs

1. Open the menu ☰ and go to **APIs & Services → Library**.
2. Search for **Google Calendar API**, open it and click **Enable**.
3. Go back to the Library, search for **Google Tasks API**, open it and click **Enable**.

## Step 3: Set up the sign-in screen (OAuth consent)

1. Open the menu ☰ and go to **APIs & Services → OAuth consent screen**. It may be called **Google Auth Platform**.
2. Click **Get started**.
   - **App name:** Task Board
   - **User support email:** your email
   - **Audience:** **External**
   - **Contact email:** your email
   - Agree to the policy, then click **Create**.
3. Go to **Audience** (or "Test users"). Under **Test users**, click **Add users** and add your own Gmail address. Add your wife's Google address too if she'll use Google. Click **Save**.
   - Leave the app in **Testing** mode. You don't need to publish it or get it verified for your own family's use. Up to 100 test users are allowed.
4. Go to **Data Access** (or "Scopes"), click **Add or remove scopes**, and in the filter box add these:
   - `https://www.googleapis.com/auth/calendar.events`
   - `https://www.googleapis.com/auth/tasks`
   - `https://www.googleapis.com/auth/calendar.calendarlist.readonly` (lets the app show your shared and subscribed calendars too)

   Tick them, then click **Update** and **Save**.

## Step 4: Create the Client ID

1. Go to **Clients** (or **APIs & Services → Credentials → Create credentials → OAuth client ID**).
2. For **Application type**, choose **Web application**. For **Name**, enter "Task Board web".
3. Under **Authorized JavaScript origins**, click **Add URI** and enter exactly:
   ```
   https://berek-dev.github.io
   ```
   Leave out the `/task-board` part and any slash at the end. Google wants just the site.
4. Leave **Authorized redirect URIs** empty, then click **Create**.
5. Copy the **Client ID**. It looks like `123456789012-abc...xyz.apps.googleusercontent.com`.
   - The Client ID is public. It's fine to paste it in the chat.
   - Don't send the **Client secret** to anyone. The app doesn't use it, so you can ignore it.

## Step 5: Put the Client ID in the app

There are two ways. Pick one.

- **Easiest:** paste the Client ID in our chat and I'll add it and publish the update.
- **Yourself on GitHub:**
  1. Open https://github.com/berek-dev/task-board/blob/main/config.js.
  2. Click the ✏️ pencil icon.
  3. Replace `YOUR_CLIENT_ID.apps.googleusercontent.com` with your Client ID, keeping the quotes.
  4. Click **Commit changes**.

Either way, GitHub Pages updates the site within a minute or two.

## Step 6: Sign in

1. Open **https://berek-dev.github.io/task-board/**. If it still shows the old version, refresh the page or close and reopen the tab.
2. Tap **Sign in with Google**. A Google window opens. If your browser blocks pop-ups, allow them for this site.
3. Choose your account. Google will warn **"Google hasn't verified this app"**. That's expected, because it's your own app in Testing mode. Tap **Continue** (on some screens it's **Advanced → Go to Task Board**).
4. Allow access to your calendar and tasks.

Your Google Tasks lists and Google Calendar events now appear on the board.

## Good to know

- **Reconnecting:** Google's browser sign-in lasts one hour. After that, a yellow banner says the Google connection expired. Tap it to reconnect, which takes one tap.
- **Phone:** open the same address on your phone and use **Add to Home Screen** (Safari: Share button; Chrome: ⋮ menu) to get an app icon.
- **What Google supports:** the status, priority, date, time, notes and event reminders/repeats all sync. Google Tasks has no subtasks, labels, files or task repeats, so the app hides those for Google items.
- **Outlook later:** once the Microsoft side is set up (see README.md), a **+ Outlook** button lets you connect both at once.
- **Work Google account:** if the account is a company Google Workspace account, the company's admin may need to allow the app.

## If something goes wrong

| What you see | What to do |
|---|---|
| "Error 400: redirect_uri_mismatch" or "origin not allowed" | The JavaScript origin in Step 4 must be exactly `https://berek-dev.github.io`. Changes can take a few minutes to start working. |
| "Access blocked: Task Board has not completed the Google verification process" | Your address isn't in **Test users** (Step 3.3). Add it and try again. |
| "This copy isn't set up yet" on the sign-in screen | The Client ID isn't in config.js yet, or the page is still the old version. Refresh after a minute. |
| The Google window flashes and closes | Allow pop-ups for berek-dev.github.io and try again. |
