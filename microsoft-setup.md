# Connect Task Board to Outlook (Microsoft 365)

The Microsoft button only appears once the app is registered with Microsoft. This takes about 10 minutes and costs nothing. You do it once.

Your app's address: **https://berek-dev.github.io/task-board/**

> **Work account:** your company's IT controls which apps can reach your work Outlook. If you can't create an app registration, or Microsoft says **"Need admin approval"** when you sign in, send IT the note at the bottom of this page.

---

## Step 1: Register the app

1. On your PC, go to **https://entra.microsoft.com** and sign in with your **work** Microsoft account.
2. In the left menu, go to **Identity → Applications → App registrations**, then click **+ New registration**.
3. Fill in the form:
   - **Name:** Task Board
   - **Supported account types:** **Accounts in any organizational directory (Any Microsoft Entra ID tenant – Multitenant)**. This lets your wife sign in with her work account too. If only you will use it, choose *this organizational directory only* instead.
   - **Redirect URI:** set the platform to **Single-page application (SPA)** and the URL to exactly:
     ```
     https://berek-dev.github.io/task-board/
     ```
     Keep the slash at the end.
4. Click **Register**.

If the "New registration" button is greyed out, or you get a permissions error, your company doesn't let employees register apps. Send IT the note below.

## Step 2: Give it calendar and task permissions

1. In your new app, open **API permissions**, click **+ Add a permission**, then **Microsoft Graph**, then **Delegated permissions**.
2. Tick these two:
   - **Calendars.ReadWrite**
   - **Tasks.ReadWrite**

   `User.Read` is already there.
3. Click **Add permissions**.
4. If you see **Grant admin consent for (your company)** and it's clickable, click it. If it's greyed out, that's the step for IT.

## Step 3: Send me the IDs

On the app's **Overview** page, copy:
- **Application (client) ID**
- **Directory (tenant) ID**. Only send this one if you chose "this organizational directory only" in Step 1.

Paste them here and I'll add them to the app. Neither one is a secret. There is no password or client secret, so don't create one.

## Step 4: Sign in

1. Open **https://berek-dev.github.io/task-board/** and refresh the page.
2. If you're already signed in with Google, tap **+ Outlook** in the top bar. Otherwise, tap **Sign in with Microsoft**.
3. Sign in with your work account and accept the permissions.

Your To Do lists and Outlook calendar then appear next to your Google items.

---

## Note to send IT (if needed)

> Hi, I'd like to use a personal task board web app with my work calendar. It's a static web page with no server that I host at https://berek-dev.github.io/task-board/ and that talks directly to Microsoft Graph as me. It needs these **delegated** Microsoft Graph permissions only: `User.Read`, `Calendars.ReadWrite`, `Tasks.ReadWrite`. Could you either approve (grant consent for) the app registration "Task Board", or register it for me as a Single-page application with the redirect URI `https://berek-dev.github.io/task-board/` and send me its Application (client) ID? Thanks!

## If something goes wrong

| What you see | What to do |
|---|---|
| "Need admin approval" | IT has to approve the app. Send them the note above. |
| "AADSTS50011: redirect URI mismatch" | The redirect URI in Step 1 must be exactly `https://berek-dev.github.io/task-board/`, with the platform set to **Single-page application**. |
| "AADSTS50194 … not configured as a multi-tenant application" | Send me the Directory (tenant) ID, or switch the registration to multitenant. |
| No Microsoft button | The Client ID isn't in the app yet, or the page is the old version. Refresh after a minute. |
