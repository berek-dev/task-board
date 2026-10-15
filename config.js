// Which accounts the board can connect. Fill in one or both Client IDs; with both, the sign-in screen
// offers Microsoft and Google, and you can connect both at once (the "+ Google Calendar" / "+ Outlook" button).
window.TM_CONFIG = {
  // Microsoft (Outlook Calendar + Microsoft To Do): the Application (client) ID from your
  // Azure / Entra app registration (see README.md). It is a public identifier, not a secret.
  MS_CLIENT_ID: "YOUR_CLIENT_ID",
  // "common" accepts work and personal Microsoft accounts. Your IT may ask you to put
  // your organization's tenant ID here instead.
  MS_TENANT: "common",

  // Google (Google Calendar + Google Tasks): the OAuth Client ID from Google Cloud (see README.md).
  GOOGLE_CLIENT_ID: "YOUR_CLIENT_ID.apps.googleusercontent.com",
  CALENDAR_ID: "primary",

  // How far back / ahead to load calendar events, in days.
  EVENTS_PAST_DAYS: 30,
  EVENTS_FUTURE_DAYS: 120,
};
