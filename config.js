// Settings shared by everyone who uses this copy of the app.
// Client IDs are public identifiers, not secrets (see README.md, microsoft-setup.md and google-setup.md).
window.TM_CONFIG = {
  // Microsoft (Outlook Calendar + Microsoft To Do): the Application (client) ID from your
  // Azure / Entra app registration. One registration can serve everyone.
  MS_CLIENT_ID: "YOUR_CLIENT_ID",
  // "common" accepts work and personal Microsoft accounts. Your IT may ask you to put
  // your organization's tenant ID here instead.
  MS_TENANT: "common",

  CALENDAR_ID: "primary",
  // How far back / ahead to load calendar events, in days.
  EVENTS_PAST_DAYS: 30,
  EVENTS_FUTURE_DAYS: 120,

  // Google (Google Calendar + Google Tasks): the OAuth Client ID from Google Cloud (see google-setup.md).
  // One Client ID serves everyone: each person signs in with their own Google account and sees only their own data.
  GOOGLE_CLIENT_ID: "79752309754-ptmlufk15tf041o2nl20jrk60kd0ft6i.apps.googleusercontent.com",

  // Optional: only if several people share one device and want to switch without signing out, list them here,
  // e.g. PROFILES: [{ id: "barak", name: "Barak" }, { id: "dana", name: "Dana" }]. Each profile keeps its own
  // sign-in, settings and saved views on that device.
};
