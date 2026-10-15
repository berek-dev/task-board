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

  // People who use the app. Each profile can have its own Client IDs, and keeps its own sign-in,
  // settings and saved views on each device. Anything set here overrides the values above.
  // With more than one profile, the app asks "Who's using Task Board?" the first time on each device.
  PROFILES: [
    { id: "barak", name: "Barak", GOOGLE_CLIENT_ID: "79752309754-ptmlufk15tf041o2nl20jrk60kd0ft6i.apps.googleusercontent.com" },
    { id: "user2", name: "Second user", GOOGLE_CLIENT_ID: "112739680940-884mqga50ne5f80ubbc2mtostbdegpbb.apps.googleusercontent.com" },
  ],
};
