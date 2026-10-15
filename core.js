/* Shared constants and helpers used by backends.js and app.js. */
(() => {
  "use strict";
  const { t } = window.i18n;

  // Labels are getters so they follow the current interface language.
  const STATUSES = [
    { id: "todo", color: "var(--s-todo)" },
    { id: "working", color: "var(--s-working)" },
    { id: "stuck", color: "var(--s-stuck)" },
    { id: "done", color: "var(--s-done)" },
  ].map((s) => Object.defineProperty(s, "label", { get: () => t("status." + s.id) }));
  const ALL_PRIORITIES = [
    { id: "", color: "var(--p-none)", cls: "p-none" },
    { id: "low", color: "var(--p-low)" },
    { id: "medium", color: "var(--p-medium)" },
    { id: "high", color: "var(--p-high)" },
  ].map((p) => Object.defineProperty(p, "label", { get: () => t("priority." + (p.id || "none")) }));
  // Outlook and To Do only know low / normal / high importance.
  const MS_PRIORITIES = ALL_PRIORITIES.filter((p) => p.id !== "medium");
  const PRIORITY_RANK = { high: 0, medium: 1, low: 2, "": 3 };
  const STATUS_RANK = { todo: 0, working: 1, stuck: 2, done: 3 };
  const REPEATS = ["", "daily", "weekdays", "weekly", "monthly", "yearly"];
  const REMIND_MINUTES = ["", "0", "5", "15", "30", "60", "1440"];
  const EVENTS_GROUP = "calendar";
  // Distinct hues for label dots; text stays in normal ink.
  const LABEL_COLORS = ["#5b5bd6", "#e2445c", "#00a2c7", "#fdab3d", "#00c875", "#a25ddc", "#ff7575", "#579bfc", "#9d6b3a", "#2b8a3e"];

  const pad = (n) => String(n).padStart(2, "0");
  const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const today = () => ymd(new Date());
  const addDays = (dateStr, n) => { const d = new Date(dateStr + "T00:00:00"); d.setDate(d.getDate() + n); return ymd(d); };
  const toMin = (hhmm) => { const [h, m] = (hhmm || "0:0").split(":").map(Number); return h * 60 + m; };
  const fromMin = (min) => `${pad(Math.floor(min / 60))}:${pad(min % 60)}`;
  const plusHour = (hhmm) => (toMin(hhmm) >= 23 * 60 ? "24:00" : fromMin(toMin(hhmm) + 60));
  const localTz = () => Intl.DateTimeFormat().resolvedOptions().timeZone;
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const newId = (prefix = "new:") => prefix + Math.random().toString(36).slice(2, 10);
  const isNewId = (id) => !id || String(id).startsWith("new:");
  const sameJson = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);

  function labelColor(name) {
    let h = 0;
    for (const ch of String(name).toLowerCase()) h = (h * 31 + ch.codePointAt(0)) >>> 0;
    return LABEL_COLORS[h % LABEL_COLORS.length];
  }

  // The date a repeating item moves to after it is completed.
  function nextRepeatDate(date, repeat) {
    const base = date || today();
    const d = new Date(base + "T00:00:00");
    switch (repeat) {
      case "daily": return addDays(base, 1);
      case "weekly": return addDays(base, 7);
      case "weekdays": { let n = addDays(base, 1); while ([0, 6].includes(new Date(n + "T00:00:00").getDay())) n = addDays(n, 1); return n; }
      case "monthly": d.setMonth(d.getMonth() + 1); return ymd(d);
      case "yearly": d.setFullYear(d.getFullYear() + 1); return ymd(d);
      default: return base;
    }
  }

  // Every item has the same shape whichever backend it came from.
  const normalize = (item) => ({
    notes: "", status: "todo", priority: "", date: "", startTime: "", endTime: "", link: "",
    checklist: [], repeat: "", remind: "", isOccurrence: false, labels: [], links: [], hasAttachments: false,
    ...item,
  });

  class AuthError extends Error { constructor() { super("Your session expired"); this.auth = true; } }

  function loadScript(src) {
    return new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = src; s.async = true; s.onload = res;
      s.onerror = () => rej(new Error(t("signin.loadFailed")));
      document.head.appendChild(s);
    });
  }
  function fileToBase64(file) {
    return new Promise((res, rej) => {
      const r = new FileReader();
      r.onload = () => res(String(r.result).split(",")[1] || "");
      r.onerror = () => rej(r.error);
      r.readAsDataURL(file);
    });
  }
  // Profiles: several people can use the app on one device, each with their own Client IDs and their
  // own saved sign-in, settings and views. The first profile keeps the plain "tb.*" keys.
  const GLOBAL_KEYS = new Set(["tb.lang", "tb.profile"]);
  let scope = "";
  const scoped = (k) => (scope && !GLOBAL_KEYS.has(k) ? k.replace(/^tb\./, `tb.${scope}.`) : k);
  const safeStore = {
    get(k) { try { return JSON.parse(localStorage.getItem(scoped(k))); } catch { return null; } },
    set(k, v) { try { localStorage.setItem(scoped(k), JSON.stringify(v)); } catch {} },
    del(k) { try { localStorage.removeItem(scoped(k)); } catch {} },
  };

  // Picks this device's profile from TM_CONFIG.PROFILES and copies its settings into TM_CONFIG,
  // before the app reads them. With several profiles and none chosen yet, the app asks who it is.
  const profiles = Array.isArray(window.TM_CONFIG?.PROFILES) ? window.TM_CONFIG.PROFILES : [];
  let profile = null;
  if (profiles.length) {
    const saved = safeStore.get("tb.profile");
    profile = profiles.find((p) => p.id === saved) || (profiles.length === 1 ? profiles[0] : null);
    if (profile) {
      Object.assign(window.TM_CONFIG, profile);
      scope = profile === profiles[0] ? "" : profile.id;
    }
  }
  const chooseProfile = (id) => { if (id) safeStore.set("tb.profile", id); else safeStore.del("tb.profile"); location.reload(); };

  window.TB = {
    t, STATUSES, ALL_PRIORITIES, MS_PRIORITIES, PRIORITY_RANK, STATUS_RANK, REPEATS, REMIND_MINUTES, EVENTS_GROUP,
    pad, ymd, today, addDays, toMin, fromMin, plusHour, localTz, esc, newId, isNewId, sameJson,
    labelColor, nextRepeatDate, normalize, AuthError, loadScript, fileToBase64, safeStore,
    profiles, profile, chooseProfile,
    MAX_ATTACHMENT_BYTES: 3 * 1024 * 1024,
  };
})();
