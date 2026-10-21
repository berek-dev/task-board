/* Storage backends. Each one turns calendar events and task-list tasks into the same item shape
 * (see normalize() in core.js) and writes changes back. There is no other storage.
 *
 * Microsoft 365 (Outlook Calendar + Microsoft To Do, via Microsoft Graph):
 *   To Do task -> status, importance, due date, recurrence, reminder, categories (labels),
 *                 checklist items (subtasks), linked resources (links), attachments: all native.
 *   Outlook event -> importance, categories, recurrence, reminder, attachments are native;
 *                    the board status lives in a hidden extended property.
 *
 * Google (Google Calendar + Google Tasks): fewer features, because the Google Tasks API has no
 *   hidden fields, subtask checklists, recurrence or reminders.
 *   Task  -> Done = completed; other statuses + priority in a last line of the notes,
 *            e.g. "#taskboard status=working priority=high".
 *   Event -> status + priority in private extendedProperties; reminders and recurrence are native.
 */
(() => {
  "use strict";
  const {
    t, STATUSES, MS_PRIORITIES, EVENTS_GROUP, pad, ymd, today, addDays, plusHour, localTz,
    newId, isNewId, sameJson, nextRepeatDate, normalize, AuthError, loadScript, fileToBase64, safeStore,
  } = window.TB;
  const cfg = window.TM_CONFIG || {};
  const enc = encodeURIComponent;

  // Start / end in the shape both calendar APIs accept, from the item's date and times.
  function eventTimes(item) {
    const date = item.date || today();
    if (!item.startTime) return { allDay: true, date, start: `${date}T00:00:00`, end: `${addDays(date, 1)}T00:00:00`, endDate: addDays(date, 1) };
    const endTime = item.endTime && item.endTime > item.startTime ? item.endTime : plusHour(item.startTime);
    const endDate = endTime === "24:00" ? addDays(date, 1) : date;
    return { allDay: false, date, start: `${date}T${item.startTime}:00`, end: `${endDate}T${endTime === "24:00" ? "00:00" : endTime}:00` };
  }
  const changed = (prev, item, ...keys) => !prev || keys.some((k) => !sameJson(prev[k] || "", item[k] || ""));

  // Applies the difference between two lists of sub-records (checklist steps, links) with REST calls.
  async function syncSubList(api, base, prevArr, nextArr, toBody, fromResp, same) {
    const prevById = new Map((prevArr || []).map((x) => [x.id, x]));
    for (const p of prevArr || []) if (!nextArr.some((n) => n.id === p.id)) await api("DELETE", `${base}/${enc(p.id)}`);
    const out = [];
    for (const n of nextArr) {
      const p = isNewId(n.id) ? null : prevById.get(n.id);
      if (!p) out.push(fromResp(await api("POST", base, toBody(n))));
      else {
        if (!same(p, n)) await api("PATCH", `${base}/${enc(n.id)}`, toBody(n));
        out.push(n);
      }
    }
    return out;
  }

  /* ================= Google ================= */
  const META_RE = /(?:^|\n)#taskboard ([^\n]*)\s*$/;
  const STORAGE_TOKEN = "tb.token";
  const MS_RANK_EXT = "com.taskboard.meta";
  const STORAGE_MS_ACCOUNT = "tb.msAccount";
  const RRULE = { daily: "FREQ=DAILY", weekdays: "FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR", weekly: "FREQ=WEEKLY", monthly: "FREQ=MONTHLY", yearly: "FREQ=YEARLY" };

  function parseNotes(raw) {
    const text = raw || "";
    const m = text.match(META_RE);
    if (!m) return { notes: text, meta: {} };
    const meta = {};
    for (const part of m[1].trim().split(/\s+/)) { const [k, v] = part.split("="); if (k && v) meta[k] = v; }
    return { notes: text.slice(0, m.index).replace(/\s+$/, ""), meta };
  }
  function buildNotes(notes, status, priority, rank, labels) {
    const parts = [];
    if (status && status !== "todo" && status !== "done") parts.push(`status=${status}`);
    if (priority) parts.push(`priority=${priority}`);
    if (rank != null) parts.push(`rank=${rank}`);
    if (labels?.length) parts.push(`tags=${labels.map(encodeURIComponent).join(",")}`);
    const body = (notes || "").replace(/\s+$/, "");
    if (!parts.length) return body;
    return (body ? body + "\n\n" : "") + "#taskboard " + parts.join(" ");
  }

  class GoogleBackend {
    constructor() {
      this.name = "google";
      this.caps = { eventRepeat: true, eventRemind: true, taskLabels: true, eventLabels: true };
      this.token = null;
      this.calendarId = cfg.CALENDAR_ID || "primary";
    }
    get configured() { return cfg.GOOGLE_CLIENT_ID && !cfg.GOOGLE_CLIENT_ID.startsWith("YOUR_"); }
    async init() {
      const saved = safeStore.get(STORAGE_TOKEN);
      if (saved && saved.exp > Date.now() + 60_000) this.token = saved;
      await loadScript("https://accounts.google.com/gsi/client");
      this.tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: cfg.GOOGLE_CLIENT_ID,
        scope: "https://www.googleapis.com/auth/calendar.events https://www.googleapis.com/auth/calendar.calendarlist.readonly https://www.googleapis.com/auth/tasks",
        callback: () => {},
      });
    }
    get signedIn() { return !!this.token && this.token.exp > Date.now(); }
    // Must be called from a click handler: Google opens a popup.
    signIn() {
      return new Promise((resolve, reject) => {
        this.tokenClient.callback = (resp) => {
          if (resp.error) return reject(new Error(resp.error_description || resp.error));
          this.token = { value: resp.access_token, exp: Date.now() + (resp.expires_in - 60) * 1000 };
          safeStore.set(STORAGE_TOKEN, this.token);
          resolve();
        };
        this.tokenClient.error_callback = (err) => reject(new Error(err.message || err.type || "Sign-in was cancelled"));
        this.tokenClient.requestAccessToken({ prompt: "" });
      });
    }
    signOut() {
      if (this.token) google.accounts.oauth2.revoke(this.token.value, () => {});
      this.token = null;
      safeStore.del(STORAGE_TOKEN);
    }
    async api(method, url, body) {
      if (!this.signedIn) throw new AuthError();
      const res = await fetch(url, {
        method,
        headers: { Authorization: `Bearer ${this.token.value}`, ...(body ? { "Content-Type": "application/json" } : {}) },
        body: body ? JSON.stringify(body) : undefined,
      });
      if (res.status === 401) { this.token = null; safeStore.del(STORAGE_TOKEN); throw new AuthError(); }
      if (!res.ok) {
        let msg = `${res.status}`;
        try { msg = (await res.json()).error.message; } catch {}
        throw new Error(msg);
      }
      return res.status === 204 ? null : res.json().catch(() => null);
    }
    async pages(url) {
      const out = [];
      let token = "";
      do {
        const sep = url.includes("?") ? "&" : "?";
        const data = await this.api("GET", url + (token ? `${sep}pageToken=${enc(token)}` : ""));
        out.push(...(data.items || []));
        token = data.nextPageToken;
      } while (token);
      return out;
    }
    taskUrl(listId, id) { return `https://tasks.googleapis.com/tasks/v1/lists/${enc(listId)}/tasks` + (id ? `/${enc(id)}` : ""); }
    eventUrl(id, calId) { return `https://www.googleapis.com/calendar/v3/calendars/${enc(calId || this.calendarId)}/events` + (id ? `/${enc(id)}` : ""); }
    // Every calendar ticked in Google Calendar's sidebar (yours, shared, subscribed).
    // Falls back to the main one if the calendar-list permission wasn't granted.
    async calendars() {
      try {
        const all = await this.pages("https://www.googleapis.com/calendar/v3/users/me/calendarList?minAccessRole=reader&maxResults=250");
        const cals = all.filter((c) => c.primary || (c.selected && !c.hidden && !c.deleted));
        cals.sort((a, b) => (b.primary ? 1 : 0) - (a.primary ? 1 : 0));
        if (cals.length) return cals.map((c) => ({ id: c.primary ? this.calendarId : c.id, name: c.summaryOverride || c.summary, readOnly: !["owner", "writer"].includes(c.accessRole), primary: !!c.primary }));
      } catch (e) { if (e.auth) throw e; }
      return [{ id: this.calendarId, primary: true, readOnly: false }];
    }

    async load() {
      const lists = await this.pages("https://tasks.googleapis.com/tasks/v1/users/@me/lists?maxResults=100");
      const now = Date.now();
      const min = new Date(now - (cfg.EVENTS_PAST_DAYS ?? 30) * 864e5).toISOString();
      const max = new Date(now + (cfg.EVENTS_FUTURE_DAYS ?? 120) * 864e5).toISOString();
      const cals = await this.calendars();
      const [taskSets, eventSets] = await Promise.all([
        Promise.all(lists.map((l) => this.pages(this.taskUrl(l.id) + "?showCompleted=true&showHidden=true&maxResults=100"))),
        // One calendar that can't be read (e.g. removed sharing) shouldn't hide the others.
        Promise.all(cals.map((c) => this.pages(this.eventUrl("", c.id) + `?singleEvents=true&orderBy=startTime&maxResults=250&timeMin=${enc(min)}&timeMax=${enc(max)}`)
          .catch((e) => { if (c.primary || e.auth) throw e; console.warn("Calendar skipped:", c.name, e.message); return []; }))),
      ]);
      const groups = lists.map((l) => ({ id: l.id, title: l.title, source: "task" }));
      groups.push({ id: EVENTS_GROUP, titleKey: "group.calendar", source: "event" });
      const items = [];
      lists.forEach((l, i) => taskSets[i].forEach((x) => { if (!x.deleted && x.title) items.push(gTaskToItem(x, l.id)); }));
      const seen = new Set();
      cals.forEach((c, i) => eventSets[i].forEach((e) => {
        // The same meeting can sit on two of your calendars; show it once.
        const dup = (e.iCalUID || e.id) + "@" + (e.start?.dateTime || e.start?.date);
        if (e.status === "cancelled" || seen.has(dup)) return;
        seen.add(dup);
        items.push(gEventToItem(e, c));
      }));
      return { groups, items };
    }
    async create(item) {
      if (item.source === "task") return gTaskToItem(await this.api("POST", this.taskUrl(item.containerId), itemToGTask(item)), item.containerId);
      return gEventToItem(await this.api("POST", this.eventUrl(), itemToGEvent(null, item)));
    }
    async update(prev, item) {
      if (prev.source !== item.source || (item.source === "task" && prev.containerId !== item.containerId)) {
        const created = await this.create(item);
        await this.remove(prev);
        return created;
      }
      if (item.source === "task") return gTaskToItem(await this.api("PATCH", this.taskUrl(item.containerId, item.id), itemToGTask(item)), item.containerId);
      if (prev.readOnly) throw new Error(t("event.readOnly", { cal: prev.calName || "" }));
      return gEventToItem(await this.api("PATCH", this.eventUrl(item.id, prev.calId), itemToGEvent(prev, item)), prev.calId ? { id: prev.calId, name: prev.calName } : null);
    }
    async remove(item) {
      if (item.source === "event" && item.readOnly) throw new Error(t("event.readOnly", { cal: item.calName || "" }));
      await this.api("DELETE", item.source === "task" ? this.taskUrl(item.containerId, item.id) : this.eventUrl(item.id, item.calId));
    }
  }
  function gTaskToItem(x, listId) {
    const { notes, meta } = parseNotes(x.notes);
    const done = x.status === "completed";
    return normalize({
      key: "task:" + x.id, id: x.id, source: "task", containerId: listId,
      title: x.title || "", notes,
      status: done ? "done" : (meta.status && meta.status !== "done" ? meta.status : "todo"),
      priority: meta.priority || "",
      rank: meta.rank != null && isFinite(meta.rank) ? Number(meta.rank) : null,
      labels: meta.tags ? meta.tags.split(",").map((x) => { try { return decodeURIComponent(x); } catch { return x; } }).filter(Boolean) : [],
      date: x.due ? x.due.slice(0, 10) : "", link: x.webViewLink || "",
    });
  }
  function itemToGTask(item) {
    return {
      title: item.title,
      notes: buildNotes(item.notes, item.status, item.priority, item.rank, item.labels),
      status: item.status === "done" ? "completed" : "needsAction",
      due: item.date ? item.date + "T00:00:00.000Z" : null,
    };
  }
  function gEventToItem(e, cal) {
    const p = e.extendedProperties?.private || {};
    let date = "", startTime = "", endTime = "";
    if (e.start?.date) date = e.start.date;
    else if (e.start?.dateTime) {
      const s = new Date(e.start.dateTime), en = new Date(e.end?.dateTime || e.start.dateTime);
      date = ymd(s);
      startTime = `${pad(s.getHours())}:${pad(s.getMinutes())}`;
      endTime = `${pad(en.getHours())}:${pad(en.getMinutes())}`;
    }
    const override = e.reminders && !e.reminders.useDefault ? e.reminders.overrides?.[0] : null;
    return normalize({
      key: "event:" + (cal && !cal.primary ? cal.id + "/" : "") + e.id, id: e.id, source: "event", containerId: EVENTS_GROUP,
      ...(cal && !cal.primary ? { calId: cal.id, calName: cal.name || "", readOnly: !!cal.readOnly } : {}),
      title: e.summary || t("item.noTitle"), notes: e.description || "",
      status: STATUSES.some((s) => s.id === p.tbStatus) ? p.tbStatus : "todo",
      priority: p.tbPriority || "",
      labels: (() => { try { const l = JSON.parse(p.tbLabels || "[]"); return Array.isArray(l) ? l : []; } catch { return []; } })(),
      date, startTime, endTime, link: e.htmlLink || "",
      isOccurrence: !!e.recurringEventId,
      remind: override ? String(override.minutes) : "",
    });
  }
  function itemToGEvent(prev, item) {
    const tm = eventTimes(item), tz = localTz();
    const out = {
      summary: item.title, description: item.notes || "",
      start: tm.allDay ? { date: tm.date } : { dateTime: tm.start, timeZone: tz },
      end: tm.allDay ? { date: tm.endDate } : { dateTime: tm.end, timeZone: tz },
      extendedProperties: { private: { tbStatus: item.status, tbPriority: item.priority || "", tbLabels: JSON.stringify(item.labels || []) } },
    };
    if (changed(prev, item, "remind")) {
      out.reminders = item.remind ? { useDefault: false, overrides: [{ method: "popup", minutes: Number(item.remind) }] } : { useDefault: true };
    }
    if (!item.isOccurrence && changed(prev, item, "repeat")) out.recurrence = item.repeat ? ["RRULE:" + RRULE[item.repeat]] : [];
    return out;
  }

  /* ================= Microsoft 365 ================= */
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const MS_SCOPES = ["User.Read", "Calendars.ReadWrite", "Tasks.ReadWrite"];
  const MS_STATUS_PROP = "String {6f3e1c52-9a4b-4d1e-8c7a-2b5d9e0f4a13} Name tbStatus";
  const TO_MS_STATUS = { todo: "notStarted", working: "inProgress", stuck: "waitingOnOthers", done: "completed" };
  const FROM_MS_STATUS = { notStarted: "todo", inProgress: "working", waitingOnOthers: "stuck", deferred: "stuck", completed: "done" };
  const DAYS = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];

  function appBaseUrl() { return location.origin + location.pathname.replace(/index\.html$/, ""); }
  function htmlToText(html) {
    const doc = new DOMParser().parseFromString(html || "", "text/html");
    doc.querySelectorAll("br").forEach((br) => br.replaceWith("\n"));
    doc.querySelectorAll("p,div,li,tr").forEach((el) => el.append("\n"));
    return (doc.body.textContent || "").replace(/ /g, " ").replace(/\n{3,}/g, "\n\n").trim();
  }
  const bodyText = (b) => (!b ? "" : b.contentType === "html" ? htmlToText(b.content) : (b.content || "").trim());
  const msPriority = (importance) => (importance === "low" || importance === "high" ? importance : "");
  const msImportance = (priority) => (priority === "low" || priority === "high" ? priority : "normal");

  function msDueDate(due) {
    if (!due?.dateTime) return "";
    const dt = due.dateTime;
    // To Do sometimes stores local midnight converted to UTC (e.g. 21:00 the day before).
    if (/^utc$/i.test(due.timeZone || "") && dt.slice(11, 16) !== "00:00") return ymd(new Date(dt.slice(0, 19) + "Z"));
    return dt.slice(0, 10);
  }
  // A To Do reminder as the board stores it: "HH:MM" when it is on the due date,
  // otherwise a one-time "YYYY-MM-DDTHH:MM" in local time.
  function msReminder(rem, dueDate) {
    if (!rem?.dateTime) return "";
    const dt = rem.dateTime;
    let local = dt.slice(0, 16);
    if (/^utc$/i.test(rem.timeZone || "")) {
      const d = new Date(dt.slice(0, 19) + "Z");
      local = `${ymd(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
    }
    return dueDate && local.slice(0, 10) === dueDate ? local.slice(11, 16) : local;
  }
  function msRepeat(rec) {
    const p = rec?.pattern;
    if (!p) return "";
    if (p.type === "daily") return "daily";
    if (p.type === "weekly") {
      const d = p.daysOfWeek || [];
      return d.length === 5 && !d.includes("saturday") && !d.includes("sunday") ? "weekdays" : "weekly";
    }
    if (/monthly/i.test(p.type)) return "monthly";
    if (/yearly/i.test(p.type)) return "yearly";
    return "";
  }
  function msRecurrence(repeat, date) {
    if (!repeat) return null;
    const startDate = date || today();
    const d = new Date(startDate + "T00:00:00");
    const pattern = {
      daily: { type: "daily", interval: 1 },
      weekdays: { type: "weekly", interval: 1, daysOfWeek: DAYS.slice(1, 6), firstDayOfWeek: "sunday" },
      weekly: { type: "weekly", interval: 1, daysOfWeek: [DAYS[d.getDay()]], firstDayOfWeek: "sunday" },
      monthly: { type: "absoluteMonthly", interval: 1, dayOfMonth: d.getDate() },
      yearly: { type: "absoluteYearly", interval: 1, dayOfMonth: d.getDate(), month: d.getMonth() + 1 },
    }[repeat];
    return { pattern, range: { type: "noEnd", startDate } };
  }

  class MicrosoftBackend {
    constructor() {
      this.name = "microsoft";
      this.priorities = MS_PRIORITIES;
      this.caps = {
        checklist: true, taskRepeat: true, taskRemind: true, eventRepeat: true, eventRemind: true,
        taskLabels: true, eventLabels: true, taskLinks: true, attachments: true,
      };
      this.account = null;
    }
    get configured() { return cfg.MS_CLIENT_ID && !cfg.MS_CLIENT_ID.startsWith("YOUR_"); }
    async init() {
      await loadScript("vendor/msal-browser.min.js");
      this.pca = new msal.PublicClientApplication({
        auth: {
          clientId: cfg.MS_CLIENT_ID,
          authority: `https://login.microsoftonline.com/${cfg.MS_TENANT || "common"}`,
          redirectUri: cfg.MS_REDIRECT_URI || appBaseUrl(),
          postLogoutRedirectUri: appBaseUrl(),
        },
        cache: { cacheLocation: "localStorage" },
      });
      await this.pca.initialize();
      // Finishes a sign-in that came back from the Microsoft login page.
      const res = await this.pca.handleRedirectPromise();
      // Remembers which Microsoft account belongs to this profile, so two people on one device stay apart.
      if (res?.account) safeStore.set(STORAGE_MS_ACCOUNT, res.account.homeAccountId);
      const mine = safeStore.get(STORAGE_MS_ACCOUNT);
      this.account = res?.account || (mine && this.pca.getAccountByHomeId(mine))
        || (window.TB.profiles.length > 1 ? null : this.pca.getAllAccounts()[0]) || null;
    }
    get signedIn() { return !!this.account; }
    // Redirects to the Microsoft login page and comes back signed in.
    signIn() {
      return this.account
        ? this.pca.acquireTokenRedirect({ scopes: MS_SCOPES, account: this.account })
        : this.pca.loginRedirect({ scopes: MS_SCOPES, prompt: "select_account" });
    }
    signOut() {
      const account = this.account;
      this.account = null;
      safeStore.del(STORAGE_MS_ACCOUNT);
      // Clears this app's session only; stays signed in to Microsoft elsewhere.
      this.pca.logoutRedirect({ account, onRedirectNavigate: () => false }).catch(() => {});
    }
    async api(method, path, body, { raw = false } = {}) {
      let token;
      try {
        token = (await this.pca.acquireTokenSilent({ scopes: MS_SCOPES, account: this.account })).accessToken;
      } catch (e) {
        if (e instanceof msal.InteractionRequiredAuthError) throw new AuthError();
        throw e;
      }
      const res = await fetch(path.startsWith("https://") ? path : GRAPH + path, {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          Prefer: `outlook.timezone="${localTz()}"`,
          ...(body ? { "Content-Type": "application/json" } : {}),
        },
        body: body ? JSON.stringify(body) : undefined,
      });
      if (res.status === 401) throw new AuthError();
      if (!res.ok) {
        let msg = `${res.status}`;
        try { msg = (await res.json()).error.message; } catch {}
        throw new Error(msg);
      }
      if (raw) return res;
      return res.status === 204 ? null : res.json().catch(() => null);
    }
    async pages(path) {
      const out = [];
      let next = path;
      while (next) {
        const data = await this.api("GET", next);
        out.push(...(data.value || []));
        next = data["@odata.nextLink"];
      }
      return out;
    }
    taskBase(listId, id) { return `/me/todo/lists/${enc(listId)}/tasks` + (id ? `/${enc(id)}` : ""); }
    itemBase(item) { return item.source === "task" ? this.taskBase(item.containerId, item.id) : `/me/events/${enc(item.id)}`; }

    async loadTasks(listId) {
      const base = this.taskBase(listId) + "?$top=100";
      const tries = [`&$expand=checklistItems,linkedResources,extensions($filter=id eq '${MS_RANK_EXT}')`, "&$expand=checklistItems,linkedResources", ""];
      for (let i = 0; ; i++) {
        try { return await this.pages(base + tries[i]); }
        catch (e) { if (e.auth || i === tries.length - 1) throw e; }
      }
    }
    // The order from the Priorities tab, kept in a small hidden extension on the To Do task.
    async saveRank(listId, taskId, rank, hadRank) {
      const base = this.taskBase(listId, taskId) + "/extensions";
      const patch = () => this.api("PATCH", `${base}/${enc(MS_RANK_EXT)}`, { rank });
      const post = () => this.api("POST", base, { "@odata.type": "microsoft.graph.openTypeExtension", extensionName: MS_RANK_EXT, rank });
      try { await (hadRank ? patch() : post()); }
      catch (e) { if (e.auth) throw e; await (hadRank ? post() : patch()); }
    }
    async load() {
      const lists = await this.pages("/me/todo/lists");
      lists.sort((a, b) => (b.wellknownListName === "defaultList") - (a.wellknownListName === "defaultList"));
      const now = Date.now();
      const from = new Date(now - (cfg.EVENTS_PAST_DAYS ?? 30) * 864e5).toISOString();
      const to = new Date(now + (cfg.EVENTS_FUTURE_DAYS ?? 120) * 864e5).toISOString();
      const evPath = `/me/calendarView?startDateTime=${enc(from)}&endDateTime=${enc(to)}&$top=200&$orderby=start/dateTime` +
        `&$select=id,subject,body,start,end,isAllDay,importance,webLink,isCancelled,categories,hasAttachments,type,isReminderOn,reminderMinutesBeforeStart` +
        `&$expand=singleValueExtendedProperties($filter=id eq '${MS_STATUS_PROP}')`;
      const [taskSets, events] = await Promise.all([Promise.all(lists.map((l) => this.loadTasks(l.id))), this.pages(evPath)]);
      const groups = lists.map((l) => ({ id: l.id, title: l.displayName, source: "task" }));
      groups.push({ id: EVENTS_GROUP, titleKey: "group.outlook", source: "event" });
      const items = [];
      lists.forEach((l, i) => taskSets[i].forEach((x) => items.push(msTaskToItem(x, l.id))));
      events.forEach((e) => { if (!e.isCancelled) items.push(msEventToItem(e)); });
      return { groups, items };
    }

    async syncTaskChildren(listId, taskId, prev, item) {
      const base = this.taskBase(listId, taskId);
      const api = this.api.bind(this);
      const checklist = !prev || !sameJson(prev.checklist, item.checklist)
        ? await syncSubList(api, base + "/checklistItems", prev?.checklist || [], item.checklist || [],
          (c) => ({ displayName: c.text, isChecked: !!c.done }),
          (r) => ({ id: r.id, text: r.displayName, done: !!r.isChecked }),
          (a, b) => a.text === b.text && !!a.done === !!b.done)
        : item.checklist;
      const links = !prev || !sameJson(prev.links, item.links)
        ? await syncSubList(api, base + "/linkedResources", prev?.links || [], item.links || [],
          (l) => ({ webUrl: l.url, displayName: l.title || l.url, applicationName: "Task Board" }),
          (r) => ({ id: r.id, url: r.webUrl, title: r.displayName || "" }),
          (a, b) => a.url === b.url && (a.title || "") === (b.title || ""))
        : item.links;
      return { checklist, links };
    }
    async create(item) {
      if (item.source === "task") {
        const x = await this.api("POST", this.taskBase(item.containerId), itemToMsTask(null, item));
        if (item.rank != null) await this.saveRank(item.containerId, x.id, item.rank, false);
        return { ...msTaskToItem(x, item.containerId), rank: item.rank ?? null, ...(await this.syncTaskChildren(item.containerId, x.id, null, item)) };
      }
      const e = await this.api("POST", "/me/events", itemToMsEvent(null, item));
      return { ...msEventToItem(e), status: item.status };
    }
    async update(prev, item) {
      if (prev.source !== item.source || (item.source === "task" && prev.containerId !== item.containerId)) {
        const created = await this.create({ ...item, checklist: (item.checklist || []).map((c) => ({ ...c, id: newId() })), links: (item.links || []).map((l) => ({ ...l, id: newId() })) });
        await this.remove(prev);
        return created;
      }
      if (item.source === "task") {
        const body = itemToMsTask(prev, item);
        const saved = Object.keys(body).length
          ? msTaskToItem(await this.api("PATCH", this.taskBase(item.containerId, item.id), body), item.containerId)
          : { ...item };
        if (item.rank != null && item.rank !== prev.rank) await this.saveRank(item.containerId, item.id, item.rank, prev.rank != null);
        return { ...saved, rank: item.rank ?? null, ...(await this.syncTaskChildren(item.containerId, item.id, prev, item)) };
      }
      const body = itemToMsEvent(prev, item);
      if (!Object.keys(body).length) return { ...item };
      return { ...msEventToItem(await this.api("PATCH", `/me/events/${enc(item.id)}`, body)), status: item.status, isOccurrence: item.isOccurrence };
    }
    async remove(item) { await this.api("DELETE", this.itemBase(item)); }

    async listAttachments(item) {
      const list = await this.pages(this.itemBase(item) + "/attachments?$select=id,name,size,contentType");
      return list.map((a) => ({ id: a.id, name: a.name, size: a.size, contentType: a.contentType }));
    }
    async addAttachment(item, file) {
      const body = {
        "@odata.type": item.source === "task" ? "#microsoft.graph.taskFileAttachment" : "#microsoft.graph.fileAttachment",
        name: file.name, contentType: file.type || "application/octet-stream", contentBytes: await fileToBase64(file),
      };
      const a = await this.api("POST", this.itemBase(item) + "/attachments", body);
      return { id: a.id, name: a.name, size: a.size ?? file.size, contentType: a.contentType };
    }
    async removeAttachment(item, att) { await this.api("DELETE", this.itemBase(item) + `/attachments/${enc(att.id)}`); }
    async downloadAttachment(item, att) {
      const res = await this.api("GET", this.itemBase(item) + `/attachments/${enc(att.id)}/$value`, null, { raw: true });
      return res.blob();
    }
  }

  function msTaskToItem(x, listId) {
    const date = msDueDate(x.dueDateTime);
    return normalize({
      key: "task:" + x.id, id: x.id, source: "task", containerId: listId,
      title: x.title || "", notes: bodyText(x.body),
      status: FROM_MS_STATUS[x.status] || "todo",
      priority: msPriority(x.importance),
      date,
      checklist: (x.checklistItems || []).map((c) => ({ id: c.id, text: c.displayName || "", done: !!c.isChecked })),
      links: (x.linkedResources || []).filter((r) => r.webUrl).map((r) => ({ id: r.id, url: r.webUrl, title: r.displayName || "" })),
      labels: x.categories || [],
      repeat: msRepeat(x.recurrence),
      remind: x.isReminderOn ? msReminder(x.reminderDateTime, date) : "",
      rank: (() => { const r = (x.extensions || []).find((e) => e.id === MS_RANK_EXT || e.extensionName === MS_RANK_EXT)?.rank; return typeof r === "number" ? r : null; })(),
      hasAttachments: !!x.hasAttachments,
    });
  }
  function itemToMsTask(prev, item) {
    const out = {};
    if (changed(prev, item, "title")) out.title = item.title;
    if (changed(prev, item, "status")) out.status = TO_MS_STATUS[item.status] || "notStarted";
    if (changed(prev, item, "priority")) out.importance = msImportance(item.priority);
    if (changed(prev, item, "date")) out.dueDateTime = item.date ? { dateTime: `${item.date}T00:00:00`, timeZone: localTz() } : null;
    if (changed(prev, item, "notes")) out.body = { contentType: "text", content: item.notes || "" };
    if (changed(prev, item, "labels")) out.categories = item.labels || [];
    if (changed(prev, item, "repeat") || (item.repeat && changed(prev, item, "date"))) out.recurrence = msRecurrence(item.repeat, item.date);
    // A one-time reminder ("YYYY-MM-DDTHH:MM") stays put; a due-date reminder ("HH:MM") moves with the date.
    const oneTime = (item.remind || "").includes("T");
    if (changed(prev, item, "remind") || (item.remind && !oneTime && changed(prev, item, "date"))) {
      out.isReminderOn = !!item.remind;
      if (item.remind) out.reminderDateTime = { dateTime: oneTime ? `${item.remind}:00` : `${item.date || today()}T${item.remind}:00`, timeZone: localTz() };
    }
    return out;
  }
  function msEventToItem(e) {
    const s = e.start?.dateTime || "", en = e.end?.dateTime || s;
    const status = (e.singleValueExtendedProperties || [])[0]?.value;
    return normalize({
      key: "event:" + e.id, id: e.id, source: "event", containerId: EVENTS_GROUP,
      title: e.subject || t("item.noTitle"), notes: bodyText(e.body),
      status: STATUSES.some((x) => x.id === status) ? status : "todo",
      priority: msPriority(e.importance),
      date: s.slice(0, 10),
      startTime: e.isAllDay ? "" : s.slice(11, 16),
      endTime: e.isAllDay ? "" : en.slice(11, 16),
      link: e.webLink || "",
      labels: e.categories || [],
      isOccurrence: e.type === "occurrence" || e.type === "exception",
      repeat: e.type === "seriesMaster" ? msRepeat(e.recurrence) : "",
      remind: e.isReminderOn ? String(e.reminderMinutesBeforeStart ?? 15) : "",
      hasAttachments: !!e.hasAttachments,
    });
  }
  // Only fields that changed are sent: keeps rich-text bodies intact and avoids editing
  // fields you can't change on meetings someone else organized.
  function itemToMsEvent(prev, item) {
    const out = {};
    if (changed(prev, item, "title")) out.subject = item.title;
    if (changed(prev, item, "notes")) out.body = { contentType: "text", content: item.notes || "" };
    if (changed(prev, item, "priority")) out.importance = msImportance(item.priority);
    if (changed(prev, item, "labels")) out.categories = item.labels || [];
    if (changed(prev, item, "status")) out.singleValueExtendedProperties = [{ id: MS_STATUS_PROP, value: item.status }];
    if (changed(prev, item, "remind")) {
      out.isReminderOn = !!item.remind;
      if (item.remind) out.reminderMinutesBeforeStart = Number(item.remind);
    }
    if (changed(prev, item, "date", "startTime", "endTime")) {
      const tm = eventTimes(item), tz = localTz();
      out.isAllDay = tm.allDay;
      out.start = { dateTime: tm.start, timeZone: tz };
      out.end = { dateTime: tm.end, timeZone: tz };
    }
    if (!item.isOccurrence && changed(prev, item, "repeat")) out.recurrence = msRecurrence(item.repeat, item.date);
    return out;
  }

  /* ================= Demo (in this browser only, for trying the app) ================= */
  class DemoBackend {
    constructor() {
      this.name = "demo";
      this.signedIn = true;
      this.caps = {
        checklist: true, taskRepeat: true, taskRemind: true, eventRepeat: true, eventRemind: true,
        taskLabels: true, eventLabels: true, taskLinks: true, attachments: true,
      };
      this.files = new Map();
      const d = (n) => addDays(today(), n);
      this.groups = [
        { id: "work", titleKey: "demo.work", source: "task" },
        { id: "home", titleKey: "demo.home", source: "task" },
        { id: EVENTS_GROUP, titleKey: "group.calendar", source: "event" },
      ];
      let n = 0;
      const mk = (source, containerId, title, status, priority, date, extra = {}) =>
        normalize({ key: `${source}:demo${++n}`, id: `demo${n}`, source, containerId, title, status, priority, date, ...extra });
      const c = (text, done) => ({ id: newId("c"), text, done });
      this.items = [
        mk("task", "work", t("demo.1"), "working", "high", d(2), {
          notes: t("demo.1n"), labels: [window.TB.joinTag("Work", t("demo.lQ4"))],
          checklist: [c(t("demo.c1"), true), c(t("demo.c2"), true), c(t("demo.c3"), false), c(t("demo.c4"), false)],
          links: [{ id: newId("l"), url: "https://example.com/q4-report", title: t("demo.link1") }],
        }),
        mk("task", "work", t("demo.2"), "stuck", "medium", d(-1), { labels: [window.TB.joinTag("Work", t("demo.lClients"))] }),
        mk("task", "work", t("demo.3"), "todo", "low", d(6), { labels: [window.TB.joinTag("Work", t("demo.lQ4"))] }),
        mk("task", "work", t("demo.4"), "done", "medium", d(-3), { repeat: "monthly" }),
        mk("task", "home", t("demo.5"), "todo", "medium", d(4), { remind: "09:00" }),
        mk("task", "home", t("demo.6"), "working", "high", d(1), { labels: [window.TB.joinTag("Personal", t("demo.lFamily"))] }),
        mk("task", "home", t("demo.7"), "todo", "", ""),
        mk("task", "home", t("demo.12"), "todo", "", d(0), { repeat: "weekly", remind: "08:00" }),
        mk("event", EVENTS_GROUP, t("demo.8"), "todo", "", d(1), { startTime: "10:00", endTime: "10:45", isOccurrence: true, remind: "15" }),
        mk("event", EVENTS_GROUP, t("demo.9"), "todo", "medium", d(3), { startTime: "16:30", endTime: "17:15", remind: "60" }),
        mk("event", EVENTS_GROUP, t("demo.10"), "working", "high", d(9), { labels: [window.TB.joinTag("Work", t("demo.lQ4"))] }),
        mk("event", EVENTS_GROUP, t("demo.11"), "done", "", d(-2), { startTime: "14:00", endTime: "15:00", labels: [window.TB.joinTag("Work", t("demo.lClients"))] }),
        mk("event", EVENTS_GROUP, t("demo.13"), "todo", "", d(0), { startTime: "12:30", endTime: "13:30" }),
      ];
    }
    async init() {}
    async load() { return { groups: this.groups, items: this.items.map((i) => structuredClone(i)) }; }
    fix(item) {
      const it = structuredClone(item);
      it.containerId = it.source === "event" ? EVENTS_GROUP : it.containerId;
      if (it.source === "event" && !it.date) it.date = today();
      it.checklist = (it.checklist || []).map((c) => (isNewId(c.id) ? { ...c, id: newId("c") } : c));
      it.links = (it.links || []).map((l) => (isNewId(l.id) ? { ...l, id: newId("l") } : l));
      return it;
    }
    async create(item) {
      const it = this.fix({ ...item, id: "demo" + Date.now() + Math.random().toString(36).slice(2, 5) });
      it.key = `${it.source}:${it.id}`;
      this.items.push(it);
      return structuredClone(it);
    }
    async update(prev, item) {
      const it = this.fix(item);
      if (prev.source !== item.source) it.id = "demo" + Date.now();
      it.key = `${it.source}:${it.id}`;
      this.items = this.items.filter((i) => i.key !== prev.key);
      this.items.push(it);
      // Like To Do: completing a repeating task creates the next one.
      if (it.source === "task" && it.repeat && it.status === "done" && prev.status !== "done") {
        await this.create({ ...it, status: "todo", date: nextRepeatDate(it.date, it.repeat), remind: (it.remind || "").includes("T") ? "" : it.remind, checklist: it.checklist.map((c) => ({ ...c, id: newId(), done: false })) });
        it.repeat = "";
      }
      return structuredClone(it);
    }
    async remove(item) { this.items = this.items.filter((i) => i.key !== item.key); }
    signOut() {}
    async listAttachments(item) { return (this.files.get(item.key) || []).map(({ blob, ...a }) => a); }
    async addAttachment(item, file) {
      const a = { id: newId("a"), name: file.name, size: file.size, contentType: file.type, blob: file };
      this.files.set(item.key, [...(this.files.get(item.key) || []), a]);
      const it = this.items.find((i) => i.key === item.key); if (it) it.hasAttachments = true;
      const { blob, ...meta } = a;
      return meta;
    }
    async removeAttachment(item, att) {
      const left = (this.files.get(item.key) || []).filter((a) => a.id !== att.id);
      this.files.set(item.key, left);
      const it = this.items.find((i) => i.key === item.key); if (it) it.hasAttachments = left.length > 0;
    }
    async downloadAttachment(item, att) { return (this.files.get(item.key) || []).find((a) => a.id === att.id)?.blob; }
  }

  /* ================= Both accounts at once ================= */
  // Shows Outlook / To Do and Google Calendar / Tasks side by side. Group ids and item keys get an
  // "<account>|" prefix so they can't collide, and every call is routed to the account that owns the item.
  const STORAGE_ACCOUNTS = "tb.accounts";
  const ACCOUNT_BADGE = { microsoft: "To Do", google: "Google" };
  const EVENT_GROUP_TITLE = { microsoft: "group.outlook", google: "group.google" };
  const SEP = "|";

  class CombinedBackend {
    constructor(children) {
      this.name = "combined";
      this.children = children; // { microsoft, google }, only the configured ones
      const saved = safeStore.get(STORAGE_ACCOUNTS);
      this.accounts = (Array.isArray(saved) ? saved : []).filter((a) => children[a]);
      this.reauth = null; // account whose sign-in expired
    }
    get configured() { return Object.keys(this.children).length > 0; }
    async init() {
      let firstError = null;
      for (const [name, b] of Object.entries(this.children)) {
        // One provider failing to load (e.g. its sign-in script is blocked) shouldn't block the other.
        try { await b.init(); }
        catch (e) { firstError = firstError || e; delete this.children[name]; this.accounts = this.accounts.filter((a) => a !== name); continue; }
        // Picks up a Microsoft sign-in that just came back from the login page.
        if (b.signedIn && !this.accounts.includes(name)) this.accounts.push(name);
      }
      if (!this.configured && firstError) throw firstError;
      this.save();
    }
    save() { safeStore.set(STORAGE_ACCOUNTS, this.accounts); }
    get signedIn() { return this.accounts.length > 0; }
    get priorities() { return this.accounts.includes("microsoft") ? MS_PRIORITIES : undefined; }
    // Union of what the connected accounts support; capsFor() gives one item's account.
    get caps() {
      const out = {};
      for (const a of this.accounts) for (const [k, v] of Object.entries(this.children[a].caps)) out[k] = out[k] || v;
      return out;
    }
    capsFor(x) { return this.children[this.accountOf(x)]?.caps || this.caps; }
    accountOf(x) {
      if (x && typeof x === "object") return x.account || this.accountOf(x.containerId);
      const i = String(x || "").indexOf(SEP);
      return i > 0 ? String(x).slice(0, i) : this.accounts[0];
    }
    canConnect(name) { return !!this.children[name] && !this.accounts.includes(name); }
    // Connects (or reconnects) one account. Google must be called straight from a click.
    async signIn(name = this.reauth || this.accounts[0] || Object.keys(this.children)[0]) {
      const b = this.children[name];
      if (!this.accounts.includes(name)) { this.accounts.push(name); this.save(); }
      try { await b.signIn(); }
      catch (e) { if (!b.signedIn) { this.accounts = this.accounts.filter((a) => a !== name); this.save(); } throw e; }
      if (this.reauth === name) this.reauth = null;
    }
    signOut() {
      for (const a of this.accounts) this.children[a].signOut();
      this.accounts = []; this.reauth = null; this.save();
    }

    wrap(acc, it) {
      return { ...it, account: acc, key: acc + SEP + it.key, containerId: acc + SEP + it.containerId };
    }
    strip(it) {
      const un = (v) => (typeof v === "string" && v.includes(SEP) ? v.slice(v.indexOf(SEP) + 1) : v);
      const { account, ...rest } = it;
      return { ...rest, key: un(it.key), containerId: un(it.containerId) };
    }
    async call(acc, fn) {
      try { return await fn(this.children[acc]); }
      catch (e) { if (e.auth) { this.reauth = acc; e.account = acc; } throw e; }
    }

    async load() {
      const both = this.accounts.length > 1;
      const groups = [], items = [], expired = [];
      const results = await Promise.allSettled(this.accounts.map((a) =>
        this.children[a].signedIn ? this.children[a].load() : Promise.reject(new AuthError())));
      results.forEach((r, i) => {
        const acc = this.accounts[i];
        if (r.status === "rejected") {
          if (!r.reason?.auth) throw r.reason;
          expired.push(acc); return;
        }
        for (const g of r.value.groups) {
          const event = g.source === "event";
          groups.push({
            ...g, id: acc + SEP + g.id, account: acc,
            ...(event ? { titleKey: both ? EVENT_GROUP_TITLE[acc] : g.titleKey } : {}),
            badge: both && !event ? ACCOUNT_BADGE[acc] : "",
          });
        }
        for (const it of r.value.items) items.push(this.wrap(acc, it));
      });
      // Task lists first, then calendars.
      groups.sort((a, b) => (a.source === "event") - (b.source === "event"));
      if (expired.length) this.reauth = expired[0];
      if (expired.length === this.accounts.length) { const e = new AuthError(); e.account = expired[0]; throw e; }
      return { groups, items, reauth: expired };
    }
    create(item) {
      const acc = this.accountOf(item.containerId);
      return this.call(acc, async (b) => this.wrap(acc, await b.create(this.strip(item))));
    }
    async update(prev, next) {
      const from = this.accountOf(prev), to = this.accountOf(next.containerId);
      if (from === to) return this.call(to, async (b) => this.wrap(to, await b.update(this.strip(prev), this.strip(next))));
      // Moving to the other account: create it there, then delete the original.
      const { id, key, ...fresh } = this.strip(next);
      const created = await this.call(to, (b) => b.create({
        ...fresh, isOccurrence: false,
        checklist: (fresh.checklist || []).map((c) => ({ ...c, id: newId() })),
        links: (fresh.links || []).map((l) => ({ ...l, id: newId() })),
      }));
      await this.call(from, (b) => b.remove(this.strip(prev)));
      return this.wrap(to, created);
    }
    remove(item) { return this.call(this.accountOf(item), (b) => b.remove(this.strip(item))); }
    listAttachments(item) { return this.call(this.accountOf(item), (b) => b.listAttachments ? b.listAttachments(this.strip(item)) : []); }
    addAttachment(item, file) { return this.call(this.accountOf(item), (b) => b.addAttachment(this.strip(item), file)); }
    removeAttachment(item, att) { return this.call(this.accountOf(item), (b) => b.removeAttachment(this.strip(item), att)); }
    downloadAttachment(item, att) { return this.call(this.accountOf(item), (b) => b.downloadAttachment(this.strip(item), att)); }
  }

  window.TBBackends = { GoogleBackend, MicrosoftBackend, DemoBackend, CombinedBackend };
})();
