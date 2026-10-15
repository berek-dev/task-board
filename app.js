/* Task Board app: views (table, board, calendar, overview), editing, filters, bulk edit, shortcuts.
 * All data comes from and goes to a backend in backends.js (Outlook/To Do, Google, or the demo). */
(() => {
  "use strict";
  const {
    t, STATUSES, ALL_PRIORITIES, PRIORITY_RANK, STATUS_RANK, REPEATS, REMIND_MINUTES, EVENTS_GROUP,
    ymd, today, addDays, toMin, fromMin, esc, newId, isNewId, sameJson, labelColor, normalize, safeStore, MAX_ATTACHMENT_BYTES,
  } = window.TB;
  const { GoogleBackend, MicrosoftBackend, DemoBackend, CombinedBackend } = window.TBBackends;
  const cfg = window.TM_CONFIG || {};
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => [...el.querySelectorAll(s)];

  const GROUP_COLORS = ["#5b5bd6", "#e2445c", "#fdab3d", "#00c875", "#a25ddc", "#ff7575", "#579bfc"];
  const STORAGE_PREFS = "tb.prefs";
  const START_H = 6, END_H = 23, HOUR_PX = 48; // calendar week grid: 06:00–23:00
  const EMPTY_FILTERS = { status: "", priority: "", due: "", label: "" };
  const PRESET_VIEWS = [
    { id: "p:week", nameKey: "view.week", filters: { due: "week" } },
    { id: "p:overdue", nameKey: "view.overdue", filters: { due: "overdue" } },
    { id: "p:stuck", nameKey: "view.stuck", filters: { status: "stuck" } },
    { id: "p:high", nameKey: "view.high", filters: { priority: "high" } },
  ];

  /* ---------- State ---------- */
  const prefs = Object.assign(
    { view: "table", kind: "all", hideDone: false, collapsed: {}, calMode: "week", groupBy: "list", sortBy: "date", savedViews: [] },
    safeStore.get(STORAGE_PREFS) || {});
  prefs.filters = { ...EMPTY_FILTERS, ...(prefs.filters || {}) };
  const state = { backend: null, groups: [], items: [], search: "", syncing: false, lastSync: null, selected: new Set(), calAnchor: today(), filtersOpen: false };
  const savePrefs = () => safeStore.set(STORAGE_PREFS, prefs);
  let PRIORITIES = ALL_PRIORITIES;
  // What the account supports; pass an item or group id to ask about the account that owns it.
  const caps = (x) => (x !== undefined && state.backend?.capsFor ? state.backend.capsFor(x) : state.backend?.caps) || {};
  const ACCOUNT_NAMES = { microsoft: "Microsoft", google: "Google" };
  const groupById = (id) => state.groups.find((g) => g.id === id);
  // The calendar that goes with an item or list: same account when both are connected.
  function eventGroupFor(x) {
    const acc = typeof x === "object" ? x?.account || groupById(x?.containerId)?.account : groupById(x)?.account;
    const evs = state.groups.filter((g) => g.source === "event");
    return (evs.find((g) => g.account && g.account === acc) || evs[0])?.id || EVENTS_GROUP;
  }
  const statusOf = (id) => STATUSES.find((s) => s.id === id) || STATUSES[0];
  const priorityOf = (id) => PRIORITIES.find((p) => p.id === (id || "")) || PRIORITIES[0];
  const findItem = (key) => state.items.find((i) => i.key === key);
  const allLabels = () => [...new Set(state.items.flatMap((i) => i.labels || []))].sort((a, b) => a.localeCompare(b));

  /* ---------- Filtering, sorting, formatting ---------- */
  const groupTitle = (g) => (g.titleKey ? t(g.titleKey) : g.title) + (g.badge ? ` · ${g.badge}` : "");
  function groupColor(groupId) {
    const idx = state.groups.findIndex((g) => g.id === groupId);
    const g = state.groups[idx];
    if (groupId === EVENTS_GROUP || g?.source === "event") return g?.account === "google" ? "#1e8e3e" : "#00a2c7";
    return GROUP_COLORS[(idx < 0 ? 0 : idx) % GROUP_COLORS.length];
  }
  const isOverdue = (i) => i.status !== "done" && i.date && i.date < today();
  function matchesDue(i, due) {
    const td = today();
    if (due === "overdue") return isOverdue(i);
    if (due === "today") return i.date === td;
    if (due === "week") return !!i.date && i.date >= td && i.date < addDays(td, 7);
    if (due === "nodate") return !i.date;
    return true;
  }
  function matches(i, { ignoreHideDone = false } = {}) {
    const f = prefs.filters, q = state.search.trim().toLowerCase();
    if (prefs.kind !== "all" && i.source !== prefs.kind) return false;
    if (!ignoreHideDone && prefs.hideDone && i.status === "done") return false;
    if (q && !(i.title.toLowerCase().includes(q) || i.notes.toLowerCase().includes(q) || (i.labels || []).some((l) => l.toLowerCase().includes(q)))) return false;
    if (f.status && i.status !== f.status) return false;
    if (f.priority && (i.priority || "none") !== f.priority) return false;
    if (f.due && !matchesDue(i, f.due)) return false;
    if (f.label && !(i.labels || []).includes(f.label)) return false;
    return true;
  }
  const visibleItems = (opts) => state.items.filter((i) => matches(i, opts));
  function sortItems(items) {
    const byDate = (a, b) => (a.date || "9999").localeCompare(b.date || "9999") || (a.startTime || "").localeCompare(b.startTime || "");
    const doneLast = (a, b) => (a.status === "done") - (b.status === "done");
    const cmp = {
      date: (a, b) => doneLast(a, b) || byDate(a, b),
      priority: (a, b) => doneLast(a, b) || PRIORITY_RANK[a.priority || ""] - PRIORITY_RANK[b.priority || ""] || byDate(a, b),
      status: (a, b) => STATUS_RANK[a.status] - STATUS_RANK[b.status] || byDate(a, b),
      title: (a, b) => 0,
    }[prefs.sortBy] || ((a, b) => 0);
    return items.slice().sort((a, b) => cmp(a, b) || a.title.localeCompare(b.title));
  }
  function fmtDay(date, opts = { month: "short", day: "numeric" }) {
    const d = new Date(date + "T00:00:00");
    const diff = Math.round((d - new Date(today() + "T00:00:00")) / 864e5);
    if (diff === 0) return t("date.today");
    if (diff === 1) return t("date.tomorrow");
    if (diff === -1) return t("date.yesterday");
    return d.toLocaleDateString(window.i18n.locale, { ...opts, ...(d.getFullYear() !== new Date().getFullYear() ? { year: "numeric" } : {}) });
  }
  const fmtDate = (i) => (i.date ? fmtDay(i.date) + (i.startTime ? " " + i.startTime : "") : "");
  const fmtSize = (n) => (n >= 1048576 ? (n / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(n / 1024)) + " KB");
  const labelChip = (l) => `<span class="lchip"><span class="dot" style="background:${labelColor(l)}"></span>${esc(l)}</span>`;
  function badges(i, { labels = true } = {}) {
    const out = [];
    if (i.checklist?.length) {
      const d = i.checklist.filter((c) => c.done).length;
      out.push(`<span class="badge${d === i.checklist.length ? " full" : ""}" title="${esc(t("edit.checklist"))}">☑ ${d}/${i.checklist.length}</span>`);
    }
    if (i.repeat || i.isOccurrence) out.push(`<span class="badge" title="${esc(t("badge.repeat"))}">🔁</span>`);
    if (i.remind) out.push(`<span class="badge" title="${esc(t("badge.remind"))}">🔔</span>`);
    if (i.links?.length) out.push(`<span class="badge" title="${esc(t("edit.links"))}">🔗 ${i.links.length}</span>`);
    if (i.hasAttachments) out.push(`<span class="badge" title="${esc(t("edit.files"))}">📎</span>`);
    if (labels) for (const l of i.labels || []) out.push(labelChip(l));
    return out.length ? `<span class="badges">${out.join("")}</span>` : "";
  }

  /* ---------- Rendering ---------- */
  function render() {
    $$(".seg [data-view]").forEach((b) => b.classList.toggle("active", b.dataset.view === prefs.view));
    $$(".seg [data-kind]").forEach((b) => b.classList.toggle("active", b.dataset.kind === prefs.kind));
    $("#hide-done").checked = prefs.hideDone;
    $("#sync-state").textContent = state.backend?.name === "demo" ? t("sync.demo")
      : state.lastSync ? t("sync.at", { time: state.lastSync.toLocaleTimeString(window.i18n.locale || [], { hour: "2-digit", minute: "2-digit" }) }) : "";
    $("#btn-refresh").classList.toggle("spin", state.syncing);
    renderConnect();
    renderFilterPanel();
    renderViewsBar();

    const prevScroll = $(".cal-scroll")?.scrollTop;
    const view = $("#view");
    view.innerHTML = { board: renderBoard, calendar: renderCalendar, overview: renderOverview }[prefs.view]?.() ?? renderTable();
    const sc = $(".cal-scroll");
    if (sc) sc.scrollTop = prevScroll ?? (8 - START_H) * HOUR_PX;
    renderBulkBar();
  }

  function fillSelect(sel, options, value) {
    if (document.activeElement === sel) return;
    sel.innerHTML = options.map(([v, label]) => `<option value="${esc(v)}">${esc(label)}</option>`).join("");
    sel.value = value;
    if (sel.value !== value) sel.value = options[0]?.[0] ?? "";
  }
  function renderFilterPanel() {
    const f = prefs.filters;
    $("#filter-panel").hidden = !state.filtersOpen;
    fillSelect($("#f-status"), [["", t("filter.any")], ...STATUSES.map((s) => [s.id, s.label])], f.status);
    fillSelect($("#f-priority"), [["", t("filter.any")], ...PRIORITIES.filter((p) => p.id).map((p) => [p.id, p.label]), ["none", t("priority.noneOption")]], f.priority);
    fillSelect($("#f-due"), [["", t("filter.any")], ...["overdue", "today", "week", "nodate"].map((d) => [d, t("due." + d)])], f.due);
    fillSelect($("#f-label"), [["", t("filter.any")], ...allLabels().map((l) => [l, l])], f.label);
    fillSelect($("#f-group"), ["list", "status", "priority", "due", "label"].map((g) => [g, t("group.by." + g)]), prefs.groupBy);
    fillSelect($("#f-sort"), ["date", "priority", "status", "title"].map((s) => [s, t("sort." + s)]), prefs.sortBy);
    $$(".table-only", $("#filter-panel")).forEach((el) => (el.hidden = prefs.view !== "table"));
    const n = Object.values(f).filter(Boolean).length;
    $("#filter-count").textContent = n ? `(${n})` : "";
    $("#btn-filters").classList.toggle("on", n > 0 || state.filtersOpen);
  }
  function normFilters(f) { return { ...EMPTY_FILTERS, ...f }; }
  function renderViewsBar() {
    const cur = normFilters(prefs.filters);
    const chip = (v, name, removable) => {
      const active = sameJson(normFilters(v.filters), cur) && (!v.kind || v.kind === prefs.kind);
      return `<span class="vchip${active ? " active" : ""}"><button data-apply-view="${esc(v.id)}">${esc(name)}</button>${removable ? `<button class="x" data-del-view="${esc(v.id)}" aria-label="${esc(t("view.delete"))}">×</button>` : ""}</span>`;
    };
    $("#views-bar").innerHTML =
      PRESET_VIEWS.map((v) => chip(v, t(v.nameKey))).join("") +
      prefs.savedViews.map((v) => chip(v, v.name, true)).join("") +
      `<button class="vchip add" data-save-view>${esc(t("view.save"))}</button>`;
  }

  /* ----- Table ----- */
  function tableGroups(items) {
    const mk = (id, title, color, list, extra = {}) => ({ id, title, color, items: list, ...extra });
    switch (prefs.groupBy) {
      case "status": return STATUSES.map((s) => mk("st:" + s.id, s.label, s.color, items.filter((i) => i.status === s.id)));
      case "priority": return PRIORITIES.map((p) => mk("pr:" + p.id, p.id ? p.label : t("priority.noneOption"), p.id ? p.color : "var(--muted)", items.filter((i) => (i.priority || "") === p.id)));
      case "due": return [
        mk("du:overdue", t("due.overdue"), "var(--danger)", items.filter(isOverdue)),
        mk("du:today", t("due.today"), "var(--accent)", items.filter((i) => !isOverdue(i) && i.date === today())),
        mk("du:week", t("due.week"), "#00a2c7", items.filter((i) => !isOverdue(i) && i.date > today() && i.date < addDays(today(), 7))),
        mk("du:later", t("due.later"), "#a25ddc", items.filter((i) => !isOverdue(i) && i.date >= addDays(today(), 7))),
        mk("du:none", t("due.nodate"), "var(--muted)", items.filter((i) => !i.date)),
        mk("du:done", t("status.done"), "var(--s-done)", items.filter((i) => i.status === "done" && i.date && i.date < today())),
      ];
      case "label": return [
        ...allLabels().map((l) => mk("lb:" + l, l, labelColor(l), items.filter((i) => (i.labels || []).includes(l)))),
        mk("lb:", t("filter.noLabel"), "var(--muted)", items.filter((i) => !(i.labels || []).length)),
      ];
      default:
        return state.groups.filter((g) => prefs.kind === "all" || g.source === prefs.kind)
          .map((g) => mk(g.id, groupTitle(g), groupColor(g.id), items.filter((i) => i.containerId === g.id), { source: g.source, quickAdd: true }));
    }
  }
  function renderTable() {
    const groups = tableGroups(visibleItems()).filter((g) => g.quickAdd || g.items.length);
    if (!groups.length) return `<div class="empty">${t("empty")}</div>`;
    return groups.map((g) => {
      const rows = sortItems(g.items);
      const allSel = rows.length && rows.every((i) => state.selected.has(i.key));
      const head = g.source === "event" ? t("col.event") : g.source === "task" ? t("col.task") : t("col.item");
      return `<section class="group${prefs.collapsed[g.id] ? " collapsed" : ""}" style="--group-color:${g.color}">
        <h3 class="group-head" data-collapse="${esc(g.id)}" style="color:${g.color}">
          <span class="caret">▾</span>${esc(g.title)} <span class="count">${rows.length === 1 ? t("group.items.one") : t("group.items.many", { n: rows.length })}</span>
        </h3>
        <div class="table">
          <div class="tr head"><label class="cell-sel"><input type="checkbox" data-select-group="${esc(g.id)}" ${allSel ? "checked" : ""} aria-label="${esc(t("bulk.selectAll"))}"></label><div>${head}</div><div>${t("col.status")}</div><div>${t("col.priority")}</div><div>${t("col.date")}</div></div>
          ${rows.map(rowHtml).join("")}
          ${g.quickAdd ? `<div class="add-row"><input data-quickadd="${esc(g.id)}" data-source="${g.source}" placeholder="${esc(t(g.source === "event" ? "add.event" : "add.task"))}" enterkeyhint="done" dir="auto"></div>` : ""}
        </div>
      </section>`;
    }).join("");
  }
  function rowHtml(i) {
    const s = statusOf(i.status), p = priorityOf(i.priority);
    const sel = state.selected.has(i.key);
    return `<div class="tr${i.status === "done" ? " done" : ""}${sel ? " selected" : ""}" data-key="${esc(i.key)}">
      <label class="cell-sel"><input type="checkbox" data-select="${esc(i.key)}" ${sel ? "checked" : ""} aria-label="${esc(t("bulk.select"))}"></label>
      <div class="cell-title" data-open="${esc(i.key)}"><span class="kind">${i.source === "event" ? "📅" : "☑"}</span><span class="t" dir="auto">${esc(i.title)}</span>${badges(i)}</div>
      <button class="pill pill-status" data-menu="status" data-key="${esc(i.key)}" style="background:${s.color}">${s.label}</button>
      <button class="pill pill-priority ${p.cls || ""}" data-menu="priority" data-key="${esc(i.key)}" style="background:${p.color}">${p.label}</button>
      <div class="cell-date${isOverdue(i) ? " overdue" : ""}">${esc(fmtDate(i))}</div>
    </div>`;
  }

  /* ----- Board ----- */
  function renderBoard() {
    const items = visibleItems();
    const cols = STATUSES.filter((s) => !(prefs.hideDone && s.id === "done"));
    return `<div class="board">${cols.map((s) => {
      const cards = sortItems(items.filter((i) => i.status === s.id));
      return `<div class="col" data-drop="${s.id}">
        <div class="col-head" style="background:${s.color}"><span>${s.label}</span><span>${cards.length}</span></div>
        ${cards.map(cardHtml).join("")}
      </div>`;
    }).join("")}</div>`;
  }
  function cardHtml(i) {
    const p = priorityOf(i.priority);
    const group = state.groups.find((g) => g.id === i.containerId);
    return `<div class="card" draggable="true" data-drag data-open="${esc(i.key)}" data-key="${esc(i.key)}" style="--group-color:${groupColor(i.containerId)}">
      <div class="t" dir="auto">${i.source === "event" ? "📅 " : ""}${esc(i.title)}</div>
      <div class="meta">
        ${group ? `<span>${esc(groupTitle(group))}</span>` : ""}
        ${i.date ? `<span class="${isOverdue(i) ? "overdue" : ""}">${esc(fmtDate(i))}</span>` : ""}
        ${i.priority ? `<span class="tag" style="background:${p.color}">${p.label}</span>` : ""}
        ${badges(i)}
      </div>
    </div>`;
  }

  /* ----- Calendar ----- */
  function weekStart(date) { const d = new Date(date + "T00:00:00"); return addDays(date, -d.getDay()); }
  function layoutDay(evs) {
    // Side-by-side lanes for overlapping events.
    const out = [];
    let cluster = [], lanes = [], clusterEnd = -1;
    const flush = () => { cluster.forEach((c) => (c.lanes = lanes.length)); cluster = []; lanes = []; };
    for (const e of evs) {
      const s = toMin(e.startTime);
      const en = Math.max(s + 15, e.endTime && e.endTime > e.startTime ? toMin(e.endTime) : s + 60);
      if (cluster.length && s >= clusterEnd) { flush(); clusterEnd = -1; }
      let lane = lanes.findIndex((end) => end <= s);
      if (lane < 0) { lane = lanes.length; lanes.push(en); } else lanes[lane] = en;
      const o = { e, s, en, lane };
      cluster.push(o); out.push(o);
      clusterEnd = Math.max(clusterEnd, en);
    }
    flush();
    return out;
  }
  const chipHtml = (i) => `<div class="cal-chip${i.status === "done" ? " done" : ""}" draggable="true" data-drag data-open="${esc(i.key)}" data-key="${esc(i.key)}" style="--c:${statusOf(i.status).color}" title="${esc(i.title)}">
      ${i.startTime ? `<span class="tm">${i.startTime}</span>` : ""}${i.source === "task" ? "☑ " : ""}<span dir="auto">${esc(i.title)}</span></div>`;

  function renderCalendar() {
    const mode = prefs.calMode === "month" ? "month" : "week";
    const items = sortItems(visibleItems().filter((i) => i.date));
    const loc = window.i18n.locale;
    const td = today();
    let title, body;

    if (mode === "week") {
      const ws = weekStart(state.calAnchor);
      const days = [...Array(7)].map((_, k) => addDays(ws, k));
      const s = new Date(days[0] + "T00:00:00"), e = new Date(days[6] + "T00:00:00");
      title = `${s.toLocaleDateString(loc, { month: "short", day: "numeric" })} – ${e.toLocaleDateString(loc, { month: "short", day: "numeric", year: "numeric" })}`;
      const hours = [...Array(END_H - START_H)].map((_, k) => `<div class="hr" style="top:${k * HOUR_PX}px">${String(START_H + k).padStart(2, "0")}:00</div>`).join("");
      const cols = days.map((d) => {
        const timed = items.filter((i) => i.date === d && i.source === "event" && i.startTime);
        const evs = layoutDay(timed).map(({ e, s, en, lane, lanes }) => {
          const top = (Math.min(Math.max(s, START_H * 60), END_H * 60 - 15) - START_H * 60) / 60 * HOUR_PX;
          const h = Math.max(20, (Math.min(en, END_H * 60) - Math.max(s, START_H * 60)) / 60 * HOUR_PX - 2);
          return `<div class="cal-ev${e.status === "done" ? " done" : ""}" draggable="true" data-drag data-open="${esc(e.key)}" data-key="${esc(e.key)}"
            style="top:${top}px;height:${h}px;inset-inline-start:${(lane / lanes) * 100}%;width:${100 / lanes}%;--c:${statusOf(e.status).color}" title="${esc(e.title)}">
            <span class="tm">${e.startTime}</span> <span dir="auto">${esc(e.title)}</span></div>`;
        }).join("");
        return `<div class="cw-col${d === td ? " today" : ""}" data-cal-day="${d}" data-cal-grid="1">${evs}</div>`;
      }).join("");
      const heads = days.map((d) => {
        const dt = new Date(d + "T00:00:00");
        return `<div class="cw-dayhead${d === td ? " today" : ""}"><span>${dt.toLocaleDateString(loc, { weekday: loc === "he-IL" ? "narrow" : "short" })}</span><b>${dt.getDate()}</b><button class="cm-add" data-cal-new="${d}" aria-label="${esc(t("top.new"))}">+</button></div>`;
      }).join("");
      const allday = days.map((d) => `<div class="cw-adcell" data-cal-day="${d}">${items.filter((i) => i.date === d && !(i.source === "event" && i.startTime)).map(chipHtml).join("")}</div>`).join("");
      body = `<div class="cal-week">
        <div class="cw-row cw-heads"><div class="gutter"></div>${heads}</div>
        <div class="cw-row cw-allday"><div class="gutter small">${t("cal.allDay")}</div>${allday}</div>
        <div class="cal-scroll"><div class="cw-row cw-grid" style="height:${(END_H - START_H) * HOUR_PX}px;--hour:${HOUR_PX}px"><div class="gutter hours">${hours}</div>${cols}</div></div>
      </div>`;
    } else {
      const a = new Date(state.calAnchor + "T00:00:00");
      const first = ymd(new Date(a.getFullYear(), a.getMonth(), 1));
      const gs = weekStart(first);
      title = a.toLocaleDateString(loc, { month: "long", year: "numeric" });
      const names = [...Array(7)].map((_, k) => `<div class="cm-name">${new Date(addDays(gs, k) + "T00:00:00").toLocaleDateString(loc, { weekday: loc === "he-IL" ? "narrow" : "short" })}</div>`).join("");
      const cells = [...Array(42)].map((_, k) => {
        const d = addDays(gs, k);
        const dayItems = items.filter((i) => i.date === d);
        const more = dayItems.length - 3;
        return `<div class="cm-cell${d.slice(0, 7) !== first.slice(0, 7) ? " dim" : ""}${d === td ? " today" : ""}" data-cal-day="${d}">
          <div class="cm-head"><b>${Number(d.slice(8))}</b><button class="cm-add" data-cal-new="${d}" aria-label="${esc(t("top.new"))}">+</button></div>
          ${dayItems.slice(0, 3).map(chipHtml).join("")}
          ${more > 0 ? `<button class="cm-more" data-cal-more="${d}">${esc(t("cal.more", { n: more }))}</button>` : ""}
        </div>`;
      }).join("");
      body = `<div class="cal-month"><div class="cm-names">${names}</div><div class="cm-grid">${cells}</div></div>`;
    }

    const toSchedule = sortItems(state.items.filter((i) => i.source === "task" && i.status !== "done" && matches(i))).slice(0, 40);
    const side = `<aside class="cal-side">
      <h3>${t("cal.unscheduled")}</h3>
      <p class="hint">${t("cal.dragHint")}</p>
      <ul>${toSchedule.map((i) => `<li class="side-task" draggable="true" data-drag data-open="${esc(i.key)}" data-key="${esc(i.key)}">
        <span class="dot" style="background:${statusOf(i.status).color}"></span><span class="t" dir="auto">${esc(i.title)}</span>
        <span class="when${isOverdue(i) ? " overdue" : ""}">${esc(fmtDate(i))}</span></li>`).join("") || `<li class="muted">${t("empty")}</li>`}</ul>
    </aside>`;

    return `<div class="calendar">
      <div class="cal-nav">
        <button class="btn ghost small arrow" data-cal-nav="-1" aria-label="${esc(t("cal.prev"))}">‹</button>
        <button class="btn ghost small" data-cal-nav="0">${t("cal.today")}</button>
        <button class="btn ghost small arrow" data-cal-nav="1" aria-label="${esc(t("cal.next"))}">›</button>
        <h3>${esc(title)}</h3>
        <div class="seg small"><button data-cal-mode="week" class="${mode === "week" ? "active" : ""}">${t("cal.week")}</button><button data-cal-mode="month" class="${mode === "month" ? "active" : ""}">${t("cal.month")}</button></div>
      </div>
      <div class="cal-body">${body}${side}</div>
    </div>`;
  }
  function calNavigate(dir) {
    if (dir === 0) state.calAnchor = today();
    else if (prefs.calMode === "month") {
      const a = new Date(state.calAnchor + "T00:00:00");
      state.calAnchor = ymd(new Date(a.getFullYear(), a.getMonth() + dir, 1));
    } else state.calAnchor = addDays(state.calAnchor, 7 * dir);
    render();
  }

  /* ----- Overview ----- */
  function renderOverview() {
    const all = visibleItems({ ignoreHideDone: true });
    const open = all.filter((i) => i.status !== "done");
    const td = today(), weekEnd = addDays(td, 7);
    const overdue = sortItems(open.filter(isOverdue));
    const thisWeek = open.filter((i) => i.date && i.date >= td && i.date < weekEnd);
    const tile = (n, key, warn) => `<div class="tile${warn && n ? " warn" : ""}"><div class="num">${n}</div><div class="lbl">${t(key)}</div></div>`;
    const statusBar = (items) => {
      const total = items.length || 1;
      return `<div class="sbar">${STATUSES.map((s) => {
        const n = items.filter((i) => i.status === s.id).length;
        return n ? `<span style="flex:${n / total};background:${s.color}" title="${esc(s.label)}: ${n}"></span>` : "";
      }).join("")}</div>`;
    };
    const legend = STATUSES.map((s) => {
      const n = all.filter((i) => i.status === s.id).length;
      const pct = all.length ? Math.round((n / all.length) * 100) : 0;
      return `<li><span class="dot" style="background:${s.color}"></span>${esc(s.label)}<b>${n}</b><span class="muted">${pct}%</span></li>`;
    }).join("");
    const byList = state.groups.filter((g) => prefs.kind === "all" || g.source === prefs.kind).map((g) => {
      const items = all.filter((i) => i.containerId === g.id);
      if (!items.length) return "";
      const done = items.filter((i) => i.status === "done").length;
      return `<div class="glist">
        <div class="glist-head"><span class="gname"><span class="dot" style="background:${groupColor(g.id)}"></span>${esc(groupTitle(g))}</span>
        <span class="muted">${t("ov.doneOf", { done, total: items.length })}</span></div>
        ${statusBar(items)}
      </div>`;
    }).join("");
    const line = (i, showDate) => `<li class="ov-item" data-open="${esc(i.key)}">
      <span class="dot" style="background:${statusOf(i.status).color}" title="${esc(statusOf(i.status).label)}"></span>
      <span class="t" dir="auto">${i.source === "event" ? "📅 " : ""}${esc(i.title)}</span>${badges(i, { labels: false })}
      <span class="when${isOverdue(i) ? " overdue" : ""}">${esc(showDate ? fmtDate(i) : i.startTime)}</span></li>`;
    let agenda = "";
    for (let d = 0; d < 7; d++) {
      const day = addDays(td, d);
      const items = sortItems(thisWeek.filter((i) => i.date === day));
      if (items.length) agenda += `<h4>${esc(fmtDay(day, { weekday: "long", day: "numeric", month: "short" }))}</h4><ul>${items.map((i) => line(i, false)).join("")}</ul>`;
    }
    return `<div class="overview">
      <div class="tiles">
        ${tile(open.length, "ov.open")}${tile(overdue.length, "ov.overdue", true)}
        ${tile(thisWeek.length, "ov.week")}${tile(open.filter((i) => i.status === "stuck").length, "ov.stuck", true)}
      </div>
      <div class="ov-cols">
        <div>
          <section class="panel"><h3>${t("ov.byStatus")}</h3>${statusBar(all)}<ul class="legend">${legend}</ul></section>
          <section class="panel"><h3>${t("ov.byList")}</h3>${byList || `<p class="muted">${t("empty")}</p>`}</section>
        </div>
        <div>
          <section class="panel"><h3>${t("ov.agenda")}</h3>${agenda || `<p class="muted">${t("ov.nothingScheduled")}</p>`}</section>
          <section class="panel"><h3>${t("ov.overdueList")}</h3>${overdue.length ? `<ul>${overdue.map((i) => line(i, true)).join("")}</ul>` : `<p class="muted">${t("ov.noOverdue")}</p>`}</section>
        </div>
      </div>
    </div>`;
  }

  /* ---------- Data operations (optimistic, rolled back on error) ---------- */
  async function sync({ quiet = false } = {}) {
    if (state.syncing || !state.backend?.signedIn) return;
    state.syncing = true; render();
    try {
      const { groups, items, reauth } = await state.backend.load();
      state.groups = groups; state.items = items.map(normalize); state.lastSync = new Date();
      PRIORITIES = state.backend.priorities || ALL_PRIORITIES;
      const keys = new Set(state.items.map((i) => i.key));
      for (const k of state.selected) if (!keys.has(k)) state.selected.delete(k);
      // One account's sign-in expired: show the other one's items and offer to reconnect.
      if (reauth?.length) showBanner(t("banner.reconnect", { name: ACCOUNT_NAMES[reauth[0]] }), reconnect);
      else hideBanner();
    } catch (e) { handleError(e, quiet); }
    state.syncing = false; render();
  }
  async function saveItem(prev, next) {
    const idx = prev ? state.items.findIndex((i) => i.key === prev.key) : -1;
    const temp = normalize({ ...next, key: next.key || "tmp:" + Date.now() + Math.random() });
    if (idx >= 0) state.items[idx] = temp; else state.items.push(temp);
    render();
    try {
      const saved = normalize(prev ? await state.backend.update(prev, next) : await state.backend.create(next));
      const j = state.items.findIndex((i) => i.key === temp.key);
      if (j >= 0) state.items[j] = saved;
      if (prev && state.selected.delete(prev.key)) state.selected.add(saved.key);
      render();
      // The server creates the next occurrence (To Do) or the series occurrences (calendar).
      const completedRepeat = next.repeat && next.status === "done" && prev?.status !== "done";
      const newSeries = next.source === "event" && next.repeat && (!prev || prev.repeat !== next.repeat);
      if (completedRepeat || newSeries) setTimeout(() => sync({ quiet: true }), 400);
      return saved;
    } catch (e) {
      const j = state.items.findIndex((i) => i.key === temp.key);
      if (prev) state.items[j] = prev; else state.items.splice(j, 1);
      render(); handleError(e);
      return null;
    }
  }
  async function deleteItem(item, { quiet = false } = {}) {
    const idx = state.items.findIndex((i) => i.key === item.key);
    if (idx < 0) return;
    state.items.splice(idx, 1); state.selected.delete(item.key); render();
    try { await state.backend.remove(item); if (!quiet) toast(t("toast.deleted")); }
    catch (e) { state.items.splice(idx, 0, item); render(); handleError(e); }
  }
  function handleError(e, quiet) {
    console.error(e);
    if (e.auth) { showBanner(e.account ? t("banner.reconnect", { name: ACCOUNT_NAMES[e.account] }) : t("banner.signedOut"), reconnect); return; }
    if (!quiet) toast(t("toast.saveFailed", { msg: e.message }));
  }
  async function reconnect() {
    try { await state.backend.signIn(); hideBanner(); await sync(); } catch (e) { toast(e.message); }
  }

  /* ---------- Status / priority popover ---------- */
  function openMenu(anchor, kind, item) {
    const menu = $("#menu");
    const opts = kind === "status" ? STATUSES : PRIORITIES;
    menu.innerHTML = opts.map((o) => `<button class="${o.cls || ""}" data-pick="${o.id}" style="background:${o.color}">${o.label}</button>`).join("");
    menu.hidden = false;
    const r = anchor.getBoundingClientRect();
    const mh = menu.offsetHeight, mw = menu.offsetWidth;
    menu.style.top = Math.max(8, r.bottom + mh + 8 > innerHeight ? r.top - mh - 4 : r.bottom + 4) + "px";
    menu.style.left = Math.min(innerWidth - mw - 8, Math.max(8, r.left + r.width / 2 - mw / 2)) + "px";
    menu.onclick = (ev) => {
      const b = ev.target.closest("[data-pick]");
      if (!b) return;
      closeMenu();
      if (item[kind] !== b.dataset.pick) saveItem(item, { ...item, [kind]: b.dataset.pick });
    };
  }
  function closeMenu() { $("#menu").hidden = true; }

  /* ---------- Editor ---------- */
  let editing = null;
  const ed = { checklist: [], labels: [], links: [], files: [] };
  const form = () => $("#edit-form");

  function openEditor(item, defaults = {}) {
    editing = item || null;
    const f = form();
    const taskGroups = state.groups.filter((g) => g.source === "task");
    fillSelect(f.elements.status, STATUSES.map((s) => [s.id, s.label]), "todo");
    fillSelect(f.elements.priority, PRIORITIES.map((p) => [p.id, p.id ? p.label : t("priority.noneOption")]), "");
    fillSelect(f.elements.repeat, REPEATS.map((r) => [r, t("repeat." + (r || "none"))]), "");
    fillSelect(f.elements.remindMin, REMIND_MINUTES.map((m) => [m, t("remind." + (m || "none"))]), "");
    const v = item ? structuredClone(item) : normalize({
      source: prefs.kind === "event" ? "event" : "task",
      title: "",
      ...defaults,
    });
    if (!v.containerId) v.containerId = v.source === "event" ? eventGroupFor(taskGroups[0]?.id) : taskGroups[0]?.id || "";
    for (const k of ["title", "source", "status", "priority", "date", "startTime", "endTime", "notes", "repeat"]) f.elements[k].value = v[k] || "";
    if (!taskGroups.length) v.source = "event";
    f.elements.source.value = v.source;
    fillContainers(v.source, v.containerId);
    f.elements.remindTime.value = v.source === "task" ? v.remind || "" : "";
    f.elements.remindMin.value = v.source === "event" ? v.remind || "" : "";
    ed.checklist = structuredClone(v.checklist || []);
    ed.labels = [...(v.labels || [])];
    ed.links = structuredClone(v.links || []);
    ed.files = [];
    $("#edit-title").textContent = t(item ? (item.source === "event" ? "edit.event" : "edit.task") : "edit.new");
    $("#btn-delete").hidden = !item;
    $("#btn-block").hidden = !(item && item.source === "task");
    f.elements.source.disabled = !taskGroups.length || !!item?.isOccurrence;
    $("#label-options").innerHTML = allLabels().map((l) => `<option value="${esc(l)}">`).join("");
    $("#label-input").value = ""; $("#checklist-add").value = ""; $("#link-url").value = ""; $("#link-title").value = "";
    renderLabelsEdit(); renderChecklistEdit(); renderLinksEdit(); loadFiles();
    syncEditorType();
    $("#editor").showModal();
    if (!item) f.elements.title.focus();
  }
  // Fills the List select with task lists, or with calendars for an event, keeping the account when possible.
  function fillContainers(source, selected) {
    const sel = form().elements.containerId;
    const groups = state.groups.filter((g) => g.source === source);
    fillSelect(sel, groups.map((g) => [g.id, groupTitle(g)]), "");
    const acc = groupById(selected)?.account;
    sel.value = groups.some((g) => g.id === selected) ? selected
      : (groups.find((g) => acc && g.account === acc) || groups[0])?.id || "";
  }
  function syncEditorType() {
    const f = form(), c = caps(f.elements.containerId.value);
    const isTask = f.elements.source.value === "task";
    f.classList.toggle("is-task", isTask);
    const calendars = state.groups.filter((g) => g.source === "event").length;
    $("#list-field").style.visibility = isTask || calendars > 1 ? "visible" : "hidden";
    $("#list-field > span").textContent = t(isTask ? "edit.list" : "edit.calendar");
    f.elements.date.required = !isTask;
    if (!isTask && !f.elements.date.value) f.elements.date.value = today();
    const show = (el, on) => (el.style.display = on ? "" : "none");
    show($("#repeat-field"), isTask ? c.taskRepeat : c.eventRepeat);
    show($("#remind-field"), isTask ? c.taskRemind : c.eventRemind);
    show($("#labels-field"), isTask ? c.taskLabels : c.eventLabels);
    show($("#checklist-field"), isTask && c.checklist);
    show($("#links-field"), isTask && c.taskLinks);
    show($("#files-field"), !!c.attachments);
    const series = !isTask && !!editing?.isOccurrence;
    f.elements.repeat.disabled = series;
    $("#series-hint").hidden = !series;
  }
  function renderLabelsEdit() {
    $("#label-chips").innerHTML = ed.labels.map((l, k) => `<span class="lchip"><span class="dot" style="background:${labelColor(l)}"></span>${esc(l)}<button type="button" class="x" data-label-del="${k}" aria-label="×">×</button></span>`).join("");
  }
  function addLabel(raw) {
    const l = String(raw || "").replace(/,/g, " ").trim();
    if (l && !ed.labels.includes(l)) { ed.labels.push(l); renderLabelsEdit(); }
  }
  function renderChecklistEdit() {
    const done = ed.checklist.filter((c) => c.done).length;
    $("#checklist-progress").textContent = ed.checklist.length ? `${done}/${ed.checklist.length}` : "";
    $("#checklist").innerHTML = ed.checklist.map((c, k) => `<li class="${c.done ? "done" : ""}">
      <input type="checkbox" data-ck-toggle="${k}" ${c.done ? "checked" : ""} aria-label="${esc(c.text)}">
      <input class="ck-text" data-ck-text="${k}" value="${esc(c.text)}" dir="auto">
      <button type="button" class="x" data-ck-del="${k}" aria-label="×">×</button></li>`).join("");
  }
  function renderLinksEdit() {
    $("#links").innerHTML = ed.links.map((l, k) => `<li><a href="${esc(l.url)}" target="_blank" rel="noopener noreferrer" dir="auto">🔗 ${esc(l.title || l.url)}</a>
      <button type="button" class="x" data-link-del="${k}" aria-label="×">×</button></li>`).join("");
  }
  function addLink() {
    let url = $("#link-url").value.trim();
    if (!url) return;
    if (!/^https?:\/\//i.test(url)) url = "https://" + url;
    try { if (!/^https?:$/.test(new URL(url).protocol)) throw 0; } catch { toast(t("edit.badLink")); return; }
    ed.links.push({ id: newId(), url, title: $("#link-title").value.trim() });
    $("#link-url").value = ""; $("#link-title").value = "";
    renderLinksEdit();
  }
  const savedItem = () => editing && !String(editing.key).startsWith("tmp:");
  async function loadFiles() {
    const list = $("#files"), hint = $("#files-hint");
    list.innerHTML = "";
    hint.hidden = true;
    $(".file-btn").hidden = !savedItem();
    if (!caps(editing || form().elements.containerId.value).attachments) return;
    if (!savedItem()) { hint.hidden = false; hint.textContent = t("edit.filesSaveFirst"); return; }
    hint.hidden = false; hint.textContent = t("edit.filesLoading");
    const forKey = editing.key;
    try {
      ed.files = await state.backend.listAttachments(editing);
      if (editing?.key !== forKey) return;
      hint.hidden = true;
      renderFiles();
    } catch (e) { hint.textContent = t("toast.saveFailed", { msg: e.message }); }
  }
  function renderFiles() {
    $("#files").innerHTML = ed.files.map((a) => `<li><button type="button" class="linkish" data-file-dl="${esc(a.id)}" dir="auto">📎 ${esc(a.name)}</button>
      <span class="muted">${fmtSize(a.size || 0)}</span><button type="button" class="x" data-file-del="${esc(a.id)}" aria-label="×">×</button></li>`).join("");
  }
  function markAttachments(item, has) {
    const it = findItem(item.key);
    if (it) { it.hasAttachments = has; editing = it; }
  }
  async function addFile(file) {
    if (!file || !savedItem()) return;
    if (file.size > MAX_ATTACHMENT_BYTES) { toast(t("edit.fileTooBig")); return; }
    const hint = $("#files-hint"); hint.hidden = false; hint.textContent = t("edit.fileUploading", { name: file.name });
    try {
      ed.files.push(await state.backend.addAttachment(editing, file));
      hint.hidden = true; renderFiles(); markAttachments(editing, true); render();
    } catch (e) { hint.hidden = true; handleError(e); }
  }
  async function removeFile(id) {
    const att = ed.files.find((a) => a.id === id);
    if (!att || !(await ask(t("edit.fileDeleteConfirm", { name: att.name }), { danger: true }))) return;
    try {
      await state.backend.removeAttachment(editing, att);
      ed.files = ed.files.filter((a) => a.id !== id);
      renderFiles(); markAttachments(editing, ed.files.length > 0); render();
    } catch (e) { handleError(e); }
  }
  async function downloadFile(id) {
    const att = ed.files.find((a) => a.id === id);
    if (!att) return;
    try {
      const blob = await state.backend.downloadAttachment(editing, att);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = att.name; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 10_000);
    } catch (e) { handleError(e); }
  }
  function submitEditor() {
    const f = form();
    const source = f.elements.source.value;
    addLabel($("#label-input").value);
    const pendingStep = $("#checklist-add").value.trim();
    if (pendingStep) ed.checklist.push({ id: newId(), text: pendingStep, done: false });
    if ($("#link-url").value.trim()) addLink();
    const next = {
      ...(editing || {}),
      source,
      containerId: f.elements.containerId.value || (source === "event" ? eventGroupFor(editing) : ""),
      title: f.elements.title.value.trim(),
      notes: f.elements.notes.value,
      status: f.elements.status.value,
      priority: f.elements.priority.value,
      date: f.elements.date.value,
      startTime: source === "event" ? f.elements.startTime.value : "",
      endTime: source === "event" ? f.elements.endTime.value : "",
      repeat: f.elements.repeat.disabled ? editing?.repeat || "" : f.elements.repeat.value,
      remind: source === "task" ? f.elements.remindTime.value : f.elements.remindMin.value,
      labels: ed.labels,
      checklist: source === "task" ? ed.checklist.filter((c) => c.text.trim()) : [],
      links: source === "task" ? ed.links : [],
    };
    if (!next.title) return;
    // A repeating item needs a start date.
    if (next.repeat && !next.date) next.date = today();
    if (!editing) delete next.key;
    $("#editor").close();
    saveItem(editing, next);
  }
  function blockTimeFor(task, date, startTime) {
    const start = startTime || "09:00";
    return {
      source: "event", containerId: eventGroupFor(task), title: task.title,
      notes: t("block.note", { title: task.title }), status: "todo", priority: task.priority,
      date, startTime: start, endTime: fromMin(Math.min(toMin(start) + 60, 23 * 60 + 59)),
      labels: caps(eventGroupFor(task)).eventLabels ? [...(task.labels || [])] : [],
    };
  }

  /* ---------- Drag and drop: board columns and calendar ---------- */
  let dragKey = null, dragOffsetY = 0;
  function wireDnD() {
    const view = $("#view");
    view.addEventListener("dragstart", (e) => {
      const el = e.target.closest("[data-drag]"); if (!el) return;
      dragKey = el.dataset.key;
      dragOffsetY = el.classList.contains("cal-ev") ? e.clientY - el.getBoundingClientRect().top : 0;
      el.classList.add("dragging");
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", dragKey);
    });
    view.addEventListener("dragend", (e) => {
      e.target.closest?.("[data-drag]")?.classList.remove("dragging");
      $$(".drop").forEach((c) => c.classList.remove("drop"));
      dragKey = null;
    });
    view.addEventListener("dragover", (e) => {
      const tgt = e.target.closest("[data-drop],[data-cal-day]"); if (!tgt || !dragKey) return;
      e.preventDefault();
      $$(".drop").forEach((c) => c !== tgt && c.classList.remove("drop"));
      tgt.classList.add("drop");
    });
    view.addEventListener("drop", (e) => {
      const tgt = e.target.closest("[data-drop],[data-cal-day]"); if (!tgt || !dragKey) return;
      e.preventDefault();
      const item = findItem(dragKey); dragKey = null;
      if (!item) return;
      if (tgt.dataset.drop) {
        if (item.status !== tgt.dataset.drop) saveItem(item, { ...item, status: tgt.dataset.drop });
        return;
      }
      const day = tgt.dataset.calDay;
      if (tgt.dataset.calGrid) {
        const r = tgt.getBoundingClientRect();
        let min = START_H * 60 + Math.round(((e.clientY - r.top - dragOffsetY) / HOUR_PX) * 4) * 15;
        min = Math.max(START_H * 60, Math.min(END_H * 60 - 15, min));
        if (item.source === "task") {
          // Time blocking: a calendar event for working on this task.
          saveItem(null, blockTimeFor(item, day, fromMin(min))).then((ok) => ok && toast(t("toast.blocked", { title: item.title })));
        } else {
          const dur = item.startTime ? Math.max(15, (item.endTime > item.startTime ? toMin(item.endTime) : toMin(item.startTime) + 60) - toMin(item.startTime)) : 60;
          saveItem(item, { ...item, date: day, startTime: fromMin(min), endTime: fromMin(Math.min(min + dur, 23 * 60 + 59)) });
        }
      } else if (item.date !== day) {
        saveItem(item, { ...item, date: day });
      }
    });
  }

  /* ---------- Bulk edit ---------- */
  function renderBulkBar() {
    const n = state.selected.size;
    const bar = $("#bulk-bar");
    bar.hidden = !n || prefs.view !== "table";
    document.body.classList.toggle("has-bulk", !bar.hidden);
    if (bar.hidden) return;
    $("#bulk-count").textContent = t("bulk.count", { n });
    fillSelect($("#bulk-status"), [["", t("bulk.setStatus")], ...STATUSES.map((s) => [s.id, s.label])], "");
    fillSelect($("#bulk-priority"), [["", t("bulk.setPriority")], ...PRIORITIES.map((p) => [p.id || "none", p.id ? p.label : t("priority.noneOption")])], "");
    fillSelect($("#bulk-list"), [["", t("bulk.moveTo")], ...state.groups.filter((g) => g.source === "task").map((g) => [g.id, groupTitle(g)])], "");
    $("#bulk-label").hidden = !(caps().taskLabels || caps().eventLabels);
  }
  async function applyBulk(fn) {
    const items = [...state.selected].map(findItem).filter(Boolean);
    let n = 0;
    for (const it of items) {
      const next = fn(it);
      if (next && (await saveItem(it, next))) n++;
    }
    state.selected.clear();
    render();
    toast(t("bulk.done", { n }));
  }

  /* ---------- Keyboard shortcuts ---------- */
  const SHORTCUTS = [
    ["N", "keys.new"], ["/", "keys.search"], ["1 – 4", "keys.views"], ["F", "keys.filters"], ["R", "keys.sync"],
    ["← →", "keys.calNav"], ["T", "keys.today"], ["Esc", "keys.escape"], ["Ctrl + Enter", "keys.save"], ["?", "keys.help"],
  ];
  function showShortcuts() {
    $("#shortcut-list").innerHTML = SHORTCUTS.map(([k, d]) => `<dt><kbd>${esc(k)}</kbd></dt><dd>${esc(t(d))}</dd>`).join("");
    $("#shortcuts").showModal();
  }
  function setView(v) { prefs.view = v; savePrefs(); render(); }
  function onKey(e) {
    if ((e.ctrlKey || e.metaKey) && e.key === "Enter" && $("#editor").open) { e.preventDefault(); form().requestSubmit(); return; }
    if (e.ctrlKey || e.metaKey || e.altKey || $("#app").hidden) return;
    if ($("#editor").open || $("#shortcuts").open) return;
    if (e.target.closest("input, textarea, select, [contenteditable]")) { if (e.key === "Escape") e.target.blur(); return; }
    const rtl = document.documentElement.dir === "rtl";
    // e.code, so shortcuts also work with a Hebrew keyboard layout.
    switch (e.code) {
      case "KeyN": openEditor(null); break;
      case "Slash": e.preventDefault(); if (e.shiftKey) showShortcuts(); else $("#search").focus(); break;
      case "Digit1": setView("table"); break;
      case "Digit2": setView("board"); break;
      case "Digit3": setView("calendar"); break;
      case "Digit4": setView("overview"); break;
      case "KeyF": state.filtersOpen = !state.filtersOpen; render(); break;
      case "KeyR": sync(); break;
      case "KeyT": if (prefs.view === "calendar") calNavigate(0); break;
      case "ArrowLeft": if (prefs.view === "calendar") calNavigate(rtl ? 1 : -1); break;
      case "ArrowRight": if (prefs.view === "calendar") calNavigate(rtl ? -1 : 1); break;
      case "Escape": closeMenu(); if (state.selected.size) { state.selected.clear(); render(); } break;
      default: return;
    }
  }

  /* ---------- UI helpers ---------- */
  let toastTimer;
  // In-page replacement for the browser's confirm and prompt boxes: resolves true / the typed text, or null when cancelled.
  function ask(msg, { input = false, danger = false } = {}) {
    const dlg = $("#ask"), field = $("#ask-input"), ok = $("#ask-ok");
    $("#ask-msg").textContent = msg;
    field.hidden = !input; field.value = "";
    ok.textContent = t(danger ? "edit.delete" : "ask.ok");
    ok.className = "btn " + (danger ? "danger" : "primary");
    return new Promise((resolve) => {
      dlg.returnValue = "";
      dlg.addEventListener("close", () => resolve(dlg.returnValue === "ok" ? (input ? field.value : true) : null), { once: true });
      dlg.showModal();
      (input ? field : ok).focus();
    });
  }
  function toast(msg) {
    const el = $("#toast"); el.textContent = msg; el.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(() => (el.hidden = true), 3500);
  }
  function showBanner(msg, onClick) { const b = $("#banner"); b.textContent = msg; b.hidden = false; b.onclick = onClick; }
  function hideBanner() { $("#banner").hidden = true; }
  function renderConnect() {
    const b = state.backend, name = b?.canConnect && ["google", "microsoft"].find((n) => b.canConnect(n));
    const btn = $("#btn-connect");
    btn.hidden = !name;
    if (name) btn.textContent = t(name === "google" ? "top.addGoogle" : "top.addMicrosoft");
  }
  function showApp() { $("#signin").hidden = true; $("#app").hidden = false; render(); }
  function showSignin(err) {
    $("#app").hidden = true; $("#signin").hidden = false;
    const e = $("#signin-error"); e.hidden = !err; e.textContent = err || "";
  }

  /* ---------- Wiring ---------- */
  function wire() {
    // Connects Microsoft or Google; with both connected, the board shows them together.
    const connect = async (name) => {
      const b = realBackend;
      if (!b?.children[name]) return state.backend?.signedIn ? toast(t("signin.notConfigured")) : showSignin(t("signin.notConfigured"));
      try { await b.signIn(name); useBackend(b); showApp(); await sync(); }
      catch (e) { if (b.signedIn) toast(e.message); else showSignin(e.message); }
    };
    $("#btn-signin-ms").onclick = () => connect("microsoft");
    $("#btn-signin-google").onclick = () => connect("google");
    $("#btn-connect").onclick = () => connect(state.backend.canConnect("google") ? "google" : "microsoft");
    $("#btn-demo").onclick = () => { useBackend(new DemoBackend()); showApp(); sync(); };
    $("#btn-signout").onclick = () => {
      state.backend.signOut(); state.items = []; state.groups = []; state.selected.clear();
      useBackend(realBackend); showSignin();
    };
    $("#btn-refresh").onclick = () => sync();
    $$(".btn-lang").forEach((b) => (b.onclick = () => { window.i18n.setLang(window.i18n.lang === "he" ? "en" : "he"); render(); renderProfiles(); }));
    $("#btn-new").onclick = () => openEditor(null);
    $("#btn-help").onclick = showShortcuts;
    $("#btn-shortcuts-close").onclick = () => $("#shortcuts").close();
    $("#search").oninput = (e) => { state.search = e.target.value; render(); };
    $("#hide-done").onchange = (e) => { prefs.hideDone = e.target.checked; savePrefs(); render(); };
    $$(".seg [data-view]").forEach((b) => (b.onclick = () => setView(b.dataset.view)));
    $$(".seg [data-kind]").forEach((b) => (b.onclick = () => { prefs.kind = b.dataset.kind; savePrefs(); render(); }));

    // Filters, grouping, sorting and saved views
    $("#btn-filters").onclick = () => { state.filtersOpen = !state.filtersOpen; render(); };
    const setFilter = (k) => (e) => { prefs.filters[k] = e.target.value; savePrefs(); e.target.blur(); render(); };
    $("#f-status").onchange = setFilter("status");
    $("#f-priority").onchange = setFilter("priority");
    $("#f-due").onchange = setFilter("due");
    $("#f-label").onchange = setFilter("label");
    $("#f-group").onchange = (e) => { prefs.groupBy = e.target.value; savePrefs(); e.target.blur(); render(); };
    $("#f-sort").onchange = (e) => { prefs.sortBy = e.target.value; savePrefs(); e.target.blur(); render(); };
    $("#btn-clear-filters").onclick = () => { prefs.filters = { ...EMPTY_FILTERS }; savePrefs(); render(); };
    $("#views-bar").addEventListener("click", (e) => {
      const apply = e.target.closest("[data-apply-view]");
      const del = e.target.closest("[data-del-view]");
      if (del) {
        prefs.savedViews = prefs.savedViews.filter((v) => v.id !== del.dataset.delView);
        savePrefs(); render(); return;
      }
      if (apply) {
        const v = [...PRESET_VIEWS, ...prefs.savedViews].find((x) => x.id === apply.dataset.applyView);
        if (!v) return;
        const active = apply.parentElement.classList.contains("active");
        prefs.filters = active ? { ...EMPTY_FILTERS } : normFilters(v.filters);
        if (!active && v.kind) prefs.kind = v.kind;
        if (!active && v.groupBy) prefs.groupBy = v.groupBy;
        if (!active && v.sortBy) prefs.sortBy = v.sortBy;
        savePrefs(); render(); return;
      }
      if (e.target.closest("[data-save-view]")) {
        ask(t("view.namePrompt"), { input: true }).then((name) => {
          if (!name?.trim()) return;
          prefs.savedViews.push({ id: newId("v:"), name: name.trim(), filters: normFilters(prefs.filters), kind: prefs.kind, groupBy: prefs.groupBy, sortBy: prefs.sortBy });
          savePrefs(); render();
        });
      }
    });

    // Main view interactions
    const view = $("#view");
    view.addEventListener("click", (e) => {
      const menuBtn = e.target.closest("[data-menu]");
      if (menuBtn) { e.stopPropagation(); return openMenu(menuBtn, menuBtn.dataset.menu, findItem(menuBtn.dataset.key)); }
      if (e.target.closest(".cell-sel")) return;
      const nav = e.target.closest("[data-cal-nav]");
      if (nav) return calNavigate(Number(nav.dataset.calNav));
      const mode = e.target.closest("[data-cal-mode]");
      if (mode) { prefs.calMode = mode.dataset.calMode; savePrefs(); return render(); }
      const add = e.target.closest("[data-cal-new]");
      if (add) return openEditor(null, { date: add.dataset.calNew });
      const more = e.target.closest("[data-cal-more]");
      if (more) { state.calAnchor = more.dataset.calMore; prefs.calMode = "week"; savePrefs(); return render(); }
      const open = e.target.closest("[data-open]");
      if (open) return openEditor(findItem(open.dataset.open));
      const col = e.target.closest("[data-collapse]");
      if (col) { prefs.collapsed[col.dataset.collapse] = !prefs.collapsed[col.dataset.collapse]; savePrefs(); render(); }
    });
    view.addEventListener("change", (e) => {
      const one = e.target.closest("[data-select]");
      if (one) { e.target.checked ? state.selected.add(one.dataset.select) : state.selected.delete(one.dataset.select); return render(); }
      const grp = e.target.closest("[data-select-group]");
      if (grp) {
        const g = tableGroups(visibleItems()).find((x) => x.id === grp.dataset.selectGroup);
        (g?.items || []).forEach((i) => (e.target.checked ? state.selected.add(i.key) : state.selected.delete(i.key)));
        render();
      }
    });
    view.addEventListener("keydown", (e) => {
      const input = e.target.closest("[data-quickadd]");
      if (!input || e.key !== "Enter" || !input.value.trim()) return;
      const source = input.dataset.source;
      saveItem(null, normalize({ source, containerId: input.dataset.quickadd, title: input.value.trim(), date: source === "event" ? today() : "" }));
      input.value = "";
      requestAnimationFrame(() => $(`[data-quickadd="${CSS.escape(input.dataset.quickadd)}"]`)?.focus());
    });
    wireDnD();

    // Bulk bar
    $("#bulk-status").onchange = (e) => { const v = e.target.value; if (v) applyBulk((i) => i.status !== v && { ...i, status: v }); };
    $("#bulk-priority").onchange = (e) => {
      const v = e.target.value === "none" ? "" : e.target.value;
      if (e.target.value) applyBulk((i) => (i.priority || "") !== v && { ...i, priority: v });
    };
    $("#bulk-date").onchange = (e) => { const v = e.target.value; if (v) applyBulk((i) => i.date !== v && { ...i, date: v }); e.target.value = ""; };
    $("#bulk-list").onchange = (e) => { const v = e.target.value; if (v) applyBulk((i) => i.source === "task" && i.containerId !== v && { ...i, containerId: v }); };
    $("#bulk-label").onkeydown = (e) => {
      if (e.key !== "Enter") return;
      const l = e.target.value.trim(); if (!l) return;
      e.target.value = "";
      applyBulk((i) => (i.source === "task" ? caps(i).taskLabels : caps(i).eventLabels) && !(i.labels || []).includes(l) && { ...i, labels: [...(i.labels || []), l] });
    };
    $("#bulk-delete").onclick = async () => {
      const items = [...state.selected].map(findItem).filter(Boolean);
      if (!items.length || !(await ask(t("bulk.confirmDelete", { n: items.length }), { danger: true }))) return;
      for (const it of items) await deleteItem(it, { quiet: true });
      toast(t("toast.deleted"));
    };
    $("#bulk-clear").onclick = () => { state.selected.clear(); render(); };

    // Editor
    const f = form();
    f.addEventListener("submit", (e) => { e.preventDefault(); submitEditor(); });
    f.elements.source.onchange = () => {
      // Carry the reminder across when switching type.
      if (f.elements.source.value === "event") f.elements.remindMin.value = f.elements.remindTime.value ? "15" : "";
      fillContainers(f.elements.source.value, f.elements.containerId.value);
      syncEditorType();
    };
    f.elements.containerId.onchange = () => { syncEditorType(); loadFiles(); };
    $("#btn-cancel").onclick = () => $("#editor").close();
    $("#btn-delete").onclick = async () => {
      const it = editing;
      if (it && await ask(t("confirm.delete", { title: it.title }), { danger: true })) { $("#editor").close(); deleteItem(it); }
    };
    $("#btn-block").onclick = () => {
      const task = editing;
      $("#editor").close();
      openEditor(null, blockTimeFor(task, task.date && task.date >= today() ? task.date : today()));
    };
    $("#label-input").addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === ",") { e.preventDefault(); addLabel(e.target.value); e.target.value = ""; }
    });
    $("#label-input").addEventListener("change", (e) => { if (allLabels().includes(e.target.value)) { addLabel(e.target.value); e.target.value = ""; } });
    $("#label-chips").addEventListener("click", (e) => {
      const b = e.target.closest("[data-label-del]");
      if (b) { ed.labels.splice(Number(b.dataset.labelDel), 1); renderLabelsEdit(); }
    });
    $("#checklist-add").addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      const text = e.target.value.trim();
      if (text) { ed.checklist.push({ id: newId(), text, done: false }); e.target.value = ""; renderChecklistEdit(); }
    });
    $("#checklist").addEventListener("change", (e) => {
      const tg = e.target.closest("[data-ck-toggle]");
      if (tg) { ed.checklist[Number(tg.dataset.ckToggle)].done = tg.checked; renderChecklistEdit(); }
    });
    $("#checklist").addEventListener("input", (e) => {
      const tx = e.target.closest("[data-ck-text]");
      if (tx) ed.checklist[Number(tx.dataset.ckText)].text = tx.value;
    });
    $("#checklist").addEventListener("keydown", (e) => { if (e.key === "Enter" && e.target.closest("[data-ck-text]")) { e.preventDefault(); $("#checklist-add").focus(); } });
    $("#checklist").addEventListener("click", (e) => {
      const d = e.target.closest("[data-ck-del]");
      if (d) { ed.checklist.splice(Number(d.dataset.ckDel), 1); renderChecklistEdit(); }
    });
    $("#btn-link-add").onclick = addLink;
    for (const id of ["#link-url", "#link-title"]) $(id).addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); addLink(); } });
    $("#links").addEventListener("click", (e) => {
      const d = e.target.closest("[data-link-del]");
      if (d) { ed.links.splice(Number(d.dataset.linkDel), 1); renderLinksEdit(); }
    });
    $("#file-input").onchange = (e) => { addFile(e.target.files[0]); e.target.value = ""; };
    $("#files").addEventListener("click", (e) => {
      const dl = e.target.closest("[data-file-dl]"); if (dl) return downloadFile(dl.dataset.fileDl);
      const del = e.target.closest("[data-file-del]"); if (del) removeFile(del.dataset.fileDel);
    });

    document.addEventListener("keydown", onKey);
    document.addEventListener("click", (e) => { if (!e.target.closest("#menu")) closeMenu(); });
    addEventListener("scroll", closeMenu, true);
    // Keep in sync with changes made in Outlook / To Do (or Google) on other devices.
    document.addEventListener("visibilitychange", () => { if (!document.hidden) sync({ quiet: true }); });
    setInterval(() => { if (!document.hidden && !$("#editor").open) sync({ quiet: true }); }, 2 * 60 * 1000);
  }

  let realBackend = null;
  function useBackend(b) { state.backend = b; PRIORITIES = b?.priorities || ALL_PRIORITIES; }
  function renderProfiles() {
    const { profiles, profile } = window.TB;
    $$(".btn-profile").forEach((b) => {
      b.hidden = !(profile && profiles.length > 1);
      if (profile) b.textContent = t("profile.switch", { name: profile.name || profile.id });
      b.onclick = () => window.TB.chooseProfile(null);
    });
    $("#profile-list").innerHTML = profiles.map((p) => `<button class="btn primary big" data-profile="${esc(p.id)}" dir="auto">${esc(p.name || p.id)}</button>`).join("");
    $$("[data-profile]").forEach((b) => (b.onclick = () => window.TB.chooseProfile(b.dataset.profile)));
  }
  async function start() {
    window.i18n.applyToDom();
    wire();
    renderProfiles();
    // Several people set up and this device hasn't said who it is yet.
    if (window.TB.profiles.length > 1 && !window.TB.profile && !cfg.DEMO_ONLY) { $("#profiles").hidden = false; return; }
    if (!cfg.DEMO_ONLY && "serviceWorker" in navigator && location.protocol !== "file:") navigator.serviceWorker.register("sw.js").catch(() => {});
    if (cfg.DEMO_ONLY || new URLSearchParams(location.search).has("demo")) {
      if (cfg.DEMO_ONLY) $("#btn-signout").hidden = true;
      useBackend(new DemoBackend()); showApp(); return sync();
    }
    const ms = new MicrosoftBackend(), google = new GoogleBackend();
    realBackend = new CombinedBackend({ ...(ms.configured ? { microsoft: ms } : {}), ...(google.configured ? { google } : {}) });
    // Offer the accounts this copy is set up for (both when neither is, so the setup hint shows).
    $("#btn-signin-ms").hidden = !ms.configured && google.configured;
    $("#btn-signin-google").hidden = !google.configured && ms.configured;
    if (!realBackend.configured) { showSignin(); return; }
    useBackend(realBackend);
    try { await realBackend.init(); } catch (e) { return showSignin(e.message); }
    if (realBackend.signedIn) { showApp(); sync(); } else showSignin();
  }
  start();
})();
