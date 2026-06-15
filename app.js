/* Day Reminders — Teams personal tab (v1.4.6)
   Thin client over the bot's REST API. Auth via Teams SSO.
   Server-side bot handles all notifications (proactive Adaptive Cards in chat).
*/
(function () {
  "use strict";

  const API_BASE = "https://func-day-reminders-17023.azurewebsites.net/api";
  const DONE_AGE_MS = 24 * 60 * 60 * 1000;
  const UNDO_MS = 5000;
  const APP_VERSION = "1.4.6";
  const VIEWS = ["lines", "grid", "calendar", "week"];
  const TAG_PALETTE = [
    "#0078d4", "#107c10", "#8764b8", "#ca5010", "#c50f1f",
    "#038387", "#d83b01", "#5c2d91", "#0099bc", "#498205",
  ];

  // ---------- state ----------
  let reminders = [];
  let userTemplates = []; // per-user templates loaded from /api/templates
  let settings = {
    eodTime: "17:00",
    leadMinutes: 10,
    weekdaysOnly: true,
    notifications: true,
    themeOverride: "auto", // tab-only, persisted in localStorage (bot doesn't care)
    groupBy: "none", // "none" | "tag" | "client" — tab-only, persisted in localStorage
    showAllDone: false,
  };
  let hasBot = false;
  let authToken = null;
  let activeTagFilter = null;
  let activeClientFilter = null;
  let quickFilter = "all"; // all | timed | anytime | high | done
  let searchText = "";
  let currentView = "lines"; // lines | grid | calendar | week
  let weekOffset = 0; // 0 = current week; not persisted across reloads
  let bulkMode = false;
  const bulkSelected = new Set();
  let teamsTheme = "default";
  let lastFocusedTrigger = null; // for dialog focus return
  const pendingDeletes = new Map(); // id -> { reminder, timer, toast }

  // localStorage keys (tab-only UI state — server is still source of truth for reminders)
  const LS_THEME = "themeOverride";
  const LS_QUICK_FILTER = "quickFilter";
  const LS_TAG_FILTER = "tagFilter";
  const LS_CLIENT_FILTER = "clientFilter";
  const LS_SEARCH = "searchText";
  const LS_VIEW = "currentView";
  const LS_GROUP_BY = "groupBy";
  const LS_ONBOARDED = "onboardingDismissed";
  const LS_GUIDE_OPEN = "guideOpen";

  try {
    const saved = localStorage.getItem(LS_THEME);
    if (saved) settings.themeOverride = saved;
    const sq = localStorage.getItem(LS_QUICK_FILTER);
    if (sq && ["all", "timed", "anytime", "high", "done"].includes(sq)) quickFilter = sq;
    const st = localStorage.getItem(LS_TAG_FILTER);
    if (st) activeTagFilter = st;
    const sc = localStorage.getItem(LS_CLIENT_FILTER);
    if (sc) activeClientFilter = sc;
    const ss = localStorage.getItem(LS_SEARCH);
    if (ss) searchText = ss;
    const sv = localStorage.getItem(LS_VIEW);
    if (sv && VIEWS.includes(sv)) currentView = sv;
    const sg = localStorage.getItem(LS_GROUP_BY);
    if (sg && ["none", "tag", "client"].includes(sg)) settings.groupBy = sg;
  } catch (_) {}

  // ---------- date helpers ----------
  // Today's date in Asia/Manila wall-clock as YYYY-MM-DD. Mirrors the bot's phToday so
  // that the tab and the scheduler agree on what "today" means regardless of where the
  // client browser sits.
  function todayPh() {
    const ph = new Date(Date.now() + 8 * 60 * 60 * 1000);
    return `${ph.getUTCFullYear()}-${String(ph.getUTCMonth() + 1).padStart(2, '0')}-${String(ph.getUTCDate()).padStart(2, '0')}`;
  }
  // Returns array of 7 YYYY-MM-DD strings starting Monday of the PH wall-clock
  // week containing today + (offset * 7 days). offset=0 = current week.
  function weekDates(offset) {
    const ph = new Date(Date.now() + 8 * 60 * 60 * 1000);
    const dow = ph.getUTCDay(); // 0=Sun ... 6=Sat
    const daysFromMonday = (dow + 6) % 7; // Mon=0, Sun=6
    const monday = new Date(ph.getTime());
    monday.setUTCDate(monday.getUTCDate() - daysFromMonday + offset * 7);
    const out = [];
    for (let i = 0; i < 7; i++) {
      const d = new Date(monday.getTime() + i * 24 * 60 * 60 * 1000);
      out.push(`${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`);
    }
    return out;
  }
  // Human label for a 7-day range like "Jun 15 – 21, 2026" or month/year spans.
  function weekRangeLabel(days) {
    const first = new Date(days[0] + "T00:00:00Z");
    const last = new Date(days[6] + "T00:00:00Z");
    const fmt = (d, opts) => d.toLocaleDateString(undefined, { ...opts, timeZone: "UTC" });
    if (first.getUTCMonth() === last.getUTCMonth() && first.getUTCFullYear() === last.getUTCFullYear()) {
      return `${fmt(first, { month: "short", day: "numeric" })} – ${fmt(last, { day: "numeric" })}, ${last.getUTCFullYear()}`;
    }
    if (first.getUTCFullYear() === last.getUTCFullYear()) {
      return `${fmt(first, { month: "short", day: "numeric" })} – ${fmt(last, { month: "short", day: "numeric" })}, ${last.getUTCFullYear()}`;
    }
    return `${fmt(first, { month: "short", day: "numeric", year: "numeric" })} – ${fmt(last, { month: "short", day: "numeric", year: "numeric" })}`;
  }

  // Short human label for a due date relative to today: "Today", "Tomorrow", "Yesterday",
  // weekday name within +/-6 days, else "MMM D" (and year if different).
  function formatDueLabel(dueAt, today) {
    if (!dueAt) return "";
    if (dueAt === today) return "Today";
    const due = new Date(dueAt + "T00:00:00Z");
    const tod = new Date(today + "T00:00:00Z");
    const diffDays = Math.round((due - tod) / (24 * 60 * 60 * 1000));
    if (diffDays === 1) return "Tomorrow";
    if (diffDays === -1) return "Yesterday";
    if (Math.abs(diffDays) <= 6) {
      const weekday = due.toLocaleDateString(undefined, { weekday: "short", timeZone: "UTC" });
      return weekday;
    }
    const opts = { month: "short", day: "numeric", timeZone: "UTC" };
    if (due.getUTCFullYear() !== tod.getUTCFullYear()) opts.year = "numeric";
    return due.toLocaleDateString(undefined, opts);
  }

  // debounce search input writes/render
  let searchTimer = null;
  function debouncedSearch(value) {
    searchText = value;
    if (searchTimer) clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      try { localStorage.setItem(LS_SEARCH, searchText); } catch (_) {}
      render();
    }, 150);
  }

  // templates: title + optional time (24h). Times use user's wall clock interpretation.
  const TEMPLATES = [
    { title: "Standup #work", time: "10:00" },
    { title: "Daily review #review", time: "17:00" },
    { title: "Lunch #personal", time: "12:30" },
    { title: "Hydration #wellness", time: "11:00" },
    { title: "PR review window #work", time: "14:00" },
    { title: "End-of-day wrap #review", time: "17:30" },
    { title: "1:1 prep #work", time: "16:00" },
    { title: "Inbox zero #work", time: "" },
    { title: "Weekly planning #planning", time: "" },
    { title: "Walk break #wellness", time: "15:00" },
  ];

  // ---------- DOM ----------
  const $ = (id) => document.getElementById(id);
  const addForm = $("addForm");
  const titleInput = $("title");
  const timeInput = $("time");
  const dueDateInput = $("dueDate");
  const clientInput = $("client");
  const clientList = $("clientList");
  const descriptionInput = $("description");
  const detailsToggleBtn = $("detailsToggle");
  const settingsDialog = $("settingsDialog");
  const settingsForm = $("settingsForm");
  const whatsNewDialog = $("whatsNewDialog");
  const permHint = $("permHint");
  const reminderRoot = $("reminderRoot");
  const groupToggle = $("groupToggle");
  const markAllDoneBtn = $("markAllDone");
  const filterBanner = $("filterBanner");
  const clearFilterBtn = $("clearFilter");
  const liveRegion = $("liveRegion");
  const toastRegion = $("toastRegion");
  const versionLabel = $("versionLabel");
  const rowOptionsDialog = $("rowOptionsDialog");
  const rowOptionsForm = $("rowOptionsForm");
  const rowOptionsTitle = $("rowOptionsTitle");
  const rowLeadMinutes = $("rowLeadMinutes");
  const rowDescription = $("rowDescription");
  const rowDueDate = $("rowDueDate");
  const rowClient = $("rowClient");
  const rowDueTomorrow = $("rowDueTomorrow");
  const searchInput = $("searchInput");
  const searchClear = $("searchClear");
  const bulkToggleBtn = $("bulkToggle");
  const bulkBar = $("bulkBar");
  const bulkCountLabel = $("bulkCount");
  const filterPills = $("filterPills");
  const filterCrumbs = $("filterCrumbs");
  const viewSwitch = $("viewSwitch");
  const onboardingCard = $("onboardingCard");
  const templatesDialog = $("templatesDialog");
  const templateGrid = $("templateGrid");
  let editingReminderId = null;

  if (versionLabel) versionLabel.textContent = `v${APP_VERSION}`;

  const todayDateString = new Date().toLocaleDateString(undefined, {
    weekday: "long", month: "short", day: "numeric",
  });

  // ---------- events ----------
  if (dueDateInput) dueDateInput.value = todayPh();
  if (detailsToggleBtn) {
    detailsToggleBtn.addEventListener("click", () => {
      const open = detailsToggleBtn.getAttribute("aria-expanded") === "true";
      const next = !open;
      detailsToggleBtn.setAttribute("aria-expanded", String(next));
      detailsToggleBtn.textContent = next ? "− Details" : "+ Details";
      descriptionInput.hidden = !next;
      if (next) {
        descriptionInput.focus();
        autoGrowTextarea(descriptionInput);
      }
    });
  }
  if (descriptionInput) {
    descriptionInput.addEventListener("input", () => autoGrowTextarea(descriptionInput));
  }

  addForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const { title, tags } = extractTagsFromTitle(titleInput.value.trim());
    if (!title) return;
    const time = timeInput.value || null;
    const today = todayPh();
    const dueAt = dueDateInput.value || today;
    const client = (clientInput && clientInput.value.trim().slice(0, 100)) || null;
    const description = descriptionInput.value.trim().slice(0, 2000) || null;
    const tempId = `tmp-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
    const optimistic = {
      id: tempId, title, time, tags, dueAt, description, rollDays: 0, client,
      done: false, priority: "normal", closedAt: null, _optimistic: true,
    };
    reminders.push(optimistic);
    titleInput.value = "";
    timeInput.value = "";
    dueDateInput.value = today;
    if (clientInput) clientInput.value = "";
    descriptionInput.value = "";
    descriptionInput.hidden = true;
    if (detailsToggleBtn) {
      detailsToggleBtn.setAttribute("aria-expanded", "false");
      detailsToggleBtn.textContent = "+ Details";
    }
    titleInput.focus();
    render();
    announce(`Added ${title}`);
    try {
      const body = { title, time, tags, dueAt };
      if (description) body.description = description;
      if (client) body.client = client;
      const created = await api("POST", "/reminders", body);
      replaceById(tempId, created.reminder);
      render();
    } catch (err) {
      reminders = reminders.filter((r) => r.id !== tempId);
      render();
      showError("Could not add reminder", err);
    }
  });

  $("openSettings").addEventListener("click", () => openDialog(settingsDialog, openSettings));
  $("settingsCancel").addEventListener("click", () => settingsDialog.close());
  settingsForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const next = {
      eodTime: $("setEodTime").value || settings.eodTime,
      leadMinutes: clampInt($("setLeadMinutes").value, 0, 240, settings.leadMinutes),
      weekdaysOnly: $("setWeekdaysOnly").checked,
      notifications: $("setNotifications").checked,
      quietStart: $("setQuietStart").value || null,
      quietEnd: $("setQuietEnd").value || null,
    };
    const newTheme = $("setThemeOverride").value;
    settings.themeOverride = newTheme;
    try { localStorage.setItem("themeOverride", newTheme); } catch (_) {}
    applyTheme(teamsTheme);
    try {
      const saved = await api("PUT", "/settings", { settings: next });
      Object.assign(settings, saved.settings);
      settingsDialog.close();
      announce("Settings saved");
    } catch (err) {
      showError("Could not save settings", err);
    }
  });

  // Cycle off → tag → client → off. Label + aria reflect current state so the
  // shortcut (g) is discoverable and the toggle reads cleanly.
  const GROUP_CYCLE = ["none", "tag", "client"];
  function applyGroupToggleLabel() {
    const labels = { none: "Group: off", tag: "Group: tag", client: "Group: client" };
    const titles = {
      none: "Not grouped — press to group by tag (shortcut: g)",
      tag: "Grouped by tag — press to group by client",
      client: "Grouped by client — press to ungroup",
    };
    groupToggle.textContent = labels[settings.groupBy] || labels.none;
    groupToggle.title = titles[settings.groupBy] || titles.none;
    groupToggle.setAttribute("aria-pressed", String(settings.groupBy !== "none"));
  }
  groupToggle.addEventListener("click", () => {
    const idx = GROUP_CYCLE.indexOf(settings.groupBy);
    settings.groupBy = GROUP_CYCLE[(idx + 1) % GROUP_CYCLE.length];
    try { localStorage.setItem(LS_GROUP_BY, settings.groupBy); } catch (_) {}
    applyGroupToggleLabel();
    const said = { none: "Ungrouped", tag: "Grouped by tag", client: "Grouped by client" };
    announce(said[settings.groupBy] || "Grouped");
    render();
  });
  applyGroupToggleLabel();

  clearFilterBtn.addEventListener("click", clearAllFilters);
  markAllDoneBtn.addEventListener("click", markAllOpenDone);

  const guideDetails = $("guideDetails");
  if (guideDetails) {
    try {
      const savedOpen = localStorage.getItem(LS_GUIDE_OPEN);
      if (savedOpen === "0") guideDetails.open = false;
      else if (savedOpen === "1") guideDetails.open = true;
    } catch (_) {}
    guideDetails.addEventListener("toggle", () => {
      try { localStorage.setItem(LS_GUIDE_OPEN, guideDetails.open ? "1" : "0"); } catch (_) {}
    });
  }
  function toggleGuide() {
    if (!guideDetails) return;
    guideDetails.open = !guideDetails.open;
    if (guideDetails.open) {
      guideDetails.scrollIntoView({ behavior: "smooth", block: "nearest" });
      const summary = guideDetails.querySelector("summary");
      if (summary) summary.focus();
    }
  }
  $("openGuide").addEventListener("click", toggleGuide);
  $("openWhatsNew").addEventListener("click", () => openDialog(whatsNewDialog));

  // search input
  if (searchInput) {
    searchInput.value = searchText;
    if (searchText) searchClear.hidden = false;
    searchInput.addEventListener("input", (e) => {
      const v = e.target.value;
      searchClear.hidden = !v;
      debouncedSearch(v);
    });
    searchInput.addEventListener("keydown", (e) => {
      if (e.key === "Escape") { e.preventDefault(); clearSearch(); searchInput.blur(); }
    });
  }
  if (searchClear) searchClear.addEventListener("click", () => { clearSearch(); searchInput.focus(); });

  // quick-filter pills
  if (filterPills) {
    filterPills.addEventListener("click", (e) => {
      const btn = e.target.closest("button.filter-pill");
      if (!btn) return;
      const key = btn.dataset.quick;
      if (!key) return;
      setQuickFilter(key);
    });
  }

  // view switch
  if (viewSwitch) {
    viewSwitch.addEventListener("click", (e) => {
      const btn = e.target.closest("button.view-btn");
      if (!btn) return;
      setView(btn.dataset.view);
    });
  }

  // bulk select
  if (bulkToggleBtn) bulkToggleBtn.addEventListener("click", () => setBulkMode(!bulkMode));
  if (bulkBar) {
    $("bulkDone").addEventListener("click", bulkMarkDone);
    $("bulkDelete").addEventListener("click", bulkDelete);
    $("bulkPriority").addEventListener("click", bulkTogglePriority);
    $("bulkCancel").addEventListener("click", () => setBulkMode(false));
  }

  // onboarding dismiss
  $("dismissOnboarding").addEventListener("click", () => {
    onboardingCard.hidden = true;
    try { localStorage.setItem(LS_ONBOARDED, "1"); } catch (_) {}
    lastFocusedTrigger = null;
  });
  try {
    if (localStorage.getItem(LS_ONBOARDED) !== "1") onboardingCard.hidden = false;
  } catch (_) { /* show by default if storage broken */ onboardingCard.hidden = false; }

  // templates
  $("openTemplates").addEventListener("click", () => {
    buildTemplateGrid();
    openDialog(templatesDialog);
  });

  $("rowOptionsCancel").addEventListener("click", () => rowOptionsDialog.close());
  $("rowOptionsDone").addEventListener("click", () => {
    const r = reminders.find((x) => x.id === editingReminderId);
    if (!r) { rowOptionsDialog.close(); return; }
    rowOptionsDialog.close();
    toggleDone(r);
  });
  $("rowOptionsSaveTemplate").addEventListener("click", () => {
    const r = reminders.find((x) => x.id === editingReminderId);
    if (!r) { rowOptionsDialog.close(); return; }
    rowOptionsDialog.close();
    saveAsTemplate(r);
  });
  if (rowDescription) {
    rowDescription.addEventListener("input", () => autoGrowTextarea(rowDescription));
  }
  if (rowDueTomorrow) {
    rowDueTomorrow.addEventListener("click", () => {
      const base = new Date(todayPh() + "T00:00:00Z");
      base.setUTCDate(base.getUTCDate() + 1);
      const yyyy = base.getUTCFullYear();
      const mm = String(base.getUTCMonth() + 1).padStart(2, "0");
      const dd = String(base.getUTCDate()).padStart(2, "0");
      rowDueDate.value = `${yyyy}-${mm}-${dd}`;
    });
  }
  rowOptionsForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const r = reminders.find((x) => x.id === editingReminderId);
    if (!r) { rowOptionsDialog.close(); return; }
    const raw = rowLeadMinutes.value.trim();
    let nextLead = null;
    if (raw !== "") {
      const n = clampInt(raw, 0, 240, NaN);
      if (isNaN(n)) { showError("Lead time must be 0 to 240", new Error("invalid")); return; }
      nextLead = n;
    }
    const nextDescRaw = rowDescription ? rowDescription.value.trim().slice(0, 2000) : "";
    const nextDesc = nextDescRaw || null;
    const nextDue = rowDueDate && rowDueDate.value ? rowDueDate.value : (r.dueAt || todayPh());
    const nextClientRaw = rowClient ? rowClient.value.trim().slice(0, 100) : "";
    const nextClient = nextClientRaw || null;
    const repeatSel = $("rowRepeat");
    const nextRepeat = repeatSel && ["daily", "weekdays", "weekly"].includes(repeatSel.value) ? repeatSel.value : "none";
    const prevLead = r.leadMinutes;
    const prevDesc = r.description;
    const prevDue = r.dueAt;
    const prevRoll = r.rollDays;
    const prevClient = r.client;
    const prevRepeat = r.repeat || "none";
    const dueInitial = prevDue || todayPh();
    const leadChanged = prevLead !== nextLead;
    const descChanged = (prevDesc || null) !== nextDesc;
    const dueChanged = nextDue !== dueInitial;
    const clientChanged = (prevClient || null) !== nextClient;
    const repeatChanged = prevRepeat !== nextRepeat;
    r.leadMinutes = nextLead;
    r.description = nextDesc;
    if (dueChanged) { r.dueAt = nextDue; r.rollDays = 0; }
    r.client = nextClient;
    r.repeat = nextRepeat;
    rowOptionsDialog.close();
    if (leadChanged || descChanged || dueChanged || clientChanged || repeatChanged) render();
    if (!leadChanged && !descChanged && !dueChanged && !clientChanged && !repeatChanged) return;
    try {
      const patch = {};
      if (leadChanged) patch.leadMinutes = nextLead;
      if (descChanged) patch.description = nextDesc || "";
      if (dueChanged) patch.dueAt = nextDue;
      if (clientChanged) patch.client = nextClient || "";
      if (repeatChanged) patch.repeat = nextRepeat;
      const updated = await api("PATCH", `/reminders/${r.id}`, patch);
      replaceLocal(updated.reminder);
      announce("Saved");
    } catch (err) {
      r.leadMinutes = prevLead;
      r.description = prevDesc;
      r.dueAt = prevDue;
      r.rollDays = prevRoll;
      r.client = prevClient;
      r.repeat = prevRepeat;
      render();
      showError("Could not save options", err);
    }
  });

  // keyboard shortcuts: ignore when typing in a form field
  document.addEventListener("keydown", (e) => {
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    const t = e.target;
    const tag = (t && t.tagName ? t.tagName.toLowerCase() : "");
    if (tag === "input" || tag === "textarea" || tag === "select" || (t && t.isContentEditable)) return;
    if (e.key === "/") { e.preventDefault(); titleInput.focus(); }
    else if (e.key === "f" || e.key === "F") { e.preventDefault(); searchInput && searchInput.focus(); }
    else if (e.key === "g" || e.key === "G") { e.preventDefault(); groupToggle.click(); }
    else if (e.key === "v" || e.key === "V") { e.preventDefault(); cycleView(); }
    else if (e.key === "?") { e.preventDefault(); toggleGuide(); }
    else if (e.key === "Escape" && (activeTagFilter || activeClientFilter || searchText || quickFilter !== "all" || bulkMode)) {
      if (bulkMode) setBulkMode(false);
      else clearAllFilters();
    }
  });

  // focus return after a dialog closes
  for (const d of [settingsDialog, whatsNewDialog, rowOptionsDialog, templatesDialog]) {
    d.addEventListener("close", () => {
      if (lastFocusedTrigger && typeof lastFocusedTrigger.focus === "function") {
        lastFocusedTrigger.focus();
      }
      lastFocusedTrigger = null;
    });
  }

  // ---------- render ----------
  function render() {
    const norm = (s) => String(s || "").toLowerCase();
    const searchNorm = norm(searchText.trim());

    // Cascade: !pendingDelete -> tag -> client -> search
    let visible = reminders.filter((r) => !pendingDeletes.has(r.id));
    if (activeTagFilter) {
      const tagNorm = norm(activeTagFilter);
      visible = visible.filter((r) => (r.tags || []).some((t) => norm(t) === tagNorm));
    }
    if (activeClientFilter) {
      const cNorm = norm(activeClientFilter);
      visible = visible.filter((r) => norm(r.client) === cNorm);
    }
    if (searchNorm) {
      visible = visible.filter((r) =>
        norm(r.title).includes(searchNorm) ||
        (r.tags || []).some((t) => norm(t).includes(searchNorm)) ||
        norm(r.client).includes(searchNorm) ||
        norm(r.description).includes(searchNorm)
      );
    }

    const open = visible.filter((r) => !r.done);
    const doneAll = visible.filter((r) => r.done);
    const recentDone = settings.showAllDone
      ? doneAll
      : doneAll.filter((r) => !r.closedAt || (Date.now() - new Date(r.closedAt).getTime()) < DONE_AGE_MS);
    const olderDoneCount = doneAll.length - recentDone.length;

    reminderRoot.innerHTML = "";
    populateClientList();
    updateFilterBanner();
    updateQuickFilterPills();
    updateViewSwitch();

    const totalNonPending = reminders.filter((r) => !pendingDeletes.has(r.id)).length;
    const noFiltersActive = !activeTagFilter && !activeClientFilter && !searchNorm && quickFilter === "all";

    // Calendar view bypasses the section/quick-filter structure entirely
    // for the "all" / "timed" / "anytime" cases — it shows ALL open items
    // (respecting tag + search filters from the cascade above) under hour rows.
    // For "done" and "high" we still defer to the per-quick-filter logic so
    // those filters keep meaning in calendar mode.
    if (currentView === "calendar" && quickFilter !== "done") {
      let calItems = open;
      if (quickFilter === "timed") calItems = open.filter((r) => !!r.time);
      else if (quickFilter === "anytime") calItems = open.filter((r) => !r.time);
      else if (quickFilter === "high") calItems = open.filter((r) => r.priority === "high");
      if (totalNonPending === 0 && noFiltersActive) {
        reminderRoot.appendChild(buildEmptyHero());
      } else if (calItems.length === 0) {
        reminderRoot.appendChild(buildEmptyState("Nothing on the calendar for this slice."));
      } else {
        reminderRoot.appendChild(buildCalendarLayout(calItems));
      }
      if (recentDone.length || olderDoneCount) {
        reminderRoot.appendChild(buildDoneSection(recentDone, olderDoneCount));
      }
      markAllDoneBtn.hidden = bulkMode || open.length === 0;
      updateBulkBar();
      updateBotHint();
      return;
    }

    // Week view: 7-column Mon–Sun grid of items by dueAt, with nav arrows.
    // Includes both open and done items so the week shows real activity; the
    // existing quick-filter for timed/anytime/high still narrows the slice.
    if (currentView === "week" && quickFilter !== "done") {
      const days = weekDates(weekOffset);
      let weekItems = visible.filter((r) => {
        const d = r.dueAt || r.createdDate;
        return d && days.includes(d);
      });
      if (quickFilter === "timed") weekItems = weekItems.filter((r) => !!r.time);
      else if (quickFilter === "anytime") weekItems = weekItems.filter((r) => !r.time);
      else if (quickFilter === "high") weekItems = weekItems.filter((r) => r.priority === "high");
      reminderRoot.appendChild(buildWeekLayout(weekItems, days));
      if (recentDone.length || olderDoneCount) {
        reminderRoot.appendChild(buildDoneSection(recentDone, olderDoneCount));
      }
      markAllDoneBtn.hidden = bulkMode || open.length === 0;
      updateBulkBar();
      updateBotHint();
      return;
    }

    if (quickFilter === "done") {
      if (doneAll.length === 0) {
        reminderRoot.appendChild(buildEmptyState(noFiltersActive
          ? "No completed reminders yet."
          : "No completed reminders match these filters."));
      } else {
        reminderRoot.appendChild(buildDoneSection(recentDone, olderDoneCount));
      }
    } else if (quickFilter === "timed") {
      const items = sortReminders(open.filter((r) => !!r.time));
      reminderRoot.appendChild(buildSection("Timed today", items, {
        showWhen: true,
        emptyText: items.length ? null : "Nothing scheduled in this slice.",
        meta: todayDateString,
        draggable: false,
      }));
    } else if (quickFilter === "anytime") {
      const items = sortByOrderThenTime(open.filter((r) => !r.time));
      reminderRoot.appendChild(buildSection("Anytime today", items, {
        showWhen: false,
        emptyText: items.length ? null : "Nothing without a time in this slice.",
        meta: "No specific time",
        draggable: true,
      }));
    } else if (quickFilter === "high") {
      const items = sortByOrderThenTime(open.filter((r) => r.priority === "high"));
      reminderRoot.appendChild(buildSection("High priority", items, {
        showWhen: true,
        emptyText: items.length ? null : "No high-priority reminders in this slice.",
        meta: "",
        draggable: true,
      }));
    } else {
      // quickFilter === "all"
      if (totalNonPending === 0 && noFiltersActive) {
        reminderRoot.appendChild(buildEmptyHero());
      } else if (open.length === 0 && !noFiltersActive) {
        reminderRoot.appendChild(buildEmptyState("No open reminders match these filters."));
      } else if (settings.groupBy === "tag") {
        renderByTag(open);
      } else if (settings.groupBy === "client") {
        renderByClient(open);
      } else {
        renderByTime(open);
      }
      if (recentDone.length || olderDoneCount) {
        reminderRoot.appendChild(buildDoneSection(recentDone, olderDoneCount));
      }
    }

    markAllDoneBtn.hidden = bulkMode || open.length === 0 || quickFilter === "done";
    updateBulkBar();
    updateBotHint();
  }

  function buildEmptyState(text) {
    const p = document.createElement("p");
    p.className = "empty";
    p.style.textAlign = "center";
    p.style.padding = "18px 8px";
    p.textContent = text;
    return p;
  }

  // ---------- grid view ----------
  function buildCard(r) {
    const card = document.createElement("article");
    card.className = "gcard";
    card.dataset.id = r.id;
    if (r.priority === "high") card.classList.add("high");
    if (r.firedAt) card.classList.add("fired");
    if (r.done) card.classList.add("done");
    if (pendingDeletes.has(r.id)) card.classList.add("pending-delete");
    if (bulkMode && bulkSelected.has(r.id)) card.classList.add("selected");
    if (r.snoozedUntil && new Date(r.snoozedUntil).getTime() > Date.now()) card.classList.add("snoozed");

    const head = document.createElement("div");
    head.className = "gcard-head";

    if (bulkMode) {
      const bcb = document.createElement("input");
      bcb.type = "checkbox";
      bcb.className = "bulk-checkbox";
      bcb.checked = bulkSelected.has(r.id);
      bcb.setAttribute("aria-label", `Select ${r.title}`);
      bcb.addEventListener("change", () => toggleBulkSelected(r));
      head.appendChild(bcb);
    } else {
      const star = document.createElement("button");
      star.type = "button";
      star.className = "icon-btn star" + (r.priority === "high" ? " on" : "");
      star.textContent = r.priority === "high" ? "★" : "☆";
      star.title = r.priority === "high" ? "Unstar" : "Mark high priority";
      star.setAttribute("aria-label", star.title);
      star.addEventListener("click", () => togglePriority(r));
      head.appendChild(star);
    }

    const title = document.createElement("h3");
    title.className = "gcard-title";
    title.textContent = displayTitle(r);
    title.title = "Click to rename";
    title.addEventListener("click", () => startTitleEdit(r, title));
    head.appendChild(title);

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "checkbox";
    cb.checked = !!r.done;
    cb.setAttribute("aria-label", r.done ? "Mark not done" : "Mark done");
    cb.addEventListener("change", () => toggleDone(r));
    head.appendChild(cb);

    card.appendChild(head);

    const cardMeta = buildTitleMeta(r);
    if (cardMeta) card.appendChild(cardMeta);

    if (r.snoozedUntil && new Date(r.snoozedUntil).getTime() > Date.now()) {
      const badge = document.createElement("button");
      badge.type = "button";
      badge.className = "snooze-badge";
      badge.textContent = `\u{1F4A4} ${formatSnoozeRelative(r.snoozedUntil)}`;
      badge.title = "Click to clear snooze";
      badge.addEventListener("click", () => unsnooze(r));
      card.appendChild(badge);
    }

    if (r.tags && r.tags.length) {
      const tagRow = document.createElement("div");
      tagRow.className = "tag-row";
      for (const t of r.tags) {
        const isActive = activeTagFilter && activeTagFilter.toLowerCase() === t.toLowerCase();
        const chip = document.createElement("button");
        chip.type = "button";
        chip.className = "tag-chip" + (isActive ? " active" : "");
        chip.style.backgroundColor = colorForTag(t);
        chip.textContent = `#${t}`;
        chip.title = isActive ? "Clear filter" : `Filter by #${t}`;
        chip.addEventListener("click", (e) => { e.stopPropagation(); setTagFilter(isActive ? null : t); });
        tagRow.appendChild(chip);
      }
      card.appendChild(tagRow);
    }

    const foot = document.createElement("div");
    foot.className = "gcard-foot";
    const when = document.createElement("span");
    when.className = "when";
    renderWhen(when, r.time, r.leadMinutes);
    when.title = r.time ? "Click to change time" : "Click to set a time";
    when.addEventListener("click", () => startTimeEdit(r, when));
    foot.appendChild(when);

    const more = document.createElement("button");
    more.type = "button";
    more.className = "icon-btn more";
    more.title = "More options";
    more.setAttribute("aria-label", "More options");
    more.textContent = "⋯";
    more.addEventListener("click", () => openRowOptions(r));
    foot.appendChild(more);

    card.appendChild(foot);
    const cardDesc = buildDescriptionBlock(r);
    if (cardDesc) card.appendChild(cardDesc);
    return card;
  }

  // ---------- calendar view (Day) ----------
  function buildCalendarLayout(items) {
    const section = document.createElement("section");
    section.className = "card calendar";

    const header = document.createElement("div");
    header.className = "calendar-header";
    header.appendChild(makeDayWeekSwitcher("calendar"));
    section.appendChild(header);

    const timed = items.filter((r) => !!r.time);
    const anytime = items.filter((r) => !r.time);

    // Hour range: 1h before earliest, 1h after latest, with a sensible 8-18 default.
    let startHour = 8, endHour = 18;
    if (timed.length) {
      const hours = timed.map((r) => parseInt(r.time.split(":")[0], 10));
      startHour = Math.max(0, Math.min(...hours) - 1);
      endHour = Math.min(23, Math.max(...hours) + 1);
    }
    const now = new Date();
    const nowHour = now.getHours();
    const nowMin = now.getMinutes();

    for (let h = startHour; h <= endHour; h++) {
      const row = document.createElement("div");
      row.className = "cal-row";
      if (h === nowHour) row.classList.add("now");

      const label = document.createElement("div");
      label.className = "cal-hour";
      label.textContent = `${String(h).padStart(2, "0")}:00`;

      const slot = document.createElement("div");
      slot.className = "cal-slot";

      const inHour = timed
        .filter((r) => parseInt(r.time.split(":")[0], 10) === h)
        .sort((a, b) => a.time.localeCompare(b.time));
      for (const r of inHour) slot.appendChild(buildCalendarItem(r));

      row.append(label, slot);
      section.appendChild(row);
    }

    if (anytime.length) {
      const at = document.createElement("div");
      at.className = "cal-anytime";
      const heading = document.createElement("h3");
      heading.textContent = "Anytime today";
      at.appendChild(heading);
      const stack = document.createElement("div");
      stack.style.display = "grid";
      stack.style.gap = "4px";
      for (const r of sortByOrderThenTime(anytime)) stack.appendChild(buildCalendarItem(r));
      at.appendChild(stack);
      section.appendChild(at);
    }

    if (nowHour >= startHour && nowHour <= endHour) {
      // tiny "now" label so the user can orient
      const note = document.createElement("p");
      note.className = "muted";
      note.style.margin = "8px 0 0";
      note.style.fontSize = "11px";
      note.textContent = `Now: ${String(nowHour).padStart(2, "0")}:${String(nowMin).padStart(2, "0")}`;
      section.appendChild(note);
    }

    return section;
  }
  // Mini segmented switcher for Day ⇄ Week. Rendered inside both views so the
  // user can flip between them without going to the top-bar view-switch.
  function makeDayWeekSwitcher(active) {
    const wrap = document.createElement("div");
    wrap.className = "day-week-switch";
    wrap.setAttribute("role", "group");
    wrap.setAttribute("aria-label", "Day or week");
    const mk = (key, label, title) => {
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = label;
      b.title = title;
      b.setAttribute("aria-pressed", String(active === key));
      b.addEventListener("click", () => { if (active !== key) setView(key); });
      return b;
    };
    wrap.appendChild(mk("calendar", "Day", "Hour-by-hour today"));
    wrap.appendChild(mk("week", "Week", "Mon–Sun grid"));
    return wrap;
  }

  // ---------- week view ----------
  function buildWeekLayout(items, days) {
    const section = document.createElement("section");
    section.className = "card week";

    const nav = document.createElement("div");
    nav.className = "week-nav";
    const prev = document.createElement("button");
    prev.type = "button";
    prev.className = "btn ghost small";
    prev.textContent = "‹ Prev";
    prev.title = "Previous week";
    prev.addEventListener("click", () => setWeek(weekOffset - 1));
    const range = document.createElement("div");
    range.className = "week-range";
    range.textContent = weekRangeLabel(days);
    const todayBtn = document.createElement("button");
    todayBtn.type = "button";
    todayBtn.className = "btn ghost small";
    todayBtn.textContent = "Today";
    todayBtn.title = "Jump to current week";
    todayBtn.disabled = weekOffset === 0;
    todayBtn.addEventListener("click", () => setWeek(0));
    const next = document.createElement("button");
    next.type = "button";
    next.className = "btn ghost small";
    next.textContent = "Next ›";
    next.title = "Next week";
    next.addEventListener("click", () => setWeek(weekOffset + 1));
    nav.append(makeDayWeekSwitcher("week"), prev, range, todayBtn, next);
    section.appendChild(nav);

    const grid = document.createElement("div");
    grid.className = "week-grid";
    const todayStr = todayPh();

    for (const dayStr of days) {
      const col = document.createElement("div");
      col.className = "week-day";
      if (dayStr === todayStr) col.classList.add("today");
      else if (dayStr < todayStr) col.classList.add("past");

      const dayDate = new Date(dayStr + "T00:00:00Z");
      const head = document.createElement("div");
      head.className = "week-day-head";
      const name = document.createElement("span");
      name.className = "week-day-name";
      name.textContent = dayDate.toLocaleDateString(undefined, { weekday: "short", timeZone: "UTC" });
      const num = document.createElement("span");
      num.className = "week-day-num";
      num.textContent = String(dayDate.getUTCDate());
      head.append(name, num);
      col.appendChild(head);

      const body = document.createElement("div");
      body.className = "week-day-body";
      const dayItems = items.filter((r) => (r.dueAt || r.createdDate) === dayStr);
      const timed = dayItems.filter((r) => !!r.time).sort((a, b) => a.time.localeCompare(b.time));
      const anytime = dayItems.filter((r) => !r.time);

      for (const r of timed) body.appendChild(buildWeekItem(r));
      if (anytime.length) {
        const sub = document.createElement("div");
        sub.className = "week-anytime-head";
        sub.textContent = "Anytime";
        body.appendChild(sub);
        for (const r of anytime) body.appendChild(buildWeekItem(r));
      }

      body.dataset.empty = dayItems.length === 0 ? "true" : "false";
      body.addEventListener("click", (e) => {
        if (e.target !== body) return;
        if (!dueDateInput || !titleInput) return;
        dueDateInput.value = dayStr;
        titleInput.focus();
        titleInput.scrollIntoView({ behavior: "smooth", block: "start" });
      });

      col.appendChild(body);
      grid.appendChild(col);
    }

    section.appendChild(grid);
    return section;
  }

  function buildWeekItem(r) {
    const item = document.createElement("div");
    item.className = "week-item";
    item.dataset.id = r.id;
    if (r.priority === "high") item.classList.add("high");
    if (r.done) item.classList.add("done");
    if (r.snoozedUntil && new Date(r.snoozedUntil).getTime() > Date.now()) item.classList.add("snoozed");

    const head = document.createElement("div");
    head.className = "week-item-head";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "checkbox week-item-check";
    cb.checked = !!r.done;
    cb.setAttribute("aria-label", r.done ? "Mark not done" : "Mark done");
    cb.addEventListener("click", (e) => e.stopPropagation());
    cb.addEventListener("change", (e) => { e.stopPropagation(); toggleDone(r); });

    const time = document.createElement("span");
    time.className = "week-item-time" + (r.time ? "" : " anytime");
    time.textContent = r.time ? formatTime(r.time) : "·";

    head.append(cb, time);

    const text = document.createElement("span");
    text.className = "week-item-text";
    text.textContent = displayTitle(r);

    item.append(head, text);
    item.addEventListener("click", (e) => {
      e.stopPropagation();
      if (e.target === cb || e.target.closest("button")) return;
      openRowOptions(r);
    });
    return item;
  }

  function setWeek(offset) {
    weekOffset = offset;
    render();
  }

  // displayTitle(): renders the title with a [Client] prefix when a client is set.
  // Keeps the underlying r.title editable as-is — the prefix is display-only and
  // updates automatically if the client changes.
  function displayTitle(r) {
    return r.client && r.client.trim()
      ? `[${r.client.trim()}] ${r.title}`
      : r.title;
  }

  function buildCalendarItem(r) {
    const item = document.createElement("div");
    item.className = "cal-item";
    item.dataset.id = r.id;
    if (r.priority === "high") item.classList.add("high");
    if (r.done) item.classList.add("done");
    if (r.snoozedUntil && new Date(r.snoozedUntil).getTime() > Date.now()) item.classList.add("snoozed");

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "checkbox";
    cb.checked = !!r.done;
    cb.setAttribute("aria-label", r.done ? "Mark not done" : "Mark done");
    cb.addEventListener("change", (e) => { e.stopPropagation(); toggleDone(r); });

    const time = document.createElement("span");
    time.className = "cal-item-time";
    time.textContent = r.time ? formatTime(r.time) : "";

    const text = document.createElement("span");
    text.className = "cal-item-text";
    text.textContent = r.rollDays > 0 ? `${displayTitle(r)}  (+${r.rollDays}d)` : displayTitle(r);

    const more = document.createElement("button");
    more.type = "button";
    more.className = "icon-btn more";
    more.textContent = "⋯";
    more.title = "More options";
    more.setAttribute("aria-label", "More options");
    more.addEventListener("click", (e) => { e.stopPropagation(); openRowOptions(r); });

    item.append(cb, time, text, more);
    item.addEventListener("click", (e) => {
      if (e.target === cb || e.target === more || e.target.closest("button")) return;
      startTitleEdit(r, text);
    });
    return item;
  }

  function buildEmptyHero() {
    const card = document.createElement("section");
    card.className = "card empty-hero";
    const h2 = document.createElement("h2");
    h2.textContent = "Nothing on the list yet";
    const p = document.createElement("p");
    p.textContent = "Try one of these to get started, or type your own above.";
    const examples = document.createElement("div");
    examples.className = "examples";
    const seeds = [
      "Send report 9:00 #work",
      "Call doctor #personal",
      "Standup 10:00 #team",
      "Review PRs #work",
    ];
    for (const s of seeds) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "example";
      b.textContent = s;
      b.addEventListener("click", () => {
        let text = s;
        const m = text.match(/\b(\d{1,2}:\d{2})\b/);
        if (m) {
          const [h, mm] = m[1].split(":");
          timeInput.value = `${h.padStart(2, "0")}:${mm}`;
          text = text.replace(m[0], "").replace(/\s+/g, " ").trim();
        }
        titleInput.value = text;
        titleInput.focus();
      });
      examples.appendChild(b);
    }
    card.append(h2, p, examples);
    return card;
  }

  function renderByTime(items) {
    const high = items.filter((r) => r.priority === "high");
    const timed = items.filter((r) => r.priority !== "high" && r.time)
      .sort((a, b) => a.time.localeCompare(b.time));
    const anytime = items.filter((r) => r.priority !== "high" && !r.time);

    if (high.length) {
      reminderRoot.appendChild(buildSection("High priority", sortByOrderThenTime(high), { showWhen: true, emptyText: null, draggable: true }));
    }
    reminderRoot.appendChild(buildSection("Today", timed, {
      showWhen: true,
      emptyText: timed.length ? null : (activeTagFilter ? `No timed reminders for ${activeTagFilter}.` : "Nothing scheduled. Add something above."),
      meta: todayDateString,
      draggable: false,
    }));
    reminderRoot.appendChild(buildSection("Anytime today", sortByOrderThenTime(anytime), {
      showWhen: false,
      emptyText: anytime.length ? null : "None.",
      meta: "No specific time",
      draggable: true,
    }));
  }

  function renderByTag(items) {
    const buckets = new Map();
    const untagged = [];
    for (const r of items) {
      if (!r.tags || r.tags.length === 0) {
        untagged.push(r);
      } else {
        for (const t of r.tags) {
          if (!buckets.has(t)) buckets.set(t, []);
          buckets.get(t).push(r);
        }
      }
    }
    const tagNames = [...buckets.keys()].sort();
    for (const tag of tagNames) {
      const section = buildSection(`#${tag}`, sortReminders(buckets.get(tag)), { showWhen: true, emptyText: null, draggable: true });
      const head = section.querySelector("h2");
      if (head) head.style.color = colorForTag(tag);
      reminderRoot.appendChild(section);
    }
    if (untagged.length) {
      reminderRoot.appendChild(buildSection("No tag", sortReminders(untagged), { showWhen: true, emptyText: null, draggable: true }));
    }
    if (tagNames.length === 0 && untagged.length === 0) {
      reminderRoot.appendChild(buildSection("Today", [], { showWhen: true, emptyText: "Nothing scheduled.", meta: todayDateString, draggable: false }));
    }
  }

  function renderByClient(items) {
    const buckets = new Map();
    const unassigned = [];
    for (const r of items) {
      const c = r.client && r.client.trim();
      if (!c) { unassigned.push(r); continue; }
      if (!buckets.has(c)) buckets.set(c, []);
      buckets.get(c).push(r);
    }
    const clientNames = [...buckets.keys()].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
    for (const c of clientNames) {
      const section = buildSection(`[${c}]`, sortReminders(buckets.get(c)), { showWhen: true, emptyText: null, draggable: true });
      const head = section.querySelector("h2");
      if (head) head.style.color = colorForClient(c);
      reminderRoot.appendChild(section);
    }
    if (unassigned.length) {
      reminderRoot.appendChild(buildSection("No client", sortReminders(unassigned), { showWhen: true, emptyText: null, draggable: true }));
    }
    if (clientNames.length === 0 && unassigned.length === 0) {
      reminderRoot.appendChild(buildSection("Today", [], { showWhen: true, emptyText: "Nothing scheduled.", meta: todayDateString, draggable: false }));
    }
  }

  function compareOrderless(a, b) {
    const ta = a.time || "zz";
    const tb = b.time || "zz";
    if (ta !== tb) return ta.localeCompare(tb);
    return (a.id || "").localeCompare(b.id || "");
  }

  function sortReminders(items) {
    return [...items].sort((a, b) => {
      if (a.priority !== b.priority) return a.priority === "high" ? -1 : 1;
      const aHas = typeof a.order === "number";
      const bHas = typeof b.order === "number";
      if (aHas !== bHas) return aHas ? -1 : 1;
      if (aHas && bHas) return a.order - b.order;
      return compareOrderless(a, b);
    });
  }

  function sortByOrderThenTime(items) {
    return [...items].sort((a, b) => {
      const aHas = typeof a.order === "number";
      const bHas = typeof b.order === "number";
      if (aHas !== bHas) return aHas ? -1 : 1;
      if (aHas && bHas) return a.order - b.order;
      return compareOrderless(a, b);
    });
  }

  function buildSection(title, items, opts) {
    const section = document.createElement("section");
    section.className = "card";

    const head = document.createElement("div");
    head.className = "section-head";
    const h2 = document.createElement("h2");
    h2.textContent = title;
    const meta = document.createElement("span");
    meta.className = "muted";
    meta.textContent = opts.meta || "";
    head.append(h2, meta);
    section.appendChild(head);

    if (items.length === 0 && opts.emptyText) {
      const p = document.createElement("p");
      p.className = "empty";
      p.textContent = opts.emptyText;
      section.appendChild(p);
      return section;
    }

    if (currentView === "grid") {
      const grid = document.createElement("div");
      grid.className = "gcard-grid";
      for (const r of items) grid.appendChild(buildCard(r));
      section.appendChild(grid);
    } else {
      const ul = document.createElement("ul");
      ul.className = "list";
      for (const r of items) ul.appendChild(buildRow(r, opts.showWhen, !!opts.draggable));
      section.appendChild(ul);
    }
    return section;
  }

  function buildDoneSection(recent, olderCount) {
    const section = document.createElement("section");
    section.className = "card";

    const head = document.createElement("div");
    head.className = "section-head";
    const h2 = document.createElement("h2");
    h2.textContent = "Done";
    const actions = document.createElement("div");
    actions.style.display = "flex";
    actions.style.gap = "8px";
    if (olderCount > 0) {
      const showAll = document.createElement("button");
      showAll.className = "btn ghost small";
      showAll.textContent = settings.showAllDone ? "Hide older" : `Show ${olderCount} older`;
      showAll.addEventListener("click", () => {
        settings.showAllDone = !settings.showAllDone;
        render();
      });
      actions.appendChild(showAll);
    }
    const clear = document.createElement("button");
    clear.className = "btn ghost small";
    clear.textContent = "Clear";
    clear.addEventListener("click", clearDoneVisible);
    actions.appendChild(clear);
    head.append(h2, actions);
    section.appendChild(head);

    const ul = document.createElement("ul");
    ul.className = "list done";
    for (const r of recent) ul.appendChild(buildRow(r, true, false));
    section.appendChild(ul);
    return section;
  }

  function buildRow(r, _showWhen, draggable) {
    const li = document.createElement("li");
    li.dataset.id = r.id;
    if (r.firedAt) li.classList.add("fired");
    if (r.priority === "high") li.classList.add("high");
    if (pendingDeletes.has(r.id)) li.classList.add("pending-delete");
    if (bulkMode && bulkSelected.has(r.id)) li.classList.add("selected");
    if (r.snoozedUntil && new Date(r.snoozedUntil).getTime() > Date.now()) li.classList.add("snoozed");

    // In bulk mode the drag handle is replaced by a select checkbox.
    let leading;
    if (bulkMode) {
      leading = document.createElement("input");
      leading.type = "checkbox";
      leading.className = "bulk-checkbox";
      leading.checked = bulkSelected.has(r.id);
      leading.setAttribute("aria-label", `Select ${r.title}`);
      leading.addEventListener("change", () => toggleBulkSelected(r));
      draggable = false; // disable drag while selecting
    } else {
      leading = document.createElement("button");
      leading.type = "button";
      leading.className = "drag-handle";
      leading.setAttribute("aria-label", "Drag to reorder");
      leading.title = draggable ? "Drag to reorder" : "Reorder not available in this view";
      leading.textContent = "☰"; // ≡
      if (!draggable) leading.style.visibility = "hidden";
    }
    const handle = leading;

    const star = document.createElement("button");
    star.className = "icon-btn star" + (r.priority === "high" ? " on" : "");
    star.title = r.priority === "high" ? "Unstar" : "Mark high priority";
    star.setAttribute("aria-label", star.title);
    star.textContent = r.priority === "high" ? "★" : "☆";
    star.addEventListener("click", () => togglePriority(r));

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.className = "checkbox";
    cb.checked = !!r.done;
    cb.setAttribute("aria-label", r.done ? "Mark not done" : "Mark done");
    cb.addEventListener("change", () => toggleDone(r));

    const titleWrap = document.createElement("div");
    titleWrap.className = "title-wrap";
    const title = document.createElement("span");
    title.className = "title";
    title.textContent = displayTitle(r);
    title.title = "Click to rename";
    title.addEventListener("click", () => startTitleEdit(r, title));
    titleWrap.appendChild(title);

    const titleMeta = buildTitleMeta(r);
    if (titleMeta) titleWrap.appendChild(titleMeta);

    if (r.snoozedUntil && new Date(r.snoozedUntil).getTime() > Date.now()) {
      const badge = document.createElement("button");
      badge.type = "button";
      badge.className = "snooze-badge";
      badge.textContent = `\u{1F4A4} ${formatSnoozeRelative(r.snoozedUntil)}`;
      badge.title = "Click to clear snooze (reminder will fire at its normal time again)";
      badge.setAttribute("aria-label", `Snoozed. Click to unsnooze.`);
      badge.addEventListener("click", () => unsnooze(r));
      titleWrap.appendChild(badge);
    }

    if (r.tags && r.tags.length) {
      const tagRow = document.createElement("div");
      tagRow.className = "tag-row";
      for (const t of r.tags) {
        const chip = document.createElement("button");
        chip.type = "button";
        const isActive = activeTagFilter && activeTagFilter.toLowerCase() === t.toLowerCase();
        chip.className = "tag-chip" + (isActive ? " active" : "");
        chip.style.backgroundColor = colorForTag(t);
        chip.textContent = `#${t}`;
        chip.title = isActive ? "Clear filter" : `Filter by #${t}`;
        chip.setAttribute("aria-label", chip.title);
        chip.addEventListener("click", (e) => {
          e.stopPropagation();
          setTagFilter(isActive ? null : t);
        });
        tagRow.appendChild(chip);
      }
      titleWrap.appendChild(tagRow);
    }

    const when = document.createElement("span");
    when.className = "when";
    renderWhen(when, r.time, r.leadMinutes);
    when.title = r.time ? "Click to change time" : "Click to set a time";
    when.tabIndex = 0;
    when.addEventListener("click", () => startTimeEdit(r, when));

    const more = document.createElement("button");
    more.type = "button";
    more.className = "icon-btn more";
    more.title = "More options";
    more.setAttribute("aria-label", "More options");
    more.textContent = "⋯";
    more.addEventListener("click", () => openRowOptions(r));

    const del = document.createElement("button");
    del.className = "icon-btn";
    del.title = "Delete";
    del.setAttribute("aria-label", "Delete reminder");
    del.textContent = "✕";
    del.addEventListener("click", () => removeReminder(r));

    li.append(handle, star, cb, titleWrap, when, more, del);
    const desc = buildDescriptionBlock(r);
    if (desc) li.appendChild(desc);
    if (draggable) attachDragHandlers(li, r);
    return li;
  }

  function attachDragHandlers(li, r) {
    li.draggable = true;
    li.addEventListener("dragstart", (e) => {
      try { e.dataTransfer.effectAllowed = "move"; } catch (_) {}
      e.dataTransfer.setData("text/reminder-id", r.id);
      li.classList.add("dragging");
    });
    li.addEventListener("dragend", () => {
      li.classList.remove("dragging");
      document.querySelectorAll(".drop-before, .drop-after").forEach((el) =>
        el.classList.remove("drop-before", "drop-after")
      );
    });
    li.addEventListener("dragover", (e) => {
      if (!e.dataTransfer.types.includes("text/reminder-id")) return;
      e.preventDefault();
      const rect = li.getBoundingClientRect();
      const before = (e.clientY - rect.top) < rect.height / 2;
      li.classList.toggle("drop-before", before);
      li.classList.toggle("drop-after", !before);
    });
    li.addEventListener("dragleave", () => {
      li.classList.remove("drop-before", "drop-after");
    });
    li.addEventListener("drop", (e) => {
      e.preventDefault();
      const sourceId = e.dataTransfer.getData("text/reminder-id");
      const rect = li.getBoundingClientRect();
      const before = (e.clientY - rect.top) < rect.height / 2;
      li.classList.remove("drop-before", "drop-after");
      if (!sourceId || sourceId === r.id) return;
      reorderInSection(sourceId, r.id, before, li.parentElement);
    });
  }

  async function reorderInSection(sourceId, targetId, before, ul) {
    const source = reminders.find((x) => x.id === sourceId);
    if (!source || !ul) return;
    const siblings = [...ul.querySelectorAll("li[data-id]")];
    const sectionItems = siblings.map((s) => reminders.find((x) => x.id === s.dataset.id)).filter(Boolean);
    const withoutSource = sectionItems.filter((x) => x.id !== sourceId);
    const targetIdx = withoutSource.findIndex((x) => x.id === targetId);
    if (targetIdx < 0) return;
    const insertAt = before ? targetIdx : targetIdx + 1;
    withoutSource.splice(insertAt, 0, source);
    const snapshots = [];
    const changed = [];
    for (let i = 0; i < withoutSource.length; i++) {
      const item = withoutSource[i];
      const newOrder = (i + 1) * 1000;
      if (item.order !== newOrder) {
        snapshots.push({ id: item.id, order: item.order });
        item.order = newOrder;
        changed.push(item);
      }
    }
    if (changed.length === 0) return;
    render();
    try {
      await Promise.all(changed.map((item) =>
        api("PATCH", `/reminders/${item.id}`, { order: item.order })
      ));
      announce(`Moved ${source.title}`);
    } catch (err) {
      for (const snap of snapshots) {
        const r2 = reminders.find((x) => x.id === snap.id);
        if (r2) r2.order = snap.order;
      }
      render();
      showError("Could not reorder", err);
    }
  }

  function autoGrowTextarea(ta, maxPx) {
    if (!ta) return;
    ta.style.height = "auto";
    const cap = typeof maxPx === "number" ? maxPx : 400;
    ta.style.height = Math.min(ta.scrollHeight, cap) + "px";
  }

  function openRowOptions(r) {
    editingReminderId = r.id;
    rowOptionsTitle.textContent = r.title;
    rowLeadMinutes.value = typeof r.leadMinutes === "number" ? String(r.leadMinutes) : "";
    if (rowDescription) rowDescription.value = r.description || "";
    if (rowDueDate) rowDueDate.value = r.dueAt || todayPh();
    if (rowClient) rowClient.value = r.client || "";
    const repeatSel = $("rowRepeat");
    if (repeatSel) repeatSel.value = r.repeat && r.repeat !== "none" ? r.repeat : "none";
    const doneBtn = $("rowOptionsDone");
    if (doneBtn) doneBtn.textContent = r.done ? "Mark not done" : (r.repeat && r.repeat !== "none" ? "Done (advance)" : "Mark done");
    openDialog(rowOptionsDialog);
    // scrollHeight needs the element to be in layout — modal is shown synchronously
    // by showModal(), but defer one frame to dodge any browser layout quirks.
    if (rowDescription) requestAnimationFrame(() => autoGrowTextarea(rowDescription));
  }

  // ---------- inline edit ----------
  function startTitleEdit(r, span) {
    if (span.querySelector("input")) return;
    const input = document.createElement("input");
    input.type = "text";
    input.value = r.title;
    input.className = "inline-edit";
    span.textContent = "";
    span.appendChild(input);
    input.focus();
    input.select();
    let committed = false;
    const commit = async () => {
      if (committed) return;
      committed = true;
      const { title, tags } = extractTagsFromTitle(input.value.trim());
      const wantTagMerge = tags.length && !sameTags(mergeTags(r.tags, tags), r.tags);
      if (!title || (title === r.title && !wantTagMerge)) {
        span.textContent = displayTitle(r);
        return;
      }
      const prevTitle = r.title;
      const prevTags = r.tags;
      r.title = title;
      if (wantTagMerge) r.tags = mergeTags(r.tags, tags);
      render();
      try {
        const patch = { title };
        if (wantTagMerge) patch.tags = r.tags;
        const updated = await api("PATCH", `/reminders/${r.id}`, patch);
        replaceLocal(updated.reminder);
        // Skipped post-API render: optimistic render above already painted the
        // new title/tags; server response matches except for server timestamps
        // we don't display in this row.
      } catch (err) {
        r.title = prevTitle;
        r.tags = prevTags;
        render();
        showError("Could not rename", err);
      }
    };
    input.addEventListener("blur", commit);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); input.blur(); }
      else if (e.key === "Escape") { committed = true; span.textContent = displayTitle(r); }
    });
  }

  function startTimeEdit(r, host) {
    if (host.querySelector("input")) return;
    const input = document.createElement("input");
    input.type = "time";
    input.value = r.time || "";
    input.className = "inline-edit time";
    host.textContent = "";
    host.appendChild(input);
    input.focus();
    let committed = false;
    const commit = async () => {
      if (committed) return;
      committed = true;
      const newTime = input.value || null;
      if (newTime === r.time) { renderWhen(host, r.time, r.leadMinutes); return; }
      const prevTime = r.time;
      r.time = newTime;
      render();
      try {
        const updated = await api("PATCH", `/reminders/${r.id}`, { time: newTime });
        replaceLocal(updated.reminder);
      } catch (err) {
        r.time = prevTime;
        render();
        showError("Could not change time", err);
      }
    };
    input.addEventListener("blur", commit);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); input.blur(); }
      else if (e.key === "Escape") { committed = true; renderWhen(host, r.time, r.leadMinutes); }
    });
  }

  function startDueDateEdit(r, host) {
    if (host.querySelector("input")) return;
    const input = document.createElement("input");
    input.type = "date";
    input.value = r.dueAt || todayPh();
    input.className = "inline-edit";
    const original = host.textContent;
    host.textContent = "";
    host.appendChild(input);
    input.focus();
    let committed = false;
    const commit = async () => {
      if (committed) return;
      committed = true;
      const newDue = input.value || null;
      if (!newDue || newDue === r.dueAt) { host.textContent = original; return; }
      const prev = r.dueAt;
      const prevRoll = r.rollDays;
      r.dueAt = newDue;
      r.rollDays = 0;
      render();
      try {
        const updated = await api("PATCH", `/reminders/${r.id}`, { dueAt: newDue });
        replaceLocal(updated.reminder);
      } catch (err) {
        r.dueAt = prev;
        r.rollDays = prevRoll;
        render();
        showError("Could not change date", err);
      }
    };
    input.addEventListener("blur", commit);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); input.blur(); }
      else if (e.key === "Escape") { committed = true; host.textContent = original; }
    });
  }

  function startDescriptionEdit(r, host) {
    if (host.querySelector("textarea")) return;
    const ta = document.createElement("textarea");
    ta.className = "inline-edit description-edit";
    ta.maxLength = 2000;
    ta.value = r.description || "";
    ta.placeholder = "Notes, links, sub-tasks... (Ctrl+Enter to save, Esc to cancel)";
    host.textContent = "";
    host.appendChild(ta);
    ta.focus();
    if (ta.value) ta.setSelectionRange(ta.value.length, ta.value.length);
    let committed = false;
    const commit = async () => {
      if (committed) return;
      committed = true;
      const newDesc = ta.value.trim().slice(0, 2000) || null;
      if (newDesc === (r.description || null)) { render(); return; }
      const prev = r.description;
      r.description = newDesc;
      render();
      try {
        const updated = await api("PATCH", `/reminders/${r.id}`, { description: newDesc || "" });
        replaceLocal(updated.reminder);
      } catch (err) {
        r.description = prev;
        render();
        showError("Could not update details", err);
      }
    };
    ta.addEventListener("blur", commit);
    ta.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); ta.blur(); }
      else if (e.key === "Escape") { committed = true; render(); }
    });
  }

  // Builds the title-meta row (client chip + due-date chip + overdue badge) shown
  // inside a reminder's title-wrap. Returns null if there's nothing to show.
  function buildTitleMeta(r) {
    const today = todayPh();
    const due = r.dueAt || r.createdDate || null;
    const showDueChip = due && due !== today;
    const showBadge = (r.rollDays || 0) > 0;
    const showClient = !!(r.client && r.client.trim());
    const showRepeat = r.repeat && r.repeat !== "none";
    if (!showDueChip && !showBadge && !showClient && !showRepeat) return null;
    const meta = document.createElement("div");
    meta.className = "title-meta";
    if (showClient) meta.appendChild(buildClientChip(r));
    if (showRepeat) {
      const tag = document.createElement("span");
      tag.className = "repeat-badge";
      const labels = { daily: "Daily", weekdays: "Weekdays", weekly: "Weekly" };
      tag.textContent = `↻ ${labels[r.repeat] || r.repeat}`;
      tag.title = `Recurring ${labels[r.repeat] || r.repeat}. Marking done advances to the next occurrence.`;
      meta.appendChild(tag);
    }
    if (showDueChip) {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "due-chip " + (due < today ? "past" : "future");
      chip.textContent = formatDueLabel(due, today);
      chip.title = `Due ${due}. Click to change.`;
      chip.addEventListener("click", (e) => { e.stopPropagation(); startDueDateEdit(r, chip); });
      meta.appendChild(chip);
    }
    if (showBadge) {
      const badge = document.createElement("span");
      badge.className = "overdue-badge";
      badge.textContent = `overdue ${r.rollDays}d`;
      badge.title = `Rolled forward from ${r.rollDays} day${r.rollDays === 1 ? "" : "s"} ago`;
      meta.appendChild(badge);
    }
    return meta;
  }

  // Client chip: distinct from tag chips. Clicking toggles the client filter
  // (matches the tag-chip filter UX). Long-press / right-click would be nice
  // for inline-edit but for now we surface edit via the bot row options dialog
  // and let the chip stay one-purpose (filter).
  function buildClientChip(r) {
    const isActive = activeClientFilter && activeClientFilter.toLowerCase() === r.client.toLowerCase();
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "client-chip" + (isActive ? " active" : "");
    chip.style.setProperty("--client-color", colorForClient(r.client));
    chip.textContent = r.client;
    chip.title = isActive
      ? `Clear client filter. Right-click or hold Shift+click to edit.`
      : `Filter to client "${r.client}". Right-click or hold Shift+click to edit.`;
    chip.addEventListener("click", (e) => {
      e.stopPropagation();
      // Shift+click or Alt+click → edit, plain click → filter toggle
      if (e.shiftKey || e.altKey) {
        startClientEdit(r, chip);
        return;
      }
      setClientFilter(isActive ? null : r.client);
    });
    chip.addEventListener("contextmenu", (e) => {
      e.preventDefault();
      e.stopPropagation();
      startClientEdit(r, chip);
    });
    return chip;
  }

  // Inline-edit a reminder's client. Uses the same datalist as the add form so
  // autocomplete works. Blank submission clears the client.
  function startClientEdit(r, host) {
    if (host.querySelector("input")) return;
    const input = document.createElement("input");
    input.type = "text";
    input.setAttribute("list", "clientList");
    input.maxLength = 100;
    input.value = r.client || "";
    input.className = "inline-edit";
    input.placeholder = "Client (blank to clear)";
    const original = host.textContent;
    host.textContent = "";
    host.appendChild(input);
    input.focus();
    input.select();
    let committed = false;
    const commit = async () => {
      if (committed) return;
      committed = true;
      const newClient = input.value.trim().slice(0, 100) || null;
      if (newClient === (r.client || null)) { host.textContent = original; return; }
      const prev = r.client;
      r.client = newClient;
      render();
      try {
        const updated = await api("PATCH", `/reminders/${r.id}`, { client: newClient || "" });
        replaceLocal(updated.reminder);
      } catch (err) {
        r.client = prev;
        render();
        showError("Could not change client", err);
      }
    };
    input.addEventListener("blur", commit);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); input.blur(); }
      else if (e.key === "Escape") { committed = true; host.textContent = original; }
    });
  }

  // Refresh the <datalist> for client autocomplete from the unique non-empty
  // client values across all current reminders. Fingerprinted so repeat calls
  // with unchanged client set are no-ops (perf — called from render()).
  let _clientListFingerprint = null;
  function populateClientList() {
    if (!clientList) return;
    const set = new Set();
    for (const r of reminders) {
      if (r.client && r.client.trim()) set.add(r.client.trim());
    }
    const sorted = [...set].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
    const fp = sorted.join("|");
    if (fp === _clientListFingerprint) return;
    _clientListFingerprint = fp;
    clientList.textContent = "";
    for (const c of sorted) {
      const opt = document.createElement("option");
      opt.value = c;
      clientList.appendChild(opt);
    }
  }

  // Builds the collapsible description block shown under a reminder row when
  // the reminder has a description. Returns null when empty to skip rendering
  // an entire <details> element per row (perf — saves ~100+ DOM nodes per
  // render in typical use). To ADD a description to an existing reminder, use
  // the "..." row options dialog which now has a description textarea.
  function buildDescriptionBlock(r) {
    if (!r.description || !r.description.trim()) return null;
    const det = document.createElement("details");
    det.className = "description-block";
    const sum = document.createElement("summary");
    sum.textContent = "Details";
    det.appendChild(sum);
    const body = document.createElement("div");
    body.className = "description-body";
    body.textContent = r.description;
    body.title = "Click to edit";
    body.addEventListener("click", (e) => { e.stopPropagation(); startDescriptionEdit(r, body); });
    det.appendChild(body);
    return det;
  }

  // ---------- model ops (optimistic) ----------
  function removeReminder(r) {
    if (pendingDeletes.has(r.id)) return;
    // optimistic delete: hide row immediately, fire DELETE after UNDO_MS unless undone
    const timer = setTimeout(async () => {
      const entry = pendingDeletes.get(r.id);
      pendingDeletes.delete(r.id);
      reminders = reminders.filter((x) => x.id !== r.id);
      render();
      if (entry && entry.toast) entry.toast.remove();
      try {
        await api("DELETE", `/reminders/${r.id}`);
      } catch (err) {
        if (err.status !== 404) showError("Could not delete", err);
      }
    }, UNDO_MS);
    const toast = showUndoToast(r);
    pendingDeletes.set(r.id, { reminder: r, timer, toast });
    render();
    announce(`Deleted ${r.title}, undo available`);
  }

  function undoDelete(id) {
    const entry = pendingDeletes.get(id);
    if (!entry) return;
    clearTimeout(entry.timer);
    pendingDeletes.delete(id);
    if (entry.toast) entry.toast.remove();
    render();
    announce(`Restored ${entry.reminder.title}`);
  }

  async function toggleDone(r) {
    const prevDone = r.done;
    const prevClosedAt = r.closedAt;
    r.done = !prevDone;
    r.closedAt = r.done ? new Date().toISOString() : null;
    render();
    try {
      const updated = await api("PATCH", `/reminders/${r.id}`, { done: r.done });
      replaceLocal(updated.reminder);
    } catch (err) {
      r.done = prevDone;
      r.closedAt = prevClosedAt;
      render();
      showError("Could not update", err);
    }
  }

  async function togglePriority(r) {
    const prev = r.priority || "normal";
    const next = prev === "high" ? "normal" : "high";
    r.priority = next;
    render();
    try {
      const updated = await api("PATCH", `/reminders/${r.id}`, { priority: next });
      replaceLocal(updated.reminder);
    } catch (err) {
      r.priority = prev;
      render();
      showError("Could not change priority", err);
    }
  }

  function currentVisiblePool() {
    const norm = (s) => String(s || "").toLowerCase();
    const sn = norm(searchText.trim());
    let v = reminders.filter((r) => !pendingDeletes.has(r.id));
    if (activeTagFilter) {
      const tn = norm(activeTagFilter);
      v = v.filter((r) => (r.tags || []).some((t) => norm(t) === tn));
    }
    if (activeClientFilter) {
      const cn = norm(activeClientFilter);
      v = v.filter((r) => norm(r.client) === cn);
    }
    if (sn) {
      v = v.filter((r) =>
        norm(r.title).includes(sn) ||
        (r.tags || []).some((t) => norm(t).includes(sn)) ||
        norm(r.client).includes(sn) ||
        norm(r.description).includes(sn)
      );
    }
    return v;
  }

  async function markAllOpenDone() {
    let openVisible = currentVisiblePool().filter((r) => !r.done);
    if (quickFilter === "timed") openVisible = openVisible.filter((r) => !!r.time);
    else if (quickFilter === "anytime") openVisible = openVisible.filter((r) => !r.time);
    else if (quickFilter === "high") openVisible = openVisible.filter((r) => r.priority === "high");
    if (openVisible.length === 0) return;
    if (openVisible.length > 1 && !confirm(`Mark all ${openVisible.length} open reminders as done?`)) return;
    const snapshot = openVisible.map((r) => ({ id: r.id, done: r.done, closedAt: r.closedAt }));
    const now = new Date().toISOString();
    for (const r of openVisible) { r.done = true; r.closedAt = now; }
    render();
    try {
      const ids = openVisible.map((r) => r.id);
      await api("POST", "/reminders/bulk", { ids, patch: { done: true } });
      announce(`${openVisible.length} reminders marked done`);
    } catch (err) {
      for (const snap of snapshot) {
        const r = reminders.find((x) => x.id === snap.id);
        if (r) { r.done = snap.done; r.closedAt = snap.closedAt; }
      }
      render();
      showError("Could not mark all done", err);
    }
  }

  async function clearDoneVisible() {
    const inFilter = currentVisiblePool();
    const targets = inFilter.filter((r) => r.done && (settings.showAllDone || !r.closedAt || (Date.now() - new Date(r.closedAt).getTime()) < DONE_AGE_MS));
    if (targets.length === 0) return;
    const snapshot = targets.map((r) => ({ ...r }));
    const ids = new Set(targets.map((r) => r.id));
    reminders = reminders.filter((r) => !ids.has(r.id));
    render();
    try {
      await Promise.all(targets.map((r) => api("DELETE", `/reminders/${r.id}`).catch((e) => { if (e.status !== 404) throw e; })));
      announce(`Cleared ${targets.length} done`);
    } catch (err) {
      // best-effort restore
      reminders = reminders.concat(snapshot);
      render();
      showError("Could not clear done", err);
    }
  }

  function replaceLocal(r) {
    const idx = reminders.findIndex((x) => x.id === r.id);
    if (idx >= 0) reminders[idx] = r;
    else reminders.push(r);
  }
  function replaceById(oldId, r) {
    const idx = reminders.findIndex((x) => x.id === oldId);
    if (idx >= 0) reminders[idx] = r;
    else reminders.push(r);
  }

  // ---------- filters ----------
  function setTagFilter(tag) {
    activeTagFilter = tag;
    try {
      if (tag) localStorage.setItem(LS_TAG_FILTER, tag);
      else localStorage.removeItem(LS_TAG_FILTER);
    } catch (_) {}
    if (tag) announce(`Filtered to ${tag}`);
    else announce("Tag filter cleared");
    render();
  }
  function setClientFilter(client) {
    activeClientFilter = client;
    try {
      if (client) localStorage.setItem(LS_CLIENT_FILTER, client);
      else localStorage.removeItem(LS_CLIENT_FILTER);
    } catch (_) {}
    if (client) announce(`Filtered to client ${client}`);
    else announce("Client filter cleared");
    render();
  }
  function setQuickFilter(key) {
    if (!["all", "timed", "anytime", "high", "done"].includes(key)) return;
    quickFilter = key;
    try { localStorage.setItem(LS_QUICK_FILTER, key); } catch (_) {}
    announce(key === "all" ? "Showing all" : `Quick filter: ${key}`);
    render();
  }
  function clearSearch() {
    searchText = "";
    if (searchInput) searchInput.value = "";
    if (searchClear) searchClear.hidden = true;
    try { localStorage.removeItem(LS_SEARCH); } catch (_) {}
    render();
  }
  function clearAllFilters() {
    activeTagFilter = null;
    activeClientFilter = null;
    quickFilter = "all";
    searchText = "";
    if (searchInput) searchInput.value = "";
    if (searchClear) searchClear.hidden = true;
    try {
      localStorage.removeItem(LS_TAG_FILTER);
      localStorage.removeItem(LS_CLIENT_FILTER);
      localStorage.removeItem(LS_SEARCH);
      localStorage.setItem(LS_QUICK_FILTER, "all");
    } catch (_) {}
    announce("All filters cleared");
    render();
  }
  function updateFilterBanner() {
    if (!filterBanner || !filterCrumbs) return;
    filterCrumbs.textContent = "";
    const crumbs = [];
    if (quickFilter !== "all") {
      const labels = { timed: "Timed", anytime: "Anytime", high: "Priority", done: "Done" };
      crumbs.push({ kind: "quick", text: labels[quickFilter] });
    }
    if (activeTagFilter) crumbs.push({ kind: "tag", text: `#${activeTagFilter}` });
    if (activeClientFilter) crumbs.push({ kind: "client", text: `client: ${activeClientFilter}` });
    if (searchText) crumbs.push({ kind: "search", text: `"${truncate(searchText, 30)}"` });
    if (crumbs.length === 0) {
      filterBanner.hidden = true;
      return;
    }
    filterBanner.hidden = false;
    const intro = document.createElement("span");
    intro.textContent = "Showing ";
    intro.style.marginRight = "2px";
    filterCrumbs.appendChild(intro);
    for (const c of crumbs) {
      const chip = document.createElement("span");
      chip.className = "crumb";
      chip.textContent = c.text;
      filterCrumbs.appendChild(chip);
    }
  }
  function updateQuickFilterPills() {
    if (!filterPills) return;
    for (const btn of filterPills.querySelectorAll("button.filter-pill[data-quick]")) {
      btn.setAttribute("aria-pressed", String(btn.dataset.quick === quickFilter));
    }
  }
  function setView(v) {
    if (!VIEWS.includes(v) || v === currentView) return;
    currentView = v;
    try { localStorage.setItem(LS_VIEW, v); } catch (_) {}
    announce(`View: ${v}`);
    render();
  }
  function cycleView() {
    const idx = VIEWS.indexOf(currentView);
    setView(VIEWS[(idx + 1) % VIEWS.length]);
  }
  function updateViewSwitch() {
    if (!viewSwitch) return;
    for (const btn of viewSwitch.querySelectorAll("button.view-btn")) {
      btn.setAttribute("aria-pressed", String(btn.dataset.view === currentView));
    }
  }

  // ---------- bulk select ----------
  function setBulkMode(on) {
    bulkMode = !!on;
    bulkSelected.clear();
    document.body.classList.toggle("bulk-mode", bulkMode);
    if (bulkToggleBtn) {
      bulkToggleBtn.textContent = bulkMode ? "Cancel select" : "Select";
      bulkToggleBtn.setAttribute("aria-pressed", String(bulkMode));
    }
    render();
  }
  function toggleBulkSelected(r) {
    if (bulkSelected.has(r.id)) bulkSelected.delete(r.id);
    else bulkSelected.add(r.id);
    render();
  }
  function updateBulkBar() {
    if (!bulkBar) return;
    if (!bulkMode || bulkSelected.size === 0) {
      bulkBar.hidden = true;
      return;
    }
    bulkBar.hidden = false;
    bulkCountLabel.textContent = `${bulkSelected.size} selected`;
  }
  function selectedReminders() {
    return reminders.filter((r) => bulkSelected.has(r.id) && !pendingDeletes.has(r.id));
  }
  async function bulkMarkDone() {
    const targets = selectedReminders().filter((r) => !r.done);
    if (targets.length === 0) { setBulkMode(false); return; }
    const snapshot = targets.map((r) => ({ id: r.id, done: r.done, closedAt: r.closedAt }));
    const now = new Date().toISOString();
    for (const r of targets) { r.done = true; r.closedAt = now; }
    setBulkMode(false);
    try {
      await api("POST", "/reminders/bulk", { ids: targets.map((r) => r.id), patch: { done: true } });
      announce(`${targets.length} marked done`);
    } catch (err) {
      for (const snap of snapshot) {
        const rr = reminders.find((x) => x.id === snap.id);
        if (rr) { rr.done = snap.done; rr.closedAt = snap.closedAt; }
      }
      render();
      showError("Could not mark all selected", err);
    }
  }
  async function bulkDelete() {
    const targets = selectedReminders();
    if (targets.length === 0) { setBulkMode(false); return; }
    setBulkMode(false);
    try {
      await Promise.all(targets.map((r) =>
        api("DELETE", `/reminders/${r.id}`).catch((e) => { if (e.status !== 404) throw e; })
      ));
      reminders = reminders.filter((r) => !targets.some((t) => t.id === r.id));
      render();
      announce(`${targets.length} deleted`);
    } catch (err) {
      showError("Could not delete all selected", err);
    }
  }
  async function bulkTogglePriority() {
    const targets = selectedReminders();
    if (targets.length === 0) { setBulkMode(false); return; }
    // If any are not high, raise all to high; else demote all to normal.
    const promoteAll = targets.some((r) => r.priority !== "high");
    const nextPriority = promoteAll ? "high" : "normal";
    const snapshot = targets.map((r) => ({ id: r.id, priority: r.priority }));
    for (const r of targets) r.priority = nextPriority;
    setBulkMode(false);
    try {
      await api("POST", "/reminders/bulk", { ids: targets.map((r) => r.id), patch: { priority: nextPriority } });
      announce(`${targets.length} updated`);
    } catch (err) {
      for (const snap of snapshot) {
        const rr = reminders.find((x) => x.id === snap.id);
        if (rr) rr.priority = snap.priority;
      }
      render();
      showError("Could not update all selected", err);
    }
  }

  // ---------- templates ----------
  function buildTemplateGrid() {
    if (!templateGrid) return;
    templateGrid.textContent = "";

    const renderTemplate = (t, isUser) => {
      const wrap = document.createElement("div");
      wrap.className = "template-wrap";

      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "template";
      const title = document.createElement("strong");
      title.textContent = t.title;
      const sub = document.createElement("span");
      sub.className = "sub";
      const parts = [];
      if (t.time) parts.push(`at ${formatTime(t.time)}`);
      else parts.push("no time");
      if (t.client) parts.push(`for ${t.client}`);
      sub.textContent = parts.join(" · ");
      btn.append(title, sub);
      btn.addEventListener("click", async () => {
        templatesDialog.close();
        await addFromTemplate(t);
      });
      wrap.appendChild(btn);

      if (isUser) {
        const del = document.createElement("button");
        del.type = "button";
        del.className = "template-delete";
        del.title = "Delete this template";
        del.setAttribute("aria-label", `Delete template ${t.title}`);
        del.textContent = "×";
        del.addEventListener("click", async (e) => {
          e.stopPropagation();
          await deleteUserTemplate(t);
        });
        wrap.appendChild(del);
      }
      return wrap;
    };

    if (userTemplates.length) {
      const head = document.createElement("h3");
      head.className = "template-section";
      head.textContent = "Your templates";
      templateGrid.appendChild(head);
      for (const t of userTemplates) templateGrid.appendChild(renderTemplate(t, true));
      const head2 = document.createElement("h3");
      head2.className = "template-section";
      head2.textContent = "Built-in";
      templateGrid.appendChild(head2);
    }
    for (const t of TEMPLATES) templateGrid.appendChild(renderTemplate(t, false));
  }

  async function addFromTemplate(t) {
    const { title, tags: titleTags } = extractTagsFromTitle(t.title);
    if (!title) return;
    const tags = Array.isArray(t.tags) && t.tags.length ? mergeTags(t.tags, titleTags) : titleTags;
    const today = todayPh();
    const tempId = `tmp-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
    const optimistic = {
      id: tempId, title, time: t.time || null, tags, dueAt: today,
      description: t.description || null, rollDays: 0, client: t.client || null,
      leadMinutes: typeof t.leadMinutes === "number" ? t.leadMinutes : null,
      done: false, priority: "normal", closedAt: null, _optimistic: true,
    };
    reminders.push(optimistic);
    render();
    try {
      const created = await api("POST", "/reminders", {
        title, time: t.time || null, tags, dueAt: today,
        client: t.client || null, description: t.description || null,
        leadMinutes: typeof t.leadMinutes === "number" ? t.leadMinutes : null,
      });
      replaceById(tempId, created.reminder);
      render();
      announce(`Added template: ${title}`);
    } catch (err) {
      reminders = reminders.filter((r) => r.id !== tempId);
      render();
      showError("Could not add template", err);
    }
  }

  async function saveAsTemplate(r) {
    if (!r || !r.title) return;
    const tpl = {
      title: r.title,
      time: r.time || null,
      client: r.client || null,
      description: r.description || null,
      leadMinutes: typeof r.leadMinutes === "number" ? r.leadMinutes : null,
      tags: Array.isArray(r.tags) ? r.tags : [],
    };
    // dedup by title+time+client so saving the same thing twice is idempotent
    const key = `${tpl.title}|${tpl.time || ""}|${tpl.client || ""}`;
    const without = userTemplates.filter((x) => `${x.title}|${x.time || ""}|${x.client || ""}` !== key);
    const next = [tpl, ...without].slice(0, 100);
    const prev = userTemplates;
    userTemplates = next;
    try {
      const saved = await api("PUT", "/templates", { templates: next });
      userTemplates = Array.isArray(saved.templates) ? saved.templates : next;
      announce(`Saved template: ${tpl.title}`);
    } catch (err) {
      userTemplates = prev;
      showError("Could not save template", err);
    }
  }

  async function deleteUserTemplate(t) {
    const key = `${t.title}|${t.time || ""}|${t.client || ""}`;
    const prev = userTemplates;
    const next = userTemplates.filter((x) => `${x.title}|${x.time || ""}|${x.client || ""}` !== key);
    userTemplates = next;
    buildTemplateGrid();
    try {
      const saved = await api("PUT", "/templates", { templates: next });
      userTemplates = Array.isArray(saved.templates) ? saved.templates : next;
      buildTemplateGrid();
    } catch (err) {
      userTemplates = prev;
      buildTemplateGrid();
      showError("Could not delete template", err);
    }
  }

  // ---------- toasts ----------
  function showUndoToast(r) {
    const toast = document.createElement("div");
    toast.className = "toast";
    const label = document.createElement("span");
    label.textContent = `Deleted "${truncate(r.title, 40)}"`;
    const action = document.createElement("button");
    action.type = "button";
    action.className = "toast-action";
    action.textContent = "Undo";
    action.addEventListener("click", () => undoDelete(r.id));
    toast.append(label, action);
    toastRegion.appendChild(toast);
    return toast;
  }

  // ---------- API ----------
  async function api(method, path, body) {
    if (!authToken) throw new Error("Not signed in");
    const res = await fetch(`${API_BASE}${path}`, {
      method,
      headers: {
        Authorization: `Bearer ${authToken}`,
        "Content-Type": "application/json",
      },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (res.status === 204) return {};
    const data = await res.json().catch(() => ({}));
    if (!res.ok) {
      const err = new Error(data.error || `HTTP ${res.status}`);
      err.status = res.status;
      throw err;
    }
    return data;
  }

  // ---------- dialogs ----------
  function openDialog(dlg, beforeShow) {
    lastFocusedTrigger = document.activeElement;
    if (beforeShow) beforeShow();
    if (typeof dlg.showModal === "function") dlg.showModal();
    else dlg.setAttribute("open", "");
  }
  function openSettings() {
    $("setEodTime").value = settings.eodTime;
    $("setLeadMinutes").value = settings.leadMinutes;
    $("setWeekdaysOnly").checked = !!settings.weekdaysOnly;
    $("setNotifications").checked = settings.notifications !== false;
    $("setQuietStart").value = settings.quietStart || "";
    $("setQuietEnd").value = settings.quietEnd || "";
    $("setThemeOverride").value = settings.themeOverride || "auto";
  }

  function updateBotHint() {
    if (hasBot) {
      permHint.hidden = true;
      permHint.textContent = "";
      return;
    }
    permHint.hidden = false;
    permHint.textContent = "Heads up: I haven't seen the Day Reminders bot in your chat yet. Open the app's chat once (left rail) so I can send reminders there.";
  }

  // ---------- utils ----------
  function extractTagsFromTitle(input) {
    if (!input) return { title: "", tags: [] };
    const tokens = input.split(/\s+/).filter(Boolean);
    const tags = [];
    const rest = [];
    for (const t of tokens) {
      if (t.startsWith("#") && t.length > 1) tags.push(t.slice(1));
      else rest.push(t);
    }
    return { title: rest.join(" ").trim(), tags: tags.slice(0, 8) };
  }
  function mergeTags(existing, incoming) {
    const seen = new Set();
    const out = [];
    for (const t of [...(existing || []), ...incoming]) {
      const k = t.toLowerCase();
      if (seen.has(k)) continue;
      seen.add(k);
      out.push(t);
    }
    return out.slice(0, 8);
  }
  function sameTags(a, b) {
    if ((a || []).length !== (b || []).length) return false;
    const sa = [...(a || [])].sort();
    const sb = [...(b || [])].sort();
    return sa.every((v, i) => v === sb[i]);
  }
  // Memoized — same tag string always hashes to the same palette index, so
  // caching avoids the per-character loop on every render with many rows.
  const _tagColorCache = new Map();
  function colorForTag(tag) {
    const hit = _tagColorCache.get(tag);
    if (hit) return hit;
    let h = 0;
    for (let i = 0; i < tag.length; i++) h = (h * 31 + tag.charCodeAt(i)) & 0xffffffff;
    const color = TAG_PALETTE[Math.abs(h) % TAG_PALETTE.length];
    _tagColorCache.set(tag, color);
    return color;
  }
  // Same palette as tags but seeded differently so a string like "Citadel" used as
  // both a tag and a client gets visually distinct colors instead of colliding.
  const _clientColorCache = new Map();
  function colorForClient(name) {
    const hit = _clientColorCache.get(name);
    if (hit) return hit;
    let h = 13;
    for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) & 0xffffffff;
    const color = TAG_PALETTE[Math.abs(h) % TAG_PALETTE.length];
    _clientColorCache.set(name, color);
    return color;
  }
  function renderWhen(el, time, leadMinutes) {
    el.classList.remove("empty");
    el.textContent = "";
    if (!time) {
      el.textContent = "Set time";
      el.classList.add("empty");
      return;
    }
    el.append(document.createTextNode(formatTime(time)));
    if (typeof leadMinutes === "number") {
      const badge = document.createElement("span");
      badge.className = "lead-badge";
      badge.textContent = `−${leadMinutes}m`;
      badge.title = `Custom lead time: ${leadMinutes} minutes before`;
      el.append(badge);
    }
  }
  function formatTime(hhmm) {
    const [h, m] = hhmm.split(":").map(Number);
    const d = new Date();
    d.setHours(h, m, 0, 0);
    return d.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" });
  }
  function formatSnoozeRelative(iso) {
    const target = new Date(iso).getTime();
    const ms = target - Date.now();
    if (ms <= 0) return "now";
    const mins = Math.round(ms / 60000);
    if (mins < 60) return `${mins}m left`;
    const hours = Math.round(mins / 60);
    if (hours < 24) return `${hours}h left`;
    // Different day — show local wall clock
    return `until ${new Date(target).toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" })}`;
  }
  async function unsnooze(r) {
    const prev = r.snoozedUntil;
    r.snoozedUntil = null;
    render();
    try {
      const updated = await api("PATCH", `/reminders/${r.id}`, { snoozedUntil: null });
      replaceLocal(updated.reminder);
      announce(`Cleared snooze on ${r.title}`);
    } catch (err) {
      r.snoozedUntil = prev;
      render();
      showError("Could not clear snooze", err);
    }
  }
  function clampInt(v, min, max, fallback) {
    const n = parseInt(v, 10);
    if (isNaN(n)) return fallback;
    return Math.max(min, Math.min(max, n));
  }
  function showError(prefix, err) {
    console.error(prefix, err);
    postTelemetry("tab.error", {
      message: `${prefix}: ${err && err.message ? err.message : String(err)}`,
      stack: err && err.stack ? String(err.stack) : "",
    });
    const banner = $("errorBanner");
    if (banner) {
      banner.textContent = `${prefix}: ${err.message || err}`;
      banner.hidden = false;
      setTimeout(() => { banner.hidden = true; }, 5000);
    }
  }

  function postTelemetry(event, payload) {
    try {
      const headers = { "Content-Type": "application/json" };
      if (authToken) headers.Authorization = `Bearer ${authToken}`;
      fetch(`${API_BASE}/telemetry`, {
        method: "POST",
        headers,
        body: JSON.stringify({
          event,
          version: APP_VERSION,
          url: location.href,
          ...(payload || {}),
        }),
        keepalive: true,
      }).catch(() => {});
    } catch (_) {}
  }

  window.addEventListener("error", (e) => {
    postTelemetry("tab.error", {
      message: String(e.message || e),
      stack: e.error && e.error.stack ? String(e.error.stack) : "",
    });
  });
  window.addEventListener("unhandledrejection", (e) => {
    const reason = e.reason || {};
    postTelemetry("tab.unhandledrejection", {
      message: reason.message ? String(reason.message) : String(reason),
      stack: reason.stack ? String(reason.stack) : "",
    });
  });
  function announce(msg) {
    if (!liveRegion) return;
    liveRegion.textContent = "";
    setTimeout(() => { liveRegion.textContent = msg; }, 50);
  }
  function truncate(s, n) {
    s = String(s || "");
    return s.length > n ? s.slice(0, n - 1) + "…" : s;
  }

  // ---------- Teams init + boot ----------
  async function boot() {
    if (!window.microsoftTeams) {
      showError("Teams SDK failed to load", new Error("microsoftTeams missing"));
      return;
    }
    try {
      await microsoftTeams.app.initialize();
      const ctx = await microsoftTeams.app.getContext();
      applyTheme(ctx.app?.theme || "default");
      microsoftTeams.app.registerOnThemeChangeHandler(applyTheme);

      try {
        authToken = await microsoftTeams.authentication.getAuthToken();
      } catch (e) {
        const msg = (e && (e.message || e.errorCode || String(e))) || "unknown";
        throw new Error(`SSO failed: ${msg}`);
      }

      const [{ settings: s, hasBot: hb }, { reminders: rems }, { templates: userTpls }] = await Promise.all([
        api("GET", "/settings"),
        api("GET", "/reminders"),
        api("GET", "/templates").catch(() => ({ templates: [] })),
      ]);
      Object.assign(settings, s);
      // restore tab-only override after server settings merge
      try {
        const saved = localStorage.getItem("themeOverride");
        if (saved) settings.themeOverride = saved;
      } catch (_) {}
      hasBot = !!hb;
      reminders = rems;
      userTemplates = Array.isArray(userTpls) ? userTpls : [];
      render();
    } catch (err) {
      console.error("Boot failed", err);
      postTelemetry("tab.boot.failed", {
        message: err && err.message ? err.message : String(err),
        stack: err && err.stack ? String(err.stack) : "",
      });
      const banner = $("errorBanner");
      if (banner) {
        banner.textContent = `Could not connect: ${err.message || err}`;
        banner.hidden = false;
      }
    }
  }
  function applyTheme(theme) {
    teamsTheme = theme;
    const effective = settings.themeOverride && settings.themeOverride !== "auto"
      ? settings.themeOverride
      : (theme === "dark" ? "dark" : theme === "contrast" ? "contrast" : "default");
    document.body.dataset.theme = effective;
  }

  boot();
})();
