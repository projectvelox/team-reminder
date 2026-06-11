/* Day Reminders — Teams personal tab (v1.2)
   Thin client over the bot's REST API. Auth via Teams SSO.
   Server-side bot handles all notifications (proactive Adaptive Cards in chat).
*/
(function () {
  "use strict";

  const API_BASE = "https://func-day-reminders-17023.azurewebsites.net/api";
  const DONE_AGE_MS = 24 * 60 * 60 * 1000;
  const TAG_PALETTE = [
    "#0078d4", "#107c10", "#8764b8", "#ca5010", "#c50f1f",
    "#038387", "#d83b01", "#5c2d91", "#0099bc", "#498205",
  ];

  // ---------- state ----------
  let reminders = [];
  let settings = {
    eodTime: "17:00",
    leadMinutes: 10,
    weekdaysOnly: true,
    notifications: true,
    groupByTag: false,
    showAllDone: false,
  };
  let hasBot = false;
  let authToken = null;

  // ---------- DOM ----------
  const $ = (id) => document.getElementById(id);
  const addForm = $("addForm");
  const titleInput = $("title");
  const timeInput = $("time");
  const settingsDialog = $("settingsDialog");
  const settingsForm = $("settingsForm");
  const permHint = $("permHint");
  const reminderRoot = $("reminderRoot");
  const groupToggle = $("groupToggle");

  const todayDateString = new Date().toLocaleDateString(undefined, {
    weekday: "long", month: "short", day: "numeric",
  });

  // ---------- events ----------
  addForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const { title, tags } = extractTagsFromTitle(titleInput.value.trim());
    if (!title) return;
    try {
      const created = await api("POST", "/reminders", { title, time: timeInput.value || null, tags });
      reminders.push(created.reminder);
      titleInput.value = "";
      timeInput.value = "";
      titleInput.focus();
      render();
    } catch (err) {
      showError("Could not add reminder", err);
    }
  });

  $("openSettings").addEventListener("click", openSettings);
  $("settingsCancel").addEventListener("click", () => settingsDialog.close());
  settingsForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const next = {
      eodTime: $("setEodTime").value || settings.eodTime,
      leadMinutes: clampInt($("setLeadMinutes").value, 0, 240, settings.leadMinutes),
      weekdaysOnly: $("setWeekdaysOnly").checked,
      notifications: $("setNotifications").checked,
    };
    try {
      const saved = await api("PUT", "/settings", { settings: next });
      Object.assign(settings, saved.settings);
      settingsDialog.close();
    } catch (err) {
      showError("Could not save settings", err);
    }
  });

  groupToggle.addEventListener("click", () => {
    settings.groupByTag = !settings.groupByTag;
    groupToggle.setAttribute("aria-pressed", String(settings.groupByTag));
    render();
  });

  // ---------- render ----------
  function render() {
    const open = reminders.filter((r) => !r.done);
    const done = reminders.filter((r) => r.done);
    const recentDone = settings.showAllDone
      ? done
      : done.filter((r) => !r.closedAt || (Date.now() - new Date(r.closedAt).getTime()) < DONE_AGE_MS);
    const olderDoneCount = done.length - recentDone.length;

    reminderRoot.innerHTML = "";

    if (settings.groupByTag) {
      renderByTag(open);
    } else {
      renderByTime(open);
    }

    if (recentDone.length || olderDoneCount) {
      reminderRoot.appendChild(buildDoneSection(recentDone, olderDoneCount));
    }

    updateBotHint();
  }

  function renderByTime(items) {
    const high = items.filter((r) => r.priority === "high");
    const timed = items.filter((r) => r.priority !== "high" && r.time)
      .sort((a, b) => a.time.localeCompare(b.time));
    const anytime = items.filter((r) => r.priority !== "high" && !r.time);

    if (high.length) {
      reminderRoot.appendChild(buildSection("High priority", high, { showWhen: true, emptyText: null }));
    }
    reminderRoot.appendChild(buildSection("Today", timed, {
      showWhen: true,
      emptyText: timed.length ? null : "Nothing scheduled. Add something above.",
      meta: todayDateString,
    }));
    reminderRoot.appendChild(buildSection("Anytime today", anytime, {
      showWhen: false,
      emptyText: anytime.length ? null : "None.",
      meta: "No specific time",
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
      const section = buildSection(`#${tag}`, sortReminders(buckets.get(tag)), { showWhen: true, emptyText: null });
      const head = section.querySelector("h2");
      if (head) head.style.color = colorForTag(tag);
      reminderRoot.appendChild(section);
    }
    if (untagged.length) {
      reminderRoot.appendChild(buildSection("No tag", sortReminders(untagged), { showWhen: true, emptyText: null }));
    }
    if (tagNames.length === 0 && untagged.length === 0) {
      reminderRoot.appendChild(buildSection("Today", [], { showWhen: true, emptyText: "Nothing scheduled.", meta: todayDateString }));
    }
  }

  function sortReminders(items) {
    return [...items].sort((a, b) => {
      if (a.priority !== b.priority) return a.priority === "high" ? -1 : 1;
      return (a.time || "zz").localeCompare(b.time || "zz");
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

    const ul = document.createElement("ul");
    ul.className = "list";
    for (const r of items) ul.appendChild(buildRow(r, opts.showWhen));
    section.appendChild(ul);
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
    for (const r of recent) ul.appendChild(buildRow(r, true));
    section.appendChild(ul);
    return section;
  }

  function buildRow(r, showWhen) {
    const li = document.createElement("li");
    if (r.firedAt) li.classList.add("fired");
    if (r.priority === "high") li.classList.add("high");

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
    title.textContent = r.title;
    title.title = "Click to rename";
    title.addEventListener("click", () => startTitleEdit(r, title));
    titleWrap.appendChild(title);

    if (r.tags && r.tags.length) {
      const tagRow = document.createElement("div");
      tagRow.className = "tag-row";
      for (const t of r.tags) {
        const chip = document.createElement("span");
        chip.className = "tag-chip";
        chip.style.backgroundColor = colorForTag(t);
        chip.textContent = `#${t}`;
        tagRow.appendChild(chip);
      }
      titleWrap.appendChild(tagRow);
    }

    const when = document.createElement("span");
    when.className = "when";
    when.textContent = showWhen && r.time ? formatTime(r.time) : (r.time ? formatTime(r.time) : "");
    when.title = r.time ? "Click to change time" : "Click to set a time";
    when.tabIndex = 0;
    when.addEventListener("click", () => startTimeEdit(r, when));

    const del = document.createElement("button");
    del.className = "icon-btn";
    del.title = "Delete";
    del.setAttribute("aria-label", "Delete reminder");
    del.textContent = "✕";
    del.addEventListener("click", () => removeReminder(r));

    li.append(star, cb, titleWrap, when, del);
    return li;
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
      if (!title || title === r.title && (!tags.length || sameTags(tags, r.tags))) {
        span.textContent = r.title;
        return;
      }
      try {
        const patch = { title };
        if (tags.length) patch.tags = mergeTags(r.tags, tags);
        const updated = await api("PATCH", `/reminders/${r.id}`, patch);
        replaceLocal(updated.reminder);
        render();
      } catch (err) {
        showError("Could not rename", err);
        span.textContent = r.title;
      }
    };
    input.addEventListener("blur", commit);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); input.blur(); }
      else if (e.key === "Escape") { committed = true; span.textContent = r.title; }
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
      if (newTime === r.time) { host.textContent = r.time ? formatTime(r.time) : ""; return; }
      try {
        const updated = await api("PATCH", `/reminders/${r.id}`, { time: newTime });
        replaceLocal(updated.reminder);
        render();
      } catch (err) {
        showError("Could not change time", err);
        host.textContent = r.time ? formatTime(r.time) : "";
      }
    };
    input.addEventListener("blur", commit);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); input.blur(); }
      else if (e.key === "Escape") { committed = true; host.textContent = r.time ? formatTime(r.time) : ""; }
    });
  }

  // ---------- model ops ----------
  async function removeReminder(r) {
    try {
      await api("DELETE", `/reminders/${r.id}`);
      reminders = reminders.filter((x) => x.id !== r.id);
      render();
    } catch (err) {
      if (err.status === 404) {
        reminders = reminders.filter((x) => x.id !== r.id);
        render();
        return;
      }
      showError("Could not delete", err);
    }
  }

  async function toggleDone(r) {
    const nextDone = !r.done;
    try {
      const updated = await api("PATCH", `/reminders/${r.id}`, { done: nextDone });
      replaceLocal(updated.reminder);
      render();
    } catch (err) {
      showError("Could not update", err);
    }
  }

  async function togglePriority(r) {
    const next = r.priority === "high" ? "normal" : "high";
    try {
      const updated = await api("PATCH", `/reminders/${r.id}`, { priority: next });
      replaceLocal(updated.reminder);
      render();
    } catch (err) {
      showError("Could not change priority", err);
    }
  }

  async function clearDoneVisible() {
    const targets = reminders.filter((r) => r.done && (settings.showAllDone || !r.closedAt || (Date.now() - new Date(r.closedAt).getTime()) < DONE_AGE_MS));
    try {
      await Promise.all(targets.map((r) => api("DELETE", `/reminders/${r.id}`).catch((e) => { if (e.status !== 404) throw e; })));
      const targetIds = new Set(targets.map((r) => r.id));
      reminders = reminders.filter((r) => !targetIds.has(r.id));
      render();
    } catch (err) {
      showError("Could not clear done", err);
    }
  }

  function replaceLocal(r) {
    const idx = reminders.findIndex((x) => x.id === r.id);
    if (idx >= 0) reminders[idx] = r;
    else reminders.push(r);
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

  // ---------- settings dialog ----------
  function openSettings() {
    $("setEodTime").value = settings.eodTime;
    $("setLeadMinutes").value = settings.leadMinutes;
    $("setWeekdaysOnly").checked = !!settings.weekdaysOnly;
    $("setNotifications").checked = settings.notifications !== false;
    if (typeof settingsDialog.showModal === "function") settingsDialog.showModal();
    else settingsDialog.setAttribute("open", "");
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
  function colorForTag(tag) {
    let h = 0;
    for (let i = 0; i < tag.length; i++) h = (h * 31 + tag.charCodeAt(i)) & 0xffffffff;
    return TAG_PALETTE[Math.abs(h) % TAG_PALETTE.length];
  }
  function formatTime(hhmm) {
    const [h, m] = hhmm.split(":").map(Number);
    const d = new Date();
    d.setHours(h, m, 0, 0);
    return d.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" });
  }
  function clampInt(v, min, max, fallback) {
    const n = parseInt(v, 10);
    if (isNaN(n)) return fallback;
    return Math.max(min, Math.min(max, n));
  }
  function showError(prefix, err) {
    console.error(prefix, err);
    const banner = $("errorBanner");
    if (banner) {
      banner.textContent = `${prefix}: ${err.message || err}`;
      banner.hidden = false;
      setTimeout(() => { banner.hidden = true; }, 5000);
    }
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

      const [{ settings: s, hasBot: hb }, { reminders: rems }] = await Promise.all([
        api("GET", "/settings"),
        api("GET", "/reminders"),
      ]);
      Object.assign(settings, s);
      hasBot = !!hb;
      reminders = rems;
      render();
    } catch (err) {
      console.error("Boot failed", err);
      const banner = $("errorBanner");
      if (banner) {
        banner.textContent = `Could not connect: ${err.message || err}`;
        banner.hidden = false;
      }
    }
  }
  function applyTheme(theme) {
    document.body.dataset.theme =
      theme === "dark" ? "dark" :
      theme === "contrast" ? "contrast" : "default";
  }

  boot();
})();
