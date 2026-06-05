/* Day Reminders — Teams personal tab
   Thin client over the bot's REST API. Auth via Teams SSO.
   All scheduling + notifications happen server-side (bot posts proactive
   Adaptive Cards into your personal chat with the Day Reminders bot).
*/
(function () {
  "use strict";

  const API_BASE = "https://func-day-reminders-17023.azurewebsites.net/api";

  // ---------- state ----------
  let reminders = [];
  let settings = {
    eodTime: "17:00",
    leadMinutes: 10,
    weekdaysOnly: true,
    notifications: true,
  };
  let hasBot = false;
  let authToken = null;

  // ---------- DOM ----------
  const $ = (id) => document.getElementById(id);
  const addForm = $("addForm");
  const titleInput = $("title");
  const timeInput = $("time");
  const todayList = $("todayList");
  const anytimeList = $("anytimeList");
  const doneList = $("doneList");
  const doneCard = $("doneCard");
  const todayEmpty = $("todayEmpty");
  const anytimeEmpty = $("anytimeEmpty");
  const todayDate = $("todayDate");
  const settingsDialog = $("settingsDialog");
  const settingsForm = $("settingsForm");
  const permHint = $("permHint");

  todayDate.textContent = new Date().toLocaleDateString(undefined, {
    weekday: "long", month: "short", day: "numeric",
  });

  // permHint is repurposed: warns if bot isn't installed yet.
  permHint.innerHTML = '';

  // ---------- events ----------
  addForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const title = titleInput.value.trim();
    if (!title) return;
    try {
      const created = await api("POST", "/reminders", { title, time: timeInput.value || null });
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
      settings = saved.settings;
      settingsDialog.close();
    } catch (err) {
      showError("Could not save settings", err);
    }
  });

  $("clearDone").addEventListener("click", async () => {
    const doneItems = reminders.filter((r) => r.done);
    try {
      await Promise.all(doneItems.map((r) => api("DELETE", `/reminders/${r.id}`)));
      reminders = reminders.filter((r) => !r.done);
      render();
    } catch (err) {
      showError("Could not clear done items", err);
    }
  });

  // ---------- render ----------
  function render() {
    const todayTimed = reminders.filter((r) => !r.done && r.time).sort((a, b) => a.time.localeCompare(b.time));
    const anytime = reminders.filter((r) => !r.done && !r.time);
    const done = reminders.filter((r) => r.done);

    renderList(todayList, todayTimed, true);
    renderList(anytimeList, anytime, false);
    renderList(doneList, done, true);

    todayEmpty.hidden = todayTimed.length > 0;
    anytimeEmpty.hidden = anytime.length > 0;
    doneCard.hidden = done.length === 0;
    updateBotHint();
  }

  function renderList(ul, items, showWhen) {
    ul.innerHTML = "";
    for (const r of items) {
      const li = document.createElement("li");
      if (r.firedAt) li.classList.add("fired");

      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.className = "checkbox";
      cb.checked = !!r.done;
      cb.setAttribute("aria-label", r.done ? "Mark not done" : "Mark done");
      cb.addEventListener("change", () => toggleDone(r));

      const title = document.createElement("span");
      title.className = "title";
      title.textContent = r.title;

      const when = document.createElement("span");
      when.className = "when";
      when.textContent = showWhen && r.time ? formatTime(r.time) : "";

      const del = document.createElement("button");
      del.className = "icon-btn";
      del.title = "Delete";
      del.setAttribute("aria-label", "Delete reminder");
      del.textContent = "✕";
      del.addEventListener("click", () => removeReminder(r));

      li.append(cb, title, when, del);
      ul.appendChild(li);
    }
  }

  function updateBotHint() {
    if (hasBot) {
      permHint.hidden = true;
      permHint.innerHTML = '';
      return;
    }
    permHint.hidden = false;
    permHint.innerHTML = 'Heads up: I haven\'t seen the Day Reminders bot in your chat yet. Open the app\'s chat once (left rail) so I can send your reminders there.';
  }

  // ---------- model ops ----------
  async function removeReminder(r) {
    try {
      await api("DELETE", `/reminders/${r.id}`);
      reminders = reminders.filter((x) => x.id !== r.id);
      render();
    } catch (err) {
      showError("Could not delete", err);
    }
  }
  async function toggleDone(r) {
    const nextDone = !r.done;
    try {
      const updated = await api("PATCH", `/reminders/${r.id}`, { done: nextDone });
      const idx = reminders.findIndex((x) => x.id === r.id);
      if (idx >= 0) reminders[idx] = updated.reminder;
      render();
    } catch (err) {
      showError("Could not update", err);
    }
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

  // ---------- utils ----------
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
      settings = s;
      hasBot = !!hb;
      reminders = rems;
      render();
    } catch (err) {
      console.error("Boot failed", err);
      const banner = $("errorBanner");
      if (banner) {
        const detail = err.message || String(err);
        banner.textContent = `Could not connect: ${detail}`;
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
