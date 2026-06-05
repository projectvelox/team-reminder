/* Day Reminders — Teams personal tab
   Storage: localStorage (per-user, per-Teams-client)
   Notifications: Browser Notification API, polled every 30s
*/
(function () {
  "use strict";

  const STORAGE_KEY = "dayReminders.v1";
  const SETTINGS_KEY = "dayReminders.settings.v1";
  const EOD_STATE_KEY = "dayReminders.eodState.v1";

  const DEFAULT_SETTINGS = {
    eodTime: "17:00",
    leadMinutes: 10,
    weekdaysOnly: true,
    notifications: true,
    sound: false,
  };

  // ---------- state ----------
  let reminders = loadReminders();
  let settings = loadSettings();
  let eodState = loadEodState(); // { date: "YYYY-MM-DD", shown: bool, snoozedUntil: ts|null, dismissed: bool }

  // ---------- Teams SDK ----------
  initTeams();

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
  const eodDialog = $("eodDialog");
  const permHint = $("permHint");

  todayDate.textContent = new Date().toLocaleDateString(undefined, {
    weekday: "long", month: "short", day: "numeric",
  });

  // ---------- events ----------
  addForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const title = titleInput.value.trim();
    if (!title) return;
    addReminder(title, timeInput.value || null);
    titleInput.value = "";
    timeInput.value = "";
    titleInput.focus();
    render();
  });

  $("openSettings").addEventListener("click", openSettings);
  $("settingsCancel").addEventListener("click", () => settingsDialog.close());
  settingsForm.addEventListener("submit", (e) => {
    e.preventDefault();
    settings = {
      eodTime: $("setEodTime").value || DEFAULT_SETTINGS.eodTime,
      leadMinutes: clampInt($("setLeadMinutes").value, 0, 240, 10),
      weekdaysOnly: $("setWeekdaysOnly").checked,
      notifications: $("setNotifications").checked,
      sound: $("setSound").checked,
    };
    saveSettings();
    if (settings.notifications) requestNotificationPermission();
    settingsDialog.close();
    updatePermHint();
  });

  $("clearDone").addEventListener("click", () => {
    reminders = reminders.filter((r) => !r.done);
    saveReminders();
    render();
  });

  $("enableNotif")?.addEventListener("click", requestNotificationPermission);

  $("eodSnooze").addEventListener("click", () => {
    eodState.snoozedUntil = Date.now() + 15 * 60 * 1000;
    eodState.dismissed = false;
    saveEodState();
    eodDialog.close();
  });
  $("eodDismiss").addEventListener("click", () => {
    eodState.dismissed = true;
    saveEodState();
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
      cb.addEventListener("change", () => toggleDone(r.id));

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
      del.addEventListener("click", () => removeReminder(r.id));

      li.append(cb, title, when, del);
      ul.appendChild(li);
    }
  }

  // ---------- model ops ----------
  function addReminder(title, time) {
    reminders.push({
      id: cryptoId(),
      title,
      time, // "HH:MM" or null
      createdDate: todayKey(),
      done: false,
      firedAt: null,
    });
    saveReminders();
  }
  function removeReminder(id) {
    reminders = reminders.filter((r) => r.id !== id);
    saveReminders();
    render();
  }
  function toggleDone(id) {
    const r = reminders.find((x) => x.id === id);
    if (!r) return;
    r.done = !r.done;
    saveReminders();
    render();
  }

  // ---------- persistence ----------
  function loadReminders() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return [];
      const arr = JSON.parse(raw);
      return Array.isArray(arr) ? arr : [];
    } catch { return []; }
  }
  function saveReminders() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(reminders));
  }
  function loadSettings() {
    try {
      const raw = localStorage.getItem(SETTINGS_KEY);
      if (!raw) return { ...DEFAULT_SETTINGS };
      return { ...DEFAULT_SETTINGS, ...JSON.parse(raw) };
    } catch { return { ...DEFAULT_SETTINGS }; }
  }
  function saveSettings() {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
  }
  function loadEodState() {
    try {
      const raw = localStorage.getItem(EOD_STATE_KEY);
      const obj = raw ? JSON.parse(raw) : null;
      if (!obj || obj.date !== todayKey()) {
        return { date: todayKey(), shown: false, snoozedUntil: null, dismissed: false };
      }
      return obj;
    } catch {
      return { date: todayKey(), shown: false, snoozedUntil: null, dismissed: false };
    }
  }
  function saveEodState() {
    localStorage.setItem(EOD_STATE_KEY, JSON.stringify(eodState));
  }

  // ---------- settings dialog ----------
  function openSettings() {
    $("setEodTime").value = settings.eodTime;
    $("setLeadMinutes").value = settings.leadMinutes;
    $("setWeekdaysOnly").checked = settings.weekdaysOnly;
    $("setNotifications").checked = settings.notifications;
    $("setSound").checked = settings.sound;
    if (typeof settingsDialog.showModal === "function") settingsDialog.showModal();
    else settingsDialog.setAttribute("open", "");
  }

  // ---------- notifications + scheduler ----------
  function requestNotificationPermission() {
    if (!("Notification" in window)) return;
    if (Notification.permission === "default") {
      Notification.requestPermission().then(updatePermHint);
    } else {
      updatePermHint();
    }
  }
  function updatePermHint() {
    const supported = "Notification" in window;
    const blocked = supported && Notification.permission === "denied";
    permHint.hidden = !(settings.notifications && (blocked || (supported && Notification.permission !== "granted")));
  }

  function notify(title, body) {
    if (!settings.notifications) return;
    if (!("Notification" in window) || Notification.permission !== "granted") return;
    try {
      const n = new Notification(title, { body, icon: "color.png", tag: "day-reminders" });
      n.onclick = () => { window.focus(); n.close(); };
    } catch { /* some Teams containers throw; ignore */ }
    if (settings.sound) playChime();
  }

  let audioCtx;
  function playChime() {
    try {
      audioCtx ||= new (window.AudioContext || window.webkitAudioContext)();
      const o = audioCtx.createOscillator();
      const g = audioCtx.createGain();
      o.type = "sine";
      o.frequency.value = 880;
      g.gain.value = 0.0001;
      o.connect(g).connect(audioCtx.destination);
      const t = audioCtx.currentTime;
      g.gain.exponentialRampToValueAtTime(0.15, t + 0.02);
      g.gain.exponentialRampToValueAtTime(0.0001, t + 0.5);
      o.start(t);
      o.stop(t + 0.55);
    } catch { /* ignore */ }
  }

  function tick() {
    // reset eod state at midnight rollover
    if (eodState.date !== todayKey()) {
      eodState = { date: todayKey(), shown: false, snoozedUntil: null, dismissed: false };
      saveEodState();
    }

    const now = new Date();

    // 1) per-reminder lead notifications
    for (const r of reminders) {
      if (r.done || !r.time || r.firedAt) continue;
      const target = parseTimeToday(r.time);
      const lead = settings.leadMinutes * 60 * 1000;
      const fireAt = target.getTime() - lead;
      if (now.getTime() >= fireAt && now.getTime() < target.getTime() + 60 * 1000) {
        notify("Reminder", settings.leadMinutes > 0
          ? `${r.title} (in ~${settings.leadMinutes} min, at ${formatTime(r.time)})`
          : `${r.title} (now)`);
        r.firedAt = Date.now();
        saveReminders();
        render();
      }
    }

    // 2) end-of-day check-in
    if (shouldShowEod(now)) {
      showEod();
    }
  }

  function shouldShowEod(now) {
    if (eodState.dismissed) return false;
    if (settings.weekdaysOnly) {
      const d = now.getDay();
      if (d === 0 || d === 6) return false;
    }
    const eod = parseTimeToday(settings.eodTime).getTime();
    if (eodState.snoozedUntil && now.getTime() < eodState.snoozedUntil) return false;
    if (eodState.snoozedUntil && now.getTime() >= eodState.snoozedUntil) return true;
    if (eodState.shown) return false;
    return now.getTime() >= eod;
  }

  function showEod() {
    const open = reminders.filter((r) => !r.done);
    const openTimed = open.filter((r) => r.time);
    const openAny = open.filter((r) => !r.time);

    $("eodSummary").textContent = open.length === 0
      ? "Nice — your list is clear. Wrap up?"
      : `You still have ${open.length} item${open.length === 1 ? "" : "s"} open. Anything carry over?`;

    const list = $("eodList");
    list.innerHTML = "";
    for (const r of [...openTimed, ...openAny]) {
      const li = document.createElement("li");
      const title = document.createElement("span");
      title.className = "title";
      title.textContent = r.title;
      const when = document.createElement("span");
      when.className = "when";
      when.textContent = r.time ? formatTime(r.time) : "anytime";
      li.append(document.createElement("span"), title, when, document.createElement("span"));
      list.appendChild(li);
    }

    notify("End of day", open.length === 0
      ? "Your list is clear. Wrap up?"
      : `${open.length} open item${open.length === 1 ? "" : "s"} — are you done?`);

    eodState.shown = true;
    eodState.snoozedUntil = null;
    saveEodState();

    if (!eodDialog.open) {
      if (typeof eodDialog.showModal === "function") eodDialog.showModal();
      else eodDialog.setAttribute("open", "");
    }
  }

  // ---------- utils ----------
  function parseTimeToday(hhmm) {
    const [h, m] = hhmm.split(":").map(Number);
    const d = new Date();
    d.setHours(h, m, 0, 0);
    return d;
  }
  function formatTime(hhmm) {
    const d = parseTimeToday(hhmm);
    return d.toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" });
  }
  function todayKey() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  }
  function clampInt(v, min, max, fallback) {
    const n = parseInt(v, 10);
    if (isNaN(n)) return fallback;
    return Math.max(min, Math.min(max, n));
  }
  function cryptoId() {
    if (window.crypto?.randomUUID) return crypto.randomUUID();
    return "id-" + Math.random().toString(36).slice(2) + Date.now().toString(36);
  }

  function initTeams() {
    if (!window.microsoftTeams) return;
    try {
      microsoftTeams.app.initialize().then(() => {
        microsoftTeams.app.getContext().then((ctx) => {
          applyTheme(ctx.app?.theme || "default");
        }).catch(() => {});
        microsoftTeams.app.registerOnThemeChangeHandler(applyTheme);
      }).catch(() => { /* not running in Teams — fine */ });
    } catch { /* SDK absent or failed to load — fine */ }
  }
  function applyTheme(theme) {
    document.body.dataset.theme =
      theme === "dark" ? "dark" :
      theme === "contrast" ? "contrast" : "default";
  }

  // ---------- boot ----------
  if (settings.notifications) requestNotificationPermission();
  updatePermHint();
  render();
  tick();
  setInterval(tick, 30 * 1000);
})();
