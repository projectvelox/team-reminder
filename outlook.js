// Day Reminders — Outlook taskpane.
//
// Loaded inside an Office Add-in taskpane in new Outlook for Windows / OWA.
// Pre-fills a reminder form from the current email's subject and sender,
// authenticates via Office SSO (same Entra app reg as the Teams tab), and
// POSTs to the existing /api/reminders endpoint. Defaults the date to
// tomorrow because email triage is usually a "deal with this later" gesture.

(function () {
  "use strict";

  const API_BASE = "https://func-day-reminders-17023.azurewebsites.net/api";

  const $ = (id) => document.getElementById(id);

  let authToken = null;

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) {
      showError("This add-in is for Outlook.");
      return;
    }
    applyTheme();
    Office.context.officeTheme && hookThemeChanges();
    populateFromEmail();
    wireForm();
    getSsoToken().then((token) => {
      authToken = token;
      preloadClientList();
    }).catch((err) => {
      showError("Could not sign in to Day Reminders. " + (err && err.message ? err.message : err));
    });
  });

  function applyTheme() {
    const theme = Office.context.officeTheme;
    if (!theme) return;
    const bg = theme.bodyBackgroundColor || "";
    const dark = isDarkBg(bg);
    document.body.setAttribute("data-theme", dark ? "dark" : "default");
  }
  function hookThemeChanges() {
    try {
      Office.context.mailbox.addHandlerAsync &&
        Office.context.mailbox.addHandlerAsync(
          Office.EventType.OfficeThemeChanged,
          applyTheme
        );
    } catch (_) { /* ignore — older clients */ }
  }
  function isDarkBg(hex) {
    if (!hex || !/^#?[0-9a-f]{6}$/i.test(hex.replace("#", ""))) return false;
    const h = hex.replace("#", "");
    const r = parseInt(h.slice(0, 2), 16);
    const g = parseInt(h.slice(2, 4), 16);
    const b = parseInt(h.slice(4, 6), 16);
    return (0.299 * r + 0.587 * g + 0.114 * b) < 128;
  }

  function populateFromEmail() {
    const item = Office.context.mailbox.item;
    const subject = (item && item.subject) ? String(item.subject).slice(0, 200) : "";
    const from = item && item.from ? item.from : null;
    const fromName = from && from.displayName ? from.displayName : "";
    const fromEmail = from && from.emailAddress ? from.emailAddress : "";

    $("outlookTitle").value = subject;
    $("outlookDate").value = tomorrowIso();

    const metaParts = [];
    if (fromName || fromEmail) {
      metaParts.push("From " + (fromName ? `${fromName}${fromEmail ? ` <${fromEmail}>` : ""}` : fromEmail));
    }
    $("outlookEmailMeta").textContent = metaParts.join(" · ");

    let details = "";
    if (subject) details += `Email: ${subject}\n`;
    if (fromName || fromEmail) details += `From: ${fromName}${fromEmail ? ` <${fromEmail}>` : ""}\n`;
    $("outlookDetails").value = details.trim();
  }

  function wireForm() {
    $("outlookTomorrow").addEventListener("click", () => {
      $("outlookDate").value = tomorrowIso();
    });
    $("outlookToday").addEventListener("click", () => {
      $("outlookDate").value = todayIso();
    });
    $("outlookCancel").addEventListener("click", () => {
      $("outlookForm").reset();
      populateFromEmail();
    });
    $("outlookForm").addEventListener("submit", onSubmit);
    $("outlookAddAnother").addEventListener("click", () => {
      $("outlookSuccess").hidden = true;
      $("outlookForm").hidden = false;
      populateFromEmail();
    });
    $("outlookDone").addEventListener("click", () => {
      // No clean "close pane" API across all hosts. The user closes the pane
      // from the host's pane controls. Reset the form so reopening is fresh.
      $("outlookSuccess").hidden = true;
      $("outlookForm").hidden = false;
      populateFromEmail();
    });
  }

  async function onSubmit(e) {
    e.preventDefault();
    if (!authToken) {
      showError("Not signed in yet. Wait a moment and try again.");
      return;
    }
    const title = $("outlookTitle").value.trim();
    if (!title) {
      showError("Title is required.");
      return;
    }
    const dueAt = $("outlookDate").value || tomorrowIso();
    const time = $("outlookTime").value || null;
    const client = $("outlookClient").value.trim().slice(0, 100) || null;
    const description = $("outlookDetails").value.trim().slice(0, 2000) || null;

    const tags = extractTags(title);
    const submitBtn = $("outlookSubmit");
    submitBtn.disabled = true;
    submitBtn.textContent = "Adding...";
    hideError();
    try {
      const body = { title, time, tags, dueAt };
      if (client) body.client = client;
      if (description) body.description = description;
      await api("POST", "/reminders", body);
      showSuccess(dueAt, time);
    } catch (err) {
      showError("Could not add reminder. " + (err && err.message ? err.message : err));
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = "Add reminder";
    }
  }

  function showSuccess(dueAt, time) {
    $("outlookForm").hidden = true;
    $("outlookSuccess").hidden = false;
    const today = todayIso();
    let when;
    if (dueAt === today) when = "today";
    else if (dueAt === tomorrowIso()) when = "tomorrow";
    else when = `on ${dueAt}`;
    $("outlookSuccessDetail").textContent = `Reminder set for ${when}${time ? ` at ${time}` : ""}. You'll get a chat ping from Day Reminders.`;
  }

  async function preloadClientList() {
    try {
      const data = await api("GET", "/reminders");
      const list = (data && data.reminders) || [];
      const set = new Set();
      for (const r of list) {
        if (r && r.client && typeof r.client === "string") {
          const v = r.client.trim();
          if (v) set.add(v);
        }
      }
      const dl = $("outlookClientList");
      dl.textContent = "";
      Array.from(set).sort((a, b) => a.localeCompare(b)).forEach((c) => {
        const opt = document.createElement("option");
        opt.value = c;
        dl.appendChild(opt);
      });
    } catch (_) { /* silent: autocomplete is nice-to-have */ }
  }

  // ---------- helpers ----------

  function getSsoToken() {
    return new Promise((resolve, reject) => {
      OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: false })
        .then(resolve)
        .catch(reject);
    });
  }

  async function api(method, path, body) {
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

  function extractTags(title) {
    const re = /(?:^|\s)#([a-z0-9_\-]{1,32})/gi;
    const tags = [];
    let m;
    while ((m = re.exec(title)) !== null) tags.push(m[1].toLowerCase());
    return tags;
  }

  // PH wall-clock dates so tab and add-in agree on "today" and "tomorrow".
  function phNow() {
    return new Date(Date.now() + 8 * 60 * 60 * 1000);
  }
  function isoFrom(d) {
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
  }
  function todayIso() {
    return isoFrom(phNow());
  }
  function tomorrowIso() {
    const d = phNow();
    d.setUTCDate(d.getUTCDate() + 1);
    return isoFrom(d);
  }

  function showError(msg) {
    const el = $("errorBanner");
    el.textContent = msg;
    el.hidden = false;
  }
  function hideError() {
    const el = $("errorBanner");
    el.textContent = "";
    el.hidden = true;
  }
})();
