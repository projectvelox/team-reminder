/* Day Reminders — Licenses tab (v1.7.0)
   Tenant-shared license tracker with Table + Calendar (Month) views.
   Auth via Teams SSO. Server is source of truth; we only cache for the session.
*/
(function () {
  "use strict";

  const API_BASE = "https://func-day-reminders-17023.azurewebsites.net/api";
  // v1.7.51 — version set in JS so a stale cached HTML still shows the
  // current build label (the JS itself is cache-busted via ?v=...).
  const LIC_VERSION = "v1.8.2";
  function paintVersionLabel() {
    const lbl = document.getElementById("licVersionLabel");
    if (lbl) lbl.textContent = LIC_VERSION;
  }
  if (document.readyState !== "loading") paintVersionLabel();
  else document.addEventListener("DOMContentLoaded", paintVersionLabel);
  const DEFAULT_LEAD_DAYS = 14;
  const STATUSES = ["notStarted", "noticeSent", "awaitingCustomer", "customerConfirmed", "renewed"];
  const STATUS_LABEL = {
    notStarted: "Not started",
    noticeSent: "Notice sent",
    awaitingCustomer: "Awaiting customer",
    customerConfirmed: "Customer confirmed",
    renewed: "Renewed",
  };
  const CYCLE_YEARS = { annual: 1, biennial: 2, triennial: 3 };
  const CYCLE_LABEL = { annual: "1/1", biennial: "Y?/2", triennial: "Y?/3" };
  // Owner pill palette, seeded differently from client-chip palette in styles.css so
  // a name reused as both client and owner doesn't get the same color in both contexts.
  const OWNER_PALETTE = [
    "#0078d4", "#107c10", "#8764b8", "#ca5010", "#c50f1f",
    "#038387", "#d83b01", "#5c2d91", "#0099bc", "#498205",
    "#3f6ec2", "#8a8886",
  ];

  // ---------- state ----------
  let licenses = [];
  let members = [];
  let customers = [];
  let emailTemplates = []; // [{ productLine, subject, body }]
  let me = { oid: null, name: null };
  let authToken = null;
  let teamsTheme = "default";
  let bulkMode = false;
  const bulkSelected = new Set();
  const pendingDeletes = new Map(); // licenseId -> { license, timer }

  // Default email templates seed the editor when no server-saved templates exist.
  // Variables: {customer}, {customerFirstWord}, {licenseType}, {users}, {expiryDate}, {ownerName}
  const DEFAULT_TEMPLATES = {
    _default: {
      subject: "Renewal reminder: {licenseType} expires {expiryDate}",
      body: "Hi {customerFirstWord} team,\n\nThis is a friendly reminder that your subscription for {licenseType} ({users} users) is set to expire on {expiryDate}.\n\nPlease let us know if you would like to:\n  1. Renew at the current {users} users\n  2. Adjust the seat count\n  3. Make any other changes\n\nLooking forward to your reply.\n\nBest regards,\n{ownerName}\nKation Technologies",
    },
    M365: {
      subject: "Microsoft 365 renewal: {customer} - {licenseType} ({expiryDate})",
      body: "Hi {customerFirstWord} team,\n\nYour Microsoft 365 subscription is coming up for renewal:\n  - {licenseType}: {users} users\n  - Expires: {expiryDate}\n\nPlease let us know if you would like to:\n  1. Renew at the current {users} seats\n  2. Adjust the seat count\n  3. Change plan\n\nLooking forward to your reply.\n\nBest regards,\n{ownerName}\nKation Technologies",
    },
    BC: {
      subject: "Dynamics 365 BC renewal: {customer} - {licenseType} ({expiryDate})",
      body: "Hi {customerFirstWord} team,\n\nYour Dynamics 365 Business Central subscription is coming up for renewal:\n  - {licenseType}: {users} users\n  - Expires: {expiryDate}\n\nPlease let us know if you would like to renew, adjust the seat count, or change plan.\n\nLooking forward to your reply.\n\nBest regards,\n{ownerName}\nKation Technologies",
    },
    BREP: {
      subject: "BREP renewal: {customer} - {licenseType} ({expiryDate})",
      body: "Hi {customerFirstWord} team,\n\nThis is a reminder for your Business Ready Enhancement Plan (BREP) renewal:\n  - {licenseType}\n  - Expires: {expiryDate}\n\nPlease confirm if you would like to renew the perpetual license enhancement plan for another term.\n\nBest regards,\n{ownerName}\nKation Technologies",
    },
  };

  let currentView = "table";
  let summaryFilter = null; // null | 'week' | 'month' | 'overdue'
  let quickFilter = "all";  // 'all' | 'mine' | 'month' | 'overdue' | 'attention'
  // v1.7.20: breakdown chips are multi-select. Empty set = no filter on that
  // axis; otherwise only rows whose value is in the set pass.
  let ownerFilter = new Set();   // Set<ownerOid>
  let productFilter = new Set(); // Set<productLine>
  let statusFilter = new Set();  // Set<status>  -- workflow status (notStarted/...)
  let expiryFilter = new Set();  // Set<'expired'|'soon'|'thisMonth'|'active'>  -- v1.7.38 date-derived
  let monthFilter = "";          // YYYY-MM, applies to expiryDate; "" = any month
  let quarterFilter = "";        // YYYY-QN (e.g. "2026-Q3"); "" = any quarter (v1.8.0)
  let dateFromFilter = "";       // YYYY-MM-DD inclusive (v1.7.42)
  let dateToFilter = "";         // YYYY-MM-DD inclusive (v1.7.42)
  let groupBy = "none";     // 'none' | 'customer' | 'ownerName' | 'productLine'
  let searchText = "";
  let sortKey = "expiryDate";
  let sortDir = 1; // 1 asc, -1 desc

  // v1.7.38 expiry buckets (date-derived "health"). The "soon" threshold tracks
  // the user's smallest lead-day default so the visual bucket agrees with what's
  // actually being notified about — instead of hardcoding 10 days like Ella's
  // spec, which would disagree with a user whose lead-days are [60,30,15,7,1].
  const EXPIRY_BUCKETS = ["expired", "soon", "thisMonth", "active"];
  const EXPIRY_LABEL = {
    expired: "Expired",
    soon: "Expiring soon",
    thisMonth: "Expiring this month",
    active: "Active",
  };
  function soonThresholdDays() {
    const arr = Array.isArray(userSettings.licenseLeadDays) && userSettings.licenseLeadDays.length
      ? userSettings.licenseLeadDays.slice().sort((a, b) => a - b)
      : [7];
    return arr[0];
  }
  function expiryBucket(daysLeft) {
    if (daysLeft === null) return null;
    if (daysLeft < 0) return "expired";
    if (daysLeft <= soonThresholdDays()) return "soon";
    if (daysLeft <= 30) return "thisMonth";
    return "active";
  }

  let calCursor = startOfMonth(new Date());
  // v1.8.0 — Quarter view cursor (year + 1-4 index). Anchors which 3-month
  // grid Quarter view renders. Defaults to the quarter that contains today.
  let quarterCursor = (function () {
    const d = new Date();
    return { year: d.getFullYear(), q: Math.floor(d.getMonth() / 3) + 1 };
  })();
  // v1.8.0 — controlled-vocab registry for productLine. Empty until first
  // fetch completes; row renderer guards against the empty case so the
  // "legacy" flag doesn't false-positive during boot.
  let productLinesRegistry = []; // [{name, sortOrder}]
  let editingId = null; // id of license being edited in licDialog; null = new
  let renewTargetId = null;
  let importRows = []; // staged rows during CSV import
  let licOwnerPicker = null;
  let bulkReassignPicker = null;
  let userSettings = {
    licenseLeadDays: [14],
    licenseSkipBriefing: false,
    licenseSkipMonthlyDigest: false,
    licenseRollupDigest: false,
    savedLicenseViews: [],
  };
  let themeOverride = "auto"; // 'auto' | 'default' | 'dark' | 'contrast'
  try {
    const saved = localStorage.getItem("lic.themeOverride");
    if (saved && ["auto", "default", "dark", "contrast"].includes(saved)) themeOverride = saved;
  } catch (_) {}
  // Collapse state for grouped table view. Keys are "axis:value" e.g.
  // "ownerName:Joshua Oducado", so switching group axes preserves collapse state
  // for each axis independently.
  const collapsedGroups = new Set();
  try {
    const saved = localStorage.getItem("lic.collapsedGroups");
    if (saved) JSON.parse(saved).forEach((k) => collapsedGroups.add(k));
  } catch (_) {}
  function saveCollapsedGroups() {
    try { localStorage.setItem("lic.collapsedGroups", JSON.stringify([...collapsedGroups])); } catch (_) {}
  }
  // Precomputed once per render so every row in the same cluster shows the
  // same Bundle: N badge. Recomputed by computeBundles().
  let licenseBundles = new Map(); // licenseId -> { id, members[], size }

  // localStorage keys (tab-only UI state)
  const LS_VIEW = "lic.view";
  const LS_QUICK = "lic.quickFilter";
  const LS_SORT = "lic.sort";
  const LS_GROUP = "lic.groupBy";
  // v1.7.38 — multi-axis filter persistence so a refresh keeps Ella's drilldown.
  const LS_OWNER_FILTER = "lic.ownerFilter";
  const LS_PRODUCT_FILTER = "lic.productFilter";
  const LS_STATUS_FILTER = "lic.statusFilter";
  const LS_EXPIRY_FILTER = "lic.expiryFilter";
  const LS_MONTH_FILTER = "lic.monthFilter";
  const LS_QUARTER_FILTER = "lic.quarterFilter"; // v1.8.0

  try {
    const v = localStorage.getItem(LS_VIEW);
    if (v === "table" || v === "calendar" || v === "quarter") currentView = v;
    const q = localStorage.getItem(LS_QUICK);
    if (["all", "mine", "month", "quarter", "overdue", "attention"].includes(q)) quickFilter = q;
    const s = localStorage.getItem(LS_SORT);
    if (s) {
      const [k, d] = s.split(":");
      if (k) sortKey = k;
      if (d === "1" || d === "-1") sortDir = Number(d);
    }
    const g = localStorage.getItem(LS_GROUP);
    if (["none", "customer", "ownerName", "productLine"].includes(g)) groupBy = g;
    const ownerJson = localStorage.getItem(LS_OWNER_FILTER);
    if (ownerJson) JSON.parse(ownerJson).forEach((x) => ownerFilter.add(x));
    const prodJson = localStorage.getItem(LS_PRODUCT_FILTER);
    if (prodJson) JSON.parse(prodJson).forEach((x) => productFilter.add(x));
    const statusJson = localStorage.getItem(LS_STATUS_FILTER);
    if (statusJson) JSON.parse(statusJson).forEach((x) => statusFilter.add(x));
    const expiryJson = localStorage.getItem(LS_EXPIRY_FILTER);
    if (expiryJson) JSON.parse(expiryJson).forEach((x) => { if (EXPIRY_BUCKETS.includes(x)) expiryFilter.add(x); });
    const mFilter = localStorage.getItem(LS_MONTH_FILTER);
    if (mFilter && /^\d{4}-\d{2}$/.test(mFilter)) monthFilter = mFilter;
    const qFilter = localStorage.getItem(LS_QUARTER_FILTER);
    if (qFilter && /^\d{4}-Q[1-4]$/.test(qFilter)) quarterFilter = qFilter;
  } catch (_) {}
  function persistFilters() {
    try {
      localStorage.setItem(LS_OWNER_FILTER, JSON.stringify([...ownerFilter]));
      localStorage.setItem(LS_PRODUCT_FILTER, JSON.stringify([...productFilter]));
      localStorage.setItem(LS_STATUS_FILTER, JSON.stringify([...statusFilter]));
      localStorage.setItem(LS_EXPIRY_FILTER, JSON.stringify([...expiryFilter]));
      localStorage.setItem(LS_MONTH_FILTER, monthFilter || "");
      localStorage.setItem(LS_QUARTER_FILTER, quarterFilter || "");
    } catch (_) {}
    syncFiltersToHash();
  }

  // v1.7.42 — URL hash state so filtered views are shareable links. Teams
  // strips/rewrites query strings on tab loads, so we use the hash fragment
  // (preserved across reloads inside the iframe). Encoded as URL-safe params.
  let suppressHashSync = false;
  function encodeFilterHash() {
    const params = new URLSearchParams();
    if (quickFilter && quickFilter !== "all") params.set("q", quickFilter);
    if (summaryFilter) params.set("s", summaryFilter);
    if (ownerFilter.size) params.set("o", [...ownerFilter].join(","));
    if (productFilter.size) params.set("p", [...productFilter].join("|"));
    if (statusFilter.size) params.set("st", [...statusFilter].join(","));
    if (expiryFilter.size) params.set("x", [...expiryFilter].join(","));
    if (monthFilter) params.set("m", monthFilter);
    if (quarterFilter) params.set("qf", quarterFilter); // v1.8.0
    if (dateFromFilter) params.set("df", dateFromFilter);
    if (dateToFilter) params.set("dt", dateToFilter);
    if (searchText) params.set("q2", searchText);
    if (currentView && currentView !== "table") params.set("v", currentView);
    const s = params.toString();
    return s ? "#" + s : "";
  }
  function syncFiltersToHash() {
    if (suppressHashSync) return;
    const next = encodeFilterHash();
    // Avoid clobbering identical state (history noise) and avoid pushing
    // entries — replaceState keeps the back-button useful for actual nav.
    if (("#" + (location.hash.slice(1) || "")) === next || next === "" && location.hash === "") return;
    try { history.replaceState(null, "", location.pathname + location.search + next); } catch (_) {}
  }
  function loadFiltersFromHash() {
    const raw = location.hash.replace(/^#/, "");
    if (!raw) return false;
    const params = new URLSearchParams(raw);
    let touched = false;
    if (params.has("q")) { quickFilter = params.get("q"); touched = true; }
    if (params.has("s")) { summaryFilter = params.get("s") || null; touched = true; }
    if (params.has("o")) { ownerFilter = new Set(params.get("o").split(",").filter(Boolean)); touched = true; }
    if (params.has("p")) { productFilter = new Set(params.get("p").split("|").filter(Boolean)); touched = true; }
    if (params.has("st")) { statusFilter = new Set(params.get("st").split(",").filter(Boolean)); touched = true; }
    if (params.has("x")) { expiryFilter = new Set(params.get("x").split(",").filter(Boolean).filter((b) => EXPIRY_BUCKETS.includes(b))); touched = true; }
    if (params.has("m")) { monthFilter = params.get("m"); touched = true; }
    if (params.has("qf")) {
      const qf = params.get("qf");
      if (/^\d{4}-Q[1-4]$/.test(qf)) { quarterFilter = qf; touched = true; }
    }
    if (params.has("df")) { dateFromFilter = params.get("df"); touched = true; }
    if (params.has("dt")) { dateToFilter = params.get("dt"); touched = true; }
    if (params.has("q2")) { searchText = params.get("q2"); touched = true; }
    if (params.has("v")) { const v = params.get("v"); if (v === "calendar" || v === "table" || v === "quarter") currentView = v; touched = true; }
    return touched;
  }

  // ---------- date helpers ----------
  function todayPh() {
    const ph = new Date(Date.now() + 8 * 60 * 60 * 1000);
    return `${ph.getUTCFullYear()}-${String(ph.getUTCMonth() + 1).padStart(2, "0")}-${String(ph.getUTCDate()).padStart(2, "0")}`;
  }
  function parseISO(d) {
    if (!d || typeof d !== "string") return null;
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d);
    if (!m) return null;
    return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  }
  function daysBetween(aStr, bStr) {
    const a = parseISO(aStr);
    const b = parseISO(bStr);
    if (!a || !b) return null;
    return Math.round((b - a) / 86400000);
  }
  function startOfMonth(d) {
    return new Date(d.getFullYear(), d.getMonth(), 1);
  }
  function fmtMonth(d) {
    return d.toLocaleString(undefined, { month: "long", year: "numeric" });
  }
  function fmtShortDate(d) {
    if (!d) return "";
    const parsed = typeof d === "string" ? parseISO(d) : d;
    if (!parsed) return "";
    return parsed.toLocaleString(undefined, { month: "short", day: "numeric", year: "numeric", timeZone: "UTC" });
  }

  // ---------- color hashing ----------
  function ownerColor(oid) {
    if (!oid) return "#888";
    let h = 0;
    for (let i = 0; i < oid.length; i++) h = ((h << 5) - h + oid.charCodeAt(i)) | 0;
    return OWNER_PALETTE[Math.abs(h) % OWNER_PALETTE.length];
  }

  // ---------- DOM helpers ----------
  const $ = (id) => document.getElementById(id);
  function showError(msg, err) {
    const banner = $("errorBanner");
    if (banner) {
      banner.textContent = err ? `${msg}: ${err.message || err}` : msg;
      banner.hidden = false;
    }
    if (err) console.error(msg, err);
  }
  // v1.7.39 — toast accepts an optional { actionLabel, onAction, durationMs }.
  // When actionLabel is set, the toast renders a clickable button (e.g. Undo)
  // and stays up longer. Calling toast() with a plain string keeps the old API.
  function toast(msg, opts) {
    const t = $("toast");
    if (!t) return;
    t.innerHTML = "";
    const span = document.createElement("span");
    span.textContent = msg;
    t.appendChild(span);
    // v1.7.50 — actionable toasts (Undo) get 8s window so users can spot them.
    const durationMs = (opts && opts.durationMs) || (opts && opts.actionLabel ? 8000 : 2400);
    if (opts && opts.actionLabel && typeof opts.onAction === "function") {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "toast-action";
      btn.textContent = opts.actionLabel;
      btn.addEventListener("click", () => {
        clearTimeout(toast._t);
        t.hidden = true;
        try { opts.onAction(); } catch (err) { console.error(err); }
      });
      t.appendChild(btn);
    }
    t.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(() => { t.hidden = true; }, durationMs);
  }

  // ---------- API ----------
  // v1.7.41 — `opts.ifMatch` adds the optimistic-concurrency If-Match header.
  // On 409 we throw an error tagged with .status=409 and .license=<currentRow>
  // so callers can show a "Rey just edited this — reload?" toast instead of
  // silently clobbering somebody else's edit.
  async function api(method, path, body, opts) {
    const headers = { "Authorization": `Bearer ${authToken}` };
    const init = { method, headers };
    if (body !== undefined) {
      headers["Content-Type"] = "application/json";
      init.body = JSON.stringify(body);
    }
    if (opts && opts.ifMatch) headers["If-Match"] = opts.ifMatch;
    const res = await fetch(API_BASE + path, init);
    if (res.status === 204) return null;
    let data = null;
    try { data = await res.json(); } catch (_) {}
    if (!res.ok) {
      const msg = data && data.error ? data.error : `HTTP ${res.status}`;
      const err = new Error(msg);
      err.status = res.status;
      if (res.status === 409 && data && data.license) err.license = data.license;
      throw err;
    }
    return data;
  }
  // Shared 409 handler: refresh from server + show a toast offering reload.
  // Returns true if the error was a conflict (caller should stop).
  function handleConflict(err, label) {
    if (!err || err.status !== 409) return false;
    if (err.license) {
      licenses = licenses.map((l) => l.id === err.license.id ? err.license : l);
      render();
    }
    const editedBy = err.license && err.license.lastEditedByName ? err.license.lastEditedByName : "someone";
    toast(`${label || "Edit"} blocked: ${editedBy} changed this row first.`, {
      actionLabel: "Reload",
      onAction: () => refreshAll(),
    });
    return true;
  }

  // ---------- filtering / sorting ----------
  function effectiveFilter() {
    // summaryFilter takes precedence over quickFilter when set.
    return summaryFilter || quickFilter;
  }
  function matchesFilter(lic) {
    const today = todayPh();
    const d = daysBetween(today, lic.expiryDate);
    const f = effectiveFilter();
    // Quick / summary filter
    let ok = true;
    if (f === "mine") ok = lic.ownerOid === me.oid;
    else if (f === "month") {
      if (!lic.expiryDate) ok = false;
      else {
        const exp = parseISO(lic.expiryDate);
        const now = new Date();
        ok = exp && exp.getUTCFullYear() === now.getFullYear() && exp.getUTCMonth() === now.getMonth();
      }
    }
    else if (f === "quarter") {
      // v1.8.0 — calendar quarter (Q1=Jan-Mar, etc.) containing today.
      if (!lic.expiryDate) ok = false;
      else {
        const exp = parseISO(lic.expiryDate);
        const now = new Date();
        const nowQ = Math.floor(now.getMonth() / 3);
        const expQ = exp ? Math.floor(exp.getUTCMonth() / 3) : -1;
        ok = exp && exp.getUTCFullYear() === now.getFullYear() && expQ === nowQ;
      }
    }
    else if (f === "overdue") ok = d !== null && d < 0 && lic.state !== "abandoned";
    else if (f === "week") ok = d !== null && d >= 0 && d <= 7;
    else if (f === "attention") ok = needsAttention(lic, today);
    // additive owner / product / status breakdown filters (multi-select)
    if (ok && ownerFilter.size > 0 && !ownerFilter.has(lic.ownerOid)) ok = false;
    if (ok && productFilter.size > 0 && !productFilter.has(lic.productLine)) ok = false;
    if (ok && statusFilter.size > 0 && !statusFilter.has(lic.status || "notStarted")) ok = false;
    // v1.7.38 expiry-bucket multi-select
    if (ok && expiryFilter.size > 0) {
      const bucket = expiryBucket(d);
      if (!bucket || !expiryFilter.has(bucket)) ok = false;
    }
    // v1.7.38 month filter (YYYY-MM)
    if (ok && monthFilter) {
      if (!lic.expiryDate || lic.expiryDate.slice(0, 7) !== monthFilter) ok = false;
    }
    // v1.8.0 quarter filter (YYYY-QN)
    if (ok && quarterFilter) {
      const m = quarterFilter.match(/^(\d{4})-Q([1-4])$/);
      if (!m || !lic.expiryDate) ok = false;
      else {
        const qYear = Number(m[1]);
        const qIdx = Number(m[2]) - 1; // 0..3
        const expYear = Number(lic.expiryDate.slice(0, 4));
        const expMonth = Number(lic.expiryDate.slice(5, 7)) - 1;
        const expQ = Math.floor(expMonth / 3);
        if (expYear !== qYear || expQ !== qIdx) ok = false;
      }
    }
    // v1.7.42 date-range filter (inclusive). Both ends optional.
    if (ok && dateFromFilter) {
      if (!lic.expiryDate || lic.expiryDate < dateFromFilter) ok = false;
    }
    if (ok && dateToFilter) {
      if (!lic.expiryDate || lic.expiryDate > dateToFilter) ok = false;
    }
    return ok;
  }

  // A row "needs attention" if:
  //   - it's stuck in noticeSent or awaitingCustomer for > 7 days since last status change, OR
  //   - status is notStarted AND expiry is within 30 days, OR
  //   - status is customerConfirmed AND expiry is in the past (i.e. confirmed but never marked renewed).
  function needsAttention(lic, today) {
    if (lic.state === "abandoned") return false;
    const d = daysBetween(today, lic.expiryDate);
    if (lic.status === "noticeSent" || lic.status === "awaitingCustomer") {
      if (!lic.statusChangedAt) return false;
      const days = Math.floor((Date.now() - Date.parse(lic.statusChangedAt)) / 86400000);
      return days >= 7;
    }
    if (lic.status === "notStarted") return d !== null && d <= 30;
    if (lic.status === "customerConfirmed") return d !== null && d < 0;
    return false;
  }
  function matchesSearch(lic) {
    if (!searchText) return true;
    const q = searchText.toLowerCase();
    return (lic.customer || "").toLowerCase().includes(q)
        || (lic.licenseType || "").toLowerCase().includes(q)
        || (lic.notes || "").toLowerCase().includes(q)
        || (lic.productLine || "").toLowerCase().includes(q)
        || (lic.ownerName || "").toLowerCase().includes(q);
  }
  function visibleLicenses() {
    return licenses.filter((l) => matchesFilter(l) && matchesSearch(l));
  }
  function sortLicenses(list) {
    const k = sortKey;
    const dir = sortDir;
    return [...list].sort((a, b) => {
      if (k === "status") {
        return (STATUSES.indexOf(a.status || "notStarted") - STATUSES.indexOf(b.status || "notStarted")) * dir;
      }
      const va = a[k] == null ? "" : a[k];
      const vb = b[k] == null ? "" : b[k];
      if (typeof va === "number" && typeof vb === "number") return (va - vb) * dir;
      return String(va).localeCompare(String(vb)) * dir;
    });
  }

  // ---------- summary chips ----------
  function recomputeSummary() {
    const today = todayPh();
    let w = 0, m = 0, o = 0;
    const now = new Date();
    for (const lic of licenses) {
      if (lic.state === "abandoned") continue;
      const d = daysBetween(today, lic.expiryDate);
      if (d === null) continue;
      if (d < 0) o++;
      if (d >= 0 && d <= 7) w++;
      const exp = parseISO(lic.expiryDate);
      if (exp && exp.getUTCFullYear() === now.getFullYear() && exp.getUTCMonth() === now.getMonth()) m++;
    }
    $("cntWeek").textContent = w;
    $("cntMonth").textContent = m;
    $("cntOverdue").textContent = o;
    document.querySelectorAll(".lic-summary-chip").forEach((b) => {
      b.setAttribute("aria-pressed", summaryFilter === b.dataset.summary ? "true" : "false");
    });
  }

  // ---------- v1.7.38 active-filter bar ----------
  // v1.7.39 — owner avatar URL backed by /api/users/{oid}/photo (60s server
  // cache). The img tag's onerror falls back to the colored initials pill.
  function ownerAvatarUrl(oid) {
    if (!oid) return null;
    return `${API_BASE}/users/${encodeURIComponent(oid)}/photo`;
  }
  function ownerInitials(name) {
    if (!name) return "?";
    const parts = name.trim().split(/\s+/).filter(Boolean);
    if (!parts.length) return "?";
    if (parts.length === 1) return parts[0][0].toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }
  function buildOwnerAvatar(lic, opts) {
    const size = (opts && opts.size) || 22;
    const wrap = document.createElement("span");
    wrap.className = "owner-avatar";
    wrap.style.width = `${size}px`;
    wrap.style.height = `${size}px`;
    wrap.style.background = ownerColor(lic.ownerOid);
    wrap.style.fontSize = `${Math.max(9, size * 0.42)}px`;
    wrap.title = lic.ownerName || (lic.ownerOid ? lic.ownerOid.slice(0, 8) : "(none)");
    const initials = document.createElement("span");
    initials.className = "owner-avatar-initials";
    initials.textContent = ownerInitials(lic.ownerName);
    wrap.appendChild(initials);
    // v1.7.51 — probe the photo off-DOM. Only attach the <img> on a real
    // successful load with non-zero dimensions, so the browser's broken-image
    // glyph never paints over the initials when the photo proxy returns 204
    // (no photo) or any non-image payload.
    if (lic.ownerOid) {
      const probe = new Image();
      probe.onload = () => {
        if (!probe.naturalWidth || !probe.naturalHeight) return;
        const img = document.createElement("img");
        img.src = probe.src;
        img.alt = "";
        wrap.appendChild(img);
      };
      probe.onerror = () => {};
      probe.src = ownerAvatarUrl(lic.ownerOid);
    }
    return wrap;
  }

  // v1.7.40 — auto-add a corner close X to every <dialog> so users don't have
  // to scroll to find a Cancel/Close button or remember Esc. Positioned via
  // CSS (top-right of the dialog). Skipped if a dialog already has its own X
  // (so this is idempotent and safe to call after dynamic inserts).
  function installDialogCloseButtons() {
    document.querySelectorAll("dialog").forEach((d) => {
      if (d.querySelector(":scope > .dialog-close-x")) return;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "dialog-close-x";
      btn.setAttribute("aria-label", "Close");
      btn.title = "Close (Esc)";
      btn.innerHTML = "&times;";
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        if (d.open) d.close();
      });
      d.prepend(btn);
    });
  }

  // v1.7.44 — customer hover preview. Attaches a 250ms-delayed enter handler
  // to any element; opens a fixed-positioned card with summary stats and
  // cleans up on leave. Reuses the in-memory `licenses` array — no API call.
  let _hoverCardEl = null;
  let _hoverTimer = null;
  function attachCustomerHoverPreview(el, customer) {
    if (!el || !customer) return;
    el.addEventListener("mouseenter", () => {
      _hoverTimer = setTimeout(() => showCustomerHoverCard(el, customer), 250);
    });
    el.addEventListener("mouseleave", () => {
      clearTimeout(_hoverTimer);
      if (_hoverCardEl) { _hoverCardEl.remove(); _hoverCardEl = null; }
    });
  }
  function showCustomerHoverCard(anchorEl, customer) {
    const norm = customer.trim().toLowerCase();
    const rows = licenses.filter((l) => (l.customer || "").trim().toLowerCase() === norm);
    if (!rows.length) return;
    const seats = rows.reduce((s, l) => s + (l.userCount || 0), 0);
    const owners = new Set();
    let nextExpiry = null;
    let overdue = 0;
    const today = todayPh();
    for (const l of rows) {
      if (l.ownerName) owners.add(l.ownerName);
      if (l.expiryDate) {
        if (l.expiryDate < today && l.state !== "abandoned" && l.status !== "renewed") overdue++;
        if (l.expiryDate >= today && (!nextExpiry || l.expiryDate < nextExpiry)) nextExpiry = l.expiryDate;
      }
    }
    const card = document.createElement("div");
    card.className = "customer-hover-card";
    card.innerHTML = `
      <h4>${escapeHtml(customer)}</h4>
      <dl>
        <dt>Licenses</dt><dd>${rows.length}</dd>
        <dt>Seats</dt><dd>${seats.toLocaleString()}</dd>
        <dt>Owners</dt><dd>${escapeHtml([...owners].join(", ") || "—")}</dd>
        <dt>Next expiry</dt><dd>${nextExpiry ? escapeHtml(fmtShortDate(nextExpiry)) : "—"}</dd>
        ${overdue ? `<dt>Overdue</dt><dd style="color: var(--color-danger-fg, #b71c1c);">${overdue}</dd>` : ""}
      </dl>
    `;
    document.body.appendChild(card);
    // Position the card next to the anchor without going off-screen.
    const ar = anchorEl.getBoundingClientRect();
    const cardRect = card.getBoundingClientRect();
    let left = ar.right + 8;
    if (left + cardRect.width > window.innerWidth - 12) left = ar.left - cardRect.width - 8;
    if (left < 8) left = 8;
    let top = ar.top - 4;
    if (top + cardRect.height > window.innerHeight - 12) top = window.innerHeight - cardRect.height - 12;
    card.style.left = `${left}px`;
    card.style.top = `${top}px`;
    _hoverCardEl = card;
  }

  function ownerNameByOid(oid) {
    if (!oid) return "(none)";
    // Prefer the in-memory members list (auto-registered), fall back to the
    // license rows where the owner was last seen, then truncated oid.
    const m = members.find((x) => x.oid === oid);
    if (m && m.displayName) return m.displayName;
    const lic = licenses.find((l) => l.ownerOid === oid && l.ownerName);
    return lic ? lic.ownerName : oid.slice(0, 8);
  }
  function fmtMonthShort(yyyyMm) {
    const [y, m] = yyyyMm.split("-").map(Number);
    return new Date(y, m - 1, 1).toLocaleString(undefined, { month: "short", year: "numeric" });
  }
  function hasAnyFilter() {
    return (
      quickFilter !== "all" ||
      summaryFilter !== null ||
      ownerFilter.size > 0 ||
      productFilter.size > 0 ||
      statusFilter.size > 0 ||
      expiryFilter.size > 0 ||
      !!monthFilter ||
      !!quarterFilter ||
      !!dateFromFilter ||
      !!dateToFilter ||
      !!searchText
    );
  }
  function fmtQuarterLabel(yyyyQn) {
    const m = (yyyyQn || "").match(/^(\d{4})-Q([1-4])$/);
    if (!m) return yyyyQn;
    return `Q${m[2]} ${m[1]}`;
  }
  function renderActiveFilterBar() {
    const visibleCount = visibleLicenses().length;
    const totalCount = licenses.length;
    const countEl = $("resultCount");
    if (!hasAnyFilter()) {
      countEl.textContent = `${totalCount} license${totalCount === 1 ? "" : "s"}`;
    } else {
      countEl.textContent = `${visibleCount} of ${totalCount}`;
    }
    const chipsEl = $("activeFilterChips");
    chipsEl.innerHTML = "";

    function addChip(label, onClear) {
      const chip = document.createElement("span");
      chip.className = "active-filter-chip";
      const lab = document.createElement("span");
      lab.textContent = label;
      const x = document.createElement("button");
      x.type = "button";
      x.className = "active-filter-chip-x";
      x.setAttribute("aria-label", `Remove filter: ${label}`);
      x.textContent = "×";
      x.addEventListener("click", onClear);
      chip.appendChild(lab);
      chip.appendChild(x);
      chipsEl.appendChild(chip);
    }

    if (quickFilter && quickFilter !== "all") {
      const QUICK_LABEL = { mine: "Mine", month: "This month", quarter: "This quarter", overdue: "Overdue", attention: "Needs attention" };
      addChip(QUICK_LABEL[quickFilter] || quickFilter, () => { setQuickFilter("all"); });
    }
    if (summaryFilter) {
      const SUM_LABEL = { week: "Expiring this week", month: "Expiring this month", overdue: "Overdue" };
      addChip(SUM_LABEL[summaryFilter] || summaryFilter, () => { summaryFilter = null; render(); });
    }
    for (const oid of ownerFilter) {
      addChip(`Owner: ${ownerNameByOid(oid)}`, () => { ownerFilter.delete(oid); persistFilters(); render(); });
    }
    for (const p of productFilter) {
      addChip(`Product: ${p || "(none)"}`, () => { productFilter.delete(p); persistFilters(); render(); });
    }
    for (const s of statusFilter) {
      addChip(`Status: ${STATUS_LABEL[s] || s}`, () => { statusFilter.delete(s); persistFilters(); render(); });
    }
    for (const b of expiryFilter) {
      addChip(`Expiry: ${EXPIRY_LABEL[b] || b}`, () => { expiryFilter.delete(b); persistFilters(); renderExpiryFilterMenu(); render(); });
    }
    if (monthFilter) {
      addChip(`Month: ${fmtMonthShort(monthFilter)}`, () => { monthFilter = ""; persistFilters(); $("monthFilter").value = ""; render(); });
    }
    if (quarterFilter) {
      addChip(`Quarter: ${fmtQuarterLabel(quarterFilter)}`, () => { quarterFilter = ""; persistFilters(); $("quarterFilter").value = ""; render(); });
    }
    if (dateFromFilter || dateToFilter) {
      const lbl = dateFromFilter && dateToFilter
        ? `Expires ${fmtShortDate(dateFromFilter)} – ${fmtShortDate(dateToFilter)}`
        : dateFromFilter ? `Expires from ${fmtShortDate(dateFromFilter)}` : `Expires until ${fmtShortDate(dateToFilter)}`;
      addChip(lbl, () => {
        dateFromFilter = ""; dateToFilter = "";
        if ($("dateFromFilter")) $("dateFromFilter").value = "";
        if ($("dateToFilter")) $("dateToFilter").value = "";
        persistFilters(); render();
      });
    }
    if (searchText) {
      addChip(`Search: "${searchText}"`, () => { searchText = ""; $("searchInput").value = ""; $("searchClear").hidden = true; render(); });
    }

    $("clearAllFiltersBtn").hidden = !hasAnyFilter();
  }
  // ---------- v1.7.39 calendar density ----------
  let calDensity = "comfortable"; // 'comfortable' | 'compact'
  try {
    const d = localStorage.getItem("lic.calDensity");
    if (d === "compact" || d === "comfortable") calDensity = d;
  } catch (_) {}
  function setCalDensity(d) {
    calDensity = d;
    try { localStorage.setItem("lic.calDensity", d); } catch (_) {}
    document.body.dataset.calDensity = d;
    document.querySelectorAll("#calDensity .view-btn").forEach((b) => {
      b.setAttribute("aria-pressed", b.dataset.density === d ? "true" : "false");
    });
  }
  function populateCalYearJump() {
    const sel = $("calYearJump");
    if (!sel) return;
    const years = new Set();
    const thisYear = new Date().getFullYear();
    years.add(thisYear);
    years.add(thisYear + 1);
    years.add(thisYear + 2);
    years.add(thisYear + 3);
    for (const l of licenses) {
      if (l.expiryDate) years.add(parseInt(l.expiryDate.slice(0, 4), 10));
    }
    const sorted = [...years].sort();
    const current = calCursor.getFullYear();
    sel.innerHTML = "";
    for (const y of sorted) {
      const opt = document.createElement("option");
      opt.value = String(y);
      opt.textContent = String(y);
      sel.appendChild(opt);
    }
    sel.value = String(current);
  }

  // ---------- v1.7.39 saved views ----------
  function currentFilterSnapshot() {
    return {
      quickFilter,
      summaryFilter,
      ownerFilter: [...ownerFilter],
      productFilter: [...productFilter],
      statusFilter: [...statusFilter],
      expiryFilter: [...expiryFilter],
      monthFilter,
      quarterFilter,
      dateFromFilter,
      dateToFilter,
      searchText,
      groupBy,
    };
  }
  function applyFilterSnapshot(snap) {
    if (!snap || typeof snap !== "object") return;
    quickFilter = typeof snap.quickFilter === "string" ? snap.quickFilter : "all";
    summaryFilter = snap.summaryFilter || null;
    ownerFilter = new Set(Array.isArray(snap.ownerFilter) ? snap.ownerFilter : []);
    productFilter = new Set(Array.isArray(snap.productFilter) ? snap.productFilter : []);
    statusFilter = new Set(Array.isArray(snap.statusFilter) ? snap.statusFilter : []);
    expiryFilter = new Set(Array.isArray(snap.expiryFilter) ? snap.expiryFilter : []);
    monthFilter = typeof snap.monthFilter === "string" ? snap.monthFilter : "";
    quarterFilter = typeof snap.quarterFilter === "string" ? snap.quarterFilter : "";
    dateFromFilter = typeof snap.dateFromFilter === "string" ? snap.dateFromFilter : "";
    dateToFilter = typeof snap.dateToFilter === "string" ? snap.dateToFilter : "";
    searchText = typeof snap.searchText === "string" ? snap.searchText : "";
    if (typeof snap.groupBy === "string") groupBy = snap.groupBy;
    persistFilters();
    try { localStorage.setItem(LS_QUICK, quickFilter); } catch (_) {}
    try { localStorage.setItem(LS_GROUP, groupBy); } catch (_) {}
    $("searchInput").value = searchText;
    $("searchClear").hidden = !searchText;
    $("monthFilter").value = monthFilter;
    if ($("quarterFilter")) $("quarterFilter").value = quarterFilter;
    if ($("dateFromFilter")) $("dateFromFilter").value = dateFromFilter;
    if ($("dateToFilter")) $("dateToFilter").value = dateToFilter;
  }
  async function saveCurrentView() {
    const name = prompt("Name this view:", `View ${(userSettings.savedLicenseViews || []).length + 1}`);
    if (!name || !name.trim()) return;
    const view = {
      id: `v-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`,
      name: name.trim().slice(0, 80),
      filters: currentFilterSnapshot(),
    };
    const next = {
      ...userSettings,
      savedLicenseViews: [...(userSettings.savedLicenseViews || []), view].slice(-25),
    };
    try {
      const { settings } = await api("PUT", "/settings", { settings: next });
      userSettings = { ...userSettings, ...settings };
      renderSavedViews();
      toast(`Saved view: ${view.name}`);
    } catch (err) { showError("Save view failed", err); }
  }
  async function deleteSavedView(id) {
    const views = (userSettings.savedLicenseViews || []).filter((v) => v.id !== id);
    const next = { ...userSettings, savedLicenseViews: views };
    try {
      const { settings } = await api("PUT", "/settings", { settings: next });
      userSettings = { ...userSettings, ...settings };
      renderSavedViews();
    } catch (err) { showError("Delete view failed", err); }
  }
  // v1.7.39 — tenant-wide activity sidebar: last 20 events across all licenses.
  // Sourced from license.events[] so no new endpoint. Reads "Dona renewed Acme
  // M365 · 3 min ago" style.
  function renderActivityStrip() {
    const strip = $("activityStrip");
    const list = $("activityList");
    if (!strip || !list) return;
    const flat = [];
    for (const lic of licenses) {
      if (!Array.isArray(lic.events)) continue;
      for (const ev of lic.events) {
        if (!ev || !ev.at) continue;
        flat.push({ ...ev, licenseId: lic.id, customer: lic.customer, licenseType: lic.licenseType });
      }
    }
    flat.sort((a, b) => b.at.localeCompare(a.at));
    const top = flat.slice(0, 20);
    list.innerHTML = "";
    if (!top.length) { strip.hidden = true; return; }
    strip.hidden = false;
    const verbMap = {
      created: "added",
      statusChanged: "changed status on",
      ownerChanged: "reassigned owner on",
      expiryChanged: "moved expiry on",
      renewed: "renewed",
      abandoned: "marked won't renew",
    };
    for (const ev of top) {
      const li = document.createElement("li");
      li.className = "activity-strip-item";
      const link = document.createElement("button");
      link.type = "button";
      link.className = "activity-strip-link";
      const verb = verbMap[ev.type] || ev.type;
      const who = ev.byName || "Someone";
      const target = `${ev.customer || "?"}, ${ev.licenseType || "?"}`;
      const when = fmtRelative(Date.parse(ev.at));
      link.innerHTML = `<strong>${escapeHtml(who)}</strong> ${escapeHtml(verb)} <em>${escapeHtml(target)}</em> · <span class="activity-when">${escapeHtml(when)}</span>`;
      link.addEventListener("click", () => {
        const lic = licenses.find((l) => l.id === ev.licenseId);
        if (lic) openEditDialog(lic);
      });
      li.appendChild(link);
      list.appendChild(li);
    }
  }
  function escapeHtml(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }

  function renderSavedViews() {
    const strip = $("savedViewsStrip");
    if (!strip) return;
    const views = userSettings.savedLicenseViews || [];
    strip.innerHTML = "";
    strip.hidden = views.length === 0;
    for (const v of views) {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "saved-view-chip";
      chip.title = `Apply saved view: ${v.name}`;
      const lab = document.createElement("span");
      lab.textContent = v.name;
      chip.appendChild(lab);
      const x = document.createElement("span");
      x.className = "saved-view-chip-x";
      x.setAttribute("role", "button");
      x.setAttribute("aria-label", `Delete saved view: ${v.name}`);
      x.textContent = "×";
      x.addEventListener("click", (e) => { e.stopPropagation(); deleteSavedView(v.id); });
      chip.appendChild(x);
      chip.addEventListener("click", () => { applyFilterSnapshot(v.filters); render(); });
      strip.appendChild(chip);
    }
  }

  function clearAllFilters() {
    quickFilter = "all";
    summaryFilter = null;
    ownerFilter.clear();
    productFilter.clear();
    statusFilter.clear();
    expiryFilter.clear();
    monthFilter = "";
    quarterFilter = "";
    dateFromFilter = "";
    dateToFilter = "";
    searchText = "";
    $("searchInput").value = "";
    $("searchClear").hidden = true;
    $("monthFilter").value = "";
    if ($("quarterFilter")) $("quarterFilter").value = "";
    if ($("dateFromFilter")) $("dateFromFilter").value = "";
    if ($("dateToFilter")) $("dateToFilter").value = "";
    try { localStorage.setItem(LS_QUICK, "all"); } catch (_) {}
    persistFilters();
    renderExpiryFilterMenu();
    render();
  }
  function setQuickFilter(next) {
    quickFilter = next;
    try { localStorage.setItem(LS_QUICK, next); } catch (_) {}
  }

  // Populate the Month dropdown with months that actually have expiring
  // licenses, in chronological order. Preserves the user's current selection
  // if the month is still present after a reload.
  function populateMonthDropdown() {
    const sel = $("monthFilter");
    if (!sel) return;
    const months = new Set();
    for (const lic of licenses) {
      if (!lic.expiryDate) continue;
      months.add(lic.expiryDate.slice(0, 7));
    }
    const sorted = [...months].sort();
    // Preserve selection if still present, else fall back to "any".
    if (monthFilter && !months.has(monthFilter)) monthFilter = "";
    const current = monthFilter;
    sel.innerHTML = "";
    const any = document.createElement("option");
    any.value = "";
    any.textContent = "Any month";
    sel.appendChild(any);
    for (const ym of sorted) {
      const opt = document.createElement("option");
      opt.value = ym;
      opt.textContent = fmtMonthShort(ym);
      sel.appendChild(opt);
    }
    sel.value = current;
  }

  // v1.8.0 — Quarter dropdown. Lists every YYYY-QN that contains at least
  // one expiring license, in chronological order. Like the month dropdown,
  // preserves the user's selection if still present.
  function populateQuarterDropdown() {
    const sel = $("quarterFilter");
    if (!sel) return;
    const quarters = new Set();
    for (const lic of licenses) {
      if (!lic.expiryDate) continue;
      const y = lic.expiryDate.slice(0, 4);
      const m = Number(lic.expiryDate.slice(5, 7));
      const q = Math.floor((m - 1) / 3) + 1;
      quarters.add(`${y}-Q${q}`);
    }
    const sorted = [...quarters].sort();
    if (quarterFilter && !quarters.has(quarterFilter)) quarterFilter = "";
    const current = quarterFilter;
    sel.innerHTML = "";
    const any = document.createElement("option");
    any.value = "";
    any.textContent = "Any quarter";
    sel.appendChild(any);
    for (const qk of sorted) {
      const opt = document.createElement("option");
      opt.value = qk;
      opt.textContent = fmtQuarterLabel(qk);
      sel.appendChild(opt);
    }
    sel.value = current;
  }

  // Expiry filter is a 4-option multi-select inside a popover (button shows
  // current selection summary). Re-rendered each time so checkbox state stays
  // in sync with `expiryFilter` mutations from chip removal or Clear all.
  function renderExpiryFilterMenu() {
    const menu = $("expiryFilterMenu");
    const valueEl = $("expiryFilterValue");
    if (!menu || !valueEl) return;
    menu.innerHTML = "";
    for (const b of EXPIRY_BUCKETS) {
      const lbl = document.createElement("label");
      lbl.className = "filter-popover-row";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.value = b;
      cb.checked = expiryFilter.has(b);
      cb.addEventListener("change", () => {
        if (cb.checked) expiryFilter.add(b);
        else expiryFilter.delete(b);
        persistFilters();
        renderExpiryFilterMenu();
        render();
      });
      const dot = document.createElement("span");
      dot.className = `expiry-dot exp-${b}`;
      const txt = document.createElement("span");
      txt.textContent = EXPIRY_LABEL[b];
      lbl.appendChild(cb);
      lbl.appendChild(dot);
      lbl.appendChild(txt);
      menu.appendChild(lbl);
    }
    if (expiryFilter.size === 0) valueEl.textContent = "Any";
    else if (expiryFilter.size === 1) valueEl.textContent = EXPIRY_LABEL[[...expiryFilter][0]];
    else valueEl.textContent = `${expiryFilter.size} selected`;
  }

  // ---------- stats + breakdowns ----------

  function recomputeStats() {
    const active = licenses.filter((l) => l.state !== "abandoned");
    const customers = new Set();
    let seats = 0;
    for (const l of active) {
      if (l.customer) customers.add(l.customer.trim().toLowerCase());
      seats += (typeof l.userCount === "number" ? l.userCount : 0);
    }
    // Renewed in last 30 days (ISO timestamp comparison)
    const cutoff = new Date(Date.now() - 30 * 86400000).toISOString();
    let renewed30 = 0;
    for (const l of active) if (l.lastRenewedAt && l.lastRenewedAt >= cutoff) renewed30++;
    $("statLicenses").textContent = active.length;
    $("statSeats").textContent = seats.toLocaleString();
    $("statCustomers").textContent = customers.size;
    $("statRenewed30").textContent = renewed30;
    renderRenewalRate();
    renderRenewalTrend();
    renderLeaderboard();
  }

  // v1.7.42 — 12-month renewal-rate trendline. Buckets license events into
  // monthly cohorts: renewed (lastRenewedAt's YYYY-MM) vs lapsed (expiryDate
  // fell in month AND not renewed/abandoned). Plotted on a hand-drawn canvas
  // (no Chart.js dependency). Hidden if every month is empty.
  function renderRenewalTrend() {
    const strip = $("renewalTrendStrip");
    const canvas = $("renewalTrendChart");
    if (!strip || !canvas) return;
    const months = [];
    const now = new Date();
    for (let i = 11; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      months.push({
        key: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`,
        label: d.toLocaleString(undefined, { month: "short" }),
        renewed: 0, lapsed: 0,
      });
    }
    const idxByKey = new Map(months.map((m, i) => [m.key, i]));
    const today = todayPh();
    for (const l of licenses) {
      if (l.lastRenewedAt) {
        const k = l.lastRenewedAt.slice(0, 7);
        if (idxByKey.has(k)) months[idxByKey.get(k)].renewed++;
      }
      if (l.expiryDate && l.expiryDate < today && l.state !== "abandoned" && l.status !== "renewed") {
        const k = l.expiryDate.slice(0, 7);
        if (idxByKey.has(k)) months[idxByKey.get(k)].lapsed++;
      }
    }
    const rates = months.map((m) => {
      const t = m.renewed + m.lapsed;
      return t > 0 ? Math.round((m.renewed / t) * 100) : null;
    });
    if (rates.every((r) => r === null)) { strip.hidden = true; return; }
    strip.hidden = false;
    const sub = $("renewalTrendSub");
    const recent = rates.filter((r) => r !== null);
    if (sub && recent.length >= 2) {
      const last = recent[recent.length - 1];
      const prev = recent[recent.length - 2];
      const delta = last - prev;
      sub.textContent = delta === 0 ? "Flat vs prior month"
        : `${delta > 0 ? "▲" : "▼"} ${Math.abs(delta)}pt vs prior month`;
      sub.classList.remove("trend-up", "trend-down", "trend-flat");
      sub.classList.add(delta > 0 ? "trend-up" : delta < 0 ? "trend-down" : "trend-flat");
    } else if (sub) sub.textContent = "";

    // Hand-drawn line + filled-area + month ticks. Resilient to dark mode
    // by reading the resolved foreground color from computed style.
    const ratio = window.devicePixelRatio || 1;
    const w = canvas.clientWidth || 600;
    const h = canvas.clientHeight || 120;
    canvas.width = Math.floor(w * ratio);
    canvas.height = Math.floor(h * ratio);
    const ctx = canvas.getContext("2d");
    ctx.scale(ratio, ratio);
    ctx.clearRect(0, 0, w, h);
    const PAD_L = 28, PAD_R = 12, PAD_T = 12, PAD_B = 22;
    const plotW = w - PAD_L - PAD_R;
    const plotH = h - PAD_T - PAD_B;
    // grid: 0 / 50 / 100
    const styles = getComputedStyle(canvas);
    const muted = styles.getPropertyValue("--muted").trim() || "#6a6a6a";
    const border = styles.getPropertyValue("--border").trim() || "#e4e4e4";
    const accent = styles.getPropertyValue("--accent").trim() || "#38AEEB";
    ctx.strokeStyle = border; ctx.lineWidth = 1; ctx.font = `${10}px Segoe UI, system-ui`;
    ctx.fillStyle = muted;
    for (const v of [0, 50, 100]) {
      const y = PAD_T + plotH - (v / 100) * plotH;
      ctx.beginPath(); ctx.moveTo(PAD_L, y); ctx.lineTo(PAD_L + plotW, y); ctx.stroke();
      ctx.fillText(`${v}%`, 4, y + 3);
    }
    // x-axis labels (every other month so they don't collide)
    for (let i = 0; i < months.length; i++) {
      if (i % 2 !== months.length % 2) continue;
      const x = PAD_L + (plotW * i) / Math.max(1, months.length - 1);
      ctx.fillText(months[i].label, x - 9, h - 6);
    }
    // Draw the line + filled area for months that have data.
    const pts = rates.map((r, i) => r === null ? null : ({
      x: PAD_L + (plotW * i) / Math.max(1, months.length - 1),
      y: PAD_T + plotH - (r / 100) * plotH,
      r,
    }));
    // Filled area under the curve (segments connecting consecutive non-null points)
    ctx.fillStyle = accent + "20"; // ~12% opacity
    ctx.beginPath();
    let started = false;
    let lastP = null;
    for (const p of pts) {
      if (!p) { if (started && lastP) { ctx.lineTo(lastP.x, PAD_T + plotH); ctx.closePath(); ctx.fill(); ctx.beginPath(); started = false; } continue; }
      if (!started) { ctx.moveTo(p.x, PAD_T + plotH); ctx.lineTo(p.x, p.y); started = true; }
      else ctx.lineTo(p.x, p.y);
      lastP = p;
    }
    if (started && lastP) { ctx.lineTo(lastP.x, PAD_T + plotH); ctx.closePath(); ctx.fill(); }
    // Line on top
    ctx.strokeStyle = accent; ctx.lineWidth = 2;
    ctx.beginPath();
    let move = true;
    for (const p of pts) {
      if (!p) { move = true; continue; }
      if (move) { ctx.moveTo(p.x, p.y); move = false; }
      else ctx.lineTo(p.x, p.y);
    }
    ctx.stroke();
    // Dots
    ctx.fillStyle = accent;
    for (const p of pts) {
      if (!p) continue;
      ctx.beginPath(); ctx.arc(p.x, p.y, 3, 0, Math.PI * 2); ctx.fill();
    }
    // v1.7.44 — click a month to filter the table to that month.
    canvas.onclick = (ev) => {
      const rect = canvas.getBoundingClientRect();
      const x = ev.clientX - rect.left;
      const stepW = plotW / Math.max(1, months.length - 1);
      let bestIdx = 0;
      let bestDist = Infinity;
      for (let i = 0; i < months.length; i++) {
        const px = PAD_L + (plotW * i) / Math.max(1, months.length - 1);
        const d = Math.abs(x - px);
        if (d < bestDist) { bestDist = d; bestIdx = i; }
      }
      const key = months[bestIdx].key; // YYYY-MM
      monthFilter = key;
      $("monthFilter").value = key;
      persistFilters();
      render();
      toast(`Filtered to ${months[bestIdx].label}`);
    };
    canvas.title = "Click a month to filter the table to it";
  }

  // v1.7.42 — owner leaderboard. Counts renewed-in-window vs total-touched
  // (renewed + lapsed + abandoned) per owner over the last 90 days. Sorts
  // by total touched so the most-active owners surface first.
  function renderLeaderboard() {
    const strip = $("leaderboardStrip");
    const list = $("leaderboardList");
    if (!strip || !list) return;
    const cutoffIso = new Date(Date.now() - 90 * 86400000).toISOString();
    const cutoffDate = cutoffIso.slice(0, 10);
    const today = todayPh();
    const byOwner = new Map(); // oid -> { name, renewed, lapsed, abandoned }
    function bump(oid, name, key) {
      if (!oid) return;
      const cur = byOwner.get(oid) || { oid, name, renewed: 0, lapsed: 0, abandoned: 0 };
      cur[key]++;
      cur.name = name || cur.name;
      byOwner.set(oid, cur);
    }
    for (const l of licenses) {
      const oid = l.ownerOid;
      const name = l.ownerName || "(no name)";
      if (l.lastRenewedAt && l.lastRenewedAt >= cutoffIso) { bump(oid, name, "renewed"); continue; }
      if (l.state === "abandoned" && Array.isArray(l.events)) {
        const ab = [...l.events].reverse().find((e) => e && e.type === "abandoned");
        if (ab && ab.at && ab.at >= cutoffIso) { bump(oid, name, "abandoned"); continue; }
      }
      if (l.expiryDate && l.expiryDate >= cutoffDate && l.expiryDate < today &&
          l.state !== "abandoned" && l.status !== "renewed") {
        bump(oid, name, "lapsed");
      }
    }
    const rows = [...byOwner.values()].filter((r) => r.renewed + r.lapsed + r.abandoned > 0);
    if (!rows.length) { strip.hidden = true; return; }
    rows.sort((a, b) => (b.renewed + b.lapsed + b.abandoned) - (a.renewed + a.lapsed + a.abandoned));
    strip.hidden = false;
    list.innerHTML = "";
    for (const r of rows.slice(0, 6)) {
      const total = r.renewed + r.lapsed + r.abandoned;
      const rate = Math.round((r.renewed / total) * 100);
      const li = document.createElement("li");
      li.className = "leaderboard-row";
      const head = document.createElement("div");
      head.className = "leaderboard-head";
      const name = document.createElement("button");
      name.type = "button";
      name.className = "leaderboard-name clickable";
      name.textContent = r.name;
      name.title = `Filter to licenses owned by ${r.name}`;
      name.addEventListener("click", () => {
        // Replace current owner filter with just this owner.
        ownerFilter.clear();
        ownerFilter.add(r.oid);
        persistFilters();
        render();
      });
      const score = document.createElement("strong");
      score.className = "leaderboard-score";
      score.textContent = `${rate}%`;
      score.classList.add(rate >= 80 ? "rate-good" : rate >= 60 ? "rate-ok" : "rate-bad");
      head.appendChild(name); head.appendChild(score);
      const meta = document.createElement("div");
      meta.className = "leaderboard-meta";
      meta.textContent = `${r.renewed} of ${total} renewed · ${r.lapsed} lapsed${r.abandoned ? ` · ${r.abandoned} won't renew` : ""}`;
      li.appendChild(head); li.appendChild(meta);
      list.appendChild(li);
    }
  }

  // v1.7.41 — last-90-days renewal-rate widget. Definitions:
  //   renewed  = lastRenewedAt within window
  //   abandoned = an "abandoned" event in window (state=abandoned today)
  //   lapsed   = expiryDate fell in window AND row is NOT renewed/abandoned
  // Rate = renewed / (renewed + abandoned + lapsed). Hidden when nothing
  // fell in the window so the widget doesn't show "—" for empty tenants.
  function renderRenewalRate() {
    const strip = $("renewalRateStrip");
    if (!strip) return;
    const windowMs = 90 * 86400000;
    const cutoffMs = Date.now() - windowMs;
    const cutoffIso = new Date(cutoffMs).toISOString();
    const cutoffDate = cutoffIso.slice(0, 10);
    const today = todayPh();
    let renewed = 0, abandoned = 0, lapsed = 0;
    for (const l of licenses) {
      // Renewed in window
      if (l.lastRenewedAt && l.lastRenewedAt >= cutoffIso) { renewed++; continue; }
      // Abandoned: state=abandoned today AND most recent "abandoned" event in window
      if (l.state === "abandoned" && Array.isArray(l.events)) {
        const ab = [...l.events].reverse().find((e) => e && e.type === "abandoned");
        if (ab && ab.at && ab.at >= cutoffIso) { abandoned++; continue; }
      }
      // Lapsed: expiryDate in window AND row is not renewed/abandoned
      if (l.expiryDate && l.expiryDate >= cutoffDate && l.expiryDate < today &&
          l.state !== "abandoned" && l.status !== "renewed") {
        lapsed++;
      }
    }
    const total = renewed + abandoned + lapsed;
    if (total === 0) { strip.hidden = true; return; }
    strip.hidden = false;
    const rate = Math.round((renewed / total) * 100);
    $("renewalRateValue").textContent = `${rate}%`;
    $("rrCntRenewed").textContent = renewed;
    $("rrCntLapsed").textContent = lapsed;
    $("rrCntAbandoned").textContent = abandoned;
    const pct = (n) => `${total ? (n / total) * 100 : 0}%`;
    $("renewalRateBarRenewed").style.width = pct(renewed);
    $("renewalRateBarLapsed").style.width = pct(lapsed);
    $("renewalRateBarAbandoned").style.width = pct(abandoned);
    // Color the headline by rate so Rey gets a glanceable health signal.
    strip.classList.remove("rate-good", "rate-ok", "rate-bad");
    strip.classList.add(rate >= 80 ? "rate-good" : rate >= 60 ? "rate-ok" : "rate-bad");
  }

  function renderOwnerChips() {
    const map = new Map(); // ownerOid -> { name, count, customers: Set, nextExpiry }
    const today = todayPh();
    for (const l of licenses) {
      if (l.state === "abandoned") continue;
      if (!l.ownerOid) continue;
      const key = l.ownerOid;
      const cur = map.get(key) || { oid: l.ownerOid, name: l.ownerName || "(no name)", count: 0, customers: new Set(), nextExpiry: null };
      cur.count++;
      if (l.customer) cur.customers.add(l.customer.trim().toLowerCase());
      // next expiry = soonest upcoming or zero-days expiry (>=today)
      if (l.expiryDate && (!cur.nextExpiry || l.expiryDate < cur.nextExpiry)) {
        if (l.expiryDate >= today) cur.nextExpiry = l.expiryDate;
        else if (!cur.nextExpiry) cur.nextExpiry = l.expiryDate; // fallback to overdue
      }
      map.set(key, cur);
    }
    const strip = $("ownerStrip");
    const wrap = $("ownerChips");
    wrap.innerHTML = "";
    if (!map.size) { strip.hidden = true; return; }
    strip.hidden = false;
    updateBreakdownLabel("ownerStrip", "Owner", ownerFilter);
    const entries = [...map.values()].sort((a, b) => b.count - a.count);
    for (const e of entries) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "lic-breakdown-chip";
      btn.setAttribute("aria-pressed", ownerFilter.has(e.oid) ? "true" : "false");
      // v1.7.44 — use the same round-avatar chip as the row owner cell, so
      // the sidebar, table, and active-filter strip read as the same person.
      const avatar = buildOwnerAvatar({ ownerOid: e.oid, ownerName: e.name }, { size: 18 });
      avatar.classList.add("chip-avatar");
      btn.appendChild(avatar);
      btn.appendChild(document.createTextNode(e.name));
      const cnt = document.createElement("span");
      cnt.className = "chip-count";
      cnt.textContent = e.count;
      btn.appendChild(cnt);
      // Subtitle: unique customer count + next expiry, so the chip distinguishes
      // workload (records, in the count pill) from concentration (customers).
      const subParts = [];
      if (e.customers.size > 0) subParts.push(`${e.customers.size} cust`);
      if (e.nextExpiry) {
        const d = daysBetween(today, e.nextExpiry);
        if (d !== null) subParts.push(d < 0 ? `overdue ${-d}d` : d === 0 ? "today" : `next ${fmtShortDate(e.nextExpiry)}`);
      }
      if (subParts.length) {
        const nx = document.createElement("span");
        nx.className = "chip-next";
        nx.textContent = subParts.join(" · ");
        btn.appendChild(nx);
      }
      btn.addEventListener("click", () => {
        if (ownerFilter.has(e.oid)) ownerFilter.delete(e.oid);
        else ownerFilter.add(e.oid);
        // Owner-chip selection overrides the Mine quick-filter so the chip wins.
        if (ownerFilter.size > 0 && quickFilter === "mine") quickFilter = "all";
        persistFilters();
        render();
      });
      wrap.appendChild(btn);
    }
  }

  // Render a small "(N) · Clear" affordance next to a breakdown label so a
  // user can see how many filters are active and reset them with one click.
  // Power BI-style multi-select: click chips to add/remove, click Clear to reset.
  function updateBreakdownLabel(stripId, labelText, filterSet) {
    const strip = document.getElementById(stripId);
    if (!strip) return;
    const labelEl = strip.querySelector(".lic-breakdown-label");
    if (!labelEl) return;
    labelEl.textContent = labelText;
    let trailing = strip.querySelector(".lic-breakdown-meta");
    if (filterSet.size > 0) {
      if (!trailing) {
        trailing = document.createElement("span");
        trailing.className = "lic-breakdown-meta";
        labelEl.after(trailing);
      }
      trailing.innerHTML = "";
      const count = document.createElement("span");
      count.className = "lic-breakdown-count";
      count.textContent = `${filterSet.size} selected`;
      const clear = document.createElement("button");
      clear.type = "button";
      clear.className = "lic-breakdown-clear";
      clear.textContent = "Clear";
      clear.title = "Clear all selections on this filter";
      clear.addEventListener("click", (e) => {
        e.stopPropagation();
        filterSet.clear();
        persistFilters();
        render();
      });
      trailing.appendChild(count);
      trailing.appendChild(clear);
    } else if (trailing) {
      trailing.remove();
    }
  }

  function renderStatusChips() {
    const map = new Map();
    for (const l of licenses) {
      if (l.state === "abandoned") continue;
      const s = l.status || "notStarted";
      map.set(s, (map.get(s) || 0) + 1);
    }
    const strip = $("statusStrip");
    const wrap = $("statusChips");
    wrap.innerHTML = "";
    if (!map.size) { strip.hidden = true; return; }
    strip.hidden = false;
    updateBreakdownLabel("statusStrip", "Status", statusFilter);
    // Keep canonical pipeline order.
    for (const status of STATUSES) {
      const count = map.get(status);
      if (!count) continue;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "lic-breakdown-chip";
      btn.setAttribute("aria-pressed", statusFilter.has(status) ? "true" : "false");
      const pill = document.createElement("span");
      pill.className = `status-pill status-${status}`;
      pill.textContent = STATUS_LABEL[status];
      btn.appendChild(pill);
      const cnt = document.createElement("span");
      cnt.className = "chip-count";
      cnt.textContent = count;
      btn.appendChild(cnt);
      btn.addEventListener("click", () => {
        if (statusFilter.has(status)) statusFilter.delete(status);
        else statusFilter.add(status);
        persistFilters();
        render();
      });
      wrap.appendChild(btn);
    }
  }

  function renderProductChips() {
    const map = new Map(); // productLine -> count
    for (const l of licenses) {
      if (l.state === "abandoned") continue;
      if (!l.productLine) continue;
      map.set(l.productLine, (map.get(l.productLine) || 0) + 1);
    }
    const strip = $("productStrip");
    const wrap = $("productChips");
    wrap.innerHTML = "";
    if (!map.size) { strip.hidden = true; return; }
    strip.hidden = false;
    updateBreakdownLabel("productStrip", "Product", productFilter);
    const entries = [...map.entries()].sort((a, b) => b[1] - a[1]);
    for (const [name, count] of entries) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "lic-breakdown-chip";
      btn.setAttribute("aria-pressed", productFilter.has(name) ? "true" : "false");
      btn.appendChild(document.createTextNode(name));
      const cnt = document.createElement("span");
      cnt.className = "chip-count";
      cnt.textContent = count;
      btn.appendChild(cnt);
      btn.addEventListener("click", () => {
        if (productFilter.has(name)) productFilter.delete(name);
        else productFilter.add(name);
        persistFilters();
        render();
      });
      wrap.appendChild(btn);
    }
  }

  function refreshDataLists() {
    const customers = new Set(), licTypes = new Set(), productLines = new Set();
    for (const lic of licenses) {
      if (lic.customer) customers.add(lic.customer);
      if (lic.licenseType) licTypes.add(lic.licenseType);
      if (lic.productLine) productLines.add(lic.productLine);
    }
    function fill(id, set) {
      const dl = $(id);
      if (!dl) return; // v1.8.0 — datalist may have been removed (e.g. productLineList)
      dl.innerHTML = "";
      Array.from(set).sort().forEach((v) => {
        const opt = document.createElement("option");
        opt.value = v;
        dl.appendChild(opt);
      });
    }
    fill("customerList", customers);
    fill("licTypeList", licTypes);
    // v1.8.0 — productLineList datalist was removed when the Edit dialog
    // input became a strict <select>. Leaving the helper-no-op behavior
    // above in case other datalists get removed later. Reference to the
    // productLines variable is preserved as a void to keep the lint clean.
    void productLines;
  }

  function ensureLicOwnerPicker(initialOid, initialName) {
    if (!licOwnerPicker) {
      licOwnerPicker = createPeoplePicker($("licOwner"), {
        initialOid,
        initialName,
        placeholder: "Search teammates by name…",
      });
    } else {
      licOwnerPicker.setValue(initialOid, initialName);
    }
  }

  // ---------- render ----------
  function render() {
    computeBundles();
    recomputeSummary();
    recomputeStats();
    refreshDataLists();
    populateMonthDropdown();
    populateQuarterDropdown();
    renderExpiryFilterMenu();
    renderOwnerChips();
    renderProductChips();
    renderStatusChips();
    renderActiveFilterBar();
    renderSavedViews();
    renderActivityStrip();
    populateCalYearJump();
    // v1.7.42 sync date-range button label
    const drv = $("dateRangeValue");
    if (drv) {
      drv.textContent = (dateFromFilter && dateToFilter)
        ? `${fmtShortDate(dateFromFilter)} – ${fmtShortDate(dateToFilter)}`
        : dateFromFilter ? `from ${fmtShortDate(dateFromFilter)}`
        : dateToFilter ? `until ${fmtShortDate(dateToFilter)}`
        : "Any";
    }
    if ($("dateFromFilter")) $("dateFromFilter").value = dateFromFilter;
    if ($("dateToFilter")) $("dateToFilter").value = dateToFilter;
    document.querySelectorAll("#viewSwitch .view-btn").forEach((b) => {
      b.setAttribute("aria-pressed", b.dataset.view === currentView ? "true" : "false");
    });
    document.querySelectorAll("#groupSwitch .view-btn").forEach((b) => {
      b.setAttribute("aria-pressed", b.dataset.group === groupBy ? "true" : "false");
    });
    document.querySelectorAll(".filter-pill").forEach((p) => {
      p.setAttribute("aria-pressed", (summaryFilter ? "false" : (p.dataset.quick === quickFilter ? "true" : "false")));
    });
    document.body.dataset.view = currentView;
    const qView = $("quarterView");
    if (currentView === "calendar") {
      $("tableView").hidden = true;
      $("calendarView").hidden = false;
      if (qView) qView.hidden = true;
      renderCalendar();
    } else if (currentView === "quarter") {
      $("tableView").hidden = true;
      $("calendarView").hidden = true;
      if (qView) qView.hidden = false;
      renderQuarter();
    } else {
      $("tableView").hidden = false;
      $("calendarView").hidden = true;
      if (qView) qView.hidden = true;
      renderTable();
    }
  }

  function buildLicenseRow(lic, today) {
    const tr = document.createElement("tr");
    tr.dataset.id = lic.id;
    if (bulkSelected.has(lic.id)) tr.classList.add("selected");

    if (bulkMode) {
      const tdCheck = document.createElement("td");
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.className = "lic-row-checkbox";
      cb.checked = bulkSelected.has(lic.id);
      cb.setAttribute("aria-label", `Select ${lic.customer} ${lic.licenseType}`);
      cb.addEventListener("click", (e) => { e.stopPropagation(); });
      cb.addEventListener("change", () => {
        if (cb.checked) bulkSelected.add(lic.id);
        else bulkSelected.delete(lic.id);
        tr.classList.toggle("selected", cb.checked);
        updateBulkBar();
      });
      tdCheck.appendChild(cb);
      tr.appendChild(tdCheck);
    }
    const d = daysBetween(today, lic.expiryDate);
    if (d !== null && d < 0 && lic.state !== "abandoned") tr.classList.add("overdue");
    if (lic.state === "abandoned") tr.classList.add("abandoned");
    // v1.7.39 — hover-tooltip with full notes so the cell text doesn't need
    // to be expanded to read context. Truncated for sanity.
    if (lic.notes) tr.title = lic.notes.length > 600 ? lic.notes.slice(0, 597) + "…" : lic.notes;

    // v1.8.0 — Product line is now the leading column. Built first so it
    // appends in column-1 position. (The previous tdProd block below is
    // skipped; we still build the same chip here, just in a different slot.)
    const tdProd = document.createElement("td");
    tdProd.className = "col-productline";
    if (lic.productLine) {
      const tag = document.createElement("span");
      tag.className = "product-tag";
      tag.textContent = lic.productLine;
      // v1.8.0 — flag legacy values (not in the registry) so users notice
      // them and can normalize. Registry is loaded into productLinesRegistry
      // once on boot; if it isn't ready yet, skip the flag rather than guess.
      if (productLinesRegistry && productLinesRegistry.length > 0) {
        const known = productLinesRegistry.some((p) => p.name === lic.productLine);
        if (!known) {
          tag.classList.add("product-tag-legacy");
          tag.title = "Legacy value — not in product line registry. Open Settings → Product lines to normalize.";
        }
      }
      tdProd.appendChild(tag);
    }
    tr.appendChild(tdProd);

    const tdCustomer = document.createElement("td");
    tdCustomer.className = "col-customer";
    const custBtn = document.createElement("button");
    custBtn.type = "button";
    custBtn.className = "customer-link";
    custBtn.textContent = lic.customer || "";
    custBtn.addEventListener("click", (e) => { e.stopPropagation(); openCustomerDialog(lic.customer); });
    // v1.7.44 — hover preview card with seats / owner / next expiry summary
    attachCustomerHoverPreview(custBtn, lic.customer);
    tdCustomer.appendChild(custBtn);
    // Y-of-N badge for non-annual cycles
    if (lic.renewalCycle && lic.renewalCycle !== "annual") {
      const badge = document.createElement("span");
      badge.className = "cycle-badge";
      badge.textContent = lic.renewalCycle === "biennial" ? "biennial" : "triennial";
      badge.title = `Renewal cycle: ${lic.renewalCycle}`;
      tdCustomer.appendChild(badge);
    }
    // Bundle membership is communicated on the Email button below (its label
    // changes to "Email 10" etc.). No badge in the customer cell so the
    // customer name has more room to breathe.
    const bundle = bundleFor(lic);
    tr.appendChild(tdCustomer);

    const tdType = document.createElement("td");
    tdType.className = "col-licensetype";
    tdType.textContent = lic.licenseType || "";
    tr.appendChild(tdType);

    const tdUsers = document.createElement("td");
    tdUsers.className = "num col-users";
    tdUsers.textContent = lic.userCount || 0;
    tr.appendChild(tdUsers);

    const tdExpiry = document.createElement("td");
    tdExpiry.className = "col-expires";
    const expSpan = document.createElement("span");
    expSpan.textContent = fmtShortDate(lic.expiryDate);
    tdExpiry.appendChild(expSpan);
    if (d !== null) {
      const badge = document.createElement("span");
      badge.className = "lic-day-badge";
      // v1.7.48 — colorize the badge by expiry bucket so days-left reads as
      // a status chip (consistent with the calendar pill stripe + expiry
      // pill colors) rather than mute grey text.
      const dayBucket = expiryBucket(d);
      if (dayBucket) badge.classList.add(`exp-${dayBucket}`);
      if (d < 0) badge.textContent = `${-d}d overdue`;
      else if (d === 0) badge.textContent = "today";
      else badge.textContent = `${d}d left`;
      tdExpiry.appendChild(badge);
      // v1.7.38 — 4-bucket Expiry pill (Expired / Soon / This month / Active)
      const bucket = expiryBucket(d);
      if (bucket) {
        const pill = document.createElement("span");
        pill.className = `expiry-pill exp-${bucket}`;
        pill.textContent = EXPIRY_LABEL[bucket];
        pill.title = `Click to filter by ${EXPIRY_LABEL[bucket]}`;
        pill.addEventListener("click", (e) => {
          e.stopPropagation();
          // Toggle: if this bucket is the only filter, clear it; else replace.
          if (expiryFilter.size === 1 && expiryFilter.has(bucket)) expiryFilter.clear();
          else { expiryFilter.clear(); expiryFilter.add(bucket); }
          persistFilters();
          render();
        });
        tdExpiry.appendChild(pill);
      }
    }
    tr.appendChild(tdExpiry);

    const tdOwner = document.createElement("td");
    tdOwner.className = "col-owner";
    const ownerWrap = document.createElement("span");
    ownerWrap.className = "owner-cell";
    ownerWrap.appendChild(buildOwnerAvatar(lic, { size: 22 }));
    const pill = document.createElement("span");
    pill.className = "owner-pill";
    pill.style.background = ownerColor(lic.ownerOid);
    pill.textContent = lic.ownerName || (lic.ownerOid ? lic.ownerOid.slice(0, 8) : "(none)");
    ownerWrap.appendChild(pill);
    tdOwner.appendChild(ownerWrap);
    tr.appendChild(tdOwner);

    // (Product line cell is the leading column now — built at the top of
    // this function. No second product cell here.)

    // Status pill column
    const tdStatus = document.createElement("td");
    tdStatus.className = "col-status";
    const sPill = document.createElement("span");
    const statusVal = lic.status || "notStarted";
    sPill.className = `status-pill status-${statusVal}`;
    sPill.textContent = STATUS_LABEL[statusVal] || statusVal;
    tdStatus.appendChild(sPill);
    tr.appendChild(tdStatus);

    const tdActions = document.createElement("td");
    tdActions.className = "actions col-actions";
    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "btn ghost small";
    editBtn.textContent = "Edit";
    // v1.7.39 — comment count badge so threads are visible from the row.
    const cmtCount = Array.isArray(lic.comments) ? lic.comments.length : 0;
    if (cmtCount > 0) {
      const badge = document.createElement("span");
      badge.className = "row-comment-badge";
      badge.textContent = `💬 ${cmtCount}`;
      badge.title = `${cmtCount} comment${cmtCount === 1 ? "" : "s"}`;
      editBtn.appendChild(document.createTextNode(" "));
      editBtn.appendChild(badge);
    }
    editBtn.addEventListener("click", (e) => { e.stopPropagation(); openEditDialog(lic); });
    tdActions.appendChild(editBtn);
    const emailBtn = document.createElement("button");
    emailBtn.type = "button";
    emailBtn.className = "btn ghost small";
    if (bundle.length >= 2) {
      emailBtn.textContent = `Email ${bundle.length}`;
      emailBtn.title = `Drafts ONE combined email covering all ${bundle.length} licenses for ${lic.customer} expiring around this date.`;
      emailBtn.classList.add("email-bundle");
    } else {
      emailBtn.textContent = "Email";
      emailBtn.title = "Open Outlook with a pre-filled renewal notice to the customer";
    }
    emailBtn.addEventListener("click", (e) => { e.stopPropagation(); emailCustomer(lic); });
    tdActions.appendChild(emailBtn);
    const renewBtn = document.createElement("button");
    renewBtn.type = "button";
    renewBtn.className = "btn primary small";
    renewBtn.textContent = "Renew";
    // v1.7.39 — single click = quick +1y (most-common path), Shift-click = open
    // the dialog for picking 2y/3y/custom-date.
    renewBtn.title = "Renew +1y (Shift-click for custom)";
    renewBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      if (e.shiftKey) { openRenewDialog(lic.id); return; }
      quickRenewOneYear(lic);
    });
    tdActions.appendChild(renewBtn);
    // Kebab overflow menu (Delete, future actions)
    const kebab = document.createElement("button");
    kebab.type = "button";
    kebab.className = "row-kebab";
    kebab.setAttribute("aria-label", "More actions");
    kebab.setAttribute("aria-haspopup", "true");
    kebab.textContent = "⋯"; // horizontal ellipsis ⋯
    kebab.addEventListener("click", (e) => { e.stopPropagation(); openRowMenu(lic, kebab); });
    tdActions.appendChild(kebab);
    tr.appendChild(tdActions);

    tr.addEventListener("click", () => openEditDialog(lic));
    return tr;
  }

  function buildGroupHeaderRow(label, group, collapsed, onToggle) {
    const tr = document.createElement("tr");
    tr.className = "lic-group-header" + (collapsed ? " collapsed" : "");
    tr.setAttribute("role", "button");
    tr.setAttribute("aria-expanded", String(!collapsed));
    tr.title = collapsed ? "Click to expand this group" : "Click to collapse this group";
    const td = document.createElement("td");
    td.colSpan = bulkMode ? 9 : 8;
    const chevron = document.createElement("span");
    chevron.className = "group-chevron";
    chevron.setAttribute("aria-hidden", "true");
    chevron.textContent = collapsed ? "▸" : "▾";
    td.appendChild(chevron);
    td.appendChild(document.createTextNode(" " + (label || "(none)")));
    const seats = group.reduce((s, l) => s + (typeof l.userCount === "number" ? l.userCount : 0), 0);
    const uniqueCustomers = new Set();
    for (const l of group) if (l.customer) uniqueCustomers.add(l.customer.trim().toLowerCase());
    const meta = document.createElement("span");
    meta.className = "group-count";
    const noun = group.length === 1 ? "license" : "licenses";
    const parts = [`${group.length} ${noun}`];
    // Show customer count only when it's actually informative (i.e. when grouped
    // by owner / product line, not when grouped by customer where it's always 1).
    if (uniqueCustomers.size > 1) parts.push(`${uniqueCustomers.size} customers`);
    parts.push(`${seats.toLocaleString()} seats`);
    meta.textContent = parts.join(" · ");
    td.appendChild(meta);
    tr.appendChild(td);
    tr.addEventListener("click", onToggle);
    return tr;
  }

  function renderTable() {
    const list = sortLicenses(visibleLicenses());
    const tbody = $("licTbody");
    const empty = $("licEmpty");
    const tbl = $("licTable");
    tbl.classList.toggle("bulk-mode", bulkMode);
    // Insert/remove the checkbox column header dynamically.
    const headerRow = tbl.querySelector("thead tr");
    const existingCheckHeader = headerRow.querySelector("th.bulk-header");
    if (bulkMode && !existingCheckHeader) {
      const th = document.createElement("th");
      th.className = "bulk-header";
      const masterCb = document.createElement("input");
      masterCb.type = "checkbox";
      masterCb.setAttribute("aria-label", "Select all visible");
      masterCb.addEventListener("change", () => {
        if (masterCb.checked) for (const l of list) bulkSelected.add(l.id);
        else bulkSelected.clear();
        renderTable();
        updateBulkBar();
      });
      th.appendChild(masterCb);
      headerRow.insertBefore(th, headerRow.firstChild);
    } else if (!bulkMode && existingCheckHeader) {
      headerRow.removeChild(existingCheckHeader);
    }
    tbody.innerHTML = "";
    if (!list.length) {
      tbl.hidden = true;
      empty.hidden = false;
      // v1.7.38 — distinguish "no data at all" from "filters hid everything"
      // so users don't think their data was wiped. Swaps the hero copy in
      // place and offers a 1-click Clear all filters out.
      const isFiltered = hasAnyFilter() && licenses.length > 0;
      const heroH = empty.querySelector("h2");
      const heroP = empty.querySelector("p");
      const actions = empty.querySelector(".lic-empty-actions");
      const filteredId = "licEmptyClearFilters";
      const addBtn = $("licEmptyAdd");
      const importBtn = $("licEmptyImport");
      let clearBtn = document.getElementById(filteredId);
      if (isFiltered) {
        if (heroH) heroH.textContent = "No licenses match these filters";
        if (heroP) heroP.textContent = `You have ${licenses.length} license${licenses.length === 1 ? "" : "s"} total. Clear the filters above to see them.`;
        if (addBtn) addBtn.hidden = true;
        if (importBtn) importBtn.hidden = true;
        if (!clearBtn && actions) {
          clearBtn = document.createElement("button");
          clearBtn.id = filteredId;
          clearBtn.type = "button";
          clearBtn.className = "btn primary";
          clearBtn.textContent = "Clear all filters";
          clearBtn.addEventListener("click", clearAllFilters);
          actions.appendChild(clearBtn);
        }
        if (clearBtn) clearBtn.hidden = false;
      } else {
        if (heroH) heroH.textContent = "No licenses yet";
        if (heroP) heroP.textContent = "Track when your customers' Microsoft licenses (and anything else with a renewal date) need to be renewed.";
        if (addBtn) addBtn.hidden = false;
        if (importBtn) importBtn.hidden = false;
        if (clearBtn) clearBtn.hidden = true;
      }
      return;
    }
    tbl.hidden = false;
    empty.hidden = true;
    const today = todayPh();

    if (groupBy === "none") {
      for (const lic of list) tbody.appendChild(buildLicenseRow(lic, today));
    } else {
      const groups = new Map();
      for (const lic of list) {
        const key = (lic[groupBy] || "(none)").toString();
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push(lic);
      }
      const orderedKeys = [...groups.keys()].sort((a, b) => a.localeCompare(b));
      for (const key of orderedKeys) {
        const group = groups.get(key);
        const sectionKey = `${groupBy}:${key}`;
        const collapsed = collapsedGroups.has(sectionKey);
        const headerRow = buildGroupHeaderRow(key, group, collapsed, () => {
          if (collapsedGroups.has(sectionKey)) collapsedGroups.delete(sectionKey);
          else collapsedGroups.add(sectionKey);
          saveCollapsedGroups();
          renderTable();
        });
        tbody.appendChild(headerRow);
        if (!collapsed) {
          for (const lic of group) tbody.appendChild(buildLicenseRow(lic, today));
        }
      }
    }

    document.querySelectorAll("th.sortable").forEach((th) => {
      if (th.dataset.sort === sortKey) {
        th.setAttribute("aria-sort", sortDir === 1 ? "ascending" : "descending");
      } else {
        th.removeAttribute("aria-sort");
      }
    });
  }

  // ---------- calendar ----------
  function renderCalendar() {
    const title = $("calTitle");
    title.textContent = fmtMonth(calCursor);
    const body = $("calBody");
    body.innerHTML = "";

    const firstOfMonth = new Date(calCursor.getFullYear(), calCursor.getMonth(), 1);
    const lastOfMonth = new Date(calCursor.getFullYear(), calCursor.getMonth() + 1, 0);
    // Monday-first grid. JS getDay: 0=Sun..6=Sat → convert to Mon=0..Sun=6.
    const startWeekday = (firstOfMonth.getDay() + 6) % 7;

    // Bucket visible licenses by ISO date.
    const visible = visibleLicenses();
    const byDate = new Map();
    for (const lic of visible) {
      if (!lic.expiryDate) continue;
      const arr = byDate.get(lic.expiryDate) || [];
      arr.push(lic);
      byDate.set(lic.expiryDate, arr);
    }

    const todayKey = todayPh();

    // Leading blanks
    for (let i = 0; i < startWeekday; i++) {
      const cell = document.createElement("div");
      cell.className = "lic-cal-cell muted";
      body.appendChild(cell);
    }

    const daysInMonth = lastOfMonth.getDate();
    for (let day = 1; day <= daysInMonth; day++) {
      const cell = document.createElement("div");
      cell.className = "lic-cal-cell";
      const yyyy = calCursor.getFullYear();
      const mm = String(calCursor.getMonth() + 1).padStart(2, "0");
      const dd = String(day).padStart(2, "0");
      const isoDate = `${yyyy}-${mm}-${dd}`;
      cell.dataset.date = isoDate;
      if (isoDate === todayKey) cell.classList.add("today");

      const num = document.createElement("div");
      num.className = "lic-cal-num";
      num.textContent = day;
      cell.appendChild(num);

      const items = byDate.get(isoDate) || [];
      const MAX_VISIBLE = 3;
      const visiblePills = items.slice(0, MAX_VISIBLE);
      for (const lic of visiblePills) {
        const pill = document.createElement("button");
        pill.type = "button";
        pill.className = "lic-cal-pill";
        // v1.7.38 — left-edge expiry-status stripe (color-codes the bucket
        // without overriding the owner-color fill that keeps the calendar
        // visually grouped by who owns each renewal).
        const d = daysBetween(todayKey, lic.expiryDate);
        const bucket = expiryBucket(d);
        if (bucket) pill.classList.add(`exp-${bucket}`);
        if (lic.state === "abandoned") pill.classList.add("abandoned");
        pill.style.background = ownerColor(lic.ownerOid);
        const text = `${lic.customer} · ${lic.licenseType}`;
        pill.textContent = text.length > 40 ? text.slice(0, 39) + "…" : text;
        const bucketTitle = bucket ? ` [${EXPIRY_LABEL[bucket]}]` : "";
        pill.title = `${lic.customer} — ${lic.licenseType} (${lic.userCount} users, owner: ${lic.ownerName || "—"})${bucketTitle}`;
        pill.addEventListener("click", (e) => { e.stopPropagation(); openEditDialog(lic); });
        cell.appendChild(pill);
      }
      if (items.length > MAX_VISIBLE) {
        const more = document.createElement("button");
        more.type = "button";
        more.className = "lic-cal-more";
        more.textContent = `+${items.length - MAX_VISIBLE} more`;
        more.addEventListener("click", (e) => { e.stopPropagation(); openDayDialog(isoDate, items); });
        cell.appendChild(more);
      }
      cell.addEventListener("click", () => openAddDialogForDate(isoDate));
      body.appendChild(cell);
    }
  }

  // v1.8.0 — Quarter view. Renders the 3 months of `quarterCursor` as
  // stacked month grids so the user can scan a quarter without losing the
  // per-day pill detail. Each month sub-grid uses the same pill/density
  // styling as Calendar view. Pill clicks open Edit; cell clicks open Add
  // for that date.
  function renderQuarter() {
    const title = $("qTitle");
    if (!title) return;
    const { year, q } = quarterCursor;
    title.textContent = `Q${q} ${year}`;

    const body = $("qBody");
    body.innerHTML = "";

    const visible = visibleLicenses();
    const byDate = new Map();
    for (const lic of visible) {
      if (!lic.expiryDate) continue;
      const arr = byDate.get(lic.expiryDate) || [];
      arr.push(lic);
      byDate.set(lic.expiryDate, arr);
    }
    const todayKey = todayPh();
    const firstMonthIdx = (q - 1) * 3; // 0-based: Q1 -> 0, Q2 -> 3, Q3 -> 6, Q4 -> 9

    for (let mOffset = 0; mOffset < 3; mOffset++) {
      const monthIdx = firstMonthIdx + mOffset;
      const monthDate = new Date(year, monthIdx, 1);
      const monthLabel = monthDate.toLocaleString(undefined, { month: "long", year: "numeric" });

      const sec = document.createElement("section");
      sec.className = "lic-quarter-month";
      const h3 = document.createElement("h3");
      h3.className = "lic-quarter-month-title";
      h3.textContent = monthLabel;
      sec.appendChild(h3);

      const grid = document.createElement("div");
      grid.className = "lic-cal-grid";
      // Day headers
      ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"].forEach((d) => {
        const h = document.createElement("div");
        h.className = "lic-cal-header";
        h.textContent = d;
        grid.appendChild(h);
      });
      const monthBody = document.createElement("div");
      monthBody.className = "lic-cal-body";
      monthBody.setAttribute("role", "grid");

      const firstOfMonth = new Date(year, monthIdx, 1);
      const lastOfMonth = new Date(year, monthIdx + 1, 0);
      const startWeekday = (firstOfMonth.getDay() + 6) % 7;
      for (let i = 0; i < startWeekday; i++) {
        const cell = document.createElement("div");
        cell.className = "lic-cal-cell muted";
        monthBody.appendChild(cell);
      }
      const daysInMonth = lastOfMonth.getDate();
      for (let day = 1; day <= daysInMonth; day++) {
        const cell = document.createElement("div");
        cell.className = "lic-cal-cell";
        const yyyy = year;
        const mm = String(monthIdx + 1).padStart(2, "0");
        const dd = String(day).padStart(2, "0");
        const isoDate = `${yyyy}-${mm}-${dd}`;
        cell.dataset.date = isoDate;
        if (isoDate === todayKey) cell.classList.add("today");

        const num = document.createElement("div");
        num.className = "lic-cal-num";
        num.textContent = day;
        cell.appendChild(num);

        const items = byDate.get(isoDate) || [];
        const MAX_VISIBLE = 3;
        for (const lic of items.slice(0, MAX_VISIBLE)) {
          const pill = document.createElement("button");
          pill.type = "button";
          pill.className = "lic-cal-pill";
          const d = daysBetween(todayKey, lic.expiryDate);
          const bucket = expiryBucket(d);
          if (bucket) pill.classList.add(`exp-${bucket}`);
          if (lic.state === "abandoned") pill.classList.add("abandoned");
          pill.style.background = ownerColor(lic.ownerOid);
          const text = `${lic.customer} · ${lic.licenseType}`;
          pill.textContent = text.length > 40 ? text.slice(0, 39) + "…" : text;
          const bucketTitle = bucket ? ` [${EXPIRY_LABEL[bucket]}]` : "";
          pill.title = `${lic.customer} — ${lic.licenseType} (${lic.userCount} users, owner: ${lic.ownerName || "—"})${bucketTitle}`;
          pill.addEventListener("click", (e) => { e.stopPropagation(); openEditDialog(lic); });
          cell.appendChild(pill);
        }
        if (items.length > MAX_VISIBLE) {
          const more = document.createElement("button");
          more.type = "button";
          more.className = "lic-cal-more";
          more.textContent = `+${items.length - MAX_VISIBLE} more`;
          more.addEventListener("click", (e) => { e.stopPropagation(); openDayDialog(isoDate, items); });
          cell.appendChild(more);
        }
        cell.addEventListener("click", () => openAddDialogForDate(isoDate));
        monthBody.appendChild(cell);
      }
      grid.appendChild(monthBody);
      sec.appendChild(grid);
      body.appendChild(sec);
    }

    // Populate the Quarter-jump dropdown (mirrors calYearJump).
    const yearJump = $("qYearJump");
    if (yearJump) {
      const quarters = new Set();
      const thisYear = new Date().getFullYear();
      [thisYear - 1, thisYear, thisYear + 1, thisYear + 2].forEach((y) => {
        for (let qi = 1; qi <= 4; qi++) quarters.add(`${y}-Q${qi}`);
      });
      for (const lic of licenses) {
        if (!lic.expiryDate) continue;
        const y = lic.expiryDate.slice(0, 4);
        const m = Number(lic.expiryDate.slice(5, 7));
        quarters.add(`${y}-Q${Math.floor((m - 1) / 3) + 1}`);
      }
      const sorted = [...quarters].sort();
      const current = `${year}-Q${q}`;
      yearJump.innerHTML = "";
      for (const qk of sorted) {
        const opt = document.createElement("option");
        opt.value = qk;
        opt.textContent = fmtQuarterLabel(qk);
        yearJump.appendChild(opt);
      }
      yearJump.value = current;
    }
  }

  // ---------- lead-day picker (v1.7.37) ----------
  // Backs a chip-style multi-select for the per-license lead-days field and
  // the per-user default in Settings. Each picker container carries data-*
  // attributes pointing at its chip strip, presets row, and custom input so
  // one factory powers both dialogs.
  const leadPickers = {};
  function attachLeadPicker(rootId) {
    const root = $(rootId);
    if (!root) return null;
    const chips = $(root.dataset.chipId);
    const presets = $(root.dataset.presetId);
    const customInput = $(root.dataset.customId);
    const customAdd = $(root.dataset.customAddId);
    const state = { values: [] };

    function asArr(input) {
      if (input === null || input === undefined) return [];
      if (typeof input === "number") return [input];
      if (!Array.isArray(input)) return [];
      return input
        .map((n) => parseInt(n, 10))
        .filter((n) => Number.isFinite(n) && n >= 0 && n <= 365);
    }
    function sortValues() {
      state.values = Array.from(new Set(state.values)).sort((a, b) => b - a).slice(0, 10);
    }
    function renderChips() {
      chips.innerHTML = "";
      if (!state.values.length) {
        const hint = document.createElement("span");
        hint.className = "lead-empty";
        hint.textContent = "Using Settings default";
        chips.appendChild(hint);
      } else {
        // v1.7.48 — mark non-preset chips so they render with a dashed border,
        // distinguishing custom values (e.g. "14 d") from preset ones (e.g.
        // "15 d") at a glance.
        const PRESETS = new Set([60, 30, 15, 7, 1]);
        for (const d of state.values) {
          const chip = document.createElement("span");
          chip.className = "lead-chip" + (PRESETS.has(d) ? "" : " is-custom");
          chip.dataset.days = String(d);
          const label = document.createElement("span");
          label.textContent = `${d} d`;
          const x = document.createElement("button");
          x.type = "button";
          x.className = "lead-chip-x";
          x.setAttribute("aria-label", `Remove ${d}-day reminder`);
          x.textContent = "×";
          x.addEventListener("click", () => {
            state.values = state.values.filter((v) => v !== d);
            renderChips();
            renderPresets();
          });
          chip.appendChild(label);
          chip.appendChild(x);
          chips.appendChild(chip);
        }
      }
    }
    function renderPresets() {
      presets.querySelectorAll(".lead-preset").forEach((btn) => {
        const d = parseInt(btn.dataset.days, 10);
        btn.classList.toggle("active", state.values.includes(d));
      });
    }
    function add(d) {
      if (!Number.isFinite(d) || d < 0 || d > 365) return;
      if (state.values.length >= 10 && !state.values.includes(d)) return;
      if (state.values.includes(d)) state.values = state.values.filter((v) => v !== d);
      else state.values.push(d);
      sortValues();
      renderChips();
      renderPresets();
    }
    presets.querySelectorAll(".lead-preset").forEach((btn) => {
      btn.addEventListener("click", () => add(parseInt(btn.dataset.days, 10)));
    });
    customAdd.addEventListener("click", () => {
      customInput.hidden = false;
      customInput.value = "";
      customInput.focus();
    });
    customInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        const n = parseInt(customInput.value, 10);
        if (Number.isFinite(n) && n >= 0 && n <= 365) {
          add(n);
          customInput.value = "";
          customInput.hidden = true;
        }
      } else if (e.key === "Escape") {
        customInput.value = "";
        customInput.hidden = true;
      }
    });
    customInput.addEventListener("blur", () => {
      if (!customInput.value.trim()) customInput.hidden = true;
    });

    const ctrl = {
      set(values) {
        state.values = asArr(values);
        sortValues();
        customInput.hidden = true;
        customInput.value = "";
        renderChips();
        renderPresets();
      },
      get() {
        sortValues();
        return state.values.slice();
      },
    };
    leadPickers[rootId] = ctrl;
    return ctrl;
  }
  function ensureLeadPicker(rootId) {
    return leadPickers[rootId] || attachLeadPicker(rootId);
  }

  // ---------- dialogs ----------
  // v1.8.0 — Populate the Edit dialog's product-line <select> from the
  // registry. If `currentValue` isn't in the registry, append it as an extra
  // option marked "(legacy)" so the row stays editable until somebody runs
  // the normalize tool. Returns the resolved value applied to the select.
  function populateProductLineSelect(currentValue) {
    const sel = $("licProductLine");
    if (!sel) return "";
    sel.innerHTML = "";
    const noneOpt = document.createElement("option");
    noneOpt.value = "";
    noneOpt.textContent = "(none)";
    sel.appendChild(noneOpt);
    const seen = new Set();
    for (const pl of productLinesRegistry || []) {
      const opt = document.createElement("option");
      opt.value = pl.name;
      opt.textContent = pl.name;
      sel.appendChild(opt);
      seen.add(pl.name);
    }
    const current = String(currentValue || "");
    if (current && !seen.has(current)) {
      const legacy = document.createElement("option");
      legacy.value = current;
      legacy.textContent = `${current} (legacy)`;
      legacy.className = "product-option-legacy";
      sel.appendChild(legacy);
    }
    sel.value = current;
    return sel.value;
  }

  function openAddDialog() {
    editingId = null;
    $("licDialogTitle").textContent = "Add license";
    $("licCustomer").value = "";
    $("licType").value = "";
    $("licUsers").value = "1";
    $("licExpiry").value = todayPh();
    ensureLicOwnerPicker(me.oid || null, me.name || "");
    populateProductLineSelect("");
    $("licStatus").value = "notStarted";
    $("licRenewalCycle").value = "annual";
    ensureLeadPicker("licLeadDaysPicker").set([]);
    $("licNotes").value = "";
    $("licActivity").hidden = true;
    $("licComments").hidden = true;
    $("licEmailBtn").hidden = true;
    $("licRenewBtn").hidden = true;
    $("licDeleteBtn").hidden = true;
    $("licDialog").showModal();
    $("licCustomer").focus();
  }
  function openAddDialogForDate(isoDate) {
    openAddDialog();
    $("licExpiry").value = isoDate;
  }
  function openEditDialog(lic) {
    editingId = lic.id;
    $("licDialogTitle").textContent = "Edit license";
    $("licCustomer").value = lic.customer || "";
    $("licType").value = lic.licenseType || "";
    $("licUsers").value = lic.userCount || 0;
    $("licExpiry").value = lic.expiryDate || todayPh();
    ensureLicOwnerPicker(lic.ownerOid || null, lic.ownerName || "");
    populateProductLineSelect(lic.productLine || "");
    $("licStatus").value = lic.status || "notStarted";
    $("licRenewalCycle").value = lic.renewalCycle || "annual";
    // leadDays may be: array (v1.7.37+), scalar (legacy row), or null/undefined.
    let leadInit = [];
    if (Array.isArray(lic.leadDays)) leadInit = lic.leadDays;
    else if (typeof lic.leadDays === "number") leadInit = [lic.leadDays];
    ensureLeadPicker("licLeadDaysPicker").set(leadInit);
    $("licNotes").value = lic.notes || "";
    renderActivityLog(lic);
    renderComments(lic);
    $("licEmailBtn").hidden = false;
    $("licRenewBtn").hidden = false;
    $("licDeleteBtn").hidden = false;
    $("licDialog").showModal();
    $("licCustomer").focus();
  }

  // v1.7.39 — per-license comments thread.
  function renderComments(lic) {
    const section = $("licComments");
    const list = $("licCommentsList");
    if (!section || !list) return;
    section.hidden = false;
    list.innerHTML = "";
    const comments = Array.isArray(lic.comments) ? lic.comments.slice().sort((a, b) => (a.at || "").localeCompare(b.at || "")) : [];
    if (!comments.length) {
      const li = document.createElement("li");
      li.className = "lic-comments-empty";
      li.textContent = "No comments yet. Add the first one below.";
      list.appendChild(li);
      return;
    }
    for (const c of comments) {
      const li = document.createElement("li");
      li.className = "lic-comment";
      li.dataset.commentId = c.id || "";
      const head = document.createElement("div");
      head.className = "lic-comment-head";
      const who = document.createElement("span");
      who.className = "lic-comment-who";
      who.textContent = c.byName || "(unknown)";
      const when = document.createElement("span");
      when.className = "lic-comment-when";
      when.textContent = c.at ? new Date(c.at).toLocaleString(undefined, { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }) : "";
      if (c.editedAt) {
        const edited = document.createElement("span");
        edited.className = "lic-comment-edited";
        edited.textContent = " (edited)";
        edited.title = `Last edited ${new Date(c.editedAt).toLocaleString()}`;
        when.appendChild(edited);
      }
      head.appendChild(who); head.appendChild(when);
      // v1.7.45 — edit + delete affordances for the original author only.
      if (c.byOid === me.oid && c.id) {
        const tools = document.createElement("span");
        tools.className = "lic-comment-tools";
        const editBtn = document.createElement("button");
        editBtn.type = "button";
        editBtn.className = "lic-comment-tool";
        editBtn.title = "Edit comment";
        editBtn.setAttribute("aria-label", "Edit comment");
        editBtn.textContent = "Edit";
        editBtn.addEventListener("click", () => beginEditComment(li, c));
        const delBtn = document.createElement("button");
        delBtn.type = "button";
        delBtn.className = "lic-comment-tool danger";
        delBtn.title = "Delete comment";
        delBtn.setAttribute("aria-label", "Delete comment");
        delBtn.textContent = "Delete";
        delBtn.addEventListener("click", () => deleteOwnComment(c));
        tools.appendChild(editBtn); tools.appendChild(delBtn);
        head.appendChild(tools);
      }
      const body = document.createElement("div");
      body.className = "lic-comment-body";
      body.textContent = c.text || "";
      li.appendChild(head); li.appendChild(body);
      list.appendChild(li);
    }
    // Auto-scroll to newest comment
    list.scrollTop = list.scrollHeight;
  }

  // v1.7.45 — inline-edit a comment row: replace body with a textarea + Save/Cancel.
  function beginEditComment(li, c) {
    if (!li || !c || !c.id) return;
    const body = li.querySelector(".lic-comment-body");
    if (!body) return;
    body.hidden = true;
    const editor = document.createElement("div");
    editor.className = "lic-comment-edit";
    const ta = document.createElement("textarea");
    ta.rows = 2; ta.maxLength = 1000;
    ta.value = c.text || "";
    editor.appendChild(ta);
    const actions = document.createElement("div");
    actions.className = "lic-comment-edit-actions";
    const save = document.createElement("button");
    save.type = "button"; save.className = "btn primary small"; save.textContent = "Save";
    const cancel = document.createElement("button");
    cancel.type = "button"; cancel.className = "btn ghost small"; cancel.textContent = "Cancel";
    actions.appendChild(save); actions.appendChild(cancel);
    editor.appendChild(actions);
    li.appendChild(editor);
    ta.focus(); ta.setSelectionRange(ta.value.length, ta.value.length);
    cancel.addEventListener("click", () => { editor.remove(); body.hidden = false; });
    ta.addEventListener("keydown", (e) => {
      if (e.key === "Escape") { editor.remove(); body.hidden = false; }
      if ((e.ctrlKey || e.metaKey) && e.key === "Enter") save.click();
    });
    save.addEventListener("click", async () => {
      const text = ta.value.trim();
      if (!text) return;
      save.disabled = true; save.textContent = "Saving…";
      try {
        const { license } = await api("PATCH", `/licenses/${editingId}/comments/${c.id}`, { text });
        licenses = licenses.map((l) => l.id === license.id ? license : l);
        renderComments(license);
      } catch (err) { showError("Save failed", err); save.disabled = false; save.textContent = "Save"; }
    });
  }

  async function deleteOwnComment(c) {
    if (!editingId || !c || !c.id) return;
    if (!confirm("Delete this comment? This can't be undone.")) return;
    try {
      const { license } = await api("DELETE", `/licenses/${editingId}/comments/${c.id}`);
      licenses = licenses.map((l) => l.id === license.id ? license : l);
      renderComments(license);
    } catch (err) { showError("Delete failed", err); }
  }
  async function sendComment() {
    if (!editingId) return;
    const input = $("licCommentInput");
    const text = (input.value || "").trim();
    if (!text) return;
    const btn = $("licCommentSendBtn");
    btn.disabled = true;
    try {
      const { license } = await api("POST", `/licenses/${editingId}/comments`, { text });
      licenses = licenses.map((l) => l.id === license.id ? license : l);
      input.value = "";
      renderComments(license);
    } catch (err) { showError("Send failed", err); }
    finally { btn.disabled = false; input.focus(); }
  }

  function renderActivityLog(lic) {
    const events = Array.isArray(lic.events) ? lic.events : [];
    const list = $("licActivityList");
    list.innerHTML = "";
    if (!events.length) { $("licActivity").hidden = true; return; }
    $("licActivity").hidden = false;
    // newest first
    const sorted = [...events].sort((a, b) => (b.at || "").localeCompare(a.at || ""));
    for (const ev of sorted.slice(0, 30)) {
      const li = document.createElement("li");
      const when = document.createElement("span");
      when.className = "lic-activity-when";
      when.textContent = ev.at ? new Date(ev.at).toLocaleString(undefined, { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }) : "";
      const who = document.createElement("span");
      who.className = "lic-activity-who";
      who.textContent = ev.byName || "(unknown)";
      const action = document.createElement("span");
      const verbMap = { created: "added the license", statusChanged: "changed status", ownerChanged: "reassigned owner", expiryChanged: "moved expiry", renewed: "marked renewed" };
      const verb = verbMap[ev.type] || ev.type;
      action.textContent = ` ${verb}${ev.detail ? ": " + ev.detail : ""}`;
      li.appendChild(when);
      li.appendChild(document.createTextNode(" "));
      li.appendChild(who);
      li.appendChild(action);
      list.appendChild(li);
    }
  }
  function closeEditDialog() {
    $("licDialog").close();
    editingId = null;
  }

  function readLicenseForm() {
    const picked = licOwnerPicker ? licOwnerPicker.getValue() : { oid: null, name: null };
    const ownerOid = picked.oid || null;
    const ownerNameFromPicker = picked.name || null;
    const leadArr = ensureLeadPicker("licLeadDaysPicker").get();
    const leadDays = leadArr.length ? leadArr : null;
    return {
      customer: $("licCustomer").value.trim(),
      licenseType: $("licType").value.trim(),
      userCount: parseInt($("licUsers").value, 10) || 0,
      expiryDate: $("licExpiry").value,
      ownerOid,
      ownerName: ownerNameFromPicker,
      productLine: $("licProductLine").value.trim() || null,
      status: $("licStatus").value || "notStarted",
      renewalCycle: $("licRenewalCycle").value || "annual",
      leadDays,
      notes: $("licNotes").value.trim() || null,
    };
  }
  async function saveLicense() {
    const payload = readLicenseForm();
    if (!payload.customer) { toast("Customer is required"); return; }
    if (!payload.licenseType) { toast("License type is required"); return; }
    if (!payload.expiryDate) { toast("Expiry date is required"); return; }
    if (!payload.ownerOid) { toast("Owner is required"); return; }
    try {
      if (editingId) {
        const current = licenses.find((l) => l.id === editingId);
        const ifMatch = current && current.lastEditedAt ? current.lastEditedAt : null;
        const { license } = await api("PATCH", `/licenses/${editingId}`, payload, { ifMatch });
        licenses = licenses.map((l) => l.id === license.id ? license : l);
        toast("Saved");
      } else {
        const { license } = await api("POST", "/licenses", payload);
        licenses.push(license);
        toast("License added");
      }
      closeEditDialog();
      render();
    } catch (err) {
      if (handleConflict(err, "Save")) return;
      showError("Save failed", err);
    }
  }
  // Open the shared row-action menu next to the clicked kebab button.
  function openRowMenu(lic, anchorEl) {
    const menu = $("rowMenu");
    menu.innerHTML = "";
    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.className = "danger";
    delBtn.textContent = "Delete this license";
    delBtn.addEventListener("click", () => {
      menu.hidden = true;
      deleteLicenseById(lic.id);
    });
    menu.appendChild(delBtn);
    // Position relative to the clicked anchor.
    const r = anchorEl.getBoundingClientRect();
    menu.style.top = `${window.scrollY + r.bottom + 4}px`;
    menu.style.left = `${window.scrollX + r.right - 160}px`;
    menu.hidden = false;
  }
  function closeRowMenu() { $("rowMenu").hidden = true; }
  document.addEventListener("click", (e) => {
    const menu = $("rowMenu");
    if (!menu || menu.hidden) return;
    if (!menu.contains(e.target)) closeRowMenu();
  });

  // Soft delete: remove from local state and show an Undo toast for 5s. Commit
  // to the server when the timer elapses or the user navigates away.
  function deleteLicense() {
    if (!editingId) return;
    deleteLicenseById(editingId);
    closeEditDialog();
  }
  function deleteLicenseById(id) {
    const lic = licenses.find((l) => l.id === id);
    if (!lic) return;
    licenses = licenses.filter((l) => l.id !== id);
    render();
    softDeleteWithUndo(lic);
  }

  function softDeleteWithUndo(lic) {
    if (pendingDeletes.has(lic.id)) clearTimeout(pendingDeletes.get(lic.id).timer);
    const t = $("toast");
    if (!t) return;
    t.innerHTML = "";
    const msg = document.createElement("span");
    msg.textContent = `Deleted ${lic.customer} - ${lic.licenseType}. `;
    t.appendChild(msg);
    const undo = document.createElement("button");
    undo.type = "button";
    undo.className = "toast-undo";
    undo.textContent = "Undo";
    undo.addEventListener("click", () => {
      const pending = pendingDeletes.get(lic.id);
      if (pending) clearTimeout(pending.timer);
      pendingDeletes.delete(lic.id);
      licenses.push(lic);
      render();
      t.hidden = true;
    });
    t.appendChild(undo);
    t.hidden = false;
    const timer = setTimeout(() => {
      // commit the delete
      api("DELETE", `/licenses/${lic.id}`).then(() => refreshTrashBadge()).catch((err) => {
        // If server delete fails, restore.
        showError("Delete failed (restored locally)", err);
        licenses.push(lic);
        render();
      });
      pendingDeletes.delete(lic.id);
      t.hidden = true;
    }, 8000); // v1.7.50: 5s -> 8s for Undo
    pendingDeletes.set(lic.id, { license: lic, timer });
  }

  // Commit any pending soft deletes immediately, e.g. before navigating away.
  function flushPendingDeletes() {
    for (const { license, timer } of pendingDeletes.values()) {
      clearTimeout(timer);
      api("DELETE", `/licenses/${license.id}`).catch(() => {});
    }
    pendingDeletes.clear();
  }
  window.addEventListener("beforeunload", flushPendingDeletes);

  // ---------- renew dialog ----------
  function openRenewDialog(id) {
    renewTargetId = id;
    const lic = licenses.find((l) => l.id === id);
    if (!lic) return;
    // Default custom date = current expiry advanced by the license's renewal cycle.
    const years = CYCLE_YEARS[lic.renewalCycle] || 1;
    if (lic.expiryDate) {
      const d = parseISO(lic.expiryDate);
      if (d) {
        d.setUTCFullYear(d.getUTCFullYear() + years);
        $("renewCustomDate").value = d.toISOString().slice(0, 10);
      }
    }
    $("renewDialog").showModal();
  }

  // ---------- email customer ----------

  function customerByName(name) {
    if (!name) return null;
    const norm = name.trim().toLowerCase();
    return customers.find((c) => (c.name || "").trim().toLowerCase() === norm) || null;
  }

  function templateFor(productLine) {
    if (productLine) {
      const t = emailTemplates.find((t) => (t.productLine || "").toLowerCase() === productLine.toLowerCase());
      if (t && (t.subject || t.body)) return t;
      // Fall through to seeded default for this product line
      if (DEFAULT_TEMPLATES[productLine]) return { productLine, ...DEFAULT_TEMPLATES[productLine] };
    }
    const def = emailTemplates.find((t) => (t.productLine || "") === "_default");
    return def || { productLine: "_default", ...DEFAULT_TEMPLATES._default };
  }

  function substitute(text, vars) {
    return String(text || "").replace(/\{(\w+)\}/g, (_, key) => vars[key] != null ? vars[key] : `{${key}}`);
  }

  // Compute disjoint renewal-package clusters once for the entire table.
  // Algorithm: for each customer, sort licenses by expiryDate, then walk in
  // order grouping consecutive licenses whose gap is ≤ 14 days. Each cluster
  // of size ≥ 2 becomes a bundle; every member sees the same size in their
  // Bundle: N badge. Abandoned and renewed rows are excluded.
  function computeBundles() {
    licenseBundles = new Map();
    const byCustomer = new Map();
    for (const lic of licenses) {
      if (lic.state === "abandoned" || lic.status === "renewed") continue;
      if (!lic.customer || !lic.expiryDate) continue;
      const key = lic.customer.trim().toLowerCase();
      if (!byCustomer.has(key)) byCustomer.set(key, []);
      byCustomer.get(key).push(lic);
    }
    const WINDOW = 14 * 86400000;
    let bundleIdCounter = 0;
    function emitCluster(cluster) {
      if (cluster.length < 2) return;
      bundleIdCounter++;
      const bundle = { id: bundleIdCounter, members: cluster.slice(), size: cluster.length };
      for (const lic of cluster) licenseBundles.set(lic.id, bundle);
    }
    for (const list of byCustomer.values()) {
      list.sort((a, b) => (a.expiryDate || "").localeCompare(b.expiryDate || ""));
      let cluster = [];
      let lastMs = null;
      for (const lic of list) {
        const ms = parseISO(lic.expiryDate)?.getTime() ?? null;
        if (ms === null) continue;
        if (lastMs !== null && ms - lastMs > WINDOW) {
          emitCluster(cluster);
          cluster = [];
        }
        cluster.push(lic);
        lastMs = ms;
      }
      if (cluster.length) emitCluster(cluster);
    }
  }

  // Returns the precomputed cluster members for this license, or [lic] if it's
  // not part of any bundle. Same result for every row in the same cluster.
  function bundleFor(lic) {
    if (!lic) return [lic];
    const bundle = licenseBundles.get(lic.id);
    return bundle ? bundle.members : [lic];
  }

  function emailCustomer(lic) {
    const ownerName = lic.ownerName || me.name || "";
    const cust = customerByName(lic.customer);
    const recipient = cust && cust.primaryEmail ? cust.primaryEmail : "";
    const cc = cust && Array.isArray(cust.secondaryEmails) && cust.secondaryEmails.length
      ? cust.secondaryEmails.join(",")
      : "";

    // Check for bundle: 2+ licenses for same customer within 14 days.
    const bundle = bundleFor(lic);
    const isBundle = bundle.length >= 2;

    let subject, body;
    if (isBundle) {
      subject = `Renewal reminder: ${lic.customer} multiple subscriptions expiring around ${fmtShortDate(lic.expiryDate)}`;
      const lines = bundle.map((b) => `  - ${b.licenseType}: ${b.userCount} users, expires ${fmtShortDate(b.expiryDate)}`).join("\n");
      const firstWord = (lic.customer || "").split(/[\s,]+/)[0] || "";
      body =
        `Hi ${firstWord} team,\n\n` +
        `Several of your subscriptions are coming up for renewal:\n\n` +
        lines + `\n\n` +
        `Please let us know if you would like to renew, adjust seat counts, or change any of these.\n\n` +
        `Looking forward to your reply.\n\n` +
        `Best regards,\n${ownerName}\nKation Technologies`;
    } else {
      const tpl = templateFor(lic.productLine);
      const firstWord = (lic.customer || "").split(/[\s,]+/)[0] || "";
      const vars = {
        customer: lic.customer || "",
        customerFirstWord: firstWord,
        licenseType: lic.licenseType || "",
        users: lic.userCount || 0,
        expiryDate: fmtShortDate(lic.expiryDate),
        ownerName,
      };
      subject = substitute(tpl.subject, vars);
      body = substitute(tpl.body, vars);
    }

    const params = [];
    if (cc) params.push(`cc=${encodeURIComponent(cc)}`);
    params.push(`subject=${encodeURIComponent(subject)}`);
    params.push(`body=${encodeURIComponent(body)}`);
    const url = `mailto:${encodeURIComponent(recipient)}?${params.join("&")}`;
    window.open(url, "_blank");

    // Per Ella's feedback 2026-06-17: do NOT auto-promote the license status to
    // Notice sent. mailto: opens an Outlook DRAFT, not a sent email -- the user
    // might never actually send it. The status should reflect what's been sent,
    // not what's been drafted. User must change the status manually in the edit
    // dialog once they've confirmed the email left their outbox.
    if (isBundle) toast(`Drafted ${bundle.length}-license bundle email in Outlook.`);
    else toast(`Drafted email in Outlook.`);
  }

  // ---------- people picker (v1.7.14) ----------
  // Searchable Entra-backed picker. Replaces the <select> dropdown.
  // Uses /api/users/search (User.Read.All app permission, consented tenant-wide).
  //
  // Each instance is a small object bound to a DOM container, with getValue /
  // setValue / focus methods. The picker handles its own DOM and event wiring.

  // Session-lifetime photo cache. Stores either a blob URL or `null` if the user
  // has no photo (204 from /api/users/.../photo).
  const photoCache = new Map(); // oid -> Promise<string|null>
  function loadPhoto(oid) {
    if (!oid) return Promise.resolve(null);
    if (photoCache.has(oid)) return photoCache.get(oid);
    const p = (async () => {
      try {
        const res = await fetch(`${API_BASE}/users/${encodeURIComponent(oid)}/photo`, {
          headers: { Authorization: `Bearer ${authToken}` },
        });
        if (res.status === 204 || !res.ok) return null;
        const blob = await res.blob();
        return URL.createObjectURL(blob);
      } catch (_) { return null; }
    })();
    photoCache.set(oid, p);
    return p;
  }

  function avatarInitials(displayName) {
    if (!displayName) return "?";
    const parts = String(displayName).trim().split(/\s+/);
    const first = (parts[0] || "")[0] || "";
    const last = parts.length > 1 ? (parts[parts.length - 1][0] || "") : "";
    return (first + last).toUpperCase() || "?";
  }

  // Render an avatar element. Photo loads async; falls back to initials.
  function renderAvatar(oid, displayName) {
    const wrap = document.createElement("div");
    wrap.className = "ppicker-avatar";
    wrap.style.background = ownerColor(oid);
    wrap.textContent = avatarInitials(displayName);
    loadPhoto(oid).then((url) => {
      if (!url || !wrap.isConnected) return;
      wrap.textContent = "";
      wrap.style.background = "none";
      const img = document.createElement("img");
      img.src = url;
      img.alt = displayName || "";
      wrap.appendChild(img);
    });
    return wrap;
  }

  function createPeoplePicker(container, options) {
    options = options || {};
    let selectedOid = options.initialOid || null;
    let selectedName = options.initialName || "";
    let results = [];
    let searching = false;
    let highlighted = -1;
    let debounceTimer = null;
    let lastQuery = "";

    container.classList.add("ppicker");

    function render() {
      container.innerHTML = "";
      if (selectedOid) {
        const chip = document.createElement("div");
        chip.className = "ppicker-chip";
        chip.appendChild(renderAvatar(selectedOid, selectedName));
        const nameSpan = document.createElement("span");
        nameSpan.className = "ppicker-name";
        nameSpan.textContent = selectedName || "(unnamed)";
        chip.appendChild(nameSpan);
        const clearBtn = document.createElement("button");
        clearBtn.type = "button";
        clearBtn.className = "ppicker-clear";
        clearBtn.setAttribute("aria-label", "Clear selection");
        clearBtn.textContent = "×";
        clearBtn.addEventListener("click", () => { clearSelection(); });
        chip.appendChild(clearBtn);
        container.appendChild(chip);
        return;
      }
      const input = document.createElement("input");
      input.type = "text";
      input.className = "ppicker-input";
      input.placeholder = options.placeholder || "Search teammates by name…";
      input.autocomplete = "off";
      input.addEventListener("input", () => onInput(input.value));
      input.addEventListener("keydown", (e) => onKeyDown(e));
      input.addEventListener("focus", () => { if (results.length) renderResults(); });
      container.appendChild(input);
      const dropdown = document.createElement("div");
      dropdown.className = "ppicker-results";
      dropdown.hidden = true;
      container.appendChild(dropdown);
    }

    function clearSelection() {
      selectedOid = null;
      selectedName = "";
      results = [];
      highlighted = -1;
      render();
      const input = container.querySelector(".ppicker-input");
      if (input) input.focus();
      if (options.onChange) options.onChange({ oid: null, name: "" });
    }

    function pick(user) {
      selectedOid = user.oid;
      selectedName = user.displayName || user.mail || "";
      results = [];
      highlighted = -1;
      render();
      if (options.onChange) options.onChange({ oid: selectedOid, name: selectedName });
    }

    function onInput(value) {
      const q = value.trim();
      lastQuery = q;
      if (debounceTimer) clearTimeout(debounceTimer);
      if (q.length < 2) {
        results = [];
        renderResults();
        return;
      }
      debounceTimer = setTimeout(() => doSearch(q), 220);
    }

    async function doSearch(q) {
      searching = true;
      renderResults();
      try {
        const { users } = await api("GET", `/users/search?q=${encodeURIComponent(q)}`);
        if (q !== lastQuery) return; // stale response
        results = users || [];
        highlighted = results.length ? 0 : -1;
      } catch (err) {
        console.warn("user search failed", err);
        results = [];
      } finally {
        searching = false;
        renderResults();
      }
    }

    function renderResults() {
      const dropdown = container.querySelector(".ppicker-results");
      if (!dropdown) return;
      dropdown.innerHTML = "";
      if (searching) {
        const loading = document.createElement("div");
        loading.className = "ppicker-loading";
        loading.textContent = "Searching…";
        dropdown.appendChild(loading);
        dropdown.hidden = false;
        return;
      }
      if (!results.length) {
        if (lastQuery.length >= 2) {
          const none = document.createElement("div");
          none.className = "ppicker-loading";
          none.textContent = "No matches.";
          dropdown.appendChild(none);
          dropdown.hidden = false;
        } else {
          dropdown.hidden = true;
        }
        return;
      }
      results.forEach((u, i) => {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "ppicker-result" + (i === highlighted ? " highlighted" : "");
        btn.appendChild(renderAvatar(u.oid, u.displayName));
        const txt = document.createElement("div");
        txt.className = "ppicker-result-text";
        const nm = document.createElement("div");
        nm.className = "ppicker-result-name";
        nm.textContent = u.displayName;
        txt.appendChild(nm);
        if (u.jobTitle || u.mail) {
          const sub = document.createElement("div");
          sub.className = "ppicker-result-sub";
          sub.textContent = u.jobTitle || u.mail || "";
          txt.appendChild(sub);
        }
        btn.appendChild(txt);
        // mousedown fires before blur on the input, so we can pick without the
        // input losing focus first. preventDefault keeps the focus on the input
        // during the synchronous render swap.
        btn.addEventListener("mousedown", (e) => { e.preventDefault(); pick(u); });
        // Update the highlighted class WITHOUT rebuilding the dropdown -- a full
        // re-render here destroys the button mid-click and was eating the pick.
        btn.addEventListener("mouseenter", () => {
          if (highlighted === i) return;
          highlighted = i;
          dropdown.querySelectorAll(".ppicker-result").forEach((el, idx) => {
            el.classList.toggle("highlighted", idx === i);
          });
        });
        dropdown.appendChild(btn);
      });
      dropdown.hidden = false;
    }

    function onKeyDown(e) {
      if (e.key === "ArrowDown") {
        e.preventDefault();
        if (results.length) { highlighted = Math.min(highlighted + 1, results.length - 1); renderResults(); }
      } else if (e.key === "ArrowUp") {
        e.preventDefault();
        if (results.length) { highlighted = Math.max(highlighted - 1, 0); renderResults(); }
      } else if (e.key === "Enter") {
        if (highlighted >= 0 && results[highlighted]) { e.preventDefault(); pick(results[highlighted]); }
      } else if (e.key === "Escape") {
        results = [];
        highlighted = -1;
        renderResults();
      }
    }

    render();
    return {
      getValue: () => ({ oid: selectedOid, name: selectedName }),
      setValue: (oid, name) => {
        selectedOid = oid || null;
        selectedName = name || "";
        render();
      },
      focus: () => {
        const input = container.querySelector(".ppicker-input");
        if (input) input.focus();
      },
    };
  }

  // ---------- bulk select + reassign ----------
  function setBulkMode(on) {
    bulkMode = on;
    $("bulkSelectBtn").setAttribute("aria-pressed", String(on));
    $("bulkSelectBtn").textContent = on ? "Done selecting" : "Select";
    if (!on) bulkSelected.clear();
    updateBulkBar();
    render();
  }
  function updateBulkBar() {
    const n = bulkSelected.size;
    $("bulkBar").hidden = !(bulkMode && n > 0);
    $("bulkBarCount").textContent = n;
  }
  function openBulkReassign() {
    const n = bulkSelected.size;
    if (!n) return;
    $("bulkReassignCount").textContent = n;
    if (!bulkReassignPicker) {
      bulkReassignPicker = createPeoplePicker($("bulkReassignOwner"), {
        placeholder: "Search teammates by name…",
      });
    } else {
      bulkReassignPicker.setValue(null, "");
    }
    $("bulkReassignDialog").showModal();
  }
  async function confirmBulkReassign() {
    const picked = bulkReassignPicker ? bulkReassignPicker.getValue() : { oid: null, name: null };
    const oid = picked.oid;
    if (!oid) { toast("Pick an owner first"); return; }
    const ids = [...bulkSelected];
    try {
      const { updated } = await api("POST", "/licenses/bulk", {
        ids,
        patch: { ownerOid: oid, ownerName: picked.name || "" },
      });
      const m = new Map(updated.map((l) => [l.id, l]));
      licenses = licenses.map((l) => m.get(l.id) || l);
      $("bulkReassignDialog").close();
      bulkSelected.clear();
      setBulkMode(false);
      toast(`Reassigned ${updated.length} licenses to ${picked.name || "new owner"}.`);
    } catch (err) {
      showError("Reassign failed", err);
    }
  }

  // v1.7.41 — bulk Renew +1y. Loops through the per-id renew endpoint
  // (no bulk-renew route on the server) and shows aggregate progress.
  // If-Match guards each call individually, so a stale row is reported
  // without aborting the whole batch.
  async function bulkRenewSelected() {
    const ids = [...bulkSelected];
    if (!ids.length) return;
    if (!confirm(`Renew +1y on ${ids.length} licenses? (Each row's lastEditedAt is used as a guard against in-flight edits.)`)) return;
    let ok = 0, conflicts = 0, failed = 0;
    for (const id of ids) {
      const lic = licenses.find((l) => l.id === id);
      if (!lic) { failed++; continue; }
      try {
        const ifMatch = lic.lastEditedAt || null;
        const { license } = await api("POST", `/licenses/${id}/renew`, { years: 1 }, { ifMatch });
        licenses = licenses.map((l) => l.id === license.id ? license : l);
        ok++;
      } catch (err) {
        if (err && err.status === 409) {
          if (err.license) licenses = licenses.map((l) => l.id === err.license.id ? err.license : l);
          conflicts++;
        } else failed++;
      }
    }
    bulkSelected.clear();
    setBulkMode(false);
    render();
    const parts = [`${ok} renewed`];
    if (conflicts) parts.push(`${conflicts} skipped (edited by someone else)`);
    if (failed) parts.push(`${failed} failed`);
    toast(parts.join(", "));
  }

  // v1.7.41 — bulk Export selected. Uses the same CSV builder as the
  // toolbar's exportCsv() but only over the bulk-selected ids.
  function bulkExportSelected() {
    const ids = bulkSelected;
    if (!ids.size) return;
    const list = sortLicenses(licenses.filter((l) => ids.has(l.id)));
    exportCsvList(list, `selected-licenses-${todayPh()}.csv`);
    toast(`Exported ${list.length} selected license${list.length === 1 ? "" : "s"}`);
  }

  // v1.7.41 — bulk Delete selected with one combined Undo. Optimistic UI:
  // remove from local state immediately, commit per-row after 6s. Click Undo
  // before the timer to roll the batch back in one toast.
  function bulkDeleteSelected() {
    const ids = [...bulkSelected];
    if (!ids.length) return;
    if (!confirm(`Delete ${ids.length} license${ids.length === 1 ? "" : "s"}? You can Undo for 6 seconds.`)) return;
    const removed = licenses.filter((l) => ids.includes(l.id));
    licenses = licenses.filter((l) => !ids.includes(l.id));
    bulkSelected.clear();
    setBulkMode(false);
    render();
    let undone = false;
    const timer = setTimeout(async () => {
      if (undone) return;
      let failed = 0;
      for (const lic of removed) {
        try { await api("DELETE", `/licenses/${lic.id}`); } catch { failed++; }
      }
      if (failed) toast(`Delete failed for ${failed} of ${removed.length}`);
    }, 8000); // v1.7.50: 6s -> 8s for bulk Undo
    toast(`Deleted ${removed.length} license${removed.length === 1 ? "" : "s"}`, {
      actionLabel: "Undo",
      onAction: () => {
        undone = true;
        clearTimeout(timer);
        licenses.push(...removed);
        render();
      },
    });
  }

  // v1.7.41 — customer merge. Shows every other customer with at least 3 chars
  // of name overlap with the current one, lets the user pick which are dupes,
  // and POSTs the merge. Cures the "Beyond Innovations" vs "Beyond Innovations,
  // Inc" inflation in counts.
  let mergeTargetCustomer = null;
  function openMergeDialog(targetName) {
    if (!targetName) return;
    mergeTargetCustomer = targetName;
    $("mergeTargetName").textContent = targetName;
    const list = $("mergeCandidateList");
    list.innerHTML = "";
    const targetNorm = targetName.trim().toLowerCase();
    const targetTokens = new Set(targetNorm.split(/[\s,.&\-]+/).filter((t) => t.length >= 3));
    // Build candidate set from BOTH customer rows + license customer strings,
    // so even a customer that doesn't have a registry row yet can be merged in.
    const candidates = new Map(); // norm -> { name, source }
    for (const c of customers) {
      const k = (c.name || "").trim().toLowerCase();
      if (!k || k === targetNorm) continue;
      const tokens = k.split(/[\s,.&\-]+/);
      if (tokens.some((t) => t.length >= 3 && targetTokens.has(t))) {
        candidates.set(k, { name: c.name, hasRegistry: true });
      }
    }
    for (const lic of licenses) {
      const name = (lic.customer || "").trim();
      const k = name.toLowerCase();
      if (!k || k === targetNorm || candidates.has(k)) continue;
      const tokens = k.split(/[\s,.&\-]+/);
      if (tokens.some((t) => t.length >= 3 && targetTokens.has(t))) {
        candidates.set(k, { name, hasRegistry: false });
      }
    }
    if (!candidates.size) {
      const li = document.createElement("li");
      li.className = "merge-empty";
      li.textContent = "No obvious duplicates. Use the search box if you have a specific one in mind.";
      list.appendChild(li);
    } else {
      for (const { name, hasRegistry } of candidates.values()) {
        const li = document.createElement("li");
        const lbl = document.createElement("label");
        lbl.className = "merge-candidate-row";
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.value = name;
        const txt = document.createElement("span");
        const licCount = licenses.filter((l) => (l.customer || "").trim().toLowerCase() === name.toLowerCase()).length;
        txt.textContent = `${name}  ·  ${licCount} license${licCount === 1 ? "" : "s"}${hasRegistry ? "" : "  ·  (no registry entry)"}`;
        lbl.appendChild(cb);
        lbl.appendChild(txt);
        li.appendChild(lbl);
        list.appendChild(li);
      }
    }
    $("customerMergeDialog").showModal();
  }
  async function confirmMerge() {
    if (!mergeTargetCustomer) return;
    const sources = [...$("mergeCandidateList").querySelectorAll("input[type=checkbox]:checked")].map((cb) => cb.value);
    if (!sources.length) { toast("Pick at least one duplicate to merge"); return; }
    if (!confirm(`Merge ${sources.length} customer row${sources.length === 1 ? "" : "s"} into "${mergeTargetCustomer}"? Every license under those names will be updated. This can't be undone.`)) return;
    const btn = $("mergeConfirmBtn");
    btn.disabled = true;
    btn.textContent = "Merging…";
    try {
      const res = await api("POST", "/customers/merge", {
        sourceNames: sources,
        targetName: mergeTargetCustomer,
      });
      // Hard refresh — merge touches lots of rows + the customer registry.
      await refreshAll();
      $("customerMergeDialog").close();
      toast(`Merged ${res.customersDeleted} duplicate${res.customersDeleted === 1 ? "" : "s"}, updated ${res.licensesUpdated} license${res.licensesUpdated === 1 ? "" : "s"}`);
    } catch (err) {
      showError("Merge failed", err);
    } finally {
      btn.disabled = false;
      btn.textContent = "Merge selected";
    }
  }

  // ---------- customer 360 drawer ----------
  function openCustomerDialog(customer) {
    if (!customer) return;
    const norm = customer.trim().toLowerCase();
    const rows = licenses.filter((l) => (l.customer || "").trim().toLowerCase() === norm);
    if (!rows.length) return;
    $("customerDialogTitle").textContent = customer;
    const seats = rows.reduce((s, l) => s + (l.userCount || 0), 0);
    const active = rows.filter((l) => l.state !== "abandoned" && l.status !== "renewed").length;
    const renewedCount = rows.filter((l) => l.status === "renewed").length;
    const today = todayPh();
    const nextExpiry = rows
      .filter((l) => l.state !== "abandoned" && l.expiryDate && l.expiryDate >= today)
      .map((l) => l.expiryDate)
      .sort()[0];
    const summary = $("customerDialogSummary");
    summary.innerHTML = "";
    function stat(label, val) {
      const s = document.createElement("div");
      s.innerHTML = `<strong>${val}</strong>${label}`;
      summary.appendChild(s);
    }
    stat(" licenses", rows.length);
    stat(" total seats", seats.toLocaleString());
    stat(" open", active);
    stat(" renewed", renewedCount);
    if (nextExpiry) stat(" next expiry", fmtShortDate(nextExpiry));

    // Show customer registry info (primary email, secondary, address, notes) if present.
    const cust = customerByName(customer);
    const info = $("customerRegistryInfo");
    info.innerHTML = "";
    if (cust && (cust.primaryEmail || (cust.secondaryEmails && cust.secondaryEmails.length) || cust.address || cust.notes)) {
      info.hidden = false;
      function infoRow(label, value) {
        if (!value) return;
        const row = document.createElement("div");
        row.className = "registry-row";
        const l = document.createElement("span");
        l.className = "registry-label";
        l.textContent = label;
        const v = document.createElement("span");
        v.textContent = value;
        row.appendChild(l);
        row.appendChild(v);
        info.appendChild(row);
      }
      infoRow("Email:", cust.primaryEmail);
      if (Array.isArray(cust.secondaryEmails) && cust.secondaryEmails.length) infoRow("CC:", cust.secondaryEmails.join(", "));
      infoRow("Address:", cust.address);
      infoRow("Notes:", cust.notes);
    } else {
      info.hidden = true;
    }
    // Wire Edit Customer button to open the customer edit dialog.
    $("customerEditBtn").onclick = () => openCustomerEditDialog(customer);
    // v1.7.41 — Merge button.
    $("customerMergeBtn").onclick = () => openMergeDialog(customer);

    const ul = $("customerDialogList");
    ul.innerHTML = "";
    const sorted = [...rows].sort((a, b) => (a.expiryDate || "").localeCompare(b.expiryDate || ""));
    for (const lic of sorted) {
      const li = document.createElement("li");
      const body = document.createElement("div");
      body.className = "cl-body";
      const title = document.createElement("div");
      title.className = "cl-title";
      title.textContent = lic.licenseType;
      body.appendChild(title);
      const meta = document.createElement("div");
      meta.className = "cl-meta";
      meta.textContent = `${lic.userCount || 0} users · expires ${fmtShortDate(lic.expiryDate)} · owner ${lic.ownerName || "—"}`;
      body.appendChild(meta);
      li.appendChild(body);
      const pill = document.createElement("span");
      const sv = lic.status || "notStarted";
      pill.className = `status-pill status-${sv}`;
      pill.textContent = STATUS_LABEL[sv] || sv;
      li.appendChild(pill);
      li.addEventListener("click", () => { $("customerDialog").close(); openEditDialog(lic); });
      ul.appendChild(li);
    }
    $("customerDialog").showModal();
  }
  async function confirmRenew(years, customDate) {
    if (!renewTargetId) return;
    const body = customDate ? { newExpiryDate: customDate } : { years };
    try {
      const current = licenses.find((l) => l.id === renewTargetId);
      const ifMatch = current && current.lastEditedAt ? current.lastEditedAt : null;
      const { license } = await api("POST", `/licenses/${renewTargetId}/renew`, body, { ifMatch });
      licenses = licenses.map((l) => l.id === license.id ? license : l);
      $("renewDialog").close();
      // If the edit dialog is open for this license, refresh its expiry field.
      if (editingId === license.id) {
        $("licExpiry").value = license.expiryDate;
      }
      renewTargetId = null;
      render();
      toast(`Renewed — new expiry ${fmtShortDate(license.expiryDate)}`);
    } catch (err) {
      if (handleConflict(err, "Renew")) return;
      showError("Renewal failed", err);
    }
  }

  // v1.7.39 — inline-row +1y renew. Optimistic UI + Undo via toast.
  async function quickRenewOneYear(lic) {
    const before = { ...lic, comments: lic.comments ? [...lic.comments] : [] };
    try {
      const ifMatch = lic.lastEditedAt || null;
      const { license } = await api("POST", `/licenses/${lic.id}/renew`, { years: 1 }, { ifMatch });
      licenses = licenses.map((l) => l.id === license.id ? license : l);
      render();
      toast(`Renewed: ${license.customer}, ${license.licenseType}`, {
        actionLabel: "Undo",
        onAction: async () => {
          try {
            await api("PATCH", `/licenses/${lic.id}`, {
              expiryDate: before.expiryDate,
              status: before.status,
            });
            const refreshed = await api("GET", "/licenses");
            licenses = refreshed.licenses || licenses;
            render();
            toast("Renewal undone");
          } catch (err) { showError("Undo failed", err); }
        },
      });
    } catch (err) {
      if (handleConflict(err, "Renew")) return;
      showError("Renewal failed", err);
    }
  }

  // ---------- day dialog (calendar overflow) ----------
  function openDayDialog(isoDate, items) {
    $("dayDialogTitle").textContent = fmtShortDate(isoDate);
    const ul = $("dayDialogList");
    ul.innerHTML = "";
    for (const lic of items) {
      const li = document.createElement("li");
      const pill = document.createElement("span");
      pill.className = "owner-pill small";
      pill.style.background = ownerColor(lic.ownerOid);
      pill.textContent = lic.ownerName || "(no owner)";
      li.appendChild(pill);
      const text = document.createElement("span");
      text.textContent = ` ${lic.customer} · ${lic.licenseType} (${lic.userCount} users)`;
      li.appendChild(text);
      li.style.cursor = "pointer";
      li.addEventListener("click", () => { $("dayDialog").close(); openEditDialog(lic); });
      ul.appendChild(li);
    }
    $("dayDialogAdd").onclick = () => { $("dayDialog").close(); openAddDialogForDate(isoDate); };
    $("dayDialog").showModal();
  }

  // ---------- v1.8.0 product lines admin dialog ----------
  let plPendingPreview = []; // [{id, customer, from, to}]
  async function openProductLinesDialog() {
    // Re-fetch in case another tab added entries since the cached load.
    try {
      const { productLines } = await api("GET", "/product-lines");
      productLinesRegistry = productLines || [];
    } catch (_) { /* fall back to in-memory */ }
    renderProductLinesList();
    renderLegacyMappingRows();
    plPendingPreview = [];
    $("plPreviewSummary").hidden = true;
    $("plNormalizeApplyBtn").disabled = true;
    $("productLinesDialog").showModal();
  }
  function renderProductLinesList() {
    const ul = $("plList");
    if (!ul) return;
    ul.innerHTML = "";
    const usage = countProductLineUsage();
    for (const pl of productLinesRegistry) {
      const li = document.createElement("li");
      li.className = "pl-item";
      const nameInput = document.createElement("input");
      nameInput.type = "text";
      nameInput.value = pl.name;
      nameInput.className = "pl-name-input";
      nameInput.maxLength = 100;
      nameInput.setAttribute("aria-label", `Product line name (${pl.name})`);
      const count = document.createElement("span");
      count.className = "pl-usage";
      const n = usage.get(pl.name) || 0;
      count.textContent = n === 1 ? "1 license" : `${n} licenses`;
      const saveBtn = document.createElement("button");
      saveBtn.type = "button";
      saveBtn.className = "btn ghost small";
      saveBtn.textContent = "Rename";
      saveBtn.addEventListener("click", () => renameProductLine(pl.name, nameInput.value.trim()));
      const delBtn = document.createElement("button");
      delBtn.type = "button";
      delBtn.className = "btn ghost small danger";
      delBtn.textContent = "Delete";
      delBtn.addEventListener("click", () => deleteProductLineEntry(pl.name, n));
      li.appendChild(nameInput);
      li.appendChild(count);
      li.appendChild(saveBtn);
      li.appendChild(delBtn);
      ul.appendChild(li);
    }
  }
  function countProductLineUsage() {
    const m = new Map();
    for (const lic of licenses) {
      const v = String(lic.productLine || "").trim();
      if (!v) continue;
      m.set(v, (m.get(v) || 0) + 1);
    }
    return m;
  }
  function legacyValues() {
    const known = new Set(productLinesRegistry.map((p) => p.name));
    const set = new Set();
    for (const lic of licenses) {
      const v = String(lic.productLine || "").trim();
      if (v && !known.has(v)) set.add(v);
    }
    return [...set].sort();
  }
  function renderLegacyMappingRows() {
    const empty = $("plLegacyEmpty");
    const ul = $("plLegacyList");
    if (!empty || !ul) return;
    ul.innerHTML = "";
    const vals = legacyValues();
    if (!vals.length) {
      empty.hidden = false;
      ul.hidden = true;
      return;
    }
    empty.hidden = true;
    ul.hidden = false;
    for (const v of vals) {
      const li = document.createElement("li");
      li.className = "pl-legacy-row";
      const from = document.createElement("span");
      from.className = "pl-legacy-from";
      from.textContent = v;
      const arrow = document.createElement("span");
      arrow.className = "pl-legacy-arrow";
      arrow.textContent = "→";
      const toSel = document.createElement("select");
      toSel.className = "pl-legacy-to";
      toSel.dataset.legacyFrom = v;
      toSel.setAttribute("aria-label", `Map ${v} to`);
      const skip = document.createElement("option");
      skip.value = "";
      skip.textContent = "(skip)";
      toSel.appendChild(skip);
      // Auto-pick a sensible default when the legacy value contains a known
      // canonical name (e.g. "ERP - BC" → "Business Central").
      const suggestion = autoSuggestCanonical(v);
      for (const pl of productLinesRegistry) {
        const opt = document.createElement("option");
        opt.value = pl.name;
        opt.textContent = pl.name;
        if (pl.name === suggestion) opt.selected = true;
        toSel.appendChild(opt);
      }
      li.appendChild(from);
      li.appendChild(arrow);
      li.appendChild(toSel);
      ul.appendChild(li);
    }
  }
  function autoSuggestCanonical(legacy) {
    const v = (legacy || "").toLowerCase();
    // Heuristics for the known Kation legacy values. Returns "" if no guess.
    if (/\bbc\b|business\s*central/.test(v)) return "Business Central";
    if (/f&o|f and o|finance/.test(v)) return "Finance and Operation";
    if (/philtax/.test(v)) return "PHILTAX";
    if (/m365|microsoft\s*365|office\s*365/.test(v)) return "M365";
    if (/crm|dynamics\s*365\s*sales/.test(v)) return "CRM";
    if (/security|defender/.test(v)) return "Security";
    return "";
  }
  function gatherMappingFromUI() {
    const map = {};
    document.querySelectorAll("#plLegacyList .pl-legacy-to").forEach((sel) => {
      const from = sel.dataset.legacyFrom;
      const to = sel.value;
      if (from && to) map[from] = to;
    });
    return map;
  }
  async function previewNormalize() {
    const mapping = gatherMappingFromUI();
    if (!Object.keys(mapping).length) {
      toast("Pick a canonical value for at least one legacy entry.");
      return;
    }
    try {
      const res = await api("POST", "/licenses/normalize-product-lines", { mapping, dryRun: true });
      plPendingPreview = res.preview || [];
      const sum = $("plPreviewSummary");
      sum.textContent = `${plPendingPreview.length} row${plPendingPreview.length === 1 ? "" : "s"} will change. Click Apply to commit.`;
      sum.hidden = false;
      $("plNormalizeApplyBtn").disabled = plPendingPreview.length === 0;
    } catch (err) {
      showError("Preview failed", err);
    }
  }
  async function applyNormalize() {
    const mapping = gatherMappingFromUI();
    if (!Object.keys(mapping).length) return;
    try {
      const res = await api("POST", "/licenses/normalize-product-lines", { mapping, dryRun: false });
      toast(`Normalized ${res.count} row${res.count === 1 ? "" : "s"}.`);
      // Re-fetch licenses so the table reflects the new values immediately.
      const { licenses: lics } = await api("GET", "/licenses");
      licenses = lics || [];
      lastLicensesSig = licensesSignature(licenses);
      renderProductLinesList();
      renderLegacyMappingRows();
      plPendingPreview = [];
      $("plPreviewSummary").hidden = true;
      $("plNormalizeApplyBtn").disabled = true;
      render();
    } catch (err) {
      showError("Apply failed", err);
    }
  }
  async function addProductLineFromInput() {
    const input = $("plAddInput");
    const name = (input.value || "").trim().slice(0, 100);
    if (!name) return;
    if (productLinesRegistry.some((p) => p.name === name)) {
      toast("Already exists.");
      return;
    }
    const next = [...productLinesRegistry, { name, sortOrder: productLinesRegistry.length }];
    try {
      const res = await api("PUT", "/product-lines", { productLines: next });
      productLinesRegistry = res.productLines || next;
      input.value = "";
      // v1.8.2 — Toast for consistency with Rename / Delete actions.
      toast(`Added: ${name}`);
      renderProductLinesList();
      renderLegacyMappingRows();
    } catch (err) {
      showError("Add failed", err);
    }
  }
  async function renameProductLine(oldName, newName) {
    if (!newName || newName === oldName) return;
    if (productLinesRegistry.some((p) => p.name === newName && p.name !== oldName)) {
      toast("Name already in use.");
      return;
    }
    const next = productLinesRegistry.map((p) => p.name === oldName ? { ...p, name: newName } : p);
    try {
      const res = await api("PUT", "/product-lines", { productLines: next });
      productLinesRegistry = res.productLines || next;
      // v1.8.2 — Short, clear toast fired BEFORE the list re-render so the
      // user can't miss it. The "still referenced by licenses" hint moved
      // to the legacy panel below (which will now show the old name as a
      // legacy value with an auto-suggested target of the new name).
      const stillUsed = countProductLineUsage().get(oldName) || 0;
      const tail = stillUsed > 0 ? ` (${stillUsed} license${stillUsed === 1 ? "" : "s"} now legacy)` : "";
      toast(`Renamed: ${oldName} → ${newName}${tail}`);
      renderProductLinesList();
      renderLegacyMappingRows();
    } catch (err) {
      showError("Rename failed", err);
    }
  }
  async function deleteProductLineEntry(name, usageCount) {
    if (usageCount > 0) {
      toast(`Cannot delete — ${usageCount} license${usageCount === 1 ? " uses" : "s use"} it. Re-map them first.`);
      return;
    }
    if (!confirm(`Delete product line "${name}"?`)) return;
    try {
      await api("DELETE", `/product-lines/${encodeURIComponent(name)}`);
      productLinesRegistry = productLinesRegistry.filter((p) => p.name !== name);
      toast(`Deleted: ${name}`);
      renderProductLinesList();
      renderLegacyMappingRows();
    } catch (err) {
      showError("Delete failed", err);
    }
  }

  // ---------- event wiring ----------
  function wireEvents() {
    // View switcher (Table / Calendar)
    document.querySelectorAll("#viewSwitch .view-btn").forEach((b) => {
      b.addEventListener("click", () => {
        currentView = b.dataset.view;
        try { localStorage.setItem(LS_VIEW, currentView); } catch (_) {}
        render();
      });
    });

    // Group switcher (None / Customer / Owner / Product line)
    document.querySelectorAll("#groupSwitch .view-btn").forEach((b) => {
      b.addEventListener("click", () => {
        groupBy = b.dataset.group;
        try { localStorage.setItem(LS_GROUP, groupBy); } catch (_) {}
        render();
      });
    });

    // Summary chips toggle filter
    document.querySelectorAll(".lic-summary-chip").forEach((b) => {
      b.addEventListener("click", () => {
        const f = b.dataset.summary;
        summaryFilter = summaryFilter === f ? null : f;
        render();
      });
    });

    // Quick filter pills
    document.querySelectorAll(".filter-pill").forEach((p) => {
      p.addEventListener("click", () => {
        quickFilter = p.dataset.quick;
        summaryFilter = null;
        // v1.8.2 — date quick-filters (month/quarter) clear the corresponding
        // dropdowns so the user never lands on an empty intersection like
        // "This quarter (Q2)" + "Quarter dropdown = Q3" → 0 rows.
        if (quickFilter === "month" || quickFilter === "quarter") {
          monthFilter = "";
          quarterFilter = "";
          const mf = $("monthFilter"); if (mf) mf.value = "";
          const qf = $("quarterFilter"); if (qf) qf.value = "";
          persistFilters();
        }
        try { localStorage.setItem(LS_QUICK, quickFilter); } catch (_) {}
        render();
      });
    });

    // v1.7.38 Month dropdown
    $("monthFilter").addEventListener("change", (e) => {
      monthFilter = e.target.value || "";
      // v1.8.2 — MONTH and QUARTER are mutually exclusive in BOTH directions.
      // Picking a specific month also clears the QUARTER dropdown + the
      // "This quarter" quick-filter pill so the user can never end up with
      // an empty intersection of date filters.
      if (monthFilter) {
        quarterFilter = "";
        const qs = $("quarterFilter");
        if (qs) qs.value = "";
        if (quickFilter === "quarter") quickFilter = "all";
      }
      persistFilters();
      // Calendar view: jump the calendar to the selected month for parity
      // (Table just filters in-place). Hands the user a single mental model.
      if (monthFilter) {
        const [y, m] = monthFilter.split("-").map(Number);
        calCursor = new Date(y, m - 1, 1);
      }
      render();
    });

    // v1.7.38 Expiry filter popover
    const expBtn = $("expiryFilterBtn");
    const expMenu = $("expiryFilterMenu");
    expBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      const open = !expMenu.hidden;
      expMenu.hidden = open;
      expBtn.setAttribute("aria-expanded", String(!open));
    });
    document.addEventListener("click", (e) => {
      if (!expMenu.hidden && !expMenu.contains(e.target) && e.target !== expBtn && !expBtn.contains(e.target)) {
        expMenu.hidden = true;
        expBtn.setAttribute("aria-expanded", "false");
      }
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && !expMenu.hidden) {
        expMenu.hidden = true;
        expBtn.setAttribute("aria-expanded", "false");
        expBtn.focus();
      }
    });

    // v1.7.38 Clear-all-filters button
    $("clearAllFiltersBtn").addEventListener("click", clearAllFilters);

    // v1.7.42 Date-range popover
    const drBtn = $("dateRangeBtn");
    const drMenu = $("dateRangeMenu");
    drBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      const open = !drMenu.hidden;
      drMenu.hidden = open;
      drBtn.setAttribute("aria-expanded", String(!open));
    });
    document.addEventListener("click", (e) => {
      if (!drMenu.hidden && !drMenu.contains(e.target) && e.target !== drBtn && !drBtn.contains(e.target)) {
        drMenu.hidden = true;
        drBtn.setAttribute("aria-expanded", "false");
      }
    });
    $("dateFromFilter").addEventListener("change", (e) => {
      dateFromFilter = e.target.value || "";
      persistFilters(); render();
    });
    $("dateToFilter").addEventListener("change", (e) => {
      dateToFilter = e.target.value || "";
      persistFilters(); render();
    });
    $("dateRangeClear").addEventListener("click", () => {
      dateFromFilter = ""; dateToFilter = "";
      $("dateFromFilter").value = ""; $("dateToFilter").value = "";
      persistFilters(); render();
    });

    // v1.7.39 Save view
    $("saveViewBtn").addEventListener("click", saveCurrentView);

    // v1.7.39 Comments
    $("licCommentSendBtn").addEventListener("click", sendComment);
    $("licCommentInput").addEventListener("keydown", (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "Enter") {
        e.preventDefault();
        sendComment();
      }
    });

    // v1.7.39 Quick-add inline row
    const qaForm = $("quickAddForm");
    if (qaForm) {
      $("qaExpiry").value = todayPh();
      qaForm.addEventListener("submit", async (e) => {
        e.preventDefault();
        const payload = {
          customer: $("qaCustomer").value.trim(),
          licenseType: $("qaLicenseType").value.trim(),
          expiryDate: $("qaExpiry").value,
          userCount: parseInt($("qaUsers").value, 10) || 0,
          ownerOid: me.oid,
          ownerName: me.name,
        };
        if (!payload.customer || !payload.licenseType || !payload.expiryDate) {
          toast("Customer, license type and expiry are required");
          return;
        }
        try {
          const { license } = await api("POST", "/licenses", payload);
          licenses = [...licenses, license];
          $("qaCustomer").value = "";
          $("qaLicenseType").value = "";
          $("qaUsers").value = "";
          $("qaExpiry").value = todayPh();
          $("qaCustomer").focus();
          render();
          toast(`Added: ${license.customer}, ${license.licenseType}`);
        } catch (err) { showError("Add failed", err); }
      });
      $("qaCancelBtn").addEventListener("click", () => {
        $("quickAddDetails").open = false;
        $("qaCustomer").value = "";
        $("qaLicenseType").value = "";
      });
    }

    // Search — v1.7.40 debounced to 150ms so render doesn't fire every keystroke
    // when the user is typing fast in a tab with hundreds of licenses.
    const searchInput = $("searchInput");
    let searchDebounce = null;
    searchInput.addEventListener("input", () => {
      $("searchClear").hidden = !searchInput.value.trim();
      clearTimeout(searchDebounce);
      searchDebounce = setTimeout(() => {
        searchText = searchInput.value.trim();
        render();
      }, 150);
    });
    $("searchClear").addEventListener("click", () => {
      searchInput.value = "";
      searchText = "";
      $("searchClear").hidden = true;
      render();
      searchInput.focus();
    });

    // Sort
    document.querySelectorAll("th.sortable").forEach((th) => {
      th.addEventListener("click", () => {
        const k = th.dataset.sort;
        if (sortKey === k) sortDir = -sortDir;
        else { sortKey = k; sortDir = 1; }
        try { localStorage.setItem(LS_SORT, `${sortKey}:${sortDir}`); } catch (_) {}
        render();
      });
    });

    // Add license buttons
    $("addLicenseBtn").addEventListener("click", openAddDialog);
    $("licEmptyAdd").addEventListener("click", openAddDialog);

    // Manual refresh
    $("refreshBtn").addEventListener("click", refreshAll);

    // Unified CSV menu (Import / Export / Download template)
    const csvMenuBtn = $("csvMenuBtn");
    const csvMenu = $("csvMenu");
    function closeCsvMenu() {
      csvMenu.hidden = true;
      csvMenuBtn.setAttribute("aria-expanded", "false");
    }
    csvMenuBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      const wasHidden = csvMenu.hidden;
      csvMenu.hidden = !wasHidden;
      csvMenuBtn.setAttribute("aria-expanded", String(wasHidden));
    });
    document.addEventListener("click", (e) => {
      if (!csvMenu.contains(e.target) && e.target !== csvMenuBtn) closeCsvMenu();
    });
    document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeCsvMenu(); });
    $("importCsvBtn").addEventListener("click", () => { closeCsvMenu(); openImportDialog(); });
    $("exportCsvBtn").addEventListener("click", () => { closeCsvMenu(); exportCsv(); });
    $("downloadTemplateBtn").addEventListener("click", () => { closeCsvMenu(); downloadCsvTemplate(); });
    $("licEmptyImport").addEventListener("click", openImportDialog);
    $("downloadTemplateBtnInDialog").addEventListener("click", downloadCsvTemplate);

    // License-tab Settings
    $("licSettingsBtn").addEventListener("click", openSettingsDialog);
    $("setSaveBtn").addEventListener("click", saveLicSettings);
    $("setCancelBtn").addEventListener("click", () => $("licSettingsDialog").close());
    $("setSendTestDigestBtn").addEventListener("click", () => sendTestDigest(false));
    $("setSendTestDigestAllBtn").addEventListener("click", () => sendTestDigest(true));
    // v1.7.43 Recovery + Privacy
    $("setOpenRecoveryBtn").addEventListener("click", openRecoveryDialog);
    $("recoveryCloseBtn").addEventListener("click", () => $("recoveryDialog").close());
    // v1.7.45 trash-can shortcut in topbar
    $("trashBtn").addEventListener("click", openRecoveryDialog);
    // refresh badge periodically (after data loads, after deletes)
    refreshTrashBadge();

    // v1.7.45 Settings tabs
    document.querySelectorAll(".settings-tab").forEach((tab) => {
      tab.addEventListener("click", () => {
        const name = tab.dataset.tab;
        document.querySelectorAll(".settings-tab").forEach((t) => {
          const active = t.dataset.tab === name;
          t.classList.toggle("active", active);
          t.setAttribute("aria-selected", String(active));
        });
        document.querySelectorAll(".settings-panel").forEach((p) => {
          p.hidden = p.dataset.panel !== name;
        });
      });
    });
    $("setExportMyDataBtn").addEventListener("click", exportMyData);
    $("setDeleteMyDataBtn").addEventListener("click", deleteMyData);
    $("privacyDetailsLink").addEventListener("click", (e) => { e.preventDefault(); $("privacyDialog").showModal(); });
    $("privacyCloseBtn").addEventListener("click", () => $("privacyDialog").close());

    // v1.7.46 — Topbar ? opens the keyboard shortcuts overlay (matches the
    // global `?` keystroke + the button's title hint "Quick guide (?)").
    // Full Quick Guide is still accessible via Cmd+K → "Open Quick guide".
    $("licGuideBtn").addEventListener("click", () => $("shortcutsDialog").showModal());
    $("licGuideCloseBtn").addEventListener("click", () => $("licGuideDialog").close());
    $("importFile").addEventListener("change", (e) => {
      const f = e.target.files && e.target.files[0];
      if (f) handleImportFile(f);
    });
    $("importCancelBtn").addEventListener("click", () => {
      $("importDialog").close();
      importRows = [];
    });
    $("importConfirmBtn").addEventListener("click", confirmImport);
    // Re-classify all rows when user toggles the dup-mode dropdown.
    $("importDupMode").addEventListener("change", () => {
      for (const r of importRows) {
        // Re-classify only rows that previously matched an existing license.
        if (r.existingId || r.status === "duplicate" || r.status === "update") {
          const cls = classifyRow(r);
          r.status = cls.status;
          r.reason = cls.reason || null;
        }
      }
      renderImportPreview();
    });
    $("importBulkProductLine").addEventListener("input", () => {
      // Just visual; bulk PL is read at confirm time.
    });

    // Add/edit dialog buttons
    $("licSaveBtn").addEventListener("click", saveLicense);
    $("licDeleteBtn").addEventListener("click", deleteLicense);
    $("licCancelBtn").addEventListener("click", closeEditDialog);
    $("licRenewBtn").addEventListener("click", () => {
      if (editingId) openRenewDialog(editingId);
    });
    $("licEmailBtn").addEventListener("click", () => {
      const lic = licenses.find((l) => l.id === editingId);
      if (lic) emailCustomer(lic);
    });

    // Customer 360 dialog
    $("customerDialogClose").addEventListener("click", () => $("customerDialog").close());
    // v1.7.41 Customer merge dialog
    $("mergeConfirmBtn").addEventListener("click", confirmMerge);
    $("mergeCancelBtn").addEventListener("click", () => $("customerMergeDialog").close());

    // Customer edit dialog
    $("custSaveBtn").addEventListener("click", saveCustomer);
    $("custCancelBtn").addEventListener("click", () => { $("customerEditDialog").close(); editingCustomerId = null; });

    // Email templates editor
    $("emailTemplatesBtn").addEventListener("click", openTemplatesDialog);
    $("tplSaveBtn").addEventListener("click", saveTemplate);
    $("tplDeleteBtn").addEventListener("click", deleteTemplate);
    $("tplNewBtn").addEventListener("click", newTemplate);
    $("templatesCloseBtn").addEventListener("click", () => $("templatesDialog").close());

    // v1.8.0 — Product Lines admin dialog
    const openPlBtn = $("openProductLinesBtn");
    if (openPlBtn) openPlBtn.addEventListener("click", () => {
      $("licSettingsDialog").close();
      openProductLinesDialog();
    });
    const manageLink = $("licProductLineManageLink");
    if (manageLink) manageLink.addEventListener("click", (e) => {
      e.preventDefault();
      $("licDialog").close();
      openProductLinesDialog();
    });
    $("plCloseBtn").addEventListener("click", () => $("productLinesDialog").close());
    $("plAddBtn").addEventListener("click", addProductLineFromInput);
    $("plAddInput").addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); addProductLineFromInput(); }
    });
    $("plNormalizePreviewBtn").addEventListener("click", previewNormalize);
    $("plNormalizeApplyBtn").addEventListener("click", applyNormalize);

    // Bulk select + reassign
    $("bulkSelectBtn").addEventListener("click", () => setBulkMode(!bulkMode));
    $("bulkReassignBtn").addEventListener("click", openBulkReassign);
    $("bulkClearBtn").addEventListener("click", () => { bulkSelected.clear(); updateBulkBar(); render(); });
    $("bulkReassignConfirm").addEventListener("click", confirmBulkReassign);
    $("bulkReassignCancel").addEventListener("click", () => $("bulkReassignDialog").close());
    // v1.7.41 bulk Renew / Export / Delete
    $("bulkRenewBtn").addEventListener("click", bulkRenewSelected);
    $("bulkExportBtn").addEventListener("click", bulkExportSelected);
    $("bulkDeleteBtn").addEventListener("click", bulkDeleteSelected);

    // Renew dialog
    document.querySelectorAll(".renew-presets [data-years]").forEach((b) => {
      b.addEventListener("click", () => confirmRenew(parseInt(b.dataset.years, 10), null));
    });
    $("renewConfirm").addEventListener("click", () => {
      const d = $("renewCustomDate").value;
      if (d) confirmRenew(null, d);
    });
    $("renewCancel").addEventListener("click", () => { $("renewDialog").close(); renewTargetId = null; });

    // Day dialog close
    $("dayDialogClose").addEventListener("click", () => $("dayDialog").close());

    // Calendar nav
    $("calPrev").addEventListener("click", () => {
      calCursor = new Date(calCursor.getFullYear(), calCursor.getMonth() - 1, 1);
      render();
    });
    $("calNext").addEventListener("click", () => {
      calCursor = new Date(calCursor.getFullYear(), calCursor.getMonth() + 1, 1);
      render();
    });
    $("calToday").addEventListener("click", () => {
      calCursor = startOfMonth(new Date());
      render();
    });
    // v1.7.39 Year jump dropdown
    $("calYearJump").addEventListener("change", (e) => {
      const y = parseInt(e.target.value, 10);
      if (!isFinite(y)) return;
      calCursor = new Date(y, calCursor.getMonth(), 1);
      render();
    });
    // v1.7.39 Calendar density toggle
    document.querySelectorAll("#calDensity .view-btn").forEach((b) => {
      b.addEventListener("click", () => setCalDensity(b.dataset.density));
    });
    setCalDensity(calDensity); // initial paint

    $("calNextYear").addEventListener("click", () => {
      calCursor = new Date(calCursor.getFullYear() + 1, calCursor.getMonth(), 1);
      render();
    });

    // v1.8.0 — Quarter view nav.
    function bumpQuarter(delta) {
      let { year, q } = quarterCursor;
      q += delta;
      while (q < 1) { q += 4; year -= 1; }
      while (q > 4) { q -= 4; year += 1; }
      quarterCursor = { year, q };
      render();
    }
    const qPrev = $("qPrev");
    if (qPrev) qPrev.addEventListener("click", () => bumpQuarter(-1));
    const qNext = $("qNext");
    if (qNext) qNext.addEventListener("click", () => bumpQuarter(1));
    const qToday = $("qToday");
    if (qToday) qToday.addEventListener("click", () => {
      const d = new Date();
      quarterCursor = { year: d.getFullYear(), q: Math.floor(d.getMonth() / 3) + 1 };
      render();
    });
    const qYearJump = $("qYearJump");
    if (qYearJump) qYearJump.addEventListener("change", (e) => {
      const m = (e.target.value || "").match(/^(\d{4})-Q([1-4])$/);
      if (!m) return;
      quarterCursor = { year: Number(m[1]), q: Number(m[2]) };
      render();
    });
    // Quarter view density toggle (mirrors Calendar density).
    document.querySelectorAll("#qDensity .view-btn").forEach((b) => {
      b.addEventListener("click", () => setCalDensity(b.dataset.density));
    });

    // v1.8.0 — quarter filter dropdown
    const qSel = $("quarterFilter");
    if (qSel) qSel.addEventListener("change", (e) => {
      quarterFilter = e.target.value || "";
      // v1.8.2 — Picking a specific quarter clears the MONTH dropdown AND
      // the "This quarter" quick-filter pill (which means "current quarter"
      // and would intersect to zero rows for any other quarter).
      if (quarterFilter) {
        monthFilter = "";
        $("monthFilter").value = "";
        if (quickFilter === "quarter" || quickFilter === "month") quickFilter = "all";
      }
      persistFilters();
      render();
    });
  }

  // ---------- customer edit dialog ----------
  let editingCustomerId = null;
  function openCustomerEditDialog(customerName) {
    const cust = customerByName(customerName);
    editingCustomerId = cust ? cust.id : null;
    const fallbackName = customerName || (cust && cust.name) || "";
    $("customerEditName").textContent = fallbackName;
    $("custPrimaryEmail").value = cust ? (cust.primaryEmail || "") : "";
    $("custSecondaryEmails").value = cust && Array.isArray(cust.secondaryEmails) ? cust.secondaryEmails.join("\n") : "";
    $("custAddress").value = cust ? (cust.address || "") : "";
    $("custNotes").value = cust ? (cust.notes || "") : "";
    $("customerDialog").close();
    $("customerEditDialog").showModal();
    $("custPrimaryEmail").focus();
  }
  async function saveCustomer() {
    const secondaryRaw = $("custSecondaryEmails").value.trim();
    const secondaryEmails = secondaryRaw
      ? secondaryRaw.split(/[\n,]/).map((s) => s.trim()).filter(Boolean)
      : [];
    const payload = {
      name: $("customerEditName").textContent || "",
      primaryEmail: $("custPrimaryEmail").value.trim() || null,
      secondaryEmails,
      address: $("custAddress").value.trim() || null,
      notes: $("custNotes").value.trim() || null,
    };
    try {
      if (editingCustomerId) {
        const { customer } = await api("PATCH", `/customers/${editingCustomerId}`, payload);
        const idx = customers.findIndex((c) => c.id === customer.id);
        if (idx >= 0) customers[idx] = customer; else customers.push(customer);
      } else {
        const { customer } = await api("POST", "/customers", payload);
        customers.push(customer);
      }
      $("customerEditDialog").close();
      editingCustomerId = null;
      toast("Customer saved.");
    } catch (err) {
      showError("Save failed", err);
    }
  }

  // ---------- email templates editor ----------
  let editingTemplateKey = null; // productLine of the template currently in the form
  function openTemplatesDialog() {
    renderTemplatesList();
    // Default-select _default or first item
    const first = emailTemplates.find((t) => t.productLine === "_default") || emailTemplates[0] || { productLine: "_default", ...DEFAULT_TEMPLATES._default };
    selectTemplate(first.productLine);
    $("templatesDialog").showModal();
  }
  function renderTemplatesList() {
    const ul = $("templatesList");
    ul.innerHTML = "";
    // Merge saved + defaults (defaults shown if no saved version exists)
    const seen = new Set();
    const all = [];
    for (const t of emailTemplates) { all.push(t); seen.add(t.productLine.toLowerCase()); }
    for (const key of Object.keys(DEFAULT_TEMPLATES)) {
      if (!seen.has(key.toLowerCase())) all.push({ productLine: key, ...DEFAULT_TEMPLATES[key], _default: true });
    }
    all.sort((a, b) => (a.productLine === "_default" ? -1 : b.productLine === "_default" ? 1 : a.productLine.localeCompare(b.productLine)));
    for (const t of all) {
      const li = document.createElement("li");
      li.dataset.key = t.productLine;
      const label = t.productLine === "_default" ? "Default (any product line)" : t.productLine;
      li.textContent = label + (t._default ? " (built-in)" : "");
      if (t.productLine === editingTemplateKey) li.classList.add("active");
      li.addEventListener("click", () => selectTemplate(t.productLine));
      ul.appendChild(li);
    }
  }
  function selectTemplate(productLine) {
    editingTemplateKey = productLine;
    const t = emailTemplates.find((x) => x.productLine === productLine)
      || (DEFAULT_TEMPLATES[productLine] ? { productLine, ...DEFAULT_TEMPLATES[productLine] } : { productLine, subject: "", body: "" });
    $("tplProductLine").value = t.productLine;
    $("tplSubject").value = t.subject || "";
    $("tplBody").value = t.body || "";
    renderTemplatesList();
  }
  async function saveTemplate() {
    const key = $("tplProductLine").value.trim() || "_default";
    const subject = $("tplSubject").value;
    const body = $("tplBody").value;
    try {
      const { templates } = await api("PUT", "/email-templates", {
        templates: [{ productLine: key, subject, body }],
      });
      // Merge into local state
      const t = templates[0];
      if (t) {
        const idx = emailTemplates.findIndex((x) => x.productLine === t.productLine);
        if (idx >= 0) emailTemplates[idx] = t; else emailTemplates.push(t);
      }
      editingTemplateKey = key;
      renderTemplatesList();
      toast("Template saved.");
    } catch (err) {
      showError("Save failed", err);
    }
  }
  async function deleteTemplate() {
    if (!editingTemplateKey) return;
    if (!confirm(`Delete template for "${editingTemplateKey}"? The built-in default will be used instead.`)) return;
    try {
      await api("DELETE", `/email-templates/${encodeURIComponent(editingTemplateKey)}`);
      emailTemplates = emailTemplates.filter((t) => t.productLine !== editingTemplateKey);
      // Reload the panel with whatever the default is now
      selectTemplate(editingTemplateKey);
      toast("Template removed.");
    } catch (err) {
      showError("Delete failed", err);
    }
  }
  function newTemplate() {
    editingTemplateKey = "";
    $("tplProductLine").value = "";
    $("tplSubject").value = DEFAULT_TEMPLATES._default.subject;
    $("tplBody").value = DEFAULT_TEMPLATES._default.body;
    renderTemplatesList();
    $("tplProductLine").focus();
  }

  // ---------- CSV import ----------

  // Quote-aware CSV row splitter. Handles "Beyond Innovations, Inc" type embedded commas
  // and "" escaped quotes inside quoted cells.
  function parseCsv(text) {
    const rows = [];
    let row = [];
    let cell = "";
    let inQuotes = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (inQuotes) {
        if (c === '"' && text[i + 1] === '"') { cell += '"'; i++; }
        else if (c === '"') { inQuotes = false; }
        else { cell += c; }
      } else {
        if (c === '"') inQuotes = true;
        else if (c === ",") { row.push(cell); cell = ""; }
        else if (c === "\n" || c === "\r") {
          if (c === "\r" && text[i + 1] === "\n") i++;
          row.push(cell); cell = "";
          rows.push(row); row = [];
        } else { cell += c; }
      }
    }
    if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
    return rows.map((r) => r.map((c) => c.trim()));
  }

  // Accept ISO (YYYY-MM-DD), 17-Jul-26 / 17-Jul-2026, and M/D/YYYY or M/D/YY.
  function parseDateMulti(s) {
    if (!s) return null;
    s = String(s).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const months = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
    let m = /^(\d{1,2})[\s\-\/]+([A-Za-z]{3,9})[\s\-\/]+(\d{2,4})$/.exec(s);
    if (m) {
      const day = parseInt(m[1], 10);
      const mon = months.indexOf(m[2].slice(0, 3).toLowerCase());
      if (mon === -1) return null;
      let year = parseInt(m[3], 10);
      if (year < 100) year += 2000;
      return `${year}-${String(mon + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
    m = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/.exec(s);
    if (m) {
      const mon = parseInt(m[1], 10);
      const day = parseInt(m[2], 10);
      let year = parseInt(m[3], 10);
      if (year < 100) year += 2000;
      if (mon < 1 || mon > 12 || day < 1 || day > 31) return null;
      return `${year}-${String(mon).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
    return null;
  }

  // Find the row index that looks like the header (has cells matching multiple
  // expected column keywords). Returns -1 if none matches.
  function findHeaderRow(rows) {
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i].map((c) => c.toLowerCase());
      let hits = 0;
      for (const c of r) {
        if (/customer|client/.test(c)) hits++;
        if (/license|product/.test(c)) hits++;
        if (/expir|due\s*date|date/.test(c)) hits++;
        if (/accountable|owner|assigned/.test(c)) hits++;
      }
      if (hits >= 3) return i;
    }
    return -1;
  }

  function mapColumns(header) {
    const idx = { customer: -1, licenseType: -1, userCount: -1, expiryDate: -1, ownerName: -1, notes: -1, productLine: -1 };
    header.forEach((c, i) => {
      const lc = c.toLowerCase().trim();
      if (idx.customer === -1 && /customer|client/.test(lc)) idx.customer = i;
      else if (idx.licenseType === -1 && /license\s*type|product\s*name/.test(lc)) idx.licenseType = i;
      else if (idx.userCount === -1 && /(num\.?\s*of\s*users|users|seats|qty|quantity|count)/.test(lc)) idx.userCount = i;
      else if (idx.expiryDate === -1 && /(expir|renewal\s*date|due\s*date)/.test(lc)) idx.expiryDate = i;
      else if (idx.ownerName === -1 && /(accountable|owner|assigned\s*to)/.test(lc)) idx.ownerName = i;
      else if (idx.notes === -1 && /(notes?|remarks?|comments?)/.test(lc)) idx.notes = i;
      else if (idx.productLine === -1 && /(product\s*line|category|module)/.test(lc)) idx.productLine = i;
    });
    return idx;
  }

  // Try to find a member by Accountable-name cell. Auto-match strategies:
  // 1. Exact (case-insensitive) match on displayName.
  // 2. Exact match on first token of displayName (e.g. "Dona" matches "Dona Apolonio") — only
  //    if exactly one member matches that first token.
  // 3. Substring contains (case-insensitive) — only if exactly one match.
  function matchOwner(rawName) {
    if (!rawName) return null;
    const n = rawName.trim().toLowerCase();
    if (!n) return null;
    const exact = members.find((m) => (m.displayName || "").trim().toLowerCase() === n);
    if (exact) return exact;
    const firstTok = members.filter((m) => {
      const dn = m.displayName || "";
      const first = dn.split(/\s+/)[0] || "";
      return first.toLowerCase() === n;
    });
    if (firstTok.length === 1) return firstTok[0];
    const sub = members.filter((m) => (m.displayName || "").toLowerCase().includes(n));
    if (sub.length === 1) return sub[0];
    return null;
  }

  function classifyRow(row) {
    if (!row.customer || !row.licenseType) return { status: "invalid", reason: "missing customer or license type" };
    if (!row.expiryDate) return { status: "invalid", reason: "missing or invalid expiry date" };
    // Match detection: same customer + licenseType + expiryDate as an existing license.
    const match = licenses.find((l) =>
      l.customer.trim().toLowerCase() === row.customer.trim().toLowerCase() &&
      l.licenseType.trim().toLowerCase() === row.licenseType.trim().toLowerCase() &&
      l.expiryDate === row.expiryDate
    );
    if (match) {
      row.existingId = match.id;
      const mode = $("importDupMode") ? $("importDupMode").value : "skip";
      if (mode === "update") return { status: "update", reason: "will update this row" };
      return { status: "duplicate", reason: "already in licenses (will skip)" };
    }
    if (!row.ownerOid) return { status: "needsOwner", reason: "owner not matched" };
    return { status: "ready" };
  }

  // Generate and download a CSV template the user can fill in. Headers match
  // what the importer recognizes; sample rows show the expected date / owner-name
  // formats. Intentionally does not include Status or Renewal Cycle because the
  // importer doesn't currently read those (set per-row in the table after import).
  function downloadCsvTemplate() {
    const headers = ["Customer", "License Type", "Number of Users", "Expiry Date", "Accountable", "Product Line", "Notes"];
    const samples = [
      ["Beyond Innovations, Inc", "Microsoft 365 Business Standard", 24, "2026-07-17", "Joshua Oducado", "M365", "Renewal contact: oky@example.com"],
      ["Sidel Industrial Packaging Corporation", "Dynamics 365 Business Central Premium", 4, "2026-07-24", "Dona", "BC", "Email notice sent 2026-06-10"],
      ["Lopez Holdings Corporation", "PhilTax Module", 5, "2026-07-17", "Dona", "PhilTax", "3rd year of Triennial Plan"],
    ];
    const csv = [headers, ...samples].map((row) => row.map(csvEscape).join(",")).join("\n") + "\n";
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "day-reminders-license-import-template.csv";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 100);
    toast("Template downloaded. Fill it in then use Import CSV.");
  }

  function openImportDialog() {
    importRows = [];
    $("importFile").value = "";
    $("importStep1").hidden = false;
    $("importStep2").hidden = true;
    $("importBulkProductLine").value = "";
    $("importTbody").innerHTML = "";
    $("importStats").textContent = "";
    $("importConfirmBtn").disabled = true;
    $("importConfirmBtn").textContent = "Import 0 licenses";
    $("importDialog").showModal();
  }

  async function handleImportFile(file) {
    if (!file) return;
    const text = await file.text();
    const rawRows = parseCsv(text).filter((r) => r.some((c) => c && c.trim()));
    if (!rawRows.length) { toast("CSV is empty"); return; }

    const headerIdx = findHeaderRow(rawRows);
    if (headerIdx === -1) {
      toast("Couldn't find a header row with Customer / License type / Expiry");
      return;
    }
    const header = rawRows[headerIdx];
    const cols = mapColumns(header);
    if (cols.customer === -1 || cols.licenseType === -1 || cols.expiryDate === -1) {
      toast("CSV needs Customer, License type, and Expiry date columns");
      return;
    }

    importRows = [];
    for (let i = headerIdx + 1; i < rawRows.length; i++) {
      const r = rawRows[i];
      const customer = (r[cols.customer] || "").trim();
      const licenseType = (r[cols.licenseType] || "").trim();
      // Skip totals/section rows: missing both customer AND licenseType.
      if (!customer && !licenseType) continue;
      // Skip rows that look like TOTAL summaries.
      if (/^total\b/i.test(customer)) continue;

      const expiryRaw = cols.expiryDate >= 0 ? (r[cols.expiryDate] || "") : "";
      const expiryDate = parseDateMulti(expiryRaw);
      const ownerRaw = cols.ownerName >= 0 ? (r[cols.ownerName] || "").trim() : "";
      const owner = matchOwner(ownerRaw);
      const userCountRaw = cols.userCount >= 0 ? (r[cols.userCount] || "").replace(/,/g, "") : "";
      const userCount = userCountRaw ? Math.max(0, parseInt(userCountRaw, 10) || 0) : 0;
      const notes = cols.notes >= 0 ? (r[cols.notes] || "").trim() : "";
      const productLine = cols.productLine >= 0 ? (r[cols.productLine] || "").trim() : "";

      const row = {
        customer,
        licenseType,
        userCount,
        expiryDate,
        ownerOid: owner ? owner.oid : null,
        ownerName: owner ? owner.displayName : ownerRaw || null,
        ownerRawName: ownerRaw || null,
        notes: notes || null,
        productLine: productLine || null,
      };
      const cls = classifyRow(row);
      row.status = cls.status;
      row.reason = cls.reason || null;
      importRows.push(row);
    }

    if (!importRows.length) {
      toast("No data rows found in CSV");
      return;
    }
    $("importStep1").hidden = true;
    $("importStep2").hidden = false;
    renderImportPreview();
  }

  function renderImportPreview() {
    const tbody = $("importTbody");
    tbody.innerHTML = "";
    importRows.forEach((row, idx) => {
      const tr = document.createElement("tr");
      tr.dataset.idx = idx;
      tr.classList.add(`status-${row.status}`);

      const tdStatus = document.createElement("td");
      tdStatus.className = "status-cell";
      tdStatus.title = row.reason || "";
      tdStatus.textContent = row.status === "ready" ? "✓" :
                             row.status === "needsOwner" ? "⚠" :
                             row.status === "duplicate" ? "↻" :
                             row.status === "update" ? "↻" :
                             "✗";
      tr.appendChild(tdStatus);

      const tdCust = document.createElement("td");
      tdCust.textContent = row.customer;
      tr.appendChild(tdCust);

      const tdType = document.createElement("td");
      tdType.textContent = row.licenseType;
      tr.appendChild(tdType);

      const tdUsers = document.createElement("td");
      tdUsers.className = "num";
      tdUsers.textContent = row.userCount;
      tr.appendChild(tdUsers);

      const tdExp = document.createElement("td");
      tdExp.textContent = row.expiryDate ? fmtShortDate(row.expiryDate) : "(invalid date)";
      tr.appendChild(tdExp);

      const tdOwner = document.createElement("td");
      if (row.status === "needsOwner" || (!row.ownerOid && row.status !== "invalid")) {
        const pickerWrap = document.createElement("div");
        createPeoplePicker(pickerWrap, {
          initialOid: row.ownerOid || null,
          initialName: row.ownerName || "",
          placeholder: row.ownerRawName ? `Search (CSV: "${row.ownerRawName}")` : "Search teammates…",
          onChange: ({ oid, name }) => {
            row.ownerOid = oid || null;
            row.ownerName = name || null;
            const cls = classifyRow(row);
            row.status = cls.status;
            row.reason = cls.reason || null;
            renderImportPreview();
            updateImportStats();
          },
        });
        tdOwner.appendChild(pickerWrap);
        if (row.ownerRawName) {
          const hint = document.createElement("div");
          hint.className = "owner-hint";
          hint.textContent = `from CSV: "${row.ownerRawName}"`;
          tdOwner.appendChild(hint);
        }
      } else if (row.ownerOid) {
        const pill = document.createElement("span");
        pill.className = "owner-pill small";
        pill.style.background = ownerColor(row.ownerOid);
        pill.textContent = row.ownerName || "";
        tdOwner.appendChild(pill);
      } else {
        tdOwner.textContent = row.ownerRawName || "—";
      }
      tr.appendChild(tdOwner);

      const tdNotes = document.createElement("td");
      tdNotes.textContent = row.notes ? (row.notes.length > 40 ? row.notes.slice(0, 39) + "…" : row.notes) : "";
      tr.appendChild(tdNotes);

      const tdRemove = document.createElement("td");
      const x = document.createElement("button");
      x.type = "button";
      x.className = "btn ghost small";
      x.textContent = "×";
      x.title = "Drop this row from import";
      x.addEventListener("click", () => {
        importRows.splice(idx, 1);
        renderImportPreview();
        updateImportStats();
      });
      tdRemove.appendChild(x);
      tr.appendChild(tdRemove);

      tbody.appendChild(tr);
    });
    updateImportStats();
  }

  function updateImportStats() {
    const ready = importRows.filter((r) => r.status === "ready").length;
    const updates = importRows.filter((r) => r.status === "update").length;
    const needsOwner = importRows.filter((r) => r.status === "needsOwner").length;
    const invalid = importRows.filter((r) => r.status === "invalid").length;
    const dup = importRows.filter((r) => r.status === "duplicate").length;
    const parts = [];
    if (ready) parts.push(`${ready} new`);
    if (updates) parts.push(`${updates} to update`);
    if (needsOwner) parts.push(`${needsOwner} need owner`);
    if (dup) parts.push(`${dup} skip (duplicate)`);
    if (invalid) parts.push(`${invalid} invalid`);
    $("importStats").textContent = parts.join(" · ");
    const btn = $("importConfirmBtn");
    const actionable = ready + updates;
    btn.disabled = actionable === 0;
    if (updates && !ready) btn.textContent = updates === 1 ? `Update 1 license` : `Update ${updates} licenses`;
    else if (updates && ready) btn.textContent = `Import ${ready} + update ${updates}`;
    else btn.textContent = ready === 1 ? `Import 1 license` : `Import ${ready} licenses`;
  }

  async function confirmImport() {
    const bulkPL = $("importBulkProductLine").value.trim() || null;
    const toCreate = importRows.filter((r) => r.status === "ready");
    const toUpdate = importRows.filter((r) => r.status === "update" && r.existingId);
    if (!toCreate.length && !toUpdate.length) return;
    const btn = $("importConfirmBtn");
    btn.disabled = true;
    btn.textContent = "Importing…";
    let created = 0, updated = 0, failed = 0;
    for (const row of toCreate) {
      try {
        const payload = {
          customer: row.customer,
          licenseType: row.licenseType,
          userCount: row.userCount,
          expiryDate: row.expiryDate,
          ownerOid: row.ownerOid,
          ownerName: row.ownerName,
          productLine: bulkPL || row.productLine || null,
          notes: row.notes,
        };
        const { license } = await api("POST", "/licenses", payload);
        licenses.push(license);
        created++;
      } catch (err) {
        console.error("import row failed", row, err);
        failed++;
      }
    }
    for (const row of toUpdate) {
      try {
        // Only patch fields the importer can carry (don't blow away status / cycle).
        const payload = {
          userCount: row.userCount,
          notes: row.notes,
        };
        if (row.ownerOid) { payload.ownerOid = row.ownerOid; payload.ownerName = row.ownerName; }
        if (bulkPL || row.productLine) payload.productLine = bulkPL || row.productLine;
        const { license } = await api("PATCH", `/licenses/${row.existingId}`, payload);
        licenses = licenses.map((l) => l.id === license.id ? license : l);
        updated++;
      } catch (err) {
        console.error("update row failed", row, err);
        failed++;
      }
    }
    $("importDialog").close();
    render();
    const parts = [];
    if (created) parts.push(`${created} added`);
    if (updated) parts.push(`${updated} updated`);
    if (failed) parts.push(`${failed} failed`);
    toast(parts.join(", "));
    importRows = [];
  }

  // ---------- license-tab settings ----------
  function openSettingsDialog() {
    $("setThemeOverride").value = themeOverride;
    let defaultLeads = userSettings.licenseLeadDays;
    if (typeof defaultLeads === "number") defaultLeads = [defaultLeads];
    if (!Array.isArray(defaultLeads) || !defaultLeads.length) defaultLeads = [14];
    ensureLeadPicker("setLicenseLeadDaysPicker").set(defaultLeads);
    $("setSkipBriefing").checked = !!userSettings.licenseSkipBriefing;
    $("setSkipMonthlyDigest").checked = !!userSettings.licenseSkipMonthlyDigest;
    $("setRollupDigest").checked = !!userSettings.licenseRollupDigest;
    $("licSettingsDialog").showModal();
  }

  // ---------- v1.7.43 Recovery (recently-deleted licenses) ----------
  async function openRecoveryDialog() {
    const list = $("recoveryList");
    list.innerHTML = '<li class="recovery-loading">Loading…</li>';
    $("recoveryDialog").showModal();
    try {
      const { licenses: items } = await api("GET", "/licenses/deleted");
      list.innerHTML = "";
      if (!items.length) {
        const li = document.createElement("li");
        li.className = "recovery-empty";
        li.textContent = "Nothing in the recycle bin. Anything you delete shows up here for 30 days.";
        list.appendChild(li);
        return;
      }
      items.sort((a, b) => (b.deletedAt || "").localeCompare(a.deletedAt || ""));
      const now = Date.now();
      for (const lic of items) {
        const li = document.createElement("li");
        li.className = "recovery-row";
        const head = document.createElement("div");
        head.className = "recovery-head";
        const name = document.createElement("strong");
        name.textContent = `${lic.customer} · ${lic.licenseType}`;
        head.appendChild(name);
        const meta = document.createElement("div");
        meta.className = "recovery-meta";
        const deletedMs = Date.parse(lic.deletedAt);
        const daysLeft = Math.max(0, 30 - Math.floor((now - deletedMs) / 86400000));
        meta.textContent = `Deleted ${fmtRelative(deletedMs)}${lic.deletedByName ? ` by ${lic.deletedByName}` : ""}  ·  ${daysLeft} day${daysLeft === 1 ? "" : "s"} until permanent purge`;
        const restoreBtn = document.createElement("button");
        restoreBtn.type = "button";
        restoreBtn.className = "btn primary small";
        restoreBtn.textContent = "Restore";
        restoreBtn.addEventListener("click", async () => {
          restoreBtn.disabled = true;
          restoreBtn.textContent = "Restoring…";
          try {
            const { license } = await api("POST", `/licenses/${lic.id}/restore`);
            licenses = [...licenses, license];
            render();
            li.remove();
            toast(`Restored: ${license.customer}, ${license.licenseType}`);
            refreshTrashBadge();
            if (!list.querySelector(".recovery-row")) openRecoveryDialog();
          } catch (err) {
            restoreBtn.disabled = false;
            restoreBtn.textContent = "Restore";
            showError("Restore failed", err);
          }
        });
        li.appendChild(head);
        li.appendChild(meta);
        li.appendChild(restoreBtn);
        list.appendChild(li);
      }
    } catch (err) {
      list.innerHTML = "";
      const li = document.createElement("li");
      li.className = "recovery-empty";
      li.textContent = `Could not load: ${err.message || err}`;
      list.appendChild(li);
    }
  }

  // ---------- v1.7.43 GDPR data subject rights ----------
  async function exportMyData() {
    const btn = $("setExportMyDataBtn");
    btn.disabled = true;
    const original = btn.textContent;
    btn.textContent = "Preparing…";
    try {
      const res = await fetch(`${API_BASE}/me/export`, {
        headers: { Authorization: `Bearer ${authToken}` },
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `day-reminders-export-${todayPh()}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 100);
      toast("Your data export was downloaded.");
    } catch (err) { showError("Export failed", err); }
    finally { btn.disabled = false; btn.textContent = original; }
  }
  async function deleteMyData() {
    const confirmed = confirm(
      "This will permanently delete your reminders, templates, settings, and Teams chat handle, and unassign you from any license rows you own. License rows themselves stay with Kation. This cannot be undone — proceed?"
    );
    if (!confirmed) return;
    const second = prompt('Type "DELETE" to confirm:');
    if (second !== "DELETE") { toast("Cancelled — nothing was deleted."); return; }
    const btn = $("setDeleteMyDataBtn");
    btn.disabled = true;
    const original = btn.textContent;
    btn.textContent = "Deleting…";
    try {
      const res = await api("DELETE", "/me/data");
      toast(`Done — ${res.reminderCount || 0} reminder${res.reminderCount === 1 ? "" : "s"} deleted, ${res.licensesUnassigned || 0} license${res.licensesUnassigned === 1 ? "" : "s"} unassigned.`);
      // Hard reload after a tick so the tab refetches everything cleanly.
      setTimeout(() => location.reload(), 1500);
    } catch (err) {
      btn.disabled = false;
      btn.textContent = original;
      showError("Delete failed", err);
    }
  }
  async function sendTestDigest(forceRollup) {
    const btn = forceRollup ? $("setSendTestDigestAllBtn") : $("setSendTestDigestBtn");
    const original = btn.textContent;
    btn.disabled = true;
    btn.textContent = "Sending…";
    try {
      const res = await api("POST", "/licenses/digest/preview", { rollup: !!forceRollup });
      const counts = res.counts || {};
      const sentTo = res.sentTo || "your mailbox";
      const total = (counts.overdue || 0) + (counts.thisMonth || 0) + (counts.nextMonth || 0);
      const variant = forceRollup ? " (all-accounts view)" : "";
      if (total === 0 && !forceRollup) {
        toast(`Sent to ${sentTo}${variant} (no licenses in window — empty digest).`);
      } else {
        const parts = [];
        if (counts.overdue) parts.push(`${counts.overdue} overdue`);
        if (counts.thisMonth) parts.push(`${counts.thisMonth} this month`);
        if (counts.nextMonth) parts.push(`${counts.nextMonth} next month`);
        toast(`Sent to ${sentTo}${variant} — ${parts.join(", ") || "all-accounts section attached"}.`);
      }
    } catch (err) {
      showError("Test digest failed", err);
    } finally {
      btn.disabled = false;
      btn.textContent = original;
    }
  }

  async function saveLicSettings() {
    // Theme override is tab-only state, persisted in localStorage. Save first
    // so users see the new theme apply instantly even if the /api/settings
    // round-trip fails.
    const newTheme = $("setThemeOverride").value;
    if (["auto", "default", "dark", "contrast"].includes(newTheme)) {
      themeOverride = newTheme;
      try { localStorage.setItem("lic.themeOverride", newTheme); } catch (_) {}
      applyTheme(teamsTheme);
    }
    const pickerLeads = ensureLeadPicker("setLicenseLeadDaysPicker").get();
    const next = {
      ...userSettings,
      licenseLeadDays: pickerLeads.length ? pickerLeads : [14],
      licenseSkipBriefing: $("setSkipBriefing").checked,
      licenseSkipMonthlyDigest: $("setSkipMonthlyDigest").checked,
      licenseRollupDigest: $("setRollupDigest").checked,
    };
    try {
      const { settings } = await api("PUT", "/settings", { settings: next });
      userSettings = { ...userSettings, ...settings };
      $("licSettingsDialog").close();
      toast("Settings saved.");
    } catch (err) {
      showError("Save failed", err);
    }
  }

  // ---------- manual refresh ----------
  async function refreshAll() {
    const btn = $("refreshBtn");
    if (!btn) return;
    const originalText = btn.innerHTML;
    btn.disabled = true;
    btn.textContent = "Refreshing…";
    try {
      const [licRes, memRes, custRes, tplRes, plRes] = await Promise.all([
        api("GET", "/licenses"),
        api("GET", "/members").catch(() => ({ members })),
        api("GET", "/customers").catch(() => ({ customers })),
        api("GET", "/email-templates").catch(() => ({ templates: emailTemplates })),
        api("GET", "/product-lines").catch(() => ({ productLines: productLinesRegistry })),
      ]);
      licenses = licRes.licenses || [];
      members = memRes.members || members;
      customers = custRes.customers || customers;
      emailTemplates = tplRes.templates || emailTemplates;
      productLinesRegistry = plRes.productLines || productLinesRegistry;
      saveCachedMembers(members);
      lastSyncedAt = Date.now();
      updateSyncIndicator();
      render();
      toast(`Refreshed (${licenses.length} licenses)`);
    } catch (err) {
      showError("Refresh failed", err);
    } finally {
      btn.disabled = false;
      btn.innerHTML = originalText;
    }
  }

  // ---------- v1.7.39 command palette (Ctrl+K) ----------
  // Two kinds of entries: actions (Add license, Refresh, Open settings, …) and
  // entities (licenses + customers). Fuzzy-ish substring search on label tokens.
  function buildCmdkActions() {
    return [
      { kind: "action", label: "Add license", hint: "Open the full Add dialog", run: () => openAddDialog() },
      { kind: "action", label: "Quick add license", hint: "Inline row at top of table", run: () => { $("quickAddDetails").open = true; $("qaCustomer").focus(); } },
      { kind: "action", label: "Switch to Table view", run: () => switchView("table") },
      { kind: "action", label: "Switch to Calendar view", run: () => switchView("calendar") },
      { kind: "action", label: "Refresh data", hint: "Reload from server", run: () => refreshAll() },
      { kind: "action", label: "Open Settings", run: () => openSettingsDialog() },
      { kind: "action", label: "Open Quick guide", run: () => $("licGuideDialog").showModal() },
      { kind: "action", label: "Export CSV (current view)", run: () => exportCsv() },
      { kind: "action", label: "Import CSV", run: () => $("importCsvFile").click() },
      { kind: "action", label: "Email templates", run: () => $("emailTemplatesBtn").click() },
      { kind: "action", label: "Toggle bulk select", run: () => setBulkMode(!bulkMode) },
      { kind: "action", label: "Clear all filters", run: () => clearAllFilters() },
      { kind: "action", label: "Save current view…", run: () => saveCurrentView() },
      { kind: "action", label: "Show keyboard shortcuts", run: () => $("shortcutsDialog").showModal() },
    ];
  }
  function switchView(v) {
    currentView = v;
    try { localStorage.setItem(LS_VIEW, v); } catch (_) {}
    render();
  }
  function rankCmdkMatch(query, label) {
    if (!query) return 1;
    const q = query.toLowerCase();
    const l = label.toLowerCase();
    if (l.startsWith(q)) return 3;
    if (l.includes(q)) return 2;
    // token-each match
    const tokens = q.split(/\s+/).filter(Boolean);
    if (tokens.every((t) => l.includes(t))) return 1;
    return 0;
  }
  function runCmdkSearch(query) {
    const results = [];
    for (const a of buildCmdkActions()) {
      const score = rankCmdkMatch(query, a.label);
      if (score > 0) results.push({ ...a, score });
    }
    if (query) {
      for (const lic of licenses) {
        const label = `${lic.customer} · ${lic.licenseType}`;
        const score = rankCmdkMatch(query, label);
        if (score > 0) results.push({ kind: "license", label, hint: lic.expiryDate || "", score, run: () => openEditDialog(lic) });
      }
      const seenCust = new Set();
      for (const lic of licenses) {
        const c = lic.customer || "";
        const k = c.trim().toLowerCase();
        if (!c || seenCust.has(k)) continue;
        seenCust.add(k);
        const score = rankCmdkMatch(query, c);
        if (score > 0) results.push({ kind: "customer", label: c, hint: "Customer profile", score, run: () => openCustomerDialog(c) });
      }
    }
    results.sort((a, b) => b.score - a.score);
    return results.slice(0, 25);
  }
  let cmdkActiveIdx = 0;
  let cmdkCurrent = [];
  function renderCmdkResults(results) {
    cmdkCurrent = results;
    const list = $("cmdkResults");
    list.innerHTML = "";
    if (!results.length) {
      const li = document.createElement("li");
      li.className = "cmdk-empty";
      li.textContent = "No matches";
      list.appendChild(li);
      return;
    }
    cmdkActiveIdx = Math.min(cmdkActiveIdx, results.length - 1);
    if (cmdkActiveIdx < 0) cmdkActiveIdx = 0;
    results.forEach((r, i) => {
      const li = document.createElement("li");
      li.className = "cmdk-result" + (i === cmdkActiveIdx ? " active" : "");
      li.setAttribute("role", "option");
      const kind = document.createElement("span");
      kind.className = `cmdk-kind cmdk-kind-${r.kind}`;
      kind.textContent = r.kind === "action" ? "⚡" : r.kind === "license" ? "📄" : "👤";
      const lab = document.createElement("span");
      lab.className = "cmdk-label";
      lab.textContent = r.label;
      const hint = document.createElement("span");
      hint.className = "cmdk-hint";
      hint.textContent = r.hint || "";
      li.appendChild(kind);
      li.appendChild(lab);
      li.appendChild(hint);
      li.addEventListener("mouseenter", () => { cmdkActiveIdx = i; refreshActiveHighlight(); });
      li.addEventListener("click", () => runCmdkResult(r));
      list.appendChild(li);
    });
  }
  function refreshActiveHighlight() {
    const items = $("cmdkResults").querySelectorAll(".cmdk-result");
    items.forEach((el, i) => el.classList.toggle("active", i === cmdkActiveIdx));
  }
  function runCmdkResult(r) {
    closeCmdk();
    try { r.run(); } catch (err) { showError("Action failed", err); }
  }
  function openCmdk() {
    const d = $("cmdkDialog");
    if (!d) return;
    $("cmdkInput").value = "";
    cmdkActiveIdx = 0;
    renderCmdkResults(runCmdkSearch(""));
    d.showModal();
    setTimeout(() => $("cmdkInput").focus(), 0);
  }
  function closeCmdk() {
    const d = $("cmdkDialog");
    if (d && d.open) d.close();
  }
  function wireCmdk() {
    $("cmdkInput").addEventListener("input", (e) => {
      cmdkActiveIdx = 0;
      renderCmdkResults(runCmdkSearch(e.target.value.trim()));
    });
    $("cmdkInput").addEventListener("keydown", (e) => {
      if (e.key === "ArrowDown") { e.preventDefault(); cmdkActiveIdx = Math.min(cmdkActiveIdx + 1, cmdkCurrent.length - 1); refreshActiveHighlight(); }
      else if (e.key === "ArrowUp") { e.preventDefault(); cmdkActiveIdx = Math.max(cmdkActiveIdx - 1, 0); refreshActiveHighlight(); }
      else if (e.key === "Enter") {
        e.preventDefault();
        const chosen = cmdkCurrent[cmdkActiveIdx];
        if (chosen) runCmdkResult(chosen);
      }
    });
    $("shortcutsCloseBtn").addEventListener("click", () => $("shortcutsDialog").close());
    // v1.7.47 — cross-link from shortcuts overlay to the full Quick guide.
    $("shortcutsOpenGuideBtn").addEventListener("click", () => {
      $("shortcutsDialog").close();
      $("licGuideDialog").showModal();
    });
  }

  // ---------- v1.7.45 table arrow-key navigation ----------
  // Up/Down moves a visual focus ring; Enter opens the edit dialog. Only
  // active when no dialog is open and the user isn't inside an input. Lives
  // alongside (not inside) the global shortcuts wiring so the two don't
  // step on each other.
  let tableNavRow = -1;
  function wireTableKeyboardNav() {
    document.addEventListener("keydown", (e) => {
      // Bail when something else owns the keystroke.
      if (e.altKey || e.ctrlKey || e.metaKey) return;
      const t = e.target;
      if (t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.tagName === "SELECT" || t.isContentEditable)) return;
      if (document.querySelectorAll("dialog[open]").length > 0) return;
      if (currentView !== "table") return;
      const rows = Array.from(document.querySelectorAll("#licTbody tr[data-id]"));
      if (!rows.length) return;
      if (e.key === "ArrowDown" || e.key === "j") {
        e.preventDefault();
        tableNavRow = Math.min(rows.length - 1, Math.max(0, tableNavRow + 1));
        focusTableRow(rows[tableNavRow]);
      } else if (e.key === "ArrowUp" || e.key === "k") {
        e.preventDefault();
        tableNavRow = Math.max(0, tableNavRow - 1);
        focusTableRow(rows[tableNavRow]);
      } else if (e.key === "Enter" && tableNavRow >= 0) {
        e.preventDefault();
        const id = rows[tableNavRow]?.dataset?.id;
        const lic = id && licenses.find((l) => l.id === id);
        if (lic) openEditDialog(lic);
      } else if (e.key === "Home") {
        e.preventDefault();
        tableNavRow = 0; focusTableRow(rows[0]);
      } else if (e.key === "End") {
        e.preventDefault();
        tableNavRow = rows.length - 1; focusTableRow(rows[tableNavRow]);
      }
    });
  }
  function focusTableRow(row) {
    if (!row) return;
    document.querySelectorAll("#licTbody tr.kbd-focus").forEach((r) => r.classList.remove("kbd-focus"));
    row.classList.add("kbd-focus");
    row.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }

  // ---------- v1.7.39 global keyboard shortcuts ----------
  let pendingG = false;
  function wireGlobalShortcuts() {
    document.addEventListener("keydown", (e) => {
      // Ctrl+K / Cmd+K — palette (works even inside inputs)
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k") {
        e.preventDefault();
        if ($("cmdkDialog").open) closeCmdk();
        else openCmdk();
        return;
      }
      // Suppress single-letter shortcuts if focus is inside an editable element.
      const t = e.target;
      const isEditing = t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.tagName === "SELECT" || t.isContentEditable);
      if (isEditing) return;
      if (e.altKey || e.ctrlKey || e.metaKey) return;

      if (e.key === "?" || (e.key === "/" && e.shiftKey)) {
        e.preventDefault();
        $("shortcutsDialog").showModal();
      } else if (e.key === "/") {
        e.preventDefault();
        $("searchInput").focus();
      } else if (e.key === "a") {
        e.preventDefault();
        openAddDialog();
      } else if (e.key === "q") {
        e.preventDefault();
        const d = $("quickAddDetails");
        d.open = !d.open;
        if (d.open) $("qaCustomer").focus();
      } else if (e.key === "r") {
        e.preventDefault();
        refreshAll();
      } else if (e.key === "g") {
        pendingG = true;
        setTimeout(() => { pendingG = false; }, 1200);
      } else if (pendingG && e.key === "t") {
        pendingG = false; switchView("table");
      } else if (pendingG && e.key === "c") {
        pendingG = false; switchView("calendar");
      }
    });
  }

  // v1.7.45 — trash-can topbar badge. Quietly polls the deleted-list endpoint
  // on boot + after deletes/restores so users see a count without opening
  // Settings. Fire-and-forget; missing endpoint just hides the badge.
  async function refreshTrashBadge() {
    const badge = $("trashBadge");
    if (!badge) return;
    try {
      const { licenses: items } = await api("GET", "/licenses/deleted");
      const n = (items || []).length;
      if (n > 0) { badge.textContent = String(n); badge.hidden = false; }
      else badge.hidden = true;
    } catch (_) { badge.hidden = true; }
  }

  // ---------- v1.7.39 live poll + relative-time freshness stamp ----------
  let lastSyncedAt = 0;
  let pollTimer = null;
  function fmtRelative(ms) {
    if (!ms) return "never";
    const diff = Math.max(0, Date.now() - ms);
    if (diff < 5000) return "just now";
    if (diff < 60000) return `${Math.floor(diff / 1000)}s ago`;
    if (diff < 3600000) return `${Math.floor(diff / 60000)}m ago`;
    if (diff < 86400000) return `${Math.floor(diff / 3600000)}h ago`;
    return `${Math.floor(diff / 86400000)}d ago`;
  }
  function updateSyncIndicator() {
    const el = $("syncIndicator");
    if (!el) return;
    // v1.7.48 — "Syncing…" until the first successful fetch (was "Updated never",
    // which read as a developer string).
    el.textContent = lastSyncedAt ? `Updated ${fmtRelative(lastSyncedAt)}` : "Syncing…";
    el.title = lastSyncedAt ? new Date(lastSyncedAt).toLocaleString() : "Initial sync in progress";
  }
  // Light background poll: refresh licenses every 60s when the tab is visible
  // so changes by another tenant user (Dona renews, Rey reassigns) appear
  // without a manual refresh. Skipped while a dialog is open to avoid clobbering
  // an in-progress edit. Errors are silent — keep-warm noise isn't worth a toast.
  // v1.7.40 — cheap signature so we can skip the full render when the poll
  // returns identical data. Compares id + the fields that drive every row:
  // expiry, status, owner, comment count, leadFired, edit timestamp.
  function licensesSignature(arr) {
    if (!Array.isArray(arr)) return "0";
    const parts = [String(arr.length)];
    for (const l of arr) {
      parts.push(
        `${l.id}|${l.expiryDate || ""}|${l.status || ""}|${l.ownerOid || ""}|${l.lastEditedAt || ""}|${Array.isArray(l.comments) ? l.comments.length : 0}|${Array.isArray(l.lastFiredLeadDays) ? l.lastFiredLeadDays.length : 0}`
      );
    }
    return parts.join("~");
  }
  let lastLicensesSig = "";

  async function pollOnce() {
    if (document.visibilityState !== "visible") return;
    const anyDialogOpen = document.querySelectorAll("dialog[open]").length > 0;
    if (anyDialogOpen) return;
    try {
      const res = await api("GET", "/licenses");
      const next = res.licenses || [];
      const pendingIds = new Set([...pendingDeletes.keys()]);
      const filtered = next.filter((l) => !pendingIds.has(l.id));
      lastSyncedAt = Date.now();
      const sig = licensesSignature(filtered);
      if (sig === lastLicensesSig) {
        // Nothing changed server-side; skip the render entirely.
        updateSyncIndicator();
        return;
      }
      lastLicensesSig = sig;
      licenses = filtered;
      render();
      updateSyncIndicator();
    } catch (_) { /* swallow — next poll will retry */ }
  }
  function startLivePoll() {
    if (pollTimer) clearInterval(pollTimer);
    pollTimer = setInterval(pollOnce, 60000);
    // Also refresh the "Updated X ago" stamp every 10s without hitting the server.
    setInterval(updateSyncIndicator, 10000);
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") pollOnce();
    });
  }

  // ---------- CSV export ----------

  function csvEscape(v) {
    if (v === null || v === undefined) return "";
    const s = String(v);
    if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  }

  // v1.7.41 — shared CSV builder so the toolbar Export and the bulk-mode
  // Export reuse the same column shape + escape rules.
  function exportCsvList(list, filename) {
    if (!list.length) { toast("Nothing to export"); return; }
    const headers = ["Customer", "License Type", "Number of Users", "Expiry Date", "Owner", "Product Line", "Lead Days", "Notes", "State", "Last Renewed"];
    const lines = [headers.map(csvEscape).join(",")];
    for (const l of list) {
      lines.push([
        l.customer,
        l.licenseType,
        l.userCount || 0,
        l.expiryDate || "",
        l.ownerName || "",
        l.productLine || "",
        Array.isArray(l.leadDays) ? l.leadDays.join(",") : (l.leadDays == null ? "" : l.leadDays),
        l.notes || "",
        l.state || "active",
        l.lastRenewedAt ? l.lastRenewedAt.slice(0, 10) : "",
      ].map(csvEscape).join(","));
    }
    const blob = new Blob([lines.join("\n") + "\n"], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename || `day-reminders-licenses-${todayPh()}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 100);
  }

  function exportCsv() {
    const list = sortLicenses(visibleLicenses());
    if (!list.length) { toast("Nothing to export with the current filters"); return; }
    exportCsvList(list, `day-reminders-licenses-${todayPh()}.csv`);
    toast(`Exported ${list.length} license${list.length === 1 ? "" : "s"}`);
  }

  // ---------- Teams init + boot ----------
  function applyTheme(theme) {
    teamsTheme = theme;
    const effective = themeOverride && themeOverride !== "auto"
      ? themeOverride
      : (theme === "dark" ? "dark" : theme === "contrast" ? "contrast" : "default");
    document.body.dataset.theme = effective;
  }

  // Cache /api/members in localStorage with a short TTL so the Owner picker
  // pre-fills instantly on subsequent loads while we refresh in the background.
  // Members change rarely (only when a new user opens Day Reminders the first time).
  const LS_MEMBERS_CACHE = "lic.membersCache";
  const MEMBERS_TTL_MS = 10 * 60 * 1000;
  function loadCachedMembers() {
    try {
      const raw = localStorage.getItem(LS_MEMBERS_CACHE);
      if (!raw) return null;
      const obj = JSON.parse(raw);
      if (!obj || !obj.at || Date.now() - obj.at > MEMBERS_TTL_MS) return null;
      return Array.isArray(obj.members) ? obj.members : null;
    } catch (_) { return null; }
  }
  function saveCachedMembers(mems) {
    try { localStorage.setItem(LS_MEMBERS_CACHE, JSON.stringify({ at: Date.now(), members: mems })); } catch (_) {}
  }

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
      me.oid = ctx.user?.id || null;
      me.name = ctx.user?.userPrincipalName || ctx.user?.displayName || null;

      // Fast path: paint the UI with cached members while licenses fetch.
      const cached = loadCachedMembers();
      if (cached) members = cached;
      wireEvents();

      // Kick off all secondary calls in parallel; render as soon as licenses arrive.
      const membersPromise = api("GET", "/members")
        .then((res) => { members = res.members || []; saveCachedMembers(members); render(); })
        .catch(() => {});
      const customersPromise = api("GET", "/customers")
        .then((res) => { customers = res.customers || []; render(); })
        .catch(() => {});
      const templatesPromise = api("GET", "/email-templates")
        .then((res) => { emailTemplates = res.templates || []; })
        .catch(() => {});
      // v1.8.0 — product-line registry. Seeded on first GET if empty.
      const productLinesPromise = api("GET", "/product-lines")
        .then((res) => { productLinesRegistry = res.productLines || []; render(); })
        .catch(() => {});
      const settingsPromise = api("GET", "/settings")
        .then((res) => { if (res && res.settings) userSettings = { ...userSettings, ...res.settings }; })
        .catch(() => {});
      const { licenses: lics } = await api("GET", "/licenses");
      licenses = lics || [];
      lastSyncedAt = Date.now();
      lastLicensesSig = licensesSignature(licenses);
      updateSyncIndicator();

      // Hide the boot indicator now that we have data.
      const bi = $("bootIndicator");
      if (bi) bi.classList.add("gone");
      render();
      // Don't await secondaries if still in flight; they just refresh state in background.
      void membersPromise; void customersPromise; void templatesPromise; void settingsPromise; void productLinesPromise;
      // v1.7.39 — start the live poll once initial paint is done.
      startLivePoll();
      // v1.7.39 — wire Ctrl+K palette and global shortcuts.
      wireCmdk();
      wireGlobalShortcuts();
      // v1.7.45 — table arrow-key nav.
      wireTableKeyboardNav();
      // v1.7.40 — add an X close button to every dialog.
      installDialogCloseButtons();
      // v1.7.42 — hydrate from URL hash if present (shareable link), then
      // listen for hash changes (back/forward, paste of a new link).
      if (loadFiltersFromHash()) {
        persistFilters();
        render();
      }
      window.addEventListener("hashchange", () => {
        suppressHashSync = true;
        if (loadFiltersFromHash()) render();
        suppressHashSync = false;
      });
    } catch (err) {
      const bi = $("bootIndicator");
      if (bi) bi.classList.add("gone");
      showError("Could not connect", err);
    }
  }

  boot();
})();
