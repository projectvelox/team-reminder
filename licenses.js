/* Day Reminders — Licenses tab (v1.7.0)
   Tenant-shared license tracker with Table + Calendar (Month) views.
   Auth via Teams SSO. Server is source of truth; we only cache for the session.
*/
(function () {
  "use strict";

  const API_BASE = "https://func-day-reminders-17023.azurewebsites.net/api";
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

  try {
    const v = localStorage.getItem(LS_VIEW);
    if (v === "table" || v === "calendar") currentView = v;
    const q = localStorage.getItem(LS_QUICK);
    if (["all", "mine", "month", "overdue", "attention"].includes(q)) quickFilter = q;
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
  } catch (_) {}
  function persistFilters() {
    try {
      localStorage.setItem(LS_OWNER_FILTER, JSON.stringify([...ownerFilter]));
      localStorage.setItem(LS_PRODUCT_FILTER, JSON.stringify([...productFilter]));
      localStorage.setItem(LS_STATUS_FILTER, JSON.stringify([...statusFilter]));
      localStorage.setItem(LS_EXPIRY_FILTER, JSON.stringify([...expiryFilter]));
      localStorage.setItem(LS_MONTH_FILTER, monthFilter || "");
    } catch (_) {}
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
    const durationMs = (opts && opts.durationMs) || (opts && opts.actionLabel ? 6000 : 2400);
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
  async function api(method, path, body) {
    const headers = { "Authorization": `Bearer ${authToken}` };
    const init = { method, headers };
    if (body !== undefined) {
      headers["Content-Type"] = "application/json";
      init.body = JSON.stringify(body);
    }
    const res = await fetch(API_BASE + path, init);
    if (res.status === 204) return null;
    let data = null;
    try { data = await res.json(); } catch (_) {}
    if (!res.ok) {
      const msg = data && data.error ? data.error : `HTTP ${res.status}`;
      throw new Error(msg);
    }
    return data;
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
    if (lic.ownerOid) {
      const img = document.createElement("img");
      img.src = ownerAvatarUrl(lic.ownerOid);
      img.alt = "";
      img.loading = "lazy";
      img.addEventListener("error", () => { img.remove(); });
      wrap.appendChild(img);
    }
    const initials = document.createElement("span");
    initials.className = "owner-avatar-initials";
    initials.textContent = ownerInitials(lic.ownerName);
    wrap.appendChild(initials);
    return wrap;
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
      !!searchText
    );
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
      const QUICK_LABEL = { mine: "Mine", month: "This month", overdue: "Overdue", attention: "Needs attention" };
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
    searchText = typeof snap.searchText === "string" ? snap.searchText : "";
    if (typeof snap.groupBy === "string") groupBy = snap.groupBy;
    persistFilters();
    try { localStorage.setItem(LS_QUICK, quickFilter); } catch (_) {}
    try { localStorage.setItem(LS_GROUP, groupBy); } catch (_) {}
    $("searchInput").value = searchText;
    $("searchClear").hidden = !searchText;
    $("monthFilter").value = monthFilter;
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
    searchText = "";
    $("searchInput").value = "";
    $("searchClear").hidden = true;
    $("monthFilter").value = "";
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
      const sw = document.createElement("span");
      sw.className = "swatch";
      sw.style.background = ownerColor(e.oid);
      btn.appendChild(sw);
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
      dl.innerHTML = "";
      Array.from(set).sort().forEach((v) => {
        const opt = document.createElement("option");
        opt.value = v;
        dl.appendChild(opt);
      });
    }
    fill("customerList", customers);
    fill("licTypeList", licTypes);
    fill("productLineList", productLines);
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
    renderExpiryFilterMenu();
    renderOwnerChips();
    renderProductChips();
    renderStatusChips();
    renderActiveFilterBar();
    renderSavedViews();
    renderActivityStrip();
    populateCalYearJump();
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
    if (currentView === "calendar") {
      $("tableView").hidden = true;
      $("calendarView").hidden = false;
      renderCalendar();
    } else {
      $("tableView").hidden = false;
      $("calendarView").hidden = true;
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

    const tdCustomer = document.createElement("td");
    tdCustomer.className = "col-customer";
    const custBtn = document.createElement("button");
    custBtn.type = "button";
    custBtn.className = "customer-link";
    custBtn.textContent = lic.customer || "";
    custBtn.addEventListener("click", (e) => { e.stopPropagation(); openCustomerDialog(lic.customer); });
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
      if (d < 0) { badge.classList.add("overdue"); badge.textContent = `${-d}d overdue`; }
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

    const tdProd = document.createElement("td");
    tdProd.className = "col-productline";
    if (lic.productLine) {
      const tag = document.createElement("span");
      tag.className = "product-tag";
      tag.textContent = lic.productLine;
      tdProd.appendChild(tag);
    }
    tr.appendChild(tdProd);

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
        for (const d of state.values) {
          const chip = document.createElement("span");
          chip.className = "lead-chip";
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
  function openAddDialog() {
    editingId = null;
    $("licDialogTitle").textContent = "Add license";
    $("licCustomer").value = "";
    $("licType").value = "";
    $("licUsers").value = "1";
    $("licExpiry").value = todayPh();
    ensureLicOwnerPicker(me.oid || null, me.name || "");
    $("licProductLine").value = "";
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
    $("licProductLine").value = lic.productLine || "";
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
      const head = document.createElement("div");
      head.className = "lic-comment-head";
      const who = document.createElement("span");
      who.className = "lic-comment-who";
      who.textContent = c.byName || "(unknown)";
      const when = document.createElement("span");
      when.className = "lic-comment-when";
      when.textContent = c.at ? new Date(c.at).toLocaleString(undefined, { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }) : "";
      head.appendChild(who); head.appendChild(when);
      const body = document.createElement("div");
      body.className = "lic-comment-body";
      body.textContent = c.text || "";
      li.appendChild(head); li.appendChild(body);
      list.appendChild(li);
    }
    // Auto-scroll to newest comment
    list.scrollTop = list.scrollHeight;
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
        const { license } = await api("PATCH", `/licenses/${editingId}`, payload);
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
      api("DELETE", `/licenses/${lic.id}`).catch((err) => {
        // If server delete fails, restore.
        showError("Delete failed (restored locally)", err);
        licenses.push(lic);
        render();
      });
      pendingDeletes.delete(lic.id);
      t.hidden = true;
    }, 5000);
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
      const { license } = await api("POST", `/licenses/${renewTargetId}/renew`, body);
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
      showError("Renewal failed", err);
    }
  }

  // v1.7.39 — inline-row +1y renew. Optimistic UI + Undo via toast.
  async function quickRenewOneYear(lic) {
    const before = { ...lic, comments: lic.comments ? [...lic.comments] : [] };
    try {
      const { license } = await api("POST", `/licenses/${lic.id}/renew`, { years: 1 });
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
    } catch (err) { showError("Renewal failed", err); }
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
        try { localStorage.setItem(LS_QUICK, quickFilter); } catch (_) {}
        render();
      });
    });

    // v1.7.38 Month dropdown
    $("monthFilter").addEventListener("change", (e) => {
      monthFilter = e.target.value || "";
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

    // Search
    const searchInput = $("searchInput");
    searchInput.addEventListener("input", () => {
      searchText = searchInput.value.trim();
      $("searchClear").hidden = !searchText;
      render();
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

    // Quick guide
    $("licGuideBtn").addEventListener("click", () => $("licGuideDialog").showModal());
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

    // Customer edit dialog
    $("custSaveBtn").addEventListener("click", saveCustomer);
    $("custCancelBtn").addEventListener("click", () => { $("customerEditDialog").close(); editingCustomerId = null; });

    // Email templates editor
    $("emailTemplatesBtn").addEventListener("click", openTemplatesDialog);
    $("tplSaveBtn").addEventListener("click", saveTemplate);
    $("tplDeleteBtn").addEventListener("click", deleteTemplate);
    $("tplNewBtn").addEventListener("click", newTemplate);
    $("templatesCloseBtn").addEventListener("click", () => $("templatesDialog").close());

    // Bulk select + reassign
    $("bulkSelectBtn").addEventListener("click", () => setBulkMode(!bulkMode));
    $("bulkReassignBtn").addEventListener("click", openBulkReassign);
    $("bulkClearBtn").addEventListener("click", () => { bulkSelected.clear(); updateBulkBar(); render(); });
    $("bulkReassignConfirm").addEventListener("click", confirmBulkReassign);
    $("bulkReassignCancel").addEventListener("click", () => $("bulkReassignDialog").close());

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
      const [licRes, memRes, custRes, tplRes] = await Promise.all([
        api("GET", "/licenses"),
        api("GET", "/members").catch(() => ({ members })),
        api("GET", "/customers").catch(() => ({ customers })),
        api("GET", "/email-templates").catch(() => ({ templates: emailTemplates })),
      ]);
      licenses = licRes.licenses || [];
      members = memRes.members || members;
      customers = custRes.customers || customers;
      emailTemplates = tplRes.templates || emailTemplates;
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
    el.textContent = `Updated ${fmtRelative(lastSyncedAt)}`;
    el.title = lastSyncedAt ? new Date(lastSyncedAt).toLocaleString() : "Not synced yet";
  }
  // Light background poll: refresh licenses every 60s when the tab is visible
  // so changes by another tenant user (Dona renews, Rey reassigns) appear
  // without a manual refresh. Skipped while a dialog is open to avoid clobbering
  // an in-progress edit. Errors are silent — keep-warm noise isn't worth a toast.
  async function pollOnce() {
    if (document.visibilityState !== "visible") return;
    const anyDialogOpen = document.querySelectorAll("dialog[open]").length > 0;
    if (anyDialogOpen) return;
    try {
      const res = await api("GET", "/licenses");
      const next = res.licenses || [];
      // Merge: keep any pending-undo items the server hasn't seen.
      const pendingIds = new Set([...pendingDeletes.keys()]);
      const filtered = next.filter((l) => !pendingIds.has(l.id));
      licenses = filtered;
      lastSyncedAt = Date.now();
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

  function exportCsv() {
    const list = sortLicenses(visibleLicenses());
    if (!list.length) { toast("Nothing to export with the current filters"); return; }
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
    a.download = `day-reminders-licenses-${todayPh()}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 100);
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
      const settingsPromise = api("GET", "/settings")
        .then((res) => { if (res && res.settings) userSettings = { ...userSettings, ...res.settings }; })
        .catch(() => {});
      const { licenses: lics } = await api("GET", "/licenses");
      licenses = lics || [];
      lastSyncedAt = Date.now();
      updateSyncIndicator();

      // Hide the boot indicator now that we have data.
      const bi = $("bootIndicator");
      if (bi) bi.classList.add("gone");
      render();
      // Don't await secondaries if still in flight; they just refresh state in background.
      void membersPromise; void customersPromise; void templatesPromise; void settingsPromise;
      // v1.7.39 — start the live poll once initial paint is done.
      startLivePoll();
      // v1.7.39 — wire Ctrl+K palette and global shortcuts.
      wireCmdk();
      wireGlobalShortcuts();
    } catch (err) {
      const bi = $("bootIndicator");
      if (bi) bi.classList.add("gone");
      showError("Could not connect", err);
    }
  }

  boot();
})();
