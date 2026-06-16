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
  let ownerFilter = null;   // null | ownerOid — set by owner breakdown chip
  let productFilter = null; // null | productLine string — set by product breakdown chip
  let statusFilter = null;  // null | status — set by status breakdown chip
  let groupBy = "none";     // 'none' | 'customer' | 'ownerName' | 'productLine'
  let searchText = "";
  let sortKey = "expiryDate";
  let sortDir = 1; // 1 asc, -1 desc

  let calCursor = startOfMonth(new Date());
  let editingId = null; // id of license being edited in licDialog; null = new
  let renewTargetId = null;
  let importRows = []; // staged rows during CSV import
  let licOwnerPicker = null;
  let bulkReassignPicker = null;
  // Precomputed once per render so every row in the same cluster shows the
  // same Bundle: N badge. Recomputed by computeBundles().
  let licenseBundles = new Map(); // licenseId -> { id, members[], size }

  // localStorage keys (tab-only UI state)
  const LS_VIEW = "lic.view";
  const LS_QUICK = "lic.quickFilter";
  const LS_SORT = "lic.sort";
  const LS_GROUP = "lic.groupBy";

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
  } catch (_) {}

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
  function toast(msg) {
    const t = $("toast");
    if (!t) return;
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(() => { t.hidden = true; }, 2400);
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
    // additive owner / product / status breakdown filters
    if (ok && ownerFilter && lic.ownerOid !== ownerFilter) ok = false;
    if (ok && productFilter && lic.productLine !== productFilter) ok = false;
    if (ok && statusFilter && lic.status !== statusFilter) ok = false;
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
    const map = new Map(); // ownerOid -> { name, count, nextExpiry }
    const today = todayPh();
    for (const l of licenses) {
      if (l.state === "abandoned") continue;
      if (!l.ownerOid) continue;
      const key = l.ownerOid;
      const cur = map.get(key) || { oid: l.ownerOid, name: l.ownerName || "(no name)", count: 0, nextExpiry: null };
      cur.count++;
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
    const entries = [...map.values()].sort((a, b) => b.count - a.count);
    for (const e of entries) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "lic-breakdown-chip";
      btn.setAttribute("aria-pressed", ownerFilter === e.oid ? "true" : "false");
      const sw = document.createElement("span");
      sw.className = "swatch";
      sw.style.background = ownerColor(e.oid);
      btn.appendChild(sw);
      btn.appendChild(document.createTextNode(e.name));
      const cnt = document.createElement("span");
      cnt.className = "chip-count";
      cnt.textContent = e.count;
      btn.appendChild(cnt);
      if (e.nextExpiry) {
        const nx = document.createElement("span");
        nx.className = "chip-next";
        const d = daysBetween(today, e.nextExpiry);
        nx.textContent = d === null ? "" : (d < 0 ? `overdue ${-d}d` : d === 0 ? "today" : `next ${fmtShortDate(e.nextExpiry)}`);
        btn.appendChild(nx);
      }
      btn.addEventListener("click", () => {
        ownerFilter = ownerFilter === e.oid ? null : e.oid;
        // Clicking an owner chip overrides the Mine quick-filter so the chip wins.
        if (ownerFilter && quickFilter === "mine") quickFilter = "all";
        render();
      });
      wrap.appendChild(btn);
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
    // Keep canonical pipeline order.
    for (const status of STATUSES) {
      const count = map.get(status);
      if (!count) continue;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "lic-breakdown-chip";
      btn.setAttribute("aria-pressed", statusFilter === status ? "true" : "false");
      const pill = document.createElement("span");
      pill.className = `status-pill status-${status}`;
      pill.textContent = STATUS_LABEL[status];
      btn.appendChild(pill);
      const cnt = document.createElement("span");
      cnt.className = "chip-count";
      cnt.textContent = count;
      btn.appendChild(cnt);
      btn.addEventListener("click", () => {
        statusFilter = statusFilter === status ? null : status;
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
    const entries = [...map.entries()].sort((a, b) => b[1] - a[1]);
    for (const [name, count] of entries) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "lic-breakdown-chip";
      btn.setAttribute("aria-pressed", productFilter === name ? "true" : "false");
      btn.appendChild(document.createTextNode(name));
      const cnt = document.createElement("span");
      cnt.className = "chip-count";
      cnt.textContent = count;
      btn.appendChild(cnt);
      btn.addEventListener("click", () => {
        productFilter = productFilter === name ? null : name;
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
    renderOwnerChips();
    renderProductChips();
    renderStatusChips();
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

    const tdCustomer = document.createElement("td");
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
    tdType.textContent = lic.licenseType || "";
    tr.appendChild(tdType);

    const tdUsers = document.createElement("td");
    tdUsers.className = "num";
    tdUsers.textContent = lic.userCount || 0;
    tr.appendChild(tdUsers);

    const tdExpiry = document.createElement("td");
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
    }
    tr.appendChild(tdExpiry);

    const tdOwner = document.createElement("td");
    const pill = document.createElement("span");
    pill.className = "owner-pill";
    pill.style.background = ownerColor(lic.ownerOid);
    pill.textContent = lic.ownerName || (lic.ownerOid ? lic.ownerOid.slice(0, 8) : "(none)");
    tdOwner.appendChild(pill);
    tr.appendChild(tdOwner);

    const tdProd = document.createElement("td");
    if (lic.productLine) {
      const tag = document.createElement("span");
      tag.className = "product-tag";
      tag.textContent = lic.productLine;
      tdProd.appendChild(tag);
    }
    tr.appendChild(tdProd);

    // Status pill column
    const tdStatus = document.createElement("td");
    const sPill = document.createElement("span");
    const statusVal = lic.status || "notStarted";
    sPill.className = `status-pill status-${statusVal}`;
    sPill.textContent = STATUS_LABEL[statusVal] || statusVal;
    tdStatus.appendChild(sPill);
    tr.appendChild(tdStatus);

    const tdActions = document.createElement("td");
    tdActions.className = "actions";
    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "btn ghost small";
    editBtn.textContent = "Edit";
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
    renewBtn.title = "Mark this license renewed";
    renewBtn.addEventListener("click", (e) => { e.stopPropagation(); openRenewDialog(lic.id); });
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

  function buildGroupHeaderRow(label, group) {
    const tr = document.createElement("tr");
    tr.className = "lic-group-header";
    const td = document.createElement("td");
    td.colSpan = bulkMode ? 9 : 8;
    const seats = group.reduce((s, l) => s + (typeof l.userCount === "number" ? l.userCount : 0), 0);
    td.textContent = label || "(none)";
    const meta = document.createElement("span");
    meta.className = "group-count";
    const noun = group.length === 1 ? "license" : "licenses";
    meta.textContent = `${group.length} ${noun} · ${seats.toLocaleString()} seats`;
    td.appendChild(meta);
    tr.appendChild(td);
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
        tbody.appendChild(buildGroupHeaderRow(key, group));
        for (const lic of group) tbody.appendChild(buildLicenseRow(lic, today));
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
        if (lic.state === "abandoned") pill.classList.add("abandoned");
        pill.style.background = ownerColor(lic.ownerOid);
        const text = `${lic.customer} · ${lic.licenseType}`;
        pill.textContent = text.length > 40 ? text.slice(0, 39) + "…" : text;
        pill.title = `${lic.customer} — ${lic.licenseType} (${lic.userCount} users, owner: ${lic.ownerName || "—"})`;
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
    $("licLeadDays").value = "";
    $("licLeadDaysCustom").value = "";
    $("licLeadDaysCustom").hidden = true;
    $("licNotes").value = "";
    $("licActivity").hidden = true;
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
    if (lic.leadDays === null || lic.leadDays === undefined) {
      $("licLeadDays").value = "";
      $("licLeadDaysCustom").hidden = true;
    } else if ([7, 14, 30, 60, 90].includes(lic.leadDays)) {
      $("licLeadDays").value = String(lic.leadDays);
      $("licLeadDaysCustom").hidden = true;
    } else {
      $("licLeadDays").value = "custom";
      $("licLeadDaysCustom").hidden = false;
      $("licLeadDaysCustom").value = lic.leadDays;
    }
    $("licNotes").value = lic.notes || "";
    renderActivityLog(lic);
    $("licEmailBtn").hidden = false;
    $("licRenewBtn").hidden = false;
    $("licDeleteBtn").hidden = false;
    $("licDialog").showModal();
    $("licCustomer").focus();
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
    const leadSel = $("licLeadDays").value;
    let leadDays = null;
    if (leadSel === "custom") {
      const c = parseInt($("licLeadDaysCustom").value, 10);
      if (isFinite(c) && c >= 0 && c <= 365) leadDays = c;
    } else if (leadSel !== "") {
      leadDays = parseInt(leadSel, 10);
    }
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

    if (isBundle) toast(`Drafted email for ${bundle.length}-license bundle.`);

    // Mark notStarted licenses in the bundle as noticeSent (click implies you sent the notice).
    const toMark = bundle.filter((b) => b.status === "notStarted");
    if (toMark.length === 1) {
      api("PATCH", `/licenses/${toMark[0].id}`, { status: "noticeSent" })
        .then(({ license }) => { licenses = licenses.map((l) => l.id === license.id ? license : l); render(); })
        .catch((err) => showError("Status update failed", err));
    } else if (toMark.length > 1) {
      api("POST", "/licenses/bulk", { ids: toMark.map((b) => b.id), patch: { status: "noticeSent" } })
        .then(({ updated }) => {
          const m = new Map(updated.map((l) => [l.id, l]));
          licenses = licenses.map((l) => m.get(l.id) || l);
          render();
        })
        .catch((err) => showError("Bulk status update failed", err));
    }
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

    // CSV export
    $("exportCsvBtn").addEventListener("click", exportCsv);

    // CSV import + split-button menu
    $("importCsvBtn").addEventListener("click", openImportDialog);
    $("licEmptyImport").addEventListener("click", openImportDialog);
    const menuToggle = $("importCsvMenuToggle");
    const menu = $("importCsvMenu");
    menuToggle.addEventListener("click", (e) => {
      e.stopPropagation();
      const isHidden = menu.hidden;
      menu.hidden = !isHidden;
      menuToggle.setAttribute("aria-expanded", String(isHidden));
    });
    document.addEventListener("click", (e) => {
      if (!menu.contains(e.target) && e.target !== menuToggle) {
        menu.hidden = true;
        menuToggle.setAttribute("aria-expanded", "false");
      }
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape") { menu.hidden = true; menuToggle.setAttribute("aria-expanded", "false"); }
    });
    $("downloadTemplateBtn").addEventListener("click", () => {
      downloadCsvTemplate();
      menu.hidden = true;
      menuToggle.setAttribute("aria-expanded", "false");
    });
    $("downloadTemplateBtnInDialog").addEventListener("click", downloadCsvTemplate);
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

    // Custom lead days toggle
    $("licLeadDays").addEventListener("change", () => {
      const v = $("licLeadDays").value;
      $("licLeadDaysCustom").hidden = v !== "custom";
      if (v === "custom") $("licLeadDaysCustom").focus();
    });

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
        l.leadDays == null ? "" : l.leadDays,
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
    document.body.dataset.theme = theme === "dark" ? "dark" : theme === "contrast" ? "contrast" : "default";
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
      const { licenses: lics } = await api("GET", "/licenses");
      licenses = lics || [];

      // Hide the boot indicator now that we have data.
      const bi = $("bootIndicator");
      if (bi) bi.classList.add("gone");
      render();
      // Don't await secondaries if still in flight; they just refresh state in background.
      void membersPromise; void customersPromise; void templatesPromise;
    } catch (err) {
      const bi = $("bootIndicator");
      if (bi) bi.classList.add("gone");
      showError("Could not connect", err);
    }
  }

  boot();
})();
