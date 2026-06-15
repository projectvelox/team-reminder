/* Day Reminders — Licenses tab (v1.7.0)
   Tenant-shared license tracker with Table + Calendar (Month) views.
   Auth via Teams SSO. Server is source of truth; we only cache for the session.
*/
(function () {
  "use strict";

  const API_BASE = "https://func-day-reminders-17023.azurewebsites.net/api";
  const DEFAULT_LEAD_DAYS = 14;
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
  let me = { oid: null, name: null };
  let authToken = null;
  let teamsTheme = "default";

  let currentView = "table";
  let summaryFilter = null; // null | 'week' | 'month' | 'overdue'
  let quickFilter = "all";  // 'all' | 'mine' | 'month' | 'overdue'
  let searchText = "";
  let sortKey = "expiryDate";
  let sortDir = 1; // 1 asc, -1 desc

  let calCursor = startOfMonth(new Date());
  let editingId = null; // id of license being edited in licDialog; null = new
  let renewTargetId = null;

  // localStorage keys (tab-only UI state)
  const LS_VIEW = "lic.view";
  const LS_QUICK = "lic.quickFilter";
  const LS_SORT = "lic.sort";

  try {
    const v = localStorage.getItem(LS_VIEW);
    if (v === "table" || v === "calendar") currentView = v;
    const q = localStorage.getItem(LS_QUICK);
    if (["all", "mine", "month", "overdue"].includes(q)) quickFilter = q;
    const s = localStorage.getItem(LS_SORT);
    if (s) {
      const [k, d] = s.split(":");
      if (k) sortKey = k;
      if (d === "1" || d === "-1") sortDir = Number(d);
    }
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
    if (f === "mine") return lic.ownerOid === me.oid;
    if (f === "month") {
      if (!lic.expiryDate) return false;
      const exp = parseISO(lic.expiryDate);
      const now = new Date();
      return exp && exp.getUTCFullYear() === now.getFullYear() && exp.getUTCMonth() === now.getMonth();
    }
    if (f === "overdue") return d !== null && d < 0 && lic.state !== "abandoned";
    if (f === "week") return d !== null && d >= 0 && d <= 7;
    return true; // all
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

  function refreshOwnerSelect() {
    const sel = $("licOwner");
    const current = sel.value;
    sel.innerHTML = '<option value="">Pick a teammate</option>';
    members.forEach((m) => {
      const opt = document.createElement("option");
      opt.value = m.oid;
      opt.textContent = m.displayName || m.upn || m.oid.slice(0, 8);
      sel.appendChild(opt);
    });
    if (current) sel.value = current;
  }

  // ---------- render ----------
  function render() {
    recomputeSummary();
    refreshDataLists();
    document.querySelectorAll(".view-btn").forEach((b) => {
      b.setAttribute("aria-pressed", b.dataset.view === currentView ? "true" : "false");
    });
    document.querySelectorAll(".filter-pill").forEach((p) => {
      p.setAttribute("aria-pressed", (summaryFilter ? "false" : (p.dataset.quick === quickFilter ? "true" : "false")));
    });
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

  function renderTable() {
    const list = sortLicenses(visibleLicenses());
    const tbody = $("licTbody");
    const empty = $("licEmpty");
    const tbl = $("licTable");
    tbody.innerHTML = "";
    if (!list.length) {
      tbl.hidden = true;
      empty.hidden = false;
      return;
    }
    tbl.hidden = false;
    empty.hidden = true;
    const today = todayPh();
    for (const lic of list) {
      const tr = document.createElement("tr");
      tr.dataset.id = lic.id;
      const d = daysBetween(today, lic.expiryDate);
      if (d !== null && d < 0 && lic.state !== "abandoned") tr.classList.add("overdue");
      if (lic.state === "abandoned") tr.classList.add("abandoned");

      const tdCustomer = document.createElement("td");
      tdCustomer.textContent = lic.customer || "";
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
        if (d < 0) {
          badge.classList.add("overdue");
          badge.textContent = `${-d}d overdue`;
        } else if (d === 0) {
          badge.textContent = "today";
        } else {
          badge.textContent = `${d}d left`;
        }
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

      const tdActions = document.createElement("td");
      tdActions.className = "actions";
      const editBtn = document.createElement("button");
      editBtn.type = "button";
      editBtn.className = "btn ghost small";
      editBtn.textContent = "Edit";
      editBtn.addEventListener("click", (e) => { e.stopPropagation(); openEditDialog(lic); });
      tdActions.appendChild(editBtn);
      const renewBtn = document.createElement("button");
      renewBtn.type = "button";
      renewBtn.className = "btn primary small";
      renewBtn.textContent = "Renewed";
      renewBtn.addEventListener("click", (e) => { e.stopPropagation(); openRenewDialog(lic.id); });
      tdActions.appendChild(renewBtn);
      tr.appendChild(tdActions);

      tr.addEventListener("click", () => openEditDialog(lic));
      tbody.appendChild(tr);
    }
    // Indicate active sort column.
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
    refreshOwnerSelect();
    $("licOwner").value = me.oid || "";
    $("licProductLine").value = "";
    $("licLeadDays").value = "";
    $("licLeadDaysCustom").value = "";
    $("licLeadDaysCustom").hidden = true;
    $("licNotes").value = "";
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
    refreshOwnerSelect();
    $("licOwner").value = lic.ownerOid || "";
    $("licProductLine").value = lic.productLine || "";
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
    $("licRenewBtn").hidden = false;
    $("licDeleteBtn").hidden = false;
    $("licDialog").showModal();
    $("licCustomer").focus();
  }
  function closeEditDialog() {
    $("licDialog").close();
    editingId = null;
  }

  function readLicenseForm() {
    const ownerOid = $("licOwner").value || null;
    const owner = members.find((m) => m.oid === ownerOid);
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
      ownerName: owner ? (owner.displayName || owner.upn || "") : null,
      productLine: $("licProductLine").value.trim() || null,
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
  async function deleteLicense() {
    if (!editingId) return;
    if (!confirm("Delete this license? This cannot be undone.")) return;
    try {
      await api("DELETE", `/licenses/${editingId}`);
      licenses = licenses.filter((l) => l.id !== editingId);
      closeEditDialog();
      render();
      toast("Deleted");
    } catch (err) {
      showError("Delete failed", err);
    }
  }

  // ---------- renew dialog ----------
  function openRenewDialog(id) {
    renewTargetId = id;
    const lic = licenses.find((l) => l.id === id);
    if (!lic) return;
    // Default custom date = current expiry + 1 year
    if (lic.expiryDate) {
      const d = parseISO(lic.expiryDate);
      if (d) {
        d.setUTCFullYear(d.getUTCFullYear() + 1);
        $("renewCustomDate").value = d.toISOString().slice(0, 10);
      }
    }
    $("renewDialog").showModal();
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
    // View switcher
    document.querySelectorAll(".view-btn").forEach((b) => {
      b.addEventListener("click", () => {
        currentView = b.dataset.view;
        try { localStorage.setItem(LS_VIEW, currentView); } catch (_) {}
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

    // Add/edit dialog buttons
    $("licSaveBtn").addEventListener("click", saveLicense);
    $("licDeleteBtn").addEventListener("click", deleteLicense);
    $("licCancelBtn").addEventListener("click", closeEditDialog);
    $("licRenewBtn").addEventListener("click", () => {
      if (editingId) openRenewDialog(editingId);
    });

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

  // ---------- Teams init + boot ----------
  function applyTheme(theme) {
    teamsTheme = theme;
    document.body.dataset.theme = theme === "dark" ? "dark" : theme === "contrast" ? "contrast" : "default";
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

      // First call hits /api/members to register self and grab the picker list.
      // Second call gets the licenses.
      const [{ members: mems }, { licenses: lics }] = await Promise.all([
        api("GET", "/members"),
        api("GET", "/licenses"),
      ]);
      members = mems || [];
      licenses = lics || [];

      wireEvents();
      render();
    } catch (err) {
      showError("Could not connect", err);
    }
  }

  boot();
})();
