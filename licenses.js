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
  let importRows = []; // staged rows during CSV import

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

    // CSV import
    $("importCsvBtn").addEventListener("click", openImportDialog);
    $("licEmptyImport").addEventListener("click", openImportDialog);
    $("importFile").addEventListener("change", (e) => {
      const f = e.target.files && e.target.files[0];
      if (f) handleImportFile(f);
    });
    $("importCancelBtn").addEventListener("click", () => {
      $("importDialog").close();
      importRows = [];
    });
    $("importConfirmBtn").addEventListener("click", confirmImport);
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
    // Duplicate detection: same customer + licenseType + expiryDate as an existing license.
    const dup = licenses.find((l) =>
      l.customer.trim().toLowerCase() === row.customer.trim().toLowerCase() &&
      l.licenseType.trim().toLowerCase() === row.licenseType.trim().toLowerCase() &&
      l.expiryDate === row.expiryDate
    );
    if (dup) return { status: "duplicate", reason: "already in licenses" };
    if (!row.ownerOid) return { status: "needsOwner", reason: "owner not matched" };
    return { status: "ready" };
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
        const sel = document.createElement("select");
        sel.innerHTML = '<option value="">Pick owner…</option>';
        members.forEach((m) => {
          const opt = document.createElement("option");
          opt.value = m.oid;
          opt.textContent = m.displayName || m.upn || m.oid.slice(0, 8);
          sel.appendChild(opt);
        });
        sel.addEventListener("change", () => {
          const oid = sel.value;
          const owner = members.find((m) => m.oid === oid);
          row.ownerOid = oid || null;
          row.ownerName = owner ? owner.displayName : null;
          const cls = classifyRow(row);
          row.status = cls.status;
          row.reason = cls.reason || null;
          renderImportPreview();
          updateImportStats();
        });
        if (row.ownerRawName) {
          const hint = document.createElement("div");
          hint.className = "owner-hint";
          hint.textContent = `from CSV: "${row.ownerRawName}"`;
          tdOwner.appendChild(sel);
          tdOwner.appendChild(hint);
        } else {
          tdOwner.appendChild(sel);
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
    const needsOwner = importRows.filter((r) => r.status === "needsOwner").length;
    const invalid = importRows.filter((r) => r.status === "invalid").length;
    const dup = importRows.filter((r) => r.status === "duplicate").length;
    const parts = [];
    if (ready) parts.push(`${ready} ready`);
    if (needsOwner) parts.push(`${needsOwner} need owner`);
    if (dup) parts.push(`${dup} duplicate`);
    if (invalid) parts.push(`${invalid} invalid`);
    $("importStats").textContent = parts.join(" · ");
    const btn = $("importConfirmBtn");
    btn.disabled = ready === 0;
    btn.textContent = ready === 1 ? `Import 1 license` : `Import ${ready} licenses`;
  }

  async function confirmImport() {
    const bulkPL = $("importBulkProductLine").value.trim() || null;
    const toImport = importRows.filter((r) => r.status === "ready");
    if (!toImport.length) return;
    const btn = $("importConfirmBtn");
    btn.disabled = true;
    btn.textContent = "Importing…";
    let success = 0;
    let failed = 0;
    for (const row of toImport) {
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
        success++;
      } catch (err) {
        console.error("import row failed", row, err);
        failed++;
      }
    }
    $("importDialog").close();
    render();
    if (failed) {
      toast(`Imported ${success}, ${failed} failed (check console)`);
    } else {
      toast(`Imported ${success} license${success === 1 ? "" : "s"}`);
    }
    importRows = [];
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
