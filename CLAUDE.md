# Day Reminders

A Microsoft Teams personal tab + bot for daily reminders, with proactive Adaptive Card notifications in the user's chat with the bot. Built for Kation Technologies; sideloaded as a custom org app. As of v1.7 also includes a second tab for tenant-shared license renewal tracking.

## Why it exists

The user (Kation owner) wanted a lightweight personal productivity surface that lives inside Teams so it gets attention without context-switching. Notifications had to work even when the tab wasn't focused, which ruled out a static tab-only design and pushed us to a bot for proactive Adaptive Cards. Tab is the management UI; bot is the notification channel.

## Architecture (current)

```
projectvelox.github.io/team-reminder/  ──►  Teams static tabs (Reminders + Licenses)
                                       └──►  Outlook taskpane (outlook.html/.js)
                                              │  Teams SSO Bearer  /  Office SSO Bearer
                                              ▼
func-day-reminders-17023.azurewebsites.net  ──►  Azure Function App (Node 22, Linux Consumption, East Asia)
  /api/messages       ◄── Bot Service ──► Teams chat (proactive Adaptive Cards)
  /api/reminders      ── CRUD (per-user)
  /api/settings       ── per-user prefs
  /api/licenses       ── CRUD + bulk + renew  (tenant-shared, v1.7)
  /api/members        ── auto-registering Owner picker source (v1.7)
  /api/customers      ── CRUD (tenant-shared customer registry, v1.7.9)
  /api/email-templates── per-product-line renewal email templates (v1.7.9)
  /api/users/search   ── Graph-backed people search (v1.7.14, User.Read.All app)
  /api/users/{oid}/photo── Graph photo proxy with in-process 60s cache (v1.7.14)
  scheduler           ── Timer trigger every minute (in Asia/Manila wall-clock)
                                              │
                                              ▼
                                         Azure Table Storage  (table: dayreminders)
                                              Per-user partitions (PK = user oid):
                                                RK = `_user`        user metadata
                                                RK = `_templates`   saved templates
                                                RK = `r:<id>`       a reminder
                                              Tenant-shared partitions (v1.7):
                                                PK = `_licenses`,      RK = `l:<id>`
                                                PK = `_members`,       RK = `m:<oid>`
                                                PK = `_customers`,     RK = `c:<id>` (v1.7.9)
                                                PK = `_emailTemplates`,RK = `t:<productLine>` (v1.7.9)
```

All Azure resources live in `rg-day-reminders`, except the Bot Service which is `global`. Storage and App Insights are in `southeastasia`, the Function App is in `eastasia` (Linux Consumption in `southeastasia` was stuck in 503 at creation; we recreated in `eastasia`). A keep-warm Logic App (`la-day-reminders-keepwarm`, southeastasia) pings `/api/ping` every 5 min during PH work hours so the Function App's HTTP triggers don't cold-start.

Secrets, GUIDs, and connection strings live in the Claude memory file `project_day_reminders_secrets.md` — **never in this repo**.

## Repo map

| Path | What |
|---|---|
| `index.html` / `app.js` / `styles.css` | The Reminders tab, served from GitHub Pages |
| `licenses.html` / `licenses.js` / `licenses.css` | The Licenses tab (v1.7). Shares `styles.css` for base theming. |
| `outlook.html` / `outlook.js` / `outlook.css` | Outlook add-in taskpane (new Outlook for Windows + OWA). Pre-fills a reminder from the open email, calls the same `/api/reminders`. |
| `manifest.json` | **Unified Microsoft 365 manifest** (v1.19) — single file covering tab + bot + composeExtension + Outlook add-in. Sideloaded via Teams admin center. |
| `outlook-manifest.xml` | Legacy Office Add-in XML manifest, kept as **per-user OWA sideload fallback only**. The unified `manifest.json` is the primary deploy path. |
| `color.png` / `outline.png` | App icons |
| `build.ps1` | Repackages `dist/team-reminder.zip` for sideload |
| `bot/package.json`, `bot/host.json` | Function App project |
| `bot/src/index.js` | Entry — registers all functions |
| `bot/src/functions/messages.js` | Bot endpoint (POST /api/messages) |
| `bot/src/functions/reminders.js` | CRUD for /api/reminders |
| `bot/src/functions/licenses.js` | CRUD + bulk + renew for /api/licenses (v1.7) |
| `bot/src/functions/members.js` | Auto-registering Owner picker source (v1.7) |
| `bot/src/functions/customers.js` | CRUD for /api/customers (v1.7.9) — annotation layer (contact emails, address, cross-license notes) |
| `bot/src/functions/emailTemplates.js` | GET/PUT/DELETE for /api/email-templates (v1.7.9) — per-product-line renewal email templates |
| `bot/src/functions/userSearch.js` | GET /api/users/search and /api/users/{oid}/photo (v1.7.14) — Graph-backed people picker source |
| `bot/src/functions/settings.js` | GET/PUT /api/settings |
| `bot/src/functions/scheduler.js` | Timer trigger — lead-time + EOD check-in |
| `bot/src/functions/ping.js` | Unauthenticated `GET /api/ping` returning `"ok"` — keep-warm target for the Logic App |
| `bot/src/lib/bot.js` | Bot adapter + `ReminderBot` activity handler + slash command parsing |
| `bot/src/lib/store.js` | Azure Tables wrappers |
| `bot/src/lib/auth.js` | Teams SSO JWT validation (jose + Entra JWKS) |
| `bot/src/lib/cards.js` | Adaptive Card templates |
| `dist/` | Build output, gitignored |

## What ships today (v1.7.x)

### v1.7.44 — UX/UI polish bundle (both tabs)

Closes the gaps from the v1.7.42 UX audit. Tab-only release. No backend.

**Reminders tab — closing the parity gap with Licenses**
- **Next-up hero** card at top of tab. Surfaces the single next timed reminder (today after now, or tomorrow's first) with countdown ("in 32 min", "tomorrow at 9:00 AM"). Click to scroll-and-flash the underlying row.
- **Keyboard shortcuts overlay** (`?` opens). Parity with Licenses. Shortcuts wired: `/`/`f` (search), `n` (new), `s` (select), `t` (templates), `v` (cycle views), `g` (cycle group), `?` (this overlay). Suppressed inside inputs.
- **First-time drag hint** banner on Week/Day views. Dismissible; once-only via localStorage. Solves the "drag-to-reschedule exists but no one knew" discoverability problem.
- **Bulk Undo toast** for mark-done AND delete. Mark-done sends a revert PATCH on Undo; delete uses optimistic remove + 6s commit timer.

**Licenses tab — UX audit fixes**
- **Sticky table header**. `thead` stays put when you scroll 30+ rows. Two lines of CSS, massive ergonomic win.
- **Comments thread moved to TOP of the Edit dialog** (was below 8 form rows). Comments are the most-changed field — surfaced.
- **Click-to-drill on the renewal-rate chart**. Click any month → filter table to that YYYY-MM and toast confirmation.
- **Click-to-filter on leaderboard owner names**. Click → owner filter set to that person; sidebar chip activates.
- **Sidebar accordion grouping** — 9 stacked sections collapsed into 4 `<details>` groups: *Numbers* (summary + stats, open by default), *Breakdowns* (status/owner/product chips, open), *Trends* (renewal trend + rate + leaderboard, collapsed), *Activity* (collapsed). Sidebar fits above the fold again.
- **Avatar consistency** — sidebar Owner chips now use the same round photo/initials avatar as the table row, instead of a flat color swatch. Same person reads as same person across surfaces.
- **Customer hover preview** — 250ms hover on any customer name → fixed-positioned card with licenses count, seats, owners, next expiry, overdue badge. Position-aware (flips left + clamps to viewport).

Cache-bust `v=1.7.44` on Licenses, `v=1.5.4` on Reminders.

### v1.7.43 — compliance bundle (soft-delete, GDPR DSR, tests, CI, types)

Closes the top gaps identified in the v1.7.42 compliance audit. Tab + bot + repo infrastructure.

**Backend — soft-delete with 30-day retention**
- `store.deleteLicense(id, actor)` now stamps `deletedAt` + `deletedByOid/Name` instead of hard-deleting. `listLicenses()` / `getLicense()` hide soft-deleted rows by default; pass `{ includeDeleted: true }` to see them.
- New helpers: `listDeletedLicenses()`, `hardDeleteLicense(id)`, `restoreLicense(id)`.
- New endpoints: `GET /api/licenses/deleted` (recovery list), `POST /api/licenses/{id}/restore` (undo).
- New scheduler branch `processSoftDeletePurge` runs daily 3:00-3:10 AM PH, hard-deletes any row whose `deletedAt` is older than 30 days. Idempotent.

**Backend — GDPR data-subject rights (`bot/src/functions/meExport.js`)**
- `GET /api/me/export` returns a JSON dump of everything we store about the caller (settings, reminders, templates, hasConversationRef, owned license rows). Sets `Content-Disposition: attachment` so the browser downloads it.
- `DELETE /api/me/data` wipes the caller's reminders + templates + settings + conversationRef, then unassigns the caller as owner on every license row they own (license rows themselves stay — tenant artifacts).

**Backend — type safety + tests**
- `bot/jsconfig.json` + `bot/src/lib/types.js` with canonical JSDoc typedefs for `License`, `Reminder`, `Settings`, `UserRecord`, `LicenseComment`, `LicenseEvent`. `// @ts-check` on the new files; legacy files run loose-mode via `npm run typecheck`.
- `bot/test/` with Node's built-in test runner. 20 tests across `store-helpers.test.js` (parseLeadDays / serializeLeadDays edge cases) and `license-validation.test.js` (required fields, length caps, leadDays array coercion, status/cycle enums, PATCH merge semantics). All green.
- `npm run lint` / `npm test` / `npm run typecheck` scripts added.

**Repo infrastructure — CI + supply chain**
- `.github/dependabot.yml`: weekly npm updates for `bot/`, monthly for GitHub Actions. Minor + patch grouped into one PR.
- `.github/workflows/ci.yml`: on every push to `main` + every PR, runs syntax check + tests + `npm audit --audit-level=high` (fails the build on high/critical vulns). Tab JS also gets syntax-checked.

**Frontend**
- **Recently-deleted recovery view** — Settings dialog gets an *Open recovery…* button. Lists every soft-deleted license with "Deleted Ns ago by NAME · N days until permanent purge" and a Restore button per row.
- **Privacy & data section** in Settings: explanation copy + *More details* link to a full privacy dialog (what we store, where, how long, your controls, third parties), plus **Export my data** (downloads the `/api/me/export` JSON) and **Delete my account data…** (confirm dialog + type-DELETE prompt → calls `DELETE /api/me/data`).

Cache-bust `v=1.7.43` on Licenses. **Backend redeploy required.**

### v1.7.42 — design language + shareable filters + strategic reporting

Polish pass to close the gap toward Linear/Stripe-tier internal tools. Tab-only release; no backend changes.

**Foundation**
- **Design-token system** in `styles.css`: `--space-{0..8}` (4px grid), `--text-{xs..xxl}`, `--radius-{sm/md/lg/xl/pill}`, semantic colors `--color-{danger/warning/caution/success/info/neutral}-{fg/bg/border}`, motion vars. Light/dark/contrast all override the semantic tokens. Legacy vars (`--bg`, `--accent`, etc.) preserved so existing CSS keeps working — new code MUST use tokens.
- `@media (prefers-reduced-motion: reduce)` globally suppresses shimmer / slide-in / transition animations (WCAG 2.3.3).

**Filtering**
- **URL hash state** — every filter axis encodes into `location.hash` (URLSearchParams; e.g. `#o=oid1,oid2&p=BC|M365&x=expired,soon&q2=acme`). Pasting a filtered link applies it on load. `hashchange` listener picks up back/forward. Implemented as `encodeFilterHash() / loadFiltersFromHash() / syncFiltersToHash()` with a `suppressHashSync` guard to prevent self-loops.
- **Date-range filter** (`from`/`to` inclusive). New popover in the filter row with two date inputs + Clear. Active selection shows in the chip strip as `Expires Jun 1, 2026 – Jul 15, 2026`. Wired into `matchesFilter`, persisted, URL-encoded, saved-view-snapshotted.

**Reports**
- **Renewal-rate trend chart** — 12-month time-series on a hand-drawn canvas (no Chart.js dependency). Buckets `lastRenewedAt` and lapsed (`expiryDate` in window AND not renewed/abandoned) per month, plots renewal-% line with filled area. Subtitle shows delta vs prior month (▲/▼ Npt). Auto-hidden if all months are empty. Uses computed-style colors so dark mode looks right.
- **Owner leaderboard** — sidebar widget ranking owners by total touched (renewed + lapsed + abandoned) over the last 90 days. Each row shows `Name · NN% · M of N renewed · K lapsed · J won't renew`. Color-coded score (green ≥80, amber ≥60, red below). Hidden when no events.

**Design**
- **Mobile/narrow viewport** layout (Licenses tab). `@media (max-width: 900px)` collapses the sidebar below the main pane, gives the table a horizontal scroll instead of crushed cells. `@media (max-width: 600px)` enlarges touch targets, wraps the stats strip.
- **Microcopy pass** — unified "Mark renewed" (was mix of "Mark as renewed" / "Mark renewed" / "Mark Renewed"). Status terms verified consistent.

What I deliberately skipped (would be over-engineering at this scale): facet preview counts, expression-language operators, i18n scaffolding (no second-language demand), cohort/funnel analytics (Power BI's job), scheduled custom reports.

Cache-bust `v=1.7.42`. No bot redeploy needed.

### v1.7.41 — concurrency, merge, bulk, dashboard, CSP

Five-feature bundle building on the v1.7.40 audit. Backend + frontend + headers.

- **Concurrent-edit conflict detection** — PATCH `/api/licenses/{id}` and POST `/api/licenses/{id}/renew` now honor an `If-Match` header carrying the row's `lastEditedAt`. Mismatch → 409 with the current row. The tab sends If-Match on every Save / Renew / quick-renew and surfaces conflicts as a toast *"Save blocked: Dona changed this row first."* with a Reload action. Cures silent overwrites in the tenant-shared dataset.
- **Customer merge** — new `POST /api/customers/merge` consolidates source customer rows into a target name: rewrites every license's `customer` string, absorbs source primary/secondary emails into the target's `secondaryEmails`, appends a "Merged from: X, Y" note, deletes the source registry rows, stamps a `customerMerged` event on each touched license. UI: **Merge in duplicates…** button on the customer drawer opens a candidate list (anything with ≥3-char token overlap) with per-row checkboxes.
- **Bulk operations beyond reassign** — bulk-mode bar gains **Renew +1y**, **Export CSV**, **Delete**. Renew loops through `/renew` with per-row If-Match guards so a stale row reports as "skipped" without aborting the batch. Export uses the shared CSV builder. Delete is optimistic with a single combined Undo toast (6s window).
- **Renewal-rate dashboard** — sidebar widget summarizing the last 90 days: `Renewal rate · 80% · 12 renewed · 2 lapsed · 1 won't renew`. Stacked bar by category; headline color-coded (green ≥80, amber ≥60, red below). Definitions: renewed = `lastRenewedAt` in window; lapsed = expiry fell in window AND not renewed/abandoned; abandoned = `state=abandoned` AND an `abandoned` event in window. Hidden when no rows fell in the window.
- **CSP + SRI** — both `licenses.html` and `index.html` get a `<meta http-equiv="Content-Security-Policy">` locking script-src to `'self' + res.cdn.office.net`, connect-src to `'self' + func-day-reminders-17023.azurewebsites.net`, frame-ancestors to the Teams host families. The Teams SDK `<script>` carries an `integrity="sha384-…"` hash + `crossorigin="anonymous"` so a compromised CDN can't inject. Future SDK version bumps require re-hashing.

Cache-bust `v=1.7.41` on Licenses, `v=1.5.3` unchanged on Reminders (only CSP+SRI added to its HTML; no JS/CSS code changes there). **Backend redeploy required** — If-Match handlers + merge endpoint.

### v1.7.40 — security + perf audit + dialog close X

Defense-in-depth + performance pass triggered by the user's "heavy security and optimization audit" ask. Plus a usability ask: every modal now has a corner **× close button** so users don't have to scroll to find Cancel or remember Esc.

**User-visible**
- Auto-injected `×` close button (top-right) on every `<dialog>` across both tabs — Edit license, Settings, Renew, Customer, CSV import, Templates, Bulk reassign, Day overflow, Guide, Command palette, Shortcuts, and the Reminders tab's Settings / What's new / Row options / Templates. Implemented as a one-shot `installDialogCloseButtons()` helper called after boot.

**Security (defense-in-depth)**
- **CORS lockdown**: `Access-Control-Allow-Origin: '*'` → `https://projectvelox.github.io` everywhere. Bearer-token auth meant CSRF was already N/A, but origin reflection prevents arbitrary pages from attaching a stolen token and reading our responses. Added `Vary: Origin`, `Access-Control-Max-Age: 600` for preflight caching.
- **Standard security headers** on every JSON response: `X-Content-Type-Options: nosniff`, `Referrer-Policy: no-referrer`, `Cache-Control: no-store`. Eleven endpoint files updated.
- **Storage-limit guard on comments**: pre-existing 100 × 2000 char cap could exceed Azure Tables' 32KB UTF-16 per-property string limit. Tightened to **30 × 1000 chars** so the JSON-encoded property stays well under. Cap applies on read + write + UI textarea maxlength.
- New `bot/src/lib/cors.js` with allowlist-based reflection helper, plus a no-arg drop-in style kept on each endpoint to minimize per-call-site refactor.

**Performance**
- **Search debounce** (150ms). Was firing `render()` on every keystroke; with 100+ licenses + multiple breakdowns + active-filter bar that adds up.
- **Skip no-op poll renders**: live poll now computes a cheap signature (`id|expiry|status|owner|lastEditedAt|commentCount|leadFireCount`) and bails out of `render()` when the server returned identical data. Cuts redundant re-paints every 60s.
- **Photo proxy cache headers**: `Cache-Control: private, max-age=86400, stale-while-revalidate=300` for 200 responses (was 1h); `private, max-age=300` for 204s. Profile photos rarely change — browsers will hit the proxy once a day per owner instead of every poll.

Cache-bust `v=1.7.40` on Licenses tab, `v=1.5.3` on Reminders tab (close-X helper added). **Backend redeploy required** (CORS + comment caps + photo cache headers).

### v1.7.39 — quality-of-life bundle

Big multi-feature release; tab + bot (comments + saved-views endpoints).

- **Saved filter views** — `+ Save view` button on the active-filter bar names the current filter combo and pins it as a chip above the bar. Stored per-user in `settings.savedLicenseViews` (max 25). Click chip = apply; chip `×` = delete.
- **Command palette** — `Ctrl+K` / `Cmd+K` opens a Linear/cmdk-style picker. Searches actions (Add license, Refresh, Open Settings, Switch view, Export CSV, …) AND entities (licenses + customers). Arrow keys navigate, Enter selects, Esc closes.
- **Keyboard shortcuts overlay** — `?` opens a help dialog. Active shortcuts: `Ctrl+K` (palette), `/` (focus search), `a` (add license), `q` (toggle quick-add), `r` (refresh), `g t` / `g c` (go to Table/Calendar), `Esc` (close any dialog).
- **Quick-renew (+1y)** on each row — clicking *Renew* now does +1y instantly with an **Undo** in the toast. Shift-click still opens the dialog for 2y / 3y / custom date.
- **Undo toast** — `toast()` extended with `{ actionLabel, onAction }`. Renew uses it; delete already had its own.
- **Owner avatars** — round photo (backed by the existing `/api/users/{oid}/photo` 60s-cached proxy) with initials-on-color fallback when no photo. Renders next to the owner pill in every row.
- **Live updates** — 60s background poll on `/api/licenses` when the tab is visible and no dialog is open. New header pill `Updated 12s ago` (refreshed every 10s without server hit) shows freshness. Pending-undo rows are excluded from poll overwrites.
- **Quick-add inline row** — collapsible `+ Quick add a license` strip above the table for Customer + License type + Expiry + Users without opening the full dialog. Owner defaults to current user.
- **Comments thread per license** — new `comments[]` field on the license schema (capped 100). Edit dialog gets a Comments section with newest-at-bottom + `Comment` button (Ctrl+Enter to send). Each row's Edit button shows a `💬 N` badge when there are comments. New endpoint: `POST /api/licenses/{id}/comments`.
- **Calendar density toggle** — `▦ / ▪` switch in the calendar nav. Compact halves the cell height + shrinks pills, for users tracking 100+ renewals per month.
- **Year jump** in calendar — dropdown of years with data (plus the next 3) so you can land on March 2027 in two clicks instead of 30 forward-arrows. "Same month next year" stays.
- **Hover-notes tooltip** on every table row (truncated at 600 chars). No more opening a dialog just to read a one-line context note.
- **Activity sidebar** — last 20 tenant-wide events ("Dona renewed Acme M365 · 3 min ago") in the right sidebar. Click any entry to open the relevant license. Sourced from `license.events[]` — no new endpoint.
- **Loading skeletons** — replaced the `Loading licenses…` text with 4 shimmer-rows so the layout doesn't jump on first paint.
- **Tab focus polish** — universal `:focus-visible` ring on every interactive element (buttons, inputs, comment thread, activity links) for keyboard nav.

Cache-bust to `v=1.7.39` on `licenses.html`/`licenses.js`/`licenses.css`. **Backend redeploy required** — new comments endpoint + license schema field + settings field.

### v1.7.38 — enterprise filter bar (Ella review)

Tab-only release. Addresses Ella's filter-UX review by adding the genuinely-new pieces (active-filters summary, month dropdown, 4-bucket expiry pill) without duplicating what was already shipped (sidebar breakdown chips, quick filters, sortable columns, owner count alongside unique customers).

- **Active-filters bar** above the table/calendar — GitHub Issues / Linear pattern. Shows `47 licenses` when no filters; switches to `12 of 47   [Owner: Dona ×] [Product: BC ×] [Expiry: Expiring soon ×] …   [Clear all]` once anything is selected. Removable per-chip + global Clear all. All filter axes (`quickFilter`, `summaryFilter`, `ownerFilter`, `productFilter`, `statusFilter`, `expiryFilter`, `monthFilter`, `searchText`) render through one builder so adding new dimensions later is one entry.
- **Month dropdown** in the quick-filter row. Populated dynamically with months that actually have expiring licenses (chrono order, `Any month` first). Selecting a month also jumps the Calendar view to it — single mental model across views.
- **Expiry pill** (4 buckets: Expired / Expiring soon / Expiring this month / Active). Rendered next to the days-left badge on every table row and as a left-edge stripe on calendar pills (so owner color is still the fill). Click any pill to filter by that bucket; toggle off by clicking again.
- **Expiry filter popover** in the filter row — Radix/cmdk-style: checkbox-per-bucket with the matching color dot. Live count summary on the trigger button (`Any` / `Expiring soon` / `2 selected`). Closes on click-outside or Escape.
- **"Soon" threshold is dynamic, not hardcoded**: tied to the user's smallest `licenseLeadDays` (default `7`). A user whose lead-days are `[60,30,15,7,1]` sees `Expiring soon` = ≤7d; switching their default to `[14]` makes it ≤14d. Matches what's actually being notified about.
- **Filter persistence**: every filter axis writes through to localStorage and re-hydrates on next load. Filters are preserved across Table↔Calendar (already true; just verified).
- **Empty-results swap**: when filters hide every row, the "No licenses yet" hero swaps in place to "No licenses match these filters" with a single **Clear all filters** button. No more thinking your data was wiped.
- **What got explicitly rejected from Ella's spec** (and why): (1) "Owner chip = unique customers only" — regresses [feedback_ella_unique_vs_record_counts](C:\Users\MSI%202\.claude\projects\c--Users-MSI-2-Documents-GitHub-team-reminder\memory\feedback_ella_unique_vs_record_counts.md). Workload (records) and concentration (customers) answer different questions; we keep both. (2) "Collapsible filter panel" — sidebar isn't dominating screen space at current scale; collapse traded discoverability for nothing. (3) Combobox-style Owner picker — overkill at ~4 owners; revisit when owner list > 8. (4) Renaming Ella's date-derived "Status" filter to "Expiry" to avoid colliding with the existing workflow `status` enum (notStarted/noticeSent/awaitingCustomer/customerConfirmed/renewed).

Cache-bust to `v=1.7.38` on `licenses.html`/`licenses.js`/`licenses.css`. No backend changes; no Function App redeploy needed.

### v1.7.37 — multi-threshold lead-day reminders

The single per-license "Notify me N days before expiry" dropdown is gone. Both the per-license override and the per-user default in Settings are now **arrays of lead-day thresholds** picked from a chip-style multi-select (presets `60 / 30 / 15 / 7 / 1` + `+ Custom…`).

- **Storage shape change**: `license.leadDays` and `settings.licenseLeadDays` are now `number[] | null`. Pre-v1.7.37 scalar rows are migrated on read (`[n]`). Both fields are stored as JSON strings in Azure Tables.
- **New per-license fields**: `lastFiredLeadDays: number[]` (which thresholds have already fired this renewal cycle, prevents re-firing) and `leadSnoozedUntil: YYYY-MM-DD | null` (set by the card's *Snooze 7 days* button to pause all renewal pings on a single license).
- **Cycle reset**: `lastFiredLeadDays` is cleared on expiry change (manual edit) and on `/api/licenses/{id}/renew`, so the new cycle's thresholds start fresh.
- **New scheduler branch `processLicenseLeadFires`**: weekday 9:00–9:10 AM PH window. For each license owned by the user, finds the smallest threshold `D` where `daysUntilExpiry <= D` and `D` is not in `lastFiredLeadDays`, sends `licenseLeadCard`, then stamps `D`. Respects `state=abandoned`, `status=renewed`, and per-license `leadSnoozedUntil`. Falls back to the user's Settings default if the license's own `leadDays` is null/empty; final fallback `[14]`.
- **New `licenseLeadCard`**: shows actual `daysUntilExpiry` ("expires in 30 days", "expires tomorrow", "expires today", "expired Nd ago"), owner + product line, with actions **Mark renewed (+1y)** / **Snooze 7 days** / **Won't renew** / **Open in tab**. "Won't renew" sets `state='abandoned'` and logs an event.
- **Bot card handlers**: new `licenseLeadSnooze` (sets `leadSnoozedUntil`) and `licenseWontRenew` (state=abandoned) in `bot.js`.
- **CSV export**: `Lead Days` column now writes the array as comma-joined values (e.g. `60,30,15,7,1`).

Cache-bust to `v=1.7.37` on `licenses.html`/`licenses.js`/`licenses.css`. Backend redeploy required (scheduler + card + validation changes).

### v1.7.0 — Licenses tab (this release)

Second static tab for tracking client license renewals as a tenant-shared dataset, distinct from per-user reminders.

- **New tab "Licenses"** alongside "Reminders". Table view (default) + Calendar (Month) view via top-bar toggle. Theme + base styles inherited from `styles.css`.
- **Tenant-shared storage**. New Azure Table partitions: `_licenses` (rows) + `_members` (the Owner picker source). Every authenticated tenant user can read and write all rows. Row has `customer, licenseType, userCount, expiryDate, ownerOid, ownerName, productLine, leadDays?, notes?, state, createdAt, lastEditedAt, lastRenewedAt`.
- **Owner picker** = self-populating member dropdown. Every `/api/licenses` and `/api/members` call auto-registers the caller's `oid + displayName`. Workaround that avoids needing `User.ReadBasic.All` Graph consent for Microsoft Graph user search.
- **Table view**: sortable columns (customer / type / users / expires / owner / product line). Days-left badge per row (`5d left`, `today`, `2d overdue`). Overdue rows get a red left border.
- **Calendar (Month) view**: Mon-Sun grid, Owner-colored pills per day (deterministic hash, separate palette from client chips), 3 visible + `+N more` overflow popover. Click pill = open edit; click empty cell = add license on that date. Nav: Prev / Today / Next + **Same month next year** jump.
- **Renewed action** with 1-year / 2-year / 3-year (triennial) presets or custom new-expiry date. Posts to `/api/licenses/{id}/renew`. Resets `lastEscalatedDays` so the escalation ladder starts fresh next cycle.
- **Summary chips** at top of tab: "N expiring this week / N this month / N overdue", click to filter the active view.
- **Quick filter chips**: All / Mine / This month / Overdue. Mirror of the Reminders quick-filter row.
- **Search box**: matches customer + license type + notes + product line + owner name.
- **Add/edit dialog** with autocomplete on customer + license type + product line (datalist sourced from existing rows), plus the Notify-N-days-before dropdown (7/14/30/60/90/Custom).
- **Manifest**: bumped to `1.7.0`, second `staticTab` entry `dayReminders.licenses` pointing at `licenses.html`. Cache-bust `?v=1.7.0` on the tab assets.

### v1.7.0 — deferred until Dei (Global Admin) consents

Needs Graph permission classified as high-privilege in Kation's consent policy. App reg manifest already has the perm requested; pending consent only.

- **Email surface** via `Mail.Send` (application), sent from `assist@kationtechnologies.com` (configurable via `LICENSE_DIGEST_FROM` app setting). Two uses:
  1. **Monthly digest** on the 1st of each month, per-owner list of expiring licenses.
  2. **Cold-owner nudge** (added 2026-06-16) — when a license is assigned to a tenant user who has never opened Day Reminders (no `conversationRef` on file), send a one-time email: "You've been assigned as owner of <customer> license. Open Day Reminders in Teams to get renewal alerts." Deduped per oid so they aren't spammed.

### Cold-owner handling (explicitly chose not to auto-install)

We **dropped** the `TeamsAppInstallation.ReadWriteForUser.All` ask (Dei's call, 2026-06-16). Granting a Function App the ability to silently install itself in any tenant user's Teams is too much blast radius if the repo or deploy pipeline is ever compromised. If we want every Kation user to have Day Reminders pre-installed, the right home for that is a Teams admin app-setup policy (admin-gated, one-time), not runtime code pulling from a public GitHub repo.

Fallback pattern that ships instead:
- **Cold-owner email nudge** (above) — covers the "owner never knew they were assigned" case.
- **Inline tab warning** — Licenses tab shows "⚠ This owner hasn't opened Day Reminders — they won't get Teams alerts until they do" on any row whose owner has no `conversationRef`. Same warning appears in the Add/Edit dialog the moment the assigner picks a cold owner.
- Once the owner opens Day Reminders once, the bot captures their `conversationRef` and escalation cards work normally from that point on.

The same pattern carries to v1.5 sharing (recipients without a `conversationRef` get the email-nudge + share-dialog banner, not a silent install).

### Other v1.7.0 backlog items (in scope, not yet built)

- Bulk reassign Owner (multi-select on table + Reassign action)
- CSV import with preview + per-row Owner resolver
- ICS calendar feed per user (`/api/licenses/ical?token=...`)
- Teams escalation cards in the scheduler (14d → 7d → 1d → daily, with Renewed / Won't renew buttons)
- Settings UI for per-user default `leadDays`

See `project_day_reminders_v17_licenses_scope.md` in Claude memory for the full locked scope.

## What ships today (v1.6.x)

- **Outlook add-in** (new Outlook for Windows + OWA). Open an email, click "Add reminder" in the message action bar (or right-click in the list → Apps → Day Reminders), taskpane slides in pre-filled with the subject as the title, sender in details, **date defaulted to tomorrow** (Tomorrow/Today quick-buttons, date picker available), optional time, client autocomplete. Submit creates the reminder in the same backend as the Teams tab. SSO via the same Entra app reg as Teams; backend audience check is unchanged.

### v1.4 baseline

- Add / delete / done-toggle / inline-edit (title, **date**, time, **description**, **client**)
- **`dueAt` per reminder** (date, defaults to today) with auto-rollover of undone past-due items (cap 30 days) and an *overdue Nd* badge
- **`description` field** per reminder (plain text, ≤2000 chars), rendered in the tab and in the proactive Adaptive Card
- **`client` field** per reminder (freeform, ≤100 chars). Autocompletes from past clients used by the user — no master list. When set, title displays with a `[Client]` prefix in every view (Lines, Grid, Day, Week) and in the proactive card. Also rendered as a dashed-outlined chip below the title — **colored per client** (deterministic hash → same client always gets the same color, seeded differently from tags so a name used as both client and tag doesn't collide visually). Click chip to filter (fills the chip with the client's color for strong feedback); Shift+click or right-click to inline-edit.
- **Week view** (Lines / Grid / Day / Week toggle in the top bar). Mon–Sun grid of reminders by `dueAt`. Today's column highlighted; timed items stack chronologically per day; anytime items sit at the bottom of each day under a sub-heading. Click an empty day to focus the add form pre-filled with that date. Prev / Today / Next nav above the grid. The pre-existing hour-by-hour view was renamed from *Calendar* to *Day*.
- **Group toggle** cycles *off → tag → client → off* (shortcut `g`). Sections per tag use the tag's chip color; sections per client use the client's chip color. Group choice persists across reloads in `localStorage`.
- Hashtag-in-title parses to tags; colored chips
- High-priority star, pinned to top of list
- Group-by-tag view toggle (acts as project sections)
- View toggle: Lines / Grid / Calendar
- Search (title + tags + client + **description**), quick filters (All / Timed / Anytime / Priority / Done), templates, bulk select
- Done items >24h hidden behind a "Show N older" button
- Slash commands in bot chat: `/add [time] [date] [#tags] title`, `/list`, `/done <substring>`, `/help` — `/add` accepts `today`, `tomorrow`, weekday names, `M/D`, or `YYYY-MM-DD`. `/done` matches title, tags, client, or description.
- Compose-extension command "Quick add reminder" — invocable from `...` menu under any Teams message box
- Proactive Adaptive Card per timed reminder with "Mark done" + snooze (15m / 1h / Tomorrow / +3 days / Next Mon)
- Snoozing to Tomorrow / +3 days / Next Mon advances `dueAt` so the rollover doesn't double-count
- EOD check-in card at configurable time
- Activity Feed notification when a reminder fires (Teams bell)

### v1.5 additions

- **Recurring reminders** — `repeat: none | daily | weekdays | weekly` per reminder. Set via the row options dialog. "Mark done" on a recurring reminder advances `dueAt` to the next occurrence and leaves it open. Recurring items are auto-advanced (not piled up as overdue) during the daily rollover.
- **Quiet hours** — per-user `quietStart` / `quietEnd` HH:MM in Settings. Scheduler short-circuits all proactive sends (lead-time, snooze fires, EOD) while in-window. Wrap-around windows (e.g. 22:00–07:00) are supported. The lead-time fire window is widened to `[target-lead, target+60min]` so a reminder whose original window fell entirely inside quiet hours still fires at most an hour late.
- **Per-user templates** — `GET / PUT /api/templates` stores up to 100 templates per user under RowKey `_templates`. Tab dialog shows "Your templates" above "Built-in". Row options dialog has a *Save as template* button. Each saved template captures title, time, client, description, leadMinutes, tags.
- **Bulk PATCH** — `POST /api/reminders/bulk` with `{ ids, patch }`. Tab's *Mark all done*, bulk done, and bulk priority now do a single round-trip instead of one PATCH per reminder.

### v1.5.1 — non-developer friendliness pass (this release)

Audit-driven release applying the new "target = non-developers" rule (see Claude memory `feedback_target_audience_non_devs.md`). Every dev-style syntax surface got a UI-control equivalent as the primary path; typing shortcuts remain for power users but never as the only way.

- **Sub-tasks (checklist)** — new structured field `subtasks: [{ id, text, done }]` per reminder (up to 50). Row options dialog has a Checklist editor section (add input + per-row checkbox + text input + delete; Enter inserts the next row; Backspace on empty removes). On the row itself, a "2/5" chip in title-meta and a "▸ Checklist (2/5)" chevron expand the list inline with click-to-toggle checkboxes that PATCH live. Proactive Adaptive Card renders sub-tasks as `☐` / `☑` lines. Replaced the earlier "could be markdown" instinct after user pushback on syntax-based UX.
- **Tag chip picker** — `<input>` + chip display below the add row and inside the row options dialog. Autocompletes from past tags via a shared `tagList` datalist. Enter / Tab / `,` commits a chip; Backspace on empty removes the last chip. Typing `#tag` inline in the title still works (merged with picker tags via `mergeTags`).
- **Lead time as duration dropdown** — Settings + row options now use `<select>` with presets (0 / 5 / 10 / 15 / 30 / 60 / 120 min) plus *Custom…* that reveals the existing number field. Internal storage is still `leadMinutes`, so no API change.
- **Bot fallback action card** — new `menuCard()` in `cards.js`. Sent on install, on `/help`, and as the fallback when the user @mentions the bot or sends non-command text. Buttons: *Add a reminder*, *What's on my list?*, *Mark something done*, *Show me how*. Handlers in `_handleCardAction`: `menuList` calls `_listOpen`, `menuHelp` re-sends the help card, `menuAdd` / `menuDone` print example-led hints.
- **Built-in templates cleanup** — `TEMPLATES` array now stores `tags: [...]` separately instead of `#tag` in the title. The tag still renders as a chip on the row via the same chip path.
- **Copy polish** — *Timed* filter pill → *Scheduled*. *+ Details* button → *+ Notes*. Title placeholder no longer mentions `#tags inline`. Empty-list hero replaced its hash-tag example seeds with clean titles + a *Browse templates* button. Quick guide rewritten to lead with the picker UI, not the syntax. Onboarding card reordered to introduce the picker before mentioning power-user shortcuts.

Cache-bust to `v=1.5.1` on `styles.css` and `app.js`.

## Backlog

- **v1.5 — sharing** (Rochelle's feedback 2026-06-15) — deferred behind the v1.5 QOL bundle that just shipped. Per-reminder share with any tenant member; recipients can edit content + mark done; creator owns the recipient list; deletion removes for everyone. Includes per-tag default share lists in Settings (e.g. `#QC = [Benex, Tim]` so any reminder tagged `#QC` auto-shares). Implementation must include proactive Graph install for recipients who haven't opened the bot. See `project_day_reminders_v15_sharing_scope.md` in Claude memory for the locked decisions. Likely v1.6 now.
- **v1.7 (conditional)** — Markdown formatting toolbar on description, only if Phase A (plain text) usage shows demand. Toolbar emits markdown that Adaptive Card's TextBlock renders natively.
- **Intra-day check-ins** (Dei's feedback EIL_0003). Settings supports a list of check-in times, not just one EOD. Each fires its own "how's it going?" card once per day. Deferred behind v1.5.
- **Outlook flag → auto-create reminder** (audit 2026-06-15) — scaffolded in v1.5 as `bot/src/functions/graphNotifications.js` (validation-token handshake + notification batch skeleton) and as `settings.autoImportFlagged` (default off, persisted but no UI yet). Not active in production: missing subscription creation/renewal flow, Graph delegated `Mail.Read` consent, OBO token exchange, and per-user subscription→oid mapping. Pick this up when ready to grant Mail.Read consent at the tenant level.
- **Workato sidecar for long-message intent parsing** (Josh's feedback EIL_0005). Backend gets an API-key auth path on `/api/reminders` (in addition to Teams SSO); a Workato recipe receives `@DayReminders <long message>` in Teams, parses via Workato AI, loops `POST /api/reminders` for each extracted item. ~30 min on our side; recipe lives in Kation's Workato workspace.

## Explicitly dropped

- **True shared-inbox groups** (a `#tag` with a real member list where any member can post and all members see all posts). User rejected 2026-06-15: "na i think it should be personal, else we're building a teams replica." Sharing in v1.5 ships as personal tag-default share lists instead (each user's own contact list).

## What we will NOT do

- **No direct LLM API contracts** (Azure OpenAI, Anthropic, OpenAI). Kation has a hard rule: no new paid infra. AI features go through Workato (which Kation already pays for) or M365 Copilot extensions (license they already pay for).
- **No localStorage as source of truth.** Storage is server-side. localStorage may be used for ephemeral UI state (`showAllDone`, `groupByTag`) but never for reminder data.
- **No breaking the manifest schema.** Always test sideload zip locally before pushing a manifest change.
- **No silent error swallowing in the UI.** Boot failures must surface the actual error string in the banner — never "something went wrong." This is how we caught the timezone bug.

## Operational gotchas (learned the hard way)

### 1. Teams desktop caches `app.js` / `styles.css` aggressively

**Every time you change those files, bump the `?v=X.Y.Z` query string in `index.html`.** Otherwise users see broken-button symptoms (Settings doesn't open, sections don't render, no error banner) because Teams serves stale JS against a fresh HTML structure. This took 3 cycles to debug the first time — don't repeat it.

Convention: bump the query string in lockstep with the manifest `version` for any tab-affecting change.

### 2. Linux Function Apps ignore `WEBSITE_TIME_ZONE`

That setting is Windows-only. On Linux the equivalent is `TZ=Asia/Manila`, **but** the Node runtime doesn't always pick it up reliably either. The reliable answer is to compute time-of-day via `Intl.DateTimeFormat` with `timeZone: 'Asia/Manila'` directly. See `phWallClock()` in `bot/src/functions/scheduler.js`. The TZ env var stays set as belt-and-suspenders but the code is the source of truth.

### 3. Azure Functions v4 forbids a body on HTTP 204

The Web Standards `Response` constructor throws `Invalid response status code 204` if you pass any body. Return `{ status: 204, headers }` with no `body` field. Don't even send `"{}"`.

### 4. Linux Consumption SCM 503 at creation

Brand-new Linux Consumption Function Apps sometimes get stuck returning 503 from SCM (Kudu) for 20+ minutes. If that happens, delete and recreate the Function App in a different region — don't wait. We hit this in `southeastasia` and recreated in `eastasia`.

### 5. Bot SSO resource URI format

For a tab hosted on a separate domain (e.g. GitHub Pages), the Entra app's identifier URI and the manifest's `webApplicationInfo.resource` must be `api://<your-tab-domain>/<botAppId>` — NOT `api://botid-<botAppId>` (that format is for bots-only personal apps where the bot owns the SSO). Teams throws "App resource defined in manifest and iframe origin do not match" if you get this wrong.

### 6. Empty `requiredResourceAccess` breaks admin consent

If the Entra app reg has no requested permissions, `/adminconsent` returns `AADSTS1003031`. Add at least Microsoft Graph `User.Read` (delegated) even if the app doesn't actually need it. We do; you should leave it.

### 7. Splitting Graph PATCHes for app reg

Setting `identifierUris` + `oauth2PermissionScopes` + `preAuthorizedApplications` in one PATCH fails because Graph validates `preAuthorizedApplications.delegatedPermissionIds` against the OLD scope list before the new one is committed. Do it as two PATCHes: scopes first, then preAuth.

## Deployment

### Tab (HTML/JS/CSS)
```powershell
git push  # GitHub Pages auto-rebuilds in ~1 min
```
Always bump `?v=` in `index.html` for tab changes.

### Bot (Azure Function App)
```powershell
cd bot
func azure functionapp publish func-day-reminders-17023 --javascript --no-build
```
The `--no-build` flag skips remote build (we ship `node_modules` in the zip — saves time on Linux Consumption first-deploys). If the Azure CLI auth has lapsed, you'll see "Unable to connect to Azure" — run `az login --tenant 705b9777-fb96-49cb-b57a-9a8fe00addad` first.

### Sideload package (manifest + icons)
```powershell
.\build.ps1
```
Outputs `dist\team-reminder.zip`. URLs are baked into `manifest.json` directly now; the optional `-BaseUrl` arg rewrites them for local-host previews.

### Outlook add-in (via unified manifest — v1.6+)
The Outlook surface is bundled into the same `manifest.json` (unified Microsoft 365 schema, v1.19) as the Teams tab and bot. **One zip uploaded to Teams admin center deploys both Teams and Outlook surfaces** to every user assigned to the Day Reminders app. No separate admin-center workflow.

Why the unified manifest: Microsoft retired the Exchange `New-App -OrganizationApp` PowerShell path (Centralized Deployment), and the M365 admin center "Integrated apps → Upload custom apps → Office Add-in" flow has a locked-to-Teams dropdown in some tenants. The unified manifest is Microsoft's current GA path that avoids both.

The `outlook.html` / `outlook.js` / `outlook.css` files are tab-side — they ship on every `git push` (GitHub Pages). No re-upload of the zip is required for taskpane code changes; bump `?v=` in `outlook.html` instead.

`outlook-manifest.xml` is kept around as a per-user OWA sideload fallback (Outlook → Get Add-ins → My add-ins → Add from file). Use only if the unified manifest can't be deployed for some reason.

### When to bump what
- Tab-only change → bump `?v=` in `index.html`, push. No re-upload needed.
- Outlook taskpane change (`outlook.*`) → bump `?v=` in `outlook.html`, push. No re-upload of the zip needed.
- Bot-only change → publish to Function App. No re-upload needed.
- `manifest.json` change (any surface) → bump `version`, run `build.ps1`, re-upload `dist/team-reminder.zip` via Teams admin center (Manage apps → Day Reminders → Update).

## Auth model

| Caller | Mechanism | Validated by |
|---|---|---|
| Tab → /api/reminders, /api/settings | Teams SSO Bearer token (audience = `api://projectvelox.github.io/<botAppId>`) | `bot/src/lib/auth.js` against Entra JWKS |
| Outlook taskpane → /api/reminders | Office SSO Bearer token via `OfficeRuntime.auth.getAccessToken()` (same audience) | same `verifyTeamsToken` — token shape is identical |
| Bot Service → /api/messages | Bot Framework JWT | `CloudAdapter` (auto) |
| Future: Workato → /api/reminders | API key + `X-User-Oid` header | TBD; not yet implemented |

`access_as_user` scope is pre-authorized for the standard Teams + Office client app IDs (seven total as of v1.6, including the Office umbrella `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`). Tenant-wide admin consent is granted (one-time, by user as Kation admin); adding the Office umbrella ID did **not** require a new consent prompt since no new scope was added.

## User collaboration notes

- **Confirm before destructive ops, but otherwise auto-execute.** Don't ask permission for normal `az` / `gh` / `git push` flows — just do them and surface what changed.
- **Check before changing.** When the user says "something's wrong," verify the actual state (table contents, App Insights logs, live HTML) before patching code. They explicitly asked this after we shipped a fix without confirming the cause first.
- **For multi-target asks, only ship what's explicitly agreed.** Don't bundle adjacent items "while you're at it."
- **No em-dashes in user-facing copy** (bot replies, error banners, Adaptive Card text). Use commas, periods, parens.
- **When shipping a release, provide a test script** the user can hand to their team. Steps with concrete actions, not just "test the new features."
- **Triage feedback before coding.** When the user pastes new requests, respond with a verdict table (Ship / Defer / Skip) and effort estimates; wait for go-ahead.

## Teams app metadata

- Manifest version is in `manifest.json` (`version` field, separate from tab `?v=` cache-bust).
- Manifest schema: `MicrosoftTeams.schema.json` v1.19 (unified Microsoft 365 app manifest, covers Teams + Outlook surfaces in one file).
- Bot App ID: `9f3711c1-c861-4da5-9664-6903bbe5bf05` (display name `Day Reminders Bot`).
- Teams app ID (manifest `id`): `5a03bfa3-63c4-417c-b668-b02234ebc11b` — same id used by the unified manifest for both Teams and Outlook surfaces.
- Outlook add-in id in legacy `outlook-manifest.xml` (fallback only): `29d36e10-9bd4-4c78-b66b-f6189ce9d7b5`.
- Tenant: Kation Technologies (`705b9777-fb96-49cb-b57a-9a8fe00addad`).
