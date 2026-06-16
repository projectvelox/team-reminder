# Day Reminders — full regression test plan (v1.7.45)

> Paste any single section (or the whole doc) into a Claude chat / Chrome tab to drive a manual test pass. Each step has an expected outcome.

## 0. Bootstrap

- [ ] Open Teams → Day Reminders app (sideloaded as `Day Reminders`).
- [ ] Reminders tab loads without the red error banner. Footer says `v1.4.6` (cosmetic label) but the loaded cache-bust is `v=1.5.4`. No console errors during boot.
- [ ] Switch to Licenses tab. Loads without error. Topbar shows the icon utility cluster (🗑 / ↻ / ⚙ / ?). Sync indicator shows "Updated Ns ago" and ticks up.
- [ ] Hard refresh both tabs (Ctrl+F5 in browser preview, or rebump from Teams admin if cached). Confirm both still load.

---

## 1. Reminders tab — core CRUD

- [ ] **Add** a timed reminder: Title `Test 1`, client `Acme`, date today, time 30 min from now. Press Add. Row appears in Lines view.
- [ ] **Inline edit** title (click it) → change to `Test 1 edited` → Enter saves. Esc cancels another edit.
- [ ] **Inline edit** time → change to 1 hour from now. Saves.
- [ ] **Toggle done** via checkbox. Row strikes through. Toggle again → returns.
- [ ] **Delete** via `…` menu → Delete. Toast shows "Deleted Test 1 edited · Undo" (5s window). Click Undo before timer → row returns.
- [ ] Delete again, let timer expire → row is permanently gone after a refresh.

## 2. Reminders tab — new in v1.7.44

- [ ] **Next-up hero**: with at least one timed reminder for today (in the future) or tomorrow, a blue card appears at the top of the tab showing the time, title, and countdown ("in 32 min" / "tomorrow at 9:00 AM"). Click the card → the matching row scrolls into view and flashes accent color for ~1.5s.
- [ ] **Keyboard shortcuts overlay**: press `?` outside any input. Overlay opens. Try `/` (focus search), `n` (focus new field), `s` (toggle Select), `t` (open templates), `v` (cycle views), `g` (cycle group), Esc (close).
- [ ] **Drag hint**: switch to Week view. Amber banner "Tip: drag a reminder…" appears at the top. Dismiss with ×. Switch away + back → banner does NOT reappear (once-only).
- [ ] **Drag-to-reschedule** (Week view): drag a row to a different day → dueAt updates. Drag back. Toast confirms.
- [ ] **Bulk mark-done Undo**: turn on Select, tick 2–3 rows, click Mark all done → toast "Marked N done · Undo". Click Undo within 6s → rows return to open state.
- [ ] **Bulk delete Undo**: select 2–3 rows, click Delete, confirm. Toast "Deleted N reminders · Undo". Click Undo within 6s → rows return.

## 3. Reminders tab — recurring + quiet hours + templates + tags

- [ ] **Templates**: open Templates → add `Standup` template (10:00, tag `work`). Save. From the add row, open templates again → click to apply → fields prefill.
- [ ] **Tag chip picker**: type a tag in the picker → press Enter → chip appears. Backspace removes the last chip.
- [ ] **Tag in title**: type `Send report #urgent` in the title field → press Add → row shows the `urgent` chip separately, title is `Send report`.
- [ ] **Recurring**: open row options (`…` menu) → Recurring → weekly. Save. Mark the row done → it advances to next week instead of closing.
- [ ] **Quiet hours**: Settings → set quiet hours 23:00–07:00. Add a reminder for the current PH time but inside that window (set system clock if needed, or use a one-off). Verify the proactive card does NOT fire until window ends.

## 4. Reminders tab — Lines / Grid / Day / Week parity

- [ ] In each of the 4 views, the same reminders are visible. Toggling does not lose state.
- [ ] **Day view**: shows hour-by-hour. New v1.5.2 date header at top reads "Today · Wed, Jun 17" or similar.
- [ ] **Week view**: Mon–Sun grid. Today's column highlighted. Drag works between days.
- [ ] **Rolled-over badge**: create an open reminder dated yesterday. Reload tab tomorrow → it appears under today with an amber `+1d` pill.

## 5. Licenses tab — view + filter

- [ ] **Table view** loads. Sortable columns work (click Customer header → sorts asc/desc; arrow indicator changes).
- [ ] **Search**: type a customer name in the topbar search → table debounces 150ms, filters down.
- [ ] **Quick filters**: All / Mine / This month / Overdue / Needs attention. Each toggles correctly.
- [ ] **Month dropdown**: pick a month → table narrows to that YYYY-MM. Calendar view jumps to that month if you switch.
- [ ] **Expiry popover**: click `Expiry: Any ▾` → 4 checkboxes with color dots (Expired / Expiring soon / This month / Active). Tick + untick. Active chip appears in the bar.
- [ ] **Date-range popover** (v1.7.42): click `Range: Any ▾` → set from + to dates → chip "Expires Jun 1, 2026 – Jul 15, 2026" appears in the bar.
- [ ] **Active-filters bar**: shows `12 of 47` style count + removable chips + Clear all.
- [ ] **Save view** (v1.7.39): with some filters set, click `+ Save view` → name it → chip appears above the bar. Refresh tab → chip persists. Click chip → filters re-apply. Click chip × → deletes.
- [ ] **URL hash** (v1.7.42): copy the browser URL after filtering → paste in a new browser window/incognito (with the same Teams session) → same filters apply on load.

## 6. Licenses tab — Calendar view

- [ ] Switch to Calendar view. Mon–Sun grid. Today's cell highlighted.
- [ ] **Pill colors**: each license pill is owner-color filled with a left-edge stripe matching its Expiry bucket (red / orange / yellow / green).
- [ ] **Density toggle**: ▦ vs ▪ switch in the calendar nav. Compact halves the cell size.
- [ ] **Year jump**: dropdown shows years with data + 3 future years. Selecting jumps the calendar.
- [ ] **Same month next year**: button advances cursor 12 months.
- [ ] **Click pill** → opens edit dialog. **Click empty cell** → opens add dialog pre-filled with that date.

## 7. Licenses tab — table row UX

- [ ] **Owner avatar** appears next to the owner pill in each row (round photo or initials fallback).
- [ ] **Days-left badge** and **Expiry pill** (4-bucket color) both show on every row.
- [ ] **Customer hover preview** (v1.7.44): hover any customer name for 250ms → card pops up with licenses count, seats, owners, next expiry, overdue badge.
- [ ] **Hover-notes tooltip** (v1.7.39): hover any row with notes set → browser tooltip shows the full notes (truncated 600 chars).
- [ ] **Sticky header** (v1.7.44): scroll the table → column headers stay anchored.
- [ ] **Comment badge** on Edit button: rows with comments show `💬 N` badge.

## 8. Licenses tab — Add + Edit + Comments

- [ ] **Quick-add inline** (v1.7.39): `+ Quick add a license` strip → fill 4 fields → Enter → row created.
- [ ] **Full Add dialog**: `+ Add license` → all fields. Customer + license-type autocomplete from past values.
- [ ] **Owner picker**: search-as-you-type via Graph; results show name + email + avatar.
- [ ] **Lead-days picker** (v1.7.37): chip-multi-select (60/30/15/7/1 + Custom). Toggle presets. Add a custom day. Remove a chip.
- [ ] **Save** with valid data → toast "Saved", row updates.
- [ ] **Concurrent edit conflict** (v1.7.41): open the same license in two windows (or simulate by editing `lastEditedAt` in the in-memory licenses array via console). Save in one → other gets 409 with "Save blocked: NAME changed this row first · Reload" toast.
- [ ] **Renew** button in row: click → quick +1y with Undo toast. Shift-click → opens dialog with 1y/2y/3y presets + custom date.
- [ ] **Comments** (v1.7.39): open Edit dialog. Comments section is at the TOP (v1.7.44 moved it up). Add a comment via Ctrl+Enter or click Comment. Newest at bottom.
- [ ] **Comments edit/delete** (v1.7.45): hover your own comment → Edit + Delete buttons fade in. Edit replaces body with textarea; Save / Cancel (Ctrl+Enter saves, Esc cancels). After edit, "(edited)" stamp appears. Delete → confirm → comment gone. Hover someone else's comment → no Edit/Delete (author-only).

## 9. Licenses tab — Bulk operations

- [ ] **Select mode**: click Select in topbar → checkbox column appears, bulk bar at bottom shows count.
- [ ] **Bulk Reassign owner** → picker → confirm → all selected get new owner.
- [ ] **Bulk Renew +1y** (v1.7.41): pick rows → Renew +1y → confirm → toast "N renewed, M skipped (edited by someone else)" if any have stale `lastEditedAt`.
- [ ] **Bulk Export CSV** → downloads `selected-licenses-YYYY-MM-DD.csv` with only selected rows.
- [ ] **Bulk Delete** → confirm → optimistic remove → toast "Deleted N · Undo". Undo within 6s → rows return.

## 10. Licenses tab — Customer drawer + merge

- [ ] **Click customer name** → side drawer opens with all that customer's licenses, totals, contact emails, notes.
- [ ] **Customer Edit** → modal for primary email, secondary emails, address, notes. Save.
- [ ] **Merge** (v1.7.41): from the drawer, click `Merge in duplicates…` → dialog lists customers/license-strings with ≥3-char token overlap. Tick 1–2 → Merge selected → confirm → all selected names rewritten, source customer rows deleted, activity shows `customerMerged`.

## 11. Licenses tab — Sidebar (v1.7.44 + v1.7.45)

- [ ] Sidebar shows 4 accordion groups:
  - **Numbers** (open) — summary chips + stats strip
  - **Breakdowns** (open) — status / owner / product chips
  - **Renewals (last 90d)** (collapsed by default, v1.7.45 rename) — trend chart + rate widget + leaderboard
  - **Activity** (collapsed) — last 20 events
- [ ] Expand each. Click summary chip → filter applies. Click breakdown chip → multi-select filter toggle.
- [ ] **Owner sidebar chips** show round avatar (v1.7.44) matching the table row.
- [ ] **Click leaderboard owner name** (v1.7.44) → owner filter set; sidebar Owner chip activates.
- [ ] **Click month on trend chart** (v1.7.44) → month filter set; toast "Filtered to <month>".
- [ ] **Activity entries** clickable → open the corresponding license.

## 12. Licenses tab — Recovery + Privacy (v1.7.43)

- [ ] **Delete a license** → wait for Undo to expire (or skip).
- [ ] **Trash can icon** (v1.7.45) in topbar shows red badge with count.
- [ ] Click trash icon → recovery dialog lists soft-deleted rows with "Deleted Ns ago · N days until permanent purge" + Restore button.
- [ ] **Restore** → row returns to main table; badge count drops.
- [ ] **Settings → Data & privacy tab** (v1.7.45 tabbed):
  - [ ] **Export my data** → JSON download with everything stored about caller.
  - [ ] **More details** → opens privacy dialog (What we store / Where / How long / Your controls / Third parties).
  - [ ] **Delete my account data** → double-confirm (confirm() + type-DELETE prompt) → wipes reminders/settings/templates, unassigns from licenses.

## 13. Licenses tab — Command palette + keyboard nav (v1.7.39 / v1.7.45)

- [ ] Press `Ctrl+K` / `Cmd+K` → palette opens. Type `add` → "Add license" action. Type a customer name → that entity appears. Arrow keys + Enter.
- [ ] Press `?` → shortcuts overlay.
- [ ] Other shortcuts: `/` (search), `a` (add), `q` (quick-add), `r` (refresh), `g t` (Table view), `g c` (Calendar view).
- [ ] **Table arrow-key nav** (v1.7.45): outside any input, press ↓ or `j` → focus ring on first row. Up/Down/j/k move it. **Enter** opens edit dialog. Home/End jump to first/last.

## 14. Settings dialog (v1.7.45 tabbed)

- [ ] Click ⚙ Settings → tabs at top: **General** / **Notifications** / **Data & privacy**.
- [ ] **General** tab: Theme override dropdown.
- [ ] **Notifications** tab: Default lead-times picker, briefing opt-out, monthly digest opt-out, roll-up opt-in, send-test-digest buttons.
- [ ] **Data & privacy** tab: Recovery button, Export my data, Delete my account data, More details link.

## 15. Dialogs — close X (v1.7.40)

- [ ] Every modal dialog (Edit license, Settings, Renew, Customer, Templates, CSV import, Bulk reassign, Day overflow, Guide, Cmdk, Shortcuts, Recovery, Privacy) has an `×` close button in the top-right corner. Click it → dialog closes.

## 16. CSV import / export / template

- [ ] **CSV dropdown menu**: Import / Export current view / Download template.
- [ ] **Download template** → minimal CSV with example rows.
- [ ] **Export current view** → CSV of currently-visible rows.
- [ ] **Import**: pick a CSV → preview table with per-row status (ready / needs owner / invalid / duplicate / update). Resolve owners via inline picker. Confirm → toast counts ready vs updated vs skipped.

## 17. Email templates

- [ ] Templates button in topbar → modal lists per-product-line templates (M365, BC, BREP, etc.). Edit subject + body. Save.
- [ ] On any license row, click **Email** → opens Outlook draft with placeholders substituted ({customer}, {customerFirstWord}, {licenseType}, {users}, {expiryDate}, {ownerName}).
- [ ] Bundle email: if a customer has 2+ licenses expiring within 14 days → Email button reads `Email N` → drafts one combined message.

## 18. Bot (proactive cards in Teams chat)

- [ ] **Add reminder** with a time in the next 15 min, leadMinutes=5 → wait → Adaptive Card arrives in the Day Reminders bot chat with title, subtitle, **Mark done / Snooze 15m / Snooze 1h / Tomorrow / +3 days / Next Mon**.
- [ ] **Snooze** → reminder re-fires at the chosen time.
- [ ] **Mark done** → row closes in tab.
- [ ] **EOD check-in**: configured EOD time fires once per day with list of open reminders.
- [ ] **License lead card** (v1.7.37): set a license expiry 7 days out with leadDays=[7]. Wait until 9:00 AM PH next weekday → card arrives with **Mark renewed (+1y) / Snooze 7 days / Won't renew / Open in tab**.
- [ ] **License follow-up card**: set a license to status=noticeSent. Wait 7+ days → follow-up card arrives during work hours.
- [ ] **Daily briefing card**: weekday 8 AM PH → if owner has overdue / expiring / stuck-in-status / needs-action items, briefing card arrives. Otherwise silent.

## 19. Bot slash commands + free-text

- [ ] `/help` or `help` → menu card with Add / List / Mark done / Show me how buttons.
- [ ] `/list` → bot replies with list of open reminders.
- [ ] `/done report` → marks matching reminder done.
- [ ] `/add 5pm Send report #work` → creates a reminder for 17:00 today with tag `work`.
- [ ] Free-text "hello" → bot replies with menu card (fallback).

## 20. Compose extension

- [ ] In any Teams chat, click `…` under compose box → Day Reminders → Quick add reminder. Type `5pm Send report` → submits → reminder appears in tab.

## 21. Outlook taskpane

- [ ] Open an email in new Outlook for Windows or OWA → click `Add reminder` in the message action bar (or right-click → Apps → Day Reminders).
- [ ] Taskpane opens with subject pre-filled as title, sender as client, date defaulted to tomorrow.
- [ ] Save → reminder appears in the Reminders tab.

## 22. Settings persistence

- [ ] Set theme override to Dark → reload tab → still dark.
- [ ] Change EOD time → reload → persists.
- [ ] Toggle off monthly digest → reload → still off.
- [ ] Add saved view → reload → still there.

## 23. Live poll + freshness (v1.7.39 / v1.7.40)

- [ ] Sync indicator in Licenses topbar shows "Updated Ns ago" and ticks every 10s.
- [ ] Open Licenses in two windows. Edit a license in window 1. Wait 60s in window 2 → row updates without manual refresh.
- [ ] When any dialog is open, the poll skips re-render (data refreshes after dialog closes).

## 24. Mobile / narrow viewport (v1.7.42)

- [ ] Resize browser to ≤900px wide → sidebar collapses BELOW the main table; table gets horizontal scroll.
- [ ] Resize to ≤600px → filter pills wrap with bigger touch targets; stats strip wraps.

## 25. Accessibility quick-pass

- [ ] Reduced-motion: enable OS reduced-motion → shimmer / slide-in animations suppressed.
- [ ] Tab through interactive elements with keyboard → focus rings visible on every button, input, chip, dialog control.
- [ ] Open command palette / shortcuts overlay → all options reachable via keyboard.

## 26. CI / dev infra (v1.7.43)

Not user-visible, but verify:
- [ ] `bot/npm test` → 20 tests green.
- [ ] `bot/npm run lint` → no syntax errors.
- [ ] GitHub Actions CI runs on every push → green badge.
- [ ] Dependabot has opened or scheduled weekly npm PRs.

---

## Quick smoke test (5 minutes)

If short on time, do at minimum:

1. Open both tabs — both load without error.
2. Add a reminder in Reminders → row appears.
3. Add a license in Licenses → row appears.
4. Open any license → comments at top, sticky header on the table.
5. Press `?` in either tab → shortcuts overlay opens.
6. Click trash icon (Licenses) → recovery dialog opens.
7. Settings (Licenses) → 3 tabs render.
8. Add a comment → hover it → Edit/Delete tools appear → edit → "(edited)" stamp shows.

If all 8 pass, no major regressions.
