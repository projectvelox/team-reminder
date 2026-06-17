# Day Reminders — v1.7.51 + jose v6 smoke test

Paste-ready prompt for Claude for Chrome (or any tester). Walks the human through ~15 minutes of verification covering:

- **jose v6 runtime auth** (just-deployed bot dep bump — highest-risk change)
- **3 UI fixes in v1.7.51** (owner avatar, yellow contrast, sidebar spacing)
- **Brief regression on v1.7.50 fixes** to confirm the deploy didn't undo them

**Before pasting:**

1. Open Microsoft Teams in Chrome and sign in to your Kation account.
2. Open the Day Reminders app, both tabs (**Reminders** and **Licenses**) reachable.
3. Make sure you have **at least 5 licenses in the dataset** with a mix of expiry dates (some <14 days, some this month, some far out) and **at least 2 different owners**, ideally one of whom does NOT have a Microsoft profile photo set (so we exercise the initials-fallback path).
4. Set Teams theme to **Dark** for the yellow-contrast test. (We'll switch to Default mid-test.)
5. Paste the prompt below into Claude for Chrome.

---

## The prompt (paste verbatim into Claude for Chrome)

````
You are running an overnight smoke test for the "Day Reminders" Microsoft Teams app. Just-deployed changes:

- jose JWT-validation library bumped 5.10 → 6.2 in the auth backend (HIGH RISK — if broken, no tab can load).
- 3 UI fixes on the Licenses tab (owner avatar broken-image, yellow contrast in dark mode, sidebar spacing).
- 4 Dependabot bumps merged (jose runtime; typescript/setup-node/checkout dev-only).

Your job: walk through the test plan below in order, observe what actually happens, and at the end produce a structured PASS / FAIL report.

GROUND RULES:
- READ ONLY. Do not delete or modify the user's data. You may add a single test license to drive certain checks, but you must delete it before ending the run (Section 8).
- If a test FAILS, capture the symptom in 1-2 sentences and (where useful) take a screenshot to "c:\Users\MSI 2\Documents\GitHub\team-reminder\docs\screenshots\smoke-FAIL-<section>.png".
- Wait at least 1 second after any click before judging the result — Teams animations can race the assertion.
- If you hit a section you cannot execute (e.g. the dataset doesn't have a row matching the precondition), mark it SKIPPED with a one-line reason. Don't fabricate a result.

============================================================
SECTION 1 — Auth (jose v6) is the gate. If this fails, stop.
============================================================

1.1 Reload the Reminders tab (right-click the tab in the Teams left bar → Reload, OR Ctrl+R inside the iframe).
    PASS = data loads (your reminders appear) within 5 seconds.
    FAIL = red "Could not connect" banner, or a 401, or an indefinite spinner.

1.2 Reload the Licenses tab the same way.
    PASS = data loads (license rows + sidebar Numbers appear) within 5 seconds.
    FAIL = error banner, 401, or no data.

1.3 Open the browser DevTools Network panel (F12 → Network). Reload the Licenses tab one more time.
    PASS = the call to /api/licenses returns HTTP 200.
    FAIL = 401 or 500. Capture the response body if so.

If Section 1 fails on ANY of 1.1/1.2/1.3, STOP HERE and write a FAIL report. The jose v6 deploy needs a rollback.

============================================================
SECTION 2 — Owner photo broken-image fix (v1.7.51)
============================================================

Precondition: at least one license assigned to an owner WITHOUT a Microsoft profile photo. If every owner has a photo, this section is INFORMATIONAL ONLY (no broken-image scenario to exercise) — mark SKIPPED.

2.1 Find a row whose owner has no profile photo. Look at the round avatar to the left of the owner name in the OWNER column.
    PASS = you see colored initials (e.g. "TA" for Tim Albert) on a colored circle. NO broken-image glyph (the tiny torn-page icon Chrome/Edge paints for failed images). NO empty space inside the circle.
    FAIL = a broken-image glyph is visible, OR the circle is empty and the initials are missing.

2.2 Open the Edit dialog for that same row (click the "Edit" button on the row). Look at the OWNER chip section near the top of the form.
    PASS = same — colored initials, no broken-image glyph.
    FAIL = broken-image glyph visible.

2.3 Look at the SIDEBAR → "Breakdowns" → "OWNER" chips on the right side of the Licenses tab.
    PASS = every owner row shows a round colored avatar with initials (or photo if they have one). No broken-image glyphs.
    FAIL = any broken-image glyph visible.

============================================================
SECTION 3 — Yellow contrast in dark mode (v1.7.51)
============================================================

Precondition: Teams theme is DARK. If you're not in dark mode, switch now (Teams … menu → Appearance → Dark).

3.1 On the Licenses tab, find a license that expires within the next 14 days (its days-left badge will say something like "7d left" or "10d left").
    PASS = the badge background is a vivid orange and the text "Nd left" is white and clearly legible at the badge's small size. No squinting required.
    FAIL = the text is hard to read (dark text on yellow, or low-contrast).

3.2 Find a license expiring later this month but more than 14 days out (e.g. "25d left", "30d left").
    PASS = the badge background is a deep amber and the text is white and clearly legible.
    FAIL = dark/yellow-on-yellow text, hard to read at the badge size.

3.3 Set a license's STATUS to "Notice sent" (open Edit dialog → Status dropdown → Notice sent → Save). Look at the resulting STATUS pill on that row.
    PASS = the "Notice sent" pill has a brighter golden-yellow foreground and the text reads cleanly against the dark row background.
    FAIL = the text looks dim or hard to read.

3.4 Toggle Teams to LIGHT theme (Appearance → Default). Look at the same badges + pills.
    PASS = light-mode rendering is unchanged from before — still readable, no visual regression.
    FAIL = something got worse in light mode.

============================================================
SECTION 4 — Sidebar spacing (v1.7.51)
============================================================

4.1 On the Licenses tab, look at the right-hand sidebar. The top group is titled "NUMBERS" (uppercase, small) and contains:
    - 3 pill chips: "N expiring this week", "N expiring this month", "N overdue"
    - A boxed panel underneath with "N LICENSES / N SEATS / N CUSTOMERS / N RENEWED LAST 30 DAYS"

    PASS = there is clearly visible vertical breathing room (~12px) between the bottom "N overdue" chip and the top edge of the "LICENSES" box. They do NOT visually touch.
    FAIL = the chip and the box look glued together with no gap.

4.2 Click the "NUMBERS" section header to collapse the group, then click again to re-expand.
    PASS = animation is smooth, spacing returns identical after re-expand.
    FAIL = layout shifts permanently or spacing disappears.

============================================================
SECTION 5 — Regression sweep: v1.7.50 fixes still hold
============================================================

5.1 Sticky table header — scroll down the licenses table past 10+ rows.
    PASS = the column headers (CUSTOMER / LICENSE TYPE / USERS / EXPIRES / OWNER / PRODUCT LINE / STATUS / actions) stay pinned to the top of the table area while rows scroll underneath.
    FAIL = the header scrolls away with the rows.

5.2 Customer drawer backdrop — click any customer name in the table to open the customer drawer.
    PASS = the page behind the drawer is clearly dimmed (about 70% darker) and slightly blurred. The drawer stands out clearly.
    FAIL = the page behind looks barely dimmed or not at all.

5.3 Reminders footer version label — switch to the Reminders tab. Scroll to the bottom of the tab.
    PASS = footer label reads "v1.5.8" (matching the current build).
    FAIL = label still says v1.5.7 or older (cached-HTML issue).

5.4 Licenses footer version label — switch to Licenses. Click the "?" help button in the topbar to open the shortcuts overlay, then close it. Open "Settings" then close it. Find the version label (usually in the Quick Guide dialog footer).
    To check: open Quick Guide via the cmdk palette (Ctrl+K → "Open Quick guide").
    PASS = version label reads "v1.7.51".
    FAIL = label says v1.7.50 or older.

5.5 Undo toast duration — delete any reminder (single delete on Reminders tab). The toast appears with an "Undo" button.
    PASS = the toast stays visible for AT LEAST 7 seconds (count: "one mississippi, two mississippi…"). Click Undo within that window and the reminder is restored.
    FAIL = toast disappears in under 6 seconds, OR Undo button does nothing.

============================================================
SECTION 6 — Authenticated end-to-end (jose v6 deeper check)
============================================================

6.1 Add a new test license: click "+ Add license" on the Licenses tab. Customer = "Smoke Test ZZ", License type = "Test", Number of users = 1, Expires = 30 days from today, Owner = yourself, Product line = (leave blank). Status = Not started. Click Save.
    PASS = row appears in the table within 2 seconds. No error banner.
    FAIL = error banner appears, or row doesn't show up.

6.2 Edit the test row — open Edit dialog, change Notes to "smoke test note", Save.
    PASS = dialog closes, row reflects the change (Edit again to verify).
    FAIL = error banner or change doesn't persist.

6.3 Add a comment on the test row — open Edit, scroll to Comments at the top, type "smoke comment", Ctrl+Enter to send.
    PASS = comment appears in the list with your name and timestamp.
    FAIL = no comment shown, or 401, or error.

6.4 On the Reminders tab, add a quick reminder: title "smoke check", date today, no time, click Add.
    PASS = reminder appears in the list.
    FAIL = error banner.

If 6.1–6.4 all pass, jose v6 is round-tripping correctly through every CRUD path the tabs use.

============================================================
SECTION 7 — Console / network sanity
============================================================

7.1 With DevTools open (F12), look at the Console tab on both tabs.
    PASS = no red errors. Yellow warnings are OK if pre-existing.
    FAIL = any "Uncaught" red error related to the app's own code (ignore third-party Teams SDK noise).

7.2 In the Network tab, filter to "api" and confirm every request is 200 or 204.
    PASS = all green.
    FAIL = any 4xx or 5xx (note the endpoint + status).

============================================================
SECTION 8 — Cleanup (REQUIRED)
============================================================

8.1 Delete the test license ("Smoke Test ZZ") from the Licenses tab.
8.2 Delete the test reminder ("smoke check") from the Reminders tab.

============================================================
FINAL REPORT
============================================================

When done, produce a single report block in this exact format and nothing else:

# v1.7.51 + jose v6 smoke test — results

Run timestamp (PH time): YYYY-MM-DD HH:MM
Tester: (your name)
Teams theme during run: Dark → Light (or whichever order)
Dataset size: N licenses, M owners

## Section results
- Section 1 (auth): PASS / FAIL — short note
- Section 2 (owner avatar): PASS / FAIL / SKIPPED — note
- Section 3 (yellow contrast): PASS / FAIL — note
- Section 4 (sidebar spacing): PASS / FAIL — note
- Section 5 (v1.7.50 regression sweep): PASS / FAIL per sub-item
- Section 6 (end-to-end CRUD): PASS / FAIL per sub-item
- Section 7 (console/network): PASS / FAIL — note any errors verbatim
- Section 8 (cleanup): DONE / SKIPPED

## Overall verdict
SHIP / HOLD — one-sentence reason.

## Anything else you noticed
Bullet list of incidental findings, even if minor. (Things that look slightly off but weren't on the checklist.)
````

---

## Notes for the human reading this

- **If Section 1 fails:** rollback in two commits — `git revert 80dcaec && git push && cd bot && func azure functionapp publish func-day-reminders-17023 --javascript --no-build`. That puts jose v5.10 back in production.
- **If Sections 2–4 fail:** the v1.7.51 UI fix didn't take effect for that tester — almost always a Teams desktop HTML cache. Have them right-click the tab → Reload, OR clear Teams cache: `Get-Process Teams | Stop-Process; Remove-Item "$env:LOCALAPPDATA\Microsoft\Teams\Cache" -Recurse -Force`.
- **If Section 5 footer labels are wrong but everything else passes:** same cache issue; force-reload usually fixes it.
- **If Section 6 fails specifically on POST/PATCH but GET works:** that's a partial jose v6 issue — auth verifies for reads but the audience/issuer check is rejecting the same token shape on mutations. Rollback as in the first bullet.
