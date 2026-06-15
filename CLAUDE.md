# Day Reminders

A Microsoft Teams personal tab + bot for daily reminders, with proactive Adaptive Card notifications in the user's chat with the bot. Built for Kation Technologies; sideloaded as a custom org app.

## Why it exists

The user (Kation owner) wanted a lightweight personal productivity surface that lives inside Teams so it gets attention without context-switching. Notifications had to work even when the tab wasn't focused, which ruled out a static tab-only design and pushed us to a bot for proactive Adaptive Cards. Tab is the management UI; bot is the notification channel.

## Architecture (current)

```
projectvelox.github.io/team-reminder/  ──►  Teams static tab (HTML/JS/CSS)
                                       └──►  Outlook taskpane (outlook.html/.js)
                                              │  Teams SSO Bearer  /  Office SSO Bearer
                                              ▼
func-day-reminders-17023.azurewebsites.net  ──►  Azure Function App (Node 22, Linux Consumption, East Asia)
  /api/messages       ◄── Bot Service ──► Teams chat (proactive Adaptive Cards)
  /api/reminders      ── CRUD
  /api/settings       ── per-user prefs
  scheduler           ── Timer trigger every minute (in Asia/Manila wall-clock)
                                              │
                                              ▼
                                         Azure Table Storage  (table: dayreminders)
                                              PK = user oid
                                              RK = `_user`  or  `r:<reminderId>`
```

All Azure resources live in `rg-day-reminders`, except the Bot Service which is `global`. Storage and App Insights are in `southeastasia`, the Function App is in `eastasia` (Linux Consumption in `southeastasia` was stuck in 503 at creation; we recreated in `eastasia`). A keep-warm Logic App (`la-day-reminders-keepwarm`, southeastasia) pings `/api/ping` every 5 min during PH work hours so the Function App's HTTP triggers don't cold-start.

Secrets, GUIDs, and connection strings live in the Claude memory file `project_day_reminders_secrets.md` — **never in this repo**.

## Repo map

| Path | What |
|---|---|
| `index.html` / `app.js` / `styles.css` | The tab, served from GitHub Pages |
| `outlook.html` / `outlook.js` / `outlook.css` | Outlook add-in taskpane (new Outlook for Windows + OWA). Pre-fills a reminder from the open email, calls the same `/api/reminders`. |
| `manifest.json` | **Unified Microsoft 365 manifest** (v1.19) — single file covering tab + bot + composeExtension + Outlook add-in. Sideloaded via Teams admin center. |
| `outlook-manifest.xml` | Legacy Office Add-in XML manifest, kept as **per-user OWA sideload fallback only**. The unified `manifest.json` is the primary deploy path. |
| `color.png` / `outline.png` | App icons |
| `build.ps1` | Repackages `dist/team-reminder.zip` for sideload |
| `bot/package.json`, `bot/host.json` | Function App project |
| `bot/src/index.js` | Entry — registers all functions |
| `bot/src/functions/messages.js` | Bot endpoint (POST /api/messages) |
| `bot/src/functions/reminders.js` | CRUD for /api/reminders |
| `bot/src/functions/settings.js` | GET/PUT /api/settings |
| `bot/src/functions/scheduler.js` | Timer trigger — lead-time + EOD check-in |
| `bot/src/functions/ping.js` | Unauthenticated `GET /api/ping` returning `"ok"` — keep-warm target for the Logic App |
| `bot/src/lib/bot.js` | Bot adapter + `ReminderBot` activity handler + slash command parsing |
| `bot/src/lib/store.js` | Azure Tables wrappers |
| `bot/src/lib/auth.js` | Teams SSO JWT validation (jose + Entra JWKS) |
| `bot/src/lib/cards.js` | Adaptive Card templates |
| `dist/` | Build output, gitignored |

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
- Search, quick filters (All / Timed / Anytime / Priority / Done), templates, bulk select
- Done items >24h hidden behind a "Show N older" button
- Slash commands in bot chat: `/add [time] [date] [#tags] title`, `/list`, `/done <substring>`, `/help` — `/add` accepts `today`, `tomorrow`, weekday names, `M/D`, or `YYYY-MM-DD`
- Compose-extension command "Quick add reminder" — invocable from `...` menu under any Teams message box
- Proactive Adaptive Card per timed reminder with "Mark done" + snooze (15m / 1h / Tomorrow)
- Snoozing to Tomorrow advances `dueAt` so the rollover doesn't double-count
- EOD check-in card at configurable time
- Activity Feed notification when a reminder fires (Teams bell)

## Backlog

- **v1.5 — sharing** (Rochelle's feedback 2026-06-15). Per-reminder share with any tenant member; recipients can edit content + mark done; creator owns the recipient list; deletion removes for everyone. Includes per-tag default share lists in Settings (e.g. `#QC = [Benex, Tim]` so any reminder tagged `#QC` auto-shares). Implementation must include proactive Graph install for recipients who haven't opened the bot. See `project_day_reminders_v15_sharing_scope.md` in Claude memory for the locked decisions.
- **v1.7 (conditional)** — Markdown formatting toolbar on description, only if Phase A (plain text) usage shows demand. Toolbar emits markdown that Adaptive Card's TextBlock renders natively.
- **Recurring reminders** (Dei's feedback EIL_0003). `repeat: none | daily | weekdays | weekly` on each reminder; scheduler uses `lastFiredDate` instead of one-shot `firedAt`. Deferred behind v1.5.
- **Intra-day check-ins** (Dei's feedback EIL_0003). Settings supports a list of check-in times, not just one EOD. Each fires its own "how's it going?" card once per day. Deferred behind v1.5.
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
