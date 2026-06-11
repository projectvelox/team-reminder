# Day Reminders

A Microsoft Teams personal tab + bot for daily reminders, with proactive Adaptive Card notifications in the user's chat with the bot. Built for Kation Technologies; sideloaded as a custom org app.

## Why it exists

The user (Kation owner) wanted a lightweight personal productivity surface that lives inside Teams so it gets attention without context-switching. Notifications had to work even when the tab wasn't focused, which ruled out a static tab-only design and pushed us to a bot for proactive Adaptive Cards. Tab is the management UI; bot is the notification channel.

## Architecture (current)

```
projectvelox.github.io/team-reminder/  ──►  Teams static tab (HTML/JS/CSS)
                                              │  Teams SSO Bearer
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

All Azure resources live in `rg-day-reminders`, except the Bot Service which is `global`. Storage and App Insights are in `southeastasia`, the Function App is in `eastasia` (Linux Consumption in `southeastasia` was stuck in 503 at creation; we recreated in `eastasia`).

Secrets, GUIDs, and connection strings live in the Claude memory file `project_day_reminders_secrets.md` — **never in this repo**.

## Repo map

| Path | What |
|---|---|
| `index.html` / `app.js` / `styles.css` | The tab, served from GitHub Pages |
| `manifest.json` | Teams app manifest (sideload) |
| `color.png` / `outline.png` | App icons |
| `build.ps1` | Repackages `dist/team-reminder.zip` for sideload |
| `bot/package.json`, `bot/host.json` | Function App project |
| `bot/src/index.js` | Entry — registers all functions |
| `bot/src/functions/messages.js` | Bot endpoint (POST /api/messages) |
| `bot/src/functions/reminders.js` | CRUD for /api/reminders |
| `bot/src/functions/settings.js` | GET/PUT /api/settings |
| `bot/src/functions/scheduler.js` | Timer trigger — lead-time + EOD check-in |
| `bot/src/lib/bot.js` | Bot adapter + `ReminderBot` activity handler + slash command parsing |
| `bot/src/lib/store.js` | Azure Tables wrappers |
| `bot/src/lib/auth.js` | Teams SSO JWT validation (jose + Entra JWKS) |
| `bot/src/lib/cards.js` | Adaptive Card templates |
| `dist/` | Build output, gitignored |

## What ships today (v1.2.x)

- Add / delete / done-toggle / inline-edit (title + time)
- Hashtag-in-title parses to tags; colored chips
- High-priority star, pinned to top of list
- Group-by-tag view toggle (acts as project sections)
- Done items >24h hidden behind a "Show N older" button
- Slash commands in bot chat: `/add [time] [#tags] title`, `/list`, `/done <substring>`, `/help`
- Proactive Adaptive Card per timed reminder with "Mark done" button
- EOD check-in card at configurable time

## Backlog

- **v1.3 — recurring + intra-day check-ins** (Dei's feedback EIL_0003). Two sub-features:
  - **3a** `repeat: none | daily | weekdays | weekly` on each reminder; scheduler uses `lastFiredDate` instead of one-shot `firedAt`.
  - **3b** Settings supports a list of check-in times, not just one EOD. Each fires its own "how's it going?" card once per day.
- **Workato sidecar for long-message intent parsing** (Josh's feedback EIL_0005). Backend gets an API-key auth path on `/api/reminders` (in addition to Teams SSO); a Workato recipe receives `@DayReminders <long message>` in Teams, parses via Workato AI, loops `POST /api/reminders` for each extracted item. ~30 min on our side; recipe lives in Kation's Workato workspace.

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

### When to bump what
- Tab-only change → bump `?v=` in `index.html`, push. No re-upload needed.
- Bot-only change → publish to Function App. No re-upload needed.
- `manifest.json` change → bump `version`, run `build.ps1`, re-upload via Teams admin center (Manage apps → Day Reminders → Update).

## Auth model

| Caller | Mechanism | Validated by |
|---|---|---|
| Tab → /api/reminders, /api/settings | Teams SSO Bearer token (audience = `api://projectvelox.github.io/<botAppId>`) | `bot/src/lib/auth.js` against Entra JWKS |
| Bot Service → /api/messages | Bot Framework JWT | `CloudAdapter` (auto) |
| Future: Workato → /api/reminders | API key + `X-User-Oid` header | TBD; not yet implemented |

`access_as_user` scope is pre-authorized for the six standard Teams + Office client app IDs. Tenant-wide admin consent is granted (one-time, by user as Kation admin).

## User collaboration notes

- **Confirm before destructive ops, but otherwise auto-execute.** Don't ask permission for normal `az` / `gh` / `git push` flows — just do them and surface what changed.
- **Check before changing.** When the user says "something's wrong," verify the actual state (table contents, App Insights logs, live HTML) before patching code. They explicitly asked this after we shipped a fix without confirming the cause first.
- **For multi-target asks, only ship what's explicitly agreed.** Don't bundle adjacent items "while you're at it."
- **No em-dashes in user-facing copy** (bot replies, error banners, Adaptive Card text). Use commas, periods, parens.
- **When shipping a release, provide a test script** the user can hand to their team. Steps with concrete actions, not just "test the new features."
- **Triage feedback before coding.** When the user pastes new requests, respond with a verdict table (Ship / Defer / Skip) and effort estimates; wait for go-ahead.

## Teams app metadata

- Manifest version is in `manifest.json` (`version` field, separate from tab `?v=` cache-bust).
- Bot App ID: `9f3711c1-c861-4da5-9664-6903bbe5bf05` (display name `Day Reminders Bot`).
- Teams app ID (manifest `id`): `5a03bfa3-63c4-417c-b668-b02234ebc11b`.
- Tenant: Kation Technologies (`705b9777-fb96-49cb-b57a-9a8fe00addad`).
