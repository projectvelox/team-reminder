# Day Reminders — Setup

A simple personal-tab Teams app. You jot down what you need to remember today, optionally with a time, and you get:

- a desktop notification a configurable number of minutes before each timed item
- an end-of-day "Are you done?" check-in at your chosen wrap-up time, listing anything still open

All data lives in the browser's `localStorage` (per user, per Teams client). No server, no database, no Azure resources.

---

## What's in the box

| File             | What it is                                                 |
|------------------|------------------------------------------------------------|
| `index.html`     | The tab UI                                                 |
| `app.js`         | Reminder logic, notifications, end-of-day check-in         |
| `styles.css`     | Teams-friendly light/dark/contrast theming                 |
| `color.png`      | 192x192 app icon                                           |
| `outline.png`    | 32x32 outline app icon                                     |
| `manifest.json`  | Teams app manifest (schema v1.17)                          |
| `build.ps1`      | Packages everything into `dist\team-reminder.zip`          |

---

## Step 1 — Host the static files (HTTPS required)

Teams personal tabs load over HTTPS. Pick whichever option you already have:

### Option A: GitHub Pages (free, easiest)

```powershell
git add .
git commit -m "Initial Day Reminders app"
git branch -M main
git remote add origin https://github.com/<you>/team-reminder.git
git push -u origin main
```

Then in the repo on GitHub: **Settings → Pages → Build from branch → `main` / `/ (root)` → Save**.
After a minute you'll have something like `https://<you>.github.io/team-reminder/`.

### Option B: Azure Static Web Apps, Vercel, Netlify, etc.

Any static host works. You just need the eventual public URL for the next step.

---

## Step 2 — Build the sideload `.zip`

From this folder, run:

```powershell
.\build.ps1 -BaseUrl 'https://<you>.github.io/team-reminder'
```

That rewrites `contentUrl` and `validDomains` in the manifest, then zips everything to `dist\team-reminder.zip`.

(If you'd rather hand-edit `manifest.json` once and stop passing `-BaseUrl`, just replace both `REPLACE-ME.example.com` strings and run `.\build.ps1` with no args.)

---

## Step 3 — Upload to Teams as an org app

You need a Teams admin (or the right delegated role) for org-wide upload. The path is:

1. Open **Microsoft Teams admin center** → https://admin.teams.microsoft.com
2. **Teams apps → Manage apps**
3. **Actions → Upload new app** (top right) → pick `dist\team-reminder.zip`
4. Once it lands in the list, open it and confirm it's set to **Allowed** for the org (or the policy/group you want).
5. Users will see it under **Apps → Built for your org** in Teams. They click **Add** to pin it as a personal tab.

> **Testing first without admin?** In Teams: **Apps → Manage your apps → Upload an app → Upload a custom app** (only works if your tenant has custom app uploads enabled for your account). Same `.zip`.

---

## Step 4 — Use it

- Type a reminder, optionally pick a time, hit **Add**.
- Click **Settings** (top right) to change:
  - End-of-day check-in time (default `17:00`)
  - Minutes of lead time before each timed reminder (default `10`)
  - Weekdays-only mode (skips Sat/Sun for the EOD check-in)
  - Desktop notifications on/off
  - Soft chime on/off
- First time you save Settings with notifications enabled, your browser/Teams will ask for permission. Allow it.

---

## Limits to know

- Notifications fire **while the tab is open in Teams.** If Teams is fully closed when your end-of-day time hits, you won't be pinged. For real push-when-closed reminders you'd need to add an Azure Bot Service registration (proactive messaging) — out of scope for v1, since it adds infra/cost.
- Storage is **per Teams client.** Open the tab on your work laptop and your phone, and each has its own list. (Cross-device sync would mean a backend.)
- Updating the app: change `version` in `manifest.json` (e.g. `1.0.0` → `1.0.1`), rebuild, and upload the new `.zip` from the same admin-center entry (use **Update** on the existing app rather than uploading a second copy with a new `id`).

---

## Troubleshooting

- **"App package validation failed"** during upload: usually a missing icon or a `validDomains` entry that doesn't cover your `contentUrl` host. Re-run `build.ps1 -BaseUrl ...` so they stay in sync.
- **Tab shows a blank page**: the static files aren't being served from the URL in `contentUrl`. Open that URL directly in a browser — if it 404s there, fix the hosting before anything else.
- **No notifications ever fire**: in Teams desktop, check **Settings → Notifications & activity → permissions** and confirm the tab is allowed to show notifications. In a browser, check the lock icon in the address bar.
