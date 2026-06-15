# Claude-for-Chrome handoff prompt

This is a ready-to-paste prompt for **Claude for Chrome** (the browser extension that lets a Claude session drive your Chrome tab). It captures the 8 screenshots listed in `docs/SCREENSHOTS.md` from your live Day Reminders app.

**Before pasting:**

1. Open Microsoft Teams in Chrome (`https://teams.microsoft.com`) and sign in to your Kation account.
2. Open the Day Reminders app — make sure it's loaded, signed-in, and has a few sample reminders visible (ideally a mix of clients, tags, priorities, and dates).
3. Have your Day Reminders **Chat** tab also opened in a tab so the bot/card screenshots are reachable.
4. Optional but helpful: set a real reminder for ~2 minutes from now so Screenshot 06 captures an authentic Adaptive Card. You can do this before or after starting the agent — either way, snap that one when it fires.
5. Decide where screenshots should land. The prompt below assumes `c:\Users\MSI 2\Documents\GitHub\team-reminder\docs\screenshots\` — create that folder first.
6. Start Claude for Chrome on the Teams tab, then paste the prompt below as a single message.

---

## The prompt (paste this verbatim into Claude for Chrome)

```
You are helping me capture 8 screenshots of my "Day Reminders" Microsoft Teams app for an end-user release doc. Save each screenshot as a PNG to "c:\Users\MSI 2\Documents\GitHub\team-reminder\docs\screenshots\" with the exact filenames listed below. Use Chrome's full-page or visible-region screenshot — whichever frames the scene cleanly. Crop tight (no extra browser chrome). Capture in the order listed; some require navigating between the tab and the bot chat.

For every screenshot, before clicking the camera: verify the relevant UI is fully visible, no loading spinners are showing, and no transient toasts are on screen. If the page is mid-render, wait 1 second and re-verify.

Pause and ask me before doing anything destructive (adding fake reminders, changing settings, sending messages). All my real data should stay intact — you should be capturing, not editing.

---

Screenshot 01 — Filename: 01-left-rail.png
Scene: The Teams left rail with the Day Reminders alarm-clock icon visible (selected if possible), and the two top tabs (Chat and Reminders) showing.
How: Make sure Day Reminders is the active app in the left rail. Frame so the rail + the top tabs of the app are both visible.

---

Screenshot 02 — Filename: 02-add-form.png
Scene: Top of the Reminders tab showing the add form with ALL fields visible and populated for demonstration: title field with some sample text, client field with a sample value, date picker showing today, time picker showing a sample time, and the "+ Details" textarea expanded with a couple lines of placeholder text.
How: Click into each input and add demo text. After capturing, CLEAR the inputs (do not click Add — we don't want a fake reminder saved). Frame so the entire add card is visible.

---

Screenshot 03 — Filename: 03-lines-view.png
Scene: Lines view of the Reminders tab showing 5–10 reminders with visible variety: at least one with a [Client] prefix, at least one with #tags, ideally one with a high-priority star (★), and ideally one with the overdue Nd badge if your data has one.
How: Make sure "Lines" is selected in the top-right view toggle, and "Group: off" so the flat list is showing. Frame to show the section header (Today/...) and the row list.

---

Screenshot 04 — Filename: 04-week-view.png
Scene: Week view (top-right toggle → Week) showing the Mon–Sun grid with today's column highlighted and at least 2 days with reminders stacked under date headers. Both the prev/today/next nav AND the Day ⇄ Week mini-switcher should be visible.
How: Click Week in the top-right toggle. Frame the entire grid plus the nav row.

---

Screenshot 05 — Filename: 05-group-by-client.png
Scene: Lines view with the Group toggle set to "Group: client". Show 2-3 client sections with colored headers and at least one reminder in each section.
How: Switch back to Lines view. Click the Group button until its label says "Group: client". Frame to show several client sections.

After capturing this, click the Group button again until it reads "Group: off" so you leave the UI as you found it.

---

Screenshot 06 — Filename: 06-proactive-card.png
Scene: A real Adaptive Card that fired in your Day Reminders chat, showing the title (ideally with [Client] prefix), the time subtitle, the description line if present, and the four buttons (Mark done · Snooze 15m · Snooze 1h · Tomorrow).
How: Navigate to the Day Reminders Chat tab. Scroll to a recent reminder card. If none is visible, wait for the user to fire one manually (or skip this screenshot and tell the user to capture it separately). Frame just the card, not the surrounding chat.

---

Screenshot 07 — Filename: 07-bot-help.png
Scene: The Day Reminders Chat tab showing the bot's "/help" output (the markdown list of commands).
How: In the Chat tab message box, type "/help" and send. Wait for the reply, then frame the bot's response message.

---

Screenshot 08 — Filename: 08-compose-extension.png
Scene: The Teams compose-extension popup for "Quick add reminder" — either the "..." menu opened under a message box with the Day Reminders option visible, OR the task module popup that appears after clicking it.
How: In any chat (your Day Reminders chat is fine), click the "..." button under the message box. Either capture the menu showing "Day Reminders → Quick add reminder", OR click through and capture the popup.

Important: do NOT submit the form. Cancel after capturing so no test reminder is created.

---

When all 8 are saved, list what you captured (filename + 1-line confirmation that the scene matched). If any failed (e.g., card not visible), say so — I'll capture those manually.
```

---

## After the screenshots are captured

Once `docs/screenshots/01-left-rail.png` through `08-compose-extension.png` exist:

1. Tell me ("screenshots are in"). I'll update `docs/USERGUIDE.md` so the `> _Screenshot: ..._` placeholders become real `![alt](screenshots/NN-name.png)` references.
2. I'll re-run `py tools/build-overview-pptx.py` after first **swapping the placeholder PNGs** — actually a faster path: in PowerPoint, right-click each "Screenshot NN" placeholder → **Change Picture** → **From File** → pick the corresponding `docs/screenshots/NN-*.png`. PowerPoint preserves the slot size and positioning.

If you'd rather have the generator script auto-embed the real screenshots (skipping the manual Change Picture step), I can update `tools/build-overview-pptx.py` to use the real PNG when it exists and the placeholder otherwise — a 5-line tweak. Ask when you're ready.
