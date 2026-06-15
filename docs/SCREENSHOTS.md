# Screenshot capture checklist

I can't access your Teams client, so the screenshots in the user guide + PowerPoint are placeholders. Capture these (1280×720 or larger is fine — the PPT will scale), drop them into `docs/screenshots/`, and update the references.

For consistency, capture all of them with the same Teams theme (Dark or Light, your call) and ideally with a populated reminder list so the views look realistic.

| # | Filename | What to capture | Used in |
|---|---|---|---|
| 1 | `01-left-rail.png` | The Teams left rail with the Day Reminders alarm-clock icon visible (and ideally selected). Show the two top tabs: **Chat** and **Reminders**. | Slide 3, USERGUIDE §1 |
| 2 | `02-add-form.png` | The top of the Reminders tab showing the add form with all fields visible: title (with some text typed), client (with text), date, time. Have the **+ Details** toggle expanded with a few lines in the textarea. | Slide 4, USERGUIDE §2a |
| 3 | `03-lines-view.png` | The Lines view with 5–10 reminders, ideally a mix of: some with `[Client]` prefix, some with `#tags`, one with a high-priority ⭐, one with an overdue badge if you have one. Today's date visible in the section header. | Slide 5 |
| 4 | `04-week-view.png` | The Week view showing today highlighted, with reminders stacked under a few different days. Show the Prev / Today / Next nav and the Day ⇄ Week mini-switcher. | Slide 6, USERGUIDE §4 |
| 5 | `05-group-by-client.png` | The Lines view with **Group: client** active. Multiple client sections visible, each with the client's colored header. | Slide 7, USERGUIDE §3 |
| 6 | `06-proactive-card.png` | A real Adaptive Card that fired in your Day Reminders chat. Include the title (with `[Client]` prefix if possible), subtitle, description line, and the four buttons (Mark done / Snooze 15m / Snooze 1h / Tomorrow). | Slide 8, USERGUIDE §6 |
| 7 | `07-bot-help.png` | The bot chat showing the `/help` output (the command list). | USERGUIDE §2b (optional) |
| 8 | `08-compose-extension.png` | The `...` menu opened under a Teams message box, with **Day Reminders → Quick add reminder** highlighted, OR the Quick add task module popup itself. | USERGUIDE §2c (optional) |

## How to update after capturing

For the USERGUIDE.md placeholders, replace the `> _Screenshot: ..._` lines with:

```markdown
![Description](screenshots/01-left-rail.png)
```

For the PowerPoint, open `docs/DayReminders-Overview.pptx`, click each placeholder image (labeled "REPLACE WITH SCREENSHOT #N"), and use **Picture Format → Change Picture**.

## Capture tips

- Use Windows **Snipping Tool** (`Win + Shift + S`) for clean rectangular grabs.
- Crop tight — no Windows chrome, no extra empty space outside the Teams window.
- For the Adaptive Card screenshot, set a real reminder for ~2 minutes out so you capture an authentic card, not a mockup.
- If your reminders contain sensitive client names, either blur them in the screenshot, or capture a workspace where you've added a few generic test reminders just for the docs.
