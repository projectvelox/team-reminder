# Day Reminders — How to use it

Hey. This is a guide for Day Reminders — the tab + bot we built so reminders stop scattering across Outlook tasks, Notion, sticky notes, and DMs to yourself. It lives in Teams. You add stuff; the bot pings you in chat when it's due.

> Current version: **v1.4.6**. Iterates fast — if something here doesn't match the app, ping me and I'll either fix the docs or the app.

---

## 1. Finding it

Open Teams. Look at the left rail — there's an alarm-clock icon labelled **Day Reminders**. Click it.

You get two tabs at the top:

- **Reminders** — where you actually do stuff. Add, edit, filter, group.
- **Chat** — your private chat with the bot. Reminder cards land here when stuff is due, and slash commands like `/add` work here.

![Teams left rail with the Day Reminders icon and the Chat / Reminders top tabs](screenshots/01-left-rail.png)

If you don't see the icon, ping me — Teams sideload propagation can be slow.

---

## 2. Adding a reminder

Three ways to do it. Use whichever fits the moment.

### a) Type it into the tab

Top of the Reminders tab there's a row of inputs. Drop your reminder text in the first one and hit **Add** (or just press Enter).

If you want to be more specific, here's what the other fields do:

- **Title** — what to remember. Throw `#tags` right in the title (`Send report #work`) and they become colored chips you can filter by later.
- **Client** — which engagement is this for? (Citadel, BII, whoever.) Autocompletes from clients you've used before. Shows up as `[Client] Title` on every row from then on.
- **Date** — defaults to today. Change it if it's for later in the week.
- **Time** — optional. Skip it for "do it sometime today" items.
- **+ Details** — click to expand. Notes, links, sub-tasks, whatever extra context you need (up to 2000 chars).

![The add form populated with a sample title, client, date, time, and an expanded Details textarea](screenshots/02-add-form.png)

### b) Tell the bot in chat

In the Day Reminders **Chat** tab, type:

```
/add 5pm tomorrow #work Send weekly report
```

Time and date can show up in either order. Dates can be:

- `today`, `tomorrow`
- a weekday name: `mon`, `tue`, `wed`, `thu`, `fri`, `sat`, `sun`
- `M/D` like `6/20`
- or a full `YYYY-MM-DD` like `2026-06-20`

Hashtags become tags. Everything else is the title.

Other commands: `/list` shows what's open today, `/done <substring>` marks the matching one done, `/help` reminds you of all this.

![Bot chat showing the /help response with the full list of Day Reminders commands](screenshots/07-bot-help.png)

### c) Quick-add from any Teams chat

This one's underrated. Anywhere in Teams — a 1:1, a channel, even your bot chat — click the **`...`** button under the message box, pick **Day Reminders → Quick add reminder**, type, submit.

Built for the moment when you're mid-conversation with someone, they say something, and you go "I need to remember that." Stash it without leaving the chat.

![The Quick add reminder compose-extension popup ready to submit](screenshots/08-compose-extension.png)

---

## 3. Organizing your day

### Clients

Set a **Client** on a reminder and it shows everywhere as `[Client] Title` (e.g. `[Citadel] Review batch 14`). Each client gets its own color, deterministically — same client name always lands on the same color across every row.

- Click any client chip to filter to just that client.
- **Shift+click** or **right-click** the chip to inline-edit. Clearing the input removes the client.

### Tags

Anything you prefix with `#` in the title becomes a tag (`#work`, `#urgent`, `#qc`). Tags get colored chips too. Click any chip to filter to that tag.

Difference vs clients: tags are free-form scribbles (you can have many per reminder), clients are structured (one per reminder, with autocomplete).

### Priority

Click the ☆ next to a row to mark it high priority. It pins to the top.

### Dates and auto-rollover

Every reminder has a due date (defaults to today). If you don't tick it done by the end of its day, it auto-rolls forward to today's list with an **overdue Nd** badge so you can see what's been sitting around.

Cap is 30 days — anything older stays where it is and won't pile up on today.

---

## 4. Viewing your list

Top-right has a four-button switcher. Press **`v`** to cycle, or click:

- **Lines** — classic row-by-row. Best for triage.
- **Grid** — compact cards. Best for a glance.
- **Day** — hour-by-hour timeline of today. Best for time-blocking.
- **Week** — Mon–Sun grid by due date. Best for "wait, am I bombarded Thursday?"

![Lines view — the default row-by-row list, today's items with [Client] prefixes and colored tag chips](screenshots/03-lines-view.png)

![Week view — Mon to Sun grid with today's column highlighted and reminders stacked under each day](screenshots/04-week-view.png)

### Group toggle

Next to the view switcher, the **Group** button cycles:

1. **Group: off** — flat list.
2. **Group: tag** — sections per tag, header colored to match.
3. **Group: client** — sections per client, header colored to match.

Shortcut: **`g`**.

### Quick filters

Above the list, a row of pills: **All / Timed / Anytime / Priority / Done**. Click one to narrow what's visible.

![Lines view with Group: client active — sections per client with colored headers](screenshots/05-group-by-client.png)

---

## 5. Editing on the fly

Almost everything on a row is clickable:

- Click a **title** — inline-edit. Enter saves, Esc cancels.
- Click a **date chip** or **time** — same idea.
- Click a **client chip** — that filters. To EDIT the client, **Shift+click** or **right-click** it.
- Click the **details chevron** under a row to expand the description, then click the text to edit.
- Click the **`⋯` menu** for the row options dialog (custom lead time + the details textarea).
- In Lines view, drag the **☰ handle** to reorder.

---

## 6. The proactive card (the whole point)

When a timed reminder is due (or your lead-time minutes before), a card lands in your Day Reminders chat:

- **Title** — `[Client] Title` if a client is set.
- **Time line** — when it's due, or "in N min."
- **Description** — if you wrote one.
- **Buttons** — **Mark done**, **Snooze 15m**, **Snooze 1h**, **Tomorrow**.

You also get a **Teams Activity Feed** notification (the bell icon) when it fires.

![Proactive Adaptive Card in the Day Reminders chat with title, time, description, and the four action buttons](screenshots/06-proactive-card.png)

At your configured **end-of-day** time, the bot also sends an "Are you done?" card listing anything still open.

---

## 7. Settings

Hit the **⚙ Settings** button (top-right of the tab) to change:

- End-of-day check-in time
- Default lead time (in minutes) before each timed reminder
- Weekdays-only (skip Sat/Sun for the EOD card)
- Notifications on/off
- Appearance — match Teams, or lock to Light / Dark / High-contrast

---

## 8. Shortcuts and tricks

| Press this | To do this |
|---|---|
| `/` | Jump to the add field |
| `f` | Jump to the search box |
| `g` | Cycle group mode (off → tag → client) |
| `v` | Cycle views (Lines → Grid → Day → Week) |
| `?` | Open the quick guide |
| `Esc` | Clear filters / cancel an inline edit |

A few other things worth knowing:

- **Select** in the top bar puts you in bulk mode — tick a bunch of rows, then bulk-done / delete / star.
- **+ Templates** has a starter set of common reminders (standup, EOD wrap, etc.) — one click to add.
- **Search** in the top bar matches title, tags, AND client. Sticks across reloads.
- **Undo delete** — a toast hangs around for 5 seconds after every delete.

---

## 9. Feedback

Bugs, requests, "this is annoying" — drop them in the team channel or DM Josh. This thing iterates fast (more than once a day on busy days).

**Coming next (v1.5):** sharing. You'll be able to assign reminders to teammates, and set per-tag default share lists so `#QC` auto-shares with the QC team without picking recipients each time.

If you have feedback to shape that, now's the time.
