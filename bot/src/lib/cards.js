// Adaptive Card templates for proactive messages.

function reminderCard(reminder, leadMinutes) {
  const subtitle = reminder.time
    ? (leadMinutes > 0
        ? `In about ${leadMinutes} min (at ${formatTime(reminder.time)})`
        : `Now (${formatTime(reminder.time)})`)
    : 'Just a heads-up';

  return {
    contentType: 'application/vnd.microsoft.card.adaptive',
    content: {
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        { type: 'TextBlock', text: 'Day Reminder', weight: 'Bolder', size: 'Small', color: 'Accent' },
        { type: 'TextBlock', text: reminder.client ? `[${reminder.client}] ${reminder.title}` : reminder.title, weight: 'Bolder', size: 'Medium', wrap: true },
        { type: 'TextBlock', text: subtitle, isSubtle: true, spacing: 'None' },
        ...(reminder.description ? [{ type: 'TextBlock', text: reminder.description, wrap: true, spacing: 'Small' }] : []),
        ...(Array.isArray(reminder.subtasks) && reminder.subtasks.length > 0
          ? [{ type: 'TextBlock',
               text: reminder.subtasks.map((s) => `${s.done ? '☑' : '☐'} ${s.text}`).join('\n'),
               wrap: true,
               spacing: 'Small' }]
          : []),
      ],
      actions: [
        { type: 'Action.Submit', title: "Mark done", data: { action: 'markDone', reminderId: reminder.id } },
        { type: 'Action.Submit', title: "Snooze 15m", data: { action: 'snooze', reminderId: reminder.id, minutes: 15 } },
        { type: 'Action.Submit', title: "Snooze 1h", data: { action: 'snooze', reminderId: reminder.id, minutes: 60 } },
        { type: 'Action.Submit', title: "Tomorrow", data: { action: 'snooze', reminderId: reminder.id, tomorrow: true } },
        { type: 'Action.Submit', title: "+3 days", data: { action: 'snooze', reminderId: reminder.id, days: 3 } },
        { type: 'Action.Submit', title: "Next Mon", data: { action: 'snooze', reminderId: reminder.id, nextMonday: true } },
      ],
    },
  };
}

function eodCard(openReminders) {
  const items = openReminders.slice(0, 10).map((r) => ({
    type: 'TextBlock',
    text: r.time ? `• ${formatTime(r.time)} — ${r.title}` : `• ${r.title}`,
    wrap: true,
    spacing: 'Small',
  }));
  const overflow = openReminders.length > 10
    ? [{ type: 'TextBlock', text: `…and ${openReminders.length - 10} more`, isSubtle: true, spacing: 'Small' }]
    : [];

  const summary = openReminders.length === 0
    ? "Your list is clear. Nice work — wrap up?"
    : `You still have ${openReminders.length} item${openReminders.length === 1 ? '' : 's'} open. Want to push them to tomorrow or call it a day?`;

  return {
    contentType: 'application/vnd.microsoft.card.adaptive',
    content: {
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        { type: 'TextBlock', text: 'End of day', weight: 'Bolder', size: 'Small', color: 'Accent' },
        { type: 'TextBlock', text: summary, wrap: true, spacing: 'Small' },
        ...items,
        ...overflow,
      ],
      actions: [
        { type: 'Action.Submit', title: "I'm done for today", data: { action: 'eodDismiss' } },
        { type: 'Action.Submit', title: 'Remind me in 15 min', data: { action: 'eodSnooze' } },
      ],
    },
  };
}

function formatTime(hhmm) {
  const [h, m] = hhmm.split(':').map(Number);
  const d = new Date();
  d.setHours(h, m, 0, 0);
  return d.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });
}

// Fallback action card sent when the user @mentions the bot, sends free text, or types /help.
// Lets non-technical users tap a button instead of remembering slash commands.
function menuCard(intro) {
  return {
    contentType: 'application/vnd.microsoft.card.adaptive',
    content: {
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        { type: 'TextBlock', text: 'Day Reminders', weight: 'Bolder', size: 'Small', color: 'Accent' },
        { type: 'TextBlock', text: intro || "What can I help with?", wrap: true },
      ],
      actions: [
        { type: 'Action.Submit', title: 'Add a reminder', data: { action: 'menuAdd' } },
        { type: 'Action.Submit', title: "What's on my list?", data: { action: 'menuList' } },
        { type: 'Action.Submit', title: 'Mark something done', data: { action: 'menuDone' } },
        { type: 'Action.Submit', title: 'Show me how', data: { action: 'menuHelp' } },
      ],
    },
  };
}

module.exports = { reminderCard, eodCard, menuCard };
