// Timer trigger — runs every minute.
// For each user that has a stored Teams conversation ref, fires:
//   * lead-time reminders for timed items due within the next minute
//   * an end-of-day "Are you done?" card at the user's configured time
//
// All time-of-day comparisons run in Asia/Manila wall-clock via Intl, so
// they don't depend on the process TZ env var (which silently ignores
// timezone settings on Azure Functions Linux Consumption).

const { app } = require('@azure/functions');
const { MessageFactory } = require('botbuilder');
const { adapter } = require('../lib/bot');
const { reminderCard, eodCard } = require('../lib/cards');
const store = require('../lib/store');
const { sendReminderActivity } = require('../lib/graph');

const TIME_ZONE = 'Asia/Manila';

function phWallClock(d) {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: TIME_ZONE,
    hour12: false,
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit',
    weekday: 'short',
  }).formatToParts(d);
  const get = (t) => parts.find((p) => p.type === t)?.value;
  const hour = +get('hour') % 24; // Intl can return "24" at midnight
  return {
    date: `${get('year')}-${get('month')}-${get('day')}`,
    hour,
    minute: +get('minute'),
    minutesOfDay: hour * 60 + +get('minute'),
    weekday: get('weekday'), // e.g. "Mon"
  };
}

function hhmmToMinutes(hhmm) {
  const [h, m] = hhmm.split(':').map(Number);
  return h * 60 + m;
}

// Returns true if minutesOfDay falls inside the user's configured quiet window.
// Supports wrap-around windows (e.g. 22:00–07:00).
function inQuietWindow(settings, minutesOfDay) {
  const qs = settings?.quietStart;
  const qe = settings?.quietEnd;
  if (!qs || !qe) return false;
  const start = hhmmToMinutes(qs);
  const end = hhmmToMinutes(qe);
  if (start === end) return false;
  if (start < end) return minutesOfDay >= start && minutesOfDay < end;
  return minutesOfDay >= start || minutesOfDay < end;
}

app.timer('scheduler', {
  schedule: '0 */1 * * * *', // every minute on the second 0
  handler: async (_timer, context) => {
    const appId = process.env.MicrosoftAppId;
    if (!appId) {
      context.error('[scheduler] MicrosoftAppId not set');
      return;
    }

    const now = new Date();
    const ph = phWallClock(now);
    const tasks = [];
    const userOids = [];

    for await (const user of store.iterateUsersWithConversationRefs()) {
      userOids.push(user.oid);
      tasks.push(processUser(appId, user, ph, context).catch((err) => {
        context.error(`[scheduler] failed for ${user.oid}: ${err?.message || err}`);
      }));
    }
    context.log(`[scheduler] tick PH=${ph.date} ${String(ph.hour).padStart(2, '0')}:${String(ph.minute).padStart(2, '0')} (${ph.weekday}) users=${userOids.length}`);
    await Promise.all(tasks);
  },
});

async function processUser(appId, user, ph, context) {
  const defaultLead = user.settings?.leadMinutes ?? 10;

  // 0) once per PH day, roll forward undone past-due reminders (cap at 30 days).
  // Rollover is storage-only and runs even during quiet hours.
  await rolloverPastDue(user, ph, context);

  // Quiet hours: suppress all proactive sends (lead-time, snooze, EOD).
  // Reminders that come due inside the window will fire on the first tick
  // after the window ends; the EOD check-in is per-day and will fire as
  // soon as we're past quietEnd if still before midnight.
  if (inQuietWindow(user.settings, ph.minutesOfDay)) {
    return;
  }

  const reminders = await store.listReminders(user.oid);

  // 1) reminders. Snoozed items fire at their snoozedUntil; otherwise, normal lead-time logic.
  const nowMs = Date.now();
  for (const r of reminders) {
    if (r.done) continue;

    // Snooze path: fire when now >= snoozedUntil, then clear the snooze.
    if (r.snoozedUntil) {
      if (nowMs < new Date(r.snoozedUntil).getTime()) continue;
      context.log(`[scheduler] firing snoozed "${r.title}" snoozedUntil=${r.snoozedUntil} now=${new Date(nowMs).toISOString()}`);
      try {
        await sendProactive(appId, user, MessageFactory.attachment(reminderCard(r, 0)));
        r.snoozedUntil = null;
        r.firedAt = new Date().toISOString();
        await store.upsertReminder(user.oid, r);
      } catch (err) {
        context.error(`[scheduler] sendProactive failed for snoozed "${r.title}": ${err?.message || err}`);
      }
      // Activity feed is best-effort: never fail the proactive flow if Graph errors.
      try {
        await sendReminderActivity(user.oid, r);
      } catch (err) {
        context.log(`[scheduler] activity feed skipped for "${r.title}": ${err?.message || err}`);
      }
      continue;
    }

    // Normal time-based path — only fire on the due date.
    if (!r.time || r.firedAt) continue;
    const effectiveDueDate = r.dueAt || r.createdDate || ph.date;
    if (effectiveDueDate !== ph.date) continue;
    const effectiveLead = typeof r.leadMinutes === 'number' ? r.leadMinutes : defaultLead;
    const targetMin = hhmmToMinutes(r.time);
    const fireAtMin = targetMin - effectiveLead;
    // Upper bound is targetMin + 60 (not just targetMin) so a reminder whose
    // lead window fell entirely inside quiet hours still fires up to an hour
    // after its scheduled time. firedAt is the once-per-day guarantee.
    if (ph.minutesOfDay >= fireAtMin && ph.minutesOfDay <= targetMin + 60) {
      context.log(`[scheduler] firing "${r.title}" target=${r.time} now=${ph.hour}:${String(ph.minute).padStart(2, '0')} lead=${effectiveLead}${typeof r.leadMinutes === 'number' ? ' (custom)' : ''}`);
      try {
        await sendProactive(appId, user, MessageFactory.attachment(reminderCard(r, effectiveLead)));
        r.firedAt = new Date().toISOString();
        await store.upsertReminder(user.oid, r);
      } catch (err) {
        context.error(`[scheduler] sendProactive failed for "${r.title}": ${err?.message || err}`);
      }
      try {
        await sendReminderActivity(user.oid, r);
      } catch (err) {
        context.log(`[scheduler] activity feed skipped for "${r.title}": ${err?.message || err}`);
      }
    } else {
      context.log(`[scheduler] skip "${r.title}" target=${r.time} now=${ph.hour}:${String(ph.minute).padStart(2, '0')} fireAtMin=${fireAtMin} nowMin=${ph.minutesOfDay}`);
    }
  }

  // 2) end-of-day check-in
  const isWeekend = ph.weekday === 'Sat' || ph.weekday === 'Sun';
  if (user.settings?.weekdaysOnly && isWeekend) return;

  const eodTargetMin = hhmmToMinutes(user.settings?.eodTime || '17:00');
  if (user.lastEodDate === ph.date) return;
  if (ph.minutesOfDay < eodTargetMin) return;

  const open = reminders.filter((r) => !r.done);
  await sendProactive(appId, user, MessageFactory.attachment(eodCard(open)));
  await store.upsertUser(user.oid, { lastEodDate: ph.date });
}

async function sendProactive(appId, user, activity) {
  await adapter.continueConversationAsync(appId, user.conversationRef, async (turnContext) => {
    await turnContext.sendActivity(activity);
  });
}

// Roll any undone past-due reminder forward to today, once per user per PH day.
// Items more than 30 days past their due date are left alone so a stale backlog
// doesn't pile into the active list. Each roll bumps rollDays so the UI can
// show an "overdue N days" badge.
//
// Recurring reminders auto-advance to the next occurrence on/after today instead
// of accumulating overdue debt — the rule is "always today's task," not "missed."
async function rolloverPastDue(user, ph, context) {
  if (user.lastRolloverDate === ph.date) return;
  const today = new Date(ph.date + 'T00:00:00Z');
  const cap = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);
  const reminders = await store.listReminders(user.oid);
  for (const r of reminders) {
    if (r.done) continue;
    const effectiveDueDate = r.dueAt || r.createdDate;
    if (!effectiveDueDate || effectiveDueDate >= ph.date) continue;
    if (r.repeat && r.repeat !== 'none') {
      let next = effectiveDueDate;
      let safety = 366;
      while (next < ph.date && safety-- > 0) next = advanceOccurrence(next, r.repeat);
      r.dueAt = next;
      r.firedAt = null;
      r.rollDays = 0;
      await store.upsertReminder(user.oid, r);
      context.log(`[scheduler] advanced recurring "${r.title}" (${r.repeat}) from ${effectiveDueDate} to ${next}`);
      continue;
    }
    const dueDate = new Date(effectiveDueDate + 'T00:00:00Z');
    if (dueDate < cap) continue;
    const daysOld = Math.round((today - dueDate) / (24 * 60 * 60 * 1000));
    r.dueAt = ph.date;
    r.rollDays = (r.rollDays || 0) + daysOld;
    r.firedAt = null;
    await store.upsertReminder(user.oid, r);
    context.log(`[scheduler] rolled "${r.title}" from ${effectiveDueDate} to ${ph.date} (+${daysOld} days, total ${r.rollDays})`);
  }
  await store.upsertUser(user.oid, { lastRolloverDate: ph.date });
}

function advanceOccurrence(dateStr, repeat) {
  const d = new Date(dateStr + 'T00:00:00Z');
  if (repeat === 'daily') {
    d.setUTCDate(d.getUTCDate() + 1);
  } else if (repeat === 'weekly') {
    d.setUTCDate(d.getUTCDate() + 7);
  } else if (repeat === 'weekdays') {
    do { d.setUTCDate(d.getUTCDate() + 1); }
    while (d.getUTCDay() === 0 || d.getUTCDay() === 6);
  }
  return d.toISOString().slice(0, 10);
}
