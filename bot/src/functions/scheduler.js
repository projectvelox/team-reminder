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
  const leadMinutes = user.settings?.leadMinutes ?? 10;
  const reminders = await store.listReminders(user.oid);

  // 1) lead-time reminders. Work in PH minute-of-day so process TZ is irrelevant.
  for (const r of reminders) {
    if (r.done || !r.time || r.firedAt) continue;
    const targetMin = hhmmToMinutes(r.time);
    const fireAtMin = targetMin - leadMinutes;
    if (ph.minutesOfDay >= fireAtMin && ph.minutesOfDay <= targetMin) {
      context.log(`[scheduler] firing "${r.title}" target=${r.time} now=${ph.hour}:${String(ph.minute).padStart(2, '0')} lead=${leadMinutes}`);
      try {
        await sendProactive(appId, user, MessageFactory.attachment(reminderCard(r, leadMinutes)));
        r.firedAt = new Date().toISOString();
        await store.upsertReminder(user.oid, r);
      } catch (err) {
        context.error(`[scheduler] sendProactive failed for "${r.title}": ${err?.message || err}`);
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
