// Timer trigger — runs every minute.
// For each user that has a stored Teams conversation ref, fires:
//   * lead-time reminders for timed items due within the next minute
//   * an end-of-day "Are you done?" card at the user's configured time

const { app } = require('@azure/functions');
const { MessageFactory } = require('botbuilder');
const { adapter } = require('../lib/bot');
const { reminderCard, eodCard } = require('../lib/cards');
const store = require('../lib/store');

const MS_PER_MIN = 60 * 1000;

app.timer('scheduler', {
  schedule: '0 */1 * * * *', // every minute on the second 0
  handler: async (_timer, context) => {
    const appId = process.env.MicrosoftAppId;
    if (!appId) {
      context.error('[scheduler] MicrosoftAppId not set');
      return;
    }

    const now = new Date();
    const tasks = [];

    for await (const user of store.iterateUsersWithConversationRefs()) {
      tasks.push(processUser(appId, user, now, context).catch((err) => {
        context.error(`[scheduler] failed for ${user.oid}`, err);
      }));
    }
    await Promise.all(tasks);
  },
});

async function processUser(appId, user, now, context) {
  // 1) lead-time reminders
  const reminders = await store.listReminders(user.oid);
  const lead = (user.settings?.leadMinutes ?? 10) * MS_PER_MIN;
  const fireWindow = MS_PER_MIN; // we tick once per minute

  for (const r of reminders) {
    if (r.done || !r.time || r.firedAt) continue;
    const targetMs = timeToTodayMs(now, r.time);
    const fireAt = targetMs - lead;
    if (now.getTime() >= fireAt && now.getTime() < targetMs + fireWindow) {
      await sendProactive(appId, user, MessageFactory.attachment(reminderCard(r, user.settings?.leadMinutes ?? 10)));
      r.firedAt = new Date().toISOString();
      await store.upsertReminder(user.oid, r);
    }
  }

  // 2) end-of-day check-in
  const todayKey = store.todayKey(now);
  const isWeekend = now.getDay() === 0 || now.getDay() === 6;
  if (user.settings?.weekdaysOnly && isWeekend) return;

  const eodMs = timeToTodayMs(now, user.settings?.eodTime || '17:00');
  const eodAlreadySentToday = user.lastEodDate === todayKey;
  if (eodAlreadySentToday) return;
  if (now.getTime() < eodMs) return;

  const open = reminders.filter((r) => !r.done);
  await sendProactive(appId, user, MessageFactory.attachment(eodCard(open)));
  await store.upsertUser(user.oid, { lastEodDate: todayKey });
}

async function sendProactive(appId, user, activity) {
  await adapter.continueConversationAsync(appId, user.conversationRef, async (turnContext) => {
    await turnContext.sendActivity(activity);
  });
}

function timeToTodayMs(now, hhmm) {
  const [h, m] = hhmm.split(':').map(Number);
  const d = new Date(now);
  d.setHours(h, m, 0, 0);
  return d.getTime();
}
