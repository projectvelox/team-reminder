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
const { reminderCard, eodCard, licenseFollowUpCard, licenseBriefingCard } = require('../lib/cards');
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

  // 1b) License follow-up cards for warm owners.
  // Fires once every 7 days for licenses stuck in noticeSent or awaitingCustomer
  // status. Only during work hours (9 AM-6 PM PH) so we don't ping at midnight.
  if (ph.minutesOfDay >= 9 * 60 && ph.minutesOfDay < 18 * 60) {
    await processLicenseFollowUps(appId, user, context);
  }

  // 1c) Daily morning license briefing. Weekday 8 AM PH (480 minutesOfDay,
  // tolerate the window 8:00-8:10). Once per PH day per user.
  const isWeekendDay = ph.weekday === 'Sat' || ph.weekday === 'Sun';
  if (!isWeekendDay && ph.minutesOfDay >= 8 * 60 && ph.minutesOfDay < 8 * 60 + 10) {
    if (user.lastBriefingDate !== ph.date && (!user.briefingSnoozedUntil || user.briefingSnoozedUntil <= ph.date)) {
      await processLicenseBriefing(appId, user, ph, context);
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

// Find licenses where this user is the owner, status is noticeSent or
// awaitingCustomer, the status hasn't changed in >= 7 days, and we haven't
// already followed up in the last 7 days. Send a follow-up card per match
// and stamp lastFollowUpAt so we don't re-fire for another week.
async function processLicenseFollowUps(appId, user, context) {
  const SEVEN_DAYS_MS = 7 * 24 * 60 * 60 * 1000;
  const now = Date.now();
  let licenses;
  try { licenses = await store.listLicenses(); }
  catch (err) { context.error(`[scheduler] listLicenses failed: ${err?.message || err}`); return; }
  const owned = licenses.filter((l) =>
    l.ownerOid === user.oid &&
    l.state !== 'abandoned' &&
    (l.status === 'noticeSent' || l.status === 'awaitingCustomer')
  );
  for (const lic of owned) {
    const anchorIso = lic.lastFollowUpAt || lic.statusChangedAt;
    if (!anchorIso) continue;
    const anchorMs = Date.parse(anchorIso);
    if (isNaN(anchorMs)) continue;
    if (now - anchorMs < SEVEN_DAYS_MS) continue;
    const statusChangedMs = lic.statusChangedAt ? Date.parse(lic.statusChangedAt) : now;
    const daysSinceStatus = Math.floor((now - statusChangedMs) / (24 * 60 * 60 * 1000));
    try {
      await sendProactive(appId, user, MessageFactory.attachment(licenseFollowUpCard(lic, daysSinceStatus)));
      lic.lastFollowUpAt = new Date().toISOString();
      await store.upsertLicense(lic);
      context.log(`[scheduler] license follow-up sent for ${lic.id} (${lic.customer}, ${lic.licenseType}) days=${daysSinceStatus}`);
    } catch (err) {
      context.error(`[scheduler] license follow-up failed for ${lic.id}: ${err?.message || err}`);
    }
  }
}

// Compute the per-owner briefing stats and send a single Teams card if any
// numbers are non-zero. Updates user.lastBriefingDate to skip the rest of today.
async function processLicenseBriefing(appId, user, ph, context) {
  let licenses;
  try { licenses = await store.listLicenses(); }
  catch (err) { context.error(`[scheduler] listLicenses (briefing) failed: ${err?.message || err}`); return; }
  const today = ph.date;
  const SEVEN_DAYS_MS = 7 * 24 * 60 * 60 * 1000;
  const now = Date.now();
  let overdue = 0, expiringThisWeek = 0, stuckInStatus = 0, needsAction = 0;
  for (const lic of licenses) {
    if (lic.ownerOid !== user.oid) continue;
    if (lic.state === 'abandoned') continue;
    if (lic.status === 'renewed') continue;
    if (!lic.expiryDate) continue;
    const expDate = new Date(lic.expiryDate + 'T00:00:00Z');
    const todayDate = new Date(today + 'T00:00:00Z');
    const days = Math.floor((expDate - todayDate) / 86400000);
    if (days < 0) overdue++;
    else if (days <= 7) expiringThisWeek++;
    if ((lic.status === 'noticeSent' || lic.status === 'awaitingCustomer') && lic.statusChangedAt) {
      if (now - Date.parse(lic.statusChangedAt) >= SEVEN_DAYS_MS) stuckInStatus++;
    }
    if (lic.status === 'notStarted' && days <= 30 && days >= 0) needsAction++;
  }
  const total = overdue + expiringThisWeek + stuckInStatus + needsAction;
  if (total === 0) {
    // No urgent queue. Skip the briefing to avoid alert fatigue.
    await store.upsertUser(user.oid, { lastBriefingDate: today });
    return;
  }
  try {
    await sendProactive(appId, user, MessageFactory.attachment(licenseBriefingCard({ overdue, expiringThisWeek, stuckInStatus, needsAction }, user.displayName)));
    await store.upsertUser(user.oid, { lastBriefingDate: today });
    context.log(`[scheduler] briefing sent to ${user.oid} (overdue=${overdue} week=${expiringThisWeek} stuck=${stuckInStatus} action=${needsAction})`);
  } catch (err) {
    context.error(`[scheduler] briefing send failed for ${user.oid}: ${err?.message || err}`);
  }
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
