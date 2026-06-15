// Bot Framework adapter + activity handler.
//
// On install (conversationUpdate with bot in membersAdded), capture the
// conversation reference so the scheduler can send proactive messages later.
// On Action.Submit (Adaptive Card buttons), handle markDone / eodDismiss / eodSnooze.

const {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  TeamsActivityHandler,
  TurnContext,
  MessageFactory,
} = require('botbuilder');

const store = require('./store');
const { menuCard } = require('./cards');

function parseTimeToken(token) {
  if (!token) return null;
  const lower = token.toLowerCase();
  // 12-hour with am/pm: 5pm, 5:30pm, 5p, 12:00am
  let m = lower.match(/^(\d{1,2})(?::(\d{2}))?\s*(am?|pm?)$/);
  if (m) {
    let h = parseInt(m[1], 10);
    const mm = m[2] ? parseInt(m[2], 10) : 0;
    const isPm = m[3].startsWith('p');
    if (h < 1 || h > 12 || mm < 0 || mm > 59) return null;
    if (h === 12) h = isPm ? 12 : 0;
    else if (isPm) h += 12;
    return `${String(h).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
  }
  // 24-hour HH:MM
  m = lower.match(/^(\d{1,2}):(\d{2})$/);
  if (m) {
    const h = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10);
    if (h > 23 || mm > 59) return null;
    return `${String(h).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
  }
  return null;
}

// Today's date in Asia/Manila wall-clock as YYYY-MM-DD. Used so /add and the
// compose-extension agree with the scheduler on what "today" means.
function phToday() {
  const ph = new Date(Date.now() + PH_OFFSET_MS);
  return `${ph.getUTCFullYear()}-${String(ph.getUTCMonth() + 1).padStart(2, '0')}-${String(ph.getUTCDate()).padStart(2, '0')}`;
}

// Parses a single token as a date. Returns YYYY-MM-DD or null.
// Accepts: today, tomorrow/tmrw/tom, weekday names (next occurrence including today),
// M/D or M-D (current year, next year if past), and full YYYY-MM-DD.
function parseDateToken(token, today) {
  if (!token) return null;
  const lower = token.toLowerCase();
  if (lower === 'today') return today;
  if (lower === 'tomorrow' || lower === 'tmrw' || lower === 'tom') {
    const d = new Date(today + 'T00:00:00Z');
    d.setUTCDate(d.getUTCDate() + 1);
    return d.toISOString().slice(0, 10);
  }
  const weekdays = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];
  const wdIdx = weekdays.findIndex((w) => lower === w || lower === w + 'day' || (w === 'thu' && lower === 'thur'));
  if (wdIdx >= 0) {
    const d = new Date(today + 'T00:00:00Z');
    const delta = (wdIdx - d.getUTCDay() + 7) % 7;
    d.setUTCDate(d.getUTCDate() + delta);
    return d.toISOString().slice(0, 10);
  }
  let m = lower.match(/^(\d{1,2})[\/-](\d{1,2})$/);
  if (m) {
    const month = parseInt(m[1], 10);
    const day = parseInt(m[2], 10);
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;
    const year = parseInt(today.slice(0, 4), 10);
    const cand = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    if (cand >= today) return cand;
    return `${year + 1}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }
  m = lower.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return lower;
  return null;
}

// ---------- snooze time helpers ----------
// Asia/Manila is UTC+8 with no DST, so we can do wall-clock math by shifting Date.
const PH_OFFSET_MS = 8 * 60 * 60 * 1000;

function snoozeMinutesIso(n) {
  return new Date(Date.now() + n * 60 * 1000).toISOString();
}
function snoozeTomorrowIso(hhmm) {
  // tomorrow's PH wall-clock at the reminder's original time (or 09:00 default),
  // converted back to a real UTC ISO.
  return snoozeDaysIso(hhmm, 1);
}
function snoozeDaysIso(hhmm, days) {
  const ph = new Date(Date.now() + PH_OFFSET_MS);
  ph.setUTCDate(ph.getUTCDate() + days);
  const [h, m] = (hhmm || '09:00').split(':').map(Number);
  ph.setUTCHours(h, m, 0, 0);
  return new Date(ph.getTime() - PH_OFFSET_MS).toISOString();
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

function snoozeNextMondayIso(hhmm) {
  // days until next Monday in PH (if today IS Monday, skip to next week).
  const ph = new Date(Date.now() + PH_OFFSET_MS);
  const todayDow = ph.getUTCDay(); // 0=Sun..6=Sat
  const daysUntilMon = ((1 - todayDow + 7) % 7) || 7;
  return snoozeDaysIso(hhmm, daysUntilMon);
}
function snoozeLabel(data, untilIso) {
  const phTime = () => {
    const ph = new Date(new Date(untilIso).getTime() + PH_OFFSET_MS);
    const hh = String(ph.getUTCHours()).padStart(2, '0');
    const mm = String(ph.getUTCMinutes()).padStart(2, '0');
    return `${hh}:${mm}`;
  };
  if (data.tomorrow) return `tomorrow at ${phTime()}`;
  if (data.nextMonday) return `next Monday at ${phTime()}`;
  if (typeof data.days === 'number' && data.days > 0) {
    return `in ${data.days} day${data.days === 1 ? '' : 's'} at ${phTime()}`;
  }
  if (typeof data.minutes === 'number') {
    if (data.minutes < 60) return `in ${data.minutes} min`;
    if (data.minutes % 60 === 0) return `in ${data.minutes / 60} h`;
    return `in ${data.minutes} min`;
  }
  return 'later';
}

function taskMessage(text) {
  return {
    task: {
      type: 'message',
      value: text,
    },
  };
}

function parseAddCommand(rest, today) {
  const tokens = rest.trim().split(/\s+/).filter(Boolean);
  if (tokens.length === 0) return null;
  let i = 0;
  let time = null;
  let dueAt = null;
  // First 2 tokens can be time and/or date in either order; rest is title + tags.
  for (let n = 0; n < 2 && i < tokens.length; n++) {
    if (!time) {
      const t = parseTimeToken(tokens[i]);
      if (t) { time = t; i++; continue; }
    }
    if (!dueAt && today) {
      const d = parseDateToken(tokens[i], today);
      if (d) { dueAt = d; i++; continue; }
    }
    break;
  }
  const tags = [];
  const titleParts = [];
  for (; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.startsWith('#') && t.length > 1) tags.push(t.slice(1));
    else titleParts.push(t);
  }
  const title = titleParts.join(' ').trim();
  if (!title) return null;
  return { title, time, dueAt, tags: tags.slice(0, 8) };
}

const auth = new ConfigurationBotFrameworkAuthentication({
  MicrosoftAppId: process.env.MicrosoftAppId,
  MicrosoftAppPassword: process.env.MicrosoftAppPassword,
  MicrosoftAppType: process.env.MicrosoftAppType || 'SingleTenant',
  MicrosoftAppTenantId: process.env.MicrosoftAppTenantId,
});

const adapter = new CloudAdapter(auth);

adapter.onTurnError = async (context, error) => {
  console.error('[onTurnError]', error);
  try {
    await context.sendActivity('Sorry — the reminder bot hit an unexpected error.');
  } catch (_) { /* ignore */ }
};

class ReminderBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.onConversationUpdate(async (context, next) => {
      const activity = context.activity;
      const added = activity.membersAdded || [];
      const botAdded = added.some((m) => m.id === activity.recipient?.id);
      if (botAdded) {
        await this._registerUser(context);
        await context.sendActivity(
          "Hi! I'm Day Reminders. Open the **Reminders** tab to add what you need to get done today, or tap a button below — I'll ping you here before each one, and at your end-of-day time I'll check in."
        );
        await context.sendActivity(MessageFactory.attachment(menuCard("Quick actions:")));
      }
      await next();
    });

    this.onMessage(async (context, next) => {
      await this._registerUser(context);
      const raw = (context.activity.text || '').replace(/^<at>[^<]*<\/at>\s*/, '').trim();
      const lower = raw.toLowerCase();
      const oid = context.activity.from.aadObjectId;

      if (lower === '/help' || lower === 'help' || lower === '?') {
        await this._sendHelp(context);
      } else if (lower === '/list' || lower === 'list') {
        await this._listOpen(context, oid);
      } else if (lower.startsWith('/done ') || lower.startsWith('done ')) {
        const query = raw.slice(raw.toLowerCase().indexOf(' ') + 1).trim();
        await this._markDoneByQuery(context, oid, query);
      } else if (lower.startsWith('/add ') || lower.startsWith('add ')) {
        const rest = raw.slice(raw.toLowerCase().indexOf(' ') + 1).trim();
        await this._addFromCommand(context, oid, rest);
      } else if (raw) {
        // Free-text fallback: show the action card so non-technical users
        // don't need to know slash commands exist.
        await context.sendActivity(MessageFactory.attachment(
          menuCard("I didn't catch a command in that. Pick something below, or open the **Reminders** tab.")
        ));
      }
      await next();
    });
  }

  async onInvokeActivity(context) {
    // Adaptive Card Action.Submit comes through as invoke "adaptiveCard/action" in Teams.
    if (context.activity.name === 'adaptiveCard/action') {
      const data = context.activity.value?.action?.data || {};
      await this._handleCardAction(context, data);
      return { status: 200, body: { statusCode: 200, type: 'application/vnd.microsoft.activity.message', value: 'OK' } };
    }
    // Compose extension command from anywhere in Teams (search bar / compose box).
    if (context.activity.name === 'composeExtension/submitAction') {
      return await this._handleComposeSubmit(context);
    }
    return await super.onInvokeActivity(context);
  }

  async _handleCardAction(context, data) {
    const oid = context.activity.from.aadObjectId;
    if (!oid) return;

    if (data.action === 'markDone' && data.reminderId) {
      const r = await store.getReminder(oid, data.reminderId);
      if (r) {
        if (r.repeat && r.repeat !== 'none') {
          // Recurring: advance to next occurrence, stay open.
          r.dueAt = advanceOccurrence(r.dueAt || phToday(), r.repeat);
          r.firedAt = null;
          r.snoozedUntil = null;
          r.rollDays = 0;
          await store.upsertReminder(oid, r);
          await context.sendActivity(`Done for today: ${r.title}. Next occurrence: ${r.dueAt}.`);
        } else {
          r.done = true;
          r.closedAt = new Date().toISOString();
          r.snoozedUntil = null;
          await store.upsertReminder(oid, r);
          await context.sendActivity(`Marked done: ${r.title}`);
        }
      }
    } else if (data.action === 'snooze' && data.reminderId) {
      const r = await store.getReminder(oid, data.reminderId);
      if (!r) return;
      let until = null;
      if (data.tomorrow) until = snoozeTomorrowIso(r.time);
      else if (data.nextMonday) until = snoozeNextMondayIso(r.time);
      else if (typeof data.days === 'number' && data.days > 0) until = snoozeDaysIso(r.time, data.days);
      else if (typeof data.minutes === 'number' && data.minutes > 0) until = snoozeMinutesIso(data.minutes);
      if (!until) return;
      r.snoozedUntil = until;
      r.firedAt = null; // re-enable firing on the snooze window
      // If the snooze crosses midnight (Tomorrow, or minutes pushing past 24:00 PH),
      // advance dueAt so the next-day rollover doesn't double-count the snooze.
      const untilPhDate = new Date(new Date(until).getTime() + PH_OFFSET_MS).toISOString().slice(0, 10);
      if (untilPhDate > phToday()) {
        r.dueAt = untilPhDate;
        r.rollDays = 0;
      }
      await store.upsertReminder(oid, r);
      await context.sendActivity(`Snoozed "${r.title}" ${snoozeLabel(data, until)}.`);
    } else if (data.action === 'eodDismiss') {
      const user = await store.getUser(oid);
      user.lastEodDate = store.todayKey();
      await store.upsertUser(oid, user);
      await context.sendActivity('See you tomorrow.');
    } else if (data.action === 'eodSnooze') {
      const user = await store.getUser(oid);
      user.lastEodDate = null; // allow re-fire
      user.eodSnoozedUntil = Date.now() + 15 * 60 * 1000;
      await store.upsertUser(oid, user);
      await context.sendActivity("OK — I'll nudge you again in 15 min.");
    } else if (data.action === 'menuAdd') {
      await context.sendActivity(
        "Type your reminder here and I'll create it. Examples:\n" +
        "* `Send report at 5pm`\n" +
        "* `Review proposal tomorrow 10am`\n" +
        "* `Standup mon 9am`\n\n" +
        "Or open the **Reminders** tab and use the form at the top."
      );
    } else if (data.action === 'menuList') {
      await this._listOpen(context, oid);
    } else if (data.action === 'menuDone') {
      await context.sendActivity(
        "Tell me what to close. Type `done` followed by a word or two from the reminder. " +
        "For example: `done report` will close the next reminder whose title or notes contain \"report\"."
      );
    } else if (data.action === 'menuHelp') {
      await this._sendHelp(context);
    } else if (data.action === 'licenseSetStatus' && data.licenseId && data.status) {
      const lic = await store.getLicense(data.licenseId);
      if (!lic) { await context.sendActivity("That license isn't in the tracker anymore."); return; }
      const VALID = ['notStarted', 'noticeSent', 'awaitingCustomer', 'customerConfirmed', 'renewed'];
      if (!VALID.includes(data.status)) return;
      const now = new Date().toISOString();
      const prev = lic.status;
      lic.status = data.status;
      lic.statusChangedAt = now;
      lic.statusChangedByOid = oid;
      lic.statusChangedByName = context.activity.from?.name || null;
      lic.lastFollowUpAt = null;
      lic.lastEditedAt = now;
      lic.lastEditedByOid = oid;
      lic.lastEditedByName = context.activity.from?.name || null;
      lic.events = Array.isArray(lic.events) ? lic.events : [];
      lic.events.push({ at: now, byOid: oid, byName: context.activity.from?.name || null, type: 'statusChanged', detail: `${prev} -> ${data.status} (via card)` });
      if (lic.events.length > 50) lic.events = lic.events.slice(-50);
      await store.upsertLicense(lic);
      await context.sendActivity(`Updated: ${lic.customer}, ${lic.licenseType}. Status is now ${data.status}.`);
    } else if (data.action === 'licenseRenew' && data.licenseId) {
      const lic = await store.getLicense(data.licenseId);
      if (!lic) { await context.sendActivity("That license isn't in the tracker anymore."); return; }
      const years = (data.years === 1 || data.years === 2 || data.years === 3) ? data.years : 1;
      const base = lic.expiryDate && /^\d{4}-\d{2}-\d{2}$/.test(lic.expiryDate)
        ? new Date(lic.expiryDate + 'T00:00:00Z') : new Date();
      base.setUTCFullYear(base.getUTCFullYear() + years);
      const newExpiry = base.toISOString().slice(0, 10);
      const now = new Date().toISOString();
      lic.expiryDate = newExpiry;
      lic.status = 'renewed';
      lic.statusChangedAt = now;
      lic.statusChangedByOid = oid;
      lic.statusChangedByName = context.activity.from?.name || null;
      lic.lastRenewedAt = now;
      lic.lastFollowUpAt = null;
      lic.lastEscalatedDays = null;
      lic.lastEditedAt = now;
      lic.lastEditedByOid = oid;
      lic.lastEditedByName = context.activity.from?.name || null;
      lic.events = Array.isArray(lic.events) ? lic.events : [];
      lic.events.push({ at: now, byOid: oid, byName: context.activity.from?.name || null, type: 'renewed', detail: `+${years}y to ${newExpiry} (via card)` });
      if (lic.events.length > 50) lic.events = lic.events.slice(-50);
      await store.upsertLicense(lic);
      await context.sendActivity(`Renewed: ${lic.customer}, ${lic.licenseType}. New expiry ${newExpiry}.`);
    } else if (data.action === 'licenseFollowUpSnooze' && data.licenseId) {
      const lic = await store.getLicense(data.licenseId);
      if (!lic) return;
      lic.lastFollowUpAt = new Date().toISOString();
      await store.upsertLicense(lic);
      await context.sendActivity(`OK, I will check in again on this in 7 days.`);
    }
  }

  async _handleComposeSubmit(context) {
    const oid = context.activity.from?.aadObjectId;
    const commandId = context.activity.value?.commandId;
    const data = context.activity.value?.data || {};
    if (!oid || commandId !== 'remind') {
      return { status: 200, body: taskMessage('Sorry, this command is not recognized.') };
    }
    const text = String(data.text || '').trim();
    if (!text) {
      return { status: 200, body: taskMessage("I need at least a title. Try `5pm Send report #work`.") };
    }
    const today = phToday();
    const parsed = parseAddCommand(text, today);
    if (!parsed) {
      return { status: 200, body: taskMessage("I couldn't parse that. Try `5pm Send report #work`.") };
    }
    try {
      const id = (globalThis.crypto?.randomUUID?.()) || `id-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
      const reminder = {
        id,
        title: parsed.title,
        time: parsed.time,
        done: false,
        firedAt: null,
        createdDate: store.todayKey(),
        closedAt: null,
        tags: parsed.tags,
        priority: 'normal',
        dueAt: parsed.dueAt || today,
        description: null,
        rollDays: 0,
      };
      await store.upsertReminder(oid, reminder);
      const when = reminder.time ? ` at ${reminder.time}` : ' (anytime)';
      const dateLabel = reminder.dueAt !== today ? ` on ${reminder.dueAt}` : '';
      const tagLabel = reminder.tags.length ? ` ${reminder.tags.map((t) => `#${t}`).join(' ')}` : '';
      return { status: 200, body: taskMessage(`Added: ${reminder.title}${when}${dateLabel}${tagLabel}`) };
    } catch (err) {
      return { status: 200, body: taskMessage(`Could not add: ${err?.message || err}`) };
    }
  }

  async _sendHelp(context) {
    await context.sendActivity(MessageFactory.attachment(
      menuCard(
        "Most things live in the **Reminders** tab — open it any time. From here in chat you can also tap a button below, or type:\n" +
        "* a reminder like `Send report at 5pm` (I'll create it)\n" +
        "* `list` to see what's open\n" +
        "* `done report` to mark something done\n\n" +
        "When a reminder fires, the card has *Mark done* + snooze buttons (15m / 1h / Tomorrow / +3 days / Next Mon)."
      )
    ));
  }

  async _listOpen(context, oid) {
    const items = await store.listReminders(oid);
    const open = items.filter((r) => !r.done);
    if (open.length === 0) {
      await context.sendActivity('Nothing open right now.');
      return;
    }
    const lines = open
      .sort((a, b) => {
        if (a.priority !== b.priority) return a.priority === 'high' ? -1 : 1;
        return (a.time || 'zz').localeCompare(b.time || 'zz');
      })
      .map((r) => {
        const star = r.priority === 'high' ? '⭐ ' : '';
        const tags = (r.tags || []).length ? ' ' + r.tags.map((t) => `#${t}`).join(' ') : '';
        return r.time
          ? `* ${star}**${r.time}** ${r.title}${tags}`
          : `* ${star}${r.title}${tags}`;
      });
    await context.sendActivity(lines.join('\n'));
  }

  async _addFromCommand(context, oid, rest) {
    const today = phToday();
    const parsed = parseAddCommand(rest, today);
    if (!parsed) {
      await context.sendActivity("I need at least a title. Try `/add 5pm Send report`.");
      return;
    }
    const id = (globalThis.crypto?.randomUUID?.()) || `id-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
    const reminder = {
      id,
      title: parsed.title,
      time: parsed.time,
      done: false,
      firedAt: null,
      createdDate: store.todayKey(),
      closedAt: null,
      tags: parsed.tags,
      priority: 'normal',
      dueAt: parsed.dueAt || today,
      description: null,
      rollDays: 0,
    };
    await store.upsertReminder(oid, reminder);
    const whenLabel = reminder.time ? ` at ${reminder.time}` : ' (anytime)';
    const dateLabel = reminder.dueAt !== today ? ` on ${reminder.dueAt}` : '';
    const tagLabel = reminder.tags.length ? ` ${reminder.tags.map((t) => `#${t}`).join(' ')}` : '';
    await context.sendActivity(`Added: ${reminder.title}${whenLabel}${dateLabel}${tagLabel}`);
  }

  async _markDoneByQuery(context, oid, query) {
    if (!query) {
      await context.sendActivity("Tell me which one to close. Try `/done report`.");
      return;
    }
    const q = query.toLowerCase();
    const items = (await store.listReminders(oid)).filter((r) => !r.done);
    const matchField = (r) =>
      (r.title || "").toLowerCase().includes(q) ||
      (r.client || "").toLowerCase().includes(q) ||
      (r.description || "").toLowerCase().includes(q) ||
      (r.tags || []).some((t) => (t || "").toLowerCase().includes(q));
    const matches = items.filter(matchField);
    if (matches.length === 0) {
      await context.sendActivity(`No open item matches "${query}".`);
      return;
    }
    if (matches.length > 1) {
      const lines = matches.slice(0, 5).map((r) => `* ${r.title}`);
      await context.sendActivity(`Several match "${query}":\n${lines.join('\n')}\nBe more specific.`);
      return;
    }
    const r = matches[0];
    r.done = true;
    r.closedAt = new Date().toISOString();
    await store.upsertReminder(oid, r);
    await context.sendActivity(`Marked done: ${r.title}`);
  }

  async _registerUser(context) {
    const activity = context.activity;
    const oid = activity.from?.aadObjectId;
    if (!oid) return;
    const ref = TurnContext.getConversationReference(activity);
    await store.upsertUser(oid, {
      conversationRef: ref,
      tenantId: activity.conversation?.tenantId || activity.channelData?.tenant?.id || null,
      serviceUrl: activity.serviceUrl,
      displayName: activity.from?.name || null,
    });
  }
}

const bot = new ReminderBot();

module.exports = { adapter, bot };
