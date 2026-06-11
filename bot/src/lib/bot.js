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
} = require('botbuilder');

const store = require('./store');

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

// ---------- snooze time helpers ----------
// Asia/Manila is UTC+8 with no DST, so we can do wall-clock math by shifting Date.
const PH_OFFSET_MS = 8 * 60 * 60 * 1000;

function snoozeMinutesIso(n) {
  return new Date(Date.now() + n * 60 * 1000).toISOString();
}
function snoozeTomorrowIso(hhmm) {
  // tomorrow's PH wall-clock at the reminder's original time (or 09:00 default),
  // converted back to a real UTC ISO.
  const ph = new Date(Date.now() + PH_OFFSET_MS);
  ph.setUTCDate(ph.getUTCDate() + 1);
  const [h, m] = (hhmm || '09:00').split(':').map(Number);
  ph.setUTCHours(h, m, 0, 0);
  return new Date(ph.getTime() - PH_OFFSET_MS).toISOString();
}
function snoozeLabel(data, untilIso) {
  if (data.tomorrow) {
    // describe in PH wall-clock terms
    const ph = new Date(new Date(untilIso).getTime() + PH_OFFSET_MS);
    const hh = String(ph.getUTCHours()).padStart(2, '0');
    const mm = String(ph.getUTCMinutes()).padStart(2, '0');
    return `tomorrow at ${hh}:${mm}`;
  }
  if (typeof data.minutes === 'number') {
    if (data.minutes < 60) return `in ${data.minutes} min`;
    if (data.minutes % 60 === 0) return `in ${data.minutes / 60} h`;
    return `in ${data.minutes} min`;
  }
  return 'later';
}

function parseAddCommand(rest) {
  const tokens = rest.trim().split(/\s+/).filter(Boolean);
  if (tokens.length === 0) return null;
  let i = 0;
  let time = null;
  const maybeTime = parseTimeToken(tokens[0]);
  if (maybeTime) { time = maybeTime; i = 1; }
  const tags = [];
  const titleParts = [];
  for (; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.startsWith('#') && t.length > 1) tags.push(t.slice(1));
    else titleParts.push(t);
  }
  const title = titleParts.join(' ').trim();
  if (!title) return null;
  return { title, time, tags: tags.slice(0, 8) };
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
          "Hi! I'm Day Reminders. Open the Reminders tab to add what you need to get done today — I'll ping you here before each one, and at your end-of-day time I'll check in to see if you're done."
        );
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
        await context.sendActivity("Try **/add 5pm Send report**, **/list**, **/done report**, or **/help**. Or open the Reminders tab.");
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
    return await super.onInvokeActivity(context);
  }

  async _handleCardAction(context, data) {
    const oid = context.activity.from.aadObjectId;
    if (!oid) return;

    if (data.action === 'markDone' && data.reminderId) {
      const r = await store.getReminder(oid, data.reminderId);
      if (r) {
        r.done = true;
        r.snoozedUntil = null;
        await store.upsertReminder(oid, r);
        await context.sendActivity(`Marked done: ${r.title}`);
      }
    } else if (data.action === 'snooze' && data.reminderId) {
      const r = await store.getReminder(oid, data.reminderId);
      if (!r) return;
      let until = null;
      if (data.tomorrow) until = snoozeTomorrowIso(r.time);
      else if (typeof data.minutes === 'number' && data.minutes > 0) until = snoozeMinutesIso(data.minutes);
      if (!until) return;
      r.snoozedUntil = until;
      r.firedAt = null; // re-enable firing on the snooze window
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
    }
  }

  async _sendHelp(context) {
    await context.sendActivity(
      "**Day Reminders commands**\n" +
      "* **/add** [time] [#tag] *title* (e.g. `/add 5pm #work Send weekly report`)\n" +
      "* **/list** to see what's open today\n" +
      "* **/done** *substring* to mark a matching item done (e.g. `/done report`)\n" +
      "* **/help** to see this again\n\n" +
      "Or use the **Reminders** tab in this app for clicking instead of typing."
    );
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
    const parsed = parseAddCommand(rest);
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
    };
    await store.upsertReminder(oid, reminder);
    const whenLabel = reminder.time ? ` at ${reminder.time}` : ' (anytime)';
    const tagLabel = reminder.tags.length ? ` ${reminder.tags.map((t) => `#${t}`).join(' ')}` : '';
    await context.sendActivity(`Added: ${reminder.title}${whenLabel}${tagLabel}`);
  }

  async _markDoneByQuery(context, oid, query) {
    if (!query) {
      await context.sendActivity("Tell me which one to close. Try `/done report`.");
      return;
    }
    const q = query.toLowerCase();
    const items = (await store.listReminders(oid)).filter((r) => !r.done);
    const matches = items.filter((r) => r.title.toLowerCase().includes(q));
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
