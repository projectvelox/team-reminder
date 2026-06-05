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
      // Capture/refresh the conversation ref on any message too.
      await this._registerUser(context);

      const text = (context.activity.text || '').trim().toLowerCase();
      if (text === 'help' || text === '/help') {
        await context.sendActivity("Open the **Reminders** tab to add or check off items. I post here before each timed reminder and at your end-of-day time.");
      } else if (text === 'list' || text === '/list') {
        const items = await store.listReminders(context.activity.from.aadObjectId);
        const open = items.filter((r) => !r.done);
        if (open.length === 0) {
          await context.sendActivity('Nothing open right now.');
        } else {
          const lines = open
            .sort((a, b) => (a.time || 'zz').localeCompare(b.time || 'zz'))
            .map((r) => (r.time ? `• ${r.time} — ${r.title}` : `• ${r.title}`));
          await context.sendActivity(lines.join('\n'));
        }
      } else if (text) {
        await context.sendActivity("Type **list** to see open items, or open the **Reminders** tab to add new ones.");
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
        await store.upsertReminder(oid, r);
        await context.sendActivity(`Marked done: ${r.title}`);
      }
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
