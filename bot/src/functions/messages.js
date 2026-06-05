// POST /api/messages — the Bot Framework endpoint.
// Bot Service forwards every activity here. Authentication is handled by
// the CloudAdapter (validates the JWT issued by Bot Service against this bot).

const { app } = require('@azure/functions');
const { adapter, bot } = require('../lib/bot');

app.http('messages', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'messages',
  handler: async (request, context) => {
    const body = await request.json().catch(() => ({}));
    const headers = Object.fromEntries(request.headers.entries());

    // Adapt to the Express-style req/res that CloudAdapter expects.
    const captured = { status: 200, body: undefined, headers: {} };
    const res = {
      status(code) { captured.status = code; return this; },
      send(b) { captured.body = b; return this; },
      end(b) { if (b !== undefined) captured.body = b; return this; },
      setHeader(k, v) { captured.headers[k] = v; return this; },
      header(k, v) { captured.headers[k] = v; return this; },
    };

    try {
      await adapter.process({ body, headers, method: 'POST' }, res, (turnContext) => bot.run(turnContext));
    } catch (err) {
      context.error('[messages] adapter.process failed', err);
      return { status: 500, body: 'Bot adapter error' };
    }

    return {
      status: captured.status,
      headers: captured.headers,
      body: typeof captured.body === 'string' || captured.body == null
        ? captured.body
        : JSON.stringify(captured.body),
    };
  },
});
