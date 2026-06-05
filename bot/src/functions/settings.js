// GET / PUT /api/settings — per-user preferences.

const { app } = require('@azure/functions');
const { verifyTeamsToken } = require('../lib/auth');
const store = require('../lib/store');

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
    'Access-Control-Allow-Methods': 'GET, PUT, OPTIONS',
    'Content-Type': 'application/json',
  };
}

function json(status, body) {
  return { status, headers: corsHeaders(), body: JSON.stringify(body) };
}

app.http('settings', {
  methods: ['GET', 'PUT', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'settings',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await verifyTeamsToken(request.headers.get('authorization')); }
    catch (err) { return json(err.status || 401, { error: err.message }); }

    if (request.method === 'GET') {
      const u = await store.getUser(user.oid);
      return json(200, { settings: u.settings, hasBot: !!u.conversationRef });
    }

    const body = await request.json().catch(() => ({}));
    const incoming = body.settings || {};
    const settings = {
      eodTime: typeof incoming.eodTime === 'string' && /^\d{2}:\d{2}$/.test(incoming.eodTime) ? incoming.eodTime : store.DEFAULT_SETTINGS.eodTime,
      leadMinutes: clampInt(incoming.leadMinutes, 0, 240, store.DEFAULT_SETTINGS.leadMinutes),
      weekdaysOnly: !!incoming.weekdaysOnly,
      notifications: incoming.notifications !== false,
    };
    await store.upsertUser(user.oid, { settings });
    return json(200, { settings });
  },
});

function clampInt(v, min, max, fallback) {
  const n = parseInt(v, 10);
  if (isNaN(n)) return fallback;
  return Math.max(min, Math.min(max, n));
}
