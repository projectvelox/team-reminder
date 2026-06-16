// Member registry — the Owner picker source for the Licenses tab (v1.7).
// Every authenticated tab call auto-registers the caller. Tabs fetch this list
// to populate the Owner combobox.

const { app } = require('@azure/functions');
const { verifyTeamsToken } = require('../lib/auth');
const store = require('../lib/store');

function corsHeaders() {
  // v1.7.40 — see customers.js for rationale.
  return {
    'Access-Control-Allow-Origin': 'https://projectvelox.github.io',
    'Vary': 'Origin',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Max-Age': '600',
    'Content-Type': 'application/json',
    'X-Content-Type-Options': 'nosniff',
    'Referrer-Policy': 'no-referrer',
    'Cache-Control': 'no-store',
  };
}

function json(status, body) {
  return { status, headers: corsHeaders(), body: JSON.stringify(body) };
}

async function authed(request) {
  const auth = request.headers.get('authorization');
  return await verifyTeamsToken(auth);
}

// GET /api/members          — list everyone we know about
// POST /api/members/me      — register self (idempotent, refreshes lastSeenAt)
app.http('membersCollection', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'members',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }
    // Auto-register the caller so the list grows naturally as people use the app.
    try { await store.registerMember({ oid: user.oid, displayName: user.name, upn: user.upn }); } catch {}
    const items = await store.listMembers();
    items.sort((a, b) => (a.displayName || '').localeCompare(b.displayName || ''));
    return json(200, { members: items });
  },
});

app.http('membersMe', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'members/me',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }
    const member = await store.registerMember({ oid: user.oid, displayName: user.name, upn: user.upn });
    return json(200, { member });
  },
});
