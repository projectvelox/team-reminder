// GET / PUT /api/templates — per-user reusable reminder templates.
//
// Stored as a single JSON blob under PartitionKey=oid, RowKey='_templates'.
// Each template is a partial reminder: { title, time, client, leadMinutes, description, tags }.
// dueAt and done are intentionally not stored — templates are date-less.

const { app } = require('@azure/functions');
const { verifyTeamsToken } = require('../lib/auth');
const store = require('../lib/store');

function corsHeaders() {
  // v1.7.40 — see customers.js for rationale.
  return {
    'Access-Control-Allow-Origin': 'https://projectvelox.github.io',
    'Vary': 'Origin',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
    'Access-Control-Allow-Methods': 'GET, PUT, OPTIONS',
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

function sanitizeTemplate(t) {
  if (!t || typeof t !== 'object') return null;
  const title = typeof t.title === 'string' ? t.title.trim().slice(0, 200) : '';
  if (!title) return null;
  const time = typeof t.time === 'string' && /^\d{2}:\d{2}$/.test(t.time) ? t.time : null;
  const client = typeof t.client === 'string' && t.client.trim() ? t.client.trim().slice(0, 100) : null;
  const description = typeof t.description === 'string' && t.description.trim()
    ? t.description.trim().slice(0, 2000) : null;
  const leadMinutes = typeof t.leadMinutes === 'number' && t.leadMinutes >= 0 && t.leadMinutes <= 240
    ? Math.floor(t.leadMinutes) : null;
  const tags = Array.isArray(t.tags) ? t.tags.map(String).map(s => s.trim()).filter(Boolean).slice(0, 8) : [];
  return { title, time, client, description, leadMinutes, tags };
}

app.http('templates', {
  methods: ['GET', 'PUT', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'templates',
  handler: async (request, _context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await verifyTeamsToken(request.headers.get('authorization')); }
    catch (err) { return json(err.status || 401, { error: err.message }); }

    if (request.method === 'GET') {
      const templates = await store.getTemplates(user.oid);
      return json(200, { templates });
    }

    const body = await request.json().catch(() => ({}));
    const incoming = Array.isArray(body.templates) ? body.templates : [];
    const cleaned = incoming.map(sanitizeTemplate).filter(Boolean).slice(0, 100);
    await store.setTemplates(user.oid, cleaned);
    return json(200, { templates: cleaned });
  },
});
