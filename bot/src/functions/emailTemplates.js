// Tenant-shared renewal-email templates per product line (v1.7.9).
// Each row has productLine (key), subject, body with {variable} placeholders
// the tab substitutes at mailto: build time.

const { app } = require('@azure/functions');
const { verifyTeamsToken } = require('../lib/auth');
const store = require('../lib/store');

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Content-Type': 'application/json',
  };
}

function json(status, body) {
  return { status, headers: corsHeaders(), body: JSON.stringify(body) };
}

async function authed(request) {
  const auth = request.headers.get('authorization');
  return await verifyTeamsToken(auth);
}

// GET /api/email-templates           - list all templates
// PUT /api/email-templates           - replace the entire set (idempotent bulk write)
app.http('emailTemplatesCollection', {
  methods: ['GET', 'PUT', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'email-templates',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }

    if (request.method === 'GET') {
      const items = await store.listEmailTemplates();
      return json(200, { templates: items });
    }
    // PUT — upsert each provided template (no implicit delete of missing keys)
    const body = await request.json().catch(() => ({}));
    const templates = Array.isArray(body.templates) ? body.templates : [];
    const now = new Date().toISOString();
    const saved = [];
    for (const t of templates.slice(0, 50)) {
      const productLine = String(t.productLine || '').trim().slice(0, 100) || '_default';
      const subject = String(t.subject || '').trim().slice(0, 500);
      const bodyText = String(t.body || '').trim().slice(0, 5000);
      if (!subject && !bodyText) continue;
      const tpl = {
        productLine,
        subject,
        body: bodyText,
        lastEditedAt: now,
        lastEditedByOid: user.oid,
        lastEditedByName: user.name || null,
      };
      await store.upsertEmailTemplate(tpl);
      saved.push(tpl);
    }
    return json(200, { templates: saved });
  },
});

// DELETE /api/email-templates/{productLine}
app.http('emailTemplatesItem', {
  methods: ['DELETE', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'email-templates/{productLine}',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }
    const ok = await store.deleteEmailTemplate(request.params.productLine);
    if (ok) return { status: 204, headers: corsHeaders() };
    return json(404, { error: 'not found' });
  },
});
