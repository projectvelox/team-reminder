// REST API for the tab to read/write reminders.
// All endpoints require a Teams SSO Bearer token; user identity = oid claim.

const { app } = require('@azure/functions');
const { verifyTeamsToken } = require('../lib/auth');
const store = require('../lib/store');

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
    'Access-Control-Allow-Methods': 'GET, POST, PATCH, DELETE, OPTIONS',
    'Content-Type': 'application/json',
  };
}

async function authed(request) {
  const auth = request.headers.get('authorization');
  return await verifyTeamsToken(auth);
}

function json(status, body) {
  return { status, headers: corsHeaders(), body: JSON.stringify(body) };
}

// GET/POST /api/reminders
app.http('remindersCollection', {
  methods: ['GET', 'POST', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'reminders',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }

    if (request.method === 'GET') {
      const items = await store.listReminders(user.oid);
      return json(200, { reminders: items });
    }

    // POST — create
    const body = await request.json().catch(() => ({}));
    const title = (body.title || '').toString().trim();
    if (!title) return json(400, { error: 'title is required' });
    const time = body.time ? String(body.time).slice(0, 5) : null;
    if (time && !/^\d{2}:\d{2}$/.test(time)) return json(400, { error: 'time must be HH:MM' });
    const tags = Array.isArray(body.tags) ? body.tags.map(String).map(s => s.trim()).filter(Boolean).slice(0, 8) : [];
    const priority = body.priority === 'high' ? 'high' : 'normal';
    const dueAt = typeof body.dueAt === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(body.dueAt)
      ? body.dueAt : store.todayKey();
    const description = typeof body.description === 'string' && body.description.trim()
      ? body.description.trim().slice(0, 2000) : null;

    const id = (globalThis.crypto?.randomUUID?.()) || `id-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
    const order = typeof body.order === 'number' && isFinite(body.order) ? body.order : null;
    const leadMinutes = typeof body.leadMinutes === 'number' && body.leadMinutes >= 0 && body.leadMinutes <= 240
      ? Math.floor(body.leadMinutes) : null;
    const reminder = { id, title, time, done: false, firedAt: null, createdDate: store.todayKey(), closedAt: null, tags, priority, order, leadMinutes, snoozedUntil: null, dueAt, description, rollDays: 0 };
    await store.upsertReminder(user.oid, reminder);
    return json(201, { reminder });
  },
});

// PATCH/DELETE /api/reminders/{id}
app.http('remindersItem', {
  methods: ['PATCH', 'DELETE', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'reminders/{id}',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }

    const id = request.params.id;
    if (!id) return json(400, { error: 'id is required' });

    if (request.method === 'DELETE') {
      const ok = await store.deleteReminder(user.oid, id);
      if (ok) return { status: 204, headers: corsHeaders() };
      return json(404, { error: 'not found' });
    }

    // PATCH — partial update
    const body = await request.json().catch(() => ({}));
    const existing = await store.getReminder(user.oid, id);
    if (!existing) return json(404, { error: 'not found' });
    if (typeof body.done === 'boolean') {
      existing.done = body.done;
      existing.closedAt = body.done ? new Date().toISOString() : null;
    }
    if (typeof body.title === 'string' && body.title.trim()) existing.title = body.title.trim();
    if (body.time === null) {
      existing.time = null;
      existing.firedAt = null;
    } else if (typeof body.time === 'string' && /^\d{2}:\d{2}$/.test(body.time)) {
      if (existing.time !== body.time) existing.firedAt = null;
      existing.time = body.time;
    }
    if (Array.isArray(body.tags)) {
      existing.tags = body.tags.map(String).map(s => s.trim()).filter(Boolean).slice(0, 8);
    }
    if (body.priority === 'high' || body.priority === 'normal') existing.priority = body.priority;
    if (typeof body.order === 'number' && isFinite(body.order)) existing.order = body.order;
    if (body.leadMinutes === null) existing.leadMinutes = null;
    else if (typeof body.leadMinutes === 'number' && body.leadMinutes >= 0 && body.leadMinutes <= 240) {
      existing.leadMinutes = Math.floor(body.leadMinutes);
    }
    // snoozedUntil: accept null to clear, or an ISO string to set
    if (body.snoozedUntil === null) existing.snoozedUntil = null;
    else if (typeof body.snoozedUntil === 'string' && !isNaN(Date.parse(body.snoozedUntil))) {
      existing.snoozedUntil = new Date(body.snoozedUntil).toISOString();
      existing.firedAt = null;
    }
    if (typeof body.dueAt === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(body.dueAt)) {
      if (existing.dueAt !== body.dueAt) existing.firedAt = null;
      existing.dueAt = body.dueAt;
      existing.rollDays = 0;
    }
    if (body.description === null || body.description === '') {
      existing.description = null;
    } else if (typeof body.description === 'string') {
      existing.description = body.description.trim().slice(0, 2000) || null;
    }
    await store.upsertReminder(user.oid, existing);
    return json(200, { reminder: existing });
  },
});
