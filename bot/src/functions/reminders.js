// REST API for the tab to read/write reminders.
// All endpoints require a Teams SSO Bearer token; user identity = oid claim.

const { app } = require('@azure/functions');
const { verifyTeamsToken } = require('../lib/auth');
const store = require('../lib/store');

function corsHeaders() {
  // v1.7.40 — see customers.js for rationale.
  return {
    'Access-Control-Allow-Origin': 'https://projectvelox.github.io',
    'Vary': 'Origin',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
    'Access-Control-Allow-Methods': 'GET, POST, PATCH, DELETE, OPTIONS',
    'Access-Control-Max-Age': '600',
    'Content-Type': 'application/json',
    'X-Content-Type-Options': 'nosniff',
    'Referrer-Policy': 'no-referrer',
    'Cache-Control': 'no-store',
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
    const client = typeof body.client === 'string' && body.client.trim()
      ? body.client.trim().slice(0, 100) : null;

    const id = (globalThis.crypto?.randomUUID?.()) || `id-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
    const order = typeof body.order === 'number' && isFinite(body.order) ? body.order : null;
    const leadMinutes = typeof body.leadMinutes === 'number' && body.leadMinutes >= 0 && body.leadMinutes <= 240
      ? Math.floor(body.leadMinutes) : null;
    const repeat = body.repeat === 'daily' || body.repeat === 'weekdays' || body.repeat === 'weekly' ? body.repeat : 'none';
    const subtasks = Array.isArray(body.subtasks)
      ? body.subtasks
          .filter((s) => s && typeof s === 'object' && typeof s.text === 'string' && s.text.trim())
          .map((s) => ({
            id: typeof s.id === 'string' && s.id ? s.id : `s-${Math.random().toString(36).slice(2, 10)}`,
            text: String(s.text).trim().slice(0, 500),
            done: !!s.done,
          }))
          .slice(0, 50)
      : [];
    const reminder = { id, title, time, done: false, firedAt: null, createdDate: store.todayKey(), closedAt: null, tags, priority, order, leadMinutes, snoozedUntil: null, dueAt, description, rollDays: 0, client, repeat, subtasks };
    await store.upsertReminder(user.oid, reminder);
    return json(201, { reminder });
  },
});

// Compute the next occurrence date for a recurring reminder.
// `currentDate` is YYYY-MM-DD; `repeat` is 'daily' | 'weekdays' | 'weekly'.
function nextOccurrence(currentDate, repeat) {
  const d = new Date(currentDate + 'T00:00:00Z');
  if (repeat === 'daily') {
    d.setUTCDate(d.getUTCDate() + 1);
  } else if (repeat === 'weekly') {
    d.setUTCDate(d.getUTCDate() + 7);
  } else if (repeat === 'weekdays') {
    do { d.setUTCDate(d.getUTCDate() + 1); }
    while (d.getUTCDay() === 0 || d.getUTCDay() === 6);
  } else {
    return currentDate;
  }
  return d.toISOString().slice(0, 10);
}

// Shared field-by-field patch logic. Mutates `existing` in place using values from `body`.
// Only validated fields are applied; unknown or invalid fields are ignored.
function applyPatch(existing, body) {
  const isRecurring = existing.repeat && existing.repeat !== 'none';
  if (typeof body.done === 'boolean') {
    if (body.done && isRecurring) {
      // Recurring "done" = complete this occurrence and roll forward to next.
      // The reminder stays open with a new dueAt; firedAt and snooze reset.
      const base = existing.dueAt || store.todayKey();
      existing.dueAt = nextOccurrence(base, existing.repeat);
      existing.firedAt = null;
      existing.snoozedUntil = null;
      existing.rollDays = 0;
      // done/closedAt are intentionally left untouched (stays open).
    } else {
      existing.done = body.done;
      existing.closedAt = body.done ? new Date().toISOString() : null;
    }
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
  if (body.client === null || body.client === '') {
    existing.client = null;
  } else if (typeof body.client === 'string') {
    existing.client = body.client.trim().slice(0, 100) || null;
  }
  if (body.repeat === null || body.repeat === 'none') existing.repeat = 'none';
  else if (body.repeat === 'daily' || body.repeat === 'weekdays' || body.repeat === 'weekly') {
    existing.repeat = body.repeat;
  }
  if (Array.isArray(body.subtasks)) {
    existing.subtasks = body.subtasks
      .filter((s) => s && typeof s === 'object' && typeof s.text === 'string' && s.text.trim())
      .map((s) => ({
        id: typeof s.id === 'string' && s.id ? s.id : `s-${Math.random().toString(36).slice(2, 10)}`,
        text: String(s.text).trim().slice(0, 500),
        done: !!s.done,
      }))
      .slice(0, 50);
  }
}

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
    applyPatch(existing, body);
    await store.upsertReminder(user.oid, existing);
    return json(200, { reminder: existing });
  },
});

// POST /api/reminders/bulk — apply the same patch to many reminders in one round-trip.
// Body: { ids: ["id1", "id2", ...], patch: { done: true } }
// Returns: { updated: [reminder...], notFound: [id...] }
app.http('remindersBulk', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'reminders/bulk',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }

    const body = await request.json().catch(() => ({}));
    const ids = Array.isArray(body.ids) ? body.ids.filter(x => typeof x === 'string').slice(0, 500) : [];
    const patch = body.patch && typeof body.patch === 'object' ? body.patch : null;
    if (!ids.length) return json(400, { error: 'ids array is required' });
    if (!patch) return json(400, { error: 'patch object is required' });

    const updated = [];
    const notFound = [];
    for (const id of ids) {
      const existing = await store.getReminder(user.oid, id);
      if (!existing) { notFound.push(id); continue; }
      applyPatch(existing, patch);
      await store.upsertReminder(user.oid, existing);
      updated.push(existing);
    }
    return json(200, { updated, notFound });
  },
});
