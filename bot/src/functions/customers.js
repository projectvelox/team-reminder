// Tenant-shared customer registry (v1.7.9). Annotation layer over the per-license
// `customer` string field. Stores contact emails, address, and notes that apply
// across all of a customer's licenses.

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

function json(status, body) {
  return { status, headers: corsHeaders(), body: JSON.stringify(body) };
}

async function authed(request) {
  const auth = request.headers.get('authorization');
  return await verifyTeamsToken(auth);
}

const EMAIL_RE = /^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/;

function validatePayload(body, existing) {
  const out = existing ? { ...existing } : {};
  if (body.name !== undefined) {
    const v = String(body.name || '').trim();
    if (!v) return { error: 'name is required' };
    if (v.length > 200) return { error: 'name max 200 chars' };
    out.name = v;
  } else if (!existing) {
    return { error: 'name is required' };
  }
  if (body.primaryEmail !== undefined) {
    const v = body.primaryEmail === null || body.primaryEmail === '' ? null : String(body.primaryEmail).trim();
    if (v && !EMAIL_RE.test(v)) return { error: 'primaryEmail is not a valid email address' };
    out.primaryEmail = v;
  }
  if (body.secondaryEmails !== undefined) {
    if (!Array.isArray(body.secondaryEmails)) return { error: 'secondaryEmails must be an array' };
    const cleaned = [];
    for (const e of body.secondaryEmails.slice(0, 10)) {
      const v = String(e || '').trim();
      if (!v) continue;
      if (!EMAIL_RE.test(v)) return { error: `secondaryEmails contains invalid address: ${v}` };
      cleaned.push(v);
    }
    out.secondaryEmails = cleaned;
  } else if (!existing) {
    out.secondaryEmails = [];
  }
  if (body.address !== undefined) {
    out.address = body.address === null || body.address === '' ? null : String(body.address).trim().slice(0, 500);
  }
  if (body.notes !== undefined) {
    out.notes = body.notes === null || body.notes === '' ? null : String(body.notes).trim().slice(0, 2000);
  }
  return { customer: out };
}

// GET/POST /api/customers
app.http('customersCollection', {
  methods: ['GET', 'POST', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'customers',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }

    if (request.method === 'GET') {
      const items = await store.listCustomers();
      return json(200, { customers: items });
    }
    // POST — create
    const body = await request.json().catch(() => ({}));
    const result = validatePayload(body, null);
    if (result.error) return json(400, { error: result.error });
    const now = new Date().toISOString();
    const customer = {
      ...result.customer,
      id: store.customerIdFromName(result.customer.name),
      createdAt: now,
      updatedAt: now,
    };
    await store.upsertCustomer(customer);
    return json(201, { customer });
  },
});

// PATCH/DELETE /api/customers/{id}
app.http('customersItem', {
  methods: ['PATCH', 'DELETE', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'customers/{id}',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }
    const id = request.params.id;
    if (!id) return json(400, { error: 'id is required' });
    if (request.method === 'DELETE') {
      const ok = await store.deleteCustomer(id);
      if (ok) return { status: 204, headers: corsHeaders() };
      return json(404, { error: 'not found' });
    }
    const existing = await store.getCustomer(id);
    if (!existing) return json(404, { error: 'not found' });
    const body = await request.json().catch(() => ({}));
    const result = validatePayload(body, existing);
    if (result.error) return json(400, { error: result.error });
    const merged = { ...result.customer, updatedAt: new Date().toISOString() };
    await store.upsertCustomer(merged);
    return json(200, { customer: merged });
  },
});
