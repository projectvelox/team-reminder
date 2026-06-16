// Tenant-shared customer registry (v1.7.9). Annotation layer over the per-license
// `customer` string field. Stores contact emails, address, and notes that apply
// across all of a customer's licenses.

const { app } = require('@azure/functions');
const { verifyTeamsToken } = require('../lib/auth');
const store = require('../lib/store');

function corsHeaders() {
  // v1.7.40 — origin locked to GitHub Pages prod. Auth is Bearer-token, so
  // CSRF is N/A, but tightening prevents arbitrary pages from attaching a
  // stolen token and reading our responses.
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

// v1.7.41 — POST /api/customers/merge
// Consolidate two (or more) customer rows into one canonical name.
// Body: { sourceNames: ["Beyond Innovations", "Beyond Innov"], targetName: "Beyond Innovations, Inc" }
//
// Effects:
//   1. Every license whose customer matches a sourceName (case/whitespace
//      insensitive) is updated to customer = targetName.
//   2. The target customer row's secondaryEmails absorbs any source's
//      primaryEmail + secondaryEmails (deduped).
//   3. The target's notes get a "Merged from: X, Y" appended.
//   4. Source customer rows are deleted.
//   5. An event is appended to each touched license so the activity feed
//      shows the merge.
app.http('customersMerge', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'customers/merge',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }

    const body = await request.json().catch(() => ({}));
    const sourceNames = Array.isArray(body.sourceNames)
      ? body.sourceNames.map((x) => String(x || '').trim()).filter(Boolean).slice(0, 10)
      : [];
    const targetName = String(body.targetName || '').trim();
    if (!targetName) return json(400, { error: 'targetName is required' });
    if (!sourceNames.length) return json(400, { error: 'sourceNames must include at least one source' });
    if (targetName.length > 200) return json(400, { error: 'targetName max 200 chars' });

    const norm = (s) => String(s || '').trim().toLowerCase();
    const sourceKeys = new Set(sourceNames.map(norm));
    sourceKeys.delete(norm(targetName)); // can't merge target into itself
    if (!sourceKeys.size) return json(400, { error: 'sourceNames cannot equal targetName' });

    // 1. Ensure / load the target customer row, then absorb source emails + notes.
    const customers = await store.listCustomers();
    const targetId = store.customerIdFromName(targetName);
    let target = customers.find((c) => c.id === targetId) || null;
    if (!target) {
      const now = new Date().toISOString();
      target = { id: targetId, name: targetName, primaryEmail: null, secondaryEmails: [], address: null, notes: null, createdAt: now, updatedAt: now };
    }
    const secondaries = new Set(Array.isArray(target.secondaryEmails) ? target.secondaryEmails : []);
    const mergedFromNames = [];
    for (const c of customers) {
      if (!sourceKeys.has(norm(c.name))) continue;
      if (c.primaryEmail && c.primaryEmail !== target.primaryEmail) secondaries.add(c.primaryEmail);
      for (const e of (Array.isArray(c.secondaryEmails) ? c.secondaryEmails : [])) {
        if (e && e !== target.primaryEmail) secondaries.add(e);
      }
      mergedFromNames.push(c.name);
    }
    target.secondaryEmails = Array.from(secondaries).slice(0, 10);
    if (mergedFromNames.length) {
      const stamp = `Merged from: ${mergedFromNames.join(', ')} (by ${user.name || user.oid} on ${new Date().toISOString().slice(0, 10)})`;
      target.notes = target.notes ? `${target.notes}\n\n${stamp}` : stamp;
      target.notes = target.notes.slice(0, 2000);
    }
    target.updatedAt = new Date().toISOString();
    await store.upsertCustomer(target);

    // 2. Rewrite license rows whose customer matches any source.
    const licenses = await store.listLicenses();
    let updated = 0;
    for (const lic of licenses) {
      if (!sourceKeys.has(norm(lic.customer))) continue;
      const prevName = lic.customer;
      lic.customer = targetName;
      lic.lastEditedAt = new Date().toISOString();
      lic.lastEditedByOid = user.oid;
      lic.lastEditedByName = user.name || null;
      lic.events = Array.isArray(lic.events) ? lic.events : [];
      lic.events.push({
        at: new Date().toISOString(),
        byOid: user.oid,
        byName: user.name || null,
        type: 'customerMerged',
        detail: `${prevName} -> ${targetName}`,
      });
      if (lic.events.length > 50) lic.events = lic.events.slice(-50);
      await store.upsertLicense(lic);
      updated++;
    }

    // 3. Delete the source customer rows.
    let deleted = 0;
    for (const c of customers) {
      if (!sourceKeys.has(norm(c.name))) continue;
      try { if (await store.deleteCustomer(c.id)) deleted++; } catch (_) { /* swallow per-row */ }
    }

    context.log(`[customers/merge] target="${targetName}" sources=${JSON.stringify(mergedFromNames)} licensesUpdated=${updated} customersDeleted=${deleted}`);
    return json(200, { target, licensesUpdated: updated, customersDeleted: deleted });
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
