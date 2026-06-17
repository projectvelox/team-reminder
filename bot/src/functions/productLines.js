// Tenant-shared product-line registry (v1.8.0).
//
// Strict controlled vocab for License.productLine. The Licenses tab's Edit
// dialog renders a <select> populated from this registry; admins manage the
// list via the Product Lines admin dialog. Canonical seed (M365 / Business
// Central / Finance and Operation / PHILTAX / CRM / Security) ships on first
// GET if the partition is empty.
//
// All authenticated tenant users can read; PUT/DELETE are not gated to admin
// at the API layer (the tenant is small, mutually-trusted Kation staff).

const { app } = require('@azure/functions');
const { verifyTeamsToken } = require('../lib/auth');
const store = require('../lib/store');

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': 'https://projectvelox.github.io',
    'Vary': 'Origin',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
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

// GET  /api/product-lines               - list (seeds canonical 6 on first call)
// PUT  /api/product-lines               - bulk replace { productLines: [{name, sortOrder?}] }
app.http('productLinesCollection', {
  methods: ['GET', 'PUT', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'product-lines',
  handler: async (request) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }

    if (request.method === 'GET') {
      let items = await store.listProductLines();
      if (items.length === 0) items = await store.ensureProductLinesSeeded(user);
      return json(200, { productLines: items });
    }

    // PUT — full set replace. Deletes any registry rows not in the payload.
    const body = await request.json().catch(() => ({}));
    const incoming = Array.isArray(body.productLines) ? body.productLines : [];
    const now = new Date().toISOString();
    const desiredNames = new Set();
    const saved = [];
    for (let i = 0; i < incoming.slice(0, 50).length; i++) {
      const item = incoming[i];
      const name = String(item.name || '').trim().slice(0, 100);
      if (!name) continue;
      desiredNames.add(name);
      const sortOrder = typeof item.sortOrder === 'number' ? item.sortOrder : i;
      await store.upsertProductLine({
        name,
        sortOrder,
        createdAt: item.createdAt || now,
        createdByOid: item.createdByOid || user.oid,
        createdByName: item.createdByName || user.name || null,
      });
      saved.push({ name, sortOrder });
    }
    // Delete any registry rows the caller dropped from the payload.
    const existing = await store.listProductLines();
    for (const e of existing) {
      if (!desiredNames.has(e.name)) {
        await store.deleteProductLine(e.name);
      }
    }
    saved.sort((a, b) => (a.sortOrder - b.sortOrder) || a.name.localeCompare(b.name));
    return json(200, { productLines: saved });
  },
});

// DELETE /api/product-lines/{name}      - single delete
app.http('productLinesItem', {
  methods: ['DELETE', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'product-lines/{name}',
  handler: async (request) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    try { await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }
    const name = decodeURIComponent(request.params.name || '');
    const ok = await store.deleteProductLine(name);
    if (ok) return { status: 204, headers: corsHeaders() };
    return json(404, { error: 'not found' });
  },
});
