// REST API for the Licenses tab (v1.7).
// Tenant-shared data — every authenticated tenant user can read and write all rows.
// Each row has an Owner (ownerOid + ownerName) who receives Teams escalation cards
// and the monthly email digest.

const { app } = require('@azure/functions');
const { verifyTeamsToken } = require('../lib/auth');
const store = require('../lib/store');
const { getUserBasic, sendMail } = require('../lib/graph');

// Cold-owner nudge (v1.7.20). Sends a one-time email to a newly-assigned owner
// who has never opened Day Reminders. Dedupe by coldNudgedAt on the user record.
//
// Fire-and-forget: callers don't await this so a Graph slowdown can't make
// /api/licenses slow. Errors are logged but don't surface to the user.
async function maybeNudgeColdOwner({ ownerOid, customer, licenseType, assignedByName }) {
  if (!ownerOid) return;
  try {
    const user = await store.getUser(ownerOid);
    if (user.conversationRef) return; // they've opened the bot, no need
    if (user.coldNudgedAt) return; // we already nudged them once
    const profile = await getUserBasic(ownerOid).catch(() => null);
    if (!profile || !profile.mail) return;
    const firstName = (profile.displayName || '').split(/\s+/)[0] || 'there';
    const deepLink = 'https://teams.microsoft.com/l/entity/5a03bfa3-63c4-417c-b668-b02234ebc11b/dayReminders.licenses';
    const subject = `You're the owner of a license renewal in Day Reminders`;
    const html = `
      <p>Hi ${escapeHtml(firstName)},</p>
      <p>${assignedByName ? escapeHtml(assignedByName) + ' has assigned' : 'You have been assigned'} you as the owner of a license renewal in Day Reminders:</p>
      <ul>
        <li><strong>${escapeHtml(customer || 'a customer')}</strong> &mdash; ${escapeHtml(licenseType || 'license')}</li>
      </ul>
      <p>To get Teams alerts ahead of each renewal, open Day Reminders once:</p>
      <p><a href="${deepLink}">Open Day Reminders &rarr; Licenses</a></p>
      <p>Once you've opened it, the bot will start chasing you 14 days before each expiry, with a quick morning briefing on weekdays summarising what needs attention.</p>
      <p style="color:#888;font-size:12px;margin-top:24px">Sent automatically by Day Reminders. You'll only get this once.</p>
    `;
    await sendMail({ to: { address: profile.mail, name: profile.displayName }, subject, body: html, contentType: 'HTML' });
    // Mark nudged so we don't spam on reassignment churn.
    user.coldNudgedAt = new Date().toISOString();
    user.displayName = user.displayName || profile.displayName || null;
    await store.upsertUser(ownerOid, user);
  } catch (err) {
    // Log only; never let nudge failure affect the parent request.
    // eslint-disable-next-line no-console
    console.error(`[coldNudge] ${ownerOid}: ${err && err.message ? err.message : err}`);
  }
}

function escapeHtml(s) {
  return String(s || '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
}

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

// Tenant-shared = every caller is implicitly an editor today.
// We still register them as a member so the Owner picker auto-populates.
async function registerCaller(user) {
  try {
    await store.registerMember({ oid: user.oid, displayName: user.name, upn: user.upn });
  } catch { /* member registration is best-effort */ }
}

function newId() {
  return (globalThis.crypto?.randomUUID?.()) || `id-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

function appendEvent(license, type, user, detail) {
  if (!Array.isArray(license.events)) license.events = [];
  license.events.push({
    at: new Date().toISOString(),
    byOid: user.oid,
    byName: user.name || null,
    type,
    detail: detail || null,
  });
  if (license.events.length > 50) license.events = license.events.slice(-50);
}

// Validate + coerce a license payload. Returns either { license: {...} } or { error: "..." }.
// `existing` is the current row when patching; null when creating.
function validatePayload(body, existing) {
  const out = existing ? { ...existing } : {};

  if (body.customer !== undefined) {
    const v = String(body.customer || '').trim();
    if (!v) return { error: 'customer is required' };
    if (v.length > 200) return { error: 'customer max 200 chars' };
    out.customer = v;
  } else if (!existing) {
    return { error: 'customer is required' };
  }

  if (body.licenseType !== undefined) {
    const v = String(body.licenseType || '').trim();
    if (!v) return { error: 'licenseType is required' };
    if (v.length > 200) return { error: 'licenseType max 200 chars' };
    out.licenseType = v;
  } else if (!existing) {
    return { error: 'licenseType is required' };
  }

  if (body.userCount !== undefined) {
    const n = Number(body.userCount);
    if (!isFinite(n) || n < 0 || n > 1000000) return { error: 'userCount must be a non-negative number' };
    out.userCount = Math.floor(n);
  } else if (!existing) {
    out.userCount = 0;
  }

  if (body.expiryDate !== undefined) {
    if (!ISO_DATE.test(String(body.expiryDate || ''))) return { error: 'expiryDate must be YYYY-MM-DD' };
    out.expiryDate = body.expiryDate;
  } else if (!existing) {
    return { error: 'expiryDate is required' };
  }

  if (body.ownerOid !== undefined) {
    const oid = String(body.ownerOid || '').trim();
    if (!oid) return { error: 'ownerOid is required' };
    out.ownerOid = oid;
  } else if (!existing) {
    return { error: 'ownerOid is required' };
  }

  if (body.ownerName !== undefined) {
    out.ownerName = String(body.ownerName || '').trim().slice(0, 100) || null;
  }

  if (body.productLine !== undefined) {
    const v = String(body.productLine || '').trim();
    out.productLine = v ? v.slice(0, 100) : null;
  }

  if (body.leadDays === null) {
    out.leadDays = null;
  } else if (body.leadDays !== undefined) {
    // v1.7.37: accept array of 0-365 ints. Scalar legacy clients are coerced
    // to [n] so older Outlook/compose paths keep working without a redeploy.
    let arr = body.leadDays;
    if (typeof arr === 'number') arr = [arr];
    if (!Array.isArray(arr)) return { error: 'leadDays must be an array of 0-365 ints' };
    const cleaned = arr
      .map((x) => Number(x))
      .filter((n) => Number.isFinite(n) && n >= 0 && n <= 365)
      .map((n) => Math.floor(n));
    if (cleaned.length > 10) return { error: 'leadDays max 10 thresholds' };
    out.leadDays = cleaned.length
      ? Array.from(new Set(cleaned)).sort((a, b) => b - a)
      : null;
  }

  if (body.notes === null || body.notes === '') {
    out.notes = null;
  } else if (body.notes !== undefined) {
    out.notes = String(body.notes).trim().slice(0, 2000) || null;
  }

  if (body.state !== undefined) {
    out.state = body.state === 'abandoned' ? 'abandoned' : 'active';
  } else if (!existing) {
    out.state = 'active';
  }

  if (body.status !== undefined) {
    if (!store.LICENSE_STATUSES.includes(body.status)) return { error: `status must be one of ${store.LICENSE_STATUSES.join(', ')}` };
    out.status = body.status;
  } else if (!existing) {
    out.status = 'notStarted';
  }

  if (body.renewalCycle !== undefined) {
    if (!store.RENEWAL_CYCLES.includes(body.renewalCycle)) return { error: `renewalCycle must be one of ${store.RENEWAL_CYCLES.join(', ')}` };
    out.renewalCycle = body.renewalCycle;
  } else if (!existing) {
    out.renewalCycle = 'annual';
  }

  return { license: out };
}

// GET/POST /api/licenses
app.http('licensesCollection', {
  methods: ['GET', 'POST', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'licenses',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }
    await registerCaller(user);

    if (request.method === 'GET') {
      const items = await store.listLicenses();
      return json(200, { licenses: items });
    }

    // POST — create
    const body = await request.json().catch(() => ({}));
    const result = validatePayload(body, null);
    if (result.error) return json(400, { error: result.error });

    const now = new Date().toISOString();
    const license = {
      ...result.license,
      id: newId(),
      createdAt: now,
      createdByOid: user.oid,
      createdByName: user.name || null,
      lastEditedAt: now,
      lastEditedByOid: user.oid,
      lastEditedByName: user.name || null,
      lastRenewedAt: null,
      lastFollowUpAt: null,
      lastEscalatedDays: null,
      lastFiredLeadDays: [],
      statusChangedAt: now,
      statusChangedByOid: user.oid,
      statusChangedByName: user.name || null,
      events: [],
    };
    appendEvent(license, 'created', user, `${license.customer} · ${license.licenseType}`);
    await store.upsertLicense(license);
    // Ensure a stub customer entry exists so the registry stays in sync.
    try { await store.ensureCustomer(license.customer); } catch {}
    // Cold-owner nudge: if the new owner has never opened Day Reminders, send
    // a one-time email so they know to install. Fire-and-forget.
    maybeNudgeColdOwner({
      ownerOid: license.ownerOid,
      customer: license.customer,
      licenseType: license.licenseType,
      assignedByName: user.name || null,
    });
    return json(201, { license });
  },
});

// PATCH/DELETE /api/licenses/{id}
app.http('licensesItem', {
  methods: ['PATCH', 'DELETE', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'licenses/{id}',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }
    await registerCaller(user);

    const id = request.params.id;
    if (!id) return json(400, { error: 'id is required' });

    if (request.method === 'DELETE') {
      const ok = await store.deleteLicense(id);
      if (ok) return { status: 204, headers: corsHeaders() };
      return json(404, { error: 'not found' });
    }

    const existing = await store.getLicense(id);
    if (!existing) return json(404, { error: 'not found' });
    const body = await request.json().catch(() => ({}));
    const result = validatePayload(body, existing);
    if (result.error) return json(400, { error: result.error });

    const now = new Date().toISOString();
    const merged = {
      ...result.license,
      lastEditedAt: now,
      lastEditedByOid: user.oid,
      lastEditedByName: user.name || null,
    };
    // Clear escalation + follow-up tracking when expiry moves forward.
    if (existing.expiryDate !== merged.expiryDate) {
      merged.lastEscalatedDays = null;
      merged.lastFiredLeadDays = [];
    }
    // Status transition: stamp + log + reset follow-up timer.
    if (existing.status !== merged.status) {
      merged.statusChangedAt = now;
      merged.statusChangedByOid = user.oid;
      merged.statusChangedByName = user.name || null;
      merged.lastFollowUpAt = null;
      appendEvent(merged, 'statusChanged', user, `${existing.status} -> ${merged.status}`);
    }
    // Owner reassignment is worth logging.
    if (existing.ownerOid !== merged.ownerOid) {
      appendEvent(merged, 'ownerChanged', user, `${existing.ownerName || existing.ownerOid || '(none)'} -> ${merged.ownerName || merged.ownerOid || '(none)'}`);
      // Cold-owner nudge for the new owner.
      maybeNudgeColdOwner({
        ownerOid: merged.ownerOid,
        customer: merged.customer,
        licenseType: merged.licenseType,
        assignedByName: user.name || null,
      });
    }
    // Expiry change worth logging too.
    if (existing.expiryDate !== merged.expiryDate) {
      appendEvent(merged, 'expiryChanged', user, `${existing.expiryDate || '?'} -> ${merged.expiryDate || '?'}`);
    }
    await store.upsertLicense(merged);
    return json(200, { license: merged });
  },
});

// POST /api/licenses/{id}/renew
// Body: { newExpiryDate: "YYYY-MM-DD" }  -- explicit new expiry
//   or: { years: 1 | 2 | 3 }              -- advance by N years from current expiry
app.http('licensesRenew', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'licenses/{id}/renew',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }
    await registerCaller(user);

    const id = request.params.id;
    if (!id) return json(400, { error: 'id is required' });

    const existing = await store.getLicense(id);
    if (!existing) return json(404, { error: 'not found' });

    const body = await request.json().catch(() => ({}));
    let newExpiry = null;
    if (typeof body.newExpiryDate === 'string' && ISO_DATE.test(body.newExpiryDate)) {
      newExpiry = body.newExpiryDate;
    } else if (body.years === 1 || body.years === 2 || body.years === 3) {
      const base = existing.expiryDate && ISO_DATE.test(existing.expiryDate)
        ? new Date(existing.expiryDate + 'T00:00:00Z')
        : new Date();
      base.setUTCFullYear(base.getUTCFullYear() + body.years);
      newExpiry = base.toISOString().slice(0, 10);
    } else {
      return json(400, { error: 'provide newExpiryDate (YYYY-MM-DD) or years (1, 2, or 3)' });
    }

    const now = new Date().toISOString();
    const merged = {
      ...existing,
      expiryDate: newExpiry,
      state: 'active',
      status: 'renewed',
      statusChangedAt: now,
      statusChangedByOid: user.oid,
      statusChangedByName: user.name || null,
      lastRenewedAt: now,
      lastFollowUpAt: null,
      lastEditedAt: now,
      lastEditedByOid: user.oid,
      lastEditedByName: user.name || null,
      lastEscalatedDays: null,
      lastFiredLeadDays: [],
    };
    appendEvent(merged, 'renewed', user, `expiry advanced to ${newExpiry}`);
    await store.upsertLicense(merged);
    return json(200, { license: merged });
  },
});

// v1.7.39 — POST /api/licenses/{id}/comments
// Append a comment to the license's thread. Body: { text }.
app.http('licensesComments', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'licenses/{id}/comments',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }
    await registerCaller(user);

    const id = request.params.id;
    if (!id) return json(400, { error: 'id is required' });
    const existing = await store.getLicense(id);
    if (!existing) return json(404, { error: 'not found' });

    const body = await request.json().catch(() => ({}));
    const text = String(body.text || '').trim();
    if (!text) return json(400, { error: 'text is required' });
    // v1.7.40 — Azure Tables strings cap at 32KB UTF-16. The whole comments
    // array is one JSON-encoded property, so cap individual text + total count
    // conservatively. 30 × 1000 chars + JSON overhead stays comfortably under.
    if (text.length > 1000) return json(400, { error: 'text max 1000 chars' });

    const now = new Date().toISOString();
    const comment = {
      id: newId(),
      at: now,
      byOid: user.oid,
      byName: user.name || null,
      text,
    };
    const comments = Array.isArray(existing.comments) ? existing.comments.slice() : [];
    comments.push(comment);
    if (comments.length > 30) comments.splice(0, comments.length - 30);

    const merged = {
      ...existing,
      comments,
      lastEditedAt: now,
      lastEditedByOid: user.oid,
      lastEditedByName: user.name || null,
    };
    await store.upsertLicense(merged);
    return json(201, { comment, license: merged });
  },
});

// POST /api/licenses/bulk — apply the same patch to many licenses (e.g. bulk reassign Owner).
// Body: { ids: ["id1", ...], patch: { ownerOid, ownerName } }
app.http('licensesBulk', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'licenses/bulk',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }
    await registerCaller(user);

    const body = await request.json().catch(() => ({}));
    const ids = Array.isArray(body.ids) ? body.ids.filter(x => typeof x === 'string').slice(0, 500) : [];
    const patch = body.patch && typeof body.patch === 'object' ? body.patch : null;
    if (!ids.length) return json(400, { error: 'ids array is required' });
    if (!patch) return json(400, { error: 'patch object is required' });

    const now = new Date().toISOString();
    const updated = [];
    const notFound = [];
    const nudgedOids = new Set();
    for (const id of ids) {
      const existing = await store.getLicense(id);
      if (!existing) { notFound.push(id); continue; }
      const result = validatePayload(patch, existing);
      if (result.error) continue;
      const merged = {
        ...result.license,
        lastEditedAt: now,
        lastEditedByOid: user.oid,
        lastEditedByName: user.name || null,
      };
      await store.upsertLicense(merged);
      updated.push(merged);
      // Cold-owner nudge once per oid even across a bulk reassign of many rows.
      if (existing.ownerOid !== merged.ownerOid && merged.ownerOid && !nudgedOids.has(merged.ownerOid)) {
        nudgedOids.add(merged.ownerOid);
        maybeNudgeColdOwner({
          ownerOid: merged.ownerOid,
          customer: merged.customer,
          licenseType: merged.licenseType,
          assignedByName: user.name || null,
        });
      }
    }
    return json(200, { updated, notFound });
  },
});
