// @ts-check
// v1.7.43 — GDPR data-subject-rights endpoints.
//
//   GET    /api/me/export — dump everything we store about the calling user as
//          a single JSON download (settings, reminders, templates, the conversation
//          ref, owner-of license rows). For data-portability + transparency.
//   DELETE /api/me/data   — wipe everything we store about the calling user.
//          Reminders + templates + settings + conversation ref. Tenant-shared
//          license rows are NOT wiped (they're tenant data); ownerOid is set
//          to null so the user is unassigned, with a clear note.
//
// Auth: same Teams SSO Bearer token as the rest of the API.

const { app } = require('@azure/functions');
const { verifyTeamsToken } = require('../lib/auth');
const store = require('../lib/store');

function corsHeaders(extra) {
  return {
    'Access-Control-Allow-Origin': 'https://projectvelox.github.io',
    'Vary': 'Origin',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
    'Access-Control-Allow-Methods': 'GET, DELETE, OPTIONS',
    'Access-Control-Max-Age': '600',
    'Content-Type': 'application/json',
    'X-Content-Type-Options': 'nosniff',
    'Referrer-Policy': 'no-referrer',
    'Cache-Control': 'no-store',
    ...(extra || {}),
  };
}

function json(status, body) {
  return { status, headers: corsHeaders(), body: JSON.stringify(body) };
}

async function authed(request) {
  const auth = request.headers.get('authorization');
  return await verifyTeamsToken(auth);
}

app.http('meExport', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'me/export',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }

    const oid = user.oid;
    const [userRow, reminders, templates, allLicenses] = await Promise.all([
      store.getUser(oid),
      store.listReminders(oid),
      store.getTemplates(oid).catch(() => []),
      store.listLicenses().catch(() => []),
    ]);
    const ownedLicenses = allLicenses.filter((l) => l.ownerOid === oid);

    const payload = {
      schema: 'day-reminders.user-export.v1',
      generatedAt: new Date().toISOString(),
      generatedFor: { oid, name: user.name || null, upn: user.upn || null },
      userRecord: {
        settings: userRow.settings,
        // The conversationRef is included so the user can see we have a Teams
        // chat handle for them; we don't expose tokens.
        hasConversationRef: !!userRow.conversationRef,
        tenantId: userRow.tenantId,
        displayName: userRow.displayName,
        lastEodDate: userRow.lastEodDate,
        lastRolloverDate: userRow.lastRolloverDate,
        lastBriefingDate: userRow.lastBriefingDate,
        lastDigestSentMonth: userRow.lastDigestSentMonth,
        coldNudgedAt: userRow.coldNudgedAt,
      },
      reminders,
      templates,
      // Tenant-shared, but include the rows where the user is the owner so the
      // export reflects every place we've recorded their identity.
      ownedLicenses,
    };

    // Force a download in the browser by hinting a filename.
    return {
      status: 200,
      headers: corsHeaders({
        'Content-Disposition': `attachment; filename="day-reminders-export-${oid.slice(0, 8)}-${new Date().toISOString().slice(0, 10)}.json"`,
      }),
      body: JSON.stringify(payload, null, 2),
    };
  },
});

app.http('meDelete', {
  methods: ['DELETE', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'me/data',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }

    const oid = user.oid;
    let reminderCount = 0;
    try {
      const rs = await store.listReminders(oid);
      reminderCount = rs.length;
      for (const r of rs) await store.deleteReminder(oid, r.id);
    } catch (err) { context.error(`[me/data] reminders wipe: ${err?.message || err}`); }

    // Templates + settings + conversation ref: reset the whole user row.
    let userReset = false;
    try {
      const defaultUser = await store.getUser('__not_a_real_oid__');
      // Reuse getUser's default-settings path by writing defaults via upsertUser.
      await store.upsertUser(oid, {
        settings: { ...defaultUser.settings },
        conversationRef: null,
        lastEodDate: null,
        lastRolloverDate: null,
        tenantId: null,
        serviceUrl: null,
        displayName: null,
        lastBriefingDate: null,
        briefingSnoozedUntil: null,
        coldNudgedAt: null,
        lastDigestSentMonth: null,
      });
      userReset = true;
    } catch (err) { context.error(`[me/data] user reset: ${err?.message || err}`); }

    // Reset templates blob.
    try { await store.setTemplates(oid, []); } catch (err) { context.error(`[me/data] templates wipe: ${err?.message || err}`); }

    // Tenant-shared licenses: unassign rows where this user is the owner so
    // their identity is removed from the operational dataset. The license itself
    // (a tenant artifact) is NOT deleted.
    let unassigned = 0;
    try {
      const all = await store.listLicenses();
      for (const lic of all) {
        if (lic.ownerOid !== oid) continue;
        lic.ownerOid = null;
        lic.ownerName = null;
        lic.lastEditedAt = new Date().toISOString();
        lic.lastEditedByOid = oid;
        lic.lastEditedByName = '(account data deleted)';
        lic.events = Array.isArray(lic.events) ? lic.events : [];
        lic.events.push({
          at: new Date().toISOString(),
          byOid: oid,
          byName: '(account data deleted)',
          type: 'ownerChanged',
          detail: 'former owner requested data deletion',
        });
        if (lic.events.length > 50) lic.events = lic.events.slice(-50);
        await store.upsertLicense(lic);
        unassigned++;
      }
    } catch (err) { context.error(`[me/data] license unassign: ${err?.message || err}`); }

    context.log(`[me/data] wiped oid=${oid} reminders=${reminderCount} userReset=${userReset} licensesUnassigned=${unassigned}`);
    return json(200, { reminderCount, userReset, licensesUnassigned: unassigned });
  },
});
