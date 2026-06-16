// Test endpoint: POST /api/licenses/digest/preview
//
// Builds the same monthly digest the scheduler would send on day 1 of the
// month, but for the calling user only, and sends it to their mailbox right
// now. Lets owners preview the format without waiting for the next month-end.
// No-op on cold owners with no mailbox.

const { app } = require('@azure/functions');
const { verifyTeamsToken } = require('../lib/auth');
const store = require('../lib/store');
const { getUserBasic, sendMail } = require('../lib/graph');
const { phToday, bucketsForOwner, aggregateAllBuckets, buildDigestHtml } = require('../lib/digest');

function corsHeaders() {
  // v1.7.40 — see customers.js for rationale.
  return {
    'Access-Control-Allow-Origin': 'https://projectvelox.github.io',
    'Vary': 'Origin',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
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

app.http('digestPreview', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'licenses/digest/preview',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await verifyTeamsToken(request.headers.get('authorization')); }
    catch (err) { return json(err.status || 401, { error: err.message }); }

    // The endpoint always sends to the calling user. Safe regardless of body.
    const ownerOid = user.oid;

    // Look up the caller's email + display name from Graph (User.Read.All is
    // already consented tenant-wide, so this works even if they haven't opened
    // Day Reminders before).
    let profile;
    try { profile = await getUserBasic(ownerOid); }
    catch (err) { return json(502, { error: `Graph lookup failed: ${err.message}` }); }
    if (!profile || !profile.mail) return json(400, { error: 'No mailbox found for the calling user.' });

    let licenses;
    try { licenses = await store.listLicenses(); }
    catch (err) { return json(500, { error: `listLicenses failed: ${err.message}` }); }

    const today = phToday();
    const buckets = bucketsForOwner(licenses, ownerOid, today);
    const empty = buckets.overdue.length === 0 && buckets.thisMonth.length === 0 && buckets.nextMonth.length === 0;

    // Honor the rollup opt-in if set, OR an explicit body.rollup override so
    // the "show me all accounts" preview button can force-include the section
    // without changing the user's actual monthly digest preference.
    const body = await request.json().catch(() => ({}));
    const forceRollup = body && body.rollup === true;
    const stored = await store.getUser(ownerOid);
    const wantsRollup = forceRollup || !!stored.settings?.licenseRollupDigest;
    const rollupBuckets = wantsRollup ? aggregateAllBuckets(licenses, today) : null;

    const html = buildDigestHtml(profile.displayName || user.name, buckets, rollupBuckets);
    const totalThis = buckets.thisMonth.length + buckets.overdue.length;
    const subject = `[TEST] Day Reminders monthly digest: ${totalThis} renewals this month`;

    try {
      await sendMail({ to: { address: profile.mail, name: profile.displayName }, subject, body: html, contentType: 'HTML' });
    } catch (err) {
      return json(502, { error: `sendMail failed: ${err.message}` });
    }

    return json(200, {
      sentTo: profile.mail,
      counts: {
        overdue: buckets.overdue.length,
        thisMonth: buckets.thisMonth.length,
        nextMonth: buckets.nextMonth.length,
        rollup: wantsRollup,
      },
      empty,
    });
  },
});
