// User search + photo proxy for the Licenses tab people picker (v1.7.13).
//
// Uses the existing app-only Graph token (User.Read.All application permission,
// already consented tenant-wide). Search results filter out Guests and rows
// with no mailbox (typically service accounts) so the picker shows real people.
//
// Auth: same Teams SSO Bearer token as /api/licenses. Anyone in the tenant who
// has the app open can search the directory.

const { app } = require('@azure/functions');
const { verifyTeamsToken } = require('../lib/auth');
const { searchUsers, getUserPhoto } = require('../lib/graph');

function corsHeaders(extra) {
  // v1.7.40 — see customers.js for rationale. `extra` is used by the photo
  // endpoint to override Content-Type and Cache-Control.
  return {
    'Access-Control-Allow-Origin': 'https://projectvelox.github.io',
    'Vary': 'Origin',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
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

// GET /api/users/search?q=<text>
app.http('usersSearch', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'users/search',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    try { await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }

    // request.query is a URLSearchParams in Functions v4.
    const q = (request.query && request.query.get && request.query.get('q')) || '';
    if (!q || String(q).trim().length < 2) return json(200, { users: [] });

    try {
      const users = await searchUsers(q);
      return json(200, { users });
    } catch (err) {
      context.error(`[users/search] ${err?.message || err}`);
      return json(err.status || 500, { error: err.message || 'search failed' });
    }
  },
});

// GET /api/users/{oid}/photo — proxies the Graph user photo. Returns 204 if
// the user has no photo (so the tab can render initials fallback without an
// error spike in the console).
app.http('usersPhoto', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'users/{oid}/photo',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    try { await authed(request); } catch (err) { return json(err.status || 401, { error: err.message }); }

    const oid = request.params.oid;
    if (!oid) return json(400, { error: 'oid is required' });

    try {
      const photo = await getUserPhoto(oid);
      if (!photo) {
        // v1.7.40 — let the browser remember the 204 for 5 min so we don't
        // re-pound Graph every render for a user who has no photo.
        return { status: 204, headers: corsHeaders({ 'Cache-Control': 'private, max-age=300' }) };
      }
      // v1.7.40 — restrict origin, extend client cache to 24h, allow ~5 min
      // staleness while we revalidate. Profile photos change rarely.
      return {
        status: 200,
        headers: corsHeaders({
          'Content-Type': photo.contentType,
          'Cache-Control': 'private, max-age=86400, stale-while-revalidate=300',
        }),
        body: photo.buffer,
      };
    } catch (err) {
      context.error(`[users/photo] ${err?.message || err}`);
      return json(err.status || 500, { error: err.message || 'photo fetch failed' });
    }
  },
});
