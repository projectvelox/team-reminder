// GET / PUT /api/settings — per-user preferences.

const { app } = require('@azure/functions');
const { verifyTeamsToken } = require('../lib/auth');
const store = require('../lib/store');

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
    'Access-Control-Allow-Methods': 'GET, PUT, OPTIONS',
    'Content-Type': 'application/json',
  };
}

function json(status, body) {
  return { status, headers: corsHeaders(), body: JSON.stringify(body) };
}

app.http('settings', {
  methods: ['GET', 'PUT', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'settings',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    let user;
    try { user = await verifyTeamsToken(request.headers.get('authorization')); }
    catch (err) { return json(err.status || 401, { error: err.message }); }

    if (request.method === 'GET') {
      const u = await store.getUser(user.oid);
      return json(200, { settings: u.settings, hasBot: !!u.conversationRef });
    }

    const body = await request.json().catch(() => ({}));
    const incoming = body.settings || {};
    const validHHMM = (s) => typeof s === 'string' && /^\d{2}:\d{2}$/.test(s);
    const settings = {
      eodTime: validHHMM(incoming.eodTime) ? incoming.eodTime : store.DEFAULT_SETTINGS.eodTime,
      leadMinutes: clampInt(incoming.leadMinutes, 0, 240, store.DEFAULT_SETTINGS.leadMinutes),
      weekdaysOnly: !!incoming.weekdaysOnly,
      notifications: incoming.notifications !== false,
      quietStart: validHHMM(incoming.quietStart) ? incoming.quietStart : null,
      quietEnd: validHHMM(incoming.quietEnd) ? incoming.quietEnd : null,
      autoImportFlagged: !!incoming.autoImportFlagged,
      // License-tab settings (v1.7.22; v1.7.37 widens to array)
      licenseLeadDays: cleanLeadDaysArray(incoming.licenseLeadDays, store.DEFAULT_SETTINGS.licenseLeadDays),
      licenseSkipBriefing: !!incoming.licenseSkipBriefing,
      licenseSkipMonthlyDigest: !!incoming.licenseSkipMonthlyDigest,
      licenseRollupDigest: !!incoming.licenseRollupDigest,
    };
    // Both must be set, or both null — partial config means quiet hours disabled.
    if (!settings.quietStart || !settings.quietEnd) {
      settings.quietStart = null;
      settings.quietEnd = null;
    }
    await store.upsertUser(user.oid, { settings });
    return json(200, { settings });
  },
});

function clampInt(v, min, max, fallback) {
  const n = parseInt(v, 10);
  if (isNaN(n)) return fallback;
  return Math.max(min, Math.min(max, n));
}

// Accept array of 0-365 ints (or a scalar from legacy clients). Dedupes,
// sorts descending, caps at 10 entries. Falls back to the system default if
// the cleaned array is empty.
function cleanLeadDaysArray(v, fallback) {
  let arr = v;
  if (typeof arr === 'number') arr = [arr];
  if (!Array.isArray(arr)) return Array.isArray(fallback) ? fallback.slice() : [14];
  const cleaned = arr
    .map((x) => parseInt(x, 10))
    .filter((n) => Number.isFinite(n) && n >= 0 && n <= 365);
  if (!cleaned.length) return Array.isArray(fallback) ? fallback.slice() : [14];
  return Array.from(new Set(cleaned)).sort((a, b) => b - a).slice(0, 10);
}
