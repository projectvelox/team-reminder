// Tab-side telemetry sink. Accepts a single error event from the personal tab
// and logs it via context.error so App Insights ingests it through the Function
// App's existing instrumentation. Auth via Teams SSO; no PII beyond oid is logged.

const { app } = require('@azure/functions');
const { verifyTeamsToken } = require('../lib/auth');

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Content-Type': 'application/json',
  };
}

function clamp(s, n) {
  return String(s == null ? '' : s).slice(0, n);
}

app.http('telemetry', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'telemetry',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    // Auth is OPTIONAL: SSO boot failures can't carry a token, so we accept
    // anonymous error reports too. Authenticated reports get the oid; others
    // log "anon" — useful for spotting unauth'd boot failures in App Insights.
    let oid = 'anon';
    const auth = request.headers.get('authorization');
    if (auth) {
      try {
        const user = await verifyTeamsToken(auth);
        oid = user.oid;
      } catch (_) { /* keep anon */ }
    }
    const body = await request.json().catch(() => ({}));
    const event = clamp(body.event || 'tab.error', 64);
    const message = clamp(body.message, 1000);
    const stack = clamp(body.stack, 4000);
    const version = clamp(body.version, 32);
    const url = clamp(body.url, 256);
    const ua = clamp(request.headers.get('user-agent'), 256);
    context.error(`[tab-telemetry] oid=${oid} event=${event} version=${version} url=${url} ua=${ua} msg=${message} stack=${stack}`);
    return { status: 204, headers: corsHeaders() };
  },
});
