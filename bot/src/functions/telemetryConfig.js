// v1.8.3 — exposes the App Insights connection string to the tab JS so the
// client-side SDK can ingest usage telemetry (Users, Sessions, custom events).
//
// The connection string is not a secret per Microsoft docs — it identifies
// the AI resource and grants write-only access to its ingestion endpoint.
// Anyone with it can SEND telemetry; they cannot read what's stored. Keeping
// it server-side anyway so it doesn't get committed to the public repo.
//
// Anonymous on purpose: the tab needs this before SSO completes so we can
// also capture the boot-failure population (SSO errors, network drops).

const { app } = require('@azure/functions');

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': 'https://projectvelox.github.io',
    'Vary': 'Origin',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Max-Age': '600',
    'Content-Type': 'application/json',
    'X-Content-Type-Options': 'nosniff',
    'Referrer-Policy': 'no-referrer',
    // 5-minute cache is fine; connection string only changes on AI resource swap.
    'Cache-Control': 'public, max-age=300',
  };
}

app.http('telemetryConfig', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous',
  route: 'telemetry-config',
  handler: async (request) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: corsHeaders() };
    const connectionString = process.env.APPLICATIONINSIGHTS_CONNECTION_STRING || '';
    return {
      status: 200,
      headers: corsHeaders(),
      body: JSON.stringify({ connectionString }),
    };
  },
});
