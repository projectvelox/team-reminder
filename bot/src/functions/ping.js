// Cheap unauthenticated keep-warm endpoint. The Logic App pings this every
// 5 minutes during PH work hours so the Function App's HTTP triggers stay
// warm and the tab doesn't pay a 5-15 s cold-start cost on first load.
const { app } = require('@azure/functions');

app.http('ping', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'ping',
  handler: async () => ({
    status: 200,
    headers: { 'Content-Type': 'text/plain', 'Cache-Control': 'no-store' },
    body: 'ok',
  }),
});
