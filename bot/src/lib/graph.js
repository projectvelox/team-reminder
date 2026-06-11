// App-only Microsoft Graph client for Teams activity feed notifications.
//
// Sends a notification to the user's Activity tab when a reminder fires, so
// even if the user has the bot chat collapsed, the bell icon lights up.
//
// Requires:
//   - Graph application permission `TeamsActivity.Send` (granted to the bot's
//     Entra app registration, with tenant-wide admin consent)
//   - Env var `TeamsAppId` set on the Function App to the manifest's `id`
//   - Env vars MicrosoftAppId / MicrosoftAppPassword / MicrosoftAppTenantId
//     (already used by the Bot Framework adapter; reused here for the
//     client credentials token request)

const TENANT_ID = process.env.MicrosoftAppTenantId;
const APP_ID = process.env.MicrosoftAppId;
const APP_SECRET = process.env.MicrosoftAppPassword;
const TEAMS_APP_ID = process.env.TeamsAppId;

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

let tokenCache = null; // { token, expiresAt }
const installCache = new Map(); // oid -> { id, fetchedAt }
const INSTALL_TTL_MS = 60 * 60 * 1000;

async function getAppToken() {
  if (tokenCache && Date.now() < tokenCache.expiresAt - 60_000) return tokenCache.token;
  if (!TENANT_ID || !APP_ID || !APP_SECRET) {
    throw new Error('Graph credentials missing (MicrosoftAppTenantId / MicrosoftAppId / MicrosoftAppPassword)');
  }
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: APP_ID,
    client_secret: APP_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => '');
    throw new Error(`token fetch ${res.status}: ${txt.slice(0, 300)}`);
  }
  const data = await res.json();
  tokenCache = {
    token: data.access_token,
    expiresAt: Date.now() + (data.expires_in || 3600) * 1000,
  };
  return tokenCache.token;
}

async function graph(method, path, body) {
  const token = await getAppToken();
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => '');
    const err = new Error(`Graph ${method} ${path} -> ${res.status}: ${txt.slice(0, 300)}`);
    err.status = res.status;
    throw err;
  }
  if (res.status === 204) return null;
  return res.json();
}

async function getInstallationId(oid) {
  const cached = installCache.get(oid);
  if (cached && Date.now() - cached.fetchedAt < INSTALL_TTL_MS) return cached.id;
  if (!TEAMS_APP_ID) throw new Error('TeamsAppId env var not set');
  const filter = `teamsApp/id eq '${TEAMS_APP_ID}'`;
  const data = await graph(
    'GET',
    `/users/${oid}/teamwork/installedApps?$expand=teamsApp&$filter=${encodeURIComponent(filter)}`
  );
  const inst = (data && data.value && data.value[0]) || null;
  if (!inst) return null;
  installCache.set(oid, { id: inst.id, fetchedAt: Date.now() });
  return inst.id;
}

async function sendReminderActivity(oid, reminder) {
  const installId = await getInstallationId(oid);
  if (!installId) return false;
  await graph('POST', `/users/${oid}/teamwork/sendActivityNotification`, {
    topic: {
      source: 'entityUrl',
      value: `https://graph.microsoft.com/v1.0/users/${oid}/teamwork/installedApps/${installId}`,
    },
    activityType: 'reminderFired',
    previewText: {
      content: reminder.time ? `${reminder.title} at ${reminder.time}` : reminder.title,
    },
    templateParameters: [
      { name: 'reminderTitle', value: reminder.title || 'Reminder' },
    ],
  });
  return true;
}

module.exports = { sendReminderActivity };
