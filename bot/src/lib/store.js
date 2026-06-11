// Azure Tables storage layer.
//
// Single table `dayreminders`, partitioned by user oid (AAD object ID).
// RowKey encodes the kind:
//   _user                          → user metadata (settings, conversation ref, last EOD)
//   r:<reminderId>                 → a reminder

const { TableClient, AzureNamedKeyCredential } = require('@azure/data-tables');

const TABLE_NAME = 'dayreminders';

let cachedClient;

function getClient() {
  if (cachedClient) return cachedClient;
  const conn = process.env.DAYREMINDERS_STORAGE || process.env.AzureWebJobsStorage;
  if (!conn) throw new Error('DAYREMINDERS_STORAGE or AzureWebJobsStorage must be set');
  cachedClient = TableClient.fromConnectionString(conn, TABLE_NAME, { allowInsecureConnection: false });
  return cachedClient;
}

async function ensureTable() {
  await getClient().createTable().catch((err) => {
    if (err.statusCode !== 409) throw err;
  });
}

const DEFAULT_SETTINGS = {
  eodTime: '17:00',
  leadMinutes: 10,
  weekdaysOnly: true,
  notifications: true,
};

// ---------- user ----------

async function getUser(oid) {
  await ensureTable();
  try {
    const entity = await getClient().getEntity(oid, '_user');
    return {
      settings: entity.settings ? JSON.parse(entity.settings) : { ...DEFAULT_SETTINGS },
      conversationRef: entity.conversationRef ? JSON.parse(entity.conversationRef) : null,
      lastEodDate: entity.lastEodDate || null,
      tenantId: entity.tenantId || null,
      serviceUrl: entity.serviceUrl || null,
      displayName: entity.displayName || null,
    };
  } catch (err) {
    if (err.statusCode === 404) {
      return {
        settings: { ...DEFAULT_SETTINGS },
        conversationRef: null,
        lastEodDate: null,
        tenantId: null,
        serviceUrl: null,
        displayName: null,
      };
    }
    throw err;
  }
}

async function upsertUser(oid, patch) {
  await ensureTable();
  const existing = await getUser(oid);
  const merged = { ...existing, ...patch };
  await getClient().upsertEntity({
    partitionKey: oid,
    rowKey: '_user',
    settings: JSON.stringify(merged.settings || DEFAULT_SETTINGS),
    conversationRef: merged.conversationRef ? JSON.stringify(merged.conversationRef) : null,
    lastEodDate: merged.lastEodDate || null,
    tenantId: merged.tenantId || null,
    serviceUrl: merged.serviceUrl || null,
    displayName: merged.displayName || null,
  }, 'Replace');
  return merged;
}

async function* iterateUsersWithConversationRefs() {
  await ensureTable();
  const iter = getClient().listEntities({
    queryOptions: { filter: "RowKey eq '_user'" },
  });
  for await (const e of iter) {
    if (!e.conversationRef) continue;
    yield {
      oid: e.partitionKey,
      settings: e.settings ? JSON.parse(e.settings) : { ...DEFAULT_SETTINGS },
      conversationRef: JSON.parse(e.conversationRef),
      lastEodDate: e.lastEodDate || null,
      tenantId: e.tenantId || null,
      serviceUrl: e.serviceUrl || null,
    };
  }
}

// ---------- reminders ----------

function entityToReminder(e) {
  let tags = [];
  if (e.tags) {
    try { tags = JSON.parse(e.tags); } catch { tags = []; }
    if (!Array.isArray(tags)) tags = [];
  }
  return {
    id: e.rowKey.slice(2),
    title: e.title,
    time: e.time || null,
    done: !!e.done,
    firedAt: e.firedAt || null,
    createdDate: e.createdDate || null,
    closedAt: e.closedAt || null,
    tags,
    priority: e.priority === 'high' ? 'high' : 'normal',
  };
}

async function listReminders(oid) {
  await ensureTable();
  const iter = getClient().listEntities({
    queryOptions: { filter: `PartitionKey eq '${oid}' and RowKey ge 'r:' and RowKey lt 'r;'` },
  });
  const out = [];
  for await (const e of iter) out.push(entityToReminder(e));
  return out;
}

async function getReminder(oid, id) {
  await ensureTable();
  try {
    const e = await getClient().getEntity(oid, `r:${id}`);
    return entityToReminder({ ...e, rowKey: `r:${id}` });
  } catch (err) {
    if (err.statusCode === 404) return null;
    throw err;
  }
}

async function upsertReminder(oid, r) {
  await ensureTable();
  await getClient().upsertEntity({
    partitionKey: oid,
    rowKey: `r:${r.id}`,
    title: r.title,
    time: r.time || null,
    done: !!r.done,
    firedAt: r.firedAt || null,
    createdDate: r.createdDate || todayKey(),
    closedAt: r.closedAt || null,
    tags: JSON.stringify(Array.isArray(r.tags) ? r.tags : []),
    priority: r.priority === 'high' ? 'high' : 'normal',
  }, 'Replace');
}

async function deleteReminder(oid, id) {
  await ensureTable();
  try {
    await getClient().deleteEntity(oid, `r:${id}`);
    return true;
  } catch (err) {
    if (err.statusCode === 404) return false;
    throw err;
  }
}

function todayKey(d = new Date()) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

module.exports = {
  DEFAULT_SETTINGS,
  getUser,
  upsertUser,
  iterateUsersWithConversationRefs,
  listReminders,
  getReminder,
  upsertReminder,
  deleteReminder,
  todayKey,
};
