// Azure Tables storage layer.
//
// Single table `dayreminders`. Partition strategy:
//   PK = <user oid>                 → per-user reminders
//     RK = _user                    → user metadata (settings, conversation ref, last EOD)
//     RK = _templates               → saved templates
//     RK = r:<reminderId>           → a reminder
//   PK = _licenses                  → tenant-shared license tracker (v1.7)
//     RK = l:<licenseId>            → a license
//   PK = _members                   → tenant-shared member registry (v1.7, Owner picker source)
//     RK = m:<oid>                  → a known user (auto-registered on tab load)

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
  quietStart: null, // HH:MM, null/empty = quiet hours disabled
  quietEnd: null,
  autoImportFlagged: false, // when true, flagged Outlook emails auto-create reminders. v1.5: setting persisted but subscription flow not yet active.
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
      lastRolloverDate: entity.lastRolloverDate || null,
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
        lastRolloverDate: null,
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
    lastRolloverDate: merged.lastRolloverDate || null,
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
      lastRolloverDate: e.lastRolloverDate || null,
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
  const repeat = e.repeat === 'daily' || e.repeat === 'weekdays' || e.repeat === 'weekly' ? e.repeat : 'none';
  let subtasks = [];
  if (e.subtasks) {
    try {
      const parsed = JSON.parse(e.subtasks);
      if (Array.isArray(parsed)) {
        subtasks = parsed
          .filter((s) => s && typeof s === 'object' && typeof s.text === 'string')
          .map((s) => ({
            id: typeof s.id === 'string' && s.id ? s.id : `s-${Math.random().toString(36).slice(2, 10)}`,
            text: String(s.text).slice(0, 500),
            done: !!s.done,
          }));
      }
    } catch { subtasks = []; }
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
    order: typeof e.order === 'number' ? e.order : null,
    leadMinutes: typeof e.leadMinutes === 'number' ? e.leadMinutes : null,
    snoozedUntil: e.snoozedUntil || null,
    dueAt: e.dueAt || e.createdDate || null,
    description: e.description || null,
    rollDays: typeof e.rollDays === 'number' ? e.rollDays : 0,
    client: e.client || null,
    repeat,
    subtasks,
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
    order: typeof r.order === 'number' ? r.order : null,
    leadMinutes: typeof r.leadMinutes === 'number' ? r.leadMinutes : null,
    snoozedUntil: r.snoozedUntil || null,
    dueAt: r.dueAt || null,
    description: r.description || null,
    rollDays: typeof r.rollDays === 'number' ? r.rollDays : 0,
    client: r.client || null,
    repeat: r.repeat === 'daily' || r.repeat === 'weekdays' || r.repeat === 'weekly' ? r.repeat : 'none',
    subtasks: JSON.stringify(Array.isArray(r.subtasks) ? r.subtasks.slice(0, 50) : []),
  }, 'Replace');
}

// ---------- templates (per-user, stored as a single JSON blob) ----------

async function getTemplates(oid) {
  await ensureTable();
  try {
    const entity = await getClient().getEntity(oid, '_templates');
    if (entity.items) {
      try {
        const parsed = JSON.parse(entity.items);
        return Array.isArray(parsed) ? parsed : [];
      } catch { return []; }
    }
    return [];
  } catch (err) {
    if (err.statusCode === 404) return [];
    throw err;
  }
}

async function setTemplates(oid, templates) {
  await ensureTable();
  await getClient().upsertEntity({
    partitionKey: oid,
    rowKey: '_templates',
    items: JSON.stringify(Array.isArray(templates) ? templates : []),
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

const PH_OFFSET_MS = 8 * 60 * 60 * 1000;

function todayKey() {
  const ph = new Date(Date.now() + PH_OFFSET_MS);
  return `${ph.getUTCFullYear()}-${String(ph.getUTCMonth() + 1).padStart(2, '0')}-${String(ph.getUTCDate()).padStart(2, '0')}`;
}

// ---------- licenses (tenant-shared, v1.7) ----------

const LICENSE_PARTITION = '_licenses';

function entityToLicense(e) {
  return {
    id: e.rowKey.slice(2),
    customer: e.customer || '',
    licenseType: e.licenseType || '',
    userCount: typeof e.userCount === 'number' ? e.userCount : (parseInt(e.userCount, 10) || 0),
    expiryDate: e.expiryDate || null,
    ownerOid: e.ownerOid || null,
    ownerName: e.ownerName || null,
    productLine: e.productLine || null,
    leadDays: typeof e.leadDays === 'number' ? e.leadDays : null,
    notes: e.notes || null,
    state: e.state === 'abandoned' ? 'abandoned' : 'active',
    createdAt: e.createdAt || null,
    createdByOid: e.createdByOid || null,
    createdByName: e.createdByName || null,
    lastEditedAt: e.lastEditedAt || null,
    lastEditedByOid: e.lastEditedByOid || null,
    lastEditedByName: e.lastEditedByName || null,
    lastRenewedAt: e.lastRenewedAt || null,
    lastEscalatedDays: typeof e.lastEscalatedDays === 'number' ? e.lastEscalatedDays : null,
  };
}

async function listLicenses() {
  await ensureTable();
  const iter = getClient().listEntities({
    queryOptions: { filter: `PartitionKey eq '${LICENSE_PARTITION}'` },
  });
  const out = [];
  for await (const e of iter) out.push(entityToLicense(e));
  return out;
}

async function getLicense(id) {
  await ensureTable();
  try {
    const e = await getClient().getEntity(LICENSE_PARTITION, `l:${id}`);
    return entityToLicense({ ...e, rowKey: `l:${id}` });
  } catch (err) {
    if (err.statusCode === 404) return null;
    throw err;
  }
}

async function upsertLicense(license) {
  await ensureTable();
  await getClient().upsertEntity({
    partitionKey: LICENSE_PARTITION,
    rowKey: `l:${license.id}`,
    customer: license.customer || '',
    licenseType: license.licenseType || '',
    userCount: typeof license.userCount === 'number' ? license.userCount : 0,
    expiryDate: license.expiryDate || null,
    ownerOid: license.ownerOid || null,
    ownerName: license.ownerName || null,
    productLine: license.productLine || null,
    leadDays: typeof license.leadDays === 'number' ? license.leadDays : null,
    notes: license.notes || null,
    state: license.state === 'abandoned' ? 'abandoned' : 'active',
    createdAt: license.createdAt || null,
    createdByOid: license.createdByOid || null,
    createdByName: license.createdByName || null,
    lastEditedAt: license.lastEditedAt || null,
    lastEditedByOid: license.lastEditedByOid || null,
    lastEditedByName: license.lastEditedByName || null,
    lastRenewedAt: license.lastRenewedAt || null,
    lastEscalatedDays: typeof license.lastEscalatedDays === 'number' ? license.lastEscalatedDays : null,
  }, 'Replace');
}

async function deleteLicense(id) {
  await ensureTable();
  try {
    await getClient().deleteEntity(LICENSE_PARTITION, `l:${id}`);
    return true;
  } catch (err) {
    if (err.statusCode === 404) return false;
    throw err;
  }
}

// ---------- members (tenant-shared, v1.7) ----------
// Auto-populating registry of users who have opened the app.
// Source of truth for the Owner picker on the Licenses tab.

const MEMBER_PARTITION = '_members';

function entityToMember(e) {
  return {
    oid: e.rowKey.slice(2),
    displayName: e.displayName || null,
    upn: e.upn || null,
    firstSeenAt: e.firstSeenAt || null,
    lastSeenAt: e.lastSeenAt || null,
  };
}

async function registerMember({ oid, displayName, upn }) {
  if (!oid) return null;
  await ensureTable();
  const now = new Date().toISOString();
  let firstSeenAt = now;
  try {
    const existing = await getClient().getEntity(MEMBER_PARTITION, `m:${oid}`);
    firstSeenAt = existing.firstSeenAt || now;
  } catch (err) {
    if (err.statusCode !== 404) throw err;
  }
  await getClient().upsertEntity({
    partitionKey: MEMBER_PARTITION,
    rowKey: `m:${oid}`,
    displayName: displayName || null,
    upn: upn || null,
    firstSeenAt,
    lastSeenAt: now,
  }, 'Replace');
  return { oid, displayName, upn, firstSeenAt, lastSeenAt: now };
}

async function listMembers() {
  await ensureTable();
  const iter = getClient().listEntities({
    queryOptions: { filter: `PartitionKey eq '${MEMBER_PARTITION}'` },
  });
  const out = [];
  for await (const e of iter) out.push(entityToMember(e));
  return out;
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
  getTemplates,
  setTemplates,
  todayKey,
  listLicenses,
  getLicense,
  upsertLicense,
  deleteLicense,
  registerMember,
  listMembers,
};
