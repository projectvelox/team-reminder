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
      lastBriefingDate: entity.lastBriefingDate || null,
      briefingSnoozedUntil: entity.briefingSnoozedUntil || null,
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
        lastBriefingDate: null,
        briefingSnoozedUntil: null,
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
    lastBriefingDate: merged.lastBriefingDate || null,
    briefingSnoozedUntil: merged.briefingSnoozedUntil || null,
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

const LICENSE_STATUSES = ['notStarted', 'noticeSent', 'awaitingCustomer', 'customerConfirmed', 'renewed'];
const RENEWAL_CYCLES = ['annual', 'biennial', 'triennial'];

function entityToLicense(e) {
  let events = [];
  if (e.events) {
    try { const parsed = JSON.parse(e.events); if (Array.isArray(parsed)) events = parsed.slice(-50); }
    catch { events = []; }
  }
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
    status: LICENSE_STATUSES.includes(e.status) ? e.status : 'notStarted',
    statusChangedAt: e.statusChangedAt || null,
    statusChangedByOid: e.statusChangedByOid || null,
    statusChangedByName: e.statusChangedByName || null,
    renewalCycle: RENEWAL_CYCLES.includes(e.renewalCycle) ? e.renewalCycle : 'annual',
    createdAt: e.createdAt || null,
    createdByOid: e.createdByOid || null,
    createdByName: e.createdByName || null,
    lastEditedAt: e.lastEditedAt || null,
    lastEditedByOid: e.lastEditedByOid || null,
    lastEditedByName: e.lastEditedByName || null,
    lastRenewedAt: e.lastRenewedAt || null,
    lastFollowUpAt: e.lastFollowUpAt || null,
    lastEscalatedDays: typeof e.lastEscalatedDays === 'number' ? e.lastEscalatedDays : null,
    events,
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
    status: LICENSE_STATUSES.includes(license.status) ? license.status : 'notStarted',
    statusChangedAt: license.statusChangedAt || null,
    statusChangedByOid: license.statusChangedByOid || null,
    statusChangedByName: license.statusChangedByName || null,
    renewalCycle: RENEWAL_CYCLES.includes(license.renewalCycle) ? license.renewalCycle : 'annual',
    createdAt: license.createdAt || null,
    createdByOid: license.createdByOid || null,
    createdByName: license.createdByName || null,
    lastEditedAt: license.lastEditedAt || null,
    lastEditedByOid: license.lastEditedByOid || null,
    lastEditedByName: license.lastEditedByName || null,
    lastRenewedAt: license.lastRenewedAt || null,
    lastFollowUpAt: license.lastFollowUpAt || null,
    lastEscalatedDays: typeof license.lastEscalatedDays === 'number' ? license.lastEscalatedDays : null,
    events: JSON.stringify(Array.isArray(license.events) ? license.events.slice(-50) : []),
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

// ---------- customers (tenant-shared, v1.7.9) ----------
// Annotation layer over the per-license `customer` string field. Adds contact
// emails, address, notes that apply across all that customer's licenses.

const CUSTOMER_PARTITION = '_customers';

function customerIdFromName(name) {
  return String(name || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 100) || 'unnamed';
}

function entityToCustomer(e) {
  let secondaryEmails = [];
  if (e.secondaryEmails) {
    try { const p = JSON.parse(e.secondaryEmails); if (Array.isArray(p)) secondaryEmails = p; } catch {}
  }
  return {
    id: e.rowKey.slice(2),
    name: e.name || '',
    primaryEmail: e.primaryEmail || null,
    secondaryEmails,
    address: e.address || null,
    notes: e.notes || null,
    createdAt: e.createdAt || null,
    updatedAt: e.updatedAt || null,
  };
}

async function listCustomers() {
  await ensureTable();
  const iter = getClient().listEntities({
    queryOptions: { filter: `PartitionKey eq '${CUSTOMER_PARTITION}'` },
  });
  const out = [];
  for await (const e of iter) out.push(entityToCustomer(e));
  return out;
}

async function getCustomer(id) {
  await ensureTable();
  try {
    const e = await getClient().getEntity(CUSTOMER_PARTITION, `c:${id}`);
    return entityToCustomer({ ...e, rowKey: `c:${id}` });
  } catch (err) {
    if (err.statusCode === 404) return null;
    throw err;
  }
}

async function upsertCustomer(customer) {
  await ensureTable();
  await getClient().upsertEntity({
    partitionKey: CUSTOMER_PARTITION,
    rowKey: `c:${customer.id}`,
    name: customer.name || '',
    primaryEmail: customer.primaryEmail || null,
    secondaryEmails: JSON.stringify(Array.isArray(customer.secondaryEmails) ? customer.secondaryEmails : []),
    address: customer.address || null,
    notes: customer.notes || null,
    createdAt: customer.createdAt || null,
    updatedAt: customer.updatedAt || null,
  }, 'Replace');
}

async function deleteCustomer(id) {
  await ensureTable();
  try {
    await getClient().deleteEntity(CUSTOMER_PARTITION, `c:${id}`);
    return true;
  } catch (err) {
    if (err.statusCode === 404) return false;
    throw err;
  }
}

// Ensure a stub customer entity exists for this name; idempotent. Called when
// a license is created or imported so the registry stays in sync with reality.
async function ensureCustomer(name) {
  if (!name) return null;
  const id = customerIdFromName(name);
  const existing = await getCustomer(id);
  if (existing) return existing;
  const now = new Date().toISOString();
  const stub = {
    id,
    name: name.trim(),
    primaryEmail: null,
    secondaryEmails: [],
    address: null,
    notes: null,
    createdAt: now,
    updatedAt: now,
  };
  await upsertCustomer(stub);
  return stub;
}

// ---------- email templates (tenant-shared, v1.7.9) ----------

const EMAIL_TEMPLATE_PARTITION = '_emailTemplates';

function entityToTemplate(e) {
  return {
    productLine: e.rowKey.slice(2),
    subject: e.subject || '',
    body: e.body || '',
    lastEditedAt: e.lastEditedAt || null,
    lastEditedByOid: e.lastEditedByOid || null,
    lastEditedByName: e.lastEditedByName || null,
  };
}

async function listEmailTemplates() {
  await ensureTable();
  const iter = getClient().listEntities({
    queryOptions: { filter: `PartitionKey eq '${EMAIL_TEMPLATE_PARTITION}'` },
  });
  const out = [];
  for await (const e of iter) out.push(entityToTemplate(e));
  return out;
}

async function upsertEmailTemplate(template) {
  await ensureTable();
  // Use a sentinel productLine "_default" for the general template.
  const key = String(template.productLine || '_default').trim().slice(0, 100) || '_default';
  await getClient().upsertEntity({
    partitionKey: EMAIL_TEMPLATE_PARTITION,
    rowKey: `t:${key}`,
    subject: template.subject || '',
    body: template.body || '',
    lastEditedAt: template.lastEditedAt || null,
    lastEditedByOid: template.lastEditedByOid || null,
    lastEditedByName: template.lastEditedByName || null,
  }, 'Replace');
}

async function deleteEmailTemplate(productLine) {
  await ensureTable();
  const key = String(productLine || '_default').trim().slice(0, 100) || '_default';
  try {
    await getClient().deleteEntity(EMAIL_TEMPLATE_PARTITION, `t:${key}`);
    return true;
  } catch (err) {
    if (err.statusCode === 404) return false;
    throw err;
  }
}

module.exports = {
  DEFAULT_SETTINGS,
  LICENSE_STATUSES,
  RENEWAL_CYCLES,
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
  listCustomers,
  getCustomer,
  upsertCustomer,
  deleteCustomer,
  ensureCustomer,
  customerIdFromName,
  listEmailTemplates,
  upsertEmailTemplate,
  deleteEmailTemplate,
};
