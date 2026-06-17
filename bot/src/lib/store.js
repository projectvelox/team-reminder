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
  // License-tab settings (v1.7.22; v1.7.37 widens licenseLeadDays to an array)
  licenseLeadDays: [14],         // per-user default lead-day thresholds, used when license.leadDays is null/empty
  licenseSkipBriefing: false,    // opt out of the morning briefing card
  licenseSkipMonthlyDigest: false, // opt out of the monthly email digest
  licenseRollupDigest: false,    // opt in to include an all-accounts section in the digest (sales lead use case)
  // v1.7.39 — per-user saved filter views. Each: { id, name, filters: {...} }
  savedLicenseViews: [],
};

// ---------- user ----------

// Normalize licenseLeadDays after read so legacy rows that stored a scalar
// (pre-v1.7.37) come out as an array. Mutates and returns the settings object.
function migrateLeadDaysSetting(s) {
  if (typeof s.licenseLeadDays === 'number') {
    s.licenseLeadDays = [s.licenseLeadDays];
  } else if (!Array.isArray(s.licenseLeadDays)) {
    s.licenseLeadDays = [14];
  }
  return s;
}

async function getUser(oid) {
  await ensureTable();
  try {
    const entity = await getClient().getEntity(oid, '_user');
    const settings = entity.settings ? JSON.parse(entity.settings) : { ...DEFAULT_SETTINGS };
    migrateLeadDaysSetting(settings);
    return {
      settings,
      conversationRef: entity.conversationRef ? JSON.parse(entity.conversationRef) : null,
      lastEodDate: entity.lastEodDate || null,
      lastRolloverDate: entity.lastRolloverDate || null,
      tenantId: entity.tenantId || null,
      serviceUrl: entity.serviceUrl || null,
      displayName: entity.displayName || null,
      lastBriefingDate: entity.lastBriefingDate || null,
      briefingSnoozedUntil: entity.briefingSnoozedUntil || null,
      coldNudgedAt: entity.coldNudgedAt || null,
      lastDigestSentMonth: entity.lastDigestSentMonth || null,
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
        coldNudgedAt: null,
        lastDigestSentMonth: null,
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
    coldNudgedAt: merged.coldNudgedAt || null,
    lastDigestSentMonth: merged.lastDigestSentMonth || null,
  }, 'Replace');
  return merged;
}

// Iterate every user record (with or without a conversationRef). Used by the
// monthly digest which mails owners regardless of whether they've opened the bot.
async function* iterateAllUsers() {
  await ensureTable();
  const iter = getClient().listEntities({
    queryOptions: { filter: "RowKey eq '_user'" },
  });
  for await (const e of iter) {
    yield {
      oid: e.partitionKey,
      settings: e.settings ? JSON.parse(e.settings) : { ...DEFAULT_SETTINGS },
      conversationRef: e.conversationRef ? JSON.parse(e.conversationRef) : null,
      lastEodDate: e.lastEodDate || null,
      lastRolloverDate: e.lastRolloverDate || null,
      tenantId: e.tenantId || null,
      serviceUrl: e.serviceUrl || null,
      displayName: e.displayName || null,
      lastBriefingDate: e.lastBriefingDate || null,
      briefingSnoozedUntil: e.briefingSnoozedUntil || null,
      coldNudgedAt: e.coldNudgedAt || null,
      lastDigestSentMonth: e.lastDigestSentMonth || null,
    };
  }
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

// Parse a leadDays cell from Tables. Returns number[] | null.
// Backward compat: pre-v1.7.37 rows stored a scalar number; convert to [n].
function parseLeadDays(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return Number.isFinite(v) ? [v] : null;
  if (typeof v === 'string') {
    try {
      const parsed = JSON.parse(v);
      if (Array.isArray(parsed)) {
        const cleaned = parsed.map(Number).filter((n) => Number.isFinite(n) && n >= 0 && n <= 365);
        return cleaned.length ? cleaned : null;
      }
      if (typeof parsed === 'number' && Number.isFinite(parsed)) return [parsed];
    } catch { /* fall through */ }
  }
  return null;
}

// Always store as a JSON string so we don't depend on Azure Tables' EDM
// inference of an untyped column. Null becomes null (cleared cell).
function serializeLeadDays(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number' && Number.isFinite(v)) return JSON.stringify([Math.floor(v)]);
  if (Array.isArray(v)) {
    const cleaned = v.map(Number).filter((n) => Number.isFinite(n) && n >= 0 && n <= 365).map(Math.floor);
    if (!cleaned.length) return null;
    return JSON.stringify(Array.from(new Set(cleaned)).sort((a, b) => b - a));
  }
  return null;
}

// v1.7.39 — per-license comment thread. v1.7.40 tightened cap to 30 so the
// JSON-encoded property stays under Azure Tables' 32KB string limit.
// Each: { id, at, byOid, byName, text }.
function parseComments(v) {
  if (!v || typeof v !== 'string') return [];
  try {
    const parsed = JSON.parse(v);
    if (!Array.isArray(parsed)) return [];
    return parsed.slice(-30);
  } catch { return []; }
}

function parseLeadFireLog(v) {
  if (!v) return [];
  if (typeof v !== 'string') return [];
  try {
    const parsed = JSON.parse(v);
    if (!Array.isArray(parsed)) return [];
    return parsed.map(Number).filter((n) => Number.isFinite(n));
  } catch { return []; }
}

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
    leadDays: parseLeadDays(e.leadDays),
    lastFiredLeadDays: parseLeadFireLog(e.lastFiredLeadDays),
    notes: e.notes || null,
    comments: parseComments(e.comments),
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
    leadSnoozedUntil: e.leadSnoozedUntil || null,
    // v1.7.43 — soft-delete. Rows with deletedAt set are hidden from default
    // queries; a scheduled task hard-deletes after 30 days. Restoring within
    // the window clears the field.
    deletedAt: e.deletedAt || null,
    deletedByOid: e.deletedByOid || null,
    deletedByName: e.deletedByName || null,
    events,
  };
}

// v1.7.43 — default queries skip soft-deleted rows. Pass `includeDeleted` to
// see them (the recovery view + the 30-day purge timer use this).
async function listLicenses(opts) {
  await ensureTable();
  const iter = getClient().listEntities({
    queryOptions: { filter: `PartitionKey eq '${LICENSE_PARTITION}'` },
  });
  const out = [];
  for await (const e of iter) {
    const lic = entityToLicense(e);
    if (lic.deletedAt && !(opts && opts.includeDeleted)) continue;
    out.push(lic);
  }
  return out;
}

// Same as listLicenses({ includeDeleted: true }), filtered to soft-deleted only.
async function listDeletedLicenses() {
  const all = await listLicenses({ includeDeleted: true });
  return all.filter((l) => l.deletedAt);
}

async function getLicense(id, opts) {
  await ensureTable();
  try {
    const e = await getClient().getEntity(LICENSE_PARTITION, `l:${id}`);
    const lic = entityToLicense({ ...e, rowKey: `l:${id}` });
    if (lic.deletedAt && !(opts && opts.includeDeleted)) return null;
    return lic;
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
    leadDays: serializeLeadDays(license.leadDays),
    lastFiredLeadDays: JSON.stringify(Array.isArray(license.lastFiredLeadDays) ? license.lastFiredLeadDays : []),
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
    leadSnoozedUntil: license.leadSnoozedUntil || null,
    deletedAt: license.deletedAt || null,
    deletedByOid: license.deletedByOid || null,
    deletedByName: license.deletedByName || null,
    events: JSON.stringify(Array.isArray(license.events) ? license.events.slice(-50) : []),
    comments: JSON.stringify(Array.isArray(license.comments) ? license.comments.slice(-30) : []),
  }, 'Replace');
}

// v1.7.43 — soft-delete: stamp deletedAt + actor and upsert. Row stays in
// storage; default queries skip it; the scheduled purge hard-deletes after
// 30 days. Returns the soft-deleted license (or null if not found).
async function deleteLicense(id, actor) {
  const lic = await getLicense(id);
  if (!lic) return null;
  lic.deletedAt = new Date().toISOString();
  lic.deletedByOid = (actor && actor.oid) || null;
  lic.deletedByName = (actor && actor.name) || null;
  await upsertLicense(lic);
  return lic;
}

// Permanent removal — used by the purge timer for rows soft-deleted > 30 days
// ago, and by admin/restore endpoints if we ever expose them. NOT called from
// the HTTP DELETE handler.
async function hardDeleteLicense(id) {
  await ensureTable();
  try {
    await getClient().deleteEntity(LICENSE_PARTITION, `l:${id}`);
    return true;
  } catch (err) {
    if (err.statusCode === 404) return false;
    throw err;
  }
}

// Restore a soft-deleted row by clearing deletedAt. Returns the restored
// license, or null if the id wasn't soft-deleted (so we don't accidentally
// "restore" a live row).
async function restoreLicense(id) {
  const lic = await getLicense(id, { includeDeleted: true });
  if (!lic || !lic.deletedAt) return null;
  lic.deletedAt = null;
  lic.deletedByOid = null;
  lic.deletedByName = null;
  await upsertLicense(lic);
  return lic;
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

// ---------- product lines registry (tenant-shared, v1.8.0) ----------
//
// Strict controlled vocab for License.productLine. Free-text product line is
// gone; the Edit dialog renders a <select> sourced from this registry. The
// canonical seed below ships on first deploy if the partition is empty;
// after that, admins can rename / add / delete via /api/product-lines.
//
// Legacy values (rows with productLine not in the registry) are tolerated
// at read time and rendered with a "(legacy)" badge in the dialog so they
// can be edited without forcing a normalize first.

const PRODUCT_LINE_PARTITION = '_productLines';
const CANONICAL_PRODUCT_LINES = [
  'M365',
  'Business Central',
  'Finance and Operation',
  'PHILTAX',
  'CRM',
  'Security',
];

function productLineId(name) {
  // RK stays stable across renames? No — we key by the canonical name so a
  // rename means delete+create. Simpler than tracking a separate id.
  return `p:${String(name || '').trim().slice(0, 100)}`;
}

function entityToProductLine(e) {
  return {
    name: e.rowKey.slice(2),
    sortOrder: typeof e.sortOrder === 'number' ? e.sortOrder : 100,
    createdAt: e.createdAt || null,
    createdByOid: e.createdByOid || null,
    createdByName: e.createdByName || null,
  };
}

async function listProductLines() {
  await ensureTable();
  const iter = getClient().listEntities({
    queryOptions: { filter: `PartitionKey eq '${PRODUCT_LINE_PARTITION}'` },
  });
  const out = [];
  for await (const e of iter) out.push(entityToProductLine(e));
  out.sort((a, b) => (a.sortOrder - b.sortOrder) || a.name.localeCompare(b.name));
  return out;
}

async function upsertProductLine(pl) {
  await ensureTable();
  const name = String(pl.name || '').trim().slice(0, 100);
  if (!name) throw new Error('product line name required');
  await getClient().upsertEntity({
    partitionKey: PRODUCT_LINE_PARTITION,
    rowKey: productLineId(name),
    sortOrder: typeof pl.sortOrder === 'number' ? pl.sortOrder : 100,
    createdAt: pl.createdAt || new Date().toISOString(),
    createdByOid: pl.createdByOid || null,
    createdByName: pl.createdByName || null,
  }, 'Replace');
  return { name, sortOrder: pl.sortOrder ?? 100 };
}

async function deleteProductLine(name) {
  await ensureTable();
  try {
    await getClient().deleteEntity(PRODUCT_LINE_PARTITION, productLineId(name));
    return true;
  } catch (err) {
    if (err.statusCode === 404) return false;
    throw err;
  }
}

async function ensureProductLinesSeeded(actor) {
  const existing = await listProductLines();
  if (existing.length > 0) return existing;
  const now = new Date().toISOString();
  const seeded = [];
  for (let i = 0; i < CANONICAL_PRODUCT_LINES.length; i++) {
    const name = CANONICAL_PRODUCT_LINES[i];
    await upsertProductLine({
      name,
      sortOrder: i,
      createdAt: now,
      createdByOid: (actor && actor.oid) || null,
      createdByName: (actor && actor.name) || 'system seed',
    });
    seeded.push({ name, sortOrder: i });
  }
  return seeded;
}

module.exports = {
  DEFAULT_SETTINGS,
  LICENSE_STATUSES,
  RENEWAL_CYCLES,
  getUser,
  upsertUser,
  iterateUsersWithConversationRefs,
  iterateAllUsers,
  listReminders,
  getReminder,
  upsertReminder,
  deleteReminder,
  getTemplates,
  setTemplates,
  todayKey,
  listLicenses,
  listDeletedLicenses,
  getLicense,
  upsertLicense,
  deleteLicense,
  hardDeleteLicense,
  restoreLicense,
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
  CANONICAL_PRODUCT_LINES,
  listProductLines,
  upsertProductLine,
  deleteProductLine,
  ensureProductLinesSeeded,
};
