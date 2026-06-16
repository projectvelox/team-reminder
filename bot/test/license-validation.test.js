// Tests for the license payload validator. Because the validator is defined
// inside licenses.js (not exported), we re-create the same logic here as a
// regression contract: if behavior changes in the handler, this test must
// change too — surfaces the intent.

const test = require('node:test');
const assert = require('node:assert/strict');

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

const LICENSE_STATUSES = ['notStarted', 'noticeSent', 'awaitingCustomer', 'customerConfirmed', 'renewed'];
const RENEWAL_CYCLES = ['annual', 'biennial', 'triennial'];

function validatePayload(body, existing) {
  const out = existing ? { ...existing } : {};
  if (body.customer !== undefined) {
    const v = String(body.customer || '').trim();
    if (!v) return { error: 'customer is required' };
    if (v.length > 200) return { error: 'customer max 200 chars' };
    out.customer = v;
  } else if (!existing) return { error: 'customer is required' };

  if (body.licenseType !== undefined) {
    const v = String(body.licenseType || '').trim();
    if (!v) return { error: 'licenseType is required' };
    if (v.length > 200) return { error: 'licenseType max 200 chars' };
    out.licenseType = v;
  } else if (!existing) return { error: 'licenseType is required' };

  if (body.userCount !== undefined) {
    const n = Number(body.userCount);
    if (!isFinite(n) || n < 0 || n > 1000000) return { error: 'userCount must be a non-negative number' };
    out.userCount = Math.floor(n);
  } else if (!existing) out.userCount = 0;

  if (body.expiryDate !== undefined) {
    if (!ISO_DATE.test(String(body.expiryDate || ''))) return { error: 'expiryDate must be YYYY-MM-DD' };
    out.expiryDate = body.expiryDate;
  } else if (!existing) return { error: 'expiryDate is required' };

  if (body.ownerOid !== undefined) {
    const oid = String(body.ownerOid || '').trim();
    if (!oid) return { error: 'ownerOid is required' };
    out.ownerOid = oid;
  } else if (!existing) return { error: 'ownerOid is required' };

  if (body.leadDays === null) out.leadDays = null;
  else if (body.leadDays !== undefined) {
    let arr = body.leadDays;
    if (typeof arr === 'number') arr = [arr];
    if (!Array.isArray(arr)) return { error: 'leadDays must be an array of 0-365 ints' };
    const cleaned = arr.map((x) => Number(x)).filter((n) => Number.isFinite(n) && n >= 0 && n <= 365).map((n) => Math.floor(n));
    if (cleaned.length > 10) return { error: 'leadDays max 10 thresholds' };
    out.leadDays = cleaned.length ? Array.from(new Set(cleaned)).sort((a, b) => b - a) : null;
  }

  if (body.status !== undefined) {
    if (!LICENSE_STATUSES.includes(body.status)) return { error: `status must be one of ${LICENSE_STATUSES.join(', ')}` };
    out.status = body.status;
  } else if (!existing) out.status = 'notStarted';

  if (body.renewalCycle !== undefined) {
    if (!RENEWAL_CYCLES.includes(body.renewalCycle)) return { error: `renewalCycle must be one of ${RENEWAL_CYCLES.join(', ')}` };
    out.renewalCycle = body.renewalCycle;
  } else if (!existing) out.renewalCycle = 'annual';

  return { license: out };
}

// ---- happy path ----

test('validatePayload: create with all required fields succeeds', () => {
  const r = validatePayload({
    customer: 'Acme',
    licenseType: 'M365 BS',
    expiryDate: '2026-08-15',
    ownerOid: 'oid-1',
  }, null);
  assert.ok(r.license);
  assert.equal(r.license.customer, 'Acme');
  assert.equal(r.license.userCount, 0);
  assert.equal(r.license.status, 'notStarted');
  assert.equal(r.license.renewalCycle, 'annual');
});

// ---- required-field guards ----

test('validatePayload: missing customer on create returns error', () => {
  const r = validatePayload({ licenseType: 'X', expiryDate: '2026-08-15', ownerOid: 'o' }, null);
  assert.equal(r.error, 'customer is required');
});

test('validatePayload: missing licenseType on create returns error', () => {
  const r = validatePayload({ customer: 'X', expiryDate: '2026-08-15', ownerOid: 'o' }, null);
  assert.equal(r.error, 'licenseType is required');
});

// ---- length caps ----

test('validatePayload: customer > 200 chars rejected', () => {
  const r = validatePayload({
    customer: 'a'.repeat(201),
    licenseType: 'X',
    expiryDate: '2026-08-15',
    ownerOid: 'o',
  }, null);
  assert.equal(r.error, 'customer max 200 chars');
});

// ---- expiry date shape ----

test('validatePayload: non-ISO expiryDate rejected', () => {
  const r = validatePayload({
    customer: 'X', licenseType: 'X', expiryDate: '08/15/2026', ownerOid: 'o',
  }, null);
  assert.equal(r.error, 'expiryDate must be YYYY-MM-DD');
});

// ---- leadDays ----

test('validatePayload: leadDays array dedupes + sorts desc', () => {
  const r = validatePayload({
    customer: 'X', licenseType: 'X', expiryDate: '2026-08-15', ownerOid: 'o',
    leadDays: [7, 30, 60, 30, 7, 1],
  }, null);
  assert.deepEqual(r.license.leadDays, [60, 30, 7, 1]);
});

test('validatePayload: leadDays > 10 entries rejected', () => {
  const r = validatePayload({
    customer: 'X', licenseType: 'X', expiryDate: '2026-08-15', ownerOid: 'o',
    leadDays: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11],
  }, null);
  assert.equal(r.error, 'leadDays max 10 thresholds');
});

test('validatePayload: leadDays scalar coerced to array', () => {
  const r = validatePayload({
    customer: 'X', licenseType: 'X', expiryDate: '2026-08-15', ownerOid: 'o',
    leadDays: 14,
  }, null);
  assert.deepEqual(r.license.leadDays, [14]);
});

test('validatePayload: leadDays out-of-range silently filtered', () => {
  const r = validatePayload({
    customer: 'X', licenseType: 'X', expiryDate: '2026-08-15', ownerOid: 'o',
    leadDays: [999, 30, -5],
  }, null);
  assert.deepEqual(r.license.leadDays, [30]);
});

// ---- status / renewalCycle enums ----

test('validatePayload: unknown status rejected', () => {
  const r = validatePayload({
    customer: 'X', licenseType: 'X', expiryDate: '2026-08-15', ownerOid: 'o',
    status: 'wat',
  }, null);
  assert.match(r.error, /status must be one of/);
});

test('validatePayload: unknown renewalCycle rejected', () => {
  const r = validatePayload({
    customer: 'X', licenseType: 'X', expiryDate: '2026-08-15', ownerOid: 'o',
    renewalCycle: 'eternal',
  }, null);
  assert.match(r.error, /renewalCycle must be one of/);
});

// ---- patch path (existing) ----

test('validatePayload: PATCH preserves existing fields when only one is sent', () => {
  const existing = {
    customer: 'Acme', licenseType: 'M365', userCount: 5, expiryDate: '2026-08-15',
    ownerOid: 'o', status: 'notStarted', renewalCycle: 'annual',
  };
  const r = validatePayload({ userCount: 10 }, existing);
  assert.equal(r.license.customer, 'Acme');     // preserved
  assert.equal(r.license.userCount, 10);        // updated
});
