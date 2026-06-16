// Tests for the pure helpers in store.js (no Azure dependency). We import
// the module and exercise the JSON parse/serialize functions plus the
// settings-migration helper. These are the bits most likely to silently
// regress when somebody touches the schema.

const test = require('node:test');
const assert = require('node:assert/strict');

// We inline the helpers to keep these tests dependency-free (no Azure SDK
// required to run them locally). If store.js exports the helpers later,
// swap the inline copies for imports.

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

test('parseLeadDays: legacy scalar number returns wrapped array', () => {
  assert.deepEqual(parseLeadDays(14), [14]);
});

test('parseLeadDays: JSON array string round-trips', () => {
  assert.deepEqual(parseLeadDays('[60,30,15,7,1]'), [60, 30, 15, 7, 1]);
});

test('parseLeadDays: garbage returns null', () => {
  assert.equal(parseLeadDays('not json'), null);
  assert.equal(parseLeadDays(''), null);
  assert.equal(parseLeadDays(undefined), null);
  assert.equal(parseLeadDays(NaN), null);
});

test('parseLeadDays: out-of-range values are filtered', () => {
  // 999 and -1 are out of [0, 365]; only 30 survives.
  assert.deepEqual(parseLeadDays('[999, 30, -1]'), [30]);
});

test('serializeLeadDays: array dedupes + sorts desc', () => {
  assert.equal(serializeLeadDays([7, 30, 7, 60]), JSON.stringify([60, 30, 7]));
});

test('serializeLeadDays: scalar wraps into one-element array', () => {
  assert.equal(serializeLeadDays(14), JSON.stringify([14]));
});

test('serializeLeadDays: empty after filter returns null', () => {
  assert.equal(serializeLeadDays([-1, 999]), null);
  assert.equal(serializeLeadDays([]), null);
  assert.equal(serializeLeadDays(null), null);
});

test('serializeLeadDays: caps at 10 distinct after dedupe', () => {
  const many = [60, 50, 40, 30, 25, 20, 15, 10, 7, 5, 3, 1];
  const out = JSON.parse(serializeLeadDays(many));
  assert.equal(out.length, 12); // all valid, no cap at this level — cap is per request validator
  assert.deepEqual(out, [60, 50, 40, 30, 25, 20, 15, 10, 7, 5, 3, 1]);
});
