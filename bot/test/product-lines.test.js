// Tests for product-line registry helpers (v1.8.0). These are pure logic —
// the canonical seed list and the case-insensitive mapping the normalize
// endpoint uses. Storage-level helpers (listProductLines, etc.) hit Azure
// Tables and aren't covered here; the seed list shape itself is what's most
// likely to silently drift if somebody renames a product line.

const test = require('node:test');
const assert = require('node:assert/strict');

const { CANONICAL_PRODUCT_LINES } = require('../src/lib/store');

// Inline copy of the mapping logic from licenses.js's normalize endpoint so
// these tests stay dependency-free. If the real function gets extracted into
// a shared helper, swap this for an import.
function buildNormMap(mapping) {
  const out = {};
  for (const [from, to] of Object.entries(mapping || {})) {
    const fromKey = String(from || '').trim().toLowerCase();
    const toVal = String(to || '').trim();
    if (!fromKey || !toVal) continue;
    out[fromKey] = toVal;
  }
  return out;
}

function computePreview(licenses, normMap) {
  const preview = [];
  for (const lic of licenses) {
    const cur = String(lic.productLine || '').trim();
    const target = normMap[cur.toLowerCase()];
    if (!target || target === cur) continue;
    preview.push({ id: lic.id, from: cur, to: target });
  }
  return preview;
}

test('CANONICAL_PRODUCT_LINES exposes the 6 v1.8.0 seed values in order', () => {
  assert.deepEqual(CANONICAL_PRODUCT_LINES, [
    'M365',
    'Business Central',
    'Finance and Operation',
    'PHILTAX',
    'CRM',
    'Security',
  ]);
});

test('buildNormMap: empty mapping returns empty object', () => {
  assert.deepEqual(buildNormMap({}), {});
  assert.deepEqual(buildNormMap(null), {});
  assert.deepEqual(buildNormMap(undefined), {});
});

test('buildNormMap: lowercases the FROM key, preserves TO value casing', () => {
  const m = buildNormMap({ 'ERP - BC': 'Business Central', 'PhilTax': 'PHILTAX' });
  assert.equal(m['erp - bc'], 'Business Central');
  assert.equal(m['philtax'], 'PHILTAX');
  assert.equal(m['ERP - BC'], undefined, 'original casing key should not be retained');
});

test('buildNormMap: drops empty / whitespace-only entries on either side', () => {
  const m = buildNormMap({ '': 'X', 'old': '', '   ': 'Y', 'good': 'New' });
  assert.deepEqual(m, { good: 'New' });
});

test('computePreview: rows whose productLine matches a mapping key get rewritten', () => {
  const licenses = [
    { id: '1', productLine: 'ERP - BC' },
    { id: '2', productLine: 'PhilTax' },
    { id: '3', productLine: 'M365' },
    { id: '4', productLine: 'unknown vendor' },
    { id: '5', productLine: '' },
  ];
  const m = buildNormMap({ 'ERP - BC': 'Business Central', 'PhilTax': 'PHILTAX' });
  const preview = computePreview(licenses, m);
  assert.equal(preview.length, 2);
  assert.deepEqual(preview[0], { id: '1', from: 'ERP - BC', to: 'Business Central' });
  assert.deepEqual(preview[1], { id: '2', from: 'PhilTax', to: 'PHILTAX' });
});

test('computePreview: case-insensitive FROM match, exact TO casing applied', () => {
  const licenses = [
    { id: 'a', productLine: 'philtax' },
    { id: 'b', productLine: 'PhilTax' },
    { id: 'c', productLine: 'PHILTAX' }, // already canonical → skipped
  ];
  const m = buildNormMap({ 'PhilTax': 'PHILTAX' });
  const preview = computePreview(licenses, m);
  assert.equal(preview.length, 2);
  assert.equal(preview[0].to, 'PHILTAX');
  assert.equal(preview[1].to, 'PHILTAX');
});

test('computePreview: no-op when target equals current (no spurious events)', () => {
  const licenses = [
    { id: '1', productLine: 'M365' },
    { id: '2', productLine: 'CRM' },
  ];
  const m = buildNormMap({ 'M365': 'M365', 'CRM': 'CRM' });
  assert.deepEqual(computePreview(licenses, m), []);
});

test('computePreview: untouched rows with empty / null productLine', () => {
  const licenses = [
    { id: '1', productLine: '' },
    { id: '2', productLine: null },
    { id: '3', productLine: undefined },
  ];
  const m = buildNormMap({ '': 'M365' }); // empty FROM key — should be dropped by buildNormMap
  assert.deepEqual(computePreview(licenses, m), []);
});
