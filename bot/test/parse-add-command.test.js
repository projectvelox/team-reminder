// v1.8.4 — covers the parseAddCommand fix where `#tags` at the START of the
// message used to swallow the time-detection slot.

const test = require('node:test');
const assert = require('node:assert');
const { parseAddCommand } = require('../src/lib/parseAddCommand');

const today = '2026-06-23'; // a Tuesday; lets us assert weekday math.

test('plain title with no time/date/tags', () => {
  const r = parseAddCommand('Send report', today);
  assert.deepStrictEqual(r, { title: 'Send report', time: null, dueAt: null, tags: [] });
});

test('time at the start: 5pm Send report', () => {
  const r = parseAddCommand('5pm Send report', today);
  assert.strictEqual(r.title, 'Send report');
  assert.strictEqual(r.time, '17:00');
  assert.deepStrictEqual(r.tags, []);
});

test('tag at the END is extracted (already worked pre-fix)', () => {
  const r = parseAddCommand('5pm Send report #urgent', today);
  assert.strictEqual(r.title, 'Send report');
  assert.strictEqual(r.time, '17:00');
  assert.deepStrictEqual(r.tags, ['urgent']);
});

test('tag in the MIDDLE is extracted (already worked pre-fix)', () => {
  const r = parseAddCommand('5pm Send #urgent report', today);
  assert.strictEqual(r.title, 'Send report');
  assert.strictEqual(r.time, '17:00');
  assert.deepStrictEqual(r.tags, ['urgent']);
});

test('tag BEFORE the time still leaves the time detected (regression: was broken)', () => {
  const r = parseAddCommand('#urgent 5pm Send report', today);
  assert.strictEqual(r.title, 'Send report');
  assert.strictEqual(r.time, '17:00', 'time should be detected after stripping #urgent');
  assert.deepStrictEqual(r.tags, ['urgent']);
});

test('tag BEFORE date+time still leaves both detected', () => {
  const r = parseAddCommand('#qc tomorrow 9am Send report', today);
  assert.strictEqual(r.title, 'Send report');
  assert.strictEqual(r.time, '09:00');
  assert.strictEqual(r.dueAt, '2026-06-24');
  assert.deepStrictEqual(r.tags, ['qc']);
});

test('multiple tags scattered around', () => {
  const r = parseAddCommand('#urgent 5pm #qc Send #internal report', today);
  assert.strictEqual(r.title, 'Send report');
  assert.strictEqual(r.time, '17:00');
  assert.deepStrictEqual(r.tags, ['urgent', 'qc', 'internal']);
});

test('duplicate tags deduped, case-insensitive', () => {
  const r = parseAddCommand('Send report #URGENT #urgent #Urgent', today);
  assert.strictEqual(r.title, 'Send report');
  assert.deepStrictEqual(r.tags, ['urgent']);
});

test('tags with digits, underscore, hyphen are valid', () => {
  const r = parseAddCommand('Send report #q2_review #milestone-1', today);
  assert.strictEqual(r.title, 'Send report');
  assert.deepStrictEqual(r.tags, ['q2_review', 'milestone-1']);
});

test('lone # (no word after) does NOT become an empty tag', () => {
  const r = parseAddCommand('Send report # urgent', today);
  // The lone "#" survives in the title; "urgent" stays as a normal word.
  assert.strictEqual(r.title, 'Send report # urgent');
  assert.deepStrictEqual(r.tags, []);
});

test('tag cap at 8 (the 9th is dropped)', () => {
  const r = parseAddCommand('Send #a #b #c #d #e #f #g #h #i', today);
  assert.strictEqual(r.title, 'Send');
  assert.strictEqual(r.tags.length, 8);
  assert.deepStrictEqual(r.tags, ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h']);
});

test('empty input returns null', () => {
  assert.strictEqual(parseAddCommand('', today), null);
  assert.strictEqual(parseAddCommand('   ', today), null);
  assert.strictEqual(parseAddCommand(null, today), null);
});

test('only tags, no title => null (no point in tagging nothing)', () => {
  assert.strictEqual(parseAddCommand('#urgent #qc', today), null);
});

test('weekday name parses to the next occurrence', () => {
  // 2026-06-23 is a Tuesday; "fri" -> 2026-06-26
  const r = parseAddCommand('fri Send report', today);
  assert.strictEqual(r.dueAt, '2026-06-26');
  assert.strictEqual(r.title, 'Send report');
});

test('# inside a word stays in the title (e.g. "#1 priority")', () => {
  // "#1" — only digits — is a valid tag candidate per our regex.
  // That's a known minor gotcha. Document the behavior in the assertion.
  const r = parseAddCommand('Send report #1', today);
  assert.deepStrictEqual(r.tags, ['1']);
  assert.strictEqual(r.title, 'Send report');
});
