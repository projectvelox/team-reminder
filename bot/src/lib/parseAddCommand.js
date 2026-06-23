// v1.8.4 — extracted from bot.js for unit testing.
//
// parseAddCommand turns a free-text "/add" payload into a structured reminder:
//
//   "5pm #urgent Send report"          -> { title: "Send report", time: "17:00", tags: ["urgent"], dueAt: null }
//   "#qc tomorrow Send report"          -> { title: "Send report", time: null, tags: ["qc"], dueAt: "<tomorrow>" }
//   "Send report #urgent #qc"           -> { title: "Send report", time: null, tags: ["urgent","qc"], dueAt: null }
//
// Tags are extracted from ANYWHERE in the message before time/date detection
// so the order the user types doesn't break parsing.

function parseTimeToken(token) {
  if (!token) return null;
  const lower = token.toLowerCase();
  // 12-hour with am/pm: 5pm, 5:30pm, 5p, 12:00am
  let m = lower.match(/^(\d{1,2})(?::(\d{2}))?\s*(am?|pm?)$/);
  if (m) {
    let h = parseInt(m[1], 10);
    const mm = m[2] ? parseInt(m[2], 10) : 0;
    const isPm = m[3].startsWith('p');
    if (h < 1 || h > 12 || mm < 0 || mm > 59) return null;
    if (h === 12) h = isPm ? 12 : 0;
    else if (isPm) h += 12;
    return `${String(h).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
  }
  // 24-hour HH:MM
  m = lower.match(/^(\d{1,2}):(\d{2})$/);
  if (m) {
    const h = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10);
    if (h > 23 || mm > 59) return null;
    return `${String(h).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
  }
  return null;
}

function parseDateToken(token, today) {
  if (!token) return null;
  const lower = token.toLowerCase();
  if (lower === 'today') return today;
  if (lower === 'tomorrow' || lower === 'tmrw' || lower === 'tom') {
    const d = new Date(today + 'T00:00:00Z');
    d.setUTCDate(d.getUTCDate() + 1);
    return d.toISOString().slice(0, 10);
  }
  const weekdays = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];
  const wdIdx = weekdays.findIndex((w) => lower === w || lower === w + 'day' || (w === 'thu' && lower === 'thur'));
  if (wdIdx >= 0) {
    const d = new Date(today + 'T00:00:00Z');
    const delta = (wdIdx - d.getUTCDay() + 7) % 7;
    d.setUTCDate(d.getUTCDate() + delta);
    return d.toISOString().slice(0, 10);
  }
  let m = lower.match(/^(\d{1,2})[\/-](\d{1,2})$/);
  if (m) {
    const month = parseInt(m[1], 10);
    const day = parseInt(m[2], 10);
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;
    const year = parseInt(today.slice(0, 4), 10);
    const cand = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    if (cand >= today) return cand;
    return `${year + 1}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }
  m = lower.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return lower;
  return null;
}

function parseAddCommand(rest, today) {
  const raw = String(rest || '').trim();
  if (!raw) return null;
  const tags = [];
  // Match `#word` anywhere (letters, digits, _, -). Strip from the text and
  // collect into tags. Case-insensitive, deduped, capped at 8 downstream.
  const cleaned = raw.replace(/(?:^|\s)#([A-Za-z0-9_-]{1,40})\b/g, (_, tag) => {
    const lower = tag.toLowerCase();
    if (!tags.includes(lower)) tags.push(lower);
    return ' ';
  }).replace(/\s+/g, ' ').trim();
  const tokens = cleaned.split(/\s+/).filter(Boolean);
  let i = 0;
  let time = null;
  let dueAt = null;
  // First 2 tokens can be time and/or date in either order; rest is title.
  for (let n = 0; n < 2 && i < tokens.length; n++) {
    if (!time) {
      const t = parseTimeToken(tokens[i]);
      if (t) { time = t; i++; continue; }
    }
    if (!dueAt && today) {
      const d = parseDateToken(tokens[i], today);
      if (d) { dueAt = d; i++; continue; }
    }
    break;
  }
  const title = tokens.slice(i).join(' ').trim();
  if (!title) return null;
  return { title, time, dueAt, tags: tags.slice(0, 8) };
}

module.exports = { parseAddCommand, parseTimeToken, parseDateToken };
