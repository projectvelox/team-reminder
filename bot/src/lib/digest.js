// Shared license-digest building blocks. Used by both the monthly scheduler
// branch (processMonthlyDigest) and the per-user test endpoint
// (POST /api/licenses/digest/preview).

const PH_OFFSET_MS = 8 * 60 * 60 * 1000;

function phToday() {
  const ph = new Date(Date.now() + PH_OFFSET_MS);
  return `${ph.getUTCFullYear()}-${String(ph.getUTCMonth() + 1).padStart(2, '0')}-${String(ph.getUTCDate()).padStart(2, '0')}`;
}

function periodPrefixes(today) {
  const [yr, mo] = today.split('-').map(Number);
  const nextMo = mo === 12 ? 1 : mo + 1;
  const nextYr = mo === 12 ? yr + 1 : yr;
  return {
    currentPrefix: `${yr}-${String(mo).padStart(2, '0')}`,
    nextPrefix: `${nextYr}-${String(nextMo).padStart(2, '0')}`,
  };
}

// Build the three buckets (overdue / this month / next month) for a single
// owner. Skips abandoned + renewed + missing-expiry rows.
function bucketsForOwner(licenses, ownerOid, today) {
  const { currentPrefix, nextPrefix } = periodPrefixes(today);
  const buckets = { thisMonth: [], nextMonth: [], overdue: [] };
  for (const lic of licenses) {
    if (lic.ownerOid !== ownerOid) continue;
    if (lic.state === 'abandoned') continue;
    if (lic.status === 'renewed') continue;
    if (!lic.expiryDate) continue;
    if (lic.expiryDate < today) buckets.overdue.push(lic);
    else if (lic.expiryDate.startsWith(currentPrefix)) buckets.thisMonth.push(lic);
    else if (lic.expiryDate.startsWith(nextPrefix)) buckets.nextMonth.push(lic);
  }
  return buckets;
}

// Same shape as bucketsForOwner but across every owner — used for the
// roll-up section of the digest for users with licenseRollupDigest = true.
function aggregateAllBuckets(licenses, today) {
  const { currentPrefix, nextPrefix } = periodPrefixes(today);
  const buckets = { thisMonth: [], nextMonth: [], overdue: [] };
  for (const lic of licenses) {
    if (lic.state === 'abandoned') continue;
    if (lic.status === 'renewed') continue;
    if (!lic.expiryDate) continue;
    if (lic.expiryDate < today) buckets.overdue.push(lic);
    else if (lic.expiryDate.startsWith(currentPrefix)) buckets.thisMonth.push(lic);
    else if (lic.expiryDate.startsWith(nextPrefix)) buckets.nextMonth.push(lic);
  }
  return buckets;
}

function esc(s) {
  return String(s || '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
}

function statusLabel(s) {
  return { notStarted: 'Not started', noticeSent: 'Notice sent', awaitingCustomer: 'Awaiting customer', customerConfirmed: 'Customer confirmed', renewed: 'Renewed' }[s] || s || 'Not started';
}

function buildDigestHtml(displayName, buckets, rollupBuckets) {
  const firstName = (displayName || '').split(/\s+/)[0] || 'there';
  const deepLink = 'https://teams.microsoft.com/l/entity/5a03bfa3-63c4-417c-b668-b02234ebc11b/dayReminders.licenses';
  const section = (title, items, color) => {
    if (!items.length) return '';
    const rows = items
      .slice()
      .sort((a, b) => (a.expiryDate || '').localeCompare(b.expiryDate || ''))
      .map((l) => `<tr>
        <td style="padding:6px 10px;border-bottom:1px solid #eee;">${esc(l.customer || '')}</td>
        <td style="padding:6px 10px;border-bottom:1px solid #eee;">${esc(l.licenseType || '')}</td>
        <td style="padding:6px 10px;border-bottom:1px solid #eee;text-align:right;">${l.userCount || 0}</td>
        <td style="padding:6px 10px;border-bottom:1px solid #eee;white-space:nowrap;">${l.expiryDate || ''}</td>
        <td style="padding:6px 10px;border-bottom:1px solid #eee;">${esc(statusLabel(l.status))}</td>
      </tr>`)
      .join('');
    return `<h3 style="margin:24px 0 8px;color:${color};">${title} (${items.length})</h3>
      <table style="border-collapse:collapse;width:100%;font-size:13px;">
        <thead><tr style="background:#f4f4f6;">
          <th style="padding:6px 10px;text-align:left;">Customer</th>
          <th style="padding:6px 10px;text-align:left;">License type</th>
          <th style="padding:6px 10px;text-align:right;">Users</th>
          <th style="padding:6px 10px;text-align:left;">Expires</th>
          <th style="padding:6px 10px;text-align:left;">Status</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>`;
  };
  const rollupHtml = rollupBuckets ? `
    <h2 style="margin:32px 0 8px;border-top:1px solid #ddd;padding-top:24px;">All accounts (across every owner)</h2>
    <p style="color:#666;">You opted in to the all-accounts roll-up in Settings.</p>
    ${section('All overdue', rollupBuckets.overdue, '#c50f1f')}
    ${section('All expiring this month', rollupBuckets.thisMonth, '#ca5010')}
    ${section('All expiring next month', rollupBuckets.nextMonth, '#0078d4')}
  ` : '';
  return `
    <p>Hi ${esc(firstName)},</p>
    <p>Here's your monthly snapshot of license renewals you own:</p>
    ${section('Overdue (action needed)', buckets.overdue, '#c50f1f')}
    ${section('Expiring this month', buckets.thisMonth, '#ca5010')}
    ${section('Expiring next month', buckets.nextMonth, '#0078d4')}
    ${rollupHtml}
    <p style="margin-top:24px"><a href="${deepLink}">Open Day Reminders in Teams</a></p>
    <p style="color:#888;font-size:12px;margin-top:24px">Sent once per month from Day Reminders. To opt out, open Day Reminders &rarr; Licenses &rarr; Settings.</p>
  `;
}

module.exports = {
  phToday,
  periodPrefixes,
  bucketsForOwner,
  aggregateAllBuckets,
  buildDigestHtml,
  statusLabel,
  esc,
};
