/**
 * MODULE: DAILY DIGEST
 *
 * Emails a morning snapshot of the operation: service level, floor counts,
 * net staffing, today's overtime, live power outages, and unplanned absences.
 *
 * Everything is compiled AT SEND TIME:
 *   - Outages are fetched live from the utility feeds (UrlFetch).
 *   - SL/ACK comes from the latest "Stats History" rows.
 *   - Floor / staffing reflect the current schedule data in the sheet.
 *   - OT / absence numbers are as fresh as the last WFM paste — the digest
 *     footer shows the sync stamps so stale data is visible, not hidden.
 *
 * SETUP (one time, from the Apps Script editor):
 *   1. Run installDailyDigestTrigger()        → daily email at 9 AM Toronto.
 *      Or installDailyDigestTrigger(7)        → any other hour.
 *   2. Optional: setDigestRecipients('lead1@telus.com, lead2@telus.com')
 *      Default recipient is the script owner.
 *   3. Run sendDailyDigest() once manually to authorize + preview.
 *   Remove with removeDailyDigestTrigger().
 */

var DailyDigest = {

  PROP_RECIPIENTS: 'DIGEST_RECIPIENTS',
  TZ: 'America/Toronto',

  _recipients: function() {
    try {
      var p = PropertiesService.getScriptProperties().getProperty(this.PROP_RECIPIENTS);
      if (p && p.trim()) return p.trim();
    } catch (e) {}
    return Session.getEffectiveUser().getEmail();
  },

  // Each section is independently guarded: one broken source must never
  // kill the whole email.
  compile: function() {
    var now = new Date();
    var todayStr = Utilities.formatDate(now, this.TZ, 'yyyy-MM-dd');
    var d = {
      date: todayStr,
      dateNice: Utilities.formatDate(now, this.TZ, 'EEE MMM d, yyyy'),
      generatedAt: Utilities.formatDate(now, this.TZ, 'HH:mm'),
      sl: null, floor: null, staffing: null, outages: null, ot: null, sync: null
    };

    // Service level — latest + today's low from Stats History.
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stats History');
      if (sheet && sheet.getLastRow() > 1) {
        var startRow = Math.max(2, sheet.getLastRow() - 200);
        var rows = sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, 3).getValues();
        var todays = [];
        rows.forEach(function(r) {
          var t = new Date(r[0]);
          if (isNaN(t.getTime())) return;
          if (Utilities.formatDate(t, DailyDigest.TZ, 'yyyy-MM-dd') !== todayStr) return;
          var svl = parseFloat(r[1]) || 0;
          if (svl > 0 && svl <= 1) svl = Math.round(svl * 100);
          todays.push({ t: t, svl: svl, ack: parseFloat(String(r[2]).replace(/[^\d.]/g, '')) || 0 });
        });
        if (todays.length) {
          todays.sort(function(a, b) { return a.t - b.t; });
          var latest = todays[todays.length - 1];
          var low = todays.reduce(function(m, x) { return x.svl < m.svl ? x : m; }, todays[0]);
          d.sl = {
            latest: latest.svl, latestAck: latest.ack,
            latestAt: Utilities.formatDate(latest.t, this.TZ, 'HH:mm'),
            low: low.svl, lowAt: Utilities.formatDate(low.t, this.TZ, 'HH:mm'),
            points: todays.length
          };
        }
      }
    } catch (e) {}

    // Floor counts + unplanned list.
    try {
      var floor = JSON.parse(compileFloorData());
      d.floor = {
        active: (floor.active || []).length,
        unplanned: (floor.unplanned || []).map(function(a) { return { name: a.name || '', reason: a.subStatus || '' }; }),
        planned: (floor.vacation || []).length + (floor.planned || []).length + (floor.training || []).length,
        safe: (floor.safe || []).length, icl: (floor.icl || []).length, ulc: (floor.ulc || []).length,
        startingSoon: (floor.startingSoon || []).length
      };
    } catch (e) {}

    // Net staffing for the current 15-min IDP bucket.
    try {
      var sb = JSON.parse(WorkforceTracker.getStaffingBalance());
      if (sb && sb.available) d.staffing = sb;
    } catch (e) {}

    // Live power outages.
    try {
      var o = JSON.parse(OutageTracker.fetchAll());
      if (o && o.byProvince) {
        var provs = [];
        for (var p in o.byProvince) {
          provs.push({ code: p, customers: o.byProvince[p].customers || 0, outages: o.byProvince[p].outages || 0 });
        }
        d.outages = { provinces: provs, errors: o.errors || [], updated: o.updated || '' };
      }
    } catch (e) {}

    // Overtime scheduled today (Onshore).
    try {
      var ot = JSON.parse(OvertimeTracker.getAnalytics('day', todayStr, 'Onshore', 'ALL'));
      if (ot && !ot.error && ot.totals) d.ot = { all: ot.totals.all || 0, x1: (ot.otTotals || {}).x1 || 0, x15: (ot.otTotals || {}).x15 || 0 };
    } catch (e) {}

    // Data-freshness stamps for the footer.
    try { d.sync = JSON.parse(fetchSyncMetadata()); } catch (e) {}

    return d;
  },

  _esc: function(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  },

  _kpiCell: function(label, value, color) {
    return '<td align="center" style="padding:14px 8px; border-right:1px solid #e7e9f4;">' +
           '<div style="font-size:11px; font-weight:700; color:#94a3b8; text-transform:uppercase; letter-spacing:1px;">' + label + '</div>' +
           '<div style="font-size:24px; font-weight:800; color:' + (color || '#0f172a') + '; margin-top:4px;">' + value + '</div>' +
           '</td>';
  },

  buildHtml: function(d) {
    var esc = this._esc;
    var slVal = d.sl ? d.sl.latest + '%' : '—';
    var slColor = d.sl ? (d.sl.latest >= 80 ? '#059669' : (d.sl.latest >= 70 ? '#d97706' : '#dc2626')) : '#94a3b8';
    var netVal = '—', netColor = '#94a3b8', netSub = '';
    if (d.staffing) {
      netVal = (d.staffing.net > 0 ? '+' : '') + d.staffing.net;
      netColor = d.staffing.status === 'critical' ? '#dc2626' : (d.staffing.status === 'warn' ? '#d97706' : (d.staffing.status === 'surplus' ? '#059669' : '#0f172a'));
      netSub = d.staffing.supply + ' open / ' + d.staffing.demand + ' req @ ' + esc(d.staffing.bucket);
    }
    var unplannedCount = d.floor ? d.floor.unplanned.length : 0;
    var outTotal = 0;
    if (d.outages) d.outages.provinces.forEach(function(p) { outTotal += p.customers; });

    var html = '<div style="font-family:Arial,Helvetica,sans-serif; max-width:640px; margin:0 auto; background:#f4f6ff; padding:24px; color:#0f172a;">';

    // Header
    html += '<div style="background:#4f46e5; border-radius:14px 14px 0 0; padding:20px 24px;">' +
            '<div style="color:#ffffff; font-size:18px; font-weight:800;">MRC Operations Digest</div>' +
            '<div style="color:#c7d2fe; font-size:12px; font-weight:600; margin-top:4px;">' + esc(d.dateNice) + ' &middot; compiled at ' + esc(d.generatedAt) + '</div>' +
            '</div>';

    // KPI strip
    html += '<table width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff; border-collapse:collapse;"><tr>' +
            this._kpiCell('Service Level', slVal, slColor) +
            this._kpiCell('Net Staffing', netVal, netColor) +
            this._kpiCell('Unplanned', String(unplannedCount), unplannedCount > 0 ? '#dc2626' : '#059669') +
            this._kpiCell('Active Floor', d.floor ? String(d.floor.active) : '—') +
            this._kpiCell('OT Today', d.ot ? d.ot.all.toFixed(1) + 'h' : '—') +
            '</tr></table>';

    var section = function(title, body) {
      return '<div style="background:#ffffff; border-top:1px solid #e7e9f4; padding:16px 24px;">' +
             '<div style="font-size:11px; font-weight:800; color:#94a3b8; text-transform:uppercase; letter-spacing:1.5px; margin-bottom:10px;">' + title + '</div>' + body + '</div>';
    };

    // Service level detail
    if (d.sl) {
      html += section('Service Level',
        '<div style="font-size:13px; color:#475569; line-height:1.7;">Latest <b style="color:' + slColor + ';">' + d.sl.latest + '%</b> at ' + esc(d.sl.latestAt) +
        ' &middot; ACK ' + d.sl.latestAck + 's &middot; today\'s low <b>' + d.sl.low + '%</b> at ' + esc(d.sl.lowAt) + ' (' + d.sl.points + ' readings)' +
        (netSub ? '<br>Staffing: ' + netSub : '') + '</div>');
    }

    // Outages
    if (d.outages) {
      var rows = d.outages.provinces.map(function(p) {
        var c = p.customers > 10000 ? '#dc2626' : (p.customers > 1000 ? '#d97706' : '#059669');
        return '<tr><td style="padding:6px 0; font-size:13px; font-weight:700; color:#475569;">' + esc(p.code) + '</td>' +
               '<td align="right" style="padding:6px 0; font-size:13px; font-weight:800; color:' + c + ';">' + p.customers.toLocaleString() + ' customers</td>' +
               '<td align="right" style="padding:6px 0 6px 16px; font-size:12px; color:#94a3b8;">' + p.outages.toLocaleString() + ' outages</td></tr>';
      }).join('');
      var errLine = d.outages.errors.length ? '<div style="font-size:11px; color:#dc2626; font-weight:600; margin-top:6px;">' + d.outages.errors.length + ' source(s) offline</div>' : '';
      html += section('Power Outages — ' + outTotal.toLocaleString() + ' customers (live)',
        '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">' + rows + '</table>' + errLine);
    }

    // Unplanned absences
    if (d.floor && d.floor.unplanned.length) {
      var ul = d.floor.unplanned.slice(0, 12).map(function(a) {
        return '<tr><td style="padding:5px 0; font-size:13px; font-weight:700; color:#0f172a;">' + esc(a.name) + '</td>' +
               '<td align="right" style="padding:5px 0; font-size:12px; font-weight:700; color:#dc2626;">' + esc(a.reason || 'Unplanned') + '</td></tr>';
      }).join('');
      var more = d.floor.unplanned.length > 12 ? '<div style="font-size:11px; color:#94a3b8; margin-top:6px;">+' + (d.floor.unplanned.length - 12) + ' more</div>' : '';
      html += section('Unplanned Absences (' + d.floor.unplanned.length + ')',
        '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">' + ul + '</table>' + more);
    }

    // Floor / roles summary
    if (d.floor) {
      html += section('Floor',
        '<div style="font-size:13px; color:#475569; line-height:1.8;">' +
        '<b>' + d.floor.active + '</b> active &middot; <b>' + d.floor.startingSoon + '</b> starting soon &middot; <b>' + d.floor.planned + '</b> planned off' +
        '<br>Roles: SAFE <b>' + d.floor.safe + '</b> &middot; ICL <b>' + d.floor.icl + '</b> &middot; ULC <b>' + d.floor.ulc + '</b>' +
        (d.ot ? '<br>Overtime today: <b>' + d.ot.all.toFixed(2) + 'h</b> (&times;1: ' + d.ot.x1.toFixed(1) + 'h &middot; &times;1.5: ' + d.ot.x15.toFixed(1) + 'h)' : '') +
        '</div>');
    }

    // Freshness footer
    var syncLine = d.sync ? 'Schedule: ' + esc(d.sync.sched) + ' &middot; IDP: ' + esc(d.sync.idp) : 'Sync metadata unavailable';
    html += '<div style="background:#ffffff; border-top:1px solid #e7e9f4; border-radius:0 0 14px 14px; padding:14px 24px;">' +
            '<div style="font-size:11px; color:#94a3b8; line-height:1.6;">Outages &amp; SL are live at send time. Schedule-derived numbers reflect the last import &mdash; ' + syncLine + '</div>' +
            '</div></div>';
    return html;
  },

  send: function() {
    var d = this.compile();
    var bits = [];
    if (d.sl) bits.push('SL ' + d.sl.latest + '%');
    if (d.staffing) bits.push('Net ' + (d.staffing.net > 0 ? '+' : '') + d.staffing.net);
    if (d.floor) bits.push(d.floor.unplanned.length + ' unplanned');
    var subject = 'MRC Ops Digest — ' + d.dateNice + (bits.length ? ' — ' + bits.join(' · ') : '');
    MailApp.sendEmail({
      to: this._recipients(),
      subject: subject,
      htmlBody: this.buildHtml(d),
      body: 'Open in an HTML-capable mail client to view the MRC Operations Digest.'
    });
    return 'Digest sent to: ' + this._recipients();
  }
};

// ── Globals (run these from the Apps Script editor) ─────────────────────────
function sendDailyDigest() { return DailyDigest.send(); }

function setDigestRecipients(emails) {
  PropertiesService.getScriptProperties().setProperty(DailyDigest.PROP_RECIPIENTS, String(emails || ''));
  return 'Digest recipients: ' + DailyDigest._recipients();
}

function installDailyDigestTrigger(hour) {
  removeDailyDigestTrigger();
  ScriptApp.newTrigger('sendDailyDigest')
    .timeBased()
    .atHour(hour == null ? 9 : hour)
    .everyDays(1)
    .create();
  return 'Daily digest trigger installed (~' + (hour == null ? 9 : hour) + ':00 ' + DailyDigest.TZ + ').';
}

function removeDailyDigestTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendDailyDigest') { ScriptApp.deleteTrigger(t); removed++; }
  });
  return removed + ' digest trigger(s) removed.';
}
