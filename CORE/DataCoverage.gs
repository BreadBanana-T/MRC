/**
 * MODULE: DATA COVERAGE
 *
 * Which days each imported dataset actually covers, so the Data Coverage
 * calendar can show filled / partial / missing days per data type and list
 * what needs updating. Read-only.
 *
 * One batchGet reads ONLY the date column(s) of each dataset sheet (not whole
 * sheets), every cell normalised with WorkforceTracker._formatDate. A second
 * read of System_Notifications gives each dataset's last import time from its
 * "Global Update: … Refreshed." message (the only per-type import log that
 * every import path writes).
 *
 * Grains:
 *   day     → days:   { 'yyyy-MM-dd': rowCount }
 *   month   → months: { 'yyyy-MM': rowCount }
 *   week    → weeks:  { 'yyyy-MM-dd' (Sunday): rowCount }
 *   period  → periods:[ { s, e, n } ]   (start–end inclusive, rows per period)
 * Status ("needs updating") is decided client-side against the viewer's today.
 *
 * Router export getDataCoverage() lives in Code.gs.
 */
var DataCoverage = {
  CACHE_KEY: 'MRC_DATA_COVERAGE_V1',
  CACHE_TTL_SECONDS: 300,

  // key, label, sheet, A1 column range, grain, notification text, how to update.
  DATASETS: [
    { key: 'wfm',      label: 'WFM schedule',       sheet: 'Schedule_History', range: 'C:C', grain: 'day',    note: 'WFM Schedule Refreshed',              how: 'Import → paste the WFM Agent Schedules report' },
    { key: 'idp',      label: 'IDP heatmap',        sheet: 'WF_IDP',           range: 'A:A', grain: 'day',    note: 'IDP Heatmap Refreshed',               how: 'Import → paste the IDP (occupied seats) report', perDay: 96 },
    { key: 'calc',     label: 'Calculator stats',   sheet: 'Stats History',    range: 'A:A', grain: 'day',    note: '',                                     how: 'Calculator → Run Analysis (logs SVL / ACK each run)', hours: true },
    { key: 'forecast', label: 'Forecast',           sheet: 'WF_FORECAST',      range: 'A:B', grain: 'day',    note: 'Forecast Data Refreshed',             how: 'Import → paste the Forecast report (Fcst Alarm / Fcst Acc)' },
    { key: 'otopen',   label: 'OT open slots',      sheet: 'WF_OT_OPEN',       range: 'A:A', grain: 'day',    note: 'OT Open Slots Refreshed',             how: 'Import → paste the open OT slots' },
    { key: 'safe',     label: 'SAFE roster',        sheet: 'WF_SAFE_AGENT',    range: 'A:B', grain: 'period', note: 'SAFE Roster Refreshed',               how: 'Import → paste the SAFE roster and pick its period' },
    { key: 'genesys',  label: 'Genesys agent perf.', sheet: 'WF_GENESYS_AGENT', range: 'A:B', grain: 'period', note: 'Genesys Agent Performance Refreshed', how: 'Import → paste the Genesys Agent Performance Summary and pick its period' },
    { key: 'activity', label: 'Activity loading',   sheet: 'WF_ACTIVITY_WK',   range: 'A:A', grain: 'week',   note: 'Activity Loading Refreshed',          how: 'Import → paste the Activity Loading report' },
    { key: 'gem',      label: 'GEM report',         sheet: 'WF_GEM_DATA_V3',   range: 'A:A', grain: 'month',  note: 'GEM Report Synced',                   how: 'Red Flags → Import GEM Data' },
    { key: 'alarms',   label: 'Alarms by type',     sheet: 'WF_ALARMS',        range: 'A:A', grain: 'month',  note: 'Alarms Data Refreshed',               how: 'Import → paste the Alarms by service type report' }
  ],

  get: function() {
    var ver = '';
    try { ver = PropertiesService.getScriptProperties().getProperty('WF_CACHE_VER') || ''; } catch (e) {}
    var ck = this.CACHE_KEY + '|' + ver;
    try { var hit = CacheService.getScriptCache().get(ck); if (hit) return hit; } catch (e) {}

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var present = this.DATASETS.filter(function(d) { return !!ss.getSheetByName(d.sheet); });
    var cols = this._readColumns(ss, present);
    // The WFM paste also lives in Raw Schedule (current period) — fold it in so
    // a fresh paste counts even before it reaches Schedule_History.
    var raw = ss.getSheetByName('Raw Schedule') ? this._readColumns(ss, [{ sheet: 'Raw Schedule', range: 'C:C' }])['Raw Schedule'] : [];
    var syncs = this._lastSyncs(ss);
    var F = WorkforceTracker._formatDate.bind(WorkforceTracker);

    var out = [];
    var self = this;
    this.DATASETS.forEach(function(d) {
      var rows = (cols[d.sheet] || []).slice(1);   // drop header
      if (d.key === 'wfm') rows = rows.concat((raw || []).slice(1));
      var ds = { key: d.key, label: d.label, grain: d.grain, how: d.how, perDay: d.perDay || 0, exists: !!ss.getSheetByName(d.sheet), lastSync: syncs[d.key] || null };

      if (d.key === 'forecast') {
        // Grain column: 'day' rows are yyyy-MM-dd, 'month' rows are yyyy-MM.
        ds.days = {}; ds.months = {};
        rows.forEach(function(r) {
          var g = String(r[0] || '').toLowerCase(), v = String(r[1] || '').trim();
          if (g === 'month') { var mk = self._monthKey(v); if (mk) ds.months[mk] = (ds.months[mk] || 0) + 1; }
          else { var dk = F(v); if (/^\d{4}-\d{2}-\d{2}$/.test(dk)) ds.days[dk] = (ds.days[dk] || 0) + 1; }
        });
      } else if (d.grain === 'day') {
        ds.days = {};
        if (d.hours) ds.hours = {};
        rows.forEach(function(r) {
          var dk = F(r[0]);
          if (!/^\d{4}-\d{2}-\d{2}$/.test(dk)) return;
          ds.days[dk] = (ds.days[dk] || 0) + 1;
          if (d.hours) {
            var hm = String(r[0]).match(/(\d{1,2}):(\d{2})(?::\d{2})?\s*([AP]M)?/i);
            if (hm) {
              var h = parseInt(hm[1], 10);
              if (hm[3] && /PM/i.test(hm[3]) && h < 12) h += 12;
              if (hm[3] && /AM/i.test(hm[3]) && h === 12) h = 0;
              (ds.hours[dk] = ds.hours[dk] || []).push((h < 10 ? '0' : '') + h + ':' + hm[2]);
            }
          }
        });
      } else if (d.grain === 'month') {
        ds.months = {};
        rows.forEach(function(r) { var mk = self._monthKey(r[0]); if (mk) ds.months[mk] = (ds.months[mk] || 0) + 1; });
      } else if (d.grain === 'week') {
        ds.weeks = {};
        rows.forEach(function(r) { var wk = F(r[0]); if (/^\d{4}-\d{2}-\d{2}$/.test(wk)) ds.weeks[wk] = (ds.weeks[wk] || 0) + 1; });
      } else if (d.grain === 'period') {
        var per = {};
        rows.forEach(function(r) {
          var s = F(r[0]), e = F(r[1]);
          if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return;
          if (!/^\d{4}-\d{2}-\d{2}$/.test(e)) e = s;
          var k = s + '|' + e; per[k] = (per[k] || 0) + 1;
        });
        ds.periods = Object.keys(per).sort().map(function(k) { var p = k.split('|'); return { s: p[0], e: p[1], n: per[k] }; });
      }
      out.push(ds);
    });

    var payload = JSON.stringify({ generated: Date.now(), tz: 'America/Toronto', datasets: out });
    try { CacheService.getScriptCache().put(ck, payload, this.CACHE_TTL_SECONDS); } catch (e) {}
    return payload;
  },

  // 'yyyy-MM' from a month cell ('2026-10', '2026-10-01', a Date, a serial…).
  _monthKey: function(v) {
    var s = String(v == null ? '' : v).trim();
    var m = s.match(/^(\d{4})-(\d{1,2})$/);
    if (m) return m[1] + '-' + (m[2].length < 2 ? '0' : '') + m[2];
    var d = WorkforceTracker._formatDate(v);
    return /^\d{4}-\d{2}-\d{2}$/.test(d) ? d.substring(0, 7) : '';
  },

  // One batchGet for just the needed columns; per-sheet fallback without the
  // Sheets advanced service.
  _readColumns: function(ss, list) {
    var out = {};
    if (!list.length) return out;
    var ranges = list.map(function(d) { return "'" + d.sheet.replace(/'/g, "''") + "'!" + d.range; });
    try {
      if (typeof Sheets !== 'undefined' && Sheets.Spreadsheets && Sheets.Spreadsheets.Values) {
        var resp = Sheets.Spreadsheets.Values.batchGet(ss.getId(), { ranges: ranges, valueRenderOption: 'FORMATTED_VALUE', dateTimeRenderOption: 'FORMATTED_STRING' });
        var vr = (resp && resp.valueRanges) || [];
        list.forEach(function(d, i) { out[d.sheet] = (vr[i] && vr[i].values) || []; });
        return out;
      }
    } catch (e) { Logger.log('[coverage] batchGet failed, per-sheet fallback: ' + e); }
    list.forEach(function(d) {
      var sh = ss.getSheetByName(d.sheet);
      var last = sh ? sh.getLastRow() : 0;
      // 'C:C' → 'C1:C<last>', 'A:B' → 'A1:B<last>'
      out[d.sheet] = last > 0 ? sh.getRange(d.range.replace(/^([A-Z]+):([A-Z]+)$/, '$11:$2' + last)).getDisplayValues() : [];
    });
    return out;
  },

  // Latest "Global Update: … <note>" per dataset from System_Notifications
  // (ID column = epoch ms). Calculator runs aren't notified; its latest
  // reading time comes from the data itself on the client.
  _lastSyncs: function(ss) {
    var out = {};
    var sh = ss.getSheetByName('System_Notifications');
    if (!sh || sh.getLastRow() < 2) return out;
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    var self = this;
    for (var i = rows.length - 1; i >= 0; i--) {
      var msg = String(rows[i][2] || ''), ep = Number(rows[i][0]);
      if (msg.indexOf('Global Update') === -1 || !(ep > 0)) continue;
      self.DATASETS.forEach(function(d) {
        if (d.note && !out[d.key] && msg.indexOf(d.note) !== -1) {
          out[d.key] = { epoch: ep, text: Utilities.formatDate(new Date(ep), 'America/Toronto', 'MMM d, HH:mm') };
        }
      });
    }
    return out;
  }
};
