/**
 * MODULE: MANAGEMENT VIEW
 *
 * Weekly aggregation engine for the Management dashboard. Reports by full
 * weeks in EITHER convention — no Week A/B cycle filters here:
 *   anchor = 'wed' → ops week, Wednesday → Tuesday ("7-7")
 *   anchor = 'sun' → corporate week, Sunday → Saturday
 *
 * Pulls straight from the raw segment sheets (WF_OVERTIME, WF_ROLES,
 * WF_COACHING, WF_FURLOUGH, WF_ABSENCES, Stats History) so it never depends
 * on the Week A/B epoch logic used by the floor trackers. Hours are capped
 * at "now" so the current week reads as week-to-date, not the full pasted
 * schedule.
 */

var ManagementView = {

  TZ: 'America/Toronto',
  ROLE_KEYS: { SAFE: 'safe', TOWER: 'tower', ICL: 'icl', ULC: 'ulc' },

  _anchorDay: function(anchor) { return anchor === 'sun' ? 0 : 3; },

  // Week windows, oldest → newest; the last one contains today.
  _weekWindows: function(anchor, weeksBack) {
    var n = Math.max(2, Math.min(16, weeksBack || 8));
    var anchorDay = this._anchorDay(anchor);
    var now = new Date();
    var todayMid = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var diff = (todayMid.getDay() - anchorDay + 7) % 7;
    var curStart = new Date(todayMid.getTime());
    curStart.setDate(curStart.getDate() - diff);
    var wins = [];
    for (var i = n - 1; i >= 0; i--) {
      var s = new Date(curStart.getTime());
      s.setDate(s.getDate() - i * 7);
      var e = new Date(s.getTime());
      e.setDate(e.getDate() + 7);
      wins.push({
        start: s.getTime(),
        end: e.getTime(), // exclusive
        label: Utilities.formatDate(s, this.TZ, 'MMM d') + '–' +
               Utilities.formatDate(new Date(e.getTime() - 86400000), this.TZ, 'MMM d'),
        isCurrent: i === 0
      });
    }
    return wins;
  },

  _emptyWeek: function(w) {
    return {
      label: w.label, isCurrent: w.isCurrent,
      ot: 0, otX1: 0, otX15: 0,
      safe: 0, icl: 0, ulc: 0, tower: 0, coach: 0, acsu: 0,
      absences: 0, slAvg: null, ackAvg: null
    };
  },

  _segHours: function(WT, startStr, endStr) {
    var s = WT._timeToMins(startStr), e = WT._timeToMins(endStr);
    if (s < 0 || e < 0) return 0;
    if (e < s) e += 1440;
    return (e - s) / 60;
  },

  // Epoch (noon local) for a 'yyyy-MM-dd' string; -1 when unparseable.
  _dateEpoch: function(dStr) {
    if (!dStr || dStr.length < 10) return -1;
    var y = parseInt(dStr.substring(0, 4), 10), m = parseInt(dStr.substring(5, 7), 10), d = parseInt(dStr.substring(8, 10), 10);
    if (isNaN(y) || isNaN(m) || isNaN(d)) return -1;
    return new Date(y, m - 1, d, 12, 0, 0, 0).getTime();
  },

  _windowIndex: function(wins, epoch) {
    for (var i = 0; i < wins.length; i++) {
      if (epoch >= wins[i].start && epoch < wins[i].end) return i;
    }
    return -1;
  },

  getDashboard: function(anchor, weeksBack) {
    anchor = anchor === 'sun' ? 'sun' : 'wed';
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null;
    if (!WT) return JSON.stringify({ error: 'Engine unavailable.' });

    var self = this;
    var wins = this._weekWindows(anchor, weeksBack);
    var weeks = wins.map(function(w) { return self._emptyWeek(w); });
    var curIdx = weeks.length - 1;
    var nowMs = Date.now();
    var topOt = {}, topSafe = {}, topTower = {};

    // Generic segment-sheet walker: [Agent, Date, <kind>, Start, End, Region]
    var walk = function(sheetName, onRow) {
      var db = WT._getDB(sheetName);
      if (!db || db.getLastRow() < 2) return;
      var seen = {};
      db.getDataRange().getDisplayValues().slice(1).forEach(function(row) {
        var dStr = WT._formatDate(row[1]);
        var epoch = self._dateEpoch(dStr);
        if (epoch < 0 || epoch > nowMs) return; // week-to-date, not future schedule
        var wi = self._windowIndex(wins, epoch);
        if (wi === -1) return;
        var hash = row[0] + '|' + dStr + '|' + String(row[2]).substring(0, 12) + '|' + row[3] + '|' + row[4];
        if (seen[hash]) return; seen[hash] = true;
        onRow(wi, row);
      });
    };

    // Overtime — WF_OVERTIME: [Agent, Date, Code, Rate, Bucket, IsBreak, Start, End, Region]
    try {
      var dbOt = WT._getDB('WF_OVERTIME');
      if (dbOt && dbOt.getLastRow() > 1) {
        var seenOt = {};
        dbOt.getDataRange().getDisplayValues().slice(1).forEach(function(row) {
          var dStr = WT._formatDate(row[1]);
          var epoch = self._dateEpoch(dStr);
          if (epoch < 0 || epoch > nowMs) return;
          var wi = self._windowIndex(wins, epoch);
          if (wi === -1) return;
          var hash = row[0] + '|' + dStr + '|' + row[2] + '|' + row[6] + '|' + row[7];
          if (seenOt[hash]) return; seenOt[hash] = true;
          var h = self._segHours(WT, row[6], row[7]);
          if (!h) return;
          var rate = parseFloat(row[3]) || 1.0;
          weeks[wi].ot += h;
          if (rate === 1.5) weeks[wi].otX15 += h; else weeks[wi].otX1 += h;
          if (wi === curIdx) {
            var nm = String(row[0]).trim();
            topOt[nm] = (topOt[nm] || 0) + h;
          }
        });
      }
    } catch (e) {}

    // Roles — WF_ROLES: [Agent, Date, Role, Start, End, Region]
    try {
      walk('WF_ROLES', function(wi, row) {
        var h = self._segHours(WT, row[3], row[4]);
        if (!h) return;
        var role = String(row[2]).toUpperCase();
        var nm = String(row[0]).trim();
        if (role.indexOf('SAFE') !== -1) {
          weeks[wi].safe += h;
          if (wi === curIdx) topSafe[nm] = (topSafe[nm] || 0) + h;
        } else if (role.indexOf('TOWER') !== -1 || role.indexOf('WOFQT') !== -1 || role.indexOf('WOQFT') !== -1) {
          weeks[wi].tower += h;
          if (wi === curIdx) topTower[nm] = (topTower[nm] || 0) + h;
        } else if (role.indexOf('ICL') !== -1) {
          weeks[wi].icl += h;
        } else if (role.indexOf('ULC') !== -1 || role.indexOf('FIRE') !== -1) {
          weeks[wi].ulc += h;
        }
      });
    } catch (e) {}

    // Coaching + ACSU hours
    try { walk('WF_COACHING', function(wi, row) { weeks[wi].coach += self._segHours(WT, row[3], row[4]); }); } catch (e) {}
    try { walk('WF_FURLOUGH', function(wi, row) { weeks[wi].acsu += self._segHours(WT, row[3], row[4]); }); } catch (e) {}

    // Absences — count distinct agent+day+type
    try { walk('WF_ABSENCES', function(wi) { weeks[wi].absences += 1; }); } catch (e) {}

    // Service level — weekly averages from Stats History
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stats History');
      if (sheet && sheet.getLastRow() > 1) {
        var sums = wins.map(function() { return { svl: 0, ack: 0, n: 0 }; });
        sheet.getDataRange().getValues().slice(1).forEach(function(r) {
          var t = new Date(r[0]).getTime();
          if (isNaN(t)) return;
          var wi = self._windowIndex(wins, t);
          if (wi === -1) return;
          var svl = parseFloat(r[1]) || 0;
          if (svl > 0 && svl <= 1) svl = svl * 100;
          if (svl <= 0) return;
          sums[wi].svl += svl;
          sums[wi].ack += parseFloat(String(r[2]).replace(/[^\d.]/g, '')) || 0;
          sums[wi].n++;
        });
        sums.forEach(function(s, i) {
          if (s.n > 0) {
            weeks[i].slAvg = Math.round((s.svl / s.n) * 10) / 10;
            weeks[i].ackAvg = Math.round(s.ack / s.n);
          }
        });
      }
    } catch (e) {}

    var round1 = function(v) { return Math.round(v * 10) / 10; };
    weeks.forEach(function(w) {
      ['ot', 'otX1', 'otX15', 'safe', 'icl', 'ulc', 'tower', 'coach', 'acsu'].forEach(function(k) { w[k] = round1(w[k]); });
    });

    var topList = function(map) {
      return Object.keys(map)
        .map(function(n) { return { name: n, hours: round1(map[n]) }; })
        .sort(function(a, b) { return b.hours - a.hours; })
        .slice(0, 5);
    };

    return JSON.stringify({
      anchor: anchor,
      generated: Utilities.formatDate(new Date(), this.TZ, 'yyyy-MM-dd HH:mm'),
      weeks: weeks,
      current: {
        label: weeks[curIdx].label,
        topOt: topList(topOt),
        topSafe: topList(topSafe),
        topTower: topList(topTower),
        otAgents: Object.keys(topOt).length
      }
    });
  }
};
