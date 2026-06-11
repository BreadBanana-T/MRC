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
 * on the Week A/B epoch logic used by the floor trackers.
 *
 * ACCURACY RULES (kept consistent with WorkforceTracker / OvertimeTracker):
 *   - Segments that wrap midnight are split at 24:00 and each part is
 *     attributed to its own calendar day, so hours never leak across a
 *     week boundary (mirrors _getShiftSplits day attribution).
 *   - A segment counts once its start time has passed ("week-to-date"),
 *     so future pasted schedule days never inflate the current week.
 *   - Rows are de-duplicated on agent+date+activity+start+end exactly like
 *     the trackers' eventHash, so re-pasted schedules can't double count.
 *   - Absences count distinct agent+day+type — a split-shift absence is
 *     ONE absence, matching getAbsenceProfiles' merge logic.
 */

var ManagementView = {

  TZ: 'America/Toronto',

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

  // Midnight epoch (local) for a 'yyyy-MM-dd' string; -1 when unparseable.
  _dayStartEpoch: function(dStr) {
    if (!dStr || dStr.length < 10) return -1;
    var y = parseInt(dStr.substring(0, 4), 10), m = parseInt(dStr.substring(5, 7), 10), d = parseInt(dStr.substring(8, 10), 10);
    if (isNaN(y) || isNaN(m) || isNaN(d)) return -1;
    return new Date(y, m - 1, d, 0, 0, 0, 0).getTime();
  },

  /**
   * Split a start/end time pair into parts that never cross midnight.
   * Each part carries the epoch of its own start moment, so:
   *   - week attribution is exact at week boundaries, and
   *   - the "has this started yet" cap works per part, not per row.
   */
  _splitParts: function(WT, dayStartMs, startStr, endStr) {
    var s = WT._timeToMins(startStr), e = WT._timeToMins(endStr);
    if (s < 0 || e < 0 || e === s) return [];
    if (e > s) return [{ epoch: dayStartMs + s * 60000, hours: (e - s) / 60 }];
    // Wraps midnight: tail of day 1 + head of day 2.
    var parts = [{ epoch: dayStartMs + s * 60000, hours: (1440 - s) / 60 }];
    if (e > 0) parts.push({ epoch: dayStartMs + 86400000, hours: e / 60 });
    return parts;
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

    // Generic hour-segment walker: [Agent, Date, <kind>, Start, End, Region].
    // onPart(wi, row, hours) is called once per midnight-split part that has
    // already started and lands in a reported week.
    var walkHours = function(sheetName, onPart) {
      var db = WT._getDB(sheetName);
      if (!db || db.getLastRow() < 2) return;
      var seen = {};
      db.getDataRange().getDisplayValues().slice(1).forEach(function(row) {
        var dStr = WT._formatDate(row[1]);
        var dayMs = self._dayStartEpoch(dStr);
        if (dayMs < 0) return;
        var hash = row[0] + '|' + dStr + '|' + String(row[2]).substring(0, 12) + '|' + row[3] + '|' + row[4];
        if (seen[hash]) return; seen[hash] = true;
        self._splitParts(WT, dayMs, row[3], row[4]).forEach(function(part) {
          if (part.epoch > nowMs) return;
          var wi = self._windowIndex(wins, part.epoch);
          if (wi === -1) return;
          onPart(wi, row, part.hours);
        });
      });
    };

    // Overtime — WF_OVERTIME: [Agent, Date, Code, Rate, Bucket, IsBreak, Start, End, Region]
    try {
      var dbOt = WT._getDB('WF_OVERTIME');
      if (dbOt && dbOt.getLastRow() > 1) {
        var seenOt = {};
        dbOt.getDataRange().getDisplayValues().slice(1).forEach(function(row) {
          var dStr = WT._formatDate(row[1]);
          var dayMs = self._dayStartEpoch(dStr);
          if (dayMs < 0) return;
          var hash = row[0] + '|' + dStr + '|' + row[2] + '|' + row[6] + '|' + row[7];
          if (seenOt[hash]) return; seenOt[hash] = true;
          var rate = parseFloat(row[3]) || 1.0;
          self._splitParts(WT, dayMs, row[6], row[7]).forEach(function(part) {
            if (part.epoch > nowMs) return;
            var wi = self._windowIndex(wins, part.epoch);
            if (wi === -1) return;
            weeks[wi].ot += part.hours;
            if (rate === 1.5) weeks[wi].otX15 += part.hours; else weeks[wi].otX1 += part.hours;
            if (wi === curIdx) {
              var nm = String(row[0]).trim();
              topOt[nm] = (topOt[nm] || 0) + part.hours;
            }
          });
        });
      }
    } catch (e) {}

    // Roles — WF_ROLES: [Agent, Date, Role, Start, End, Region]
    try {
      walkHours('WF_ROLES', function(wi, row, h) {
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
    try { walkHours('WF_COACHING', function(wi, row, h) { weeks[wi].coach += h; }); } catch (e) {}
    try { walkHours('WF_FURLOUGH', function(wi, row, h) { weeks[wi].acsu += h; }); } catch (e) {}

    // Absences — distinct agent+day+type, so split-shift segments collapse
    // to ONE absence exactly like the Absence tracker's merge step. Counts
    // appear as soon as the day has started.
    try {
      var dbAbs = WT._getDB('WF_ABSENCES');
      if (dbAbs && dbAbs.getLastRow() > 1) {
        var seenAbs = {};
        dbAbs.getDataRange().getDisplayValues().slice(1).forEach(function(row) {
          var dStr = WT._formatDate(row[1]);
          var dayMs = self._dayStartEpoch(dStr);
          if (dayMs < 0 || dayMs > nowMs) return;
          var wi = self._windowIndex(wins, dayMs + 43200000); // midday, robust vs DST
          if (wi === -1) return;
          var key = row[0] + '|' + dStr + '|' + String(row[2]).trim();
          if (seenAbs[key]) return; seenAbs[key] = true;
          weeks[wi].absences += 1;
        });
      }
    } catch (e) {}

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
