/**
 * MODULE: MANAGEMENT VIEW
 *
 * Upper-management lens: CORPORATE CALENDAR ONLY — days, Sunday–Saturday
 * weeks, calendar months, quarters. Purely graphical and digestible. The
 * ops dashboard/trackers keep their own 7-7 Wed–Wed world and Week A/B
 * cycles; none of that exists here.
 *
 * getDashboard(grain, refDateStr):
 *   grain  : 'day' (14 buckets) | 'week' (12, Sun–Sat) | 'month' (12) |
 *            'quarter' (8). Default 'week'.
 *   refDate: 'yyyy-MM-dd'; the LAST bucket is the one containing refDate.
 *
 * ACCURACY RULES (consistent with WorkforceTracker / OvertimeTracker):
 *   - Segments wrapping midnight are split at 24:00; each part lands in its
 *     own calendar day, so hours never leak across bucket boundaries.
 *   - A segment part counts once its start time has passed (to-date view).
 *   - Rows are de-duplicated on agent+date+activity+start+end, so re-pasted
 *     schedules can't double count.
 *   - ABSENCES are reported as distinct AGENT-DAYS (an agent absent that day
 *     = 1, however many schedule segments or codes the day contains), with a
 *     by-type and by-region breakdown so the headline explains itself.
 *   - COACHING is reported as sessions (deduped rows) + hours.
 */

var ManagementView = {

  TZ: 'America/Toronto',

  _fmt: function(ms, pattern) { return Utilities.formatDate(new Date(ms), this.TZ, pattern); },

  // Bucket windows, oldest → newest; the LAST one contains refDate.
  _windows: function(grain, refStr) {
    var y = parseInt(refStr.substring(0, 4), 10), m = parseInt(refStr.substring(5, 7), 10), d = parseInt(refStr.substring(8, 10), 10);
    if (isNaN(y) || isNaN(m) || isNaN(d)) { var n2 = new Date(); y = n2.getFullYear(); m = n2.getMonth() + 1; d = n2.getDate(); }
    var wins = [];
    var self = this;
    var push = function(s, e, label) { wins.push({ start: s.getTime(), end: e.getTime(), label: label }); };

    if (grain === 'day') {
      for (var i = 13; i >= 0; i--) {
        var s = new Date(y, m - 1, d - i, 0, 0, 0, 0);
        var e = new Date(y, m - 1, d - i + 1, 0, 0, 0, 0);
        push(s, e, self._fmt(s.getTime(), 'MMM d'));
      }
    } else if (grain === 'month') {
      for (var i2 = 11; i2 >= 0; i2--) {
        var s2 = new Date(y, m - 1 - i2, 1, 0, 0, 0, 0);
        var e2 = new Date(y, m - i2, 1, 0, 0, 0, 0);
        push(s2, e2, self._fmt(s2.getTime(), 'MMM yyyy'));
      }
    } else if (grain === 'quarter') {
      var qStartMonth = Math.floor((m - 1) / 3) * 3;
      for (var i3 = 7; i3 >= 0; i3--) {
        var s3 = new Date(y, qStartMonth - i3 * 3, 1, 0, 0, 0, 0);
        var e3 = new Date(y, qStartMonth - i3 * 3 + 3, 1, 0, 0, 0, 0);
        push(s3, e3, 'Q' + (Math.floor(s3.getMonth() / 3) + 1) + ' ' + s3.getFullYear());
      }
    } else { // week — corporate Sunday → Saturday
      var refMid = new Date(y, m - 1, d, 0, 0, 0, 0);
      var sun = new Date(refMid.getTime());
      sun.setDate(sun.getDate() - sun.getDay());
      for (var i4 = 11; i4 >= 0; i4--) {
        var s4 = new Date(sun.getTime()); s4.setDate(s4.getDate() - i4 * 7);
        var e4 = new Date(s4.getTime()); e4.setDate(e4.getDate() + 7);
        push(s4, e4, self._fmt(s4.getTime(), 'MMM d') + '–' + self._fmt(e4.getTime() - 86400000, 'MMM d'));
      }
    }
    wins.forEach(function(w, i) { w.isSel = (i === wins.length - 1); });
    return wins;
  },

  _emptyBucket: function(w) {
    return {
      label: w.label, isSel: w.isSel,
      ot: 0, otX1: 0, otX15: 0,
      safe: 0, icl: 0, ulc: 0, tower: 0, coach: 0, acsu: 0,
      coachSessions: 0,
      absences: 0, absOn: 0, absOff: 0, absTypes: {},
      slAvg: null, ackAvg: null
    };
  },

  // Midnight epoch (local) for a 'yyyy-MM-dd' string; -1 when unparseable.
  _dayStartEpoch: function(dStr) {
    if (!dStr || dStr.length < 10) return -1;
    var y = parseInt(dStr.substring(0, 4), 10), m = parseInt(dStr.substring(5, 7), 10), d = parseInt(dStr.substring(8, 10), 10);
    if (isNaN(y) || isNaN(m) || isNaN(d)) return -1;
    return new Date(y, m - 1, d, 0, 0, 0, 0).getTime();
  },

  // Split a start/end pair into parts that never cross midnight; each part
  // carries its own start epoch for exact bucket attribution + to-date caps.
  _splitParts: function(WT, dayStartMs, startStr, endStr) {
    var s = WT._timeToMins(startStr), e = WT._timeToMins(endStr);
    if (s < 0 || e < 0 || e === s) return [];
    if (e > s) return [{ epoch: dayStartMs + s * 60000, hours: (e - s) / 60 }];
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

  getDashboard: function(grain, refDateStr) {
    grain = (grain === 'day' || grain === 'month' || grain === 'quarter') ? grain : 'week';
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null;
    if (!WT) return JSON.stringify({ error: 'Engine unavailable.' });
    if (!refDateStr) refDateStr = Utilities.formatDate(new Date(), this.TZ, 'yyyy-MM-dd');

    var self = this;
    var wins = this._windows(grain, refDateStr);
    var buckets = wins.map(function(w) { return self._emptyBucket(w); });
    var selIdx = buckets.length - 1;
    var nowMs = Date.now();
    var topOt = {}, topSafe = {}, topTower = {}, topCoach = {};

    // Generic hour-segment walker: [Agent, Date, <kind>, Start, End, Region].
    // onPart(wi, row, hours, isFirstPart) per midnight-split part that has
    // started and lands in a reported bucket.
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
        var parts = self._splitParts(WT, dayMs, row[3], row[4]);
        for (var pi = 0; pi < parts.length; pi++) {
          if (parts[pi].epoch > nowMs) continue;
          var wi = self._windowIndex(wins, parts[pi].epoch);
          if (wi === -1) continue;
          onPart(wi, row, parts[pi].hours, pi === 0);
        }
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
            buckets[wi].ot += part.hours;
            if (rate === 1.5) buckets[wi].otX15 += part.hours; else buckets[wi].otX1 += part.hours;
            if (wi === selIdx) {
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
          buckets[wi].safe += h;
          if (wi === selIdx) topSafe[nm] = (topSafe[nm] || 0) + h;
        } else if (role.indexOf('TOWER') !== -1 || role.indexOf('WOFQT') !== -1 || role.indexOf('WOQFT') !== -1) {
          buckets[wi].tower += h;
          if (wi === selIdx) topTower[nm] = (topTower[nm] || 0) + h;
        } else if (role.indexOf('ICL') !== -1) {
          buckets[wi].icl += h;
        } else if (role.indexOf('ULC') !== -1 || role.indexOf('FIRE') !== -1) {
          buckets[wi].ulc += h;
        }
      });
    } catch (e) {}

    // Coaching — hours + sessions (a session = one deduped schedule row;
    // counted in the bucket where it starts).
    try {
      walkHours('WF_COACHING', function(wi, row, h, isFirst) {
        buckets[wi].coach += h;
        if (isFirst) {
          buckets[wi].coachSessions += 1;
          if (wi === selIdx) {
            var nm = String(row[0]).trim();
            if (!topCoach[nm]) topCoach[nm] = { sessions: 0, hours: 0 };
            topCoach[nm].sessions += 1;
          }
        }
        if (wi === selIdx) {
          var nm2 = String(row[0]).trim();
          if (!topCoach[nm2]) topCoach[nm2] = { sessions: 0, hours: 0 };
          topCoach[nm2].hours += h;
        }
      });
    } catch (e) {}

    // ACSU hours
    try { walkHours('WF_FURLOUGH', function(wi, row, h) { buckets[wi].acsu += h; }); } catch (e) {}

    // Absences — distinct AGENT-DAYS for the headline (an agent absent that
    // day = 1 no matter how many segments/codes), plus by-type and by-region
    // breakdowns. Counts appear as soon as the day has started.
    try {
      var dbAbs = WT._getDB('WF_ABSENCES');
      if (dbAbs && dbAbs.getLastRow() > 1) {
        var seenDay = {}, seenType = {};
        dbAbs.getDataRange().getDisplayValues().slice(1).forEach(function(row) {
          var dStr = WT._formatDate(row[1]);
          var dayMs = self._dayStartEpoch(dStr);
          if (dayMs < 0 || dayMs > nowMs) return;
          var wi = self._windowIndex(wins, dayMs + 43200000); // midday, robust vs DST
          if (wi === -1) return;
          var agent = String(row[0]).trim();
          var type = String(row[2]).trim() || 'OTHER';
          var region = String(row[5] || '').indexOf('Offshore') !== -1 ? 'Offshore' : 'Onshore';

          var dayKey = agent + '|' + dStr;
          if (!seenDay[dayKey]) {
            seenDay[dayKey] = true;
            buckets[wi].absences += 1;
            if (region === 'Offshore') buckets[wi].absOff += 1; else buckets[wi].absOn += 1;
          }
          var typeKey = agent + '|' + dStr + '|' + type;
          if (!seenType[typeKey]) {
            seenType[typeKey] = true;
            buckets[wi].absTypes[type] = (buckets[wi].absTypes[type] || 0) + 1;
          }
        });
      }
    } catch (e) {}

    // Service level — per-bucket averages from Stats History
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
            buckets[i].slAvg = Math.round((s.svl / s.n) * 10) / 10;
            buckets[i].ackAvg = Math.round(s.ack / s.n);
          }
        });
      }
    } catch (e) {}

    var round1 = function(v) { return Math.round(v * 10) / 10; };
    buckets.forEach(function(b) {
      ['ot', 'otX1', 'otX15', 'safe', 'icl', 'ulc', 'tower', 'coach', 'acsu'].forEach(function(k) { b[k] = round1(b[k]); });
    });

    var topList = function(map) {
      return Object.keys(map)
        .map(function(n) { return { name: n, hours: round1(map[n]) }; })
        .sort(function(a, b) { return b.hours - a.hours; })
        .slice(0, 5);
    };
    var topCoachList = Object.keys(topCoach)
      .map(function(n) { return { name: n, sessions: topCoach[n].sessions, hours: round1(topCoach[n].hours) }; })
      .sort(function(a, b) { return (b.hours - a.hours) || (b.sessions - a.sessions); })
      .slice(0, 5);

    var selWin = wins[selIdx];
    return JSON.stringify({
      grain: grain,
      refDate: refDateStr,
      generated: Utilities.formatDate(new Date(), this.TZ, 'yyyy-MM-dd HH:mm'),
      weeks: buckets, // kept name for the frontend chart code
      sel: {
        label: buckets[selIdx].label,
        startStr: this._fmt(selWin.start, 'yyyy-MM-dd'),
        endStr: this._fmt(selWin.end - 86400000, 'yyyy-MM-dd'),
        topOt: topList(topOt),
        topSafe: topList(topSafe),
        topTower: topList(topTower),
        topCoach: topCoachList,
        otAgents: Object.keys(topOt).length
      }
    });
  }
};
