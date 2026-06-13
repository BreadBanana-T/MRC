/**
 * MODULE: SAFE TRACKER  (forensic + workforce-deployment lens)
 *
 * Per-agent view of SAFE hours, built to answer two questions with receipts:
 *   1. "Where do this agent's SAFE hours come from, and when did they spike?"
 *   2. WHEN exactly — open any day to see the morning/evening/night split
 *      and the actual segment windows (a scheduler view of SAFE).
 *
 * SOURCES (two doors; manual floor "Assign SAFE" is status-only, never hours):
 *   SCHED — WF_ROLES SAFE blocks   |   OT — WF_OVERTIME SAFE-bucket overtime
 *
 * EXTRAS over the plain forensic view:
 *   - 14-month trailing TREND per agent (independent of the selected window)
 *     so a post-Feb-2026 explosion is visible at a glance (peak highlighted).
 *   - MasterList profile per agent (data only, no recommendations): ERC
 *     Level, inferred Language (EN/FR/BL from Skills), supervisor.
 *   - Monthly threshold band per agent (normalized to a 30-day month when
 *     not viewing a full month):
 *       RED  ≥ 50 h/mo   |   WARN 40–50   |   NORMAL 25–40   |   LOW < 25
 *
 * Flags: DOUBLE (same block in SCHED and OT), OVERLAP (same-source overlap).
 * Length is NOT flagged — long SAFE segments are normal.
 */

var SafeTracker = {

  _r2: function (v) { return Math.round(v * 100) / 100; },
  _MON: ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],

  // Language inference from the MasterList Skills string. No explicit column
  // exists, so this is a heuristic; the raw skills are shown in the UI.
  _inferLang: function (skills) {
    var s = String(skills || '').toLowerCase();
    var hasFr = /biling|bilingue|\bfr\b|french|fran[cç]ais|_fr\b|\bqc\b/.test(s);
    var hasEn = /\ben\b|english|anglais|_en\b/.test(s);
    if (/biling|bilingue/.test(s) || (hasFr && hasEn)) return 'BL';
    if (hasFr) return 'FR';
    return 'EN';
  },

  getAnalytics: function (mode, refDate, regionFilter, cycleFilter) {
    regionFilter = regionFilter || 'All';
    cycleFilter = cycleFilter || 'ALL';
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null;
    if (!WT) return JSON.stringify({ error: 'Engine unavailable.' });

    var self = this;
    var bounds = WT._calculateEpochBoundaries(mode, refDate);
    var searchStart = new Date(bounds.start); searchStart.setDate(searchStart.getDate() - 1);
    var sStr = Utilities.formatDate(searchStart, 'America/Toronto', 'yyyy-MM-dd');
    var eStr = Utilities.formatDate(new Date(bounds.end), 'America/Toronto', 'yyyy-MM-dd');

    // ── 14-month trailing trend window, anchored on the reference month ──
    var rp = String(refDate || '').split('-');
    var rY = parseInt(rp[0], 10), rM = parseInt(rp[1], 10);
    if (isNaN(rY) || isNaN(rM)) { var nn = new Date(); rY = nn.getFullYear(); rM = nn.getMonth() + 1; }
    var pad2 = function (n) { return (n < 10 ? '0' : '') + n; };
    var trendMonths = [], trendIdx = {};
    for (var t = 13; t >= 0; t--) {
      var dt = new Date(rY, (rM - 1) - t, 1);
      var key = dt.getFullYear() + '-' + pad2(dt.getMonth() + 1);
      trendIdx[key] = trendMonths.length;
      trendMonths.push({ key: key, label: self._MON[dt.getMonth()], year: dt.getFullYear() });
    }
    var trendStartKey = trendMonths[0].key;
    var trendEndKey = trendMonths[trendMonths.length - 1].key;
    var trendByAgent = {};
    var seenTrend = {};
    var addTrend = function (agent, dStr, sMins, eRaw, src) {
      var mo = dStr.substring(0, 7);
      if (trendIdx[mo] === undefined) return;
      var td = src + '|' + agent + '|' + dStr + '|' + sMins + '|' + eRaw;
      if (seenTrend[td]) return; seenTrend[td] = true;
      var em = eRaw <= sMins ? eRaw + 1440 : eRaw;
      if (!trendByAgent[agent]) trendByAgent[agent] = {};
      trendByAgent[agent][mo] = (trendByAgent[agent][mo] || 0) + (em - sMins) / 60;
    };

    // ── MasterList profiles ──
    var mlByKey = {};
    var dbML = WT._getDB('WF_MASTERLIST');
    if (dbML && dbML.getLastRow() > 1) {
      dbML.getDataRange().getDisplayValues().slice(1).forEach(function (r) {
        var nm = String(r[0]).trim();
        if (!nm) return;
        var key = (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey(nm) : nm.toLowerCase();
        mlByKey[key] = { level: parseInt(r[1], 10) || 2, sup: String(r[2] || '').trim(),
                         skills: String(r[3] || '').trim(), lang: self._inferLang(r[3]) };
      });
    }
    var profileOf = function (agent) {
      var key = (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey(agent) : String(agent).toLowerCase();
      return mlByKey[key] || { level: 2, sup: '', skills: '', lang: 'EN' };
    };

    var seen = {};
    var grouped = {};
    var rawByAgentDay = {};

    var regionOf = function (name, rowRegion) {
      var reg = rowRegion ? String(rowRegion).trim() : '';
      if (typeof RegionRegistry !== 'undefined') {
        var rg = RegionRegistry.getRegion(name);
        if (rg) reg = rg;
      }
      return reg || 'Onshore';
    };

    var addSegment = function (agent, dStr, sMins, eRaw, region, src) {
      if (sMins < 0 || eRaw < 0) return;
      if (regionFilter !== 'All' && region !== regionFilter) return;
      var dedup = src + '|' + agent + '|' + dStr + '|' + sMins + '|' + eRaw;
      if (seen[dedup]) return; seen[dedup] = true;
      var eMins = eRaw <= sMins ? eRaw + 1440 : eRaw;
      var p = dStr.split('-').map(Number);
      if (p.length < 3 || isNaN(p[0]) || isNaN(p[1]) || isNaN(p[2])) return;

      var odKey = agent + '|' + dStr;
      if (!rawByAgentDay[odKey]) rawByAgentDay[odKey] = [];
      rawByAgentDay[odKey].push({ s: sMins, e: eMins, src: src });

      WT._getShiftSplits(sMins, eMins).forEach(function (split) {
        var epoch = new Date(p[0], p[1] - 1, p[2], Math.floor(split.startMins / 60), split.startMins % 60, 0, 0).getTime();
        if (epoch < bounds.start || epoch > bounds.end) return;
        if ((mode === 'month' || mode === 'quarter') && cycleFilter !== 'ALL') {
          if (WT._getCycleForEpoch(epoch) !== cycleFilter) return;
        }
        var effDate = Utilities.formatDate(new Date(epoch), 'America/Toronto', 'yyyy-MM-dd');
        var key = src + '|' + agent + '|' + effDate + '|' + split.shift + '|' + sMins;
        if (!grouped[key]) {
          grouped[key] = { date: effDate, agent: agent, region: region, shift: split.shift, src: src, hours: split.hours,
                           startMin: split.startMins % 1440, timeStart: WT._minsToTime(split.startMins), timeEnd: WT._minsToTime(split.endMins), rawDay: dStr };
        } else {
          grouped[key].hours += split.hours;
          grouped[key].timeEnd = WT._minsToTime(split.endMins);
        }
      });
    };

    // 1) SCHED — WF_ROLES SAFE rows
    var dbRoles = WT._getDB('WF_ROLES');
    if (dbRoles && dbRoles.getLastRow() > 1) {
      dbRoles.getDataRange().getDisplayValues().slice(1).forEach(function (row) {
        if (String(row[2]).toUpperCase().indexOf('SAFE') === -1) return;
        var agent = String(row[0]).trim();
        if (!agent) return;
        var dStr = WT._formatDate(row[1]);
        if (!dStr) return;
        var sMins = WT._timeToMins(row[3]), eRaw = WT._timeToMins(row[4]);
        if (dStr >= trendStartKey && dStr <= (trendEndKey + '-31') && sMins >= 0 && eRaw >= 0) addTrend(agent, dStr, sMins, eRaw, 'SCHED');
        if (dStr < sStr || dStr > eStr) return;
        addSegment(agent, dStr, sMins, eRaw, regionOf(agent, row[5]), 'SCHED');
      });
    }

    // 2) OT — WF_OVERTIME SAFE-bucket rows
    var dbOt = WT._getDB('WF_OVERTIME');
    if (dbOt && dbOt.getLastRow() > 1) {
      dbOt.getDataRange().getDisplayValues().slice(1).forEach(function (row) {
        if (String(row[4]).trim().toUpperCase() !== 'SAFE') return;
        var agent = String(row[0]).trim();
        if (!agent) return;
        var dStr = WT._formatDate(row[1]);
        if (!dStr) return;
        var sMins = WT._timeToMins(row[6]), eRaw = WT._timeToMins(row[7]);
        if (dStr >= trendStartKey && dStr <= (trendEndKey + '-31') && sMins >= 0 && eRaw >= 0) addTrend(agent, dStr, sMins, eRaw, 'OT');
        if (dStr < sStr || dStr > eStr) return;
        addSegment(agent, dStr, sMins, eRaw, regionOf(agent, row[8]), 'OT');
      });
    }

    // Flag pass: same-source overlaps + cross-source doubles.
    var overlapDays = {}, doubleDays = {};
    Object.keys(rawByAgentDay).forEach(function (k) {
      var list = rawByAgentDay[k];
      for (var i = 0; i < list.length; i++) {
        for (var j = i + 1; j < list.length; j++) {
          var a = list[i], b = list[j];
          if (a.s < b.e && b.s < a.e) {
            if (a.src === b.src) overlapDays[k] = true; else doubleDays[k] = true;
          }
        }
      }
    });

    var agents = {};
    var events = Object.keys(grouped).map(function (k) {
      var g = grouped[k];
      var h = self._r2(g.hours);
      if (!agents[g.agent]) {
        agents[g.agent] = { name: g.agent, region: g.region, total: 0, morning: 0, evening: 0, night: 0,
                            srcSched: 0, srcOt: 0, days: {}, segs: 0 };
      }
      var a = agents[g.agent];
      a.total += h; a.segs++;
      if (g.shift === 'Morning') a.morning += h;
      else if (g.shift === 'Evening') a.evening += h;
      else a.night += h;
      if (g.src === 'SCHED') a.srcSched += h; else a.srcOt += h;
      a.days[g.date] = self._r2((a.days[g.date] || 0) + h);
      var dayKey = g.agent + '|' + g.rawDay;
      var flags = [];
      if (doubleDays[dayKey]) flags.push('DOUBLE');
      if (overlapDays[dayKey]) flags.push('OVERLAP');
      return { date: g.date, agent: g.agent, shift: g.shift, src: g.src, hours: h,
               startMin: g.startMin, dur: h, time: g.timeStart + ' - ' + g.timeEnd, flags: flags };
    });

    var winDays = Math.max(1, Math.round((bounds.end - bounds.start) / 86400000));
    var bandOf = function (mEq) { return mEq >= 50 ? 'RED' : (mEq >= 40 ? 'WARN' : (mEq < 25 ? 'LOW' : 'NORMAL')); };
    var totals = { all: 0, morning: 0, evening: 0, night: 0, sched: 0, ot: 0, count: events.length,
                   bRed: 0, bWarn: 0, bNormal: 0, bLow: 0 };

    var perAgent = Object.keys(agents).map(function (n) {
      var a = agents[n];
      totals.all += a.total; totals.morning += a.morning; totals.evening += a.evening; totals.night += a.night;
      totals.sched += a.srcSched; totals.ot += a.srcOt;
      var dayKeys = Object.keys(a.days);
      var maxDay = null;
      dayKeys.forEach(function (d) { if (!maxDay || a.days[d] > a.days[maxDay]) maxDay = d; });
      var nDouble = 0, nOverlap = 0;
      Object.keys(doubleDays).forEach(function (k) { if (k.indexOf(n + '|') === 0) nDouble++; });
      Object.keys(overlapDays).forEach(function (k) { if (k.indexOf(n + '|') === 0) nOverlap++; });

      var pf = profileOf(n);
      var totalH = self._r2(a.total);
      var monthEq = (mode === 'month') ? totalH : self._r2(totalH / winDays * 30.42);
      var band = bandOf(monthEq);
      if (band === 'RED') totals.bRed++; else if (band === 'WARN') totals.bWarn++; else if (band === 'LOW') totals.bLow++; else totals.bNormal++;

      // Trend array (14 months) — pure data; the chart speaks for itself.
      var tb = trendByAgent[n] || {};
      var trend = trendMonths.map(function (m) { return { key: m.key, label: m.label, year: m.year, hours: self._r2(tb[m.key] || 0) }; });
      var peak = trend.reduce(function (mx, x) { return x.hours > mx.hours ? x : mx; }, { hours: 0, key: '' });

      return {
        name: n, region: a.region,
        level: pf.level, lang: pf.lang, skills: pf.skills, sup: pf.sup,
        band: band, monthEq: monthEq,
        total: totalH, morning: self._r2(a.morning), evening: self._r2(a.evening), night: self._r2(a.night),
        srcSched: self._r2(a.srcSched), srcOt: self._r2(a.srcOt),
        days: dayKeys.length, segs: a.segs,
        avgPerDay: dayKeys.length ? self._r2(a.total / dayKeys.length) : 0,
        maxDay: maxDay ? { date: maxDay, hours: a.days[maxDay] } : null,
        dayMap: a.days,
        doubleDays: nDouble, overlapDays: nOverlap,
        trend: trend, trendPeak: { month: peak.key, hours: peak.hours }
      };
    }).sort(function (x, y) { return y.total - x.total; });

    ['all', 'morning', 'evening', 'night', 'sched', 'ot'].forEach(function (k) { totals[k] = self._r2(totals[k]); });
    events.sort(function (a, b) { return a.date.localeCompare(b.date) || a.agent.localeCompare(b.agent) || a.src.localeCompare(b.src); });

    return JSON.stringify({
      mode: mode, trackerType: 'safe', label: bounds.label, cycle: bounds.cycle,
      grid: [], events: events, totals: totals, perAgent: perAgent,
      trendMonths: trendMonths.map(function (m) { return m.label + (m.label === 'Jan' ? " '" + String(m.year).slice(-2) : ''); }),
      hasMasterList: Object.keys(mlByKey).length > 0,
      winDays: winDays,
      audit: { agents: perAgent.length, doubleAgentDays: Object.keys(doubleDays).length,
               overlapAgentDays: Object.keys(overlapDays).length }
    });
  },

  // ───────────────────────── SCHEDULE BOARD ─────────────────────────
  // A real "tableau": pick ANY agents + a date, get each agent's full day
  // (shift envelope + breaks/lunch + every off-phone activity), so SAFE can
  // be shown in the context of the whole schedule.
  //
  // Full fidelity (shift + breaks) comes from the "Raw Schedule" sheet, which
  // only holds the CURRENT pasted period. For older dates we fall back to the
  // historical activity sheets (SAFE/ICL/ULC/coaching/ACSU/OT/absence) — no
  // shift envelope or breaks, flagged hasFull=false.

  _normKey: function (n) { return (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey(n) : String(n).trim().toLowerCase(); },

  getScheduleRoster: function () {
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null;
    if (!WT) return '[]';
    var self = this;
    var byKey = {};
    var ml = {};
    var dbML = WT._getDB('WF_MASTERLIST');
    if (dbML && dbML.getLastRow() > 1) {
      dbML.getDataRange().getDisplayValues().slice(1).forEach(function (r) {
        var nm = String(r[0]).trim(); if (!nm) return;
        ml[self._normKey(nm)] = { level: parseInt(r[1], 10) || 2, lang: self._inferLang(r[3]) };
        byKey[self._normKey(nm)] = nm;
      });
    }
    ['Schedule_History', 'Raw Schedule', 'WF_ROLES'].forEach(function (sh) {
      var db = WT._getDB(sh);
      if (!db || db.getLastRow() < 2) return;
      db.getDataRange().getDisplayValues().slice(1).forEach(function (r) {
        var nm = String(r[0]).trim(); if (!nm) return;
        var k = self._normKey(nm); if (!byKey[k]) byKey[k] = nm;
      });
    });
    var out = Object.keys(byKey).map(function (k) {
      var prof = ml[k] || { level: 2, lang: 'EN' };
      return { name: byKey[k], level: prof.level, lang: prof.lang };
    }).sort(function (a, b) { return a.name.localeCompare(b.name); });
    return JSON.stringify(out);
  },

  getScheduleBoard: function (dateStr, agentsPipe) {
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null;
    if (!WT) return JSON.stringify({ error: 'Engine unavailable.' });
    var self = this;
    var wanted = String(agentsPipe || '').split('|').map(function (x) { return x.trim(); }).filter(Boolean);
    if (!wanted.length) return JSON.stringify({ date: dateStr, agents: [] });
    var wantKey = {};
    wanted.forEach(function (n) { wantKey[self._normKey(n)] = n; });

    var toMin = function (t) { return WT._timeToMins(t); };
    var byAgent = {};
    var ensure = function (name) {
      var k = self._normKey(name);
      if (!byAgent[k]) byAgent[k] = { name: name, shiftStart: null, shiftEnd: null, hasFull: false, region: '', segments: [] };
      return byAgent[k];
    };
    var pushSeg = function (ag, type, label, sRaw, eRaw) {
      var sm = toMin(sRaw), em = toMin(eRaw);
      if (sm < 0 || em < 0) return;
      if (em <= sm) em += 1440;
      ag.segments.push({ type: type, label: label, startMin: sm % 1440, endMin: em,
                         time: WT._minsToTime(sm) + ' - ' + WT._minsToTime(em) });
    };

    // 1) Full fidelity (shift + breaks). Schedule_History is the permanent
    //    archive (any date); Raw Schedule holds only the current pasted period.
    //    Read history first, then fill any gaps from Raw Schedule. An agent that
    //    is already "full" from history is never overwritten.
    var anyFull = false;
    var BRK = { 'LUNCH': ['LUNCH', 'Lunch'], 'BREAK': ['BREAK', 'Break'], 'TRAINING': ['COACH', 'Training'],
                'ACSU': ['ACSU', 'ACSU'], 'SAFE': ['SAFE', 'SAFE'], 'ICL': ['ICL', 'ICL'], 'ULC FIRE': ['ULC', 'ULC FIRE'] };
    var readFull = function (sheetName) {
      var rs = WT._getDB(sheetName);
      if (!rs || rs.getLastRow() < 2) return;
      rs.getDataRange().getDisplayValues().slice(1).forEach(function (row) {
        var nm = String(row[0]).trim(); if (!nm) return;
        var k = self._normKey(nm); if (!wantKey[k]) return;
        if (WT._formatDate(row[2]) !== dateStr) return;
        if (byAgent[k] && byAgent[k].hasFull) return; // already covered (history wins over Raw)
        var ag = ensure(wantKey[k]);
        ag.hasFull = true; anyFull = true;
        ag.region = String(row[6] || '').trim();
        // Empty start/end ("Off"/absent days) must NOT become a 00:00–00:00
        // envelope — only set an envelope when both times are present.
        var ssRaw = String(row[3] || '').trim(), seRaw = String(row[4] || '').trim();
        if (ssRaw && seRaw) {
          var ss = toMin(ssRaw), se = toMin(seRaw);
          ag.shiftStart = ss % 1440;
          ag.shiftEnd = (se <= ss ? se + 1440 : se);
        }
        try {
          var brks = JSON.parse(row[7] || '[]');
          brks.forEach(function (b) {
            var m = BRK[String(b.type || '').toUpperCase()] || ['OTHER', String(b.type || 'Activity')];
            pushSeg(ag, m[0], m[1], b.start, b.end);
          });
        } catch (e) {}
        if (String(row[9] || '').trim()) ag.absent = String(row[9]).trim();
      });
    };
    readFull('Schedule_History');
    readFull('Raw Schedule');

    // 2) Historical fallback — activity sheets for agents not covered by Raw Schedule
    var pull = function (sheet, mapper, sIdx, eIdx) {
      var db = WT._getDB(sheet);
      if (!db || db.getLastRow() < 2) return;
      db.getDataRange().getDisplayValues().slice(1).forEach(function (row) {
        var nm = String(row[0]).trim(); var k = self._normKey(nm);
        if (!wantKey[k]) return;
        if (byAgent[k] && byAgent[k].hasFull) return; // already full from Raw Schedule
        if (WT._formatDate(row[1]) !== dateStr) return;
        var t = mapper(row);
        pushSeg(ensure(wantKey[k]), t[0], t[1], row[sIdx], row[eIdx]);
      });
    };
    pull('WF_ROLES', function (r) { var v = String(r[2]).toUpperCase(); return v.indexOf('SAFE') !== -1 ? ['SAFE', 'SAFE'] : (v.indexOf('ICL') !== -1 ? ['ICL', 'ICL'] : ((v.indexOf('ULC') !== -1 || v.indexOf('FIRE') !== -1) ? ['ULC', 'ULC FIRE'] : ['TOWER', 'Tower'])); }, 3, 4);
    pull('WF_COACHING', function () { return ['COACH', 'Coaching']; }, 3, 4);
    pull('WF_FURLOUGH', function () { return ['ACSU', 'ACSU']; }, 3, 4);
    pull('WF_OVERTIME', function (r) { return ['OT', 'OT ' + String(r[4] || '')]; }, 6, 7);
    pull('WF_ABSENCES', function (r) { return ['ABS', String(r[2] || 'Absence')]; }, 3, 4);

    // attach level/lang from MasterList
    var ml = {};
    var dbML = WT._getDB('WF_MASTERLIST');
    if (dbML && dbML.getLastRow() > 1) {
      dbML.getDataRange().getDisplayValues().slice(1).forEach(function (r) {
        var nm = String(r[0]).trim(); if (!nm) return;
        ml[self._normKey(nm)] = { level: parseInt(r[1], 10) || 2, lang: self._inferLang(r[3]) };
      });
    }

    var result = wanted.map(function (n) {
      var k = self._normKey(n);
      var ag = byAgent[k] || { name: n, shiftStart: null, shiftEnd: null, hasFull: false, region: '', segments: [] };
      ag.segments.sort(function (a, b) { return a.startMin - b.startMin; });
      // Envelope: use shift if present, else min/max of segments.
      if (ag.shiftStart == null && ag.segments.length) {
        ag.shiftStart = ag.segments[0].startMin;
        ag.shiftEnd = ag.segments.reduce(function (mx, s2) { return Math.max(mx, s2.endMin); }, 0);
      }
      var p = ml[k] || { level: 2, lang: 'EN' };
      ag.level = p.level; ag.lang = p.lang;
      ag.shiftStartStr = ag.shiftStart != null ? WT._minsToTime(ag.shiftStart) : '';
      ag.shiftEndStr = ag.shiftEnd != null ? WT._minsToTime(ag.shiftEnd) : '';
      // Nothing in any source for this agent+date → honest "not archived" state,
      // never a misleading 00:00–00:00.
      ag.noData = !ag.hasFull && !(ag.segments && ag.segments.length);
      return ag;
    });
    // Whether ANY full (shift+breaks) record exists for this date in either the
    // permanent archive or the current period — drives the board-level banner.
    var hasArchive = (function () {
      var h = WT._getDB('Schedule_History');
      return !!(h && h.getLastRow() > 1);
    })();
    return JSON.stringify({ date: dateStr, agents: result, archived: anyFull, hasArchive: hasArchive,
                            hasMasterList: Object.keys(ml).length > 0 });
  }
};
