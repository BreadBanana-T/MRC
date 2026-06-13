/**
 * MODULE: SAFE TRACKER  (forensic + workforce-deployment lens)
 *
 * Per-agent view of SAFE hours, built to answer two questions with receipts:
 *   1. "Where do this agent's SAFE hours come from, and when did they spike?"
 *   2. "Is this the RIGHT agent for SAFE, or is a higher-value profile
 *       (bilingual / high ERC level) being spent on work a junior EN agent
 *       could do?"  → redeployment ammunition.
 *
 * SOURCES (two doors; manual floor "Assign SAFE" is status-only, never hours):
 *   SCHED — WF_ROLES SAFE blocks   |   OT — WF_OVERTIME SAFE-bucket overtime
 *
 * EXTRAS over the plain forensic view:
 *   - 14-month trailing TREND per agent (independent of the selected window)
 *     so a post-Feb-2026 explosion is visible at a glance, with surge
 *     detection (recent months vs prior baseline).
 *   - MasterList profile per agent: ERC Level, inferred Language (EN/FR/BL
 *     from the Skills field — raw skills surfaced for verification),
 *     supervisor.
 *   - FIT classification:
 *       ideal — EN, level ≤ 2  → SAFE is an appropriate use of this profile
 *       over  — level ≥ 4 OR bilingual → overqualified; SAFE hours here are
 *               "redeployable" to scarce bilingual / senior work
 *       review— everything between (level 3, or FR-mono)
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

  _fitOf: function (level, lang) {
    if (level >= 4 || lang === 'BL') return 'over';
    if (lang === 'EN' && level <= 2) return 'ideal';
    return 'review';
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
                           timeStart: WT._minsToTime(split.startMins), timeEnd: WT._minsToTime(split.endMins), rawDay: dStr };
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
               time: g.timeStart + ' - ' + g.timeEnd, flags: flags };
    });

    var totals = { all: 0, morning: 0, evening: 0, night: 0, sched: 0, ot: 0, count: events.length,
                   redeploy: 0, fitIdeal: 0, fitReview: 0, fitOver: 0 };

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
      var fit = self._fitOf(pf.level, pf.lang);
      var totalH = self._r2(a.total);
      if (fit === 'ideal') totals.fitIdeal += totalH;
      else if (fit === 'over') { totals.fitOver += totalH; totals.redeploy += totalH; }
      else totals.fitReview += totalH;

      // Trend array (14 months) + surge detection.
      var tb = trendByAgent[n] || {};
      var trend = trendMonths.map(function (m) { return { key: m.key, label: m.label, year: m.year, hours: self._r2(tb[m.key] || 0) }; });
      var peak = trend.reduce(function (mx, x) { return x.hours > mx.hours ? x : mx; }, { hours: 0, key: '' });
      var recent = (trend[11].hours + trend[12].hours + trend[13].hours) / 3;
      var prior = (trend[5].hours + trend[6].hours + trend[7].hours + trend[8].hours + trend[9].hours + trend[10].hours) / 6;
      var surge = prior >= 1 && recent >= prior * 2 && recent >= 8;
      var surgeMonth = '';
      if (surge) {
        var base = prior * 1.6;
        for (var ti = 1; ti < trend.length; ti++) {
          if (trend[ti].hours >= base && trend[ti - 1].hours < base) { surgeMonth = trend[ti].label + ' ' + trend[ti].year; break; }
        }
      }

      return {
        name: n, region: a.region,
        level: pf.level, lang: pf.lang, skills: pf.skills, sup: pf.sup, fit: fit,
        total: totalH, morning: self._r2(a.morning), evening: self._r2(a.evening), night: self._r2(a.night),
        srcSched: self._r2(a.srcSched), srcOt: self._r2(a.srcOt),
        days: dayKeys.length, segs: a.segs,
        avgPerDay: dayKeys.length ? self._r2(a.total / dayKeys.length) : 0,
        maxDay: maxDay ? { date: maxDay, hours: a.days[maxDay] } : null,
        dayMap: a.days,
        doubleDays: nDouble, overlapDays: nOverlap,
        trend: trend, trendPeak: { month: peak.key, hours: peak.hours }, surge: surge, surgeMonth: surgeMonth
      };
    }).sort(function (x, y) { return y.total - x.total; });

    ['all', 'morning', 'evening', 'night', 'sched', 'ot', 'redeploy', 'fitIdeal', 'fitReview', 'fitOver'].forEach(function (k) { totals[k] = self._r2(totals[k]); });
    events.sort(function (a, b) { return a.date.localeCompare(b.date) || a.agent.localeCompare(b.agent) || a.src.localeCompare(b.src); });

    return JSON.stringify({
      mode: mode, trackerType: 'safe', label: bounds.label, cycle: bounds.cycle,
      grid: [], events: events, totals: totals, perAgent: perAgent,
      trendMonths: trendMonths.map(function (m) { return m.label + (m.label === 'Jan' ? " '" + String(m.year).slice(-2) : ''); }),
      hasMasterList: Object.keys(mlByKey).length > 0,
      audit: { agents: perAgent.length, doubleAgentDays: Object.keys(doubleDays).length,
               overlapAgentDays: Object.keys(overlapDays).length, surges: perAgent.filter(function (a) { return a.surge; }).length }
    });
  }
};
