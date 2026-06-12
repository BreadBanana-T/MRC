/**
 * MODULE: SAFE TRACKER
 *
 * Forensic, per-agent view of SAFE hours with SOURCE ATTRIBUTION — built to
 * answer "where exactly do this agent's 90h come from", because SAFE hours
 * enter through TWO doors:
 *
 *   SCHED — WF_ROLES: SAFE blocks coded in the pasted WFM schedule
 *   OT    — WF_OVERTIME: overtime decoded into the SAFE bucket
 *           ("OTST … SAFE OnQueue" overtime). NOTE: such an activity ALSO
 *           matches the role classifier, so the same time block can exist
 *           in BOTH sheets — any report that sums them counts it twice.
 *
 * Manual "Assign SAFE" from the Floor menu is deliberately NOT a source:
 * per product owner it is floor status only and must never count as hours
 * (StatusTracker also no longer logs SAFE sessions at write time).
 *
 * Flags (length alone is NOT flagged — long SAFE segments are normal):
 *   DOUBLE  — a SCHED and an OT segment for the same agent overlap in time:
 *             the same hours exist in two sheets (the classic inflation).
 *   OVERLAP — two segments of the SAME source overlap on the same day
 *             (double-pasted schedules).
 *
 * Hours split Morning/Evening/Night with the same _getShiftSplits as the
 * Furlough tracker. Only agents with SAFE hours in the window appear.
 */

var SafeTracker = {

  _r2: function (v) { return Math.round(v * 100) / 100; },

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

    var seen = {};
    var grouped = {};
    var rawByAgentDay = {}; // agent|date → [{s,e,src}] for overlap/double detection

    var regionOf = function (name, rowRegion) {
      var reg = rowRegion ? String(rowRegion).trim() : '';
      if (typeof RegionRegistry !== 'undefined') {
        var rg = RegionRegistry.getRegion(name);
        if (rg) reg = rg;
      }
      return reg || 'Onshore';
    };

    // Shared segment ingestion: splits by shift, honors bounds + cycle,
    // groups for the event log, and records the raw window for flagging.
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

    // 1) SCHED — WF_ROLES rows whose role is SAFE
    var dbRoles = WT._getDB('WF_ROLES');
    if (dbRoles && dbRoles.getLastRow() > 1) {
      dbRoles.getDataRange().getDisplayValues().slice(1).forEach(function (row) {
        if (String(row[2]).toUpperCase().indexOf('SAFE') === -1) return;
        var agent = String(row[0]).trim();
        if (!agent) return;
        var dStr = WT._formatDate(row[1]);
        if (!dStr || dStr < sStr || dStr > eStr) return;
        addSegment(agent, dStr, WT._timeToMins(row[3]), WT._timeToMins(row[4]), regionOf(agent, row[5]), 'SCHED');
      });
    }

    // 2) OT — WF_OVERTIME rows in the SAFE bucket
    var dbOt = WT._getDB('WF_OVERTIME');
    if (dbOt && dbOt.getLastRow() > 1) {
      dbOt.getDataRange().getDisplayValues().slice(1).forEach(function (row) {
        if (String(row[4]).trim().toUpperCase() !== 'SAFE') return;
        var agent = String(row[0]).trim();
        if (!agent) return;
        var dStr = WT._formatDate(row[1]);
        if (!dStr || dStr < sStr || dStr > eStr) return;
        addSegment(agent, dStr, WT._timeToMins(row[6]), WT._timeToMins(row[7]), regionOf(agent, row[8]), 'OT');
      });
    }

    // Flag pass: same-source overlaps + cross-source (SCHED×OT etc.) doubles.
    var overlapDays = {}, doubleDays = {};
    Object.keys(rawByAgentDay).forEach(function (k) {
      var list = rawByAgentDay[k];
      for (var i = 0; i < list.length; i++) {
        for (var j = i + 1; j < list.length; j++) {
          var a = list[i], b = list[j];
          if (a.s < b.e && b.s < a.e) {
            if (a.src === b.src) overlapDays[k] = true;
            else doubleDays[k] = true;
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

    var totals = { all: 0, morning: 0, evening: 0, night: 0, sched: 0, ot: 0, count: events.length };
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
      return {
        name: n, region: a.region,
        total: self._r2(a.total), morning: self._r2(a.morning), evening: self._r2(a.evening), night: self._r2(a.night),
        srcSched: self._r2(a.srcSched), srcOt: self._r2(a.srcOt),
        days: dayKeys.length, segs: a.segs,
        avgPerDay: dayKeys.length ? self._r2(a.total / dayKeys.length) : 0,
        maxDay: maxDay ? { date: maxDay, hours: a.days[maxDay] } : null,
        doubleDays: nDouble, overlapDays: nOverlap
      };
    }).sort(function (x, y) { return y.total - x.total; });

    ['all', 'morning', 'evening', 'night', 'sched', 'ot'].forEach(function (k) { totals[k] = self._r2(totals[k]); });
    events.sort(function (a, b) { return a.date.localeCompare(b.date) || a.agent.localeCompare(b.agent) || a.src.localeCompare(b.src); });

    return JSON.stringify({
      mode: mode, trackerType: 'safe', label: bounds.label, cycle: bounds.cycle,
      grid: [], events: events, totals: totals, perAgent: perAgent,
      audit: {
        agents: perAgent.length,
        doubleAgentDays: Object.keys(doubleDays).length,
        overlapAgentDays: Object.keys(overlapDays).length
      }
    });
  }
};
