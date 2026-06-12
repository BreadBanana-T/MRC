/**
 * MODULE: SAFE TRACKER
 *
 * Forensic, per-agent view of SAFE hours, reading the same WF_ROLES rows the
 * other dashboards use. Built to answer one question with receipts:
 * "this agent supposedly did 90h of SAFE this month — where exactly is it
 * coming from?"
 *
 * - Hours split Morning / Evening / Night with the same _getShiftSplits
 *   attribution as the Furlough tracker, so the shift columns reconcile.
 * - Only agents with SAFE hours in the window appear (not everyone is a
 *   SAFE agent).
 * - Every segment is listed (date, time window, shift, hours) so a month
 *   total is auditable line by line.
 * - Suspicious patterns are flagged automatically — the usual sources of
 *   inflated totals:
 *     OVERLAP — two SAFE segments on the same day overlapping in time
 *               (double-pasted / double-coded schedules)
 *     LONG    — a single segment of 9h or more
 *     heavy   — a single day totalling more than 10h of SAFE
 */

var SafeTracker = {

  _r2: function (v) { return Math.round(v * 100) / 100; },

  getAnalytics: function (mode, refDate, regionFilter, cycleFilter) {
    regionFilter = regionFilter || 'All';
    cycleFilter = cycleFilter || 'ALL';
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null;
    if (!WT) return JSON.stringify({ error: 'Engine unavailable.' });

    var db = WT._getDB('WF_ROLES');
    if (!db || db.getLastRow() < 2) {
      return JSON.stringify({ error: 'No role data yet. Paste a WFM schedule first.' });
    }

    var self = this;
    var bounds = WT._calculateEpochBoundaries(mode, refDate);
    var searchStart = new Date(bounds.start); searchStart.setDate(searchStart.getDate() - 1);
    var sStr = Utilities.formatDate(searchStart, 'America/Toronto', 'yyyy-MM-dd');
    var eStr = Utilities.formatDate(new Date(bounds.end), 'America/Toronto', 'yyyy-MM-dd');

    var rows = db.getDataRange().getDisplayValues().slice(1);
    var seen = {};
    var grouped = {};
    var rawByAgentDay = {}; // agent|date → [{s,e}] raw segments, for overlap detection

    rows.forEach(function (row) {
      // WF_ROLES holds SAFE / ICL / ULC FIRE / TOWER — SAFE only here.
      if (String(row[2]).toUpperCase().indexOf('SAFE') === -1) return;
      var agent = String(row[0]).trim();
      if (!agent) return;
      var dStr = WT._formatDate(row[1]);
      if (!dStr || dStr < sStr || dStr > eStr) return;
      var region = row[5] ? String(row[5]).trim() : 'Onshore';
      if (regionFilter !== 'All' && region !== regionFilter) return;
      var p = dStr.split('-').map(Number);
      if (p.length < 3 || isNaN(p[0]) || isNaN(p[1]) || isNaN(p[2])) return;
      var sMins = WT._timeToMins(row[3]);
      var eRaw = WT._timeToMins(row[4]);
      if (sMins < 0 || eRaw < 0) return;

      var dedup = agent + '|' + dStr + '|' + sMins + '|' + eRaw;
      if (seen[dedup]) return; seen[dedup] = true;
      var eMins = eRaw < sMins ? eRaw + 1440 : eRaw;

      // Raw segment (pre-split) for overlap detection.
      var odKey = agent + '|' + dStr;
      if (!rawByAgentDay[odKey]) rawByAgentDay[odKey] = [];
      rawByAgentDay[odKey].push({ s: sMins, e: eMins });

      WT._getShiftSplits(sMins, eMins).forEach(function (split) {
        var epoch = new Date(p[0], p[1] - 1, p[2], Math.floor(split.startMins / 60), split.startMins % 60, 0, 0).getTime();
        if (epoch < bounds.start || epoch > bounds.end) return;
        if ((mode === 'month' || mode === 'quarter') && cycleFilter !== 'ALL') {
          if (WT._getCycleForEpoch(epoch) !== cycleFilter) return;
        }
        var effDate = Utilities.formatDate(new Date(epoch), 'America/Toronto', 'yyyy-MM-dd');
        var key = agent + '|' + effDate + '|' + split.shift + '|' + sMins;
        if (!grouped[key]) {
          grouped[key] = { date: effDate, agent: agent, region: region, shift: split.shift, hours: split.hours,
                           timeStart: WT._minsToTime(split.startMins), timeEnd: WT._minsToTime(split.endMins) };
        } else {
          grouped[key].hours += split.hours;
          grouped[key].timeEnd = WT._minsToTime(split.endMins);
        }
      });
    });

    // Overlapping raw segments per agent-day — the classic double-count.
    var overlapDays = {};
    Object.keys(rawByAgentDay).forEach(function (k) {
      var list = rawByAgentDay[k].slice().sort(function (a, b) { return a.s - b.s; });
      for (var i = 1; i < list.length; i++) {
        if (list[i].s < list[i - 1].e) { overlapDays[k] = true; break; }
      }
    });

    var agents = {};
    var events = Object.keys(grouped).map(function (k) {
      var g = grouped[k];
      var h = self._r2(g.hours);
      if (!agents[g.agent]) {
        agents[g.agent] = { name: g.agent, region: g.region, total: 0, morning: 0, evening: 0, night: 0, days: {}, segs: 0 };
      }
      var a = agents[g.agent];
      a.total += h; a.segs++;
      if (g.shift === 'Morning') a.morning += h;
      else if (g.shift === 'Evening') a.evening += h;
      else a.night += h;
      a.days[g.date] = self._r2((a.days[g.date] || 0) + h);
      var flags = [];
      if (h >= 9) flags.push('LONG');
      if (overlapDays[g.agent + '|' + g.date]) flags.push('OVERLAP');
      return { date: g.date, agent: g.agent, shift: g.shift, hours: h, time: g.timeStart + ' - ' + g.timeEnd, flags: flags };
    });

    var totals = { all: 0, morning: 0, evening: 0, night: 0, count: events.length };
    var perAgent = Object.keys(agents).map(function (n) {
      var a = agents[n];
      totals.all += a.total; totals.morning += a.morning; totals.evening += a.evening; totals.night += a.night;
      var dayKeys = Object.keys(a.days);
      var maxDay = null;
      dayKeys.forEach(function (d) { if (!maxDay || a.days[d] > a.days[maxDay]) maxDay = d; });
      return {
        name: n, region: a.region,
        total: self._r2(a.total), morning: self._r2(a.morning), evening: self._r2(a.evening), night: self._r2(a.night),
        days: dayKeys.length, segs: a.segs,
        avgPerDay: dayKeys.length ? self._r2(a.total / dayKeys.length) : 0,
        maxDay: maxDay ? { date: maxDay, hours: a.days[maxDay] } : null,
        heavyDays: dayKeys.filter(function (d) { return a.days[d] > 10; }).length,
        overlapDays: dayKeys.filter(function (d) { return overlapDays[n + '|' + d]; }).length
      };
    }).sort(function (x, y) { return y.total - x.total; });

    ['all', 'morning', 'evening', 'night'].forEach(function (k) { totals[k] = self._r2(totals[k]); });
    events.sort(function (a, b) { return a.date.localeCompare(b.date) || a.agent.localeCompare(b.agent); });

    return JSON.stringify({
      mode: mode, trackerType: 'safe', label: bounds.label, cycle: bounds.cycle,
      grid: [], events: events, totals: totals, perAgent: perAgent,
      audit: {
        agents: perAgent.length,
        overlapAgentDays: Object.keys(overlapDays).length,
        longSegs: events.filter(function (e) { return e.flags.indexOf('LONG') !== -1; }).length
      }
    });
  }
};
