/**
 * MODULE: MANAGEMENT VIEW
 *
 * Upper-management lens: CORPORATE CALENDAR ONLY — day, Sunday–Saturday
 * week, calendar month, quarter. Purely graphical and digestible. The ops
 * dashboard/trackers keep their own 7-7 Wed–Wed world and Week A/B cycles.
 *
 * getDashboard(grain, refDateStr):
 *   grain  : 'day' | 'week' | 'month' | 'quarter' (default 'week')
 *   refDate: 'yyyy-MM-dd'; the report covers the period CONTAINING refDate.
 *
 * The charts show the SELECTED PERIOD SUBDIVIDED — not N periods of history:
 *   day     → 24 hourly buckets
 *   week    → the 7 Sun–Sat days
 *   month   → that month's days
 *   quarter → that quarter's Sun–Sat weeks (clamped to quarter bounds)
 * KPI deltas compare the selected period against the immediately previous
 * period of the same grain.
 *
 * ACCURACY RULES:
 *   - Hour metrics are distributed by exact interval overlap with each
 *     bucket (overnight segments included), so totals always equal the sum
 *     of the parts and nothing leaks across boundaries.
 *   - Work is counted only up to "now" (to-date view; in-progress segments
 *     count their elapsed portion).
 *   - Rows are de-duplicated on agent+date+activity+start+end.
 *   - ABSENCES are distinct AGENT-DAYS with type/region breakdowns; ALU is
 *     a LATE (own series), ASCLU/SLU/Furlough/ACSU are approved voluntary
 *     leave (excluded). Day priority: real > late > approved.
 *   - UNAB (and every type) carries an offshore count so "most UNABs are
 *     offshore" is visible directly in the graphs.
 *   - ABSENCE HOURS: real-absence hours are also summed (absHours/absHoursOff)
 *     and split Night/Morning/Evening (absShift, via _getShiftSplits) per the
 *     totals AND per bucket (absSeries[].absShift). The client renders "% of
 *     the Sun–Sat week" as absHours ÷ schedH and a concentration-by-shift view.
 *   - CODED/OFFLINE TOTAL: the client sums ot + acsu(VFurlough) + reading +
 *     elearn + tower into one figure shown beside its breakdown. `elearn` reads
 *     WF_ROLES roles matching ELEARN/E-LEARN/VILT/VIRTUAL — 0 until a feed exists.
 *   - lunch (stacking concurrency) and abs (per-type onshore/offshore mix) are
 *     still emitted but are now rendered in the operational Furlough / Absence
 *     trackers, not here (decluttered per management direction).
 */

var ManagementView = {

  TZ: 'America/Toronto',

  _fmt: function(ms, pattern) { return Utilities.formatDate(new Date(ms), this.TZ, pattern); },

  _dayStartEpoch: function(dStr) {
    if (!dStr || dStr.length < 10) return -1;
    var y = parseInt(dStr.substring(0, 4), 10), m = parseInt(dStr.substring(5, 7), 10), d = parseInt(dStr.substring(8, 10), 10);
    if (isNaN(y) || isNaN(m) || isNaN(d)) return -1;
    return new Date(y, m - 1, d, 0, 0, 0, 0).getTime();
  },

  // Bounds of the grain period containing ref, plus the previous period.
  _periodBounds: function(grain, y, m, d) {
    var s, e, ps, pe;
    if (grain === 'day') {
      s = new Date(y, m - 1, d); e = new Date(y, m - 1, d + 1);
      ps = new Date(y, m - 1, d - 1); pe = s;
    } else if (grain === 'month') {
      s = new Date(y, m - 1, 1); e = new Date(y, m, 1);
      ps = new Date(y, m - 2, 1); pe = s;
    } else if (grain === 'quarter') {
      var q = Math.floor((m - 1) / 3) * 3;
      s = new Date(y, q, 1); e = new Date(y, q + 3, 1);
      ps = new Date(y, q - 3, 1); pe = s;
    } else if (grain === 'ytd') {
      // Calendar year so far; previous period is the same span a year earlier.
      s = new Date(y, 0, 1); e = new Date(y + 1, 0, 1);
      ps = new Date(y - 1, 0, 1); pe = s;
    } else { // week, Sunday → Saturday
      var ref = new Date(y, m - 1, d);
      s = new Date(ref.getTime()); s.setDate(s.getDate() - s.getDay());
      e = new Date(s.getTime()); e.setDate(e.getDate() + 7);
      ps = new Date(s.getTime()); ps.setDate(ps.getDate() - 7);
      pe = s;
    }
    return { selStart: s.getTime(), selEnd: e.getTime(), prevStart: ps.getTime(), prevEnd: pe.getTime() };
  },

  // Subdivisions of [selStart, selEnd) for the charts.
  _subWindows: function(grain, selStart, selEnd) {
    var wins = [];
    if (grain === 'day') {
      for (var h = 0; h < 24; h++) {
        wins.push({ start: selStart + h * 3600000, end: selStart + (h + 1) * 3600000, label: (h < 10 ? '0' + h : h) + 'h' });
      }
    } else if (grain === 'week') {
      var DN = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
      for (var i = 0; i < 7; i++) {
        var ds = new Date(selStart); ds.setDate(ds.getDate() + i);
        var de = new Date(ds.getTime()); de.setDate(de.getDate() + 1);
        wins.push({ start: ds.getTime(), end: de.getTime(), label: DN[ds.getDay()] + ' ' + ds.getDate() });
      }
    } else if (grain === 'month') {
      var cur = new Date(selStart);
      while (cur.getTime() < selEnd) {
        var nx = new Date(cur.getTime()); nx.setDate(nx.getDate() + 1);
        wins.push({ start: cur.getTime(), end: Math.min(nx.getTime(), selEnd), label: String(cur.getDate()) });
        cur = nx;
      }
    } else if (grain === 'ytd') { // calendar year → 12 month columns
      var mc = new Date(selStart);
      while (mc.getTime() < selEnd) {
        var mn = new Date(mc.getFullYear(), mc.getMonth() + 1, 1);
        wins.push({ start: mc.getTime(), end: Math.min(mn.getTime(), selEnd), label: this._fmt(mc.getTime(), 'MMM') });
        mc = mn;
      }
    } else { // quarter → Sun–Sat weeks clamped to the quarter
      var wk = new Date(selStart); wk.setDate(wk.getDate() - wk.getDay()); // Sunday on/before quarter start
      while (wk.getTime() < selEnd) {
        var we = new Date(wk.getTime()); we.setDate(we.getDate() + 7);
        var cs = Math.max(wk.getTime(), selStart), ce = Math.min(we.getTime(), selEnd);
        if (ce > cs) wins.push({ start: cs, end: ce, label: this._fmt(cs, 'MMM d') });
        wk = we;
      }
    }
    return wins;
  },

  _windowIndex: function(wins, epoch) {
    for (var i = 0; i < wins.length; i++) {
      if (epoch >= wins[i].start && epoch < wins[i].end) return i;
    }
    return -1;
  },

  // Absolute [start, end) ms interval for a segment on a given day.
  // Overnight wrap (end <= start) extends past midnight; interval-overlap
  // bucketing then attributes each portion to the right window exactly.
  _segInterval: function(WT, dayStartMs, startStr, endStr) {
    var s = WT._timeToMins(startStr), e = WT._timeToMins(endStr);
    if (s < 0 || e < 0 || e === s) return null;
    if (e <= s) e += 1440;
    return { s: dayStartMs + s * 60000, e: dayStartMs + e * 60000 };
  },

  _overlapH: function(iv, start, end) {
    var o = Math.min(iv.e, end) - Math.max(iv.s, start);
    return o > 0 ? o / 3600000 : 0;
  },

  // Map a raw WF_COACHING activity string to its IEX code (matches the manager's
  // "IEX Coding Compiled" + "Misc" tables). Returns null when it isn't one of
  // the tracked codes, so we don't invent buckets.
  _coachCode: function(act) {
    act = String(act || '').toLowerCase();
    if (/huddle/.test(act)) return 'CE-HUDDLE';
    if (/qual|quality/.test(act)) return 'QUAL';
    if (/one on one|one-on-one|individuelle/.test(act)) return 'ONE';
    if (/sbys|side ?by ?side|sidexside/.test(act)) return 'SBYS';
    if (/vilt|virtual/.test(act)) return 'VILT';
    if (/e[- ]?learning|elearn/.test(act)) return 'E Learning';
    if (/classroom/.test(act)) return 'CLASSROOM';
    if (/roadshow|\bmeet\b|réunion/.test(act)) return 'MEET';
    if (/lfqi|on loan|off ?q|opex/.test(act)) return 'LFQI';
    if (/\bauth\b|authorized/.test(act)) return 'AUTH';
    if (/\bsme\b/.test(act)) return 'SME';
    return null;
  },

  getDashboard: function(grain, refDateStr) {
    grain = (grain === 'day' || grain === 'month' || grain === 'quarter' || grain === 'ytd') ? grain : 'week';
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null;
    if (!WT) return JSON.stringify({ error: 'Engine unavailable.' });
    if (!refDateStr || refDateStr.length < 10) refDateStr = Utilities.formatDate(new Date(), this.TZ, 'yyyy-MM-dd');
    var ry = parseInt(refDateStr.substring(0, 4), 10), rm = parseInt(refDateStr.substring(5, 7), 10), rd = parseInt(refDateStr.substring(8, 10), 10);
    if (isNaN(ry) || isNaN(rm) || isNaN(rd)) {
      var n0 = new Date(); ry = n0.getFullYear(); rm = n0.getMonth() + 1; rd = n0.getDate();
    }

    var self = this;
    var nowMs = Date.now();
    var pb = this._periodBounds(grain, ry, rm, rd);

    // Payload cache — ONLY for fully-past periods (static data). The current /
    // in-progress period is never cached, because its SL/IDP and imported
    // Activity-report data change live and must always show fresh (this is what
    // caused stale "pending feed" after an import). Past periods still get a
    // fast cache, invalidated by WF_CACHE_VER (bumped on every import).
    var _isPast = pb.selEnd < nowMs - 86400000;
    var _ver = ''; try { _ver = PropertiesService.getScriptProperties().getProperty('WF_CACHE_VER') || ''; } catch (e) {}
    var _ck = 'mgmtDash|' + grain + '|' + refDateStr + '|' + _ver;
    var _cache = null;
    if (_isPast) {
      try { _cache = CacheService.getScriptCache(); } catch (e) {}
      if (_cache) { try { var _hit = _cache.get(_ck); if (_hit) return Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(_hit), 'application/x-gzip', 'c.gz')).getDataAsString(); } catch (e) {} }
    }
    var wins = this._subWindows(grain, pb.selStart, pb.selEnd);
    var selCap = Math.min(pb.selEnd, nowMs);   // to-date cap
    var prevCap = Math.min(pb.prevEnd, nowMs);

    var newTotals = function() {
      return { ot: 0, otX1: 0, otX15: 0, safe: 0, tower: 0, coach: 0, acsu: 0, elearn: 0,
               coachSessions: 0, absences: 0, absOn: 0, absOff: 0, lates: 0, approvedLeave: 0,
               absTypes: {}, absTypesOff: {}, slAvg: null, ackAvg: null,
               // Absence HOURS (not just agent-day counts) + shift concentration.
               // absShift is real-absence hours split Night/Morning/Evening; the
               // "% of week" is derived on the client as absHours ÷ schedH.
               absHours: 0, absHoursOff: 0, absShift: { Night: 0, Morning: 0, Evening: 0 },
               // Per-shift AGENT counts (a real-absence agent-day counted in each
               // shift it touches) — for the SL-vs-absence chart (agents, not hours).
               absShiftCount: { Night: 0, Morning: 0, Evening: 0 },
               // Per-IEX-code hours for the "IEX Coding Compiled / Misc" card
               // (period to-date totals). Keyed by code, e.g. {'CE-HUDDLE':8.4}.
               codes: {},
               openOt: 0, openOtToDate: 0, openSlots: 0, openSkills: {}, idpDeficit: 0, idpNet: null, idpReq: null, idpOpen: null, _idpSum: 0, _idpN: 0, _idpReq: 0, _idpOpen: 0,
               preCodedOt: 0,
               // Shrinkage inputs: scheduled hours (denominator) + off-phone hour buckets.
               reading: 0, breakH: 0, lunchH: 0, schedH: 0 };
    };
    var selT = newTotals(), prevT = newTotals();
    var buckets = wins.map(function(w) {
      return { label: w.label, ot: 0, otX1: 0, otX15: 0, safe: 0, tower: 0,
               coach: 0, acsu: 0, elearn: 0, coachSessions: 0, slAvg: null, ackAvg: null,
               openOt: 0, openSlots: 0, idpDeficit: 0, idpNet: null, idpReq: null, idpOpen: null, _idpSum: 0, _idpN: 0, _idpReq: 0, _idpOpen: 0,
               reading: 0, breakH: 0, lunchH: 0, schedH: 0 };
    });
    // Absence series: aligned with buckets for week/month/quarter; a single
    // whole-period entry for day grain (hourly absence bars are meaningless).
    // absHours + absShift power the per-shift concentration + SL-vs-absence-by-shift charts.
    var newAbsEntry = function(label) {
      return { label: label, absences: 0, absOn: 0, absOff: 0, lates: 0, approvedLeave: 0,
               absTypes: {}, absTypesOff: {}, absHours: 0, absShift: { Night: 0, Morning: 0, Evening: 0 }, absShiftCount: { Night: 0, Morning: 0, Evening: 0 } };
    };
    var absSeries = (grain === 'day')
      ? [newAbsEntry(wins.length ? this._fmt(pb.selStart, 'MMM d') : '')]
      : wins.map(function(w) { return newAbsEntry(w.label); });

    var topOt = {}, topSafe = {}, topTower = {}, topCoach = {};

    // Cold-load speedup: pull every display-value sheet this dashboard reads in
    // ONE Sheets API round-trip, instead of a separate getDataRange() per sheet.
    // (Stats History stays on its own typed read below.) rowsOf() returns the
    // full rows incl. header, so call sites keep their .slice(1).
    var MVROWS = WT._batchDisplayValues(['WF_OVERTIME', 'WF_ROLES', 'WF_COACHING', 'WF_FURLOUGH', 'WF_OT_OPEN', 'WF_IDP', 'WF_ABSENCES', 'Schedule_History', 'Raw Schedule']);
    var rowsOf = function(n) { return MVROWS[n] || []; };

    // Distribute one segment's hours into sub-buckets + sel/prev totals.
    // addFn(target, hours) writes into whichever totals object.
    // uncapped=true is for PLAN data (posted OT slots, IDP forecast): a
    // posting/forecast for the back half of the period is already known, so
    // it shows for the whole period. Actuals (worked hours, absences) stay
    // to-date capped.
    var distribute = function(iv, addFn, bucketFn, uncapped) {
      var capSel = uncapped ? pb.selEnd : selCap;
      var capPrev = uncapped ? pb.prevEnd : prevCap;
      var selH = self._overlapH(iv, pb.selStart, capSel);
      if (selH > 0) {
        addFn(selT, selH);
        for (var i = 0; i < wins.length; i++) {
          var h = self._overlapH(iv, wins[i].start, Math.min(wins[i].end, capSel));
          if (h > 0 && bucketFn) bucketFn(buckets[i], h);
        }
      }
      var prevH = self._overlapH(iv, pb.prevStart, capPrev);
      if (prevH > 0) addFn(prevT, prevH);
      return selH;
    };

    // Generic hour-segment walker: [Agent, Date, <kind>, Start, End, Region]
    var walkHours = function(sheetName, onSeg) {
      var rows = rowsOf(sheetName);
      if (rows.length < 2) return;
      var seen = {};
      rows.slice(1).forEach(function(row) {
        var dStr = WT._formatDate(row[1]);
        var dayMs = self._dayStartEpoch(dStr);
        if (dayMs < 0) return;
        // Cheap range gate before hashing: segment can only matter if its day
        // touches [prevStart, selEnd).
        if (dayMs + 86400000 * 2 < pb.prevStart || dayMs >= pb.selEnd) return;
        var hash = row[0] + '|' + dStr + '|' + String(row[2]).substring(0, 12) + '|' + row[3] + '|' + row[4];
        if (seen[hash]) return; seen[hash] = true;
        var iv = self._segInterval(WT, dayMs, row[3], row[4]);
        if (iv) onSeg(iv, row);
      });
    };

    // Overtime — WF_OVERTIME: [Agent, Date, Code, Rate, Bucket, IsBreak, Start, End, Region]
    try {
      var otRows = rowsOf('WF_OVERTIME');
      if (otRows.length > 1) {
        var seenOt = {};
        otRows.slice(1).forEach(function(row) {
          var dStr = WT._formatDate(row[1]);
          var dayMs = self._dayStartEpoch(dStr);
          if (dayMs < 0) return;
          if (dayMs + 86400000 * 2 < pb.prevStart || dayMs >= pb.selEnd) return;
          var hash = row[0] + '|' + dStr + '|' + row[2] + '|' + row[6] + '|' + row[7];
          if (seenOt[hash]) return; seenOt[hash] = true;
          var iv = self._segInterval(WT, dayMs, row[6], row[7]);
          if (!iv) return;
          var prem = (parseFloat(row[3]) || 1.0) === 1.5;
          var selH = distribute(iv,
            function(t, h) { t.ot += h; if (prem) t.otX15 += h; else t.otX1 += h; },
            function(b, h) { b.ot += h; if (prem) b.otX15 += h; else b.otX1 += h; });
          if (selH > 0) {
            var nm = String(row[0]).trim();
            topOt[nm] = (topOt[nm] || 0) + selH;
          }
          // Pre-coded future OT: the portion of this OT beyond the to-date cap.
          // Worked OT (selH) stays capped at today; this is shown separately so
          // already-coded future OT isn't invisible (it appears in the OT Tracker).
          var futH = self._overlapH(iv, Math.max(pb.selStart, selCap), pb.selEnd);
          if (futH > 0) selT.preCodedOt += futH;
        });
      }
    } catch (e) {}

    // Roles — WF_ROLES: [Agent, Date, Role, Start, End, Region]
    try {
      walkHours('WF_ROLES', function(iv, row) {
        var role = String(row[2]).toUpperCase();
        var key = null;
        if (role.indexOf('SAFE') !== -1) key = 'safe';
        else if (role.indexOf('TOWER') !== -1 || role.indexOf('WOFQT') !== -1 || role.indexOf('WOQFT') !== -1) key = 'tower';
        else if (role.indexOf('READING') !== -1 || role.indexOf('LECTURE') !== -1) key = 'reading';
        else if (role.indexOf('ELEARN') !== -1 || role.indexOf('E-LEARN') !== -1 || role.indexOf('VILT') !== -1 || role.indexOf('VIRTUAL') !== -1) key = 'elearn';
        if (!key) return; // ICL / ULC FIRE are no longer tracked (live Floor only)
        var selH = distribute(iv,
          function(t, h) { t[key] += h; },
          function(b, h) { b[key] += h; });
        if (selH > 0) {
          var nm = String(row[0]).trim();
          if (key === 'safe') topSafe[nm] = (topSafe[nm] || 0) + selH;
          else if (key === 'tower') topTower[nm] = (topTower[nm] || 0) + selH;
          // IEX coding rows derived from roles (offtask group).
          var rcode = (key === 'reading') ? 'READ' : (key === 'tower') ? 'WOFQT' : (key === 'elearn') ? 'VILT' : null;
          if (rcode) selT.codes[rcode] = (selT.codes[rcode] || 0) + selH;
        }
      });
    } catch (e) {}

    // Coaching — hours + sessions (session counted where it starts)
    try {
      walkHours('WF_COACHING', function(iv, row) {
        var selH = distribute(iv,
          function(t, h) { t.coach += h; },
          function(b, h) { b.coach += h; });
        var nm = String(row[0]).trim();
        if (iv.s >= pb.selStart && iv.s < selCap) {
          selT.coachSessions += 1;
          var wi = self._windowIndex(wins, iv.s);
          if (wi !== -1) buckets[wi].coachSessions += 1;
          if (!topCoach[nm]) topCoach[nm] = { sessions: 0, hours: 0 };
          topCoach[nm].sessions += 1;
        }
        if (iv.s >= pb.prevStart && iv.s < prevCap) prevT.coachSessions += 1;
        if (selH > 0) {
          if (!topCoach[nm]) topCoach[nm] = { sessions: 0, hours: 0 };
          topCoach[nm].hours += selH;
          var ccode = self._coachCode(row[2]);
          if (ccode) selT.codes[ccode] = (selT.codes[ccode] || 0) + selH;
        }
      });
    } catch (e) {}

    // ACSU hours
    try { walkHours('WF_FURLOUGH', function(iv) { distribute(iv, function(t, h) { t.acsu += h; }, function(b, h) { b.acsu += h; }); }); } catch (e) {}

    // Open OT slots — WF_OT_OPEN (human layout): [Date, Start Time, End Time,
    // Slots, Activity, Min Length, ADG, ADV, Type, Rate, Skill, WindowHours,
    // OpenHours, Visible, OID]. Only Type=OT rows count (ACSU released-time
    // is NOT overtime); hidden postings excluded; hours = window overlap ×
    // slot count, to-date capped like the rest of the view.
    try {
      var openRows = rowsOf('WF_OT_OPEN');
      if (openRows.length > 1) {
        var seenOp = {};
        openRows.slice(1).forEach(function(row) {
          if (String(row[8]) !== 'OT') return;
          if (String(row[13]) === 'N') return;
          var dStr = WT._formatDate(row[0]);
          var dayMs = self._dayStartEpoch(dStr);
          if (dayMs < 0) return;
          if (dayMs + 86400000 * 2 < pb.prevStart || dayMs >= pb.selEnd) return;
          var oid = String(row[14] || (dStr + '|' + row[1] + '|' + row[2] + '|' + row[10]));
          if (seenOp[oid]) return; seenOp[oid] = true;
          var slots = parseInt(row[3], 10) || 0;
          if (!slots) return;
          var iv = self._segInterval(WT, dayMs, row[1], row[2]);
          if (!iv) return;
          var skill = String(row[10]) || 'Any agent';
          var selH = distribute(iv,
            function(t, h) { t.openOt += h * slots; },
            function(b, h) { b.openOt += h * slots; },
            true); // postings are plan data — show the full period
          if (selH > 0) selT.openSkills[skill] = Math.round(((selT.openSkills[skill] || 0) + selH * slots) * 10) / 10;
          // To-date portion so the fill rate compares like-for-like with
          // worked (to-date) OT hours.
          var tdH = self._overlapH(iv, pb.selStart, selCap);
          if (tdH > 0) selT.openOtToDate += tdH * slots;
          if (iv.s >= pb.selStart && iv.s < pb.selEnd) {
            selT.openSlots += slots;
            var wi2 = self._windowIndex(wins, iv.s);
            if (wi2 !== -1) buckets[wi2].openSlots += slots;
          }
          if (iv.s >= pb.prevStart && iv.s < pb.prevEnd) prevT.openSlots += slots;
        });
      }
    } catch (e) {}

    // IDP — WF_IDP: [Day, Interval, Required, Open] per 15-min bucket.
    // idpDeficit = Σ max(0, required − open) × 0.25 = agent-hours short.
    // idpNet = average net seats over the bucket's intervals.
    try {
      var idpRows = rowsOf('WF_IDP');
      if (idpRows.length > 1) {
        var seenIdp = {};
        idpRows.slice(1).forEach(function(row) {
          var dStr = WT._formatDate(row[0]);
          var dayMs = self._dayStartEpoch(dStr);
          if (dayMs < 0) return;
          if (dayMs + 86400000 < pb.prevStart || dayMs >= pb.selEnd) return;
          var tMin = WT._timeToMins(WT._formatTimeStr(String(row[1])));
          if (tMin < 0) return;
          var key = dStr + '|' + tMin;
          if (seenIdp[key]) return; seenIdp[key] = true;
          var epoch = dayMs + tMin * 60000;
          var req = parseFloat(String(row[2]).replace(',', '.')) || 0;
          var seats = parseFloat(String(row[3]).replace(',', '.')) || 0;
          var net = seats - req;
          var defH = net < 0 ? -net * 0.25 : 0;
          // IDP is a forecast grid — the deficit for the rest of the period
          // is already known, so no to-date cap here.
          if (epoch >= pb.selStart && epoch < pb.selEnd) {
            selT.idpDeficit += defH; selT._idpSum += net; selT._idpN++; selT._idpReq += req; selT._idpOpen += seats;
            var wi3 = self._windowIndex(wins, epoch);
            if (wi3 !== -1) { buckets[wi3].idpDeficit += defH; buckets[wi3]._idpSum += net; buckets[wi3]._idpN++; buckets[wi3]._idpReq += req; buckets[wi3]._idpOpen += seats; }
          } else if (epoch >= pb.prevStart && epoch < pb.prevEnd) {
            prevT.idpDeficit += defH; prevT._idpSum += net; prevT._idpN++; prevT._idpReq += req; prevT._idpOpen += seats;
          }
        });
      }
    } catch (e) {}

    // Absences — distinct AGENT-DAYS. real > late > approved priority per day.
    // Every type tracks an offshore count so e.g. "UNAB 12 (9 Off)" is direct.
    try {
      var absRows = rowsOf('WF_ABSENCES');
      if (absRows.length > 1) {
        var apprRgx = (typeof APPROVED_LEAVE_RGX !== 'undefined') ? APPROVED_LEAVE_RGX : /\b(asclu|slu|furlough|acsu)\b/i;
        var lateRgx = (typeof LATE_RGX !== 'undefined') ? LATE_RGX : /\balu\b/i;
        var dayInfo = {};  // agent|date → { dayMid, region, real:{}, late, appr }
        absRows.slice(1).forEach(function(row) {
          var dStr = WT._formatDate(row[1]);
          var dayMs = self._dayStartEpoch(dStr);
          if (dayMs < 0 || dayMs > nowMs) return;
          var dayMid = dayMs + 43200000;
          var inSel = dayMid >= pb.selStart && dayMid < pb.selEnd;
          var inPrev = dayMid >= pb.prevStart && dayMid < pb.prevEnd;
          if (!inSel && !inPrev) return;
          var agent = String(row[0]).trim();
          var type = String(row[2]).trim() || 'OTHER';
          var k = agent + '|' + dStr;
          if (!dayInfo[k]) {
            dayInfo[k] = { dayMid: dayMid, inSel: inSel,
                           region: String(row[5] || '').indexOf('Offshore') !== -1 ? 'Offshore' : 'Onshore',
                           real: {}, late: false, appr: false,
                           hrs: 0, shift: { Night: 0, Morning: 0, Evening: 0 } };
          }
          if (apprRgx.test(type)) dayInfo[k].appr = true;
          else if (lateRgx.test(type)) dayInfo[k].late = true;
          else {
            dayInfo[k].real[type] = true;
            // Real-absence HOURS, split across the Night/Morning/Evening shift
            // windows (07-15 / 15-23 / 23-07) so concentration-by-shift is exact.
            var sMin = WT._timeToMins(row[3]), eMin = WT._timeToMins(row[4]);
            if (sMin >= 0 && eMin >= 0 && eMin !== sMin) {
              if (eMin <= sMin) eMin += 1440;
              WT._getShiftSplits(sMin, eMin).forEach(function(sp) {
                dayInfo[k].hrs += sp.hours;
                if (dayInfo[k].shift[sp.shift] != null) dayInfo[k].shift[sp.shift] += sp.hours;
              });
            }
          }
        });
        Object.keys(dayInfo).forEach(function(k) {
          var di = dayInfo[k];
          var t = di.inSel ? selT : prevT;
          var realTypes = Object.keys(di.real);
          if (realTypes.length) {
            t.absences += 1;
            if (di.region === 'Offshore') t.absOff += 1; else t.absOn += 1;
            t.absHours += di.hrs;
            if (di.region === 'Offshore') t.absHoursOff += di.hrs;
            t.absShift.Night += di.shift.Night; t.absShift.Morning += di.shift.Morning; t.absShift.Evening += di.shift.Evening;
            if (di.shift.Night > 0) t.absShiftCount.Night++; if (di.shift.Morning > 0) t.absShiftCount.Morning++; if (di.shift.Evening > 0) t.absShiftCount.Evening++;
            realTypes.forEach(function(ty) {
              t.absTypes[ty] = (t.absTypes[ty] || 0) + 1;
              if (di.region === 'Offshore') t.absTypesOff[ty] = (t.absTypesOff[ty] || 0) + 1;
            });
          } else if (di.late) {
            t.lates += 1;
          } else if (di.appr) {
            t.approvedLeave += 1;
          }
          if (di.inSel) {
            var si = (grain === 'day') ? 0 : self._windowIndex(wins, di.dayMid);
            if (si === -1) return;
            var srs = absSeries[si];
            if (realTypes.length) {
              srs.absences += 1;
              if (di.region === 'Offshore') srs.absOff += 1; else srs.absOn += 1;
              srs.absHours += di.hrs;
              srs.absShift.Night += di.shift.Night; srs.absShift.Morning += di.shift.Morning; srs.absShift.Evening += di.shift.Evening;
              if (di.shift.Night > 0) srs.absShiftCount.Night++; if (di.shift.Morning > 0) srs.absShiftCount.Morning++; if (di.shift.Evening > 0) srs.absShiftCount.Evening++;
              realTypes.forEach(function(ty) {
                srs.absTypes[ty] = (srs.absTypes[ty] || 0) + 1;
                if (di.region === 'Offshore') srs.absTypesOff[ty] = (srs.absTypesOff[ty] || 0) + 1;
              });
            } else if (di.late) {
              srs.lates += 1;
            } else if (di.appr) {
              srs.approvedLeave += 1;
            }
          }
        });
      }
    } catch (e) {}

    // Service level — per-bucket averages + sel/prev period averages
    try {
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stats History');
      if (sheet && sheet.getLastRow() > 1) {
        var sums = wins.map(function() { return { svl: 0, ack: 0, n: 0 }; });
        var selS = { svl: 0, ack: 0, n: 0 }, prevS = { svl: 0, ack: 0, n: 0 };
        sheet.getDataRange().getValues().slice(1).forEach(function(r) {
          var t = new Date(r[0]).getTime();
          if (isNaN(t)) return;
          var svl = parseFloat(r[1]) || 0;
          if (svl > 0 && svl <= 1) svl = svl * 100;
          if (svl <= 0) return;
          var ack = parseFloat(String(r[2]).replace(/[^\d.]/g, '')) || 0;
          if (t >= pb.selStart && t < pb.selEnd) {
            selS.svl += svl; selS.ack += ack; selS.n++;
            var wi = self._windowIndex(wins, t);
            if (wi !== -1) { sums[wi].svl += svl; sums[wi].ack += ack; sums[wi].n++; }
          } else if (t >= pb.prevStart && t < pb.prevEnd) {
            prevS.svl += svl; prevS.ack += ack; prevS.n++;
          }
        });
        sums.forEach(function(s2, i) {
          if (s2.n > 0) { buckets[i].slAvg = Math.round((s2.svl / s2.n) * 10) / 10; buckets[i].ackAvg = Math.round(s2.ack / s2.n); }
        });
        if (selS.n > 0) { selT.slAvg = Math.round((selS.svl / selS.n) * 10) / 10; selT.ackAvg = Math.round(selS.ack / selS.n); }
        if (prevS.n > 0) { prevT.slAvg = Math.round((prevS.svl / prevS.n) * 10) / 10; prevT.ackAvg = Math.round(prevS.ack / prevS.n); }
      }
    } catch (e) {}

    // ── Shrinkage inputs + Stacking-lunch concurrency ──────────────────────────
    // One pass over the full schedule (Schedule_History wins over Raw Schedule):
    //  • scheduled hours per bucket (the shrinkage denominator),
    //  • break/lunch hours per bucket (shrinkage numerator pieces),
    //  • a 15-min intraday concurrency curve (agents on lunch/break at once),
    //    split Onshore vs Offshore, for the stacking-lunch chart.
    var NIV = 96;                                   // 96 × 15-min buckets in a day
    var lunchAgg = { onL: [], offL: [], onB: [], offB: [] };
    for (var _i = 0; _i < NIV; _i++) { lunchAgg.onL[_i] = 0; lunchAgg.offL[_i] = 0; lunchAgg.onB[_i] = 0; lunchAgg.offB[_i] = 0; }
    var lunchDays = {};
    try {
      var schedSeen = {};
      var addConc = function(sMin, eMin, region, isLunch) {
        if (eMin <= sMin) eMin += 1440;
        var arr = region === 'Offshore' ? (isLunch ? lunchAgg.offL : lunchAgg.offB) : (isLunch ? lunchAgg.onL : lunchAgg.onB);
        for (var m = sMin; m < eMin; m += 15) { arr[Math.floor((m % 1440) / 15)] += 1; }
      };
      ['Schedule_History', 'Raw Schedule'].forEach(function(shName) {
        var rows = rowsOf(shName);
        if (rows.length < 2) return;
        rows.slice(1).forEach(function(row) {
          var nm = String(row[0]).trim(); if (!nm) return;
          var dStr = WT._formatDate(row[2]); if (!dStr) return;
          var dayMs = self._dayStartEpoch(dStr); if (dayMs < 0) return;
          if (dayMs + 86400000 * 2 < pb.prevStart || dayMs >= pb.selEnd) return;          // window gate
          var sk = nm.toLowerCase() + '|' + dStr; if (schedSeen[sk]) return; schedSeen[sk] = true;  // history wins
          var region = String(row[6] || '').indexOf('Offshore') !== -1 ? 'Offshore' : 'Onshore';
          // Scheduled hours from epochs (locale-proof); fall back to text shift cols.
          var msS = Number(String(row[10]).replace(/[,\s]/g, '')), msE = Number(String(row[11]).replace(/[,\s]/g, ''));
          if (row[10] !== '' && row[11] !== '' && !isNaN(msS) && !isNaN(msE) && msS > 0 && msE > msS && (msE - msS) <= 86400000) {
            distribute({ s: msS, e: msE }, function(t, h) { t.schedH += h; }, function(b, h) { b.schedH += h; });
          } else {
            var siv = self._segInterval(WT, dayMs, row[3], row[4]);
            if (siv) distribute(siv, function(t, h) { t.schedH += h; }, function(b, h) { b.schedH += h; });
          }
          if (dayMs >= pb.selStart && dayMs < pb.selEnd) lunchDays[dStr] = true;
          var brks; try { brks = JSON.parse(row[7] || '[]'); } catch (e2) { brks = []; }
          brks.forEach(function(bk) {
            var sMin = WT._timeToMins(bk.start), eMin = WT._timeToMins(bk.end);
            if (eMin === sMin) return;
            var isLunch = /lunch|repas|meal/i.test(String(bk.type || ''));
            var isBreak = /break|pause/i.test(String(bk.type || ''));
            if (!isLunch && !isBreak) return;
            var em = eMin <= sMin ? eMin + 1440 : eMin;
            var iv = { s: dayMs + sMin * 60000, e: dayMs + em * 60000 };
            distribute(iv, function(t, h) { if (isLunch) t.lunchH += h; else t.breakH += h; },
                           function(b, h) { if (isLunch) b.lunchH += h; else b.breakH += h; });
            if (dayMs >= pb.selStart && dayMs < pb.selEnd) addConc(sMin, eMin, region, isLunch);
          });
        });
      });
    } catch (e) {}
    // Average the intraday curve into a "typical day" across the window's days.
    var nLunchDays = Math.max(1, Object.keys(lunchDays).length);
    ['onL', 'offL', 'onB', 'offB'].forEach(function(kk) { lunchAgg[kk] = lunchAgg[kk].map(function(v) { return Math.round((v / nLunchDays) * 10) / 10; }); });
    // Intraday SVL/ACK overlay (averaged per 15-min interval-of-day across the window).
    var lunchSvl = [], lunchAck = [];
    for (var _j = 0; _j < NIV; _j++) { lunchSvl[_j] = null; lunchAck[_j] = null; }
    try {
      var sh2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stats History');
      if (sh2 && sh2.getLastRow() > 1) {
        var accS = [], accA = [], accN = [];
        for (var _k = 0; _k < NIV; _k++) { accS[_k] = 0; accA[_k] = 0; accN[_k] = 0; }
        sh2.getDataRange().getValues().slice(1).forEach(function(r) {
          var t = new Date(r[0]).getTime(); if (isNaN(t) || t < pb.selStart || t >= pb.selEnd) return;
          var svl = parseFloat(r[1]) || 0; if (svl > 0 && svl <= 1) svl = svl * 100; if (svl <= 0) return;
          var ack = parseFloat(String(r[2]).replace(/[^\d.]/g, '')) || 0;
          var dd = new Date(t); var idx = Math.floor((dd.getHours() * 60 + dd.getMinutes()) / 15);
          if (idx < 0 || idx >= NIV) return;
          accS[idx] += svl; accA[idx] += ack; accN[idx] += 1;
        });
        for (var _m = 0; _m < NIV; _m++) { if (accN[_m] > 0) { lunchSvl[_m] = Math.round((accS[_m] / accN[_m]) * 10) / 10; lunchAck[_m] = Math.round(accA[_m] / accN[_m]); } }
      }
    } catch (e) {}
    var peakOf = function(arr) { var mx = 0, mi = -1; arr.forEach(function(v, i) { if (v > mx) { mx = v; mi = i; } }); return { val: mx, label: mi < 0 ? '' : (('0' + Math.floor(mi / 4)).slice(-2) + ':' + ('0' + ((mi % 4) * 15)).slice(-2)) }; };
    var lunchLabels = []; for (var _n = 0; _n < NIV; _n++) { lunchLabels[_n] = ('0' + Math.floor(_n / 4)).slice(-2) + ':' + ('0' + ((_n % 4) * 15)).slice(-2); }
    var lunchObj = { labels: lunchLabels, onLunch: lunchAgg.onL, offLunch: lunchAgg.offL, onBreak: lunchAgg.onB, offBreak: lunchAgg.offB,
                     svl: lunchSvl, ack: lunchAck, peakOn: peakOf(lunchAgg.onL), peakOff: peakOf(lunchAgg.offL), days: Object.keys(lunchDays).length };

    var round1 = function(v) { return Math.round(v * 10) / 10; };
    var HOUR_KEYS = ['ot', 'otX1', 'otX15', 'safe', 'tower', 'coach', 'acsu', 'elearn', 'openOt', 'openOtToDate', 'idpDeficit', 'preCodedOt', 'reading', 'breakH', 'lunchH', 'schedH'];
    var roundShift = function(s) { if (!s) return; s.Night = round1(s.Night); s.Morning = round1(s.Morning); s.Evening = round1(s.Evening); };
    var finIdp = function(o) {
      o.idpNet = o._idpN > 0 ? round1(o._idpSum / o._idpN) : null;
      o.idpReq = o._idpN > 0 ? round1(o._idpReq / o._idpN) : null;   // avg required seats
      o.idpOpen = o._idpN > 0 ? round1(o._idpOpen / o._idpN) : null; // avg staffed/open seats
      delete o._idpSum; delete o._idpN; delete o._idpReq; delete o._idpOpen;
    };
    buckets.forEach(function(b) { HOUR_KEYS.forEach(function(k2) { b[k2] = round1(b[k2]); }); finIdp(b); });
    [selT, prevT].forEach(function(t) { HOUR_KEYS.forEach(function(k2) { t[k2] = round1(t[k2]); }); t.absHours = round1(t.absHours); t.absHoursOff = round1(t.absHoursOff); roundShift(t.absShift); finIdp(t); });
    absSeries.forEach(function(a) { a.absHours = round1(a.absHours); roundShift(a.absShift); });
    // Codes derived directly from category totals, then round everything.
    selT.codes.ACSU = selT.acsu; selT.codes.OT = selT.ot;
    Object.keys(selT.codes).forEach(function(c) { selT.codes[c] = round1(selT.codes[c]); });

    // Imported reports (authoritative weekly activity hours + monthly alarms).
    // When the Activity-loading report covers the selected period, its hours
    // OVERRIDE the schedule-derived codes (so LFQI/CLASSROOM/VILT/etc. show real).
    var activityImported = null, alarmsImported = null, forecastImported = null, safeAgentsImported = null;
    try {
      if (typeof ReportImport !== 'undefined') {
        var _startStr = this._fmt(pb.selStart, 'yyyy-MM-dd');
        var _endStr = this._fmt(pb.selEnd - 86400000, 'yyyy-MM-dd');
        var ai = ReportImport.getActivityCodes(_startStr, _endStr);
        if (ai && ai.has) {
          activityImported = ai;
          var CODEMAP = { CE: 'CE-HUDDLE', QUAL: 'QUAL', ONE: 'ONE', SBYS: 'SBYS', READ: 'READ', LFQI: 'LFQI', WOFQT: 'WOFQT', VILT: 'VILT', CLASSROOM: 'CLASSROOM', TEAM: 'MEET' };
          Object.keys(ai.byCode).forEach(function(c) { if (ai.byCode[c] > 0) selT.codes[CODEMAP[c] || c] = ai.byCode[c]; });
        }
        var al = ReportImport.getAlarms(_startStr, _endStr);
        if (al && al.has) alarmsImported = al;
        // Forecast: daily rows for day/week views, monthly rows for month+ views.
        var fGrain = (grain === 'month' || grain === 'quarter' || grain === 'ytd') ? 'month' : 'day';
        var fc = ReportImport.getForecast(fGrain, _startStr, _endStr);
        if (fc && fc.has) forecastImported = fc;
        var sa = ReportImport.getSafeAgents(_startStr, _endStr);
        if (sa && sa.has) safeAgentsImported = sa;
      }
    } catch (e) {}

    var topList = function(map) {
      return Object.keys(map)
        .map(function(n) { return { name: n, hours: round1(map[n]) }; })
        .filter(function(x) { return x.hours > 0; })
        .sort(function(a, b) { return b.hours - a.hours; })
        .slice(0, 5);
    };
    var topCoachList = Object.keys(topCoach)
      .map(function(n) { return { name: n, sessions: topCoach[n].sessions, hours: round1(topCoach[n].hours) }; })
      .filter(function(x) { return x.sessions > 0 || x.hours > 0; })
      .sort(function(a, b) { return (b.hours - a.hours) || (b.sessions - a.sessions); })
      .slice(0, 5);

    var periodLabel;
    if (grain === 'day') periodLabel = this._fmt(pb.selStart, 'EEE MMM d, yyyy');
    else if (grain === 'month') periodLabel = this._fmt(pb.selStart, 'MMMM yyyy');
    else if (grain === 'quarter') periodLabel = 'Q' + (Math.floor(new Date(pb.selStart).getMonth() / 3) + 1) + ' ' + new Date(pb.selStart).getFullYear();
    else if (grain === 'ytd') periodLabel = 'YTD ' + new Date(pb.selStart).getFullYear() + ' · Jan 1 – ' + this._fmt(Math.min(pb.selEnd, nowMs) - 86400000, 'MMM d');
    else periodLabel = this._fmt(pb.selStart, 'MMM d') + ' – ' + this._fmt(pb.selEnd - 86400000, 'MMM d');

    var __payload = JSON.stringify({
      grain: grain,
      refDate: refDateStr,
      generated: Utilities.formatDate(new Date(), this.TZ, 'yyyy-MM-dd HH:mm'),
      buckets: buckets,
      abs: absSeries,
      lunch: lunchObj,
      activity: activityImported,
      alarms: alarmsImported,
      forecast: forecastImported,
      safeAgents: safeAgentsImported,
      totals: { sel: selT, prev: prevT },
      sel: {
        label: periodLabel,
        startStr: this._fmt(pb.selStart, 'yyyy-MM-dd'),
        endStr: this._fmt((grain === 'ytd' ? Math.min(pb.selEnd, nowMs) : pb.selEnd) - 86400000, 'yyyy-MM-dd'),
        topOt: topList(topOt),
        topSafe: topList(topSafe),
        topTower: topList(topTower),
        topCoach: topCoachList,
        otAgents: topList(topOt).length ? Object.keys(topOt).filter(function(n) { return topOt[n] > 0; }).length : 0
      }
    });
    // _cache is non-null only for fully-past periods (set above), so this writes
    // a long-lived cache for history and never for the live period.
    if (_cache) { try { var _z = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(__payload)).getBytes()); if (_z.length < 99000) _cache.put(_ck, _z, 21600); } catch (e) {} }
    return __payload;
  }
};
