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

  // SAFE capability lives in the MasterList Skills text as the skill
  // "SmartWear" (the SAFE program's skill name). Tolerant of spacing/hyphen.
  _canSafe: function (skills) { return /smart[\s-]*wear/i.test(String(skills || '')); },

  // MasterList Skills are a SharePoint multi-value cell: "Name;#Lvl;#Name;#Lvl…".
  // Parse it into a readable list and pull the SmartWear proficiency level.
  _parseSkills: function (raw) {
    var s = String(raw || '').trim();
    if (!s) return { pretty: '', swLevel: null };
    var list = [];
    if (s.indexOf(';#') !== -1) {
      var parts = s.split(';#');
      for (var i = 0; i < parts.length; i += 2) {
        var nm = String(parts[i] || '').trim(); if (!nm) continue;
        var lv = (i + 1 < parts.length) ? String(parts[i + 1]).trim() : '';
        list.push({ name: nm, lvl: /^\d+$/.test(lv) ? parseInt(lv, 10) : null });
      }
    } else {
      s.split(/[;,]/).forEach(function (t) { t = t.trim(); if (t) list.push({ name: t, lvl: null }); });
    }
    var swLevel = null;
    list.forEach(function (x) { if (/smart[\s-]*wear/i.test(x.name) && x.lvl != null) swLevel = x.lvl; });
    var pretty = list.map(function (x) { return x.lvl != null ? x.name + ' (' + x.lvl + ')' : x.name; }).join(' · ');
    return { pretty: pretty, swLevel: swLevel };
  },

  getAnalytics: function (mode, refDate, regionFilter, cycleFilter) {
    regionFilter = regionFilter || 'All';
    cycleFilter = cycleFilter || 'ALL';
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null;
    if (!WT) return JSON.stringify({ error: 'Engine unavailable.' });

    var self = this;
    // Cache the whole analytics payload per (mode·date·region·cycle), invalidated by
    // a version stamp the import bumps. Recomputing the full read + 14-month trend on
    // every view/mode-switch is what made a full-year DB time out; now it's computed
    // once per slice. If the first compute exceeds the client watchdog, the server
    // still finishes and caches, so a retry is instant.
    var _ver = ''; try { _ver = PropertiesService.getScriptProperties().getProperty('WF_CACHE_VER') || ''; } catch (e) {}
    var _ck = 'safeAn|' + mode + '|' + refDate + '|' + regionFilter + '|' + (cycleFilter || 'ALL') + '|' + _ver;
    var _cache = null; try { _cache = CacheService.getScriptCache(); } catch (e) {}
    if (_cache) { try { var _hit = _cache.get(_ck); if (_hit) return Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(_hit), 'application/x-gzip', 'c.gz')).getDataAsString(); } catch (e) {} }
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
    var firstSafeEpoch = {};   // normKey -> earliest SAFE-delivery epoch (capability-onset proxy)
    var addTrend = function (agent, dStr, sMins, eRaw, src) {
      var mo = dStr.substring(0, 7);
      if (trendIdx[mo] === undefined) return;
      var fk = self._normKey(agent), fep = new Date(dStr + 'T12:00:00').getTime();
      if (firstSafeEpoch[fk] === undefined || fep < firstSafeEpoch[fk]) firstSafeEpoch[fk] = fep;
      var td = src + '|' + agent + '|' + dStr + '|' + sMins + '|' + eRaw;
      if (seenTrend[td]) return; seenTrend[td] = true;
      var em = eRaw <= sMins ? eRaw + 1440 : eRaw;
      if (!trendByAgent[agent]) trendByAgent[agent] = {};
      trendByAgent[agent][mo] = (trendByAgent[agent][mo] || 0) + (em - sMins) / 60;
    };

    // ── MasterList profiles ──
    var langMap = self._loadLangMap();
    var mlByKey = {};
    var dbML = WT._getDB('WF_MASTERLIST');
    if (dbML && dbML.getLastRow() > 1) {
      dbML.getDataRange().getDisplayValues().slice(1).forEach(function (r) {
        var nm = String(r[0]).trim();
        if (!nm) return;
        var key = (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey(nm) : nm.toLowerCase();
        var parsed = self._parseSkills(r[3]);
        mlByKey[key] = { name: nm, level: parseInt(r[1], 10) || 2, sup: String(r[2] || '').trim(),
                         skills: parsed.pretty, swLevel: parsed.swLevel, lang: langMap[key] || self._inferLang(r[3]), canSafe: self._canSafe(r[3]) };
      });
    }
    var profileOf = function (agent) {
      var key = (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey(agent) : String(agent).toLowerCase();
      return mlByKey[key] || { name: agent, level: 2, sup: '', skills: '', swLevel: null, lang: 'EN', canSafe: false };
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
                            srcSched: 0, srcOt: 0, otByDay: {}, days: {}, segs: 0 };
      }
      var a = agents[g.agent];
      a.total += h; a.segs++;
      if (g.shift === 'Morning') a.morning += h;
      else if (g.shift === 'Evening') a.evening += h;
      else a.night += h;
      if (g.src === 'SCHED') a.srcSched += h; else { a.srcOt += h; a.otByDay[g.rawDay] = self._r2((a.otByDay[g.rawDay] || 0) + h); }
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
    var totals = { all: 0, morning: 0, evening: 0, night: 0, sched: 0, ot: 0, otShift: 0, otIncr: 0, count: events.length,
                   bRed: 0, bWarn: 0, bNormal: 0, bLow: 0 };

    // ── Capability + hour-of-day COVERAGE (the management proof) ──────────
    // Capability is real data: an agent CAN do SAFE iff their MasterList skill
    // text contains "Smart wear". Offshore never does SAFE, so it's excluded.
    // We prove that high-SAFE agents work the hours when few capable agents are
    // staffed, by comparing SAFE demand to capable supply per clock hour.
    var winStartStr = Utilities.formatDate(new Date(bounds.start), 'America/Toronto', 'yyyy-MM-dd');
    var winEndStr = eStr;
    var nk = function (n) { return self._normKey(n); };
    var mk24 = function () { var a = []; for (var i = 0; i < 24; i++) a.push(0); return a; };
    var addHourly = function (arr, sMin, eMin) {
      var t = sMin;
      while (t < eMin) {
        var hod = Math.floor((t % 1440) / 60);
        var chunkEnd = Math.min(eMin, (Math.floor(t / 60) + 1) * 60);
        arr[hod] += (chunkEnd - t);
        t = chunkEnd;
      }
    };

    // SAFE-capable, onshore roster (the only people who CAN carry SAFE).
    var capable = {};                       // normKey -> {name, level, lang, region}
    Object.keys(mlByKey).forEach(function (k) {
      var m = mlByKey[k]; if (!m.canSafe) return;
      if (regionOf(m.name) === 'Offshore') return;
      capable[k] = { name: m.name, level: m.level, lang: m.lang };
    });

    // DEMAND + distinct providers + per-agent hour profile (activity — reliable).
    var safeHourly = mk24();
    var providerSets = []; for (var ps = 0; ps < 24; ps++) providerSets.push({});
    var byHourByAgent = {};                 // display name -> [24] minutes of SAFE
    var deliverBands = {};                  // delivering normKey -> {Morning,Evening,Night}
    events.forEach(function (e) {
      var sMin = e.startMin, eMin = e.startMin + (e.dur || 0) * 60;
      addHourly(safeHourly, sMin, eMin);
      var ek = nk(e.agent);
      var t = sMin;
      while (t < eMin) { providerSets[Math.floor((t % 1440) / 60)][ek] = true; t = (Math.floor(t / 60) + 1) * 60; }
      if (!byHourByAgent[e.agent]) byHourByAgent[e.agent] = mk24();
      addHourly(byHourByAgent[e.agent], sMin, eMin);
      (deliverBands[ek] = deliverBands[ek] || {})[e.shift] = true;
    });
    var providersHourly = providerSets.map(function (s) { return Object.keys(s).length; });

    // SUPPLY: capable agents on shift per clock hour (schedule — rotation-aware,
    // because it only counts capable agents who ACTUALLY have a schedule row).
    var capHourMin = mk24(), schedDaySet = {}, shiftBands = {}, schedSeen = {}, scheduledCapable = {};
    var agentShift = {};   // normKey -> { "startMin|endMin": count } → most common shift window
    var schedDayByAgent = {};   // normKey|date -> true: agent had a real scheduled shift that day
    var capDay = {};            // date -> { all:{}, Morning:{}, Evening:{}, Night:{} } distinct capable agents scheduled that day
    var readSched = function (sheet) {
      var db = WT._getDB(sheet);
      if (!db || db.getLastRow() < 2) return;
      db.getDataRange().getDisplayValues().slice(1).forEach(function (row) {
        var nm = String(row[0]).trim(); if (!nm) return;
        var k = nk(nm); if (!capable[k]) return;           // only SAFE-capable agents
        var dStr = WT._formatDate(row[2]); if (!dStr || dStr < winStartStr || dStr > winEndStr) return;
        if ((mode === 'month' || mode === 'quarter') && cycleFilter !== 'ALL') {   // respect Week A/B rotation
          if (WT._getCycleForEpoch(new Date(dStr + 'T12:00:00').getTime()) !== cycleFilter) return;
        }
        // Epoch-first shift envelope (same robust path as the board). Reading the
        // TEXT shift columns via _timeToMins made every coerced "12/30/1899"/"9 h 00"
        // cell parse to 0→0, so capable agents were counted on-shift 00:00–24:00 and
        // the "capable on shift" line pinned flat at the headcount (~20, "maxed out").
        var sh = self._shiftFromRow(row); if (!sh) return;   // need a real shift window
        var ss = sh.start, se = sh.end;                      // se may exceed 1440 (overnight)
        var dk = k + '|' + dStr; if (schedSeen[dk]) return; schedSeen[dk] = true;   // history wins
        schedDaySet[dStr] = true;
        scheduledCapable[k] = true;                          // actually worked the window
        schedDayByAgent[k + '|' + dStr] = true;              // this agent was scheduled this day
        var asm = agentShift[k] = agentShift[k] || {};       // tally this shift window
        var spk = (ss % 1440) + '|' + se; asm[spk] = (asm[spk] || 0) + 1;
        addHourly(capHourMin, ss, se);
        var sb = shiftBands[k] = shiftBands[k] || {};
        var splits = WT._getShiftSplits(ss, se);
        splits.forEach(function (sp) { sb[sp.shift] = true; });
        // Time-aware per-day capable: count an agent only on/after the date they FIRST
        // actually did SAFE. The current MasterList tag must NOT be applied retroactively
        // to dates before they were a SAFE agent (fixes the historical over-count).
        if (firstSafeEpoch[k] !== undefined && firstSafeEpoch[k] <= new Date(dStr + 'T12:00:00').getTime()) {
          var cd = capDay[dStr] = capDay[dStr] || {};   // normKey -> minutes-equivalent per band
          var rec = cd[k] = cd[k] || { Morning: 0, Evening: 0, Night: 0 };
          splits.forEach(function (sp) { rec[sp.shift] += (sp.hours || 0); });
        }
      });
    };
    readSched('Schedule_History');
    readSched('Raw Schedule');
    var numSchedDays = Object.keys(schedDaySet).length;
    var hasSchedule = numSchedDays > 0;
    // TRUE per-day capable staffing (NOT averaged): each day, the distinct capable
    // agents actually scheduled, split by shift. This is the scarcity proof — a day
    // with only one capable name on the clock shows 1, with that name listed.
    var capableByDay = Object.keys(capDay).sort().map(function (d) {
      var cd = capDay[d], keys = Object.keys(cd), m = 0, e = 0, n = 0, names = [];
      keys.forEach(function (k) {
        var r = cd[k];   // count each agent ONCE, in the band where they're scheduled most
        if (r.Morning >= r.Evening && r.Morning >= r.Night) m++;
        else if (r.Evening >= r.Night) e++;
        else n++;
        names.push((capable[k] || {}).name || k);
      });
      return { date: d, count: keys.length, morning: m, evening: e, night: n, names: names.sort() };
    });
    // Day-level scarcity: a day is "thin" when its distinct capable-scheduled count
    // sits at/below half the median day (same rule the hourly view uses). This lets
    // the UI surface ONLY the dangerous days instead of every date in the window.
    var dayThinThreshold = 0;
    (function () {
      var dc = capableByDay.filter(function (x) { return x.count > 0; }).map(function (x) { return x.count; });
      if (!dc.length) return;
      var s = dc.slice().sort(function (p, q) { return p - q; }), mid = Math.floor(s.length / 2);
      var med = s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
      dayThinThreshold = self._r2(0.5 * med);
      capableByDay.forEach(function (x) { x.thin = x.count > 0 && x.count <= dayThinThreshold; });
    })();
    // avg concurrent capable headcount per clock hour = agent-minutes / 60 / days.
    var capableOnShiftHourly = capHourMin.map(function (m) { return hasSchedule ? self._r2(m / 60 / numSchedDays) : 0; });

    // coverage basis = scheduled capable headcount if we have it, else distinct providers.
    var coverageBasis = hasSchedule ? capableOnShiftHourly : providersHourly.slice();
    var median = function (arr) {
      if (!arr || !arr.length) return null;
      var s = arr.slice().sort(function (p, q) { return p - q; });
      var m = Math.floor(s.length / 2);
      return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
    };
    var demandCov = [];
    for (var dc = 0; dc < 24; dc++) if (safeHourly[dc] > 0) demandCov.push(coverageBasis[dc]);
    var medCov = median(demandCov);
    var thinThreshold = medCov != null ? self._r2(0.5 * medCov) : 0;
    var thinHours = [], thinSet = {};
    for (var th = 0; th < 24; th++) if (safeHourly[th] > 0 && coverageBasis[th] <= thinThreshold) { thinHours.push(th); thinSet[th] = true; }

    // capable roster by shift band ever worked (schedule, else delivery as proxy).
    var bandsSource = hasSchedule ? shiftBands : deliverBands;
    var byShift = { morning: 0, evening: 0, night: 0 };
    Object.keys(bandsSource).forEach(function (k) {
      var b = bandsSource[k];
      if (b.Morning) byShift.morning++;
      if (b.Evening) byShift.evening++;
      if (b.Night) byShift.night++;
    });

    // Most common scheduled shift window per agent (for the thin-hour carrier list).
    var shiftModeOf = function (key) {
      var m = agentShift[key]; if (!m) return null;
      var best = null, bc = 0;
      Object.keys(m).forEach(function (pk) { if (m[pk] > bc) { bc = m[pk]; best = pk; } });
      if (!best) return null;
      var p = best.split('|'); return { start: parseInt(p[0], 10), end: parseInt(p[1], 10) };
    };
    var winEndStr = Utilities.formatDate(new Date(bounds.end), 'America/Toronto', 'yyyy-MM-dd');
    // SAFE HOURS now come from the imported SAFE report (ACTIVE TIME) for this
    // period, not the schedule, when a report covers it. The schedule still drives
    // coverage / Compare / the timeline — only the per-agent hours TOTAL is swapped.
    var safeMap = (typeof ReportImport !== 'undefined' && ReportImport.getSafeForPeriod) ? ReportImport.getSafeForPeriod(winStartStr, winEndStr) : { has: false, map: {} };
    var perAgent = Object.keys(agents).map(function (n) {
      var a = agents[n];
      var schedH = self._r2(a.total);
      var totalH = safeMap.has ? (safeMap.map[nk(n)] != null ? Math.round(safeMap.map[nk(n)] * 100) / 100 : 0) : schedH;
      var ash = shiftModeOf(nk(n));
      // Split SAFE-via-OT: hours on a day the agent was scheduled = extending their
      // shift; hours on a non-scheduled day = an OT agent brought in to do SAFE
      // (incremental capacity, not their normal shift).
      var _ak = nk(n), otShiftH = 0, otIncrH = 0;
      Object.keys(a.otByDay || {}).forEach(function (d) { if (schedDayByAgent[_ak + '|' + d]) otShiftH += a.otByDay[d]; else otIncrH += a.otByDay[d]; });
      otShiftH = self._r2(otShiftH); otIncrH = self._r2(otIncrH);
      totals.all += totalH; totals.morning += a.morning; totals.evening += a.evening; totals.night += a.night;
      totals.sched += a.srcSched; totals.ot += a.srcOt;
      totals.otShift += otShiftH; totals.otIncr += otIncrH;
      var dayKeys = Object.keys(a.days);
      var maxDay = null;
      dayKeys.forEach(function (d) { if (!maxDay || a.days[d] > a.days[maxDay]) maxDay = d; });
      var nDouble = 0, nOverlap = 0;
      Object.keys(doubleDays).forEach(function (k) { if (k.indexOf(n + '|') === 0) nDouble++; });
      Object.keys(overlapDays).forEach(function (k) { if (k.indexOf(n + '|') === 0) nOverlap++; });

      var pf = profileOf(n);
      var monthEq = (mode === 'month') ? totalH : self._r2(totalH / winDays * 30.42);
      var band = bandOf(monthEq);
      if (band === 'RED') totals.bRed++; else if (band === 'WARN') totals.bWarn++; else if (band === 'LOW') totals.bLow++; else totals.bNormal++;

      // Trend array (14 months) — pure data; the chart speaks for itself.
      var tb = trendByAgent[n] || {};
      var trend = trendMonths.map(function (m) { return { key: m.key, label: m.label, year: m.year, hours: self._r2(tb[m.key] || 0) }; });
      var peak = trend.reduce(function (mx, x) { return x.hours > mx.hours ? x : mx; }, { hours: 0, key: '' });

      // Hour-of-day SAFE for this agent + how much falls in thin-coverage hours.
      var bh = byHourByAgent[n] || mk24();
      var byHour = bh.map(function (m) { return self._r2(m / 60); });
      var thinMin = 0, allMin = 0;
      for (var hx = 0; hx < 24; hx++) { allMin += bh[hx]; if (thinSet[hx]) thinMin += bh[hx]; }
      var thinShare = allMin > 0 ? Math.round(thinMin / allMin * 100) : 0;

      return {
        name: n, region: a.region,
        level: pf.level, lang: pf.lang, skills: pf.skills, sup: pf.sup, canSafe: !!pf.canSafe, swLevel: (pf.swLevel != null ? pf.swLevel : null),
        band: band, monthEq: monthEq, safeFromReport: safeMap.has, schedHrs: schedH,
        total: totalH, morning: self._r2(a.morning), evening: self._r2(a.evening), night: self._r2(a.night),
        srcSched: self._r2(a.srcSched), srcOt: self._r2(a.srcOt), srcOtShift: otShiftH, srcOtIncr: otIncrH,
        days: dayKeys.length, segs: a.segs,
        byHour: byHour, thinShare: thinShare,
        shiftStartStr: ash ? WT._minsToTime(ash.start) : '', shiftEndStr: ash ? WT._minsToTime(ash.end) : '',
        shiftPattern: ash ? (Math.floor(ash.start / 60) + '–' + Math.floor((ash.end % 1440) / 60)) : '',
        avgPerDay: dayKeys.length ? self._r2(totalH / dayKeys.length) : 0,
        maxDay: maxDay ? { date: maxDay, hours: a.days[maxDay] } : null,
        dayMap: a.days,
        doubleDays: nDouble, overlapDays: nOverlap,
        trend: trend, trendPeak: { month: peak.key, hours: peak.hours }
      };
    }).sort(function (x, y) { return y.total - x.total; });

    ['all', 'morning', 'evening', 'night', 'sched', 'ot'].forEach(function (k) { totals[k] = self._r2(totals[k]); });
    events.sort(function (a, b) { return a.date.localeCompare(b.date) || a.agent.localeCompare(b.agent) || a.src.localeCompare(b.src); });

    // Headline proof: % of SAFE delivered in thin-coverage hours + scarcest hour.
    var totalSafeH = safeHourly.reduce(function (s, m) { return s + m; }, 0) / 60;
    var thinSafeH = thinHours.reduce(function (s, h) { return s + safeHourly[h]; }, 0) / 60;
    var thinPct = totalSafeH > 0 ? Math.round(thinSafeH / totalSafeH * 100) : 0;
    var minCovHour = null, minCov = Infinity;
    for (var mh = 0; mh < 24; mh++) if (safeHourly[mh] > 0 && coverageBasis[mh] < minCov) { minCov = coverageBasis[mh]; minCovHour = mh; }

    // Data-health: SAFE delivered by agents not flagged capable / not in MasterList.
    var dnc = {}, unm = {};
    events.forEach(function (e) {
      var k = nk(e.agent);
      if (!mlByKey[k]) unm[e.agent] = true; else if (!capable[k]) dnc[e.agent] = true;
    });

    var __payload = JSON.stringify({
      mode: mode, trackerType: 'safe', label: bounds.label, cycle: bounds.cycle,
      grid: [], events: events, totals: totals, perAgent: perAgent,
      safeFromReport: safeMap.has, safeReportLabel: safeMap.label || '',
      trendMonths: trendMonths.map(function (m) { return m.label + (m.label === 'Jan' ? " '" + String(m.year).slice(-2) : ''); }),
      hasMasterList: Object.keys(mlByKey).length > 0,
      winDays: winDays,
      coverage: {
        hasSchedule: hasSchedule, schedDays: numSchedDays,
        safeHourly: safeHourly.map(function (m) { return self._r2(m / 60); }),
        providersHourly: providersHourly, capableOnShiftHourly: capableOnShiftHourly,
        thinHours: thinHours, thinThreshold: thinThreshold,
        capable: { total: Object.keys(scheduledCapable).length, trainedTotal: Object.keys(capable).length, byShift: byShift },
        byDay: capableByDay, dayThinThreshold: dayThinThreshold,
        headline: { thinPct: thinPct, thinSafeH: self._r2(thinSafeH), totalSafeH: self._r2(totalSafeH),
                    thinHourCount: thinHours.length, minCovHour: minCovHour }
      },
      dataHealth: { deliveredNotCapable: Object.keys(dnc).slice(0, 25), unmatchedMasterList: Object.keys(unm).slice(0, 25) },
      audit: { agents: perAgent.length, doubleAgentDays: Object.keys(doubleDays).length,
               overlapAgentDays: Object.keys(overlapDays).length }
    });
    if (_cache) { try { var _z = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(__payload)).getBytes()); if (_z.length < 99000) _cache.put(_ck, _z, 21600); } catch (e) {} }
    return __payload;
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
  // Like _normKey but drops single-letter tokens (middle initials) so
  // "Wong Lawrence J" and "Wong Lawrence" resolve to the same person.
  _coreKey: function (n) { return this._normKey(n).split(' ').filter(function (t) { return t.length > 1; }).join(' '); },
  _regionOf: function (nm) {
    if (typeof RegionRegistry !== 'undefined') { var r = RegionRegistry.getRegion(nm); if (r) return r; }
    return 'Onshore';
  },

  // Parse a schedule cell into minutes-since-midnight, or -1 when it is NOT a
  // real time-of-day. Deliberately stricter than WT._timeToMins, which maps
  // date serials and unparseable junk to 0 — that's what was rendering every
  // shift as a bogus 00:00–00:00. Handles 12h ("9:00 AM"), coerced 12h with
  // seconds ("5:00:00 p.m.", meridian anywhere), 24h ("17:00"), the French-
  // Canadian "13 h 30" form that a fr-CA spreadsheet returns from
  // getDisplayValues(), day-fractions (0.375) and Date objects; rejects date
  // serials, bare integers, date strings and empties.
  _safeTime: function (raw) {
    if (raw == null) return -1;
    if (raw instanceof Date) return raw.getHours() * 60 + raw.getMinutes();
    var s = String(raw).trim(); if (!s) return -1;
    var num = Number(s);
    if (!isNaN(num) && s.indexOf(':') === -1) { return (num > 0 && num < 1.5) ? Math.round(num * 1440) % 1440 : -1; }
    // French-Canadian 24h clock: "13 h 30", "9 h", "11h00".
    var fr = s.match(/^(\d{1,2})\s*h\s*(\d{2})?$/i);
    if (fr) { var fh = parseInt(fr[1], 10), fm = fr[2] ? parseInt(fr[2], 10) : 0; return (fh > 23 || fm > 59) ? -1 : fh * 60 + fm; }
    var m = s.match(/(\d{1,2})[:.](\d{2})/); if (!m) return -1;   // grab HH:MM (ignore trailing :SS)
    var h = parseInt(m[1], 10), mm = parseInt(m[2], 10);
    var ap = /p\.?\s*m/i.test(s) ? 'PM' : (/a\.?\s*m/i.test(s) ? 'AM' : null);   // meridian anywhere
    if (ap === 'PM' && h < 12) h += 12; if (ap === 'AM' && h === 12) h = 0;
    if (h > 23 || mm > 59) return -1;
    return h * 60 + mm;
  },

  // Minutes-of-day (America/Toronto) for a Unix-ms epoch. Locale-proof.
  _epochToMin: function (ms) {
    var hm = Utilities.formatDate(new Date(Number(ms)), 'America/Toronto', 'HH:mm');
    var p = hm.split(':'); return parseInt(p[0], 10) * 60 + parseInt(p[1], 10);
  },

  // The shift window for a schedule row as { start, end } minutes (end may
  // exceed 1440 for an overnight shift), or null when the row has no real
  // shift. PREFERS the StartEpoch/EndEpoch columns (10,11) — the same
  // unambiguous values the floor view uses — because the "Shift Start/End"
  // text columns are often time-values that a fr-CA sheet formats as
  // date-only ("12/30/1899"), losing the time entirely. Falls back to parsing
  // the text columns when epochs are absent (older archive rows).
  _shiftFromRow: function (row) {
    var rs = row[10], re = row[11];
    var msS = Number(String(rs).replace(/[,\s]/g, '')), msE = Number(String(re).replace(/[,\s]/g, ''));   // tolerate "1,781,…" display
    if (rs !== '' && rs != null && re !== '' && re != null && !isNaN(msS) && !isNaN(msE) && msS > 0 && msE > msS) {
      var dur = Math.round((msE - msS) / 60000);
      if (dur > 0 && dur <= 1440) { var sm = this._epochToMin(msS); return { start: sm, end: sm + dur }; }
    }
    var ss = this._safeTime(row[3]), se = this._safeTime(row[4]);
    if (ss >= 0 && se >= 0) return { start: ss % 1440, end: (se <= ss ? se + 1440 : se) };
    return null;
  },

  // Manual language overrides (skills rarely encode language). Stored in
  // WF_LANG_MAP [Agent Key, Display Name, Lang] so they persist across imports.
  _loadLangMap: function () {
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null; var map = {};
    if (!WT) return map;
    var db = WT._getDB('WF_LANG_MAP');
    if (db && db.getLastRow() > 1) {
      db.getDataRange().getDisplayValues().slice(1).forEach(function (r) {
        var k = String(r[0] || '').trim(); var lang = String(r[2] || '').trim().toUpperCase();
        if (k && lang) map[k] = lang;
      });
    }
    return map;
  },
  setAgentLang: function (name, lang) {
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null; if (!WT) return 'Error';
    name = String(name || '').trim(); lang = String(lang || '').trim().toUpperCase();
    if (!name) return 'Error';
    if (['EN', 'FR', 'BL'].indexOf(lang) === -1) lang = 'EN';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('WF_LANG_MAP');
    if (!sh) { sh = ss.insertSheet('WF_LANG_MAP'); sh.appendRow(['Agent Key', 'Display Name', 'Lang']); }
    var key = this._normKey(name);
    var last = sh.getLastRow();
    var rows = last > 1 ? sh.getRange(2, 1, last - 1, 3).getValues() : [];
    for (var i = 0; i < rows.length; i++) { if (String(rows[i][0]) === key) { sh.getRange(i + 2, 2, 1, 2).setValues([[name, lang]]); return 'OK'; } }
    sh.appendRow([key, name, lang]); return 'OK';
  },
  // Distinct dates that actually have a full shift envelope (for the Compare
  // "jump to latest schedule" hint). Newest first, capped.
  getScheduleDates: function () {
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null; if (!WT) return '[]';
    var set = {};
    ['Schedule_History', 'Raw Schedule'].forEach(function (sh) {
      var db = WT._getDB(sh); if (!db || db.getLastRow() < 2) return;
      db.getDataRange().getDisplayValues().slice(1).forEach(function (r) {
        if (!String(r[3] || '').trim() || !String(r[4] || '').trim()) return;
        var d = WT._formatDate(r[2]); if (d) set[d] = true;
      });
    });
    return JSON.stringify(Object.keys(set).sort().reverse().slice(0, 120));
  },

  // Onshore SmartWear-capable agents who actually have a REAL shift on `dateStr`
  // overlapping the [hStart,hEnd] minute window (hEnd may exceed 1440 for an
  // overnight focus). This is what "Pin capable" uses so it only pins people
  // who are genuinely working that date/shift — never someone from another week
  // who merely did SAFE elsewhere in the period (which is what produced the
  // misleading "activities only" rows).
  getCapableScheduledForDate: function (dateStr, hStart, hEnd) {
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null; if (!WT || !dateStr) return '[]';
    var self = this;
    hStart = (hStart == null) ? 0 : Number(hStart); hEnd = (hEnd == null) ? 1440 : Number(hEnd);
    var fullDay = (hStart <= 0 && hEnd >= 1440);
    var cap = {}; var langMap = self._loadLangMap();
    var dbML = WT._getDB('WF_MASTERLIST');
    if (dbML && dbML.getLastRow() > 1) {
      dbML.getDataRange().getDisplayValues().slice(1).forEach(function (r) {
        var nm = String(r[0]).trim(); if (!nm) return;
        if (!self._canSafe(r[3])) return; if (self._regionOf(nm) === 'Offshore') return;
        cap[self._normKey(nm)] = { name: nm, level: parseInt(r[1], 10) || 2, lang: langMap[self._normKey(nm)] || self._inferLang(r[3]) };
      });
    }
    var overlaps = function (s, e) {
      if (fullDay) return true;
      var wins = (hEnd <= 1440) ? [[hStart, hEnd]] : [[hStart, 1440], [0, hEnd - 1440]];
      var segs = []; var b1 = Math.min(e, 1440); if (b1 > s) segs.push([s, b1]); if (e > 1440) segs.push([0, e - 1440]);
      return segs.some(function (sg) { return wins.some(function (w) { return sg[1] > w[0] && sg[0] < w[1]; }); });
    };
    var seen = {}, out = [];
    var read = function (sheet) {
      var rs = WT._getDB(sheet); if (!rs || rs.getLastRow() < 2) return;
      rs.getDataRange().getDisplayValues().slice(1).forEach(function (row) {
        var nm = String(row[0]).trim(); if (!nm) return;
        if (WT._formatDate(row[2]) !== dateStr) return;
        var k = self._normKey(nm); var c = cap[k]; if (!c || seen[k]) return;
        // "Scheduled" = the Coverage rule: a shift envelope exists in either the
        // text columns or the epoch columns. Don't require a PARSEABLE text time
        // (fr-CA cells display as "12/30/1899"); the epoch path / lenient parse
        // recovers the real window so a scheduled capable agent is never dropped.
        var ssRaw = String(row[3] || '').trim(), seRaw = String(row[4] || '').trim();
        var hasEpoch = String(row[10] || '').trim() && String(row[11] || '').trim();
        if (!ssRaw && !seRaw && !hasEpoch) return;
        var sh = self._shiftFromRow(row);                    // epoch-first, accurate
        if (!sh) {                                           // last resort: lenient parse (parity with Coverage)
          var ls = WT._timeToMins(ssRaw), le = WT._timeToMins(seRaw);
          sh = { start: ls % 1440, end: (le <= ls ? le + 1440 : le) };
        }
        if (!fullDay && !overlaps(sh.start % 1440, sh.end)) return;
        seen[k] = true;
        out.push({ name: c.name, level: c.level, lang: c.lang, shiftStartStr: WT._minsToTime(sh.start % 1440), shiftEndStr: WT._minsToTime(sh.end) });
      });
    };
    read('Schedule_History'); read('Raw Schedule');
    out.sort(function (a, b) { return (b.level - a.level) || a.name.localeCompare(b.name); });
    return JSON.stringify(out);
  },

  // Roster for the Compare "Add agent" search — ONLY onshore SmartWear-capable
  // agents (offshore never does SAFE; non-capable agents can't carry it).
  getScheduleRoster: function () {
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null;
    if (!WT) return '[]';
    var self = this;
    var langMap = self._loadLangMap();
    var out = [];
    var dbML = WT._getDB('WF_MASTERLIST');
    if (dbML && dbML.getLastRow() > 1) {
      dbML.getDataRange().getDisplayValues().slice(1).forEach(function (r) {
        var nm = String(r[0]).trim(); if (!nm) return;
        if (!self._canSafe(r[3])) return;                  // must have SmartWear
        if (self._regionOf(nm) === 'Offshore') return;     // offshore never does SAFE
        var parsed = self._parseSkills(r[3]);
        out.push({ name: nm, level: parseInt(r[1], 10) || 2, lang: langMap[self._normKey(nm)] || self._inferLang(r[3]), swLevel: parsed.swLevel });
      });
    }
    out.sort(function (a, b) { return a.name.localeCompare(b.name); });
    return JSON.stringify(out);
  },

  getScheduleBoard: function (dateStr, agentsPipe) {
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null;
    if (!WT) return JSON.stringify({ error: 'Engine unavailable.' });
    var self = this;
    var wanted = String(agentsPipe || '').split('|').map(function (x) { return x.trim(); }).filter(Boolean);
    if (!wanted.length) return JSON.stringify({ date: dateStr, agents: [] });
    var wantKey = {}, wantCore = {};
    wanted.forEach(function (n) { wantKey[self._normKey(n)] = n; var c = self._coreKey(n); if (c && !wantCore[c]) wantCore[c] = n; });
    // Resolve a sheet row name to a requested agent — exact key first, then a
    // looser core-key (ignores middle initials / token order) so name variants join.
    var resolve = function (nm) { return wantKey[self._normKey(nm)] || wantCore[self._coreKey(nm)] || null; };
    var schedNamesForDate = {};   // distinct schedule names present on this date (any agent)

    var toMin = function (t) { return self._safeTime(t); };
    var byAgent = {};
    var shiftDiag = [];   // first few raw shift cells, surfaced in diag for transparency
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
        if (WT._formatDate(row[2]) !== dateStr) return;
        var ssR = String(row[3] || '').trim(), seR = String(row[4] || '').trim();
        if (ssR && seR) schedNamesForDate[self._normKey(nm)] = true;   // someone scheduled this date
        var want = resolve(nm); if (!want) return;
        var k = self._normKey(want);
        if (byAgent[k] && byAgent[k].hasFull) return; // already covered (history wins over Raw)
        var ag = ensure(want);
        ag.hasFull = true; anyFull = true;
        ag.region = String(row[6] || '').trim();
        // Empty start/end ("Off"/absent days) must NOT become a 00:00–00:00
        // envelope — only set an envelope when both times are present.
        var sh = self._shiftFromRow(row);
        if (shiftDiag.length < 4) shiftDiag.push({ name: want, src: sheetName, rawStart: String(row[3] || ''), rawEnd: String(row[4] || ''), epochS: String(row[10] || ''), epochE: String(row[11] || ''), parsedStart: sh ? sh.start : -1, parsedEnd: sh ? sh.end : -1 });
        if (sh) { ag.shiftStart = sh.start; ag.shiftEnd = sh.end; }
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
        var nm = String(row[0]).trim(); if (!nm) return;
        var want = resolve(nm); if (!want) return;
        var k = self._normKey(want);
        if (byAgent[k] && byAgent[k].hasFull) return; // already full from Raw Schedule
        if (WT._formatDate(row[1]) !== dateStr) return;
        var t = mapper(row);
        pushSeg(ensure(want), t[0], t[1], row[sIdx], row[eIdx]);
      });
    };
    pull('WF_ROLES', function (r) { var v = String(r[2]).toUpperCase(); return v.indexOf('SAFE') !== -1 ? ['SAFE', 'SAFE'] : (v.indexOf('ICL') !== -1 ? ['ICL', 'ICL'] : ((v.indexOf('ULC') !== -1 || v.indexOf('FIRE') !== -1) ? ['ULC', 'ULC FIRE'] : ['TOWER', 'Tower'])); }, 3, 4);
    pull('WF_COACHING', function () { return ['COACH', 'Coaching']; }, 3, 4);
    pull('WF_FURLOUGH', function () { return ['ACSU', 'ACSU']; }, 3, 4);
    pull('WF_OVERTIME', function (r) { return ['OT', 'OT ' + String(r[4] || '')]; }, 6, 7);
    pull('WF_ABSENCES', function (r) { return ['ABS', String(r[2] || 'Absence')]; }, 3, 4);

    // attach level/lang from MasterList
    var ml = {}; var lmO = self._loadLangMap();
    var dbML = WT._getDB('WF_MASTERLIST');
    if (dbML && dbML.getLastRow() > 1) {
      dbML.getDataRange().getDisplayValues().slice(1).forEach(function (r) {
        var nm = String(r[0]).trim(); if (!nm) return;
        ml[self._normKey(nm)] = { level: parseInt(r[1], 10) || 2, lang: lmO[self._normKey(nm)] || self._inferLang(r[3]) };
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
    var matchedFull = result.filter(function (a) { return a.hasFull; }).length;
    return JSON.stringify({ date: dateStr, agents: result, archived: anyFull, hasArchive: hasArchive,
                            hasMasterList: Object.keys(ml).length > 0,
                            diag: { schedRowsForDate: Object.keys(schedNamesForDate).length, matchedFull: matchedFull, requested: wanted.length, shiftSamples: shiftDiag } });
  },

  // RANGE version of the board for the Compare "Week" view: each agent's full
  // schedule (shift + breaks/lunch + activities) for every date in [startStr,endStr].
  // Returns agents[].days keyed by date. Range is capped to 31 days.
  getScheduleRange: function (startStr, endStr, agentsPipe) {
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null;
    if (!WT) return JSON.stringify({ error: 'Engine unavailable.' });
    var self = this;
    var wanted = String(agentsPipe || '').split('|').map(function (x) { return x.trim(); }).filter(Boolean);
    if (!wanted.length || !startStr || !endStr) return JSON.stringify({ start: startStr, end: endStr, dates: [], agents: [] });
    if (endStr < startStr) { var t0 = startStr; startStr = endStr; endStr = t0; }
    // build the list of dates (cap 31)
    var dates = [], dCur = new Date(startStr + 'T12:00:00'), dEnd = new Date(endStr + 'T12:00:00'), guard = 0;
    while (dCur <= dEnd && guard < 31) { dates.push(Utilities.formatDate(dCur, 'America/Toronto', 'yyyy-MM-dd')); dCur.setDate(dCur.getDate() + 1); guard++; }
    var endCap = dates[dates.length - 1] || endStr;

    var wantKey = {}, wantCore = {};
    wanted.forEach(function (n) { wantKey[self._normKey(n)] = n; var c = self._coreKey(n); if (c && !wantCore[c]) wantCore[c] = n; });
    var resolve = function (nm) { return wantKey[self._normKey(nm)] || wantCore[self._coreKey(nm)] || null; };
    var toMin = function (t) { return self._safeTime(t); };
    var byAD = {};   // normKey -> { dateStr -> dayObj }
    var dayOf = function (want, dStr) {
      var k = self._normKey(want); byAD[k] = byAD[k] || {};
      if (!byAD[k][dStr]) byAD[k][dStr] = { shiftStart: null, shiftEnd: null, hasFull: false, segments: [], absent: '' };
      return byAD[k][dStr];
    };
    var pushSeg = function (day, type, label, sRaw, eRaw) {
      var sm = toMin(sRaw), em = toMin(eRaw); if (sm < 0 || em < 0) return; if (em <= sm) em += 1440;
      day.segments.push({ type: type, label: label, startMin: sm % 1440, endMin: em, time: WT._minsToTime(sm) + ' - ' + WT._minsToTime(em) });
    };
    var BRK = { 'LUNCH': ['LUNCH', 'Lunch'], 'BREAK': ['BREAK', 'Break'], 'TRAINING': ['COACH', 'Training'],
                'ACSU': ['ACSU', 'ACSU'], 'SAFE': ['SAFE', 'SAFE'], 'ICL': ['ICL', 'ICL'], 'ULC FIRE': ['ULC', 'ULC FIRE'] };
    var readFull = function (sheetName) {
      var rs = WT._getDB(sheetName); if (!rs || rs.getLastRow() < 2) return;
      rs.getDataRange().getDisplayValues().slice(1).forEach(function (row) {
        var nm = String(row[0]).trim(); if (!nm) return;
        var dStr = WT._formatDate(row[2]); if (!dStr || dStr < startStr || dStr > endCap) return;
        var want = resolve(nm); if (!want) return;
        var day = dayOf(want, dStr); if (day.hasFull) return;       // history wins over Raw
        day.hasFull = true;
        var ssRaw = String(row[3] || '').trim(), seRaw = String(row[4] || '').trim();
        var shR = self._shiftFromRow(row); if (shR) { day.shiftStart = shR.start; day.shiftEnd = shR.end; }
        try { JSON.parse(row[7] || '[]').forEach(function (b) { var m = BRK[String(b.type || '').toUpperCase()] || ['OTHER', String(b.type || 'Activity')]; pushSeg(day, m[0], m[1], b.start, b.end); }); } catch (e) {}
        if (String(row[9] || '').trim()) day.absent = String(row[9]).trim();
      });
    };
    readFull('Schedule_History'); readFull('Raw Schedule');
    var pull = function (sheet, mapper, sIdx, eIdx) {
      var db = WT._getDB(sheet); if (!db || db.getLastRow() < 2) return;
      db.getDataRange().getDisplayValues().slice(1).forEach(function (row) {
        var nm = String(row[0]).trim(); if (!nm) return;
        var want = resolve(nm); if (!want) return;
        var dStr = WT._formatDate(row[1]); if (!dStr || dStr < startStr || dStr > endCap) return;
        var day = dayOf(want, dStr); if (day.hasFull) return;
        var t = mapper(row); pushSeg(day, t[0], t[1], row[sIdx], row[eIdx]);
      });
    };
    pull('WF_ROLES', function (r) { var v = String(r[2]).toUpperCase(); return v.indexOf('SAFE') !== -1 ? ['SAFE', 'SAFE'] : (v.indexOf('ICL') !== -1 ? ['ICL', 'ICL'] : ((v.indexOf('ULC') !== -1 || v.indexOf('FIRE') !== -1) ? ['ULC', 'ULC FIRE'] : ['TOWER', 'Tower'])); }, 3, 4);
    pull('WF_COACHING', function () { return ['COACH', 'Coaching']; }, 3, 4);
    pull('WF_FURLOUGH', function () { return ['ACSU', 'ACSU']; }, 3, 4);
    pull('WF_OVERTIME', function (r) { return ['OT', 'OT ' + String(r[4] || '')]; }, 6, 7);
    pull('WF_ABSENCES', function (r) { return ['ABS', String(r[2] || 'Absence')]; }, 3, 4);

    var ml = {}; var lmO = self._loadLangMap();
    var dbML = WT._getDB('WF_MASTERLIST');
    if (dbML && dbML.getLastRow() > 1) {
      dbML.getDataRange().getDisplayValues().slice(1).forEach(function (r) {
        var nm = String(r[0]).trim(); if (!nm) return;
        ml[self._normKey(nm)] = { level: parseInt(r[1], 10) || 2, lang: lmO[self._normKey(nm)] || self._inferLang(r[3]) };
      });
    }
    var agentsOut = wanted.map(function (n) {
      var k = self._normKey(n); var p = ml[k] || { level: 2, lang: 'EN' };
      var daysMap = byAD[k] || {}; var total = 0;
      var days = dates.map(function (dStr) {
        var day = daysMap[dStr] || { shiftStart: null, shiftEnd: null, hasFull: false, segments: [], absent: '' };
        day.segments.sort(function (a, b) { return a.startMin - b.startMin; });
        if (day.shiftStart == null && day.segments.length) {
          day.shiftStart = day.segments[0].startMin;
          day.shiftEnd = day.segments.reduce(function (mx, s2) { return Math.max(mx, s2.endMin); }, 0);
        }
        day.date = dStr;
        day.shiftStartStr = day.shiftStart != null ? WT._minsToTime(day.shiftStart) : '';
        day.shiftEndStr = day.shiftEnd != null ? WT._minsToTime(day.shiftEnd) : '';
        day.noData = !day.hasFull && !(day.segments && day.segments.length);
        day.safeH = day.segments.reduce(function (s, x) { return s + (x.type === 'SAFE' ? (x.endMin - x.startMin) / 60 : 0); }, 0);
        total += day.safeH;
        return day;
      });
      return { name: n, level: p.level, lang: p.lang, days: days, safeTotal: self._r2(total) };
    });
    return JSON.stringify({ start: startStr, end: endCap, dates: dates, agents: agentsOut, hasMasterList: Object.keys(ml).length > 0 });
  }
};
