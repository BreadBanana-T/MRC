/**
 * MODULE: REPORT IMPORT
 *
 * Parses two pasted IEX/BI exports (tab-delimited, as copied from Excel) and
 * stores them in their own sheets so the Management View can show real numbers:
 *
 *   1. Activity loading hours — weekly hours per activity CODE, grouped by
 *      Exception Grp (Vacation/Absence/Project/Meeting/Coaching/Training/
 *      PC-Sys/Break/Lunch). Week columns are week-start Sundays. → WF_ACTIVITY_WK
 *   2. Alarms by service type — monthly Alarm Volume + AHT by priority/service
 *      type. → WF_ALARMS
 *
 * Both are full-report replaces (each paste overwrites the sheet).
 */
var ReportImport = {
  TZ: 'America/Toronto',

  _ss: function() { return SpreadsheetApp.getActiveSpreadsheet(); },

  // Overwrite a sheet with header + rows, clearing any leftover tail.
  _write: function(name, headers, rows) {
    var ss = this._ss();
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    var nCols = headers.length;
    var all = [headers].concat(rows).map(function(r) {
      r = r.slice(0, nCols);
      while (r.length < nCols) r.push('');
      return r;
    });
    var prevRows = sh.getLastRow(), prevCols = sh.getLastColumn();
    sh.getRange(1, 1, all.length, nCols).setValues(all);
    if (prevRows > all.length) sh.getRange(all.length + 1, 1, prevRows - all.length, Math.max(nCols, prevCols)).clearContent();
    if (prevCols > nCols) sh.getRange(1, nCols + 1, Math.max(all.length, prevRows), prevCols - nCols).clearContent();
    return rows.length;
  },

  // "Jun 21 - Sun" / "May 31" → yyyy-MM-dd of the most recent matching date
  // (so weeks bind correctly across a year boundary).
  _resolveSunday: function(label) {
    var m = String(label).match(/([A-Za-z]{3,})\s+(\d{1,2})/);
    if (!m) return '';
    var MO = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
    var mo = MO[m[1].substring(0, 3).toLowerCase()];
    if (mo == null) return '';
    var day = parseInt(m[2], 10);
    var now = new Date();
    var cand = new Date(now.getFullYear(), mo, day);
    if (cand.getTime() > now.getTime() + 10 * 86400000) cand = new Date(now.getFullYear() - 1, mo, day);
    return Utilities.formatDate(cand, this.TZ, 'yyyy-MM-dd');
  },

  // Accepts "January 2026", "Jan-26", "Jan 26", "Jan-2026", "2026-01", etc.
  _monthKey: function(lbl) {
    var s = String(lbl).trim();
    var iso = s.match(/^(20\d\d)[-\/](\d{1,2})$/);              // 2026-01
    if (iso) { var mmi = ('0' + iso[2]).slice(-2); return iso[1] + '-' + mmi; }
    var m = s.match(/([A-Za-z]{3,})\s*[-\/ ]\s*(\d{2,4})/);     // Jan-26 / January 2026
    if (!m) return '';
    var MO = { jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06', jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12' };
    var mm = MO[m[1].substring(0, 3).toLowerCase()];
    if (!mm) return '';
    var yr = m[2]; if (yr.length === 2) yr = '20' + yr;
    return yr + '-' + mm;
  },

  // Leading token of the activity name is the IEX code; Break/Lunch keyed by group.
  _code: function(act, grp) {
    grp = String(grp || '').toLowerCase();
    if (grp.indexOf('break') !== -1) return 'BREAK';
    if (grp.indexOf('lunch') !== -1) return 'LUNCH';
    var m = String(act || '').trim().match(/^([A-Za-z0-9\/]+)/);
    return m ? m[1].toUpperCase().replace(/[^A-Z0-9]/g, '') : '';
  },

  // Does this paste look like the activity-loading report?
  looksLikeActivity: function(t) {
    t = String(t || '').toLowerCase();
    return t.indexOf('exception grp') !== -1 || (t.indexOf('activity name') !== -1 && t.indexOf('week of ref') !== -1);
  },
  looksLikeAlarms: function(t) {
    t = String(t || '').toLowerCase();
    return t.indexOf('aht secs') !== -1 || t.indexOf('servicetype') !== -1 || (t.indexOf('alarm vol') !== -1 && t.indexOf('servtype') !== -1);
  },
  // Forecast report (weekly "Day of Ref Date" or monthly "Month of Ref Date"):
  // forecast vs actual alarm volume, AHT, SVL%, category mix. Distinct from the
  // alarms-by-service-type report by its "Fcst Alarm"/"Fcst Acc" columns.
  looksLikeForecast: function(t) {
    t = String(t || '').toLowerCase();
    return (t.indexOf('fcst alarm') !== -1 || t.indexOf('fcst acc') !== -1) && t.indexOf('servtype') === -1;
  },
  _num: function(s) { s = String(s == null ? '' : s).replace(/[, %]/g, '').trim(); var n = parseFloat(s); return isFinite(n) ? n : null; },
  // Accepts "21 June 2026", "21-Jun-26", "Jun 21 2026", "2026-06-21", etc.
  _dayDate: function(lbl) {
    var s = String(lbl).trim();
    var MO = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
    var iso = s.match(/^(20\d\d)[-\/](\d{1,2})[-\/](\d{1,2})$/);
    if (iso) return Utilities.formatDate(new Date(+iso[1], +iso[2] - 1, +iso[3]), this.TZ, 'yyyy-MM-dd');
    var day, mo, yr;
    var m = s.match(/(\d{1,2})\s*[-\/ ]\s*([A-Za-z]{3,})\s*[-\/ ]?\s*(\d{2,4})?/);   // 21 June 2026 / 21-Jun-26
    if (m) { day = parseInt(m[1], 10); mo = MO[m[2].substring(0, 3).toLowerCase()]; yr = m[3]; }
    else { var m2 = s.match(/([A-Za-z]{3,})\s*[-\/ ]\s*(\d{1,2})\s*[-\/ ]?\s*(\d{2,4})?/); if (m2) { mo = MO[m2[1].substring(0, 3).toLowerCase()]; day = parseInt(m2[2], 10); yr = m2[3]; } }
    if (mo == null || !day) return '';
    if (yr) { yr = parseInt(yr, 10); if (yr < 100) yr += 2000; return Utilities.formatDate(new Date(yr, mo, day), this.TZ, 'yyyy-MM-dd'); }
    var now = new Date(), cand = new Date(now.getFullYear(), mo, day);   // no year → most recent occurrence ≤ ~today
    if (cand.getTime() > now.getTime() + 10 * 86400000) cand = new Date(now.getFullYear() - 1, mo, day);
    return Utilities.formatDate(cand, this.TZ, 'yyyy-MM-dd');
  },

  importActivity: function(text) {
    try {
      var self = this;
      var lines = String(text || '').split(/\r?\n/).filter(function(l) { return l.trim().length; });
      var hi = -1;
      for (var i = 0; i < lines.length; i++) { if (/activity\s*name/i.test(lines[i])) { hi = i; break; } }
      if (hi < 0) return 'Activity import: header row (Activity Name) not found.';
      var hcols = lines[hi].split('\t').map(function(c) { return c.trim(); });
      var weekCols = [];
      for (var c = 2; c < hcols.length; c++) {
        if (/grand\s*total/i.test(hcols[c])) break;
        if (!hcols[c]) continue;
        weekCols.push({ idx: c, date: self._resolveSunday(hcols[c]) });
      }
      var rows = [];
      for (var j = hi + 1; j < lines.length; j++) {
        var cols = lines[j].split('\t');
        var grp = (cols[0] || '').trim(), act = (cols[1] || '').trim();
        if (!act || /^total$/i.test(act) || /grand\s*total/i.test(grp)) continue;
        var code = self._code(act, grp);
        weekCols.forEach(function(w) {
          if (!w.date) return;
          var h = parseFloat((cols[w.idx] || '').replace(/[, ]/g, ''));
          if (!isFinite(h) || h === 0) return;
          rows.push([w.date, grp, code, act, h]);
        });
      }
      this._write('WF_ACTIVITY_WK', ['WeekStart', 'ExceptionGrp', 'Code', 'ActivityName', 'Hours'], rows);
      return 'Activity loading synced: ' + rows.length + ' rows across ' + weekCols.length + ' week(s).';
    } catch (e) { return 'Activity import error: ' + e.message; }
  },

  importAlarms: function(text) {
    try {
      var lines = String(text || '').split(/\r?\n/).filter(function(l) { return l.trim().length; });
      var hdrRow = -1, monthRow = -1;
      for (var i = 0; i < lines.length; i++) {
        if (hdrRow < 0 && /alarm\s*vol|aht\s*sec/i.test(lines[i])) hdrRow = i;
        if (monthRow < 0 && /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+20\d\d/i.test(lines[i])) monthRow = i;
      }
      if (hdrRow < 0) return 'Alarms import: header (Alarm Vol / AHT) not found.';
      var hdr = lines[hdrRow].split('\t').map(function(c) { return c.trim(); });
      var monthCells = monthRow >= 0 ? lines[monthRow].split('\t') : [];
      var firstVal = -1;
      for (var c = 0; c < hdr.length; c++) { if (/alarm\s*vol/i.test(hdr[c])) { firstVal = c; break; } }
      if (firstVal < 0) firstVal = 4;
      var months = [];
      for (var c2 = firstVal; c2 < hdr.length; c2 += 2) {
        var lbl = (monthCells[c2] || '').trim();
        months.push({ idx: c2, month: this._monthKey(lbl), label: lbl });
      }
      var rows = [];
      for (var j = hdrRow + 1; j < lines.length; j++) {
        var cols = lines[j].split('\t');
        var rank = (cols[0] || '').trim(), grp = (cols[1] || '').trim(), desc = (cols[2] || '').trim(), sid = (cols[3] || '').trim();
        if (/grand\s*total/i.test(rank)) continue;
        if (!sid && !desc) continue;
        months.forEach(function(mn) {
          if (!mn.month || /grand\s*total/i.test(mn.label)) return;
          var vol = parseFloat((cols[mn.idx] || '').replace(/[, ]/g, ''));
          if (!isFinite(vol)) return;
          var aht = parseFloat((cols[mn.idx + 1] || '').replace(/[, ]/g, ''));
          rows.push([mn.month, rank, grp, desc, sid, vol, isFinite(aht) ? aht : '']);
        });
      }
      this._write('WF_ALARMS', ['Month', 'PrioRank', 'PrioGrp', 'Desc', 'ServId', 'Vol', 'AHT'], rows);
      return 'Alarms synced: ' + rows.length + ' rows across ' + months.length + ' month(s).';
    } catch (e) { return 'Alarms import error: ' + e.message; }
  },

  // Forecast vs actuals — daily ("Day of Ref Date") or monthly ("Month of Ref
  // Date"). Stored together in WF_FORECAST, keyed by Grain+Period; importing one
  // grain leaves the other intact.
  FCAST_HEADERS: ['Grain', 'Period', 'FcstAcc', 'FcstAlarm', 'AlarmVol', 'InSvlVol', 'FcstAht', 'ActAht', 'Svl', 'PctHigh', 'PctMed', 'PctLow', 'PctMas', 'PctMix', 'PctOperator'],
  importForecast: function(text) {
    try {
      var self = this;
      var lines = String(text || '').split(/\r?\n/).filter(function(l) { return l.trim().length; });
      var hi = -1;
      for (var i = 0; i < lines.length; i++) { var L = lines[i].toLowerCase(); if (L.indexOf('fcst alarm') !== -1 || L.indexOf('fcst acc') !== -1) { hi = i; break; } }
      if (hi < 0) return 'Forecast import: header (Fcst Alarm) not found.';
      var h0 = lines[hi].split('\t')[0].toLowerCase();
      var grain = h0.indexOf('month') !== -1 ? 'month' : 'day';
      var rows = [];
      for (var j = hi + 1; j < lines.length; j++) {
        var c = lines[j].split('\t');
        var lbl = (c[0] || '').trim();
        if (!lbl || /grand\s*total/i.test(lbl)) continue;
        var period = grain === 'month' ? self._monthKey(lbl) : self._dayDate(lbl);
        if (!period) continue;
        rows.push([grain, period, self._num(c[1]), self._num(c[2]), self._num(c[3]), self._num(c[4]),
                   self._num(c[5]), self._num(c[6]), self._num(c[7]), self._num(c[8]), self._num(c[9]),
                   self._num(c[10]), self._num(c[11]), self._num(c[12]), self._num(c[13])]);
      }
      // Grain-scoped replace: keep rows of the OTHER grain, swap in this grain's.
      var sh = this._ss().getSheetByName('WF_FORECAST');
      var keep = [];
      if (sh && sh.getLastRow() > 1) {
        sh.getDataRange().getDisplayValues().slice(1).forEach(function(r) { if (r[0] && r[0] !== grain) keep.push(r.slice(0, 15)); });
      }
      this._write('WF_FORECAST', this.FCAST_HEADERS, keep.concat(rows));
      return 'Forecast synced: ' + rows.length + ' ' + grain + ' row(s).';
    } catch (e) { return 'Forecast import error: ' + e.message; }
  },

  // Forecast rows for the period. grainWanted 'day' (for week/day views) or
  // 'month' (month/quarter/ytd). startStr/endStr are yyyy-MM-dd.
  getForecast: function(grainWanted, startStr, endStr) {
    var out = { has: false, grain: grainWanted, rows: [], fcstAlarm: 0, alarmVol: 0, accAvg: null };
    try {
      var sh = this._ss().getSheetByName('WF_FORECAST');
      if (!sh || sh.getLastRow() < 2) return out;
      var sKey = grainWanted === 'month' ? startStr.substring(0, 7) : startStr;
      var eKey = grainWanted === 'month' ? endStr.substring(0, 7) : endStr;
      var nf = function(v) { var n = parseFloat(v); return isFinite(n) ? n : null; };
      sh.getDataRange().getDisplayValues().slice(1).forEach(function(r) {
        if (r[0] !== grainWanted) return;
        var p = r[1]; if (!p || p < sKey || p > eKey) return;
        out.rows.push({ period: p, fcstAcc: nf(r[2]), fcstAlarm: nf(r[3]), alarmVol: nf(r[4]), inSvl: nf(r[5]),
                        fcstAht: nf(r[6]), actAht: nf(r[7]), svl: nf(r[8]),
                        pctHigh: nf(r[9]), pctMed: nf(r[10]), pctLow: nf(r[11]), pctMas: nf(r[12]), pctMix: nf(r[13]), pctOperator: nf(r[14]) });
      });
      out.rows.sort(function(a, b) { return a.period < b.period ? -1 : 1; });
      var accSum = 0, accN = 0;
      out.rows.forEach(function(r) { out.fcstAlarm += r.fcstAlarm || 0; out.alarmVol += r.alarmVol || 0; if (r.fcstAcc != null && r.alarmVol) { accSum += r.fcstAcc; accN++; } });
      out.accAvg = accN ? Math.round(accSum / accN) : null;
      out.has = out.rows.length > 0;
    } catch (e) {}
    return out;
  },

  // ── SAFE / SmartWear per-agent export ───────────────────────────────────────
  // ACTIVE TIME = per-agent SAFE hours. Also carries EMPLOYEE ID (TID) and
  // AGENT LANGUAGES, which we fold into WF_LANG_MAP for the SAFE board.
  _hms: function(s) {
    s = String(s == null ? '' : s).trim();
    if (/^\d+:\d{1,2}:\d{1,2}$/.test(s)) { var p = s.split(':'); return (+p[0]) + (+p[1]) / 60 + (+p[2]) / 3600; }
    if (/^\d+:\d{1,2}$/.test(s)) { var q = s.split(':'); return (+q[0]) + (+q[1]) / 60; }
    var n = parseFloat(s.replace(/[, ]/g, '')); return isFinite(n) ? n : 0;
  },
  _langCode: function(s) {
    s = String(s || '').toLowerCase();
    var en = /english|anglais/.test(s) || /\ben\b/.test(s), fr = /french|fran[cç]ais|fran[cç]/.test(s) || /\bfr\b/.test(s);
    if (en && fr) return 'BL'; if (fr) return 'FR'; if (en) return 'EN'; return '';
  },
  looksLikeSafe: function(t) {
    t = String(t || '').toLowerCase();
    return t.indexOf('active time') !== -1 && (t.indexOf('agent languages') !== -1 || t.indexOf('mobile sos') !== -1 || t.indexOf('volume act') !== -1);
  },
  // The SAFE export carries no dates, so the manager tags it with the period it
  // covers (startStr/endStr yyyy-MM-dd + a label) at import time. Rows are stored
  // period-scoped so monthly/weekly views read the matching report.
  importSafe: function(text, startStr, endStr, label) {
    try {
      var self = this;
      startStr = (startStr && String(startStr).length >= 10) ? startStr : Utilities.formatDate(new Date(), this.TZ, 'yyyy-MM-dd');
      endStr = (endStr && String(endStr).length >= 10) ? endStr : startStr;
      label = label || (startStr === endStr ? startStr : (startStr + ' → ' + endStr));
      var lines = String(text || '').split(/\r?\n/).filter(function(l) { return l.trim().length; });
      var hi = -1;
      for (var i = 0; i < lines.length; i++) { var L = lines[i].toLowerCase(); if (L.indexOf('active time') !== -1 && L.indexOf('agent') !== -1) { hi = i; break; } }
      if (hi < 0) return 'SAFE import: header (ACTIVE TIME) not found.';
      var H = lines[hi].split('\t').map(function(c) { return c.trim(); });
      var exact = function(name) { for (var k = 0; k < H.length; k++) if (H[k].toLowerCase() === name) return k; return -1; };
      var incl = function(name) { for (var k = 0; k < H.length; k++) if (H[k].toLowerCase().indexOf(name) !== -1) return k; return -1; };
      var iTid = exact('employee id'); if (iTid < 0) iTid = incl('employee id'); if (iTid < 0) iTid = incl('emp id');
      var iAgent = exact('agent'); if (iAgent < 0) iAgent = (iTid >= 0 ? iTid + 1 : 1);
      var iActive = exact('active time'); if (iActive < 0) iActive = incl('active time');
      var iLang = exact('agent languages'); if (iLang < 0) iLang = incl('languages');
      if (iActive < 0) return 'SAFE import: ACTIVE TIME column not found.';
      var rows = [], langPairs = [];
      for (var j = hi + 1; j < lines.length; j++) {
        var c = lines[j].split('\t');
        var name = (c[iAgent] || '').trim();
        var tid = iTid >= 0 ? (c[iTid] || '').trim() : '';
        if (!name || /^total$/i.test(name) || /^grand\s*total$/i.test(name)) continue;
        var act = self._hms(c[iActive]);
        var lang = iLang >= 0 ? self._langCode(c[iLang]) : '';
        rows.push([startStr, endStr, label, name, tid, lang, Math.round(act * 100) / 100]);
        if (lang) langPairs.push([name, lang]);
      }
      // Period-scoped replace: keep other periods, swap in this one.
      var sh = this._ss().getSheetByName('WF_SAFE_AGENT');
      var keep = [];
      if (sh && sh.getLastRow() > 1) {
        sh.getDataRange().getDisplayValues().slice(1).forEach(function(r) { if (!(r[0] === startStr && r[1] === endStr)) keep.push(r.slice(0, 7)); });
      }
      this._write('WF_SAFE_AGENT', ['PeriodStart', 'PeriodEnd', 'Label', 'Agent', 'TID', 'Lang', 'ActiveHrs'], keep.concat(rows));
      var langN = this._mergeLangMap(langPairs);
      return 'SAFE roster synced: ' + rows.length + ' agents for ' + label + ' · ' + langN + ' languages mapped.';
    } catch (e) { return 'SAFE import error: ' + e.message; }
  },
  // Bulk upsert into WF_LANG_MAP [Agent Key, Display Name, Lang] in ONE write.
  _mergeLangMap: function(pairs) {
    if (!pairs || !pairs.length) return 0;
    try {
      var ss = this._ss();
      var sh = ss.getSheetByName('WF_LANG_MAP') || ss.insertSheet('WF_LANG_MAP');
      var nk = (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey : function(s) { return String(s).trim().toLowerCase(); };
      var vals = sh.getLastRow() > 0 ? sh.getDataRange().getValues() : [];
      var hasHdr = vals.length && String(vals[0][0]).toLowerCase().indexOf('agent key') !== -1;
      var body = hasHdr ? vals.slice(1) : vals;
      var idx = {};
      body.forEach(function(r, i) { idx[String(r[0])] = i; });
      pairs.forEach(function(p) {
        var key = nk(p[0]);
        if (idx[key] != null) { body[idx[key]][1] = p[0]; body[idx[key]][2] = p[1]; }
        else { body.push([key, p[0], p[1]]); idx[key] = body.length - 1; }
      });
      var out = [['Agent Key', 'Display Name', 'Lang']].concat(body.map(function(r) { return [r[0], r[1], r[2]]; }));
      sh.getRange(1, 1, out.length, 3).setValues(out);
      return pairs.length;
    } catch (e) { return 0; }
  },
  // Pick the stored SAFE report period that best OVERLAPS the view window
  // [viewStart, viewEnd] (yyyy-MM-dd). Robust to Wed–Wed vs Sun–Sat week edges.
  // No overlap on a valid view → empty (that period has no SAFE report). When no
  // view bounds are given, fall back to the latest stored period.
  getSafeForPeriod: function(viewStart, viewEnd) {
    var out = { has: false, periodStart: '', periodEnd: '', label: '', agents: [], totalHrs: 0, count: 0, map: {}, idx: [] };
    try {
      var sh = this._ss().getSheetByName('WF_SAFE_AGENT');
      if (!sh || sh.getLastRow() < 2) return out;
      var data = sh.getDataRange().getDisplayValues().slice(1);
      var periods = {};
      data.forEach(function(r) { var ps = r[0], pe = r[1]; if (!ps) return; var k = ps + '|' + pe; if (!periods[k]) periods[k] = { ps: ps, pe: pe || ps, label: r[2], rows: [] }; periods[k].rows.push(r); });
      var dnum = function(s) { return s ? Date.parse(s + 'T00:00:00') : NaN; };
      var haveView = viewStart && String(viewStart).length >= 10;
      var vs = dnum(viewStart), ve = dnum(viewEnd || viewStart) + 86400000;
      var best = null, bestOv = 0, latest = null;
      Object.keys(periods).forEach(function(k) {
        var p = periods[k];
        if (!latest || p.ps > latest.ps) latest = p;
        if (haveView) { var ps = dnum(p.ps), pe = dnum(p.pe) + 86400000; var ov = Math.min(pe, ve) - Math.max(ps, vs); if (ov > bestOv) { bestOv = ov; best = p; } }
      });
      if (!best) { if (haveView) return out; best = latest; }
      if (!best) return out;
      var nk = (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey : function(s) { return String(s).trim().toLowerCase(); };
      best.rows.forEach(function(r) {
        var h = parseFloat(r[6]) || 0; var k = nk(r[3]);
        out.agents.push({ name: r[3], tid: r[4], lang: r[5], hrs: Math.round(h * 10) / 10 });
        out.totalHrs += h; out.map[k] = h;
        out.idx.push({ toks: k.split(' ').filter(function(t) { return t.length > 1; }), hrs: h });
      });
      out.agents.sort(function(a, b) { return b.hrs - a.hrs; });
      out.totalHrs = Math.round(out.totalHrs * 10) / 10;
      out.count = out.agents.length;
      out.periodStart = best.ps; out.periodEnd = best.pe; out.label = best.label;
      out.has = out.count > 0;
    } catch (e) {}
    return out;
  },
  // Resolve a schedule agent name to SAFE hours for a getSafeForPeriod() result.
  // 1) exact token-sorted key; 2) conservative fuzzy: token overlap (handles
  // middle names / initials / extra tokens). Multi-token names need ≥2 shared
  // tokens (first+last); single-token only matches when it's the lone candidate,
  // so we never silently mis-assign hours.
  matchSafeHours: function(name, period) {
    if (!period || !period.has) return null;
    var nk = (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey : function(s) { return String(s).trim().toLowerCase(); };
    var key = nk(name);
    if (period.map[key] != null) return period.map[key];
    var aToks = key.split(' ').filter(function(t) { return t.length > 1; });
    if (!aToks.length) return null;
    var cands = [];
    (period.idx || []).forEach(function(r) {
      var rToks = r.toks; if (!rToks.length) return;
      var shared = 0; aToks.forEach(function(t) { if (rToks.indexOf(t) !== -1) shared++; });
      if (!shared) return;
      var minLen = Math.min(aToks.length, rToks.length);
      var ok = (shared >= 2) || (shared >= 1 && minLen === 1) || (shared === minLen);
      if (ok) cands.push({ shared: shared, diff: Math.abs(aToks.length - rToks.length), hrs: r.hrs });
    });
    if (!cands.length) return null;
    cands.sort(function(a, b) { return (b.shared - a.shared) || (a.diff - b.diff); });
    if (cands[0].shared >= 2) return cands[0].hrs;     // first+last (or more) matched
    return cands.length === 1 ? cands[0].hrs : null;   // single shared token → only if unambiguous
  },
  // Normalized-key → ACTIVE TIME hours for the view period (overlay source).
  getSafeHoursMap: function(viewStart, viewEnd) {
    var p = this.getSafeForPeriod(viewStart, viewEnd);
    return { has: p.has, map: p.map, label: p.label };
  },
  getSafeAgents: function(viewStart, viewEnd) {
    var p = this.getSafeForPeriod(viewStart, viewEnd);
    return { has: p.has, agents: p.agents, totalHrs: p.totalHrs, count: p.count, label: p.label, periodStart: p.periodStart, periodEnd: p.periodEnd };
  },

  // Activity hours per CODE / per GRP within [startStr, endStr] (week-start Sundays).
  getActivityCodes: function(startStr, endStr) {
    var out = { byCode: {}, byGrp: {}, total: 0, weeks: 0, has: false };
    try {
      var sh = this._ss().getSheetByName('WF_ACTIVITY_WK');
      if (!sh || sh.getLastRow() < 2) return out;
      var weeks = {};
      sh.getDataRange().getDisplayValues().slice(1).forEach(function(r) {
        var ws = r[0]; if (!ws || ws < startStr || ws > endStr) return;
        var h = parseFloat(r[4]) || 0;
        out.byCode[r[2]] = (out.byCode[r[2]] || 0) + h;
        out.byGrp[r[1]] = (out.byGrp[r[1]] || 0) + h;
        out.total += h; weeks[ws] = true;
      });
      out.weeks = Object.keys(weeks).length;
      out.has = out.weeks > 0;
      var r1 = function(v) { return Math.round(v * 10) / 10; };
      Object.keys(out.byCode).forEach(function(k) { out.byCode[k] = r1(out.byCode[k]); });
      Object.keys(out.byGrp).forEach(function(k) { out.byGrp[k] = r1(out.byGrp[k]); });
      out.total = r1(out.total);
    } catch (e) {}
    return out;
  },

  // Alarm volume + volume-weighted AHT per service type for months overlapping the period.
  getAlarms: function(startStr, endStr) {
    var out = { rows: [], totalVol: 0, months: [], has: false };
    try {
      var sh = this._ss().getSheetByName('WF_ALARMS');
      if (!sh || sh.getLastRow() < 2) return out;
      var sMonth = startStr.substring(0, 7), eMonth = endStr.substring(0, 7);
      var agg = {}, months = {};
      sh.getDataRange().getDisplayValues().slice(1).forEach(function(r) {
        var mo = r[0]; if (!mo || mo < sMonth || mo > eMonth) return;
        months[mo] = true;
        var key = r[4] || r[3];
        if (!agg[key]) agg[key] = { rank: r[1], grp: r[2], desc: r[3], sid: r[4], vol: 0, ahtSum: 0, ahtW: 0 };
        var vol = parseFloat(r[5]) || 0, aht = parseFloat(r[6]) || 0;
        agg[key].vol += vol;
        if (aht > 0 && vol > 0) { agg[key].ahtSum += aht * vol; agg[key].ahtW += vol; }
        out.totalVol += vol;
      });
      out.months = Object.keys(months).sort();
      out.has = out.months.length > 0;
      out.rows = Object.keys(agg).map(function(k) {
        var a = agg[k];
        return { rank: a.rank, grp: a.grp, desc: a.desc, sid: a.sid, vol: a.vol, aht: a.ahtW > 0 ? Math.round(a.ahtSum / a.ahtW) : null };
      }).sort(function(a, b) { return b.vol - a.vol; });
    } catch (e) {}
    return out;
  }
};
