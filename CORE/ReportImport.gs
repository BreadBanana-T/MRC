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

  _monthKey: function(lbl) {
    var m = String(lbl).match(/([A-Za-z]{3,})\s+(20\d\d)/);
    if (!m) return '';
    var MO = { jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06', jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12' };
    var mm = MO[m[1].substring(0, 3).toLowerCase()];
    return mm ? (m[2] + '-' + mm) : '';
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
  _dayDate: function(lbl) {
    var m = String(lbl).match(/(\d{1,2})\s+([A-Za-z]{3,})\s+(20\d\d)/);
    if (!m) return '';
    var MO = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
    var mo = MO[m[2].substring(0, 3).toLowerCase()];
    if (mo == null) return '';
    return Utilities.formatDate(new Date(parseInt(m[3], 10), mo, parseInt(m[1], 10)), this.TZ, 'yyyy-MM-dd');
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
