/**
 * MODULE: DEMAND SIGNALS — how much staffing we actually need, beyond the IDP.
 *
 * Feeds the break planner (MRC Workforce) a list of explainable uplifts on top
 * of the IEX IDP "Required" curve:
 *
 *   forecast  WF_FORECAST day rows: when actual alarm volume has been running
 *             above (or below) the IEX forecast lately, the IDP — which is built
 *             from that forecast — is off by the same amount.
 *   live      The latest Calculator run (Stats History): alarms vs the IEX trend
 *             over the last 60 min, and SVL vs target. Today only, for the next
 *             ~90 min, fading out.
 *   learned   Conditions (power outages, red weather alerts, high wind,
 *             Alarm.com down) are logged every hour NEXT TO our own alarm trend.
 *             A condition only moves the plan once the log shows it really
 *             raised OUR volume (enough samples, enough days). Outage counts
 *             alone mean little — 55k customers out can bring almost no alarms —
 *             so nothing is assumed.
 *   headsup   A condition that's present but not learned yet (or learned as "no
 *             effect"): shown, never applied.
 *
 *   Writes  WF_CONDITIONS_LOG (hidden): Timestamp | Hour | Dow | Conditions JSON |
 *           SVL | ACK | Trend % | BURG Vol | Fresh (1 = stats were ≤ 45 min old)
 *
 * Run setupDemandTrigger() once from the editor (hourly logger).
 */
var DemandSignals = {
  TZ: 'America/Toronto',
  LOG_SHEET: 'WF_CONDITIONS_LOG',
  LOG_HEADERS: ['Timestamp', 'Hour', 'Dow', 'Conditions', 'SVL', 'ACK', 'Trend %', 'BURG Vol', 'Fresh'],
  LOG_MAX_ROWS: 4000,
  CACHE_KEY: 'MRC_DEMAND_V1',
  CACHE_TTL: 300,
  FRESH_MIN: 45,

  // Mirrors WFEngine.DEFAULT_RULES.demand; an admin's saved planner rules win.
  DEFAULTS: { svlTarget: 90, liveWindowMin: 90, liveCap: 0.30, minSamples: 12, minDays: 3, minUplift: 0.03, forecastDays: 14, forecastMinDays: 5 },

  PROVS: { BC: 'BC', ON: 'Ontario', QC: 'Québec', AB: 'Alberta' },

  _cfg: function () {
    var c = JSON.parse(JSON.stringify(this.DEFAULTS));
    try {
      var r = (typeof BreakPlanner !== 'undefined') ? BreakPlanner.getRules() : null;
      if (r && r.demand) for (var k in r.demand) if (r.demand[k] !== '' && r.demand[k] != null && !isNaN(Number(r.demand[k]))) c[k] = Number(r.demand[k]);
    } catch (e) {}
    return c;
  },

  // ── Conditions right now ──────────────────────────────────────────────────
  // Returns { keys: [...], raw: {...} }. Every key is a plain yes/no bucket so
  // the learning stays explainable: out5k_QC, out25k_QC, alert_ON, wind50_BC,
  // alarmcom_down.
  currentConditions: function () {
    var keys = [], raw = { out: {}, alert: {}, wind: {}, alarmcom: '' };
    try {
      if (typeof OutageTracker !== 'undefined') {
        var o = JSON.parse(OutageTracker.fetchAll() || '{}');
        var bp = (o && o.byProvince) || {};
        Object.keys(bp).forEach(function (p) {
          var d = bp[p] || {};
          if (d.stale) return;                       // only trust fresh numbers
          var n = Number(d.customers) || 0;
          raw.out[p] = n;
          if (n >= 5000) keys.push('out5k_' + p);
          if (n >= 25000) keys.push('out25k_' + p);
        });
      }
    } catch (e) {}
    try {
      if (typeof WeatherService !== 'undefined') {
        var w = WeatherService.fetch() || {};
        (w.alerts || []).forEach(function (a) {
          if (!a || !a.province) return;
          raw.alert[a.province] = a.summary || 'alert';
          keys.push('alert_' + a.province);
        });
        var wx = w.weather || {};
        Object.keys(wx).forEach(function (p) {
          var mx = 0;
          (wx[p] || []).forEach(function (c) { mx = Math.max(mx, parseFloat(c.wind) || 0); });
          raw.wind[p] = mx;
          if (mx >= 50) keys.push('wind50_' + p);
        });
      }
    } catch (e) {}
    try {
      if (typeof AlarmComMonitor !== 'undefined') {
        var s = AlarmComMonitor.getStatus();
        s = (typeof s === 'string') ? JSON.parse(s) : s;
        if (s && !s.disabled) {
          raw.alarmcom = s.state || '';
          if (s.state === 'down' || s.state === 'issue') keys.push('alarmcom_down');
        }
      }
    } catch (e) {}
    return { keys: keys, raw: raw };
  },

  // Wind conditions forecast for a future date (the only condition we can see
  // ahead of time: WeatherService gives daily max wind for the next 3 days).
  forecastConditions: function (dateStr) {
    var keys = [], raw = { wind: {} };
    try {
      if (typeof WeatherService === 'undefined') return { keys: keys, raw: raw };
      var p = dateStr.split('-').map(Number);
      var dn = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][new Date(Date.UTC(p[0], p[1] - 1, p[2], 12)).getUTCDay()];
      var days = Math.round((Date.UTC(p[0], p[1] - 1, p[2]) - Date.UTC.apply(null, this._todayParts())) / 86400000);
      if (days < 1 || days > 3) return { keys: keys, raw: raw };
      var wx = (WeatherService.fetch() || {}).weather || {};
      Object.keys(wx).forEach(function (pr) {
        var mx = 0;
        (wx[pr] || []).forEach(function (c) { (c.forecast || []).forEach(function (f) { if (f.day === dn) mx = Math.max(mx, parseFloat(f.wind) || 0); }); });
        raw.wind[pr] = mx;
        if (mx >= 50) keys.push('wind50_' + pr);
      });
    } catch (e) {}
    return { keys: keys, raw: raw };
  },

  // ── Hourly logger ─────────────────────────────────────────────────────────
  logConditions: function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(this.LOG_SHEET);
    if (!sh) {
      sh = ss.insertSheet(this.LOG_SHEET);
      sh.getRange(1, 1, 1, this.LOG_HEADERS.length).setValues([this.LOG_HEADERS]).setFontWeight('bold');
      sh.setFrozenRows(1);
      try { sh.hideSheet(); } catch (e) {}
    }
    var now = new Date();
    var cond = this.currentConditions();
    var st = this.latestStats(ss);
    var fresh = !!(st && (now.getTime() - st.at) <= this.FRESH_MIN * 60000);
    var hour = parseInt(Utilities.formatDate(now, this.TZ, 'H'), 10);
    var dow = parseInt(Utilities.formatDate(now, this.TZ, 'u'), 10);   // 1 = Mon … 7 = Sun
    sh.appendRow([now, hour, dow, JSON.stringify(cond), st ? st.svl : '', st ? st.ack : '', (st && st.trend != null) ? st.trend : '', (st && st.burg != null) ? st.burg : '', fresh ? 1 : 0]);
    var last = sh.getLastRow();
    if (last - 1 > this.LOG_MAX_ROWS + 100) sh.deleteRows(2, last - 1 - this.LOG_MAX_ROWS);
    try { CacheService.getScriptCache().remove(this.CACHE_KEY + '|' + this._today()); } catch (e) {}
    return { logged: true, conditions: cond.keys, fresh: fresh };
  },

  // Latest Calculator run from Stats History: { at (ms), svl (0-100), ack, trend (%) | null, burg | null }.
  latestStats: function (ss) {
    var sh = (ss || SpreadsheetApp.getActiveSpreadsheet()).getSheetByName('Stats History');
    if (!sh || sh.getLastRow() < 2) return null;
    var width = Math.min(5, Math.max(3, sh.getLastColumn()));
    var r = sh.getRange(sh.getLastRow(), 1, 1, width).getValues()[0];
    var at = (r[0] instanceof Date) ? r[0].getTime() : new Date(r[0]).getTime();
    if (isNaN(at)) return null;
    var svl = parseFloat(r[1]); if (svl > 0 && svl <= 1) svl *= 100;
    var tr = (width >= 4 && r[3] !== '' && r[3] != null && !isNaN(Number(r[3]))) ? Number(r[3]) : null;
    var bv = (width >= 5 && r[4] !== '' && r[4] != null && !isNaN(Number(r[4]))) ? Number(r[4]) : null;
    return { at: at, svl: isNaN(svl) ? null : svl, ack: parseFloat(String(r[2]).replace(/[^\d.]/g, '')) || null, trend: tr, burg: bv };
  },

  // ── Learning ──────────────────────────────────────────────────────────────
  // For every condition seen in the log: mean alarm trend (% over the IEX trend)
  // in hours WITH it vs hours WITHOUT it, over fresh samples only.
  learn: function (ss, cfg) {
    cfg = cfg || this._cfg();
    var sh = (ss || SpreadsheetApp.getActiveSpreadsheet()).getSheetByName(this.LOG_SHEET);
    var out = {};
    if (!sh || sh.getLastRow() < 2) return out;
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, this.LOG_HEADERS.length).getValues();
    var samples = [], self = this;
    rows.forEach(function (r) {
      if (Number(r[8]) !== 1 || r[6] === '' || r[6] == null || isNaN(Number(r[6]))) return;
      var keys = [];
      try { keys = (JSON.parse(r[3] || '{}').keys) || []; } catch (e) {}
      var day = (r[0] instanceof Date) ? Utilities.formatDate(r[0], self.TZ, 'yyyy-MM-dd') : String(r[0]).slice(0, 10);
      samples.push({ trend: Number(r[6]) / 100, keys: keys, day: day });
    });
    var all = {};
    samples.forEach(function (s) { s.keys.forEach(function (k) { all[k] = true; }); });
    Object.keys(all).forEach(function (k) {
      var w = { n: 0, sum: 0, days: {} }, wo = { n: 0, sum: 0 };
      samples.forEach(function (s) {
        if (s.keys.indexOf(k) !== -1) { w.n++; w.sum += s.trend; w.days[s.day] = 1; }
        else { wo.n++; wo.sum += s.trend; }
      });
      var nd = Object.keys(w.days).length;
      var enough = w.n >= cfg.minSamples && nd >= cfg.minDays && wo.n >= cfg.minSamples;
      var uplift = (w.n && wo.n) ? (w.sum / w.n) - (wo.sum / wo.n) : 0;
      out[k] = {
        key: k, samples: w.n, days: nd, baseline: wo.n, uplift: Math.round(uplift * 1000) / 1000,
        status: !enough ? 'learning' : (uplift >= cfg.minUplift ? 'raises' : 'no-effect')
      };
    });
    return out;
  },

  // ── Factors for a date ────────────────────────────────────────────────────
  getFactors: function (dateStr, isToday) {
    var ck = this.CACHE_KEY + '|' + dateStr;
    try { var hit = CacheService.getScriptCache().get(ck); if (hit) return JSON.parse(hit); } catch (e) {}
    var ss = SpreadsheetApp.getActiveSpreadsheet(), cfg = this._cfg(), factors = [];

    var fb = this.forecastBias(ss, dateStr, cfg);
    if (fb) factors.push(fb);

    var st = this.latestStats(ss), freshStatsAt = null;
    if (isToday && st) {
      var ageMin = (Date.now() - st.at) / 60000;
      if (ageMin <= this.FRESH_MIN) {
        freshStatsAt = Utilities.formatDate(new Date(st.at), this.TZ, 'HH:mm');
        var tr = st.trend != null ? st.trend / 100 : 0;
        var svlGap = (st.svl != null && st.svl < cfg.svlTarget) ? (cfg.svlTarget - st.svl) / 100 : 0;
        var up = Math.min(cfg.liveCap, Math.max(0, tr) + 0.5 * svlGap);
        if (up >= 0.02) {
          var nowMin = this._nowMin();
          factors.push({
            id: 'live', kind: 'live', pct: Math.round(up * 1000) / 1000, from: nowMin, to: nowMin + cfg.liveWindowMin, decay: true,
            label: 'Live +' + Math.round(up * 100) + '% until ' + this._hm(nowMin + cfg.liveWindowMin),
            detail: 'Calculator at ' + freshStatsAt + ': SVL ' + (st.svl != null ? Math.round(st.svl) + '%' : '—') +
              (st.trend != null ? ', alarms ' + (st.trend >= 0 ? '+' : '') + (Math.round(st.trend * 10) / 10) + '% vs the IEX trend (last 60 min)' : '') +
              '. Demand raised for the next ' + cfg.liveWindowMin + ' min, fading out.',
            svl: st.svl, trend: st.trend
          });
        }
      }
    }

    var learned = this.learn(ss, cfg);
    var cond = isToday ? this.currentConditions() : this.forecastConditions(dateStr);
    var self = this;
    cond.keys.forEach(function (k) {
      var L = learned[k] || { key: k, samples: 0, days: 0, status: 'learning', uplift: 0 };
      var name = self._condLabel(k, cond.raw);
      if (L.status === 'raises') {
        factors.push({ id: 'cond_' + k, kind: 'learned', pct: Math.min(0.3, L.uplift), label: name + ' +' + Math.round(L.uplift * 100) + '%',
          detail: 'In ' + L.samples + ' logged hours over ' + L.days + ' days with this condition, our alarms ran ' + Math.round(L.uplift * 100) + ' pts further over the IEX trend than without it.' + (isToday ? '' : ' Forecast for this date.'),
          samples: L.samples });
      } else {
        factors.push({ id: 'cond_' + k, kind: 'headsup', pct: 0, label: name,
          detail: L.status === 'no-effect'
            ? 'Logged ' + L.samples + ' hours with this condition: no real change in our alarm volume (' + (L.uplift >= 0 ? '+' : '') + Math.round(L.uplift * 100) + ' pts). Not applied.'
            : 'Effect on our alarms not known yet — ' + L.samples + '/' + cfg.minSamples + ' logged hours (' + L.days + '/' + cfg.minDays + ' days). Shown only.',
          samples: L.samples, status: L.status });
      }
    });

    // Conditions travel together (a QC wind warning, QC wind ≥ 50 and QC 25k
    // out are usually the same storm), and the 5k bucket is a subset of the
    // 25k one — summing them would count one event several times. So per
    // province only the largest learned uplift applies; the rest are shown as
    // covered by it.
    var bestBy = {};
    factors.forEach(function (fx) {
      if (fx.kind !== 'learned') return;
      var g = (fx.id.match(/_([A-Z]{2})$/) || [])[1] || fx.id;
      if (!bestBy[g] || fx.pct >= bestBy[g].pct) bestBy[g] = fx;   // tie: the later, more specific bucket (25k after 5k)
    });
    factors.forEach(function (fx) {
      if (fx.kind !== 'learned') return;
      var g = (fx.id.match(/_([A-Z]{2})$/) || [])[1] || fx.id;
      if (bestBy[g] !== fx) {
        fx.detail += ' Same area as "' + bestBy[g].label + '", which already applies — not added twice.';
        fx.covered = true; fx.kind = 'headsup'; fx.pct = 0;
      }
    });

    var res = { factors: factors, learned: Object.keys(learned).map(function (k) { return learned[k]; }), freshStatsAt: freshStatsAt, conditions: cond.keys };
    try { CacheService.getScriptCache().put(ck, JSON.stringify(res), this.CACHE_TTL); } catch (e) {}
    return res;
  },

  // Actual vs forecast alarm volume over the last N imported days before date.
  forecastBias: function (ss, dateStr, cfg) {
    var sh = ss.getSheetByName('WF_FORECAST');
    if (!sh || sh.getLastRow() < 2) return null;
    var self = this, rows = [];
    sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(function (r) {
      if (String(r[0]).toLowerCase() !== 'day') return;
      var d = (r[1] instanceof Date) ? Utilities.formatDate(r[1], self.TZ, 'yyyy-MM-dd') : String(r[1]).slice(0, 10);
      var fc = Number(r[3]), av = Number(r[4]);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || d >= dateStr || !(fc > 0) || !(av > 0)) return;
      rows.push({ d: d, fc: fc, av: av });
    });
    rows.sort(function (a, b) { return a.d < b.d ? 1 : -1; });
    rows = rows.slice(0, cfg.forecastDays);
    if (rows.length < cfg.forecastMinDays) return null;
    var F = 0, A = 0; rows.forEach(function (x) { F += x.fc; A += x.av; });
    var bias = Math.max(-0.15, Math.min(0.25, A / F - 1));
    if (Math.abs(bias) < 0.03) return null;
    return {
      id: 'forecast', kind: 'forecast', pct: Math.round(bias * 1000) / 1000,
      label: 'Forecast ' + (bias >= 0 ? '+' : '−') + Math.abs(Math.round(bias * 100)) + '%',
      detail: 'Over the last ' + rows.length + ' days with data (' + rows[rows.length - 1].d + ' → ' + rows[0].d + '), actual alarms were ' +
        Math.round(A).toLocaleString() + ' vs ' + Math.round(F).toLocaleString() + ' forecast. The IDP is built from that forecast, so it is scaled by the same amount.',
      days: rows.length
    };
  },

  _condLabel: function (k, raw) {
    var m = k.match(/^(out5k|out25k|alert|wind50)_(\w+)$/), P = this.PROVS;
    if (k === 'alarmcom_down') return 'Alarm.com ' + (raw.alarmcom || 'issue');
    if (!m) return k;
    var pv = P[m[2]] || m[2];
    if (m[1] === 'out5k' || m[1] === 'out25k') return pv + ' outages ≥ ' + (m[1] === 'out5k' ? '5k' : '25k') + (raw.out && raw.out[m[2]] ? ' (' + Math.round(raw.out[m[2]] / 1000) + 'k out)' : '');
    if (m[1] === 'alert') return pv + ' weather alert' + (raw.alert && raw.alert[m[2]] ? ' (' + raw.alert[m[2]] + ')' : '');
    return pv + ' wind ≥ 50 km/h' + (raw.wind && raw.wind[m[2]] ? ' (' + raw.wind[m[2]] + ')' : '');
  },
  _today: function () { return Utilities.formatDate(new Date(), this.TZ, 'yyyy-MM-dd'); },
  _todayParts: function () { var p = this._today().split('-').map(Number); return [p[0], p[1] - 1, p[2]]; },
  _nowMin: function () { var hm = Utilities.formatDate(new Date(), this.TZ, 'HH:mm').split(':'); return parseInt(hm[0], 10) * 60 + parseInt(hm[1], 10); },
  _hm: function (m) { m = ((m % 1440) + 1440) % 1440; return ('0' + Math.floor(m / 60)).slice(-2) + ':' + ('0' + (m % 60)).slice(-2); }
};

// ── Globals — run these from the Apps Script editor ─────────────────────────
function logDemandConditions() { return DemandSignals.logConditions(); }

function setupDemandTrigger() {
  removeDemandTrigger();
  ScriptApp.newTrigger('logDemandConditions').timeBased().everyHours(1).create();
  return 'Demand-signal logger installed (hourly). Run the Calculator through the day so each hour has a fresh alarm trend to learn from.';
}

function removeDemandTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'logDemandConditions') { ScriptApp.deleteTrigger(t); removed++; }
  });
  return removed + ' demand-signal trigger(s) removed.';
}

/** What the planner would apply today, and what has been learned so far. */
function debugDemandSignals() {
  try { CacheService.getScriptCache().remove(DemandSignals.CACHE_KEY + '|' + DemandSignals._today()); } catch (e) {}
  var f = DemandSignals.getFactors(DemandSignals._today(), true);
  Logger.log(JSON.stringify(f, null, 2));
  return f;
}
