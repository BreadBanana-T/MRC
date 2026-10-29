/**
 * MODULE: ALARM.COM STATUS MONITOR
 *
 * Probes the Alarm.com standby status page and keeps an uptime history so the
 * Dashboard can show, at a glance, whether the operational system is up.
 *
 * ────────────────────────────────────────────────────────────────────
 * WHAT IT CHECKS
 *   GET https://standby.alarm.com/  (browser User-Agent — a status page can
 *   serve differently to a non-browser client). Expected 200; 200–399 counts
 *   as reachable. Then two content checks against the raw HTML:
 *     • /All systems operational/i          → the global status banner
 *     • /Partner Portal[\s\S]*?\bUp\b/i     → the "Monitors" component row
 *
 *   The page's banner is often painted by JavaScript, which UrlFetchApp does
 *   not run, so _readStatus also tries status JSON embedded in the page and
 *   the Statuspage / Instatus JSON APIs on the same host. Result is tri-state:
 *   up | issue (the page reports a problem) | unknown (page answered but no
 *   reader could tell). Unknown is logged blank and kept OUT of the uptime %,
 *   so an unreadable page no longer reads as a 0% outage. partnerPortalUp is
 *   recorded and surfaced as a component detail; it does not gate isUp.
 *
 * SETUP (run once from the Apps Script editor — this project has no trigger UI)
 *   1. debugAlarmCom()        — read-only. Prints the resolved URL, status,
 *                               latency, both regex results and the first bytes
 *                               of the body. Run this FIRST: UrlFetchApp does
 *                               not execute JavaScript, so if the status page
 *                               paints its banner client-side the regexes will
 *                               miss on a perfectly healthy site. A 200 with
 *                               isOperational:false means exactly that.
 *   2. setupAlarmComTrigger() — installs the 5-minute background probe. Safe to
 *                               re-run; it removes any existing one first.
 *   3. setAlarmComUrl('...')  — optional. Overrides the probe URL without a
 *                               redeploy. '' disables probing entirely.
 * ────────────────────────────────────────────────────────────────────
 *
 * The payload is cached in CacheService for 5 minutes so N dashboard viewers
 * cause at most one probe per interval, and the last good payload is mirrored
 * into a Script Property so a transient failure shows "last known" rather than
 * an empty tile.
 *
 * NOTE ON NAMING: `WF_ALARMS` / importAlarms already mean the IEX alarm-volume
 * report. Nothing here is called "Alarms" — this is Alarm.com the vendor.
 */

var AlarmComMonitor = {

  TZ: 'America/Toronto',

  CACHE_KEY: 'MRC_ALARMCOM_V1',
  CACHE_TTL_SECONDS: 300,
  LAST_GOOD_PROP: 'ALARMCOM_LAST_GOOD',
  URL_PROP: 'ALARMCOM_URL',

  _DEFAULT_URL: 'https://standby.alarm.com/',

  // A status page can serve a different (or no) body to a non-browser client,
  // so present as one.
  _UA: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',

  // Raw-HTML signals. A status page often paints its banner with JavaScript
  // (UrlFetchApp runs none), so these are only the first of several readers —
  // see _readStatus.
  OPERATIONAL_RE: /All systems (?:are )?operational|All services (?:are )?(?:online|operational)|Fully operational|No (?:known )?(?:issues|incidents) (?:reported|at this time)/i,
  // Diagnostic only (debugAlarmCom) — see the note in _readStatus.
  ISSUE_RE: /Major outage|Partial outage|Service disruption|Degraded performance|Under maintenance|Investigating/i,
  PARTNER_RE: /Partner Portal[\s\S]*?\bUp\b/i,

  HISTORY_SHEET: 'WF_ALARMCOM_UPTIME',
  // The probe runs every 5 min; allow a little slack so a trigger that fires a
  // few seconds early is not silently dropped as a duplicate.
  HISTORY_MIN_GAP_MS: 4.5 * 60 * 1000,
  HISTORY_MAX_ROWS: 2600, // ≈9 days at one row / 5 min

  // Script Property override; null/undefined → default, '' → disabled.
  _getUrl: function() {
    try {
      var v = PropertiesService.getScriptProperties().getProperty(this.URL_PROP);
      if (v === null || v === undefined) return this._DEFAULT_URL;
      return v;
    } catch (e) { return this._DEFAULT_URL; }
  },

  _headers: function() {
    return {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': this._UA }
    };
  },

  // ── Reading the status ────────────────────────────────────────────────────
  // Returns { state: 'up'|'issue'|'unknown', partnerPortalUp: true|false|null,
  // source, detail }. Tried in order, first decisive answer wins:
  //   1. the raw HTML (banner text, or status JSON embedded in the page)
  //   2. the Atlassian Statuspage API   <origin>/api/v2/summary.json
  //   3. the Instatus API               <origin>/summary.json
  // 'unknown' means the page answered but none of these could read it — that is
  // NOT an outage and is kept out of the uptime % (it used to count as down,
  // which is how the tile sat at 0%).
  _readStatus: function(url, body) {
    var partner = this.PARTNER_RE.test(body) ? true : null;
    if (this.OPERATIONAL_RE.test(body) || /"indicator"\s*:\s*"none"/.test(body)) return { state: 'up', partnerPortalUp: partner, source: 'page', detail: '' };
    var ind = body.match(/"indicator"\s*:\s*"(minor|major|critical|maintenance)"/);
    if (ind) return { state: 'issue', partnerPortalUp: partner, source: 'page', detail: ind[1] };

    var origin = String(url).replace(/^(https?:\/\/[^\/]+).*$/, '$1') + '/';
    var sp = this._fetchJson(origin + 'api/v2/summary.json');
    if (sp && sp.status && sp.status.indicator) {
      var pp = null;
      (sp.components || []).forEach(function(c) { if (/partner portal/i.test(c.name || '')) pp = (c.status === 'operational'); });
      return { state: sp.status.indicator === 'none' ? 'up' : 'issue', partnerPortalUp: pp, source: 'statuspage-api', detail: sp.status.description || sp.status.indicator };
    }
    var ins = this._fetchJson(origin + 'summary.json');
    if (ins && ins.page && ins.page.status) {
      return { state: String(ins.page.status).toUpperCase() === 'UP' ? 'up' : 'issue', partnerPortalUp: null, source: 'instatus-api', detail: String(ins.page.status) };
    }

    // Nothing decisive. Outage WORDS in the raw HTML are not used to call an
    // issue: status pages print "Major outage / Partial outage" in their legend
    // and past-incident list even when everything is up.
    return { state: 'unknown', partnerPortalUp: partner, source: '', detail: '' };
  },

  _fetchJson: function(url) {
    try {
      var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': this._UA, 'Accept': 'application/json' } });
      if (res.getResponseCode() !== 200) return null;
      var t = res.getContentText();
      if (!/^\s*[\[{]/.test(t)) return null;
      return JSON.parse(t);
    } catch (e) { return null; }
  },

  // ── The probe ─────────────────────────────────────────────────────────────
  // Returns { skipped } | { up, state, statusCode, isReachable, isOperational,
  //                         partnerPortalUp, latencyMs, source, note, error }
  probe: function() {
    var url = this._getUrl();
    if (!url) return { skipped: true };

    var started = Date.now();
    try {
      var res = UrlFetchApp.fetch(url, this._headers());
      var latencyMs = Date.now() - started;
      var code = res.getResponseCode();
      var isReachable = code >= 200 && code <= 399;
      var body = '';
      try { body = res.getContentText(); } catch (e) { body = ''; }

      if (!isReachable) {
        return { up: false, state: 'down', statusCode: code, isReachable: false, isOperational: false, partnerPortalUp: null, latencyMs: latencyMs, source: '', note: 'HTTP ' + code, error: '' };
      }
      var st = this._readStatus(url, body);
      var note = st.state === 'issue' ? ('reported: ' + st.detail)
               : (st.state === 'unknown' ? 'status not readable' : '');
      return {
        up: st.state === 'up',
        state: st.state,
        statusCode: code,
        isReachable: true,
        isOperational: st.state === 'up',
        partnerPortalUp: st.partnerPortalUp,
        latencyMs: latencyMs,
        source: st.source,
        note: note,
        error: ''
      };
    } catch (e) {
      return {
        up: false,
        state: 'down',
        statusCode: 0,
        isReachable: false,
        isOperational: false,
        partnerPortalUp: null,
        latencyMs: Date.now() - started,
        source: '',
        note: 'fetch failed',
        error: e.message
      };
    }
  },

  // ── Uptime history (hidden sheet, ring-buffered) ───────────────────────────
  _historySheet: function(createIfMissing) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(this.HISTORY_SHEET);
    if (!sheet && createIfMissing) {
      sheet = ss.insertSheet(this.HISTORY_SHEET);
      sheet.appendRow(['Timestamp', 'Up', 'LatencyMs', 'HttpCode', 'PartnerPortalUp', 'Note']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      sheet.setFrozenRows(1);
      try { sheet.hideSheet(); } catch (e) {}
    }
    return sheet;
  },

  _logHistory: function(p) {
    try {
      var sheet = this._historySheet(true);
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var lastTs = sheet.getRange(lastRow, 1).getValue();
        if (lastTs && (Date.now() - new Date(lastTs).getTime()) < this.HISTORY_MIN_GAP_MS) return false;
      }
      sheet.appendRow([
        new Date(),
        p.state === 'unknown' ? '' : (p.up ? 1 : 0),   // blank = page answered, status unreadable
        p.latencyMs || 0,
        p.statusCode || 0,
        p.partnerPortalUp === null || p.partnerPortalUp === undefined ? '' : (p.partnerPortalUp ? 1 : 0),
        p.note || p.error || ''
      ]);
      var rows = sheet.getLastRow() - 1;
      if (rows > this.HISTORY_MAX_ROWS) sheet.deleteRows(2, rows - this.HISTORY_MAX_ROWS);
      return true;
    } catch (e) { return false; }
  },

  // Uptime % over the trailing `hoursBack`, plus the most recent down sample.
  _uptimeStats: function() {
    var out = { uptime24h: null, uptime7d: null, samples24h: 0, lastDownAt: '' };
    try {
      var sheet = this._historySheet(false);
      if (!sheet || sheet.getLastRow() < 2) return out;
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
      var now = Date.now();
      var c24 = now - 24 * 3600 * 1000, c7 = now - 7 * 24 * 3600 * 1000;
      var up24 = 0, n24 = 0, up7 = 0, n7 = 0, lastDown = 0;
      for (var i = 0; i < data.length; i++) {
        var ts = new Date(data[i][0]).getTime();
        if (isNaN(ts)) continue;
        if (this._isUnknownRow(data[i])) continue;   // can't count what we couldn't read
        var up = Number(data[i][1]) === 1;
        if (!up && ts > lastDown) lastDown = ts;
        if (ts >= c7) { n7++; if (up) up7++; }
        if (ts >= c24) { n24++; if (up) up24++; }
      }
      if (n24) { out.uptime24h = Math.round(up24 / n24 * 1000) / 10; out.samples24h = n24; }
      if (n7) out.uptime7d = Math.round(up7 / n7 * 1000) / 10;
      if (lastDown) out.lastDownAt = Utilities.formatDate(new Date(lastDown), this.TZ, 'MMM d, HH:mm');
      return out;
    } catch (e) { return out; }
  },

  // A sample whose status couldn't be read: logged blank, or a legacy row from
  // before the multi-reader probe ("banner text not matched", logged as 0).
  _isUnknownRow: function(row) {
    return row[1] === '' || /banner text not matched|status not readable/i.test(String(row[5] || ''));
  },

  getHistorySeries: function(hoursBack) {
    hoursBack = hoursBack || 24;
    try {
      var sheet = this._historySheet(false);
      if (!sheet || sheet.getLastRow() < 2) return '[]';
      var cutoff = Date.now() - hoursBack * 3600 * 1000;
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
      var out = [];
      for (var i = 0; i < data.length; i++) {
        var ts = new Date(data[i][0]).getTime();
        if (isNaN(ts) || ts < cutoff) continue;
        // up: 1 / 0, or null when the status couldn't be read (grey on the strip).
        out.push({ t: ts, up: this._isUnknownRow(data[i]) ? null : (Number(data[i][1]) === 1 ? 1 : 0), ms: parseInt(data[i][2], 10) || 0 });
      }
      out.sort(function(a, b) { return a.t - b.t; });
      return JSON.stringify(out);
    } catch (e) { return '[]'; }
  },

  // Turn a probe result into the dashboard payload, log it, and cache it. Shared
  // by getStatus() and the trigger handler so a background probe leaves the cache
  // WARM — otherwise the next dashboard load would immediately probe again.
  _publish: function(p) {
    this._logHistory(p);
    var stats = this._uptimeStats();
    var props = null;
    try { props = PropertiesService.getScriptProperties(); } catch (e) {}

    var result = {
      up: !!p.up,
      state: p.state || (p.up ? 'up' : 'down'),
      source: p.source || '',
      statusCode: p.statusCode,
      isReachable: !!p.isReachable,
      isOperational: !!p.isOperational,
      partnerPortalUp: (p.partnerPortalUp === null || p.partnerPortalUp === undefined) ? null : !!p.partnerPortalUp,
      latencyMs: p.latencyMs || 0,
      checked: Utilities.formatDate(new Date(), this.TZ, 'HH:mm'),
      uptime24h: stats.uptime24h,
      uptime7d: stats.uptime7d,
      samples24h: stats.samples24h,
      lastDownAt: stats.lastDownAt,
      note: p.note || '',
      error: p.error || ''
    };

    // When the host is unreachable, carry the last time we saw it healthy so the
    // tile can say "last good HH:mm" instead of just going blank. Resolved BEFORE
    // caching, or the cached copy would be missing it for the next 5 minutes.
    if (!result.isReachable && props) {
      try {
        var lg = props.getProperty(this.LAST_GOOD_PROP);
        if (lg) result.lastGoodChecked = (JSON.parse(lg).checked) || '';
      } catch (e) {}
    }

    var payload = JSON.stringify(result);
    try { CacheService.getScriptCache().put(this.CACHE_KEY, payload, this.CACHE_TTL_SECONDS); } catch (e) {}
    if (result.up && props) { try { props.setProperty(this.LAST_GOOD_PROP, payload); } catch (e) {} }
    return payload;
  },

  // ── Client-facing payload ─────────────────────────────────────────────────
  getStatus: function() {
    try {
      var cached = CacheService.getScriptCache().get(this.CACHE_KEY);
      if (cached) return cached;
    } catch (e) {}

    var p = this.probe();
    if (p.skipped) {
      return JSON.stringify({ disabled: true, checked: Utilities.formatDate(new Date(), this.TZ, 'HH:mm') });
    }
    return this._publish(p);
  }
};

// ── Globals — run these from the Apps Script editor ─────────────────────────

/** Time-driven handler: probe + log, bypassing the read cache. */
function probeAlarmCom() {
  if (typeof AlarmComMonitor === 'undefined') return 'AlarmComMonitor unavailable.';
  var p = AlarmComMonitor.probe();
  if (p.skipped) return 'Alarm.com probe disabled (ALARMCOM_URL is empty).';
  // _publish logs the row AND refreshes the cached payload, so the dashboard
  // shows this probe's result without triggering another fetch of its own.
  AlarmComMonitor._publish(p);
  return 'Alarm.com: ' + String(p.state || (p.up ? 'up' : 'down')).toUpperCase() + ' (HTTP ' + p.statusCode + ', ' + p.latencyMs + 'ms)'
       + (p.note ? ' — ' + p.note : '');
}

/**
 * Read-only diagnostic. Run this FIRST after deploying: it is the only way to
 * confirm the status banner is present in the RAW HTML (UrlFetchApp does not
 * run JavaScript). A 200 with isOperational:false means the page is rendered
 * client-side and isUp should fall back to the status code alone.
 */
function debugAlarmCom() {
  var url = AlarmComMonitor._getUrl();
  Logger.log('Alarm.com probe URL: ' + (url || '(disabled)'));
  if (!url) return { disabled: true };
  var origin = String(url).replace(/^(https?:\/\/[^\/]+).*$/, '$1') + '/';
  var started = Date.now();
  try {
    var res = UrlFetchApp.fetch(url, AlarmComMonitor._headers());
    var body = '';
    try { body = res.getContentText(); } catch (e) { body = ''; }
    Logger.log('Page: HTTP ' + res.getResponseCode() + ' in ' + (Date.now() - started) + 'ms · ' + body.length + ' bytes');
    Logger.log('  operational text in HTML → ' + AlarmComMonitor.OPERATIONAL_RE.test(body));
    Logger.log('  issue text in HTML       → ' + ((body.match(AlarmComMonitor.ISSUE_RE) || [''])[0] || 'none'));
    Logger.log('  Partner Portal "Up"      → ' + AlarmComMonitor.PARTNER_RE.test(body));
    Logger.log('  embedded "indicator"     → ' + ((body.match(/"indicator"\s*:\s*"\w+"/) || ['none'])[0]));
    ['api/v2/summary.json', 'summary.json'].forEach(function(path) {
      try {
        var r = UrlFetchApp.fetch(origin + path, { muteHttpExceptions: true, headers: { 'User-Agent': AlarmComMonitor._UA, 'Accept': 'application/json' } });
        Logger.log(origin + path + ' → HTTP ' + r.getResponseCode() + ' · ' + r.getContentText().substring(0, 200).replace(/\s+/g, ' '));
      } catch (e) { Logger.log(origin + path + ' → ' + e.message); }
    });
    var st = AlarmComMonitor._readStatus(url, body);
    Logger.log('RESULT: ' + JSON.stringify(st));
    Logger.log('First 800 bytes of the page:\n' + body.substring(0, 800));
    return st;
  } catch (e) {
    Logger.log('Fetch failed after ' + (Date.now() - started) + 'ms: ' + e.message);
    return { error: e.message };
  }
}

/** Installs the 5-minute background probe. Idempotent — safe to re-run. */
function setupAlarmComTrigger() {
  removeAlarmComTrigger();
  ScriptApp.newTrigger('probeAlarmCom').timeBased().everyMinutes(5).create();
  return 'Alarm.com uptime trigger installed (every 5 minutes).';
}

function removeAlarmComTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'probeAlarmCom') { ScriptApp.deleteTrigger(t); removed++; }
  });
  return removed + ' Alarm.com trigger(s) removed.';
}

/** Override the probe URL without a redeploy. '' disables probing. */
function setAlarmComUrl(url) {
  PropertiesService.getScriptProperties().setProperty(AlarmComMonitor.URL_PROP, String(url == null ? '' : url).trim());
  try { CacheService.getScriptCache().remove(AlarmComMonitor.CACHE_KEY); } catch (e) {}
  return 'Alarm.com probe URL set to: ' + (url ? url : '(disabled)');
}
