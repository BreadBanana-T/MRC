/**
 * MODULE: OUTAGE TRACKER
 * Pulls real-time power outage counts for BC, ON (Hydro One), and QC
 * (Hydro-Québec). All endpoints are publicly served (no API key, no auth).
 *
 * URLs are overridable via Script Properties without editing code. Keys:
 *   BC_HYDRO_URL, HYDRO_ONE_URL, HYDRO_QUEBEC_URL
 *
 * If a URL is blank (Script Property set to empty string), that fetcher
 * silently skips and the UI shows "source offline" for that province.
 *
 * ────────────────────────────────────────────────────────────────────
 * FINDING THE HYDRO ONE URL
 * ────────────────────────────────────────────────────────────────────
 * Hydro One's Storm Centre backend shifts periodically. Current discovery
 * recipe:
 *   1. Open https://stormcentre.hydroone.com/ in Chrome
 *   2. Open DevTools -> Network tab, filter "query" or "FeatureServer"
 *   3. Refresh the page. Look for a request hitting a URL like
 *        https://services*.arcgis.com/<orgId>/arcgis/rest/services/<SvcName>/FeatureServer/0/query
 *      or a self-hosted variant such as
 *        https://<host>.hydroone.com/.../rest/services/Outages/MapServer/0/query
 *   4. Copy the FULL URL including query string, or just the base /query
 *      path and this module will add the usual params.
 *   5. Set it via the Apps Script editor:
 *        setOutageUrl('ON', '<paste full URL here>')
 *      (See `setOutageUrl` at bottom of this file.)
 *
 * BC Hydro's outage-list.json has been stable for years. Hydro-Québec's
 * bilan.json has drifted a few times — if QC fails, do the same recipe
 * against info-pannes.hydroquebec.com.
 * ────────────────────────────────────────────────────────────────────
 *
 * Payload is cached in CacheService for 5 minutes so polling the dashboard
 * doesn't spam the utilities or eat UrlFetch quota.
 */

var OutageTracker = {
  CACHE_KEY: 'MRC_OUTAGES_V1',
  CACHE_TTL_SECONDS: 300,

  _DEFAULT_BC_URL: 'https://www.bchydro.com/power-outages/app/outage-list.json',
  // Intentionally blank by default: Hydro One's ArcGIS endpoint shifts. Set
  // via setOutageUrl('ON', '...') once discovered (see recipe above).
  _DEFAULT_HYDRO_ONE_URL: '',
  _DEFAULT_HQ_URL: 'https://www.hydroquebec.com/data/pannes/donnees/bilan.json',

  _getUrl: function(key, fallback) {
    try {
      var v = PropertiesService.getScriptProperties().getProperty(key);
      if (v === null || v === undefined) return fallback;
      return v; // explicit empty string = "disabled"
    } catch (e) { return fallback; }
  },

  fetchAll: function() {
    try {
      var cached = CacheService.getScriptCache().get(this.CACHE_KEY);
      if (cached) return cached;
    } catch (e) {}

    var result = {
      updated: Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm'),
      byProvince: {},
      errors: []
    };

    var bc = this._fetchBCHydro();
    if (bc.error) result.errors.push('BC: ' + bc.error);
    else if (bc.skipped) result.errors.push('BC: not configured');
    else result.byProvince.BC = bc.data;

    var on = this._fetchHydroOne();
    if (on.error) result.errors.push('ON: ' + on.error);
    else if (on.skipped) result.errors.push('ON: URL not configured — run setOutageUrl("ON", "...")');
    else result.byProvince.ON = on.data;

    var qc = this._fetchHQ();
    if (qc.error) result.errors.push('QC: ' + qc.error);
    else if (qc.skipped) result.errors.push('QC: not configured');
    else result.byProvince.QC = qc.data;

    var payload = JSON.stringify(result);
    try { CacheService.getScriptCache().put(this.CACHE_KEY, payload, this.CACHE_TTL_SECONDS); } catch (e) {}

    // Log the national total so the dashboard can draw outage history as a
    // band behind the Service Level chart. Only runs on cache misses, and
    // _logHistory itself throttles to one row per 10 minutes.
    try {
      var total = 0;
      for (var prov in result.byProvince) total += (result.byProvince[prov].customers || 0);
      this._logHistory(total);
    } catch (e) {}

    return payload;
  },

  // ── Outage history (for the SL-chart overlay) ──────────────────────────
  HISTORY_SHEET: 'Outage History',
  HISTORY_MIN_GAP_MS: 10 * 60 * 1000,
  HISTORY_MAX_ROWS: 600, // ≈4 days at one row / 10 min

  _logHistory: function(totalCustomers) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(this.HISTORY_SHEET);
      if (!sheet) {
        sheet = ss.insertSheet(this.HISTORY_SHEET);
        sheet.appendRow(['Timestamp', 'Total Customers']);
        sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
        try { sheet.hideSheet(); } catch (e) {}
      }
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var lastTs = sheet.getRange(lastRow, 1).getValue();
        if (lastTs && (Date.now() - new Date(lastTs).getTime()) < this.HISTORY_MIN_GAP_MS) return;
      }
      sheet.appendRow([new Date(), totalCustomers]);
      var rows = sheet.getLastRow() - 1;
      if (rows > this.HISTORY_MAX_ROWS) sheet.deleteRows(2, rows - this.HISTORY_MAX_ROWS);
    } catch (e) {}
  },

  getHistorySeries: function(hoursBack) {
    hoursBack = hoursBack || 26;
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(this.HISTORY_SHEET);
      if (!sheet || sheet.getLastRow() < 2) return '[]';
      var cutoff = Date.now() - hoursBack * 3600 * 1000;
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
      var out = [];
      for (var i = 0; i < data.length; i++) {
        var t = new Date(data[i][0]).getTime();
        if (isNaN(t) || t < cutoff) continue;
        out.push({ t: t, customers: parseInt(data[i][1], 10) || 0 });
      }
      out.sort(function(a, b) { return a.t - b.t; });
      return JSON.stringify(out);
    } catch (e) { return '[]'; }
  },

  _safeInt: function(v) {
    var n = parseInt(String(v == null ? 0 : v).replace(/[^\d-]/g, ''), 10);
    return isNaN(n) ? 0 : n;
  },

  // Utility outage layers name the "customers affected" column inconsistently
  // (NumCustomersAffected, CUST_QTY, customersaffected, ...). Try the known
  // names, then fall back to sniffing any attribute key that looks like a
  // customer count — so an unexpected field name can't silently zero the total.
  _CUST_FIELDS: [
    'NumCustomersAffected', 'CustomersAffected', 'CUSTOMERS_AFFECTED', 'CUSTOMERSAFFECTED',
    'customersAffected', 'customersaffected', 'customers', 'Customers', 'CUSTOMERS',
    'CUST_QTY', 'CustQty', 'custQty', 'CUST_AFF', 'CustAffected', 'CUSTOMERS_OUT',
    'TOTAL_CUST', 'TOTALCUST', 'totalCustomersAffected', 'AffectedCustomers', 'affectedCustomers',
    'numCustomersOut', 'NumCustomersOut', 'customersOut', 'CustomersOut', 'customerCount',
    'CustomerCount', 'cust_a', 'CustomersImpacted', 'customersImpacted', 'nbClientInterrompu'
  ],
  _extractCustomers: function(attrs) {
    if (!attrs) return 0;
    for (var i = 0; i < this._CUST_FIELDS.length; i++) {
      var k = this._CUST_FIELDS[i];
      if (attrs[k] !== undefined && attrs[k] !== null && attrs[k] !== '') return this._safeInt(attrs[k]);
    }
    var keys = Object.keys(attrs);
    // Prefer keys that mention customers AND an affected/count-ish token.
    for (var a = 0; a < keys.length; a++) {
      if (/cust|client/i.test(keys[a]) && /(out|aff|num|qty|total|impact|interru|count)/i.test(keys[a])) {
        var v = this._safeInt(attrs[keys[a]]); if (v) return v;
      }
    }
    for (var b = 0; b < keys.length; b++) {
      if (/cust|client/i.test(keys[b])) { var v2 = this._safeInt(attrs[keys[b]]); if (v2) return v2; }
    }
    return 0;
  },

  _REGION_FIELDS: [
    'MunicipalityName', 'Municipality', 'MUNICIPALITY', 'LocationDescription', 'Location',
    'Town', 'TOWNSHIP', 'Township', 'Community', 'COMMUNITY', 'City', 'CITY',
    'area', 'Area', 'AREA', 'region', 'Region', 'REGION', 'name', 'NAME', '_regionName', 'regionName'
  ],
  _extractRegion: function(attrs) {
    if (!attrs) return 'Unknown';
    for (var i = 0; i < this._REGION_FIELDS.length; i++) {
      var k = this._REGION_FIELDS[i];
      if (attrs[k] !== undefined && attrs[k] !== null && String(attrs[k]).trim() !== '') return String(attrs[k]).trim();
    }
    return 'Unknown';
  },

  // ETAs come as Unix epoch in milliseconds (BC Hydro), pre-formatted strings
  // (Hydro One), or null. Normalize to a short human-readable string.
  _fmtEta: function(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number' || /^\d{10,}$/.test(String(v))) {
      var ms = Number(v);
      if (String(ms).length === 10) ms *= 1000; // seconds → ms
      try {
        return Utilities.formatDate(new Date(ms), 'America/Toronto', "MMM d, h:mm a");
      } catch (e) { return String(v); }
    }
    return String(v);
  },

  _commonHeaders: function() {
    return {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'MRC-Ops-Dashboard/1.0', 'Accept': 'application/json' }
    };
  },

  _fetchBCHydro: function() {
    try {
      var url = this._getUrl('BC_HYDRO_URL', this._DEFAULT_BC_URL);
      if (!url) return { skipped: true };
      var res = UrlFetchApp.fetch(url, this._commonHeaders());
      if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
      var body = JSON.parse(res.getContentText());

      // BC Hydro nests outages two levels deep: body.regions[].outages[].
      // Older shapes used a flat list at body.outages or the body itself.
      var list = [];
      if (Array.isArray(body)) list = body;
      else if (Array.isArray(body.outages)) list = body.outages;
      else if (Array.isArray(body.data)) list = body.data;
      else if (Array.isArray(body.regions)) {
        body.regions.forEach(function(r) {
          if (Array.isArray(r.outages)) {
            r.outages.forEach(function(o) {
              if (o && !o._regionName) o._regionName = r.name;
              list.push(o);
            });
          }
        });
      }

      var total = 0;
      var unplannedCount = 0;
      var top = [];
      for (var i = 0; i < list.length; i++) {
        var o = list[i];
        var cause = String(o.cause || o.reason || '');
        // Skip planned/scheduled work — operationally we only care about
        // unplanned outages affecting customers right now.
        if (/\b(planned|scheduled)\b/i.test(cause)) continue;
        var n = this._safeInt(o.numCustomersOut || o.customersAffected || o.customers || o.customerCount);
        total += n;
        unplannedCount += 1;
        top.push({
          region: o.municipality || o._regionName || o.regionName || o.area || o.location || 'Unknown',
          customers: n,
          cause: cause || '—',
          eta: this._fmtEta(o.crewEta || o.crewEtr || o.estTimeOn || o.estimatedRestoration || null)
        });
      }
      top.sort(function(a, b) { return b.customers - a.customers; });
      return { data: { outages: unplannedCount, customers: total, top: top.slice(0, 10), source: 'BC Hydro' } };
    } catch (e) {
      return { error: e.message };
    }
  },

  _fetchHydroOne: function() {
    try {
      var url = this._getUrl('HYDRO_ONE_URL', this._DEFAULT_HYDRO_ONE_URL);
      if (!url) return { skipped: true };
      var isKubra = /kubra\.io|\/data\.json(\?|$)/i.test(url);

      // ---- Kubra / self-hosted summary JSON (totals only, no /query) ----
      if (isKubra) {
        var resK = UrlFetchApp.fetch(url, this._commonHeaders());
        if (resK.getResponseCode() !== 200) throw new Error('HTTP ' + resK.getResponseCode());
        var bodyK = JSON.parse(resK.getContentText());
        if (bodyK.summaryFileData && Array.isArray(bodyK.summaryFileData.totals) && bodyK.summaryFileData.totals.length) {
          var t = bodyK.summaryFileData.totals[0] || {};
          var custK = (t.total_cust_a && t.total_cust_a.val != null) ? t.total_cust_a.val : (t.total_cust_s && t.total_cust_s.val);
          return { data: { outages: this._safeInt(t.total_outages), customers: this._safeInt(custK), top: [], source: 'Hydro One' } };
        }
        throw new Error('Unrecognized Kubra summary shape');
      }

      // ---- ArcGIS FeatureServer / MapServer /query, WITH PAGINATION ----
      // The old code fetched a single page capped at 1000 records, so a large
      // storm (thousands of outages) was silently truncated — the likely cause
      // of "437 customers" when the true total is tens of thousands. We now page
      // through every record via resultOffset until the server stops returning
      // a full page.
      var base = url;
      if (!/\/query(\?|$)/i.test(base)) base = base.replace(/(FeatureServer\/\d+|MapServer\/\d+)(\/query)?/i, '$1/query');
      base = base.replace(/[?&]result(Offset|RecordCount)=\d+/gi, '');
      var addParam = function(u, k, v) { return (new RegExp('[?&]' + k + '=', 'i').test(u)) ? u : (u + (u.indexOf('?') === -1 ? '?' : '&') + k + '=' + v); };
      base = addParam(base, 'where', '1%3D1');
      base = addParam(base, 'outFields', '*');
      base = addParam(base, 'returnGeometry', 'false');
      base = addParam(base, 'f', 'json');

      var features = [];
      var offset = 0, pageSize = 2000, guard = 0;
      while (guard++ < 30) {
        var pageUrl = base + '&resultOffset=' + offset + '&resultRecordCount=' + pageSize;
        var res = UrlFetchApp.fetch(pageUrl, this._commonHeaders());
        if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
        var body = JSON.parse(res.getContentText());
        if (body.error) throw new Error('ArcGIS: ' + (body.error.message || 'unknown'));

        // Some Hydro One deployments answer the /query with a Kubra-style summary.
        if (body.summaryFileData && Array.isArray(body.summaryFileData.totals) && body.summaryFileData.totals.length) {
          var ts = body.summaryFileData.totals[0] || {};
          return { data: { outages: this._safeInt(ts.total_outages), customers: this._safeInt(ts.total_cust_a && ts.total_cust_a.val), top: [], source: 'Hydro One' } };
        }

        var page = body.features || [];
        features = features.concat(page);
        var more = (body.exceededTransferLimit === true) || (page.length === pageSize);
        if (!more || page.length === 0) break;
        offset += page.length;
      }

      var total = 0, top = [];
      for (var i = 0; i < features.length; i++) {
        var a = features[i].attributes || features[i].properties || {};
        var n = this._extractCustomers(a);
        total += n;
        top.push({
          region: this._extractRegion(a),
          customers: n,
          cause: a.Cause || a.CauseDescription || a.CauseCategory || a.cause || '—',
          eta: this._fmtEta(a.EstimatedRestoration || a.ETR || a.etr || a.CrewEta || a.CrewEtr || null)
        });
      }
      top.sort(function(x, y) { return y.customers - x.customers; });
      return { data: { outages: features.length, customers: total, top: top.slice(0, 10), source: 'Hydro One' } };
    } catch (e) {
      return { error: e.message };
    }
  },

  _fetchHQ: function() {
    try {
      var url = this._getUrl('HYDRO_QUEBEC_URL', this._DEFAULT_HQ_URL);
      if (!url) return { skipped: true };
      var res = UrlFetchApp.fetch(url, this._commonHeaders());
      if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
      var body = JSON.parse(res.getContentText());

      // New `bis` endpoint: top-level array of {id, nbClientInterrompu,
      // nbPanne, nbClientRaccorde}. The same outage is reported at every
      // hierarchy level (region → sub-region → sub-zone), so naively summing
      // double- or triple-counts. Use the magic id "HQ" for the official
      // province total, and only 2-char IDs (the 17 admin regions) for the
      // drawer breakdown to avoid duplicates.
      if (Array.isArray(body)) {
        var globalEntry = null;
        var top = [];
        var fallbackTotal = 0, fallbackOutages = 0;
        var regionNames = {
          '01': 'Bas-Saint-Laurent', '02': 'Saguenay – Lac-Saint-Jean',
          '03': 'Capitale-Nationale', '04': 'Mauricie', '05': 'Estrie',
          '06': 'Montréal', '07': 'Outaouais', '08': 'Abitibi-Témiscamingue',
          '09': 'Côte-Nord', '10': 'Nord-du-Québec',
          '11': 'Gaspésie – Îles-de-la-Madeleine', '12': 'Chaudière-Appalaches',
          '13': 'Laval', '14': 'Lanaudière', '15': 'Laurentides',
          '16': 'Montérégie', '17': 'Centre-du-Québec'
        };
        for (var i = 0; i < body.length; i++) {
          var e = body[i] || {};
          var id = String(e.id == null ? '' : e.id);
          var c = this._safeInt(e.nbClientInterrompu);
          var p = this._safeInt(e.nbPanne);
          if (id.toUpperCase() === 'HQ') {
            globalEntry = { customers: c, outages: p };
            continue;
          }
          if (id.length === 2 && /^\d{2}$/.test(id)) {
            fallbackTotal += c;
            fallbackOutages += p;
            if (c > 0 || p > 0) {
              top.push({ region: regionNames[id] || ('Region ' + id), customers: c, cause: '—', eta: null });
            }
          }
        }
        top.sort(function(a, b) { return b.customers - a.customers; });
        var total = globalEntry ? globalEntry.customers : fallbackTotal;
        var outages = globalEntry ? globalEntry.outages : fallbackOutages;
        return { data: { outages: outages, customers: total, top: top.slice(0, 10), source: 'Hydro-Québec' } };
      }

      // Legacy bilan.json shape.
      var regions = body.regions || body.list || body.data || [];
      var global = body.global || body.summary || body.total || {};

      var totalLegacy = this._safeInt(global.clients || global.clientsAffected || global.clientsSansElectricite);
      var outagesLegacy = this._safeInt(global.pannes || global.outages);

      var topLegacy = [];
      for (var j = 0; j < regions.length; j++) {
        var r = regions[j];
        topLegacy.push({
          region: r.nom || r.name || r.region || r.regionNom || 'Unknown',
          customers: this._safeInt(r.clients || r.customers || r.clientsSansElectricite),
          cause: '—',
          eta: null
        });
      }
      topLegacy.sort(function(a, b) { return b.customers - a.customers; });

      if (!totalLegacy) totalLegacy = topLegacy.reduce(function(s, r) { return s + r.customers; }, 0);
      if (!outagesLegacy) outagesLegacy = regions.length;

      return { data: { outages: outagesLegacy, customers: totalLegacy, top: topLegacy.slice(0, 10), source: 'Hydro-Québec' } };
    } catch (e) {
      return { error: e.message };
    }
  }
};

// Router export getPowerOutages() lives in Code.gs.

/**
 * Configure a utility's outage URL at runtime. Called from the Apps Script
 * editor or through google.script.run. Invalidates the cache so the new URL
 * takes effect on the next dashboard poll.
 *
 *   setOutageUrl('BC', 'https://...')     -> set BC Hydro URL
 *   setOutageUrl('ON', 'https://...')     -> set Hydro One URL (required to enable)
 *   setOutageUrl('QC', 'https://...')     -> set Hydro-Québec URL
 *   setOutageUrl('BC', '')                -> disable BC column
 *   setOutageUrl('ON', null)              -> clear override, revert to default
 */
function setOutageUrl(provinceCode, url) {
  var map = { BC: 'BC_HYDRO_URL', ON: 'HYDRO_ONE_URL', QC: 'HYDRO_QUEBEC_URL' };
  var key = map[String(provinceCode).toUpperCase()];
  if (!key) return 'Invalid province. Use BC, ON, or QC.';
  var props = PropertiesService.getScriptProperties();
  if (url === null || url === undefined) props.deleteProperty(key);
  else props.setProperty(key, String(url));
  try { CacheService.getScriptCache().remove('MRC_OUTAGES_V1'); } catch (e) {}
  return (url === null || url === undefined) ? ('Cleared ' + key) : ('Set ' + key + ' = ' + url);
}

/**
 * DIAGNOSTIC — run from the Apps Script editor (Run > debugOutages), then open
 * View > Logs. For each configured province it prints the HTTP status, the raw
 * response shape (top-level keys, feature count, whether the server capped the
 * result), and a SAMPLE feature's attribute keys + values. Paste the ON block
 * back and we can lock the customer-count field and confirm pagination against
 * the real Hydro One schema (same approach that fixed the weather feed).
 * Safe / read-only.
 */
function debugOutages() {
  var urls = {
    BC: OutageTracker._getUrl('BC_HYDRO_URL', OutageTracker._DEFAULT_BC_URL),
    ON: OutageTracker._getUrl('HYDRO_ONE_URL', OutageTracker._DEFAULT_HYDRO_ONE_URL),
    QC: OutageTracker._getUrl('HYDRO_QUEBEC_URL', OutageTracker._DEFAULT_HQ_URL)
  };
  var out = {};
  Object.keys(urls).forEach(function(prov) {
    var url = urls[prov];
    var info = { url: url || '(not set)' };
    if (url) {
      try {
        var res = UrlFetchApp.fetch(url, OutageTracker._commonHeaders());
        info.httpStatus = res.getResponseCode();
        var text = res.getContentText();
        info.bytes = text.length;
        var body = JSON.parse(text);
        info.topLevelType = Array.isArray(body) ? 'array' : typeof body;
        if (Array.isArray(body)) { info.arrayLength = body.length; info.sampleItem = body[0]; }
        else if (body && typeof body === 'object') {
          info.topLevelKeys = Object.keys(body).slice(0, 40);
          if (body.features) {
            info.featureCount = body.features.length;
            info.exceededTransferLimit = body.exceededTransferLimit || false;
            var attrs = (body.features[0] || {}).attributes || (body.features[0] || {}).properties || {};
            info.sampleAttributeKeys = Object.keys(attrs);
            info.sampleAttributes = attrs;
            info.detectedCustomers = OutageTracker._extractCustomers(attrs);
            info.detectedRegion = OutageTracker._extractRegion(attrs);
          }
          if (body.summaryFileData) info.summaryFileData = body.summaryFileData;
        }
      } catch (e) { info.error = e.message; }
    }
    out[prov] = info;
  });
  var s = JSON.stringify(out, null, 2);
  Logger.log(s);
  return s;
}

function getOutageUrls() {
  var props = PropertiesService.getScriptProperties();
  return JSON.stringify({
    BC: props.getProperty('BC_HYDRO_URL') || '(default)',
    ON: props.getProperty('HYDRO_ONE_URL') || '(not set — source disabled)',
    QC: props.getProperty('HYDRO_QUEBEC_URL') || '(default)'
  });
}

/**
 * Outage × Onshore Agent correlation.
 * Returns the subset of province-level outages above a customer threshold,
 * joined against MasterList Location (office city) and Raw Schedule
 * (shifts starting in the next N hours).
 *
 * Honest framing: this is OFFICE-LEVEL, not home address. The UI should
 * label it as a situational-awareness hint, not a per-agent alert.
 */
function getOutageAgentCorrelation(customerThreshold, lookaheadHours) {
  try {
    customerThreshold = customerThreshold || 5000;
    lookaheadHours = lookaheadHours || 2;

    var outageRaw = OutageTracker.fetchAll();
    var outages = JSON.parse(outageRaw);
    if (!outages || !outages.byProvince) return JSON.stringify({ banners: [] });

    // Province -> list of office locations we know about
    var officeByProv = {
      BC: ['VANCOUVER', 'BURNABY', 'VICTORIA', 'PRINCE GEORGE'],
      QC: ['MONTREAL', 'QUEBEC', 'LAVAL', 'GATINEAU', 'LONGUEUIL', 'RIMOUSKI', 'SHERBROOKE'],
      ON: ['OTTAWA', 'TORONTO', 'MISSISSAUGA', 'HAMILTON', 'CAMBRIDGE', 'LONDON']
    };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mlSheet = ss.getSheetByName('WF_MASTERLIST');
    var mlRows = (mlSheet && mlSheet.getLastRow() > 1)
      ? mlSheet.getRange(2, 1, mlSheet.getLastRow() - 1, 6).getDisplayValues()
      : [];

    // Count onshore agents per office province. Offshore agents are excluded
    // entirely — Canadian outages don't affect them.
    var agentsByProv = { BC: 0, QC: 0, ON: 0 };
    var mlByKey = {};
    mlRows.forEach(function(r) {
      var name = String(r[0] || '').trim();
      if (!name) return;
      var loc = String(r[4] || '').toUpperCase();
      if (!loc) return;
      // Skip offshore locations
      if (loc.includes('EL SALVADOR') || loc.includes('GUATEMALA') || loc.startsWith('TI')) return;
      // Consult registry; skip if manually flagged offshore
      if (typeof RegionRegistry !== 'undefined') {
        var rg = RegionRegistry.getRegion(name);
        if (rg === 'Offshore') return;
      }
      Object.keys(officeByProv).forEach(function(prov) {
        if (officeByProv[prov].indexOf(loc) !== -1) {
          agentsByProv[prov] = (agentsByProv[prov] || 0) + 1;
          mlByKey[(typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey(name) : name.toLowerCase()] = { name: name, prov: prov, loc: loc };
        }
      });
    });

    // Shift-start risk: count agents with shift start within next N hours
    // from the Raw Schedule. Limits to agents we've already mapped to a
    // province via MasterList.
    var startingByProv = { BC: 0, QC: 0, ON: 0 };
    var rawSheet = ss.getSheetByName('Raw Schedule');
    if (rawSheet && rawSheet.getLastRow() > 1) {
      var raw = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 12).getValues();
      var now = new Date().getTime();
      var until = now + lookaheadHours * 3600 * 1000;
      raw.forEach(function(row) {
        var name = String(row[0] || '').trim();
        if (!name) return;
        var startEpoch = Number(row[10]);
        if (!startEpoch) return;
        if (startEpoch < now || startEpoch > until) return;
        var key = (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey(name) : name.toLowerCase();
        var hit = mlByKey[key];
        if (hit) startingByProv[hit.prov] = (startingByProv[hit.prov] || 0) + 1;
      });
    }

    var banners = [];
    ['BC', 'QC', 'ON'].forEach(function(prov) {
      var o = outages.byProvince[prov];
      if (!o || o.customers < customerThreshold) return;
      banners.push({
        province: prov,
        source: o.source,
        customers: o.customers,
        outages: o.outages,
        topRegion: (o.top && o.top[0]) ? o.top[0].region : null,
        agentsInProvince: agentsByProv[prov] || 0,
        agentsStarting: startingByProv[prov] || 0,
        lookaheadHours: lookaheadHours
      });
    });
    banners.sort(function(a, b) { return b.customers - a.customers; });

    return JSON.stringify({ banners: banners, threshold: customerThreshold });
  } catch (e) {
    return JSON.stringify({ banners: [], error: e.message });
  }
}
