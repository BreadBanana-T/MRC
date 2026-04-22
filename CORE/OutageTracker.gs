/**
 * MODULE: OUTAGE TRACKER
 * Pulls real-time power outage counts for BC, ON (Hydro One), and QC
 * (Hydro-Québec). All endpoints are publicly served (no API key, no auth).
 *
 * - BC Hydro: documented outage-list.json
 * - Hydro One: public ArcGIS FeatureServer powering their Storm Centre map
 * - Hydro-Québec: undocumented but publicly-served aggregate used by their
 *   info-pannes web map. If HQ changes their URL, update _HQ_URL below;
 *   inspect info-pannes.hydroquebec.com network tab to find the new path.
 *
 * Payload is cached in CacheService for 5 minutes so polling the dashboard
 * doesn't spam the utilities or eat UrlFetch quota.
 */

var OutageTracker = {
  CACHE_KEY: 'MRC_OUTAGES_V1',
  CACHE_TTL_SECONDS: 300,

  _BC_URL: 'https://www.bchydro.com/power-outages/app/outage-list.json',
  _HYDRO_ONE_URL: 'https://services1.arcgis.com/qAo1OsXi67t7XgmS/arcgis/rest/services/Outages_External_View/FeatureServer/0/query?where=1%3D1&outFields=*&returnGeometry=false&f=json&resultRecordCount=1000',
  _HQ_URL: 'https://www.hydroquebec.com/data/pannes/donnees/bilan.json',

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
    else result.byProvince.BC = bc.data;

    var on = this._fetchHydroOne();
    if (on.error) result.errors.push('ON: ' + on.error);
    else result.byProvince.ON = on.data;

    var qc = this._fetchHQ();
    if (qc.error) result.errors.push('QC: ' + qc.error);
    else result.byProvince.QC = qc.data;

    var payload = JSON.stringify(result);
    try { CacheService.getScriptCache().put(this.CACHE_KEY, payload, this.CACHE_TTL_SECONDS); } catch (e) {}
    return payload;
  },

  _safeInt: function(v) {
    var n = parseInt(String(v == null ? 0 : v).replace(/[^\d-]/g, ''), 10);
    return isNaN(n) ? 0 : n;
  },

  _commonHeaders: function() {
    return {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'MRC-Ops-Dashboard/1.0', 'Accept': 'application/json' }
    };
  },

  _fetchBCHydro: function() {
    try {
      var res = UrlFetchApp.fetch(this._BC_URL, this._commonHeaders());
      if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
      var body = JSON.parse(res.getContentText());
      var list = Array.isArray(body) ? body : (body.outages || body.data || []);

      var total = 0;
      var top = [];
      for (var i = 0; i < list.length; i++) {
        var o = list[i];
        var n = this._safeInt(o.numCustomersOut || o.customersAffected || o.customers || o.customerCount);
        total += n;
        top.push({
          region: o.area || o.region || o.location || o.regionNm || 'Unknown',
          customers: n,
          cause: o.cause || o.reason || '—',
          eta: o.crewEta || o.estTimeOn || o.estimatedRestoration || null
        });
      }
      top.sort(function(a, b) { return b.customers - a.customers; });
      return { data: { outages: list.length, customers: total, top: top.slice(0, 10), source: 'BC Hydro' } };
    } catch (e) {
      return { error: e.message };
    }
  },

  _fetchHydroOne: function() {
    try {
      var res = UrlFetchApp.fetch(this._HYDRO_ONE_URL, this._commonHeaders());
      if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
      var body = JSON.parse(res.getContentText());
      if (body.error) throw new Error('ArcGIS: ' + (body.error.message || 'unknown'));
      var features = body.features || [];

      var total = 0;
      var top = [];
      for (var i = 0; i < features.length; i++) {
        var a = features[i].attributes || {};
        var n = this._safeInt(a.NumCustomersAffected || a.CUSTOMERS_AFFECTED || a.customersAffected || a.customers);
        total += n;
        top.push({
          region: a.MunicipalityName || a.Municipality || a.LocationDescription || a.Town || a.region || 'Unknown',
          customers: n,
          cause: a.Cause || a.CauseDescription || a.CauseCategory || '—',
          eta: a.EstimatedRestoration || a.ETR || a.etr || null
        });
      }
      top.sort(function(a, b) { return b.customers - a.customers; });
      return { data: { outages: features.length, customers: total, top: top.slice(0, 10), source: 'Hydro One (rural ON)' } };
    } catch (e) {
      return { error: e.message };
    }
  },

  _fetchHQ: function() {
    try {
      var res = UrlFetchApp.fetch(this._HQ_URL, this._commonHeaders());
      if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
      var body = JSON.parse(res.getContentText());

      // HQ bilan.json historically: { global:{clients,pannes}, regions:[{nom,clients,pannes}] }
      // Shape occasionally drifts; be defensive.
      var regions = body.regions || body.list || body.data || [];
      var global = body.global || body.summary || body.total || {};

      var total = this._safeInt(global.clients || global.clientsAffected || global.clientsSansElectricite);
      var outages = this._safeInt(global.pannes || global.outages);

      var top = [];
      for (var i = 0; i < regions.length; i++) {
        var r = regions[i];
        top.push({
          region: r.nom || r.name || r.region || r.regionNom || 'Unknown',
          customers: this._safeInt(r.clients || r.customers || r.clientsSansElectricite),
          cause: '—',
          eta: null
        });
      }
      top.sort(function(a, b) { return b.customers - a.customers; });

      if (!total) total = top.reduce(function(s, r) { return s + r.customers; }, 0);
      if (!outages) outages = regions.length;

      return { data: { outages: outages, customers: total, top: top.slice(0, 10), source: 'Hydro-Québec (unofficial)' } };
    } catch (e) {
      return { error: e.message };
    }
  }
};

function getPowerOutages() { return OutageTracker.fetchAll(); }

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
