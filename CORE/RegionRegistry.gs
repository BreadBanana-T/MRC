/**
 * MODULE: REGION REGISTRY
 * Single source of truth for agent region (Onshore / Offshore).
 *
 * Prior logic re-derived region on every read from the Raw Schedule row,
 * which meant a single weak signal (UNAB code, missing "TI " prefix, etc.)
 * could silently flip an agent. This module persists each determination
 * in a dedicated sheet with a source tag, and enforces a hierarchy so
 * cheap signals cannot override trusted ones:
 *
 *   manual   > masterlist > auto-wfm-id > auto-wfm-keyword > (nothing)
 *
 * Once an offshore agent is detected by a high-confidence signal (ID
 * prefix 3, MasterList marker), they stay offshore forever unless a
 * manual override says otherwise.
 *
 * Positive signals only. UNAB / SICK / lack-of-UNAB are NEVER used as
 * region inputs.
 */

var RegionRegistry = {
  SHEET: 'WF_REGION_MAP',
  HEADERS: ['Agent Key', 'Display Name', 'Region', 'Source', 'Last Confirmed'],

  _SOURCE_RANK: { 'manual': 4, 'masterlist': 3, 'auto-wfm-id': 2, 'auto-wfm-keyword': 1 },

  _cache: null,
  _cacheTs: 0,
  _CACHE_TTL_MS: 30000,

  _getSheet: function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(this.SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(this.SHEET);
      sheet.appendRow(this.HEADERS);
      sheet.getRange(1, 1, 1, this.HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    return sheet;
  },

  _loadMap: function() {
    var now = new Date().getTime();
    if (this._cache && (now - this._cacheTs) < this._CACHE_TTL_MS) return this._cache;
    var sheet = this._getSheet();
    var map = {};
    if (sheet.getLastRow() > 1) {
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, this.HEADERS.length).getValues();
      for (var i = 0; i < data.length; i++) {
        var r = data[i];
        if (!r[0]) continue;
        map[String(r[0])] = {
          key: String(r[0]),
          display: String(r[1] || ''),
          region: String(r[2] || 'Onshore'),
          source: String(r[3] || 'auto-wfm-keyword'),
          lastConfirmed: r[4] || null
        };
      }
    }
    this._cache = map;
    this._cacheTs = now;
    return map;
  },

  _invalidate: function() { this._cache = null; this._cacheTs = 0; },

  /** Returns "Onshore", "Offshore", or null if agent unknown. */
  getRegion: function(agentName) {
    if (!agentName) return null;
    var key = (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey(agentName) : String(agentName).trim().toLowerCase();
    var m = this._loadMap();
    return m[key] ? m[key].region : null;
  },

  /** Returns the source tag or null. */
  getSource: function(agentName) {
    if (!agentName) return null;
    var key = (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey(agentName) : String(agentName).trim().toLowerCase();
    var m = this._loadMap();
    return m[key] ? m[key].source : null;
  },

  /**
   * Write/update an entry. Hierarchy rule: if existing source outranks the
   * new source, the write is refused. A null existing source is rank 0.
   * Returns true if the row was written.
   */
  upsert: function(agentName, region, source) {
    if (!agentName || !region) return false;
    if (region !== 'Onshore' && region !== 'Offshore') return false;
    source = source || 'auto-wfm-keyword';

    var key = (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey(agentName) : String(agentName).trim().toLowerCase();
    if (!key) return false;

    var m = this._loadMap();
    var existing = m[key];
    if (existing) {
      var oldRank = this._SOURCE_RANK[existing.source] || 0;
      var newRank = this._SOURCE_RANK[source] || 0;
      if (newRank < oldRank) return false;
      // Same rank + same region = refresh timestamp only. Different region at same rank = accept (latest wins).
    }

    var sheet = this._getSheet();
    var lastRow = sheet.getLastRow();
    var rowNum = -1;
    if (lastRow > 1) {
      var keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (String(keys[i][0]) === key) { rowNum = i + 2; break; }
      }
    }

    var display = String(agentName).trim();
    var now = new Date();
    if (rowNum === -1) {
      sheet.appendRow([key, display, region, source, now]);
    } else {
      sheet.getRange(rowNum, 1, 1, 5).setValues([[key, display, region, source, now]]);
    }
    this._invalidate();
    return true;
  },

  /** Manual upsert always wins. Used by the Region Manager UI. */
  upsertManual: function(agentName, region) {
    return this.upsert(agentName, region, 'manual');
  },

  remove: function(agentName) {
    if (!agentName) return false;
    var key = (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey(agentName) : String(agentName).trim().toLowerCase();
    var sheet = this._getSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;
    var keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = keys.length - 1; i >= 0; i--) {
      if (String(keys[i][0]) === key) { sheet.deleteRow(i + 2); this._invalidate(); return true; }
    }
    return false;
  },

  /** Returns array of {key, display, region, source, lastConfirmed}. */
  list: function() {
    var m = this._loadMap();
    var out = [];
    Object.keys(m).forEach(function(k) { out.push(m[k]); });
    out.sort(function(a, b) {
      if (a.region !== b.region) return a.region === 'Offshore' ? -1 : 1;
      return a.display.localeCompare(b.display);
    });
    return out;
  }
};

// ---- Global wrappers for google.script.run ----
function getRegionRegistry() { return JSON.stringify(RegionRegistry.list()); }
function setAgentRegionManual(name, region) {
  var ok = RegionRegistry.upsertManual(name, region);
  return ok ? 'OK' : 'Rejected';
}
function clearAgentRegion(name) {
  var ok = RegionRegistry.remove(name);
  return ok ? 'OK' : 'Not found';
}
