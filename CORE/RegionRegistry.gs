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

  // Batch mode: accumulate upserts in memory during a bulk import, flush
  // them all in a single sheet write at the end. Without this, each agent
  // costs 2 sheet API round-trips (~100-500ms each) and big imports hang
  // for minutes.
  _batchMode: false,
  _batchChanges: null,

  beginBatch: function() {
    this._batchMode = true;
    this._batchChanges = {};
    this._loadMap(); // pre-warm cache so per-row reads are free
  },

  commitBatch: function() {
    if (!this._batchMode) return 0;
    var changes = this._batchChanges || {};
    this._batchMode = false;
    this._batchChanges = null;
    var changeKeys = Object.keys(changes);
    if (!changeKeys.length) return 0;
    var sheet = this._getSheet();
    var lastRow = sheet.getLastRow();
    var existing = (lastRow > 1) ? sheet.getRange(2, 1, lastRow - 1, this.HEADERS.length).getValues() : [];
    var indexByKey = {};
    for (var i = 0; i < existing.length; i++) indexByKey[String(existing[i][0])] = i;
    var added = [];
    for (var k = 0; k < changeKeys.length; k++) {
      var key = changeKeys[k];
      var c = changes[key];
      var row = [key, c.display, c.region, c.source, c.lastConfirmed];
      if (indexByKey[key] !== undefined) existing[indexByKey[key]] = row;
      else added.push(row);
    }
    if (existing.length) sheet.getRange(2, 1, existing.length, this.HEADERS.length).setValues(existing);
    if (added.length) sheet.getRange(sheet.getLastRow() + 1, 1, added.length, this.HEADERS.length).setValues(added);
    this._invalidate();
    return changeKeys.length;
  },

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
    }

    var display = String(agentName).trim();
    var now = new Date();

    // Batch path: stash in memory + update cache so subsequent reads in
    // the same batch see the new value. Single sheet write happens later
    // in commitBatch().
    if (this._batchMode) {
      this._batchChanges[key] = { display: display, region: region, source: source, lastConfirmed: now };
      m[key] = { key: key, display: display, region: region, source: source, lastConfirmed: now };
      return true;
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

/**
 * One-shot cleanup: scans WF_REGION_MAP, regroups rows by the current
 * normalized key (which now folds "Last, First" and "First Last" into
 * the same hash), and keeps one row per agent picking the highest-
 * priority source (manual > masterlist > auto-wfm-id > auto-wfm-keyword)
 * with most recent timestamp as tiebreaker.
 *
 * Run once from the Apps Script editor after pulling the new code.
 * Returns a summary like "Deduped 312 rows → 287 unique agents".
 */
function dedupeRegionRegistry() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('WF_REGION_MAP');
  if (!sheet || sheet.getLastRow() < 2) return 'No data to dedupe.';
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  var groups = {};
  rows.forEach(function(r) {
    if (!r[0] && !r[1]) return; // skip blank
    var displayName = String(r[1] || r[0] || '').trim();
    var newKey = (typeof _normalizeAgentKey === 'function')
                    ? _normalizeAgentKey(displayName)
                    : String(displayName).toLowerCase().trim();
    if (!newKey) return;
    if (!groups[newKey]) groups[newKey] = [];
    groups[newKey].push({
      display: displayName,
      region: String(r[2] || 'Onshore'),
      source: String(r[3] || 'auto-wfm-keyword'),
      ts: r[4] || null
    });
  });
  var rank = { 'manual': 4, 'masterlist': 3, 'auto-wfm-id': 2, 'auto-wfm-keyword': 1, 'auto-wfm-default': 0 };
  var winners = [];
  var dupesRemoved = 0;
  Object.keys(groups).forEach(function(k) {
    var g = groups[k];
    if (g.length > 1) dupesRemoved += g.length - 1;
    g.sort(function(a, b) {
      var d = (rank[b.source] || 0) - (rank[a.source] || 0);
      if (d !== 0) return d;
      return (new Date(b.ts || 0).getTime()) - (new Date(a.ts || 0).getTime());
    });
    var w = g[0];
    // Prefer the longer/more readable display name among the duplicates.
    var bestDisplay = g.reduce(function(acc, x) {
      return (x.display.length > acc.length) ? x.display : acc;
    }, w.display);
    winners.push([k, bestDisplay, w.region, w.source, w.ts || new Date()]);
  });
  // Replace sheet contents. clearContent (not deleteRows) avoids the
  // "cannot delete all non-frozen rows" error Sheets throws when the
  // dedupe leaves the sheet temporarily empty.
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn() || 5).clearContent();
  if (winners.length) sheet.getRange(2, 1, winners.length, 5).setValues(winners);
  RegionRegistry._invalidate();
  return 'Deduped ' + rows.length + ' rows → ' + winners.length + ' unique agents (' + dupesRemoved + ' duplicates removed).';
}
