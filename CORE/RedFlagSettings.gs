/**
 * MODULE: RED FLAG SETTINGS  (runtime-editable thresholds)
 *
 * Every number the Red Flags rules test against, editable from the dashboard
 * without a redeploy. Stored as one JSON blob in a Script Property — same
 * pattern as CalcSettings / RoleManager.
 *
 * Defaults mirror the Assignment Tracker 2026 workbook exactly
 * (docs/DATA_FLOW.md §2.2), so an untouched install behaves like the sheet:
 *
 *   Inbound Red Flags     In%  >= inPct   AND Tasks/Calls >= tasksXCalls
 *   Inbound Out of Scope  In%  >= inPct   AND Tasks/Calls <  tasksXCalls
 *   Outbound Red Flags    Out% >= outPct  AND ERC level  >= outMinLevel
 *   CP/H Red Flags        CPH  <= cph
 *   all four              Worked days >= minWorkedDays
 *
 * WHY THIS IS EDITABLE, and the trap: when In%/Out% are ESTIMATED from the AUX
 * report they run ~13pp high, so raising `inPct` is a legitimate short-term
 * correction — but it is a correction for BAD INPUT, not a rule change. Once a
 * GEM export with real Inbound/Outbound Time % columns is imported, put it back
 * to 0.75 or the flags will under-report. The dashboard says which mode the
 * month is in (measured vs estimated) right next to these numbers.
 */
var RedFlagSettings = {
  KEY: 'RF_THRESHOLDS',

  DEFAULTS: {
    inPct: 0.75,          // Inbound / Out of Scope entry gate
    tasksXCalls: 0.50,    // splits Inbound Red Flag from Out of Scope
    outPct: 0.75,         // Outbound Red Flags
    outMinLevel: 2,       // Outbound only applies at this ERC level or above
    cph: 7.3,             // CP/H at or below this is flagged
    minWorkedDays: 5      // every rule ignores agents below this
  },

  // Bounds keep a typo from silently emptying or flooding the dashboard.
  LIMITS: {
    inPct:         { min: 0.30, max: 1,   pct: true,  label: 'Inbound %',        hint: 'In% at or above this enters the Inbound / Out of Scope pool' },
    tasksXCalls:   { min: 0,    max: 1,   pct: true,  label: 'Tasks / Calls',    hint: 'At or above → Inbound Red Flag; below → Out of Scope' },
    outPct:        { min: 0.30, max: 1,   pct: true,  label: 'Outbound %',       hint: 'Out% at or above this is flagged' },
    outMinLevel:   { min: 0,    max: 9,   pct: false, label: 'Outbound min ERC', hint: 'Outbound rule only applies from this level up' },
    cph:           { min: 0,    max: 40,  pct: false, label: 'CP/H at or below', hint: 'Contacts per hour at or below this is flagged' },
    minWorkedDays: { min: 0,    max: 31,  pct: false, label: 'Min worked days',  hint: 'Agents below this are ignored by every rule' }
  },

  get: function() {
    var out = {};
    for (var k in this.DEFAULTS) out[k] = this.DEFAULTS[k];
    try {
      var raw = PropertiesService.getScriptProperties().getProperty(this.KEY);
      if (raw) {
        var saved = JSON.parse(raw);
        for (var k2 in this.DEFAULTS) {
          if (saved[k2] === undefined || saved[k2] === null || saved[k2] === '') continue;
          var v = parseFloat(saved[k2]);
          if (isNaN(v)) continue;
          out[k2] = this._clamp(k2, v);
        }
      }
    } catch (e) {}
    return out;
  },

  _clamp: function(key, v) {
    var lim = this.LIMITS[key];
    if (!lim) return v;
    return Math.max(lim.min, Math.min(lim.max, v));
  },

  /**
   * @param {string} json  {inPct, tasksXCalls, outPct, outMinLevel, cph, minWorkedDays}
   *   Percent fields accept either a fraction (0.75) or a percentage (75).
   */
  save: function(json) {
    try {
      var incoming = (typeof json === 'string') ? JSON.parse(json) : (json || {});
      var next = {};
      for (var k in this.DEFAULTS) {
        var v = incoming[k];
        if (v === undefined || v === null || v === '') { next[k] = this.DEFAULTS[k]; continue; }
        v = parseFloat(String(v).replace('%', '').trim());
        if (isNaN(v)) { next[k] = this.DEFAULTS[k]; continue; }
        // 75 and 0.75 both mean 75% — a manager typing "88" must not become 8800%.
        if (this.LIMITS[k] && this.LIMITS[k].pct && v > 1) v = v / 100;
        next[k] = this._clamp(k, v);
      }
      PropertiesService.getScriptProperties().setProperty(this.KEY, JSON.stringify(next));
      if (typeof _bustWfCache === 'function') _bustWfCache();
      return JSON.stringify({ ok: true, thresholds: next, isDefault: this._isDefault(next) });
    } catch (e) {
      return JSON.stringify({ ok: false, error: e.message });
    }
  },

  reset: function() {
    try { PropertiesService.getScriptProperties().deleteProperty(this.KEY); } catch (e) {}
    if (typeof _bustWfCache === 'function') _bustWfCache();
    return JSON.stringify({ ok: true, thresholds: this.get(), isDefault: true });
  },

  _isDefault: function(t) {
    for (var k in this.DEFAULTS) if (Math.abs((t[k] || 0) - this.DEFAULTS[k]) > 1e-9) return false;
    return true;
  },

  /** For the editor UI: current values, the built-in defaults, and the bounds. */
  payload: function() {
    var t = this.get();
    return JSON.stringify({ ok: true, thresholds: t, defaults: this.DEFAULTS, limits: this.LIMITS, isDefault: this._isDefault(t) });
  }
};
