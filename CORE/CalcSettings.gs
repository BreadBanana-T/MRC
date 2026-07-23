/**
 * MODULE: CALC SETTINGS
 *
 * Runtime-editable store for the Calculator's report template, so the layout
 * (the "(Genesys)" suffixes, the NOTES line, DELAYS, labels, …) can change on
 * the fly without redeploying the Apps Script. Follows the same Script-Property
 * JSON pattern the app already uses (RoleManager floor map, FeedbackTracker
 * config). The default lives in code as DEFAULT_CALC_TEMPLATE (CORE/Calculator.gs)
 * so a missing or blank property always renders a valid report.
 *
 * Routers getCalcTemplate() / saveCalcTemplate() / resetCalcTemplate() live in
 * Code.gs; the edit UI is in the Calculator view (JavaScript.html).
 */
var CalcSettings = {
  KEY: 'CALC_TEMPLATE',

  // Stored custom template, or "" when none is saved.
  raw: function() {
    return PropertiesService.getScriptProperties().getProperty(this.KEY) || "";
  },

  // The template to render with: the saved one if non-blank, else the default.
  template: function() {
    var t = this.raw();
    return (t && t.trim() !== "") ? t : DEFAULT_CALC_TEMPLATE;
  },

  // Persist a new template. Rejects empty input so the report can't be blanked.
  save: function(t) {
    if (typeof t !== 'string' || t.trim() === "") {
      return JSON.stringify({ ok: false, error: "Template can't be empty." });
    }
    PropertiesService.getScriptProperties().setProperty(this.KEY, t);
    return JSON.stringify({ ok: true, template: t, custom: true });
  },

  // Drop the custom template and fall back to the built-in default.
  reset: function() {
    PropertiesService.getScriptProperties().deleteProperty(this.KEY);
    return JSON.stringify({ ok: true, template: DEFAULT_CALC_TEMPLATE, custom: false });
  },

  // What the edit UI needs: current template, the default (for Reset preview),
  // and whether a custom template is currently in effect.
  payload: function() {
    return JSON.stringify({
      template: this.template(),
      default: DEFAULT_CALC_TEMPLATE,
      custom: this.raw().trim() !== ""
    });
  }
};
