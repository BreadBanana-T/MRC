/**
 * MODULE: CALC SETTINGS  (report-template library)
 *
 * Runtime-editable store for the Calculator's report templates, so the layout
 * ("(Genesys)" suffixes, NOTES wording, DELAYS, labels, …) can change on the
 * fly without redeploying. Holds a LIBRARY of named templates plus which one
 * is active; the report is rendered from the active template's body.
 *
 * Stored as one JSON blob in a Script Property (same pattern as RoleManager /
 * FeedbackTracker):
 *   { active: "<name>|''", templates: [ { name, body }, ... ] }
 * active === "" means "use the built-in default" (DEFAULT_CALC_TEMPLATE in
 * CORE/Calculator.gs), which is always available and can't be edited/deleted.
 *
 * Routers live in Code.gs; the edit UI is in the Calculator view.
 */
var CalcSettings = {
  KEY: 'CALC_LIBRARY',
  OLD_KEY: 'CALC_TEMPLATE',   // legacy single-template key (auto-migrated)

  _load: function() {
    var sp = PropertiesService.getScriptProperties();
    var lib = null;
    var raw = sp.getProperty(this.KEY);
    if (raw) { try { lib = JSON.parse(raw); } catch (e) { lib = null; } }
    if (!lib || !Array.isArray(lib.templates)) {
      // First run — seed from the legacy single-template property if present.
      lib = { active: "", templates: [] };
      var old = sp.getProperty(this.OLD_KEY);
      if (old && old.trim() !== "") {
        lib.templates.push({ name: "Saved template", body: old });
        lib.active = "Saved template";
      }
    }
    if (typeof lib.active !== 'string') lib.active = "";
    return lib;
  },

  _persist: function(lib) {
    PropertiesService.getScriptProperties().setProperty(this.KEY, JSON.stringify(lib));
  },

  _find: function(lib, name) {
    for (var i = 0; i < lib.templates.length; i++) if (lib.templates[i].name === name) return lib.templates[i];
    return null;
  },

  // Success shape shared by every call the UI makes.
  _state: function(lib, extra) {
    var hasOverride = !!(lib.defaultBody && lib.defaultBody.trim() !== "");
    var o = {
      ok: true, active: lib.active, templates: lib.templates,
      default: DEFAULT_CALC_TEMPLATE,                                   // factory text (for "load factory")
      defaultBody: hasOverride ? lib.defaultBody : DEFAULT_CALC_TEMPLATE, // effective Default (editable)
      defaultCustom: hasOverride
    };
    if (extra) for (var k in extra) o[k] = extra[k];
    return JSON.stringify(o);
  },

  // The effective built-in Default: the saved override, else the factory text.
  _default: function(lib) {
    return (lib.defaultBody && lib.defaultBody.trim() !== "") ? lib.defaultBody : DEFAULT_CALC_TEMPLATE;
  },

  // Body used to render the report: the active named template, else the Default.
  template: function() {
    var lib = this._load();
    if (lib.active === "") return this._default(lib);
    var t = this._find(lib, lib.active);
    return (t && t.body && t.body.trim() !== "") ? t.body : this._default(lib);
  },

  // Edit the built-in Default itself (stored as an override; factory text is
  // always recoverable via resetDefault / the "load factory text" button).
  saveDefault: function(body) {
    if (typeof body !== 'string' || body.trim() === "") return JSON.stringify({ ok: false, error: "Template can't be empty." });
    var lib = this._load();
    lib.defaultBody = body;
    lib.active = "";   // saving the Default implies you want it in use
    this._persist(lib);
    return this._state(lib);
  },

  // Restore the built-in Default to the factory text.
  resetDefault: function() {
    var lib = this._load();
    delete lib.defaultBody;
    this._persist(lib);
    return this._state(lib);
  },

  // Full library state for the editor.
  payload: function() {
    return this._state(this._load());
  },

  // Create or update a named template (upsert). Optionally make it active.
  saveTemplate: function(name, body, makeActive) {
    name = (name || "").trim();
    if (name === "") return JSON.stringify({ ok: false, error: "Template name is required." });
    // Saving under the name "Default" edits the built-in Default itself
    // (that's what the user wants when they type it), rather than erroring.
    if (name.toLowerCase() === "default") return this.saveDefault(body);
    if (typeof body !== 'string' || body.trim() === "") return JSON.stringify({ ok: false, error: "Template can't be empty." });
    var lib = this._load();
    var t = this._find(lib, name);
    if (t) t.body = body; else lib.templates.push({ name: name, body: body });
    if (makeActive || lib.active === "") lib.active = name;
    this._persist(lib);
    return this._state(lib, { savedName: name });
  },

  // Switch the active template. "" selects the built-in default.
  setActive: function(name) {
    var lib = this._load();
    if (name !== "" && !this._find(lib, name)) return JSON.stringify({ ok: false, error: "No such template." });
    lib.active = name;
    this._persist(lib);
    return this._state(lib);
  },

  rename: function(oldName, newName) {
    newName = (newName || "").trim();
    if (newName === "") return JSON.stringify({ ok: false, error: "New name is required." });
    if (newName.toLowerCase() === "default") return JSON.stringify({ ok: false, error: '"Default" is reserved.' });
    var lib = this._load();
    if (this._find(lib, newName)) return JSON.stringify({ ok: false, error: "A template with that name already exists." });
    var t = this._find(lib, oldName);
    if (!t) return JSON.stringify({ ok: false, error: "No such template." });
    t.name = newName;
    if (lib.active === oldName) lib.active = newName;
    this._persist(lib);
    return this._state(lib);
  },

  remove: function(name) {
    var lib = this._load();
    lib.templates = lib.templates.filter(function (x) { return x.name !== name; });
    if (lib.active === name) lib.active = "";   // fall back to built-in default
    this._persist(lib);
    return this._state(lib);
  }
};
