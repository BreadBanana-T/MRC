/**
 * MODULE: ROLE MANAGER
 * Fast-access storage for Agent Status.
 * UPDATE: Supports MULTIPLE roles (Additive).
 */

const RoleManager = {
  PROP_KEY: "MRC_FLOOR_STATUS_MAP",

  // ── AGENT KEY ─────────────────────────────────────────────────────────────
  // Everything that stores or reads floor state MUST key on the same thing.
  // AgentMonitor reads with _normalizeAgentKey (token-SORTED, accent/comma
  // stripped), so writing a plain lowercase key meant "Bennani, Mohammed" was
  // saved as `bennani, mohammed` and looked up as `bennani mohammed` — the
  // assignment was stored fine and then never found again, which is what made
  // roles and Genesys assignments silently revert a couple of minutes later.
  _key: function(name) {
    return (typeof _normalizeAgentKey === 'function')
      ? _normalizeAgentKey(name)
      : String(name).trim().toLowerCase();
  },
  // The key shape written before that fix, kept so existing assignments in the
  // property map are migrated instead of being dropped on deploy.
  _legacyKey: function(name) { return String(name).trim().toLowerCase(); },

  setStatus: function(name, type, value) {
    const cleanName = this._key(name);
    const legacyName = this._legacyKey(name);
    let finalValue = value;

    // 1. FAST SAVE (Script Properties)
    try {
      const props = PropertiesService.getScriptProperties();
      let map = JSON.parse(props.getProperty(this.PROP_KEY) || "{}");

      // Carry a pre-fix entry over to the correct key rather than stranding it.
      if (!map[cleanName] && legacyName !== cleanName && map[legacyName]) {
        map[cleanName] = map[legacyName];
        delete map[legacyName];
      }
      if (!map[cleanName]) map[cleanName] = {};

      if (type === 'role') {
          if (value === 'Active') {
              // Clear Roles
              map[cleanName].role = "";
              finalValue = "";
          } else {
              // Additive Logic: Check if already exists
              let current = map[cleanName].role || "";
              if (!current.includes(value)) {
                  // Append new role
                  map[cleanName].role = (current + " " + value).trim();
              }
              finalValue = map[cleanName].role;
          }
      }
      
      if (type === 'absent') map[cleanName].absent = value;

      map[cleanName].timestamp = new Date().getTime();

      // Garbage Collection (24h) — Genesys assignments are exempt: they
      // persist until the agent is explicitly moved back to Avaya.
      const now = new Date().getTime();
      for (const k in map) {
        if (now - map[k].timestamp > 86400000 && !map[k].genesys) delete map[k];
      }

      props.setProperty(this.PROP_KEY, JSON.stringify(map));
    } catch (e) {
      console.error("Fast Save Failed", e);
    }

    // 2. HARD SAVE (Pass the COMBINED value to Sheet)
    if (typeof StatusTracker !== 'undefined') {
       StatusTracker.updateStatus(name, type, finalValue);
    }

    return "Updated";
  },

  // Genesys Cloud manual assignment. channel: 'inbound' | 'outbound' | ''
  // (empty = back to Avaya). Outbound carries no queues. Persists until
  // changed — exempt from the fast-map GC and hard-saved to a dedicated,
  // date-agnostic sheet so it survives day to day.
  setGenesys: function(name, channel, queuesJson) {
    return this.setGenesysBatch(JSON.stringify([name]), channel, queuesJson);
  },

  // Batch variant: one Script-Properties read-modify-write for the whole
  // group (parallel single saves would clobber each other on the shared
  // property), then one consolidated sheet update.
  setGenesysBatch: function(namesJson, channel, queuesJson) {
    let names = [];
    try { names = Array.isArray(namesJson) ? namesJson : JSON.parse(namesJson || "[]"); } catch (e) { names = []; }
    names = names.filter(function(n) { return n && String(n).trim(); });
    if (!names.length) return "No agents";

    channel = channel || "";
    let queues = [];
    try { queues = Array.isArray(queuesJson) ? queuesJson : JSON.parse(queuesJson || "[]"); } catch (e) { queues = []; }
    if (channel !== 'inbound') queues = [];

    // 1. FAST SAVE (single write for all names)
    try {
      const props = PropertiesService.getScriptProperties();
      let map = JSON.parse(props.getProperty(this.PROP_KEY) || "{}");
      const now = new Date().getTime();

      const self = this;
      names.forEach(function(name) {
        const cleanName = self._key(name);
        const legacyName = self._legacyKey(name);
        if (!map[cleanName] && legacyName !== cleanName && map[legacyName]) {
          map[cleanName] = map[legacyName];
          delete map[legacyName];
        }
        if (!map[cleanName]) map[cleanName] = {};
        if (!channel) delete map[cleanName].genesys;
        else map[cleanName].genesys = { channel: channel, queues: queues };
        map[cleanName].timestamp = now;
      });

      for (const k in map) {
        if (now - map[k].timestamp > 86400000 && !map[k].genesys) delete map[k];
      }
      props.setProperty(this.PROP_KEY, JSON.stringify(map));
    } catch (e) {
      console.error("Genesys Fast Save Failed", e);
    }

    // 2. HARD SAVE (one sheet pass for the whole batch)
    if (typeof StatusTracker !== 'undefined') {
       StatusTracker.setGenesysBatch(names, channel, JSON.stringify(queues));
    }

    return "Updated " + names.length;
  },

  // Always hand back a map keyed the way AgentMonitor looks things up. Entries
  // written before the key fix are re-keyed here on the way out, so an existing
  // assignment keeps working without needing a write first. If two spellings
  // collapse onto one key, the more recent entry wins.
  getFastMap: function() {
    try {
      const json = PropertiesService.getScriptProperties().getProperty(this.PROP_KEY);
      const raw = json ? JSON.parse(json) : {};
      const out = {};
      for (const k in raw) {
        const nk = this._key(k);
        const prev = out[nk];
        if (!prev || (raw[k] && (raw[k].timestamp || 0) >= (prev.timestamp || 0))) out[nk] = raw[k];
      }
      return out;
    } catch (e) { return {}; }
  }
};
