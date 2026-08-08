/**
 * MODULE: ROLE MANAGER
 * Fast-access storage for Agent Status.
 * UPDATE: Supports MULTIPLE roles (Additive).
 */

const RoleManager = {
  PROP_KEY: "MRC_FLOOR_STATUS_MAP",

  setStatus: function(name, type, value) {
    const cleanName = String(name).trim().toLowerCase();
    let finalValue = value;

    // 1. FAST SAVE (Script Properties)
    try {
      const props = PropertiesService.getScriptProperties();
      let map = JSON.parse(props.getProperty(this.PROP_KEY) || "{}");
      
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

      names.forEach(function(name) {
        const cleanName = String(name).trim().toLowerCase();
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

  getFastMap: function() {
    try {
      const json = PropertiesService.getScriptProperties().getProperty(this.PROP_KEY);
      return json ? JSON.parse(json) : {};
    } catch (e) { return {}; }
  }
};
