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
      
      // Garbage Collection (24h)
      const now = new Date().getTime();
      for (const k in map) {
        if (now - map[k].timestamp > 86400000) delete map[k];
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

  getFastMap: function() {
    try {
      const json = PropertiesService.getScriptProperties().getProperty(this.PROP_KEY);
      return json ? JSON.parse(json) : {};
    } catch (e) { return {}; }
  }
};
