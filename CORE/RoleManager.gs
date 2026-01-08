/**
 * MODULE: ROLE MANAGER
 * Fast-access storage for Agent Status to bypass Sheet latency.
 * Uses ScriptProperties for instant updates across the floor.
 */

const RoleManager = {
  PROP_KEY: "MRC_FLOOR_STATUS_MAP",

  /**
   * Saves status instantly to memory, then backs up to Sheet.
   */
  setStatus: function(name, type, value) {
    const cleanName = String(name).trim().toLowerCase();
    
    // 1. FAST SAVE (Script Properties)
    try {
      const props = PropertiesService.getScriptProperties();
      let map = JSON.parse(props.getProperty(this.PROP_KEY) || "{}");
      
      if (!map[cleanName]) map[cleanName] = {};
      
      if (type === 'role') map[cleanName].role = (value === 'Active' ? "" : value);
      if (type === 'absent') map[cleanName].absent = value;
      
      // Update timestamp to keep data fresh
      map[cleanName].timestamp = new Date().getTime();
      
      // Garbage Collection: Remove entries older than 24 hours to keep property size low
      const now = new Date().getTime();
      for (const k in map) {
        if (now - map[k].timestamp > 86400000) delete map[k];
      }
      
      props.setProperty(this.PROP_KEY, JSON.stringify(map));
    } catch (e) {
      console.error("Fast Save Failed", e);
    }

    // 2. HARD SAVE (Sheet - Persistent)
    // We still call the original tracker so data survives resets
    if (typeof StatusTracker !== 'undefined') {
       StatusTracker.updateStatus(name, type, value);
    }

    return "Updated";
  },

  /**
   * Retrieves the fast map to merge with Sheet data.
   */
  getFastMap: function() {
    try {
      const json = PropertiesService.getScriptProperties().getProperty(this.PROP_KEY);
      return json ? JSON.parse(json) : {};
    } catch (e) {
      return {};
    }
  }
};
