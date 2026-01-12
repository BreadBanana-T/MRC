const MasterConnector = {
  // *** PASTE YOUR MASTER DB ID HERE ***
  DB_ID: "1RZOdFH3q0Zvgp2L2LbYD9NzpsfmzLGZNyd7O7j81Q-U", 
  
  THRESHOLDS: { 'SICK': {days:30,limit:3}, 'NCNS': {days:90,limit:1}, 'AWOL': {days:90,limit:1} },

  logAbsence: function(agentName, type) {
    if (!type) return;
    try {
      const ss = SpreadsheetApp.openById(this.DB_ID);
      const sheet = ss.getSheetByName("Attendance_Log");
      const flagCheck = this._checkHistory(sheet, agentName, type);
      
      sheet.appendRow([new Date(), agentName, type, flagCheck.isFlagged?"🚩 FLAGGED":"OK", SpreadsheetApp.getActiveSpreadsheet().getName()]);
      if (flagCheck.isFlagged) this._createFlag(ss, agentName, type, flagCheck.count);
    } catch (e) { console.error("DB Error", e); }
  },

  logStats: function(svl, ack, law, idp) {
    try {
      const ss = SpreadsheetApp.openById(this.DB_ID);
      // Ensure 'Stats_Log' exists in the Master DB!
      ss.getSheetByName("Stats_Log").appendRow([new Date(), svl, ack, law, idp, SpreadsheetApp.getActiveSpreadsheet().getName()]);
    } catch (e) { 
      // This will show up in the 'Executions' tab of Apps Script
      console.error("MasterConnector Error: " + e.toString()); 
    }
  },

  logRoleSession: function(agentName, role, start, end) {
    if (!role) return;
    try {
      const ss = SpreadsheetApp.openById(this.DB_ID);
      const mins = Math.round((end - start) / 60000);
      if (mins < 1) return;
      ss.getSheetByName("Role_Log").appendRow([new Date(), agentName, role, mins, new Date(start), new Date(end)]);
    } catch (e) { console.error("DB Error", e); }
  },

  _checkHistory: function(sheet, name, type) {
    const config = this.THRESHOLDS[type];
    if (!config) return { isFlagged: false, count: 0 };
    const data = sheet.getDataRange().getValues();
    const cutoff = new Date(new Date().getTime() - (config.days * 24 * 60 * 60 * 1000));
    let pastCount = 0;
    for (let i=1; i<data.length; i++) {
      if (String(data[i][1]).toLowerCase() === String(name).toLowerCase() && 
          String(data[i][2]).toUpperCase() === String(type).toUpperCase() && 
          new Date(data[i][0]) >= cutoff) pastCount++;
    }
    const total = pastCount + 1;
    return { isFlagged: total >= config.limit, count: total };
  },

  _createFlag: function(ss, agent, type, count) {
    ss.getSheetByName("Flags").appendRow([new Date(), agent, `${type} Limit Reached`, count, "OPEN"]);
  }
};
