const LogSync = {
  logSession: function(agent, role, startEpoch, endEpoch) {
    const sheet = this._getOrMakeSheet("DB_Sessions");
    const duration = Math.round((endEpoch - startEpoch) / 60000);
    const hours = (duration / 60).toFixed(2);
    if (duration > 0) {
      sheet.appendRow([new Date(), agent, role, new Date(startEpoch), new Date(endEpoch), duration, hours]);
    }
  },
  writeToJournal: function(category, text, user) {
    this._getOrMakeSheet("DB_Journal").appendRow([new Date(), category, user, text]);
  },
  commitShift: function(note) {
    this._getOrMakeSheet("DB_Shift_History").appendRow([new Date(), "SHIFT_COMMIT", note]);
  },
  fillWinds: function(data) { return "Winds Logged"; },
  _getOrMakeSheet: function(name) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      if (name === "DB_Sessions") sheet.appendRow(["Timestamp", "Agent", "Role", "Start", "End", "Mins", "Hours"]);
      sheet.setFrozenRows(1);
    }
    return sheet;
  }
};
