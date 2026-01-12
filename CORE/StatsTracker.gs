/**
 * MODULE: STATS TRACKER
 * Handles historical data for the Dashboard Graph
 */

const StatsTracker = {
  
  // Configuration
  SHEET_NAME: "Stats History",
  IDP_SHEET: "IDP_History",
  MAX_HISTORY_POINTS: 24, // Keep last 24 entries

  /**
   * Appends a new SVL and ACK record with timestamp.
   */
  logHourlyStats: function(svl, ack) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(this.SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(this.SHEET_NAME);
      sheet.appendRow(["Timestamp", "SVL", "ACK"]);
      sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
    }

    let sVal = parseFloat(svl) || 0;
    let aVal = parseFloat(String(ack).replace(/[^\d.]/g, '')) || 0;
    const timestamp = new Date();
    sheet.appendRow([timestamp, sVal, aVal]);
    return "Stats Logged";
  },

  /**
   * Logs IDP Value specifically.
   */
  logIdp: function(val) {
     const ss = SpreadsheetApp.getActiveSpreadsheet();
     let sheet = ss.getSheetByName(this.IDP_SHEET);
     if (!sheet) {
       sheet = ss.insertSheet(this.IDP_SHEET);
       sheet.appendRow(["Timestamp", "IDP Value"]);
       sheet.getRange(1,1,1,2).setFontWeight("bold");
     }
     sheet.appendRow([new Date(), parseFloat(val)||0]);
     return "IDP Saved";
  },

  getIdpHistory: function() {
     const ss = SpreadsheetApp.getActiveSpreadsheet();
     let sheet = ss.getSheetByName(this.IDP_SHEET);
     if (!sheet || sheet.getLastRow() < 2) return "[]";
     
     // Get today's data or last 24h
     const data = sheet.getDataRange().getValues().slice(1);
     const tz = ss.getSpreadsheetTimeZone();
     
     // Sort by time
     data.sort((a,b) => new Date(a[0]) - new Date(b[0]));
     
     const history = data.slice(-20).map(r => ({
         name: Utilities.formatDate(new Date(r[0]), tz, "HH:mm"),
         val: r[1]
     }));
     return JSON.stringify(history);
  },

  /**
   * Retrieves historical stats, SORTED by time.
   */
  getHistory: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(this.SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      return JSON.stringify([]);
    }

    const lastRow = sheet.getLastRow();
    // Grab more rows to ensure we cover the 24h window, then filter
    const startRow = Math.max(2, lastRow - 48);
    const numRows = lastRow - startRow + 1;
    
    const data = sheet.getRange(startRow, 1, numRows, 3).getValues();
    // 1. Map to Objects
    let history = data.map(row => ({
      time: new Date(row[0]), // Keep full Date object for sorting
      label: "",
      val: parseFloat(row[1]) || 0,
      ack: parseFloat(row[2]) || 0
    }));
    // 2. SORT Chronologically (Fixes the "jumping time" bug)
    history.sort((a, b) => a.time - b.time);
    // 3. Slice to last 24 points and Format Label
    history = history.slice(-this.MAX_HISTORY_POINTS);
    // 4. Create readable labels (HH:mm)
    const tz = ss.getSpreadsheetTimeZone();
    history.forEach(h => {
        h.label = Utilities.formatDate(h.time, tz, "HH:mm");
    });
    return JSON.stringify(history.map(h => ({ name: h.label, val: h.val, ack: h.ack })));
  }
};
