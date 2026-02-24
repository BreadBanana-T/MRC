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
     
     const data = sheet.getDataRange().getValues().slice(1);
     const tz = ss.getSpreadsheetTimeZone();
     
     data.sort((a,b) => new Date(a[0]) - new Date(b[0]));
     const history = data.slice(-20).map(r => ({
         name: Utilities.formatDate(new Date(r[0]), tz, "HH:mm"),
         val: r[1]
     }));
     return JSON.stringify(history);
  },

  /**
   * Retrieves historical stats, SORTED by time, directly from MASTER DB.
   */
  getHistory: function() {
    let sheet;
    
    // 1. Tries to Read from the Master DB's Stats_Log FIRST
    if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
        try {
            const ssMaster = SpreadsheetApp.openById(MasterConnector.DB_ID);
            sheet = ssMaster.getSheetByName("Stats_Log");
        } catch(e) {}
    }
    
    // 2. Fallback to Local Sheet if Master DB fails
    if (!sheet) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        sheet = ss.getSheetByName(this.SHEET_NAME);
    }

    if (!sheet || sheet.getLastRow() < 2) {
      return JSON.stringify([]);
    }

    const lastRow = sheet.getLastRow();
    const startRow = Math.max(2, lastRow - 48);
    const numRows = lastRow - startRow + 1;
    
    // Columns might be different in Master DB (Timestamp, SVL, ACK, LAW, IDP)
    // We only need the first 3 columns.
    const data = sheet.getRange(startRow, 1, numRows, 3).getValues();

    let history = data.map(row => {
      let svlRaw = parseFloat(row[1]) || 0;
      let ackRaw = String(row[2]).replace(/[^\d.]/g, ''); // Strips out "s" safely
      ackRaw = parseFloat(ackRaw) || 0;

      // Clean decimals (e.g., 0.87 -> 87)
      if (svlRaw > 0 && svlRaw <= 1) {
          svlRaw = Math.round(svlRaw * 100);
      }

      return {
        time: new Date(row[0]),
        label: "",
        val: svlRaw,
        ack: ackRaw
      };
    });

    // Clean out any corrupt dates
    history = history.filter(h => !isNaN(h.time.getTime()));
    
    // Sort Chronologically
    history.sort((a, b) => a.time - b.time);
    
    // Slice to the last 24 points
    history = history.slice(-this.MAX_HISTORY_POINTS);
    
    const tz = Session.getScriptTimeZone();
    history.forEach(h => {
        h.label = Utilities.formatDate(h.time, tz, "HH:mm");
    });

    return JSON.stringify(history.map(h => ({ name: h.label, val: h.val, ack: h.ack })));
  }
};
