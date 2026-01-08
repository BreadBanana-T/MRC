/**
 * MODULE: STATS TRACKER
 * Handles historical data for the Dashboard Graph
 */

const StatsTracker = {
  
  // Configuration
  SHEET_NAME: "Stats History",
  MAX_HISTORY_POINTS: 24, // Keep last 24 entries (e.g., 24 hours or 12 hours of 30m intervals)

  /**
   * Appends a new SVL and ACK record with timestamp.
   * @param {string|number} svl - Service Level (e.g. 85 or "85%")
   * @param {string|number} ack - ACK Time (e.g. 25 or "25s")
   */
  logHourlyStats: function(svl, ack) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(this.SHEET_NAME);
    
    // Create sheet if missing
    if (!sheet) {
      sheet = ss.insertSheet(this.SHEET_NAME);
      sheet.appendRow(["Timestamp", "SVL", "ACK"]); // Added ACK column
      sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
    }

    // Parse values
    let sVal = parseFloat(svl);
    let aVal = parseFloat(String(ack).replace(/[^\d.]/g, '')); // Strip 's' if present

    if (isNaN(sVal)) sVal = 0;
    if (isNaN(aVal)) aVal = 0;

    const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    sheet.appendRow([timestamp, sVal, aVal]);
    return "Stats Logged";
  },

  /**
   * Retrieves historical stats for the dashboard graph.
   * Returns: JSON string of array [{name: "10:00", val: 85, ack: 20}, ...]
   */
  getHistory: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(this.SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return JSON.stringify([]); 
    }

    const lastRow = sheet.getLastRow();
    // Get last N rows
    const startRow = Math.max(2, lastRow - this.MAX_HISTORY_POINTS + 1);
    const numRows = lastRow - startRow + 1;
    
    // Read 3 columns: Timestamp, SVL, ACK
    const lastCol = sheet.getLastColumn();
    const data = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
    
    const history = data.map(row => {
      // Parse timestamp to simple time string (e.g., "14:00")
      let timeLabel = "";
      if (row[0] instanceof Date) {
        timeLabel = Utilities.formatDate(row[0], ss.getSpreadsheetTimeZone(), "HH:mm");
      } else {
        timeLabel = String(row[0]).split(" ")[1] ? String(row[0]).split(" ")[1].substring(0, 5) : row[0];
      }
      
      return {
        name: timeLabel,
        val: parseFloat(row[1]) || 0,
        ack: parseFloat(row[2]) || 0
      };
    });

    return JSON.stringify(history);
  }
};
