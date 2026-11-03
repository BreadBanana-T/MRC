/**
 * MODULE: STATS TRACKER (LOCAL HOST ONLY)
 */

const StatsTracker = {
  
  SHEET_NAME: "Stats History",
  IDP_SHEET: "IDP_History",
  MAX_HISTORY_POINTS: 24,

  // Columns: Timestamp | SVL | ACK | Trend % | BURG Vol. The last two were
  // added for the break planner (alarms vs the IEX trend over the last 60 min,
  // and the BURG count); older rows simply have them blank, and every other
  // reader of this sheet reads columns 1-3 by position.
  logHourlyStats: function(svl, ack, extra) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(this.SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(this.SHEET_NAME);
      sheet.appendRow(["Timestamp", "SVL", "ACK", "Trend %", "BURG Vol"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    } else if (sheet.getLastColumn() < 5) {
      sheet.getRange(1, 4, 1, 2).setValues([["Trend %", "BURG Vol"]]).setFontWeight("bold");
    }

    let sVal = parseFloat(svl) || 0;
    let aVal = parseFloat(String(ack).replace(/[^\d.]/g, '')) || 0;
    const tr = (extra && extra.trendPct != null && !isNaN(Number(extra.trendPct))) ? Math.round(Number(extra.trendPct) * 10) / 10 : "";
    const bv = (extra && extra.burg != null && !isNaN(Number(extra.burg)) && Number(extra.burg) > 0) ? Number(extra.burg) : "";
    const timestamp = new Date();
    sheet.appendRow([timestamp, sVal, aVal, tr, bv]);
    return "Stats Logged";
  },

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

  getHistory: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(this.SHEET_NAME);

    if (!sheet || sheet.getLastRow() < 2) {
      return JSON.stringify([]);
    }

    const lastRow = sheet.getLastRow();
    const startRow = Math.max(2, lastRow - 48);
    const numRows = lastRow - startRow + 1;
    
    const data = sheet.getRange(startRow, 1, numRows, 3).getValues();

    let history = data.map(row => {
      let svlRaw = parseFloat(row[1]) || 0;
      let ackRaw = String(row[2]).replace(/[^\d.]/g, ''); 
      ackRaw = parseFloat(ackRaw) || 0;

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

    history = history.filter(h => !isNaN(h.time.getTime()));
    history.sort((a, b) => a.time - b.time);
    history = history.slice(-this.MAX_HISTORY_POINTS);
    
    const tz = Session.getScriptTimeZone();
    history.forEach(h => {
        h.label = Utilities.formatDate(h.time, tz, "HH:mm");
    });
    // "t" (epoch ms) lets the dashboard align other time series (e.g. the
    // power-outage band) against these points.
    return JSON.stringify(history.map(h => ({ name: h.label, val: h.val, ack: h.ack, t: h.time.getTime() })));
  }
};
