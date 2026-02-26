/**
 * NOTIFICATION HANDLER
 * Manages persistent notifications for agent coding and system syncs.
 */

const NotificationHandler = {
  SHEET_NAME: "System_Notifications",

  _getSheet: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(this.SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(this.SHEET_NAME);
      sheet.hideSheet();
      sheet.appendRow(["ID", "Timestamp", "Message", "Status", "AgentName"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    }
    return sheet;
  },

  add: function(agentName, action, isGlobal = false) {
    const sheet = this._getSheet();
    const id = new Date().getTime().toString();
    const timestamp = Utilities.formatDate(new Date(), "America/Toronto", "MM/dd HH:mm");
    
    // Allows sending standard messages vs global network alerts
    const message = isGlobal ? action : `Code ${agentName} as ${action} in IEX`;
    
    sheet.appendRow([id, timestamp, message, "PENDING", agentName]);
    return { id, timestamp, message };
  },

  getPending: function() {
    const sheet = this._getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify([]);

    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();

    const pending = data
      .filter(row => row[3] === "PENDING")
      .map(row => {
        let timeStr = row[1];
        if (row[1] instanceof Date) timeStr = Utilities.formatDate(row[1], tz, "MM/dd HH:mm");
        return { id: row[0], time: timeStr, text: row[2], agent: row[4] };
      })
      .reverse();

    return JSON.stringify(pending);
  },

  dismiss: function(id) {
    const sheet = this._getSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 4).setValue("DONE");
        return "Dismissed";
      }
    }
    return "Not Found";
  }
};

// --- GLOBAL EXPORTS ---
function addSystemNotification(agentName, action) { return NotificationHandler.add(agentName, action); }
function addGlobalNotification(msg) { return NotificationHandler.add("SYSTEM", msg, true); }
function getSystemNotifications() { return NotificationHandler.getPending(); }
function dismissSystemNotification(id) { return NotificationHandler.dismiss(id); }
