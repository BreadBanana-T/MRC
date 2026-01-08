/**
 * NOTIFICATION HANDLER
 * Manages persistent notifications for agent coding updates.
 * Stores data in a hidden sheet named "System_Notifications".
 */

const NotificationHandler = {
  
  SHEET_NAME: "System_Notifications",

  /**
   * Gets or creates the notification sheet.
   */
  _getSheet: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(this.SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(this.SHEET_NAME);
      sheet.hideSheet();
      // Keep it hidden to avoid clutter
      sheet.appendRow(["ID", "Timestamp", "Message", "Status", "AgentName"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    }
    return sheet;
  },

  /**
   * Adds a new notification.
   */
  add: function(agentName, action) {
    const sheet = this._getSheet();
    const id = new Date().getTime().toString();
    const timestamp = Utilities.formatDate(new Date(), "America/Toronto", "MM/dd HH:mm");
    const message = `Code ${agentName} as ${action} in IEX`;
    sheet.appendRow([id, timestamp, message, "PENDING", agentName]);
    return { id, timestamp, message };
  },

  /**
   * Gets all PENDING notifications.
   */
  getPending: function() {
    const sheet = this._getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return JSON.stringify([]);

    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();

    // Filter for "PENDING" status
    const pending = data
      .filter(row => row[3] === "PENDING")
      .map(row => {
        // FIX: Check if time is a Date object and format it
        let timeStr = row[1];
        if (row[1] instanceof Date) {
           timeStr = Utilities.formatDate(row[1], tz, "MM/dd HH:mm");
        }

        return {
          id: row[0],
          time: timeStr, // Now correctly formatted
          text: row[2],
          agent: row[4]
        };
      })
      .reverse();
    // Show newest first

    return JSON.stringify(pending);
  },

  /**
   * Marks a notification as DONE (Coded).
   */
  dismiss: function(id) {
    const sheet = this._getSheet();
    const data = sheet.getDataRange().getValues();
    // Find row by ID (Column A)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        // Update Status column (Col 4 / D)
        sheet.getRange(i + 1, 4).setValue("DONE");
        return "Dismissed";
      }
    }
    return "Not Found";
  }
};

// --- GLOBAL EXPORTS FOR CLIENT SIDE ---

function addSystemNotification(agentName, action) {
  return NotificationHandler.add(agentName, action);
}

function getSystemNotifications() {
  return NotificationHandler.getPending();
}

function dismissSystemNotification(id) {
  return NotificationHandler.dismiss(id);
}
