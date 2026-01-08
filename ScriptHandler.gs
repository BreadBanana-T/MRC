/**
 * MODULE: SCRIPT HANDLER
 * Manages shared clipboard scripts for the team.
 */

const ScriptHandler = {
  SHEET_NAME: "Team_Scripts",

  _getSheet: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(this.SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(this.SHEET_NAME);
      // Added "Category" to header
      sheet.appendRow(["Title", "Script Body", "Category"]);
      sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
      // Add some defaults with Categories
      sheet.appendRow(["Long Call Check", "Hi, I noticed you've been on this call for 15+ mins. Do you need supervisor assistance?", "Quality"]);
      sheet.appendRow(["Late from Break", "Hi, checking in—you're showing as away past your break time. Everything good?", "Supervision"]);
      sheet.appendRow(["ACW Nudge", "Hi, quick check-in on your ACW status. Need any help wrapping up?", "Supervision"]);
    }
    return sheet;
  },

  getAll: function() {
    const sheet = this._getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // Get all rows (Cols 1-3: Title, Body, Category)
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    return data.map((row, index) => ({
      index: index, // Used for deletion
      title: row[0],
      body: row[1],
      category: row[2] || "General" // Default if empty
    }));
  },

  add: function(title, body, category) {
    if(!title || !body) return "Missing Info";
    const sheet = this._getSheet();
    // Save with Category
    sheet.appendRow([title, body, category || "General"]);
    return "Saved";
  },

  delete: function(index) {
    const sheet = this._getSheet();
    // Index is 0-based from data array, but sheet is 1-based and has header (row 1)
    // So Row = index + 2
    sheet.deleteRow(parseInt(index) + 2);
    return "Deleted";
  }
};

// --- GLOBAL EXPORTS ---
function getTeamScripts() { return JSON.stringify(ScriptHandler.getAll()); }
function addTeamScript(t, b, c) { return ScriptHandler.add(t, b, c); }
function deleteTeamScript(i) { return ScriptHandler.delete(i); }
