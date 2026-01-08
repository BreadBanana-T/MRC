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
      sheet.appendRow(["Title", "Script Body"]);
      sheet.getRange(1, 1, 1, 2).setFontWeight("bold");
      // Add some defaults
      sheet.appendRow(["Long Call Check", "Hi, I noticed you've been on this call for 15+ mins. Do you need supervisor assistance?"]);
      sheet.appendRow(["Late from Break", "Hi, checking in—you're showing as away past your break time. Everything good?"]);
      sheet.appendRow(["ACW Nudge", "Hi, quick check-in on your ACW status. Need any help wrapping up?"]);
    }
    return sheet;
  },

  getAll: function() {
    const sheet = this._getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // Get all rows
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    return data.map((row, index) => ({
      index: index, // Used for deletion
      title: row[0],
      body: row[1]
    }));
  },

  add: function(title, body) {
    if(!title || !body) return "Missing Info";
    const sheet = this._getSheet();
    sheet.appendRow([title, body]);
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
function addTeamScript(t, b) { return ScriptHandler.add(t, b); }
function deleteTeamScript(i) { return ScriptHandler.delete(i); }
