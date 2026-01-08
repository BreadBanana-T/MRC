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
      index: index, // Used for deletion/editing
      title: row[0],
      body: row[1],
      category: row[2] || "General" 
    }));
  },

  save: function(index, title, body, category) {
    if(!title || !body) return "Missing Info";
    const sheet = this._getSheet();
    
    // UPDATE EXISTING (If index is valid)
    if (index !== null && index !== undefined && index !== "") {
       const row = parseInt(index) + 2; // Data starts at row 2
       if (row <= sheet.getLastRow()) {
          sheet.getRange(row, 1, 1, 3).setValues([[title, body, category || "General"]]);
          return "Updated";
       }
    }

    // CREATE NEW
    sheet.appendRow([title, body, category || "General"]);
    return "Saved";
  },

  delete: function(index) {
    const sheet = this._getSheet();
    sheet.deleteRow(parseInt(index) + 2);
    return "Deleted";
  }
};

// --- GLOBAL EXPORTS ---
function getTeamScripts() { return JSON.stringify(ScriptHandler.getAll()); }
// Updated to accept index for editing
function saveTeamScript(i, t, b, c) { return ScriptHandler.save(i, t, b, c); }
function deleteTeamScript(i) { return ScriptHandler.delete(i); }
