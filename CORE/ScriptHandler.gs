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
      // Create new if missing
      sheet = ss.insertSheet(this.SHEET_NAME);
      sheet.appendRow(["Title", "Script Body", "Category"]);
      sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
      sheet.appendRow(["Long Call Check", "Hi, I noticed you've been on this call for 15+ mins. Do you need supervisor assistance?", "Quality"]);
    } 
    
    // --- SELF-REPAIR: Ensure Column C (Category) exists ---
    const maxCols = sheet.getMaxColumns();
    if (maxCols < 3) {
       sheet.insertColumnsAfter(maxCols, 3 - maxCols);
    }
    
    // Ensure Header is correct
    const header = sheet.getRange(1, 3).getValue();
    if (header !== "Category") {
       sheet.getRange(1, 3).setValue("Category").setFontWeight("bold");
    }
    
    return sheet;
  },

  getAll: function() {
    const sheet = this._getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // Safe Read: Grab all available columns, handle missing data in JS
    const numCols = sheet.getLastColumn();
    const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    
    return data.map((row, index) => ({
      index: index,
      title: row[0] || "",
      body: row[1] || "",
      category: row[2] || "General" // Default if column is empty/missing
    }));
  },

  save: function(index, title, body, category) {
    if(!title || !body) return "Missing Info";
    const sheet = this._getSheet();
    const catVal = category || "General";

    // UPDATE EXISTING
    if (index !== null && index !== undefined && index !== "") {
       const row = parseInt(index) + 2; 
       if (row <= sheet.getLastRow()) {
          // Force write to first 3 columns
          sheet.getRange(row, 1, 1, 3).setValues([[title, body, catVal]]);
          return "Updated";
       }
    }

    // CREATE NEW
    sheet.appendRow([title, body, catVal]);
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
function saveTeamScript(i, t, b, c) { return ScriptHandler.save(i, t, b, c); }
function deleteTeamScript(i) { return ScriptHandler.delete(i); }
