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
      // CREATE NEW SHEET
      sheet = ss.insertSheet(this.SHEET_NAME);
      sheet.appendRow(["Title", "Script Body", "Category"]);
      sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
      // Add defaults
      sheet.appendRow(["Long Call Check", "Hi, I noticed you've been on this call for 15+ mins. Do you need supervisor assistance?", "Quality"]);
      sheet.appendRow(["Late from Break", "Hi, checking in—you're showing as away past your break time. Everything good?", "Supervision"]);
    } else {
      // AUTO-UPGRADE LEGACY SHEETS
      // If the sheet exists but only has 2 columns, add the "Category" header
      if (sheet.getLastColumn() < 3) {
         sheet.getRange(1, 3).setValue("Category").setFontWeight("bold");
         // Fill existing rows with "General" to prevent null errors
         const lastRow = sheet.getLastRow();
         if (lastRow > 1) {
             sheet.getRange(2, 3, lastRow - 1, 1).setValue("General");
         }
      }
    }
    return sheet;
  },

  getAll: function() {
    const sheet = this._getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // Force grab 3 columns even if empty
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    
    return data.map((row, index) => ({
      index: index,
      title: row[0],
      body: row[1],
      category: row[2] || "General" 
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
