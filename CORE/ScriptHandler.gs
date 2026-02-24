/**
 * MODULE: SCRIPT HANDLER
 * Dual-Write enabled: Reads locally for speed, saves/deletes to BOTH Local and Master DB.
 */

const ScriptHandler = {
  SHEET_NAME: "Team_Scripts",

  _getSheet: function(ss) {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(this.SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(this.SHEET_NAME);
      sheet.appendRow(["Title", "Script Body", "Category"]);
      sheet.getRange(1, 1, 1, 3).setFontWeight("bold");
    } 
    
    const maxCols = sheet.getMaxColumns();
    if (maxCols < 3) sheet.insertColumnsAfter(maxCols, 3 - maxCols);
    if (sheet.getRange(1, 3).getValue() !== "Category") sheet.getRange(1, 3).setValue("Category").setFontWeight("bold");
    
    return sheet;
  },

  getAll: function() {
    // Reading uses the FAST Local RAM Sheet
    const sheet = this._getSheet(SpreadsheetApp.getActiveSpreadsheet());
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const numCols = sheet.getLastColumn();
    const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    
    return data.map((row, index) => ({
      index: index,
      title: row[0] || "",
      body: row[1] || "",
      category: row[2] || "General"
    }));
  },

  save: function(index, title, body, category) {
    if(!title || !body) return "Missing Info";
    const catVal = category || "General";
    
    // Dual Write Logic
    const writeToSheet = (spreadsheet) => {
        const sheet = this._getSheet(spreadsheet);
        if (index !== null && index !== undefined && index !== "") {
           const row = parseInt(index) + 2;
           if (row <= sheet.getLastRow()) {
              sheet.getRange(row, 1, 1, 3).setValues([[title, body, catVal]]);
              return;
           }
        }
        sheet.appendRow([title, body, catVal]);
    };

    // 1. Write Local
    writeToSheet(SpreadsheetApp.getActiveSpreadsheet());
    
    // 2. Write Master DB
    if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
        try { writeToSheet(SpreadsheetApp.openById(MasterConnector.DB_ID)); } catch(e) {}
    }

    return "Saved";
  },

  delete: function(index) {
    const rowToDelete = parseInt(index) + 2;

    const deleteFromSheet = (spreadsheet) => {
        const sheet = this._getSheet(spreadsheet);
        if (rowToDelete <= sheet.getLastRow()) sheet.deleteRow(rowToDelete);
    };

    // 1. Local Delete
    deleteFromSheet(SpreadsheetApp.getActiveSpreadsheet());
    
    // 2. Master DB Delete
    if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
        try { deleteFromSheet(SpreadsheetApp.openById(MasterConnector.DB_ID)); } catch(e) {}
    }

    return "Deleted";
  }
};

// --- GLOBAL EXPORTS ---
function getTeamScripts() { return JSON.stringify(ScriptHandler.getAll()); }
function saveTeamScript(i, t, b, c) { return ScriptHandler.save(i, t, b, c); }
function deleteTeamScript(i) { return ScriptHandler.delete(i); }
