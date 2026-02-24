/**
 * MODULE: STATUS TRACKER
 * Dual-Write Enabled: Bridges local overrides natively into the Master Database.
 */

const StatusTracker = {
  
  // --- 1. ROLES & ABSENCE (Status_Overrides) ---
  updateStatus: function(name, type, value) {
    const dateStr = this._getTodayStr(); 
    const cleanName = String(name).trim().toLowerCase();
    const now = new Date();

    const writeToSheet = (spreadsheet) => {
        const sheet = this._getSheet("Status_Overrides", ["Timestamp", "Agent Name", "Role", "Absent Type", "DateStr"], spreadsheet);
        const data = sheet.getDataRange().getValues();
        let foundRow = -1;
        let oldData = null;

        for (let i = 1; i < data.length; i++) {
          if (this._matchRow(data[i], cleanName, dateStr, 1, 4)) {
            foundRow = i + 1; oldData = data[i]; break;
          }
        }

        // Bridge Logging (Only trigger once per Local call, skipped in Master loop)
        if (spreadsheet === SpreadsheetApp.getActiveSpreadsheet()) {
            if (type === 'absent' && value !== "") {
                if (typeof MasterConnector !== 'undefined') MasterConnector.logAbsence(name, value);
            }
            if (foundRow > -1 && oldData) {
                const oldRole = oldData[2]; const oldTime = oldData[0];
                if (type === 'role' && oldRole && oldRole !== "" && oldRole !== value) {
                     const startEpoch = new Date(oldTime).getTime(); const endEpoch = now.getTime();
                     if (typeof MasterConnector !== 'undefined') MasterConnector.logRoleSession(name, oldRole, startEpoch, endEpoch);
                     if (typeof LogSync !== 'undefined') LogSync.logSession(name, oldRole, startEpoch, endEpoch);
                }
            }
        }

        // Update Record
        if (foundRow > -1) {
          sheet.getRange(foundRow, 1).setValue(now);
          if (type === 'role') sheet.getRange(foundRow, 3).setValue(value);
          if (type === 'absent') sheet.getRange(foundRow, 4).setValue(value);
        } else {
          let roleVal = type === 'role' ? value : "";
          let absentVal = type === 'absent' ? value : "";
          sheet.appendRow([now, name, roleVal, absentVal, dateStr]);
        }
    };

    writeToSheet(SpreadsheetApp.getActiveSpreadsheet());
    if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
        try { writeToSheet(SpreadsheetApp.openById(MasterConnector.DB_ID)); } catch(e) {}
    }

    return "Status Saved";
  },

  // --- 2. BREAK MODIFICATIONS (Break_Overrides) ---
  updateBreaks: function(name, jsonBreaks) {
    const dateStr = this._getTodayStr();
    const cleanName = String(name).trim().toLowerCase();
    
    const writeToSheet = (spreadsheet) => {
        const sheet = this._getSheet("Break_Overrides", ["Timestamp", "Agent Name", "DateStr", "Breaks JSON"], spreadsheet);
        const data = sheet.getDataRange().getValues();
        let foundRow = -1;
        
        for (let i = 1; i < data.length; i++) {
          if (this._matchRow(data[i], cleanName, dateStr, 1, 2)) { foundRow = i + 1; break; }
        }

        if (foundRow > -1) {
          sheet.getRange(foundRow, 1).setValue(new Date());
          sheet.getRange(foundRow, 4).setValue(jsonBreaks);
        } else {
          sheet.appendRow([new Date(), name, dateStr, jsonBreaks]);
        }
    };

    writeToSheet(SpreadsheetApp.getActiveSpreadsheet());
    if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
        try { writeToSheet(SpreadsheetApp.openById(MasterConnector.DB_ID)); } catch(e) {}
    }
    
    return "Breaks Saved";
  },

  // --- 3. OVERTIME (Overtime_Tracking) ---
  logOvertime: function(name, start, end, bStart, bEnd) {
    const dateStr = this._getTodayStr();
    
    const writeToSheet = (spreadsheet) => {
        const sheet = this._getSheet("Overtime_Tracking", ["Timestamp", "Agent Name", "OT Start", "OT End", "Break Start", "Break End", "DateStr"], spreadsheet);
        sheet.appendRow([new Date(), name, start, end, bStart || "-", bEnd || "-", dateStr]);
    };

    writeToSheet(SpreadsheetApp.getActiveSpreadsheet());
    if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
        try { writeToSheet(SpreadsheetApp.openById(MasterConnector.DB_ID)); } catch(e) {}
        if (typeof MasterConnector.logOvertime === 'function') MasterConnector.logOvertime(name, start, end, bStart, bEnd);
    }
    
    return "OT Logged";
  },

  // --- DATA RETRIEVAL (INTELLIGENT FAST READ) ---
  getConsolidatedData: function() {
    const todayStr = this._getTodayStr();
    const yesterdayStr = this._getYesterdayStr();
    const map = new Map();

    const getEntry = (name) => {
      const k = String(name).trim().toLowerCase();
      if (!map.has(k)) map.set(k, { role: "", absent: "", breaks: null, ot: [] });
      return map.get(k);
    };

    // ALWAYS Read from FAST Local RAM
    const readSheet = (sheetName, nameIdx, dateIdx, callback) => {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        const data = sheet.getDataRange().getValues().slice(1);
        data.forEach(row => {
           if (this._matchDate(row[dateIdx], todayStr) || this._matchDate(row[dateIdx], yesterdayStr)) { callback(row); }
        });
      }
    };

    readSheet("Status_Overrides", 1, 4, (row) => {
       const e = getEntry(row[1]); if(row[2]) e.role = row[2]; if(row[3]) e.absent = row[3];
    });
    readSheet("Break_Overrides", 1, 2, (row) => {
       const e = getEntry(row[1]); e.breaks = row[3];
    });
    readSheet("Overtime_Tracking", 1, 6, (row) => {
       const e = getEntry(row[1]); e.ot.push({ start: row[2], end: row[3], bStart: row[4], bEnd: row[5] });
    });
    return map;
  },

  // --- HELPERS ---
  _getSheet: function(name, headers, targetSS) {
    if (!targetSS) targetSS = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = targetSS.getSheetByName(name);
    if (!sheet) {
      sheet = targetSS.insertSheet(name);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
    return sheet;
  },
  _getTodayStr: function() { return Utilities.formatDate(new Date(), "America/Toronto", "yyyy-MM-dd"); },
  _getYesterdayStr: function() {
    const d = new Date(); d.setDate(d.getDate() - 1);
    return Utilities.formatDate(d, "America/Toronto", "yyyy-MM-dd");
  },
  _matchDate: function(val, targetStr) {
    if (!val) return false; let d = val;
    if (d instanceof Date) d = Utilities.formatDate(d, "America/Toronto", "yyyy-MM-dd");
    return d === targetStr;
  },
  _matchRow: function(row, cleanName, dateStr, nameIdx, dateIdx) {
    const rName = String(row[nameIdx]).trim().toLowerCase();
    return rName === cleanName && this._matchDate(row[dateIdx], dateStr);
  }
};
