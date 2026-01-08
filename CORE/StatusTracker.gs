/**
 * MODULE: STATUS TRACKER
 * The Persistent Brain. 
 * Manages overrides (Roles, Breaks, OT) so they survive refreshes and re-imports.
 * Includes "Overnight Awareness" logic.
 */

const StatusTracker = {
  
  // --- 1. ROLES & ABSENCE (Status_Overrides) ---
  updateStatus: function(name, type, value) {
    const sheet = this._getSheet("Status_Overrides", ["Timestamp", "Agent Name", "Role", "Absent Type", "DateStr"]);
    
    // We store using "Logical Date" - if it's 1 AM, we likely mean the shift that started yesterday.
    // However, to keep it simple and robust, we default to TODAY. 
    // AgentMonitor will handle the "reading" intelligence.
    const dateStr = this._getTodayStr(); 
    const cleanName = String(name).trim().toLowerCase();
    
    const data = sheet.getDataRange().getValues();
    let foundRow = -1;

    for (let i = 1; i < data.length; i++) {
      if (this._matchRow(data[i], cleanName, dateStr, 1, 4)) {
        foundRow = i + 1; break;
      }
    }

    if (foundRow > -1) {
      sheet.getRange(foundRow, 1).setValue(new Date());
      if (type === 'role') sheet.getRange(foundRow, 3).setValue(value);
      if (type === 'absent') sheet.getRange(foundRow, 4).setValue(value);
    } else {
      let roleVal = type === 'role' ? value : "";
      let absentVal = type === 'absent' ? value : "";
      sheet.appendRow([new Date(), name, roleVal, absentVal, dateStr]);
    }
    return "Status Saved";
  },

  // --- 2. BREAK MODIFICATIONS (Break_Overrides) ---
  updateBreaks: function(name, jsonBreaks) {
    const sheet = this._getSheet("Break_Overrides", ["Timestamp", "Agent Name", "DateStr", "Breaks JSON"]);
    const dateStr = this._getTodayStr();
    const cleanName = String(name).trim().toLowerCase();
    
    const data = sheet.getDataRange().getValues();
    let foundRow = -1;

    for (let i = 1; i < data.length; i++) {
      if (this._matchRow(data[i], cleanName, dateStr, 1, 2)) {
        foundRow = i + 1; break;
      }
    }

    if (foundRow > -1) {
      sheet.getRange(foundRow, 1).setValue(new Date());
      sheet.getRange(foundRow, 4).setValue(jsonBreaks);
    } else {
      sheet.appendRow([new Date(), name, dateStr, jsonBreaks]);
    }
    return "Breaks Saved";
  },

  // --- 3. OVERTIME (Overtime_Tracking) ---
  logOvertime: function(name, start, end, bStart, bEnd) {
    const sheet = this._getSheet("Overtime_Tracking", ["Timestamp", "Agent Name", "OT Start", "OT End", "Break Start", "Break End", "DateStr"]);
    const dateStr = this._getTodayStr();
    sheet.appendRow([new Date(), name, start, end, bStart || "-", bEnd || "-", dateStr]);
    return "OT Logged";
  },

  // --- DATA RETRIEVAL (INTELLIGENT) ---
  
  /**
   * Returns ALL overrides relevant to the current operational window.
   * Checks TODAY and YESTERDAY to handle overnight shifts.
   */
  getConsolidatedData: function() {
    const todayStr = this._getTodayStr();
    const yesterdayStr = this._getYesterdayStr();
    
    const map = new Map();

    const getEntry = (name) => {
      const k = String(name).trim().toLowerCase();
      if (!map.has(k)) map.set(k, { role: "", absent: "", breaks: null, ot: [] });
      return map.get(k);
    };

    // Helper to read a sheet and populate map
    const readSheet = (sheetName, nameIdx, dateIdx, callback) => {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        const data = sheet.getDataRange().getValues().slice(1);
        data.forEach(row => {
           // We accept row if date is Today OR Yesterday (Overnight safety)
           if (this._matchDate(row[dateIdx], todayStr) || this._matchDate(row[dateIdx], yesterdayStr)) {
              callback(row);
           }
        });
      }
    };

    // 1. Status
    readSheet("Status_Overrides", 1, 4, (row) => {
       const e = getEntry(row[1]);
       // If multiple entries (yesterday/today), we prefer the freshest one usually, 
       // but here we just overwrite, meaning today's entry wins if both exist.
       if(row[2]) e.role = row[2];
       if(row[3]) e.absent = row[3];
    });

    // 2. Breaks
    readSheet("Break_Overrides", 1, 2, (row) => {
       const e = getEntry(row[1]);
       e.breaks = row[3];
    });

    // 3. Overtime
    readSheet("Overtime_Tracking", 1, 6, (row) => {
       const e = getEntry(row[1]);
       e.ot.push({ start: row[2], end: row[3], bStart: row[4], bEnd: row[5] });
    });

    return map;
  },

  // --- HELPERS ---
  _getSheet: function(name, headers) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
    return sheet;
  },
  
  _getTodayStr: function() {
    return Utilities.formatDate(new Date(), "America/Toronto", "yyyy-MM-dd");
  },
  
  _getYesterdayStr: function() {
    const d = new Date();
    d.setDate(d.getDate() - 1);
    return Utilities.formatDate(d, "America/Toronto", "yyyy-MM-dd");
  },

  _matchDate: function(val, targetStr) {
    if (!val) return false;
    let d = val;
    if (d instanceof Date) d = Utilities.formatDate(d, "America/Toronto", "yyyy-MM-dd");
    return d === targetStr;
  },

  _matchRow: function(row, cleanName, dateStr, nameIdx, dateIdx) {
    const rName = String(row[nameIdx]).trim().toLowerCase();
    return rName === cleanName && this._matchDate(row[dateIdx], dateStr);
  }
};
