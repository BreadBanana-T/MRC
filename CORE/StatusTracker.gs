/**
 * MODULE: STATUS TRACKER
 * Manages overrides (Roles, Breaks, OT) so they survive refreshes.
 *
 * Note: the legacy MasterConnector bridge was removed — the spreadsheet
 * is now the single source of truth. logHourlyStats/getHistory moved
 * exclusively to StatsTracker.gs.
 */

const StatusTracker = {

  // Same key contract as RoleManager and AgentMonitor — see the note in
  // RoleManager._key. Sheet rows keep the human-readable display name in the
  // name column; only the in-memory lookup key is normalized, so legacy rows
  // are matched correctly with no migration pass over the sheets.
  _key: function(name) {
    return (typeof _normalizeAgentKey === 'function')
      ? _normalizeAgentKey(name)
      : String(name).trim().toLowerCase();
  },

  // --- 1. ROLES & ABSENCE (Status_Overrides) ---
  updateStatus: function(name, type, value) {
    const sheet = this._getSheet("Status_Overrides", ["Timestamp", "Agent Name", "Role", "Absent Type", "DateStr"]);
    const dateStr = this._getTodayStr();
    const cleanName = this._key(name);

    const data = sheet.getDataRange().getValues();
    let foundRow = -1;
    let oldData = null;

    for (let i = 1; i < data.length; i++) {
      if (this._matchRow(data[i], cleanName, dateStr, 1, 4)) {
        foundRow = i + 1;
        oldData = data[i];
        break;
      }
    }

    const now = new Date();

    // Role session logging (local log used by Insights graphs).
    // SAFE is NEVER logged as session hours (per product owner): a manual
    // "Assign SAFE" is floor status only. Additive combos like "ICL SAFE"
    // keep their non-SAFE part so ICL/ULC hours still log.
    if (foundRow > -1 && oldData && type === 'role') {
      const oldRole = oldData[2];
      const oldTime = oldData[0];
      if (oldRole && oldRole !== '' && oldRole !== value) {
        const logRole = String(oldRole).split(/\s+/)
          .filter(t => t && t.toUpperCase().indexOf('SAFE') === -1)
          .join(' ');
        if (logRole) {
          const startEpoch = new Date(oldTime).getTime();
          const endEpoch = now.getTime();
          if (typeof LogSync !== 'undefined') {
            LogSync.logSession(name, logRole, startEpoch, endEpoch);
          }
        }
      }
    }

    if (foundRow > -1) {
      // One setValues instead of up to three single-cell setValue round-trips.
      const cur = data[foundRow - 1];
      const roleVal = (type === 'role') ? value : (cur[2] || "");
      const absentVal = (type === 'absent') ? value : (cur[3] || "");
      sheet.getRange(foundRow, 1, 1, 5).setValues([[now, name, roleVal, absentVal, dateStr]]);
    } else {
      let roleVal = type === 'role' ? value : "";
      let absentVal = type === 'absent' ? value : "";
      sheet.appendRow([now, name, roleVal, absentVal, dateStr]);
    }
    this._prune(sheet, 4);
    return "Status Saved";
  },

  // --- 2. BREAK MODIFICATIONS (Break_Overrides) ---
  updateBreaks: function(name, jsonBreaks) {
    const sheet = this._getSheet("Break_Overrides", ["Timestamp", "Agent Name", "DateStr", "Breaks JSON"]);
    const dateStr = this._getTodayStr();
    const cleanName = this._key(name);

    const data = sheet.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (this._matchRow(data[i], cleanName, dateStr, 1, 2)) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow > -1) {
      sheet.getRange(foundRow, 1).setValue(new Date());
      sheet.getRange(foundRow, 4).setValue(jsonBreaks);
    } else {
      sheet.appendRow([new Date(), name, dateStr, jsonBreaks]);
    }
    this._prune(sheet, 2);
    return "Breaks Saved";
  },

  // --- 3. OVERTIME (Overtime_Tracking) ---
  logOvertime: function(name, start, end, bStart, bEnd) {
    const sheet = this._getSheet("Overtime_Tracking", ["Timestamp", "Agent Name", "OT Start", "OT End", "Break Start", "Break End", "DateStr"]);
    const dateStr = this._getTodayStr();
    sheet.appendRow([new Date(), name, start, end, bStart || "-", bEnd || "-", dateStr]);
    this._prune(sheet, 6);
    return "OT Logged";
  },

  // --- 4. GENESYS ASSIGNMENT (Genesys_Assignments) ---
  // One row per agent (latest wins), NO date stamp — assignments persist
  // until the agent is explicitly moved back to Avaya (channel = "").
  setGenesys: function(name, channel, queuesJson) {
    return this.setGenesysBatch([name], channel, queuesJson);
  },

  // Batch write: one sheet read to index existing rows, then per-name
  // in-place updates and a single append for the new ones.
  setGenesysBatch: function(names, channel, queuesJson) {
    const sheet = this._getSheet("Genesys_Assignments", ["Timestamp", "Agent Name", "Channel", "Queues JSON"]);
    const q = channel === 'inbound' ? (queuesJson || "[]") : "[]";
    const now = new Date();

    const data = sheet.getDataRange().getValues();
    const rowByName = {};
    for (let i = 1; i < data.length; i++) {
      rowByName[this._key(data[i][1])] = i + 1;
    }

    const self = this;
    const toAppend = [];
    names.forEach(function(name) {
      const cleanName = self._key(name);
      if (!cleanName) return;
      const row = rowByName[cleanName];
      if (row) sheet.getRange(row, 1, 1, 4).setValues([[now, name, channel || "", q]]);
      else toAppend.push([now, name, channel || "", q]);
    });
    if (toAppend.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, 4).setValues(toAppend);
    }
    return "Genesys Saved";
  },

  // --- DATA RETRIEVAL (INTELLIGENT) ---
  getConsolidatedData: function() {
    const todayStr = this._getTodayStr();
    const yesterdayStr = this._getYesterdayStr();

    const map = new Map();

    const getEntry = (name) => {
      const k = this._key(name);
      if (!map.has(k)) map.set(k, { role: "", absent: "", breaks: null, ot: [], genesys: null });
      return map.get(k);
    };
    
    // Reads overrides for Today AND Yesterday (to handle overnight shifts)
    const readSheet = (sheetName, nameIdx, dateIdx, callback) => {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        const data = sheet.getDataRange().getValues().slice(1);
        data.forEach(row => {
           if (this._matchDate(row[dateIdx], todayStr) || this._matchDate(row[dateIdx], yesterdayStr)) {
              callback(row);
           }
        });
      }
    };

    readSheet("Status_Overrides", 1, 4, (row) => {
       const e = getEntry(row[1]);
       if(row[2]) e.role = row[2];
       if(row[3]) e.absent = row[3];
    });
    readSheet("Break_Overrides", 1, 2, (row) => {
       const e = getEntry(row[1]);
       e.breaks = row[3];
    });
    readSheet("Overtime_Tracking", 1, 6, (row) => {
       const e = getEntry(row[1]);
       e.ot.push({ start: row[2], end: row[3], bStart: row[4], bEnd: row[5] });
    });

    // Genesys assignments are date-agnostic — read every row, latest wins.
    const gSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Genesys_Assignments");
    if (gSheet && gSheet.getLastRow() > 1) {
      gSheet.getDataRange().getValues().slice(1).forEach(row => {
        const e = getEntry(row[1]);
        const channel = row[2];
        if (channel) {
          let q = [];
          try { q = JSON.parse(row[3] || "[]"); } catch (err) { q = []; }
          e.genesys = { channel: String(channel), queues: q };
        } else {
          e.genesys = null;
        }
      });
    }
    return map;
  },

  // --- HELPERS ---
  // Status_Overrides / Break_Overrides / Overtime_Tracking are appended to and
  // never pruned, yet getConsolidatedData reads each one IN FULL on every poll
  // and filters to today/yesterday in JS — so the floor got slower every day it
  // ran. Drop rows older than the read window (with slack), the same way
  // Schedule_History already self-prunes. dateIdx is the DateStr column.
  PRUNE_KEEP_DAYS: 7,
  _prune: function(sheet, dateIdx) {
    try {
      const last = sheet.getLastRow();
      if (last < 400) return 0;   // don't pay for this on a small sheet
      const cutoff = Utilities.formatDate(
        new Date(Date.now() - this.PRUNE_KEEP_DAYS * 86400000), "America/Toronto", "yyyy-MM-dd");
      const data = sheet.getDataRange().getValues();
      let firstKeep = 1;
      while (firstKeep < data.length) {
        let d = data[firstKeep][dateIdx];
        if (d instanceof Date) d = Utilities.formatDate(d, "America/Toronto", "yyyy-MM-dd");
        if (String(d || '') >= cutoff) break;
        firstKeep++;
      }
      // Rows are appended in time order, so everything before the first row we
      // want to keep is stale. Bail out if that isn't a clean prefix.
      const drop = firstKeep - 1;
      if (drop > 0) { sheet.deleteRows(2, drop); return drop; }
      return 0;
    } catch (e) { return 0; }
  },

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

  // cleanName is already normalized by the caller; normalize the row too so a
  // re-spelled name updates the existing row instead of appending a duplicate.
  _matchRow: function(row, cleanName, dateStr, nameIdx, dateIdx) {
    return this._key(row[nameIdx]) === cleanName && this._matchDate(row[dateIdx], dateStr);
  }
};
