/**
 * MODULE: TRAINING TRACKER (LOCAL HOST ONLY)
 *
 * Self-contained corporate-training tracker. Piggybacks off the live floor
 * (compileFloorData) to surface who hasn't done a given training AND is
 * currently working, and produces a copy-pastable list to update the external
 * Corporate Training Tracker Google Sheet.
 *
 * Fully isolated: all logic + global wrappers live in this file. It only READS
 * existing global helpers (compileFloorData, _normalizeAgentKey). No other
 * module is modified. Storage = two dedicated tabs inside the bound workbook.
 */

const TrainingTracker = {
  DEFS_SHEET: "Training_Defs",
  TRK_SHEET: "Training_Tracker",

  EMPLOYEE_STATUS_OPTS: ["Active", "Long Term Disability", "Maternity/Paternity Leave", "Short Term Disability"],
  TRAINING_STATUS_OPTS: ["Completed", "LOA", "No Longer Employed"],
  REGION_SCOPE_OPTS: ["Onshore", "Both"],

  DEFS_HEADERS: ["id", "title", "videoLinkEN", "videoLinkFR", "regionScope", "createdDate", "notes"],
  TRK_HEADERS: ["trainingId", "employeeName", "employeeStatus", "trainingStatus", "completionDate", "comment"],

  // ── Sheet bootstrap ──────────────────────────────────────────────────────
  _getDefsSheet: function(ss) {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(this.DEFS_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(this.DEFS_SHEET);
      sheet.appendRow(this.DEFS_HEADERS);
      sheet.getRange(1, 1, 1, this.DEFS_HEADERS.length).setFontWeight("bold").setBackground("#e0e0e0");
      sheet.setFrozenRows(1);
      this._applyValidations(sheet, "defs");
    }
    return sheet;
  },

  _getTrackerSheet: function(ss) {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(this.TRK_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(this.TRK_SHEET);
      sheet.appendRow(this.TRK_HEADERS);
      sheet.getRange(1, 1, 1, this.TRK_HEADERS.length).setFontWeight("bold").setBackground("#e0e0e0");
      sheet.setFrozenRows(1);
      // Keep completionDate (col E) as literal text so yyyy-MM-dd round-trips
      // with the date input instead of being coerced to a locale date value.
      sheet.getRange(2, 5, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
      this._applyValidations(sheet, "tracker");
      this._applyConditionalFormats(sheet);
    }
    return sheet;
  },

  // Dropdowns. Sticky once set — only applied at sheet creation.
  _applyValidations: function(sheet, kind) {
    const build = (opts) => SpreadsheetApp.newDataValidation().requireValueInList(opts, true).setAllowInvalid(false).build();
    const maxRows = sheet.getMaxRows() - 1;
    if (kind === "defs") {
      // regionScope = column E (5)
      sheet.getRange(2, 5, maxRows, 1).setDataValidation(build(this.REGION_SCOPE_OPTS));
    } else {
      // employeeStatus = column C (3), trainingStatus = column D (4)
      sheet.getRange(2, 3, maxRows, 1).setDataValidation(build(this.EMPLOYEE_STATUS_OPTS));
      sheet.getRange(2, 4, maxRows, 1).setDataValidation(build(this.TRAINING_STATUS_OPTS));
    }
  },

  // Replicate the source Corporate Training Tracker visuals.
  _applyConditionalFormats: function(sheet) {
    const range = sheet.getRange("A2:F" + sheet.getMaxRows());

    const ruleCompleted = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D2="Completed"')
      .setStrikethrough(true)
      .setFontColor("#999999")
      .setRanges([range]).build();

    const ruleGone = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D2="No Longer Employed"')
      .setBackground("#f4cccc")
      .setRanges([range]).build();

    const ruleAttention = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($C2="Active",$D2="")')
      .setBackground("#fde0e0")
      .setRanges([range]).build();

    // Completed wins over the attention highlight — order matters (first match applies background).
    sheet.setConditionalFormatRules([ruleGone, ruleCompleted, ruleAttention]);
  },

  // ── Batched reads ────────────────────────────────────────────────────────
  _readDefs: function() {
    const sheet = this._getDefsSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, this.DEFS_HEADERS.length).getDisplayValues();
    return data.filter(r => r[0]).map(r => ({
      id: r[0], title: r[1], videoLinkEN: r[2], videoLinkFR: r[3],
      regionScope: r[4] || "Both", createdDate: r[5], notes: r[6]
    }));
  },

  _readTrackerRows: function() {
    const sheet = this._getTrackerSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, this.TRK_HEADERS.length).getDisplayValues();
    return data.map((r, i) => ({
      rowIndex: i + 2,
      trainingId: r[0], employeeName: r[1], employeeStatus: r[2],
      trainingStatus: r[3], completionDate: r[4], comment: r[5]
    })).filter(r => r.trainingId || r.employeeName);
  },

  // ── Training definitions CRUD ────────────────────────────────────────────
  getTrainings: function() {
    return JSON.stringify(this._readDefs());
  },

  saveTraining: function(payloadJson) {
    let p;
    try { p = (typeof payloadJson === "string") ? JSON.parse(payloadJson) : payloadJson; }
    catch (e) { return JSON.stringify({ ok: false, error: "Bad payload" }); }
    if (!p || !p.title) return JSON.stringify({ ok: false, error: "Missing title" });

    const sheet = this._getDefsSheet();
    const scope = (this.REGION_SCOPE_OPTS.indexOf(p.regionScope) !== -1) ? p.regionScope : "Both";
    let id = p.id;

    if (id) {
      // Update in place.
      const rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues() : [];
      for (let i = 0; i < rows.length; i++) {
        if (rows[i][0] === id) {
          const cur = sheet.getRange(i + 2, 1, 1, this.DEFS_HEADERS.length).getDisplayValues()[0];
          sheet.getRange(i + 2, 1, 1, this.DEFS_HEADERS.length).setValues([[
            id, p.title, p.videoLinkEN || "", p.videoLinkFR || "", scope, cur[5] || "", p.notes || ""
          ]]);
          return JSON.stringify({ ok: true, id: id });
        }
      }
    }
    // New record.
    id = "TR-" + Date.now();
    const createdDate = Utilities.formatDate(new Date(), "America/Toronto", "yyyy-MM-dd");
    sheet.appendRow([id, p.title, p.videoLinkEN || "", p.videoLinkFR || "", scope, createdDate, p.notes || ""]);
    return JSON.stringify({ ok: true, id: id });
  },

  deleteTraining: function(id) {
    if (!id) return JSON.stringify({ ok: false });
    // Remove the def row.
    const defs = this._getDefsSheet();
    if (defs.getLastRow() > 1) {
      const dRows = defs.getRange(2, 1, defs.getLastRow() - 1, 1).getDisplayValues();
      for (let i = dRows.length - 1; i >= 0; i--) {
        if (dRows[i][0] === id) { defs.deleteRow(i + 2); break; }
      }
    }
    // Bulk-remove all tracker rows for this id (filter + rewrite, no per-row deletes).
    const trk = this._getTrackerSheet();
    if (trk.getLastRow() > 1) {
      const all = trk.getRange(2, 1, trk.getLastRow() - 1, this.TRK_HEADERS.length).getValues();
      const keep = all.filter(r => String(r[0]) !== id);
      trk.getRange(2, 1, all.length, this.TRK_HEADERS.length).clearContent();
      if (keep.length > 0) trk.getRange(2, 1, keep.length, this.TRK_HEADERS.length).setValues(keep);
    }
    return JSON.stringify({ ok: true });
  },

  // ── Floor cross-reference ────────────────────────────────────────────────
  // key(normalized name) -> { display, region, active }
  getFloorIndex: function() {
    const idx = {};
    let floor = {};
    try { floor = JSON.parse(compileFloorData()); } catch (e) { floor = {}; }
    const workingBuckets = { active: true, startingSoon: true };
    Object.keys(floor).forEach(bucket => {
      const list = floor[bucket];
      if (!Array.isArray(list)) return;
      list.forEach(a => {
        if (!a || !a.name) return;
        const key = _normalizeAgentKey(a.name);
        const isWorking = !!workingBuckets[bucket];
        if (!idx[key]) {
          idx[key] = { display: a.name, region: a.region || "—", active: isWorking };
        } else if (isWorking) {
          idx[key].active = true;
          if (a.region) idx[key].region = a.region;
        }
      });
    });

    // Fold in WF_MASTERLIST so people not on today's floor still get a region.
    try {
      const ml = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("WF_MASTERLIST");
      if (ml && ml.getLastRow() > 1) {
        ml.getDataRange().getDisplayValues().slice(1).forEach(r => {
          if (!r[0]) return;
          const key = _normalizeAgentKey(r[0]);
          const isOffshore = String(r[4]).toUpperCase().includes("TI") || String(r[5]).includes("@") ||
                             String(r[4]).toUpperCase().includes("EL SALVADOR") || String(r[4]).toUpperCase().includes("GUATEMALA");
          if (!idx[key]) {
            idx[key] = { display: r[0], region: isOffshore ? "Offshore" : "Onshore", active: false };
          }
        });
      }
    } catch (e) {}

    return idx;
  },

  // ── Union / cross-reference roster for one training ──────────────────────
  getTrainingRoster: function(trainingId) {
    const defs = this._readDefs();
    const training = defs.filter(d => d.id === trainingId)[0] || null;
    if (!training) return JSON.stringify({ training: null, rows: [], counts: {}, copyList: "", discrepancies: { newHires: [], notOnFloor: [] } });

    const onshoreOnly = training.regionScope === "Onshore";
    const floorIdx = this.getFloorIndex();
    const trkRows = this._readTrackerRows().filter(r => r.trainingId === trainingId);

    const union = {};

    // 1. Seed from tracker rows (source of truth for status/date/comment).
    trkRows.forEach(r => {
      const key = _normalizeAgentKey(r.employeeName);
      const f = floorIdx[key];
      union[key] = {
        name: r.employeeName,
        region: f ? f.region : "—",
        active: f ? !!f.active : false,
        employeeStatus: r.employeeStatus || "Active",
        trainingStatus: r.trainingStatus || "",
        completionDate: r.completionDate || "",
        comment: r.comment || "",
        flag: f ? "matched" : "tracker-only"
      };
    });

    // 2. Add floor/masterlist people not yet in the tracker (possible new hires).
    Object.keys(floorIdx).forEach(key => {
      if (union[key]) return;
      const f = floorIdx[key];
      if (onshoreOnly && f.region !== "Onshore") return; // smart region scope
      union[key] = {
        name: f.display, region: f.region, active: !!f.active,
        employeeStatus: "Active", trainingStatus: "", completionDate: "", comment: "",
        flag: "floor-only"
      };
    });

    const rows = Object.keys(union).map(k => union[k]).sort((a, b) => a.name.localeCompare(b.name));

    const DONE = { "Completed": true, "LOA": true, "No Longer Employed": true };
    const counts = {
      totalEmployees: rows.length,
      completed: rows.filter(r => r.trainingStatus === "Completed").length,
      outstanding: rows.filter(r => !DONE[r.trainingStatus]).length
    };

    // Copy list: NON-COMPLETERS who are Active AND currently working on the floor.
    const copyList = rows
      .filter(r => r.employeeStatus === "Active" && r.trainingStatus !== "Completed" && r.active === true)
      .map(r => r.name + "\t" + r.employeeStatus)
      .join("\n");

    const discrepancies = {
      newHires: rows.filter(r => r.flag === "floor-only").map(r => r.name),
      notOnFloor: rows.filter(r => r.flag === "tracker-only").map(r => r.name)
    };

    return JSON.stringify({ training: training, rows: rows, counts: counts, copyList: copyList, discrepancies: discrepancies });
  },

  // ── Single completion upsert ─────────────────────────────────────────────
  setTrainingCompletion: function(trainingId, name, employeeStatus, trainingStatus, date, comment) {
    if (!trainingId || !name) return JSON.stringify({ ok: false });
    const sheet = this._getTrackerSheet();
    const targetKey = _normalizeAgentKey(name);
    const es = (this.EMPLOYEE_STATUS_OPTS.indexOf(employeeStatus) !== -1) ? employeeStatus : "Active";
    const ts = (this.TRAINING_STATUS_OPTS.indexOf(trainingStatus) !== -1) ? trainingStatus : "";

    if (sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]) === trainingId && _normalizeAgentKey(data[i][1]) === targetKey) {
          sheet.getRange(i + 2, 1, 1, this.TRK_HEADERS.length).setValues([[trainingId, name, es, ts, date || "", comment || ""]]);
          return JSON.stringify({ ok: true });
        }
      }
    }
    sheet.appendRow([trainingId, name, es, ts, date || "", comment || ""]);
    return JSON.stringify({ ok: true });
  },

  // ── Seed a training's roster from the floor/masterlist union ──────────────
  seedRoster: function(trainingId) {
    const defs = this._readDefs();
    const training = defs.filter(d => d.id === trainingId)[0];
    if (!training) return JSON.stringify({ ok: false, error: "Unknown training" });

    const onshoreOnly = training.regionScope === "Onshore";
    const floorIdx = this.getFloorIndex();
    const sheet = this._getTrackerSheet();

    const existing = {};
    this._readTrackerRows().filter(r => r.trainingId === trainingId).forEach(r => {
      existing[_normalizeAgentKey(r.employeeName)] = true;
    });

    const toAdd = [];
    Object.keys(floorIdx).forEach(key => {
      if (existing[key]) return;
      const f = floorIdx[key];
      if (onshoreOnly && f.region !== "Onshore") return;
      toAdd.push([trainingId, f.display, "Active", "", "", ""]);
    });

    if (toAdd.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, this.TRK_HEADERS.length).setValues(toAdd);
    }
    return JSON.stringify({ ok: true, added: toAdd.length });
  }
};

// ── Global wrappers (kept in this file to stay isolated) ────────────────────
function getTrainings() { return TrainingTracker.getTrainings(); }
function saveTraining(payloadJson) { return TrainingTracker.saveTraining(payloadJson); }
function deleteTraining(id) { return TrainingTracker.deleteTraining(id); }
function getTrainingRoster(id) { return TrainingTracker.getTrainingRoster(id); }
function setTrainingCompletion(id, name, es, ts, date, comment) { return TrainingTracker.setTrainingCompletion(id, name, es, ts, date, comment); }
function seedTrainingRoster(id) { return TrainingTracker.seedRoster(id); }
