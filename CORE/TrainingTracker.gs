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

  EMPLOYEE_STATUS_OPTS: ["Active", "Long Term Disability", "Maternity/Paternity Leave", "Short Term Disability", "Personal Leave"],
  TRAINING_STATUS_OPTS: ["Completed", "LOA", "No Longer Employed"],
  REGION_SCOPE_OPTS: ["Onshore", "Both"],

  DEFS_HEADERS: ["id", "title", "videoLinkEN", "videoLinkFR", "regionScope", "createdDate", "notes", "links"],
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
    return data.filter(r => r[0]).map(r => {
      let links = [];
      try { links = r[7] ? (JSON.parse(r[7]) || []) : []; } catch (e) { links = []; }
      // Backward compat: synthesize links from the legacy EN/FR video columns.
      if ((!links || !links.length)) {
        if (r[2]) links.push({ label: "Video EN", url: r[2] });
        if (r[3]) links.push({ label: "Video FR", url: r[3] });
      }
      return {
        id: r[0], title: r[1], videoLinkEN: r[2], videoLinkFR: r[3],
        regionScope: r[4] || "Both", createdDate: r[5], notes: r[6], links: links
      };
    });
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
          // Preserve auto-detected links (cur[7]); the manual editor doesn't manage them.
          sheet.getRange(i + 2, 1, 1, this.DEFS_HEADERS.length).setValues([[
            id, p.title, p.videoLinkEN || "", p.videoLinkFR || "", scope, cur[5] || "", p.notes || "", cur[7] || ""
          ]]);
          return JSON.stringify({ ok: true, id: id });
        }
      }
    }
    // New record.
    id = "TR-" + Date.now();
    const createdDate = Utilities.formatDate(new Date(), "America/Toronto", "yyyy-MM-dd");
    const links = JSON.stringify(this._linksFromEnFr(p.videoLinkEN, p.videoLinkFR));
    sheet.appendRow([id, p.title, p.videoLinkEN || "", p.videoLinkFR || "", scope, createdDate, p.notes || "", links]);
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

    // Mirror the corporate tracker's two formula columns so the UI + copy list
    // stay 1:1 with the source sheet.
    //   Planning (D)   = 1 when Active and not "No Longer Employed", else 0  (→ "expecting")
    //   Completion (F) = 1 when Training Status is "Completed", else 0       (→ "completed")
    rows.forEach(r => {
      r.planningRule = (r.employeeStatus === "Active" && r.trainingStatus !== "No Longer Employed") ? 1 : 0;
      r.completionCount = (r.trainingStatus === "Completed") ? 1 : 0;
    });

    const DONE = { "Completed": true, "LOA": true, "No Longer Employed": true };
    const counts = {
      totalEmployees: rows.length,
      expecting: rows.filter(r => r.planningRule === 1).length,
      completed: rows.filter(r => r.completionCount === 1).length,
      outstanding: rows.filter(r => !DONE[r.trainingStatus]).length
    };

    // Paste-ready copy list: an exact, row-for-row replica of the corporate
    // tracker's editable block (columns B→H), every employee in the same
    // alphabetical order. Drop floor-only new hires so the row count matches
    // the source sheet 1:1 — pasting at the first name cell refills it blind.
    // Dates are emitted as M/D/YYYY to match the source formatting exactly.
    const copyList = rows
      .filter(r => r.flag !== "floor-only")
      .map(r => [
        r.name,
        r.employeeStatus,
        r.planningRule,
        r.trainingStatus,
        r.completionCount,
        this._toMDY(r.completionDate),
        r.comment || ""
      ].join("\t"))
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
  },

  // ── Bulk import from an existing Corporate Training Tracker tab ────────────
  // One paste = one training. The tab name is the title (supplied by the user
  // since it isn't in the copied cells). Creates/updates the def, then merges
  // every parsed person row into the tracker. Re-importing is safe: people are
  // matched by normalized name and overwritten in place rather than duplicated.
  importTrainingSheet: function(payloadJson) {
    let p;
    try { p = (typeof payloadJson === "string") ? JSON.parse(payloadJson) : payloadJson; }
    catch (e) { return JSON.stringify({ ok: false, error: "Bad payload" }); }
    if (!p || !p.title || !p.raw) return JSON.stringify({ ok: false, error: "Missing title or pasted data" });

    const title = String(p.title).trim();
    const scope = (this.REGION_SCOPE_OPTS.indexOf(p.regionScope) !== -1) ? p.regionScope : "Both";

    // 1. Resolve or create the training definition (match by title, case-insensitive).
    const defsSheet = this._getDefsSheet();
    const defs = this._readDefs();
    let existing = null;
    for (let i = 0; i < defs.length; i++) {
      if (defs[i].title && defs[i].title.trim().toLowerCase() === title.toLowerCase()) { existing = defs[i]; break; }
    }

    // Auto-detect every link in the pasted block (videos, quizzes, forms…),
    // labelled by the keyword sitting next to each URL.
    const links = this._detectLinks(p.raw);
    const linksJson = JSON.stringify(links);
    const pick = (re) => { const m = links.filter(l => re.test(l.label.toLowerCase()))[0]; return m ? m.url : ""; };
    const vEN = pick(/video.*en|en.*video/), vFR = pick(/video.*fr|fr.*video/);

    let trainingId, created = false;
    if (existing) {
      trainingId = existing.id;
      // Refresh the detected links whenever any were found in this paste.
      if (links.length) {
        const dRows = defsSheet.getRange(2, 1, defsSheet.getLastRow() - 1, this.DEFS_HEADERS.length).getDisplayValues();
        for (let j = 0; j < dRows.length; j++) {
          if (dRows[j][0] === trainingId) {
            defsSheet.getRange(j + 2, 1, 1, this.DEFS_HEADERS.length).setValues([[
              trainingId, existing.title,
              vEN || dRows[j][2], vFR || dRows[j][3],
              existing.regionScope || scope, dRows[j][5] || "", existing.notes || "", linksJson
            ]]);
            break;
          }
        }
      }
    } else {
      trainingId = "TR-" + Date.now();
      const createdDate = Utilities.formatDate(new Date(), "America/Toronto", "yyyy-MM-dd");
      defsSheet.appendRow([trainingId, title, vEN, vFR, scope, createdDate, p.notes || "", linksJson]);
      created = true;
    }

    // 2. Parse the pasted tab into clean rows.
    const parsed = this._parseImportRows(p.raw);
    if (parsed.length === 0) {
      return JSON.stringify({ ok: true, trainingId: trainingId, title: title, created: created, imported: 0, added: 0, updated: 0, completed: 0, warn: "No employee rows detected — check that you pasted the data cells." });
    }

    // 3. Merge-upsert into the tracker (match by trainingId + normalized name).
    const trkSheet = this._getTrackerSheet();
    const all = (trkSheet.getLastRow() > 1)
      ? trkSheet.getRange(2, 1, trkSheet.getLastRow() - 1, this.TRK_HEADERS.length).getValues()
      : [];
    const indexByKey = {};
    for (let k = 0; k < all.length; k++) {
      if (String(all[k][0]) === trainingId) indexByKey[_normalizeAgentKey(all[k][1])] = k;
    }

    let added = 0, updated = 0, completed = 0;
    parsed.forEach(function(row) {
      if (row.trainingStatus === "Completed") completed++;
      const key = _normalizeAgentKey(row.name);
      const newRow = [trainingId, row.name, row.employeeStatus, row.trainingStatus, row.completionDate, row.comment];
      if (Object.prototype.hasOwnProperty.call(indexByKey, key)) { all[indexByKey[key]] = newRow; updated++; }
      else { indexByKey[key] = all.length; all.push(newRow); added++; }
    });

    // 4. Write the full data region back, keeping the date column as literal text.
    if (all.length > 0) {
      trkSheet.getRange(2, 5, all.length, 1).setNumberFormat("@");
      trkSheet.getRange(2, 1, all.length, this.TRK_HEADERS.length).setValues(all);
    }

    return JSON.stringify({ ok: true, trainingId: trainingId, title: title, created: created, imported: parsed.length, added: added, updated: updated, completed: completed, linksDetected: links.length });
  },

  // Build a links array from the legacy EN/FR video fields (manual entry path).
  _linksFromEnFr: function(en, fr) {
    const out = [];
    if (en) out.push({ label: "Video EN", url: String(en).trim() });
    if (fr) out.push({ label: "Video FR", url: String(fr).trim() });
    return out;
  },

  // Scan pasted text for every URL and label each by the keywords next to it
  // (video/quiz/form + EN/FR/onshore/offshore). Order preserved, duplicates
  // (same url + same label) dropped.
  _detectLinks: function(raw) {
    const lines = String(raw || "").replace(/\r/g, "").split("\n");
    const urlRe = /(https?:\/\/[^\s"'<>)\]]+)/g;
    const out = [], seen = {};
    lines.forEach(function(line) {
      let m;
      while ((m = urlRe.exec(line)) !== null) {
        const url = m[1].replace(/[.,;]+$/, "");
        const prefix = line.slice(0, m.index).toLowerCase();
        const u = url.toLowerCase();

        let type = "Link";
        if (prefix.indexOf("video") !== -1 || u.indexOf("benevity") !== -1 || u.indexOf("youtu") !== -1 || u.indexOf("vimeo") !== -1) type = "Video";
        else if (prefix.indexOf("quiz") !== -1) type = "Quiz";
        else if (prefix.indexOf("form") !== -1 || u.indexOf("docs.google.com/forms") !== -1 || u.indexOf("forms.gle") !== -1 || u.indexOf("forms.office") !== -1) type = "Form";

        let tag = "";
        if (prefix.indexOf("onshore") !== -1) tag = "Onshore";
        else if (prefix.indexOf("offshore") !== -1) tag = "Offshore";
        else if (/(^|[^a-z])(fr|french|francais|français)([^a-z]|$)/.test(prefix)) tag = "FR";
        else if (/(^|[^a-z])(en|english|anglais)([^a-z]|$)/.test(prefix)) tag = "EN";

        const label = (type + (tag ? " " + tag : "")).trim();
        const dedupeKey = label + "|" + url;
        if (!seen[dedupeKey]) { seen[dedupeKey] = true; out.push({ label: label, url: url }); }
      }
    });
    return out;
  },

  // Parse a pasted Corporate Training Tracker tab. Anchors on the Employee
  // Status column (a controlled vocabulary) to locate genuine data rows, which
  // sidesteps the summary block, instructions and multi-line headers entirely.
  _parseImportRows: function(raw) {
    const EMP = {}; this.EMPLOYEE_STATUS_OPTS.forEach(function(o){ EMP[o.toLowerCase()] = o; });
    const TRN = {}; this.TRAINING_STATUS_OPTS.forEach(function(o){ TRN[o.toLowerCase()] = o; });

    const out = [];
    const lines = String(raw).replace(/\r/g, "").split("\n");
    for (let i = 0; i < lines.length; i++) {
      const cells = lines[i].split("\t").map(function(c){ return (c || "").trim(); });

      // Locate the employee-status anchor.
      let si = -1;
      for (let c = 0; c < cells.length; c++) {
        if (cells[c] && Object.prototype.hasOwnProperty.call(EMP, cells[c].toLowerCase())) { si = c; break; }
      }
      if (si < 1) continue; // need at least a name cell in front of the status

      // Name = nearest non-empty cell before the status column.
      let name = "";
      for (let n = si - 1; n >= 0; n--) { if (cells[n]) { name = cells[n]; break; } }
      if (!name) continue;

      // Fixed source layout: name | status | rule(1/0) | trainingStatus | count(1/0) | date | comment
      const employeeStatus = EMP[cells[si].toLowerCase()];
      const trainingStatus = TRN[(cells[si + 2] || "").toLowerCase()] || "";
      const completionDate = this._parseMDY(cells[si + 4] || "");
      const comment = cells[si + 5] || "";

      out.push({ name: name, employeeStatus: employeeStatus, trainingStatus: trainingStatus, completionDate: completionDate, comment: comment });
    }
    return out;
  },

  // "5/3/2026" -> "2026-05-03". Pass through already-ISO dates; blank otherwise.
  _parseMDY: function(v) {
    v = String(v || "").trim();
    if (!v) return "";
    const m = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return m[3] + "-" + ("0" + m[1]).slice(-2) + "-" + ("0" + m[2]).slice(-2);
    if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
    return "";
  },

  // "2026-05-03" -> "5/3/2026" (no leading zeros) to mirror the source sheet.
  _toMDY: function(v) {
    v = String(v || "").trim();
    const m = v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return parseInt(m[2], 10) + "/" + parseInt(m[3], 10) + "/" + m[1];
    return v; // already M/D/YYYY or blank — pass through
  }
};

// ── Global wrappers (kept in this file to stay isolated) ────────────────────
function getTrainings() { return TrainingTracker.getTrainings(); }
function saveTraining(payloadJson) { return TrainingTracker.saveTraining(payloadJson); }
function deleteTraining(id) { return TrainingTracker.deleteTraining(id); }
function getTrainingRoster(id) { return TrainingTracker.getTrainingRoster(id); }
function setTrainingCompletion(id, name, es, ts, date, comment) { return TrainingTracker.setTrainingCompletion(id, name, es, ts, date, comment); }
function seedTrainingRoster(id) { return TrainingTracker.seedRoster(id); }
function importTrainingSheet(payloadJson) { return TrainingTracker.importTrainingSheet(payloadJson); }
