/**
 * MODULE: WORKFORCE TRACKER
 */

/**
 * Canonical key for agent-name matching across MasterList, WFM, GEM, and DB_Sessions.
 * The same person shows up in different sources in any of these forms:
 *   "Bennani, Mohammed"   (Last, First — comma)
 *   "Mohammed Bennani"    (First Last  — no comma)
 *   "Bennani Mohammed"    (Last First  — no comma, French-Canadian style)
 * To collapse all of them into the same key, we strip diacritics + commas
 * + hyphens, lowercase, and SORT the tokens alphabetically. Order-agnostic
 * so any combination of first/last + comma yes/no produces the same hash.
 */
function _normalizeAgentKey(s) {
  var raw = String(s == null ? '' : s).trim();
  if (!raw) return '';
  var clean = raw
    .replace(/,/g, ' ')        // commas treated as separators, not part of the name
    .normalize('NFD')          // decompose accented chars
    .replace(/[̀-ͯ]/g, '') // strip combining marks
    .toLowerCase()
    .replace(/[-_]+/g, ' ')    // hyphens/underscores → spaces
    .replace(/[^\w\s]/g, ' ')  // any other punctuation → space
    .replace(/\s+/g, ' ')
    .trim();
  if (!clean) return '';
  var tokens = clean.split(' ').filter(Boolean);
  tokens.sort();
  return tokens.join(' ');
}

/**
 * Unicode-safe title-case. JS \b\w treats accented chars as non-word,
 * so 'Jaén' would turn into 'JaéN'. This preserves Unicode letters.
 */
function _titleCaseName(s) {
  return String(s == null ? '' : s)
    .replace(/(^|[\s,\-'])(\S)/g, function(m, p, c) { return p + c.toUpperCase(); })
    .trim();
}

// Absence taxonomy (per product owner):
//   APPROVED  — ASCLU / SLU / Furlough / ACSU: company-approved voluntary
//               early outs on low demand. NOT absenteeism; never counted in
//               absence totals, monthly flags, Bradford or Mgmt graphs.
//   LATE      — ALU: the agent showed up late / missed part of the day.
//               Not a full absence; tracked as its own "Lates" series.
//   ABSENT    — SICK / UNAB / COMP / COMPU / TI AWOL / LOA (incl. TI LOA):
//               real absenteeism.
var APPROVED_LEAVE_RGX = /\b(asclu|slu|furlough|acsu)\b/i;
var LATE_RGX = /\balu\b/i;
// "TI Mentor" is offshore mentoring, NOT a coaching session (per product
// owner) — it must never land in WF_COACHING even though it contains the
// 'mentor' coaching code.
var COACH_EXCLUDE_RGX = /ti[\s_\-]*mentor/i;

var WorkforceTracker = {

  _getDB: function(sheetName) {
      return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  },

  // normalized-name-key → WFM agent ID (TID), from WF_AGENT_IDS (written at
  // import). Lets SAFE hours match on Employee ID exactly. Empty if never built.
  _agentTidMap: function() {
      var out = {};
      try {
          var sh = this._getDB('WF_AGENT_IDS');
          if (sh && sh.getLastRow() > 1) {
              sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function(r) {
                  var k = String(r[0] || ''), tid = String(r[2] || '').trim();
                  if (k && tid) out[k] = tid;
              });
          }
      } catch (e) {}
      return out;
  },

  // Read several sheets in ONE Sheets API round-trip instead of one
  // getDataRange() call per sheet. Returns { sheetName: rows-incl-header },
  // formatted values padded rectangular so it's a drop-in for
  // getDataRange().getDisplayValues(). Falls back to per-sheet SpreadsheetApp
  // reads if the Sheets advanced service isn't enabled — so it degrades
  // gracefully and never hard-fails.
  _batchDisplayValues: function(names) {
      var out = {};
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var existing = names.filter(function(n) { return ss.getSheetByName(n); });
      names.forEach(function(n) { out[n] = []; });
      if (existing.length) {
          try {
              if (typeof Sheets !== 'undefined' && Sheets.Spreadsheets && Sheets.Spreadsheets.Values) {
                  var resp = Sheets.Spreadsheets.Values.batchGet(ss.getId(), {
                      ranges: existing,
                      valueRenderOption: 'FORMATTED_VALUE',
                      dateTimeRenderOption: 'FORMATTED_STRING'
                  });
                  var vr = (resp && resp.valueRanges) || [];
                  for (var i = 0; i < existing.length; i++) {
                      out[existing[i]] = this._rectangular((vr[i] && vr[i].values) ? vr[i].values : []);
                  }
                  return out;
              }
          } catch (e) { Logger.log('[WT] batchGet failed, falling back to per-sheet reads: ' + e); }
          // Fallback: original per-sheet behavior.
          existing.forEach(function(n) {
              var sh = ss.getSheetByName(n);
              out[n] = (sh && sh.getLastRow() > 0) ? sh.getDataRange().getDisplayValues() : [];
          });
      }
      return out;
  },

  // batchGet omits trailing empty cells (ragged rows); pad to a rectangle so
  // row[i] matches getDisplayValues (which returns '' for blanks, not undefined).
  _rectangular: function(rows) {
      var w = 0, i;
      for (i = 0; i < rows.length; i++) if (rows[i].length > w) w = rows[i].length;
      for (i = 0; i < rows.length; i++) {
          var r = rows[i];
          for (var k = 0; k < w; k++) if (r[k] == null) r[k] = '';
      }
      return rows;
  },

  importData: function(schedRaw, idpRaw) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ['WF_COACHING', 'WF_FURLOUGH', 'WF_ROLES', 'WF_IDP', 'WF_ABSENCES'].forEach(n => {
       if(!ss.getSheetByName(n)) ss.insertSheet(n);
    });
    // Same per-agent sheet-write problem as ImportHandler — RegionRegistry
    // is consulted/upserted inside flushAgentBuffer for every agent in the
    // import. Wrap the whole import in a batch so writes are coalesced
    // into one sheet read + one sheet write at the end.
    var _regionBatch = (typeof RegionRegistry !== 'undefined');
    if (_regionBatch) RegionRegistry.beginBatch();
    try {
    let msg = [];
    let schedDates = [];
    let idpDates = [];
    let muSet = "";
    if (schedRaw && schedRaw.trim().length > 0) {
      let cleanCoach = [];
      let cleanFurlough = [];
      let cleanRoles = [];
      let cleanAbsences = [];
      let agentIds = {};   // normalized key → { id (WFM Agent: NNNNN), name } for ID-exact SAFE matching

      const lines = schedRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      let currentAgent = "", currentID = "", currentDateStr = "";
      let currentY = 0, currentM = 0, currentD = 0;
      let lastTimeMins = -1, daysAdded = 0;
      let agentBuffer = []; 
      
      const segmentRegex = /([a-zA-ZÀ-ÿ0-9\/\(\)\s\-\.&']+?)\s+(\d{1,2}:\d{2}(?:\s?[AP]M)?)\s+(\d{1,2}:\d{2}(?:\s?[AP]M)?)\s*$/i;
      const dateRegex = /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/;
      const COACHING_CODES = ['ce séance', 'huddle', 'echo', 'mentor', 'hsc', 'health and safety', 'meet', 'roadshow', 'one on one', 'individuelle', 'pulsecheck', 'qual session', 'quality', 'sbys', 'survey', 'sondage en ligne', 'team'];
      const ACSU_CODES = ['acsu', 'solicited', 'libération', 'voluntary', 'asclu', 'slu', 'furlough'];
      // Word-aware role matcher. Bare substring checks turned words like
      // "vehicle"/"article" (icl) and "feuille" (feu) into role hours.
      // Leading-only boundaries on icl/ulc keep variants like "ICL2" matching.
      const ROLE_RGX = /safe onqueue|safe en ligne|wofqt|woqft|\btower\b|\breading\b|\blecture\b/; // ICL/ULC FIRE intentionally excluded — not tracked (live Floor only)
      
      const flushAgentBuffer = () => {
          if (agentBuffer.length === 0) return;
          // Capture the WFM agent ID (the "Agent: NNNNN" number) keyed by name, so
          // SAFE-report hours can be matched on Employee ID (exact) instead of name.
          if (currentID && currentAgent) agentIds[_normalizeAgentKey(currentAgent)] = { id: String(currentID).trim(), name: currentAgent };
          let isOffshore = false;
          let detectedBy = null;
          if (currentID && String(currentID).startsWith("3")) { isOffshore = true; detectedBy = 'auto-wfm-id'; }
          for (let obj of agentBuffer) {
              if (obj.raw.toUpperCase().includes("TI ") || obj.raw.toUpperCase().includes("OFFSHORE")) {
                  isOffshore = true;
                  if (!detectedBy) detectedBy = 'auto-wfm-keyword';
                  break;
              }
          }

          // Region Registry is the source of truth. A manual / masterlist entry
          // always wins over anything we detect here. Otherwise we upsert with
          // the confidence we detected (id-prefix > keyword > default).
          if (typeof RegionRegistry !== 'undefined') {
              const registered = RegionRegistry.getRegion(currentAgent);
              const src = RegionRegistry.getSource(currentAgent);
              if (registered && (src === 'manual' || src === 'masterlist')) {
                  isOffshore = registered === 'Offshore';
              } else {
                  RegionRegistry.upsert(currentAgent, isOffshore ? 'Offshore' : 'Onshore', detectedBy || 'auto-wfm-default');
              }
          }
          let reg = isOffshore ? "Offshore" : "Onshore";
          
          agentBuffer.forEach(obj => { 
              let actLower = obj.act.toLowerCase();
              let isTeamLead = actLower.includes('team lead') || actLower.includes('équipe') || actLower.includes('equipe');
              
              let isCoach = !isTeamLead && !COACH_EXCLUDE_RGX.test(actLower) && COACHING_CODES.some(c => actLower.includes(c));
              let isFurlough = ACSU_CODES.some(c => actLower.includes(c));
              let isRole = ROLE_RGX.test(actLower);
              
              let isAbsence = false;
              let absType = "";
              if (actLower.includes('sick') || actLower.includes('maladie') || actLower.includes('sicu')) { isAbsence = true; absType = 'SICK'; }
              else if (actLower.includes('unab')) { isAbsence = true; absType = 'UNAB'; }
              else if (/\bcompu\b/.test(actLower)) { isAbsence = true; absType = 'COMPU'; }
              // Whole-word matches only: a bare substring check turned activities
              // like "Compliance" into COMP and "Evaluation" into ALU absences,
              // silently inflating absence counts.
              else if (/\bcomp\b/.test(actLower)) { isAbsence = true; absType = 'COMP'; }
              else if (/\balu\b/.test(actLower)) { isAbsence = true; absType = 'ALU'; }
              else if (/\bloa\b/.test(actLower)) { isAbsence = true; absType = 'LOA'; }
              else if (actLower.includes('awol') || actLower.includes('ncns')) { isAbsence = true; absType = 'TI AWOL'; }

              let roleType = "";
              if (isRole) {
                  if (actLower.includes('safe')) roleType = "SAFE";
                  else if (actLower.includes('wofqt') || actLower.includes('woqft') || actLower.includes('tower')) roleType = "TOWER";
                  else if (actLower.includes('reading') || actLower.includes('lecture')) roleType = "READING";
              }
              
              if (isCoach) cleanCoach.push([currentAgent, obj.dateStr, obj.act, obj.start, obj.end, reg]);
              if (isFurlough && !isOffshore) cleanFurlough.push([currentAgent, obj.dateStr, obj.act, obj.start, obj.end, reg]);
              if (isRole) cleanRoles.push([currentAgent, obj.dateStr, roleType, obj.start, obj.end, reg]);
              if (isAbsence) cleanAbsences.push([currentAgent, obj.dateStr, absType, obj.start, obj.end, reg]);
              
              schedDates.push(obj.dateStr);
          });
          agentBuffer = [];
      };

      lines.forEach(line => {
        let text = line.trim();
        if(text.startsWith('Agent Name') || text.startsWith('"Agent Name"')) return;

        if (text.includes('Agent:')) {
          flushAgentBuffer();
          let parts = text.split(':');
          if (parts.length > 1) {
              let agentData = parts[1].trim();
              let idMatch = agentData.match(/^(\d+)/);
              currentID = idMatch ? idMatch[1] : "";
              currentAgent = agentData.replace(/^\d+\s+/, '').trim();
          }
          return;
        } 
        
        if (text.includes('"') && text.includes(',')) {
          flushAgentBuffer(); 
          let csvParts = this._parseCSVLine(text);
          if (csvParts.length >= 6) { 
              let actLower = this._cleanActivity(csvParts[2]).toLowerCase();
              let isTeamLead = actLower.includes('team lead') || actLower.includes('équipe') || actLower.includes('equipe');
              
              let isCoach = !isTeamLead && !COACH_EXCLUDE_RGX.test(actLower) && COACHING_CODES.some(c => actLower.includes(c));
              let isFurlough = ACSU_CODES.some(c => actLower.includes(c));
              let isRole = ROLE_RGX.test(actLower);

              let isAbsence = false;
              let absType = "";
              if (actLower.includes('sick') || actLower.includes('maladie') || actLower.includes('sicu')) { isAbsence = true; absType = 'SICK'; }
              else if (actLower.includes('unab')) { isAbsence = true; absType = 'UNAB'; }
              else if (/\bcompu\b/.test(actLower)) { isAbsence = true; absType = 'COMPU'; }
              // Whole-word matches only: a bare substring check turned activities
              // like "Compliance" into COMP and "Evaluation" into ALU absences,
              // silently inflating absence counts.
              else if (/\bcomp\b/.test(actLower)) { isAbsence = true; absType = 'COMP'; }
              else if (/\balu\b/.test(actLower)) { isAbsence = true; absType = 'ALU'; }
              else if (/\bloa\b/.test(actLower)) { isAbsence = true; absType = 'LOA'; }
              else if (actLower.includes('awol') || actLower.includes('ncns')) { isAbsence = true; absType = 'TI AWOL'; }
              
              let roleType = "";
              if (isRole) {
                  if (actLower.includes('safe')) roleType = "SAFE";
                  else if (actLower.includes('wofqt') || actLower.includes('woqft') || actLower.includes('tower')) roleType = "TOWER";
                  else if (actLower.includes('reading') || actLower.includes('lecture')) roleType = "READING";
              }

              let isOff = csvParts[5] && csvParts[5].includes("Offshore");
              // Registry consultation + upsert, same hierarchy rule as flushAgentBuffer.
              if (typeof RegionRegistry !== 'undefined' && csvParts[0]) {
                  const registered = RegionRegistry.getRegion(csvParts[0]);
                  const src = RegionRegistry.getSource(csvParts[0]);
                  if (registered && (src === 'manual' || src === 'masterlist')) {
                      isOff = registered === 'Offshore';
                  } else {
                      RegionRegistry.upsert(csvParts[0], isOff ? 'Offshore' : 'Onshore', 'auto-wfm-keyword');
                  }
              }
              let pDate = this._parseDate(csvParts[1]);
              
              if (isCoach) cleanCoach.push([csvParts[0], pDate, this._cleanActivity(csvParts[2]), csvParts[3], csvParts[4], csvParts[5]]);
              if (isFurlough && !isOff) cleanFurlough.push([csvParts[0], pDate, this._cleanActivity(csvParts[2]), csvParts[3], csvParts[4], csvParts[5]]);
              if (isRole) cleanRoles.push([csvParts[0], pDate, roleType, csvParts[3], csvParts[4], csvParts[5]]);
              if (isAbsence) cleanAbsences.push([csvParts[0], pDate, absType, csvParts[3], csvParts[4], csvParts[5]]);
              
              schedDates.push(pDate);
          }
          return;
        }

        let dMatch = text.match(dateRegex);
        if (dMatch) {
            currentDateStr = this._parseDate(dMatch[1]);
            [currentY, currentM, currentD] = currentDateStr.split('-').map(Number);
            lastTimeMins = -1;
            daysAdded = 0;
        }

        if (currentAgent && currentY) {
          let segMatch = text.match(segmentRegex);
          if (segMatch) {
            let act = this._cleanActivity(segMatch[1].trim());
            let tStartMins = this._timeToMins(segMatch[2].trim());
            if (lastTimeMins > -1 && tStartMins < lastTimeMins) daysAdded++;
            lastTimeMins = tStartMins;
            let actDateStr = Utilities.formatDate(new Date(currentY, currentM - 1, currentD + daysAdded), Session.getScriptTimeZone(), "yyyy-MM-dd");
            if (!act.toLowerCase().match(/^activity|^scheduled/)) agentBuffer.push({ raw: text, dateStr: actDateStr, act: act, start: segMatch[2].trim(), end: segMatch[3].trim() });
          }
        }
      });
      flushAgentBuffer();
      
      if (cleanCoach.length > 0) {
        this._executeDestructiveUpsert('WF_COACHING', cleanCoach, ['Agent Name', 'Date', 'Activity', 'Start Time', 'End Time', 'Region']);
        msg.push(`Coaching`);
      }
      if (cleanFurlough.length > 0) {
        this._executeDestructiveUpsert('WF_FURLOUGH', cleanFurlough, ['Agent Name', 'Date', 'Activity', 'Start Time', 'End Time', 'Region']);
        msg.push(`Furloughs`);
      }
      if (cleanRoles.length > 0) {
        cleanRoles.sort((a, b) => b[1].localeCompare(a[1]));
        this._executeDestructiveUpsert('WF_ROLES', cleanRoles, ['Agent Name', 'Date', 'Role', 'Start Time', 'End Time', 'Region']);
        msg.push(`Roles`);
      }
      if (cleanAbsences.length > 0) {
        this._executeDestructiveUpsert('WF_ABSENCES', cleanAbsences, ['Agent Name', 'Date', 'Type', 'Start Time', 'End Time', 'Region']);
        msg.push(`Absences`);
      }

      // Persist the WFM agent-ID map (MERGE — IDs survive partial imports), so
      // SAFE-report hours can be matched on Employee ID instead of name.
      try {
        var idKeys = Object.keys(agentIds);
        if (idKeys.length) {
          var ssIds = SpreadsheetApp.getActiveSpreadsheet();
          var idSh = ssIds.getSheetByName('WF_AGENT_IDS') || ssIds.insertSheet('WF_AGENT_IDS');
          var prevRows = idSh.getLastRow();
          var existing = {};
          if (prevRows > 1) {
            idSh.getRange(2, 1, prevRows - 1, 3).getValues().forEach(function(r) { var k = String(r[0]); if (k) existing[k] = { name: String(r[1] || ''), id: String(r[2] || '') }; });
          }
          idKeys.forEach(function(k) { existing[k] = { name: agentIds[k].name, id: agentIds[k].id }; });
          var outIds = [['Key', 'Name', 'TID']].concat(Object.keys(existing).map(function(k) { return [k, existing[k].name, existing[k].id]; }));
          idSh.getRange(1, 1, outIds.length, 3).setValues(outIds);
          if (prevRows > outIds.length) idSh.getRange(outIds.length + 1, 1, prevRows - outIds.length, 3).clearContent();
        }
      } catch (idErr) { Logger.log('[AgentIDs] map skipped: ' + idErr); }

      // Overtime is parsed in its own isolated module (WF_OVERTIME sheet).
      // Wrapped so any OT parse error can never break the core import above.
      try {
        if (typeof OvertimeTracker !== 'undefined') {
          var otRes = OvertimeTracker.importFromSchedule(schedRaw);
          if (otRes && otRes.indexOf('segment') !== -1) msg.push('Overtime');
        }
      } catch (otErr) { Logger.log('[OT] import skipped: ' + otErr); }
    }

    if (idpRaw && idpRaw.trim().length > 0) {
      let cleanIDP = [];
      const lines = idpRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      
      for (let i = 0; i < Math.min(lines.length, 15); i++) {
          let muMatch = lines[i].match(/MU Set:\s*(\d+)/i);
          if (muMatch) { muSet = muMatch[1]; break; }
      }

      let headerIdx = lines.findIndex(l => { const low = l.toLowerCase(); return (low.includes('requirements') || low.includes('besoins')) && (low.includes('occupied seats') || low.includes('sièges occupés')); });
      if (headerIdx > -1) {
        let headers = this._parseCSVLine(lines[headerIdx]);
        let colMap = {};
        headers.forEach((h, i) => {
          let match = h.match(/(\w+\s\d{1,2},?\s\d{4})/);
          if (match) {
            let dStr = this._parseDate(match[1]);
            if (h.toLowerCase().includes('req') || h.toLowerCase().includes('besoin')) colMap[i] = { date: dStr, type: 'req' };
            else if ((h.toLowerCase().includes('open') || h.toLowerCase().includes('ouvert')) && !h.toLowerCase().includes('+/-')) colMap[i] = { date: dStr, type: 'open' };
          }
        });
        let dataByDay = {};
        const TIME_ROW_RE = /^\s*\d{1,2}[:\.]\d{2}(?:[:\.]\d{2})?(?:\s*[AP]M)?\s*$/i;
        for (let i = headerIdx + 1; i < lines.length; i++) {
          let cols = this._parseCSVLine(lines[i]);
          if (cols[0] && TIME_ROW_RE.test(cols[0])) {
            let tNorm = this._formatTimeStr(cols[0]);
            Object.keys(colMap).forEach(idx => {
               if (cols[idx] !== undefined) {
                 if (!dataByDay[colMap[idx].date]) dataByDay[colMap[idx].date] = {};
                 if (!dataByDay[colMap[idx].date][tNorm]) dataByDay[colMap[idx].date][tNorm] = { req:0, open:0 };
                 let val = parseFloat(String(cols[idx]).replace(/,/g, '')) || 0;
                 dataByDay[colMap[idx].date][tNorm][colMap[idx].type] = val;
               }
            });
          }
        }
        Object.keys(dataByDay).forEach(day => {
            idpDates.push(day);
            Object.keys(dataByDay[day]).forEach(time => cleanIDP.push([day, time, dataByDay[day][time].req, dataByDay[day][time].open]))
        });
        if (cleanIDP.length > 0) {
          this._executeDestructiveUpsert('WF_IDP', cleanIDP, ['Day', 'Interval', 'Required', 'Open']);
          msg.push(`IDP`);
        }
      }
    }

    try {
        const props = PropertiesService.getDocumentProperties();
        const curTime = Utilities.formatDate(new Date(), "America/Toronto", "MMM dd, HH:mm");
        
        const getMinMax = (sheetName, colIdx) => {
            const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
            if (!s || s.getLastRow() < 2) return null;
            let vals = s.getRange(2, colIdx, s.getLastRow() - 1, 1).getDisplayValues().flat().filter(v => String(v).trim() !== "");
            if (vals.length === 0) return null;
            
            vals = vals.map(v => {
                let m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
                return m ? v : null;
            }).filter(v => v);
            if (vals.length === 0) return null;
            vals.sort();
            return { min: vals[0], max: vals[vals.length - 1] };
        };
        const roleRange = getMinMax('WF_ROLES', 2);
        const furlRange = getMinMax('WF_FURLOUGH', 2);
        
        let minS = null, maxS = null;
        if (roleRange) { minS = roleRange.min; maxS = roleRange.max; }
        if (furlRange) { 
            if (!minS || furlRange.min < minS) minS = furlRange.min;
            if (!maxS || furlRange.max > maxS) maxS = furlRange.max;
        }
        
        if (minS && maxS) props.setProperty('SYNC_SCHED', `WFM: ${minS} to ${maxS} (Sync: ${curTime})`);
        const idpRange = getMinMax('WF_IDP', 1);
        if (idpRange) props.setProperty('SYNC_IDP', `IDP: ${idpRange.min} to ${idpRange.max} (Sync: ${curTime})`);
        
        if (muSet) props.setProperty('MU_SET', muSet);
    } catch(e) {}
    
    // --- DEFERRED AUTO-ARCHIVE ---
    // archiveUnifiedReport reads ~8 sheets × 3 cycles per affected month and was the
    // dominant import cost (it alone could blow the 6-min limit). QUEUE the months
    // instead; the client calls flushPendingArchives() right after import, in its own
    // execution budget. Past-month cycle filters get archived a beat later, not never.
    let affectedMonths = new Set();
    schedDates.concat(idpDates).forEach(d => {
        if (d && d.length >= 7) affectedMonths.add(d.substring(0, 7) + "-01");
    });
    let queued = [];
    if (affectedMonths.size) {
        try {
            let aprops = PropertiesService.getScriptProperties();
            let pend = {}; try { pend = JSON.parse(aprops.getProperty('PENDING_ARCHIVE') || '{}'); } catch (e) {}
            affectedMonths.forEach(m => { pend[m] = true; queued.push(m.substring(0, 7)); });
            aprops.setProperty('PENDING_ARCHIVE', JSON.stringify(pend));
        } catch (e) {}
    }

    // Invalidate the SAFE analytics cache so the next view recomputes on fresh data.
    try { PropertiesService.getScriptProperties().setProperty('WF_CACHE_VER', String(Date.now())); } catch (e) {}

    if (msg.length === 0) return `Basic Schedule Synced. (Archiving queued: ${queued.join(', ')})`;
    return `Synced: ${msg.join(' | ')}. Archiving queued: ${queued.join(', ')}`;
    } finally {
      if (_regionBatch) RegionRegistry.commitBatch();
    }
  },

  // NEW: Engine to fetch Yearly Agent Profiles
  getAbsenceProfiles: function(year) {
      if (!year) year = new Date().getFullYear().toString();
      
      const dbAbs = this._getDB('WF_ABSENCES');
      const dbML = this._getDB('WF_MASTERLIST');
      
      let agents = {};
      
      // 1. Cross-reference with MasterList to get clean Supervisor names and Region
      if (dbML && dbML.getLastRow() > 1) {
          dbML.getDataRange().getDisplayValues().slice(1).forEach(r => {
              let cleanName = String(r[0]).trim();
              let key = _normalizeAgentKey(cleanName);
              let isOffshore = String(r[4]).toUpperCase().includes("TI") || String(r[5]).includes("@") || String(r[4]).toUpperCase().includes("EL SALVADOR") || String(r[4]).toUpperCase().includes("GUATEMALA");
              // Registry is authoritative if set.
              if (typeof RegionRegistry !== 'undefined') {
                  const rg = RegionRegistry.getRegion(cleanName);
                  if (rg) isOffshore = rg === 'Offshore';
              }
              agents[key] = {
                  name: cleanName,
                  supervisor: r[2] || "Unassigned",
                  region: isOffshore ? "Offshore" : "Onshore",
                  records: [],
                  totals: { SICK: 0, UNAB: 0, COMP: 0, COMPU: 0, ALU: 0, 'TI AWOL': 0, 'LOA': 0 },
                  monthlyCounts: {},
                  totalAbsences: 0,
                  flags: 0,
                  flaggedMonths: [] // NEW: Collect the specific months
              };
          });
      }

      if (!dbAbs || dbAbs.getLastRow() < 2) return JSON.stringify({ year: year, profiles: [] });

      // 2. Iterate through Absences Database (de-dup identical rows)
      let seenAbs = {};
      dbAbs.getDataRange().getDisplayValues().slice(1).forEach(row => {
          let dStr = this._formatDate(row[1]);
          if (dStr.startsWith(year)) {
              let aName = String(row[0]).trim();
              let key = _normalizeAgentKey(aName);
              let type = String(row[2]).trim();
              let start = String(row[3]).trim();
              let end = String(row[4]).trim();

              let dedupKey = key + '|' + dStr + '|' + type + '|' + start + '|' + end;
              if (seenAbs[dedupKey]) return;
              seenAbs[dedupKey] = true;

              if (!agents[key]) {
                  let reg = String(row[5]).trim();
                  if (!reg) reg = "Onshore";
                  if (typeof RegionRegistry !== 'undefined') {
                      const rg = RegionRegistry.getRegion(aName);
                      if (rg) reg = rg;
                  }
                  agents[key] = {
                      name: aName, supervisor: "Unknown", region: reg,
                      records: [], totals: { SICK: 0, UNAB: 0, COMP: 0, COMPU: 0, ALU: 0, 'TI AWOL': 0, 'LOA': 0 },
                      monthlyCounts: {}, totalAbsences: 0, flags: 0, flaggedMonths: []
                  };
              }

              agents[key].records.push({ date: dStr, type: type, start: start, end: end });
          }
      });

      // 2b. Merge same-day same-type segments (e.g. split-shift UNAB
      // morning + afternoon → single entry spanning earliest start to
      // latest end). Then compute totals from the merged records.
      const self = this;
      Object.values(agents).forEach(ag => {
          let groups = {};
          ag.records.forEach(r => {
              let gk = r.date + '|' + r.type;
              if (!groups[gk]) groups[gk] = [];
              groups[gk].push(r);
          });
          let merged = [];
          Object.keys(groups).forEach(gk => {
              let segs = groups[gk];
              if (segs.length === 1) { merged.push(segs[0]); return; }
              let earliest = segs[0], latest = segs[0];
              let earliestMins = self._timeToMins(segs[0].start);
              let latestMins = self._timeToMins(segs[0].end);
              segs.forEach(s => {
                  let sm = self._timeToMins(s.start);
                  let em = self._timeToMins(s.end);
                  if (sm < earliestMins) { earliestMins = sm; earliest = s; }
                  if (em > latestMins) { latestMins = em; latest = s; }
              });
              merged.push({ date: segs[0].date, type: segs[0].type, start: earliest.start, end: latest.end });
          });
          ag.records = merged;
          ag.totals = { SICK: 0, UNAB: 0, COMP: 0, COMPU: 0, ALU: 0, 'TI AWOL': 0, 'LOA': 0 };
          ag.monthlyCounts = {};
          ag.totalAbsences = 0;
          ag.approvedLeave = 0;
          ag.lates = 0;
          const isApproved = (t) => APPROVED_LEAVE_RGX.test(String(t));
          const isLate = (t) => LATE_RGX.test(String(t));
          merged.forEach(r => {
              if (ag.totals[r.type] !== undefined) ag.totals[r.type]++;
              else ag.totals[r.type] = 1;
              // Approved voluntary leave: visible (badge + timeline), never
              // absenteeism. ALU lates: own counter, not a full absence.
              if (isApproved(r.type)) { ag.approvedLeave++; return; }
              if (isLate(r.type)) { ag.lates++; return; }
              ag.totalAbsences++;
              let m = r.date.substring(0, 7);
              ag.monthlyCounts[m] = (ag.monthlyCounts[m] || 0) + 1;
          });

          // Bradford Factor: S² × D, where S = spells (runs of consecutive
          // calendar days) and D = distinct absence days. Weights frequent
          // short absences far heavier than one long one. Calendar-day
          // adjacency is used (a Fri + Mon pair counts as two spells) since
          // per-agent working patterns aren't known here.
          let daySet = {};
          merged.forEach(r => { if (!isApproved(r.type) && !isLate(r.type)) daySet[r.date] = true; });
          let sortedDays = Object.keys(daySet).sort();
          let spells = 0, prevEpoch = null;
          sortedDays.forEach(ds => {
              let e = new Date(ds + 'T12:00:00').getTime();
              if (prevEpoch === null || (e - prevEpoch) > 86400000 * 1.5) spells++;
              prevEpoch = e;
          });
          ag.bradfordSpells = spells;
          ag.bradfordDays = sortedDays.length;
          ag.bradford = spells * spells * sortedDays.length;
      });
      
      let profiles = [];
      Object.values(agents).forEach(ag => {
          if (ag.totalAbsences > 0 || ag.lates > 0 || ag.approvedLeave > 0) {
              // FLAG Logic: Any month where count >= 4
              Object.keys(ag.monthlyCounts).forEach(month => {
                  if (ag.monthlyCounts[month] >= 4) {
                      ag.flags++;
                      ag.flaggedMonths.push(month);
                  }
              });
              ag.flaggedMonths.sort().reverse();
              // Sort timeline chronologically (newest first)
              ag.records.sort((a,b) => b.date.localeCompare(a.date));
              profiles.push(ag);
          }
      });
      
      profiles.sort((a, b) => b.totalAbsences - a.totalAbsences);
      return JSON.stringify({ year: year, profiles: profiles });
  },

  _executeDestructiveUpsert: function(sheetName, newRows, headersArray) {
    const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
    this._runDestructiveLogic(ssLocal, sheetName, newRows, headersArray);
  },

  _runDestructiveLogic: function(targetSpreadsheet, sheetName, newRows, headersArray) {
      let sheet = targetSpreadsheet.getSheetByName(sheetName);
      if (!sheet) sheet = targetSpreadsheet.insertSheet(sheetName);

      let existingData = [];
      if (sheet.getLastRow() > 0) existingData = sheet.getDataRange().getDisplayValues();
      if (existingData.length === 1 && existingData[0].join('') === "") existingData = [];
      const headers = existingData.length > 0 ? existingData.shift() : headersArray;

      let isIDP = sheetName === 'WF_IDP';
      let wipeKeys = new Set();
      newRows.forEach(r => {
          let dateStr = this._formatDate(r[isIDP ? 0 : 1]);
          if (isIDP) wipeKeys.add(dateStr);
          else wipeKeys.add(String(r[0]).trim().toLowerCase() + "_" + dateStr);
      });
      const retainedRows = existingData.filter(row => {
          if (!row[0]) return false;
          let rDate = this._formatDate(row[isIDP ? 0 : 1]);
          if (!rDate) return false; 
          let k = isIDP ? rDate : String(row[0]).trim().toLowerCase() + "_" + rDate;
          if (wipeKeys.has(k)) return false; 
          return true;
      });
      const combined = retainedRows.concat(newRows);
      // Crash-safe write: overwrite in place + clear the tail, never clearContents()
      // first. A timed-out import previously left the sheet EMPTY in the gap between
      // clearContents() and setValues() — which silently broke whichever tracker reads
      // it (e.g. the Absence Tracker). Header is rewritten in place; data overwrites;
      // any leftover rows below are blanked.
      const ncol = combined.length > 0 ? combined[0].length : headers.length;
      const prevLast = sheet.getLastRow();
      // Force time columns to PLAIN TEXT before writing so "9:00 AM" is never coerced
      // into a date/serial that getDisplayValues() would return as "12/30/1899".
      try {
        headers.forEach(function (h, i) {
          if (/time|start|end|interval/i.test(String(h)) && !/epoch/i.test(String(h))) {
            sheet.getRange(1, i + 1, Math.max(2, combined.length + 1), 1).setNumberFormat('@');
          }
        });
      } catch (e) {}
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      if (combined.length > 0) sheet.getRange(2, 1, combined.length, ncol).setValues(combined);
      const dlNewLast = combined.length + 1;
      if (prevLast > dlNewLast) sheet.getRange(dlNewLast + 1, 1, prevLast - dlNewLast, ncol).clearContent();
  },

  _getCycleForEpoch: function(epoch) {
      let d = new Date(epoch);
      let targetUTC = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0);
      if (d.getHours() >= 23) targetUTC += 86400000;
      let diffWeeks = Math.floor(Math.floor((targetUTC - Date.UTC(2026, 0, 29, 12, 0, 0)) / 86400000) / 7);
      return (Math.abs(diffWeeks) % 2 === 1) ? "WEEK B" : "WEEK A";
  },

  _calculateEpochBoundaries: function(mode, refDateStr) {
      let rY = parseInt(refDateStr.substring(0,4));
      let rM = parseInt(refDateStr.substring(5,7));
      let rD = parseInt(refDateStr.substring(8,10));
      let tObj = new Date(rY, rM-1, rD, 12, 0, 0, 0);
      let tStart = 0, tEnd = 0, label = "", cycle = "";
      if (mode === 'day') {
          tStart = new Date(rY, rM-1, rD, 0, 0, 0, 0).getTime();
          tEnd = new Date(rY, rM-1, rD, 23, 59, 59, 999).getTime();
          label = Utilities.formatDate(tObj, "America/Toronto", "yyyy-MM-dd");
          cycle = this._getCycleForEpoch(tStart);
      } 
      else if (mode === 'week') {
          let wStart = new Date(rY, rM-1, rD, 0, 0, 0, 0);
          let dayOfWeek = wStart.getDay(); 
          let offset = (dayOfWeek >= 3) ? (dayOfWeek - 3) : (dayOfWeek + 4);
          wStart.setDate(wStart.getDate() - offset);
          wStart.setHours(23, 0, 0, 0);
          
          let wEnd = new Date(wStart);
          wEnd.setDate(wStart.getDate() + 7);
          wEnd.setHours(22, 59, 59, 999);
          tStart = wStart.getTime();
          tEnd = wEnd.getTime();
          
          cycle = this._getCycleForEpoch(tStart);
          label = `${Utilities.formatDate(wStart, "America/Toronto", "MMM dd, HH:mm")} to ${Utilities.formatDate(wEnd, "America/Toronto", "MMM dd, HH:mm")}`;
      } 
      else if (mode === 'week_sun') {
          // Sunday → Saturday CALENDAR week (matches the Management View). No
          // Week A/B cycle — calendar weeks don't rotate; cycle filtering is
          // skipped in week modes anyway.
          let wStart = new Date(rY, rM-1, rD, 0, 0, 0, 0);
          wStart.setDate(wStart.getDate() - wStart.getDay()); // back to Sunday 00:00
          let wEnd = new Date(wStart);
          wEnd.setDate(wStart.getDate() + 6);
          wEnd.setHours(23, 59, 59, 999); // Saturday 23:59:59
          tStart = wStart.getTime();
          tEnd = wEnd.getTime();
          cycle = "WEEK";
          label = `${Utilities.formatDate(wStart, "America/Toronto", "MMM dd")} to ${Utilities.formatDate(wEnd, "America/Toronto", "MMM dd")}`;
      }
      else if (mode === 'ytd') {
          // Year-to-date: Jan 1 of the reference year through the reference day.
          let sDate = new Date(rY, 0, 1, 0, 0, 0, 0);
          tStart = sDate.getTime();
          let eDate = new Date(rY, rM - 1, rD, 23, 59, 59, 999);
          tEnd = eDate.getTime();
          label = `YTD ${rY}: Jan 01 to ${Utilities.formatDate(eDate, "America/Toronto", "MMM dd")}`;
          cycle = "YTD";
      }
      else if (mode === 'month' || mode === 'quarter') {
          let sMonth = (mode === 'month') ? rM - 1 : Math.floor((rM - 1) / 3) * 3;
          let eMonth = (mode === 'month') ? rM : sMonth + 3;

          let sDate = new Date(rY, sMonth, 1, 0, 0, 0, 0);
          tStart = sDate.getTime();
          let eDate = new Date(rY, eMonth, 0, 23, 59, 59, 999);
          tEnd = eDate.getTime();
          label = mode === 'month'
              ? `Month: ${Utilities.formatDate(sDate, "America/Toronto", "MMM dd")} to ${Utilities.formatDate(eDate, "America/Toronto", "MMM dd")}`
              : `Q${Math.floor((rM - 1) / 3) + 1}: ${Utilities.formatDate(sDate, "America/Toronto", "MMM dd")} to ${Utilities.formatDate(eDate, "America/Toronto", "MMM dd")}`;
          cycle = mode === 'month' ? "MONTH" : "QUARTER";
      }
      return { start: tStart, end: tEnd, label: label, cycle: cycle, startStr: Utilities.formatDate(new Date(tStart), "America/Toronto", "yyyy-MM-dd") };
  },

  getAnalytics: function(mode, refDate, trackerType, regionFilter = 'All', cycleFilter = 'ALL') {
    // Cache the whole payload per (type·mode·date·region·cycle), invalidated by
    // the WF_CACHE_VER stamp the import bumps. The common manager action —
    // flipping dates/modes — re-read the full sheet + recomputed the trend every
    // time; now repeat views are served from cache (gzip+base64, same scheme as
    // SafeTracker).
    var _ver = ''; try { _ver = PropertiesService.getScriptProperties().getProperty('WF_CACHE_VER') || ''; } catch (e) {}
    var _ck = 'wfAn|' + mode + '|' + refDate + '|' + trackerType + '|' + regionFilter + '|' + (cycleFilter || 'ALL') + '|' + _ver;
    var _cache = null; try { _cache = CacheService.getScriptCache(); } catch (e) {}
    if (_cache) { try { var _hit = _cache.get(_ck); if (_hit) return Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(_hit), 'application/x-gzip', 'c.gz')).getDataAsString(); } catch (e) {} }

    const dbIDP = this._getDB('WF_IDP');
    let dbSched;
    if (trackerType === 'coaching') dbSched = this._getDB('WF_COACHING');
    else if (trackerType === 'furlough') dbSched = this._getDB('WF_FURLOUGH');
    else if (trackerType === 'roles') dbSched = this._getDB('WF_ROLES');
    
    if (!dbIDP || !dbSched) return JSON.stringify({ error: "Databases missing. Please run WFM Import first." });
    const bounds = this._calculateEpochBoundaries(mode, refDate);
    let searchStart = new Date(bounds.start); searchStart.setDate(searchStart.getDate() - 1);
    const searchStartStr = Utilities.formatDate(searchStart, "America/Toronto", "yyyy-MM-dd");
    const endStr = Utilities.formatDate(new Date(bounds.end), "America/Toronto", "yyyy-MM-dd");

    const idpData = dbIDP.getLastRow() > 1 ? dbIDP.getDataRange().getDisplayValues().slice(1) : [];
    const schedData = dbSched.getLastRow() > 1 ? dbSched.getDataRange().getDisplayValues().slice(1) : [];

    let buckets = [];
    let combinedEvents = [];
    let groupedLogs = {}; 

    if (mode === 'day' && trackerType === 'furlough') {
      buckets = Array.from({length: 96}, (_, i) => ({ index: i, label: this._indexToTime(i), supply: 0, demand: 0, net: 0 }));
      idpData.forEach(row => {
        let rowDateStr = this._formatDate(row[0]);
        if (!rowDateStr) return;
        
        let rY = parseInt(rowDateStr.substring(0,4));
        let rM = parseInt(rowDateStr.substring(5,7));
        let rD = parseInt(rowDateStr.substring(8,10));
        if (isNaN(rY) || isNaN(rM) || isNaN(rD)) return; 
        
        let mins = this._timeToMins(row[1]);
        let blockTime = new Date(rY, rM-1, rD, Math.floor(mins/60), mins%60, 0, 0).getTime();
        
        if (blockTime >= bounds.start && blockTime <= bounds.end) { 
          let idx = this._timeToBucket(row[1]);
          if (idx > -1) { 
              let dem = parseFloat(String(row[2]).replace(',', '.')) || 0;
              let sup = parseFloat(String(row[3]).replace(',', '.')) || 0;
              buckets[idx].demand += dem; 
              buckets[idx].supply += sup; 
          }
        }
      });
    }

    let processedEvents = new Set();

    schedData.forEach(row => {
        let rowDateStr = this._formatDate(row[1]);
        if (!rowDateStr) return;
        let rowRegion = row[5] ? String(row[5]).trim() : 'Onshore';
        if (regionFilter !== 'All' && rowRegion !== regionFilter) return;

        if (rowDateStr >= searchStartStr && rowDateStr <= endStr) {
            let agent = String(row[0]).trim();
            let rY = parseInt(rowDateStr.substring(0,4));
            let rM = parseInt(rowDateStr.substring(5,7));
            let rD = parseInt(rowDateStr.substring(8,10));
            if (isNaN(rY) || isNaN(rM) || isNaN(rD)) return;

            let startMins = this._timeToMins(row[3]);
            let endMinsRaw = this._timeToMins(row[4]);
            let actSlice = String(row[2]).trim().substring(0, 10);
            let eventHash = `${agent}_${rowDateStr}_${startMins}_${endMinsRaw}_${actSlice}`;
            if (processedEvents.has(eventHash)) return; 
            processedEvents.add(eventHash);

            let endMins = endMinsRaw < startMins ? endMinsRaw + 1440 : endMinsRaw; 

            this._getShiftSplits(startMins, endMins).forEach(split => {
                // A shift split can still straddle midnight (e.g. a Night chunk
                // that runs 23:00 -> 00:30). Break it at each midnight so every
                // piece lands wholly inside one calendar day and is credited to
                // the day it actually occurs in -- not the day the chunk started.
                // Without this, the post-midnight minutes of a cross-day furlough
                // were attributed to the start day and went missing on the day
                // they really belonged to.
                let pieceStart = split.startMins;
                while (pieceStart < split.endMins) {
                    let nextMidnight = (Math.floor(pieceStart / 1440) + 1) * 1440;
                    let pieceEnd = Math.min(split.endMins, nextMidnight);
                    let piece = { shift: split.shift, startMins: pieceStart, endMins: pieceEnd, hours: (pieceEnd - pieceStart) / 60 };
                    pieceStart = pieceEnd;

                    let splitStartEpoch = new Date(rY, rM-1, rD, Math.floor(piece.startMins/60), piece.startMins%60, 0, 0).getTime();

                    if (splitStartEpoch >= bounds.start && splitStartEpoch <= bounds.end) {
                        if ((mode === 'month' || mode === 'quarter') && cycleFilter !== 'ALL') {
                            if (this._getCycleForEpoch(splitStartEpoch) !== cycleFilter) continue;
                        }

                        let effDateStr = Utilities.formatDate(new Date(splitStartEpoch), "America/Toronto", "yyyy-MM-dd");

                        if (trackerType === 'furlough' && mode === 'day') {
                            for (let min = piece.startMins; min < piece.endMins; min += 15) {
                                let blockTime = new Date(rY, rM-1, rD, Math.floor(min/60), min%60, 0, 0).getTime();
                                if (blockTime >= bounds.start && blockTime <= bounds.end) {
                                    let idx = Math.floor((min % 1440) / 15);
                                    if (idx >= 0 && idx < 96) buckets[idx].supply = Math.max(0, buckets[idx].supply - 1);
                                }
                            }
                        }

                        let rawStartStr = String(row[3]).trim();
                        let groupKey = `${agent}_${effDateStr}_${piece.shift}_${rawStartStr}`;

                        let actName = "Time Off";
                        if (trackerType === 'coaching' || trackerType === 'roles') actName = row[2];
                        groupKey += `_${actName}`;

                        if (!groupedLogs[groupKey]) {
                            groupedLogs[groupKey] = {
                                date: effDateStr, agent: agent,
                                activityName: actName,
                                shift: piece.shift, hours: piece.hours,
                                cycle: this._getCycleForEpoch(splitStartEpoch),
                                timeStart: this._minsToTime(piece.startMins),
                                timeEnd: this._minsToTime(piece.endMins)
                            };
                        } else {
                            groupedLogs[groupKey].hours += piece.hours;
                            groupedLogs[groupKey].timeEnd = this._minsToTime(piece.endMins);
                        }
                    }
                }
            });
        }
    });

    if (mode === 'day' && trackerType === 'furlough') buckets.forEach(b => { b.net = parseFloat((b.supply - b.demand).toFixed(2)); });
    combinedEvents = Object.values(groupedLogs).map(g => ({ date: g.date, agent: g.agent, activityName: g.activityName, shift: g.shift, cycle: g.cycle || '', hours: parseFloat(g.hours.toFixed(2)), time: `${g.timeStart} - ${g.timeEnd}` }));
    let totals = { all: 0, morning: 0, evening: 0, night: 0, count: combinedEvents.length };
    combinedEvents.forEach(f => { totals.all += f.hours; if (f.shift === 'Morning') totals.morning += f.hours; else if (f.shift === 'Evening') totals.evening += f.hours; else totals.night += f.hours; });

    // Month-over-month trend (whole DB, region-aware, cycle-aware) so the UI can
    // show whether furlough is trending up or down and compare WEEK A vs WEEK B.
    // Computed from the already-loaded schedData — no extra sheet read.
    let trend = (mode === 'month' || mode === 'quarter' || mode === 'ytd') ? this._computeTrendMonthly(schedData, regionFilter) : [];
    // Which month keys the currently-selected period spans (used to highlight
    // the trend and to sum the A-vs-B split for this period).
    let periodKeys = [];
    let pCur = new Date(bounds.start);
    let pEndKey = Utilities.formatDate(new Date(bounds.end), "America/Toronto", "yyyy-MM");
    for (let guard = 0; guard < 24; guard++) {
        let k = Utilities.formatDate(pCur, "America/Toronto", "yyyy-MM");
        if (periodKeys.indexOf(k) === -1) periodKeys.push(k);
        if (k === pEndKey) break;
        pCur.setDate(1); pCur.setMonth(pCur.getMonth() + 1);
    }

    var __payload = JSON.stringify({ mode: mode, trackerType: trackerType, label: bounds.label, cycle: bounds.cycle, grid: buckets, events: combinedEvents, totals: totals, trend: trend, periodKeys: periodKeys });
    if (_cache) { try { var _z = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(__payload)).getBytes()); if (_z.length < 99000) _cache.put(_ck, _z, 21600); } catch (e) {} }
    return __payload;
  },

  // Monthly hours across the full schedule DB, split by WEEK A / WEEK B cycle.
  // Mirrors the day-view attribution: shift-split, then split again at midnight
  // so every slice is credited to the month (and cycle) it actually occurs in.
  _computeTrendMonthly: function(schedData, regionFilter) {
      const MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      let months = {};
      let processed = new Set();
      (schedData || []).forEach(row => {
          let dStr = this._formatDate(row[1]);
          if (!dStr) return;
          let reg = row[5] ? String(row[5]).trim() : 'Onshore';
          if (regionFilter && regionFilter !== 'All' && reg !== regionFilter) return;
          let rY = parseInt(dStr.substring(0,4)), rM = parseInt(dStr.substring(5,7)), rD = parseInt(dStr.substring(8,10));
          if (isNaN(rY) || isNaN(rM) || isNaN(rD)) return;
          let agent = String(row[0]).trim();
          let sM = this._timeToMins(row[3]), eMr = this._timeToMins(row[4]);
          let actSlice = String(row[2]).trim().substring(0, 10);
          let hash = `${agent}_${dStr}_${sM}_${eMr}_${actSlice}`;
          if (processed.has(hash)) return; processed.add(hash);
          let eM = eMr < sM ? eMr + 1440 : eMr;
          this._getShiftSplits(sM, eM).forEach(split => {
              let ps = split.startMins;
              while (ps < split.endMins) {
                  let nm = (Math.floor(ps / 1440) + 1) * 1440;
                  let pe = Math.min(split.endMins, nm);
                  let hrs = (pe - ps) / 60;
                  let epoch = new Date(rY, rM-1, rD, Math.floor(ps/60), ps%60, 0, 0).getTime();
                  let mKey = Utilities.formatDate(new Date(epoch), "America/Toronto", "yyyy-MM");
                  let cyc = this._getCycleForEpoch(epoch);
                  if (!months[mKey]) months[mKey] = { hours: 0, weekA: 0, weekB: 0 };
                  months[mKey].hours += hrs;
                  if (cyc === 'WEEK A') months[mKey].weekA += hrs; else months[mKey].weekB += hrs;
                  ps = pe;
              }
          });
      });
      return Object.keys(months).sort().map(k => {
          let mm = parseInt(k.substring(5,7));
          return {
              key: k,
              label: MN[mm-1] + " '" + k.substring(2,4),
              hours: parseFloat(months[k].hours.toFixed(2)),
              weekA: parseFloat(months[k].weekA.toFixed(2)),
              weekB: parseFloat(months[k].weekB.toFixed(2))
          };
      });
  },

  getUnifiedReport: function(refDateStr, cycleFilter) {
        const targetCycle = cycleFilter || 'ALL';
        const bounds = this._calculateEpochBoundaries("month", refDateStr);
        let report = { cycle: targetCycle === 'ALL' ? "FULL MONTH" : targetCycle, period: bounds.label };
        
        const nowEpoch = new Date().getTime();

        let mlData = {};
        const dbML = this._getDB('WF_MASTERLIST');
        if (dbML && dbML.getLastRow() > 1) {
            dbML.getDataRange().getDisplayValues().slice(1).forEach(r => {
                let displayName = String(r[0]).trim();
                let key = _normalizeAgentKey(displayName);
                let isOffshore = String(r[4]).toUpperCase().includes("TI") || String(r[5]).includes("@") || String(r[4]).toUpperCase().includes("EL SALVADOR") || String(r[4]).toUpperCase().includes("GUATEMALA");
                // Registry override wins if user manually flipped this agent.
                if (typeof RegionRegistry !== 'undefined') {
                    const rg = RegionRegistry.getRegion(displayName);
                    const src = RegionRegistry.getSource(displayName);
                    if (rg && src === 'manual') isOffshore = rg === 'Offshore';
                }
                mlData[key] = {
                    name: displayName,
                    level: r[1],
                    manager: r[2],
                    skills: r[3],
                    isOffshore: isOffshore,
                    isBackup: String(r[3]).includes("Backup MRC")
                };
            });
        }

        // Internal registry keyed by normalized key so accented / hyphenated names
        // (e.g. "Jaén-Benitez" in MasterList vs "Jaen Benitez" in WFM/GEM) collapse
        // to a single agent.
        let agentsByKey = {};
        const getAg = (name, reg) => {
            let key = _normalizeAgentKey(name);
            if (!agentsByKey[key]) {
                // Prefer MasterList's display name when we have one.
                let displayName = (mlData[key] && mlData[key].name) ? mlData[key].name : String(name).trim();
                agentsByKey[key] = { name: displayName, region: reg || 'Onshore', acsu: 0, coach: 0, safe: 0, tower: 0, total: 0, ot: 0 };
            }
            if (reg && String(reg).includes('Offshore')) agentsByKey[key].region = 'Offshore';
            // Registry is authoritative. Overrides both the row-level region and any MasterList inference.
            if (typeof RegionRegistry !== 'undefined') {
                const registryRegion = RegionRegistry.getRegion(name);
                if (registryRegion) agentsByKey[key].region = registryRegion;
            }
            return agentsByKey[key];
        };
        Object.keys(mlData).forEach(k => {
            let ml = mlData[k];
            if (ml.isBackup) getAg(ml.name, ml.isOffshore ? "Offshore" : "Onshore");
        });
        let searchStart = new Date(bounds.start); searchStart.setDate(searchStart.getDate() - 1);
        const sStr = Utilities.formatDate(searchStart, "America/Toronto", "yyyy-MM-dd");
        const eStr = Utilities.formatDate(new Date(bounds.end), "America/Toronto", "yyyy-MM-dd");
        
        const parseDB = (sheetName, metricName) => {
            const db = this._getDB(sheetName);
            if (!db || db.getLastRow() < 2) return;
            
            let processedEvents = new Set();
            db.getDataRange().getDisplayValues().slice(1).forEach(row => {
                let dStr = this._formatDate(row[1]);
                if (dStr >= sStr && dStr <= eStr) {
                    let rY = parseInt(dStr.substring(0,4)); let rM = parseInt(dStr.substring(5,7)); let rD = parseInt(dStr.substring(8,10));
                    if(isNaN(rY) || isNaN(rM) || isNaN(rD)) return;
                    
                    let agent = String(row[0]).trim();
                    let sMins = this._timeToMins(row[3]); let eMinsR = this._timeToMins(row[4]);
                    let region = row[5] ? String(row[5]).trim() : 'Onshore';
  
                    let actSlice = String(row[2]).trim().substring(0, 10);
                    let eventHash = `${agent}_${dStr}_${sMins}_${eMinsR}_${actSlice}`;
                    if (processedEvents.has(eventHash)) return; processedEvents.add(eventHash);

                    let eMins = eMinsR < sMins ? eMinsR + 1440 : eMinsR; 
                    this._getShiftSplits(sMins, eMins).forEach(s => {
                        let epoch = new Date(rY, rM-1, rD, Math.floor(s.startMins/60), s.startMins%60, 0, 0).getTime();
                        if (epoch >= bounds.start && epoch <= bounds.end && epoch <= nowEpoch) { 
                            if (targetCycle !== 'ALL' && this._getCycleForEpoch(epoch) !== targetCycle) return;
                            getAg(agent, region)[metricName] += s.hours; 
                            getAg(agent, region).total += s.hours;
                        }
                    });
                }
            });
        };

        parseDB('WF_COACHING', 'coach');
        parseDB('WF_FURLOUGH', 'acsu');
        const dbRoles = this._getDB('WF_ROLES');
        if (dbRoles && dbRoles.getLastRow() > 1) {
            let processedRoles = new Set();
            dbRoles.getDataRange().getDisplayValues().slice(1).forEach(row => {
                let dStr = this._formatDate(row[1]);
                if (dStr >= sStr && dStr <= eStr) {
                    let rY = parseInt(dStr.substring(0,4)); let rM = parseInt(dStr.substring(5,7)); let rD = parseInt(dStr.substring(8,10));
                    if(isNaN(rY) || isNaN(rM) || isNaN(rD)) return;
                    
                    let agent = String(row[0]).trim();
                    let sMins = this._timeToMins(row[3]); let eMinsR = this._timeToMins(row[4]);
                    let region = row[5] ? String(row[5]).trim() : 'Onshore';
  
                    let actSlice = String(row[2]).trim().substring(0, 10);
                    let eventHash = `${agent}_${dStr}_${sMins}_${eMinsR}_${actSlice}`;
                    if (processedRoles.has(eventHash)) return; processedRoles.add(eventHash);

                    let eMins = eMinsR < sMins ? eMinsR + 1440 : eMinsR; 
                    let roleType = String(row[2]).toUpperCase();
                    this._getShiftSplits(sMins, eMins).forEach(s => {
                        let epoch = new Date(rY, rM-1, rD, Math.floor(s.startMins/60), s.startMins%60, 0, 0).getTime();
                        if (epoch >= bounds.start && epoch <= bounds.end && epoch <= nowEpoch) { 
                            if (targetCycle !== 'ALL' && this._getCycleForEpoch(epoch) !== targetCycle) return;

                            if (roleType.includes('SAFE')) { getAg(agent, region).safe += s.hours; getAg(agent, region).total += s.hours; }
                            else if (roleType.includes('TOWER') || roleType.includes('WOFQT') || roleType.includes('WOQFT')) { getAg(agent, region).tower += s.hours; getAg(agent, region).total += s.hours; }
                        }
                     });
                }
            });
        }

        // DB_Sessions previously fed ICL / ULC FIRE session hours — no longer
        // tracked (those roles are live-Floor only), so the block was removed.

        // Overtime (Overtime_Tracking): worked OT hours, kept SEPARATE from `total`
        // (total = off-phone hours; OT is extra time on the phones). Columns:
        // [0]Timestamp [1]Agent [2]OT Start [3]OT End [4]Break Start [5]Break End [6]DateStr.
        const dbOT = this._getDB('Overtime_Tracking');
        if (dbOT && dbOT.getLastRow() > 1) {
            let processedOT = new Set();
            dbOT.getDataRange().getDisplayValues().slice(1).forEach(row => {
                let agent = String(row[1]).trim();
                if (!agent) return;
                let dStr = this._formatDate(row[6]) || this._formatDate(row[0]);
                if (!dStr || dStr < sStr || dStr > eStr) return;
                let rY = parseInt(dStr.substring(0,4)), rM = parseInt(dStr.substring(5,7)), rD = parseInt(dStr.substring(8,10));
                if (isNaN(rY) || isNaN(rM) || isNaN(rD)) return;

                let sMins = this._timeToMins(row[2]);
                let eMinsR = this._timeToMins(row[3]);
                let hash = `${agent}_${dStr}_${sMins}_${eMinsR}`;
                if (processedOT.has(hash)) return; processedOT.add(hash);

                let eMins = eMinsR < sMins ? eMinsR + 1440 : eMinsR;
                let gross = eMins - sMins;
                // Subtract the OT break, if one was logged.
                let bs = this._timeToMins(row[4]), be = this._timeToMins(row[5]);
                let brk = 0;
                if (bs !== be) { let beA = be < bs ? be + 1440 : be; brk = beA - bs; }
                let netHours = (gross - brk) / 60;
                if (netHours <= 0) return;

                let epoch = new Date(rY, rM-1, rD, Math.floor(sMins/60), sMins%60, 0, 0).getTime();
                if (epoch >= bounds.start && epoch <= bounds.end && epoch <= nowEpoch) {
                    if (targetCycle !== 'ALL' && this._getCycleForEpoch(epoch) !== targetCycle) return;
                    getAg(agent).ot += netHours; // intentionally NOT added to .total
                }
            });
        }

        const dbGEM = this._getDB('WF_GEM_DATA_V3');
        let gemData = {};
        if (dbGEM && dbGEM.getLastRow() > 1) {
            const targetMonth = refDateStr.substring(0, 7);
            const fixTime = (v) => {
                if (!v || v === '-' || v === '00:00') return '-';
                let s = String(v).trim();
                if (s.startsWith("'")) s = s.substring(1); 
                let m = s.match(/T(\d{2}:\d{2})/); 
                if (m) return m[1];
                if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "HH:mm");
                return s;
            };
            dbGEM.getDataRange().getValues().slice(1).forEach(row => {
                let rowDateStr = row[0] instanceof Date ? Utilities.formatDate(row[0], "America/Toronto", "yyyy-MM") : String(row[0]).substring(0, 7);
                if (rowDateStr === targetMonth) {
                    let agName = _titleCaseName(String(row[1]).trim());
                    gemData[agName] = {
                        cph: row[2], inPct: row[3], outPct: row[4], aht: fixTime(row[5]),
                        transfers: row[10] || 0, transfPct: row[11] || 0,
                        avgTalk: fixTime(row[12]), avgHold: fixTime(row[13]), avgAcw: fixTime(row[14]),
                        outCalls: row[15] || 0, outTalk: fixTime(row[16]),
                        extCalls: row[17] || 0, extTalk: fixTime(row[18])
                    };
                }
            });
        }

        let gemKeys = Object.keys(gemData);
        // Index GEM data by normalized key so "Jaen Benitez" matches "Jaén-Benitez".
        let gemByKey = {};
        gemKeys.forEach(g => { gemByKey[_normalizeAgentKey(g)] = gemData[g]; });

        let finalArr = Object.values(agentsByKey);
        // SAFE hours are now sourced from the imported SAFE report (ACTIVE TIME),
        // not the schedule codes. Overlay by normalized name; agents absent from
        // the report read 0. `total` is adjusted by the swap so off-phone math stays consistent.
        var _safeStartStr = Utilities.formatDate(new Date(bounds.start), "America/Toronto", "yyyy-MM-dd");
        var _safeRpt = (typeof ReportImport !== 'undefined' && ReportImport.getSafeForPeriod) ? ReportImport.getSafeForPeriod(_safeStartStr, eStr) : { has: false };
        var _tidMap = this._agentTidMap();   // normalized name → WFM Employee ID
        finalArr.forEach(a => {
            let aKey = _normalizeAgentKey(a.name);
            if (_safeRpt.has) {
                // Match on Employee ID first (exact), then fall back to name.
                var _tid = _tidMap[aKey];
                var _m = (_tid && _safeRpt.byTid && _safeRpt.byTid[_tid] != null)
                    ? _safeRpt.byTid[_tid]
                    : ReportImport.matchSafeHours(a.name, _safeRpt);
                var rptSafe = (_m != null) ? _m : 0;
                a.total = (a.total || 0) - (a.safe || 0) + rptSafe;
                a.safe = rptSafe;
                a.safeFromReport = true;
            }
            let ml = mlData[aKey];
            if (ml) {
                a.level = ml.level;
                a.manager = ml.manager;
                a.skills = ml.skills;
                a.isBackupMRC = ml.isBackup;
                if (ml.isOffshore) a.region = "Offshore";
            }

            ['acsu','coach','safe','tower','total','ot'].forEach(k => a[k] = parseFloat(a[k].toFixed(2)));
            let matchedGem = gemData[a.name] || gemByKey[aKey];
            if (!matchedGem) {
                let aWords = aKey.replace(/,/g, ' ').split(' ').filter(x => x.length > 1);
                for(let i=0; i<gemKeys.length; i++) {
                    let gName = gemKeys[i];
                    let gKey = _normalizeAgentKey(gName);
                    let gWords = gKey.replace(/,/g, ' ').split(' ').filter(x => x.length > 1);

                    if (aWords.length > 1 && gWords.length > 1 && aWords[0] === gWords[0] && aWords[1] === gWords[1]) {
                        matchedGem = gemData[gName];
                        break;
                    }
                    // Substring fallback: only if both names are long enough to make a spurious
                    // short-substring match (e.g. "Ali" inside "Alison") unlikely.
                    let minLen = Math.min(aKey.length, gKey.length);
                    if (minLen >= 8 && (aKey.includes(gKey) || gKey.includes(aKey))) {
                        matchedGem = gemData[gName];
                        break;
                    }
                 }
            }

            if (matchedGem) {
                a.gem = matchedGem;
                a.cph = parseFloat(matchedGem.cph).toFixed(2);
                a.aht = matchedGem.aht;
                a.inOut = Math.round(parseFloat(matchedGem.inPct) * 100) + "% / " + Math.round(parseFloat(matchedGem.outPct) * 100) + "%";
            } else {
                a.gem = null;
                a.cph = '-'; a.aht = '-'; a.inOut = '-';
            }
        });
        finalArr = finalArr.filter(a => a.total > 0 || a.gem !== null || a.isBackupMRC === true);
        report.data = finalArr.sort((a,b) => b.total - a.total);
        return JSON.stringify(report);
  },

  // Drains the queue of months that importData deferred, archiving each (ALL/
  // WEEK A/WEEK B) within a time budget so it never hits the 6-min limit. The
  // client calls this after an import and loops while `remaining` > 0.
  flushPendingArchives: function() {
      var props = PropertiesService.getScriptProperties();
      var pend = {}; try { pend = JSON.parse(props.getProperty('PENDING_ARCHIVE') || '{}'); } catch (e) {}
      var months = Object.keys(pend);
      if (!months.length) return JSON.stringify({ done: true, archived: [], remaining: 0 });
      var done = [], t0 = Date.now();
      for (var i = 0; i < months.length; i++) {
          if (Date.now() - t0 > 90000) break;   // ~90s per call — stay gentle so it never hogs the server
          var m = months[i];
          try {
              this.archiveUnifiedReport(m, 'ALL');
              this.archiveUnifiedReport(m, 'WEEK A');
              this.archiveUnifiedReport(m, 'WEEK B');
              done.push(m.substring(0, 7)); delete pend[m];
          } catch (e) {}
      }
      props.setProperty('PENDING_ARCHIVE', JSON.stringify(pend));
      var remaining = Object.keys(pend).length;
      return JSON.stringify({ done: remaining === 0, archived: done, remaining: remaining });
  },

  archiveUnifiedReport: function(refDateStr, cycleFilter) {
      const reportStr = this.getUnifiedReport(refDateStr, cycleFilter);
      const report = JSON.parse(reportStr);
      const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
      
      const writeToArchiveSheet = (spreadsheet) => {
          let archiveSheet = spreadsheet.getSheetByName('Weekly_Archives_V3');
          if (!archiveSheet) { 
              archiveSheet = spreadsheet.insertSheet('Weekly_Archives_V3');
              archiveSheet.appendRow(["Timestamp", "Cycle", "Period", "Agent Name", "Region", "Total Off-Phone", "JSON Payload"]); 
              archiveSheet.getRange(1,1,1,7).setFontWeight("bold");
          }
          const data = archiveSheet.getDataRange().getValues();
          for (let i = data.length - 1; i >= 1; i--) { if (data[i][2] === report.period && data[i][1] === report.cycle) archiveSheet.deleteRow(i + 1);
          }
          
          const rows = report.data.map(a => [ new Date(), report.cycle, report.period, a.name, a.region, a.total, JSON.stringify(a) ]);
          if (rows.length > 0) archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
      };

      writeToArchiveSheet(ssLocal);
      return `Successfully archived ${report.data.length} agents for ${report.cycle} (${report.period}).`;
  },

  /**
   * Real-time staffing balance for the current 15-min IDP bucket.
   * Positive net = surplus; negative = understaffed. Based on:
   *   supply  = IDP "open" seats for this 15-min slot (today)
   *   demand  = IDP "required" seats for same slot
   * If IDP data isn't available for today, returns null so the UI can hide.
   */
  getStaffingBalance: function() {
    try {
      const dbIDP = this._getDB('WF_IDP');
      if (!dbIDP || dbIDP.getLastRow() < 2) return JSON.stringify({ available: false, reason: 'No IDP data' });

      const now = new Date();
      const tz = 'America/Toronto';
      const todayStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
      const mins = parseInt(Utilities.formatDate(now, tz, 'H'), 10) * 60 + parseInt(Utilities.formatDate(now, tz, 'm'), 10);
      const bucketMin = Math.floor(mins / 15) * 15;
      const bucketLabel = (bucketMin < 600 ? '0' : '') + Math.floor(bucketMin / 60) + ':' + (bucketMin % 60 < 10 ? '0' : '') + (bucketMin % 60);

      const data = dbIDP.getRange(2, 1, dbIDP.getLastRow() - 1, 4).getDisplayValues();
      let demand = 0, supply = 0, found = false;
      for (let i = 0; i < data.length; i++) {
        const rowDate = this._formatDate(data[i][0]);
        if (rowDate !== todayStr) continue;
        const rowTime = this._formatTimeStr(String(data[i][1]));
        if (rowTime !== bucketLabel) continue;
        demand = parseFloat(String(data[i][2]).replace(',', '.')) || 0;
        supply = parseFloat(String(data[i][3]).replace(',', '.')) || 0;
        found = true;
        break;
      }
      if (!found) return JSON.stringify({ available: false, reason: 'No IDP row for ' + todayStr + ' ' + bucketLabel });

      const net = parseFloat((supply - demand).toFixed(2));
      let status = 'ok';
      if (net < -2) status = 'critical';
      else if (net < 0) status = 'warn';
      else if (net > 3) status = 'surplus';

      return JSON.stringify({
        available: true,
        bucket: bucketLabel,
        date: todayStr,
        demand: demand,
        supply: supply,
        net: net,
        status: status
      });
    } catch (e) {
      return JSON.stringify({ available: false, reason: 'Error: ' + e.message });
    }
  },

  getArchiveList: function() {
      let sheet = this._getDB('Weekly_Archives_V3');
      if (!sheet || sheet.getLastRow() < 2) return "[]";
      
      const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getDisplayValues();
      const unique = [];
      const seen = new Set();
      for (let i = data.length - 1; i >= 0; i--) {
          const cycle = data[i][0];
          const period = data[i][1];
          let key = cycle + "|" + period;
          if (!seen.has(key)) { seen.add(key);
          unique.push({ cycle: cycle, period: period }); }
      }
      return JSON.stringify(unique);
  },

  getArchivedReport: function(targetPeriod, targetCycle) {
      let sheet = this._getDB('Weekly_Archives_V3');
      if (!sheet || sheet.getLastRow() < 2) return "{}";

      // Map the UI cycle filter to the label stored in the archive.
      let cycleLabel = null;
      if (targetCycle === 'WEEK A' || targetCycle === 'WEEK B') cycleLabel = targetCycle;
      else if (!targetCycle || targetCycle === 'ALL') cycleLabel = 'FULL MONTH';
      else cycleLabel = targetCycle;

      const data = sheet.getDataRange().getDisplayValues();
      let report = { cycle: cycleLabel, period: targetPeriod, data: [] };
      for (let i = 1; i < data.length; i++) {
          if (data[i][2] === targetPeriod && data[i][1] === cycleLabel) {
              try { report.data.push(JSON.parse(data[i][6])); } catch(e) {}
          }
      }
      report.data.sort((a,b) => b.total - a.total);
      return JSON.stringify(report);
  },

  _getShiftSplits: function(startMins, endMins) {
      let splits = [];
      let current = startMins;
      while (current < endMins) {
          let timeOfDay = current % 1440;
          let shiftType = "Night"; let nextBound = current - timeOfDay + 420;
          if (timeOfDay >= 420 && timeOfDay < 900) { shiftType = "Morning"; nextBound = current - timeOfDay + 900;
          } 
          else if (timeOfDay >= 900 && timeOfDay < 1380) { shiftType = "Evening";
          nextBound = current - timeOfDay + 1380; } 
          else if (timeOfDay >= 1380) nextBound = current - timeOfDay + 1860;
          let chunkEnd = Math.min(endMins, nextBound);
          splits.push({ shift: shiftType, hours: (chunkEnd - current) / 60, startMins: current, endMins: chunkEnd });
          current = chunkEnd;
      }
      return splits;
  },
  
  _minsToTime: function(mins) { let m = mins % 1440; let h = Math.floor(m / 60);
      let mm = m % 60; return `${h < 10 ? '0'+h : h}:${mm < 10 ? '0'+mm : mm}`;
  },
  _timeToMins: function(tStr) {
       if (tStr == null || tStr === "") return 0;
       // Date objects: format in the sheet TZ so getHours() can't drift.
       let s = (tStr instanceof Date) ? Utilities.formatDate(tStr, "America/Toronto", 'HH:mm') : String(tStr).trim();
       if (!s) return 0;
       let num = Number(s);
       if (!isNaN(num) && s.indexOf(':') === -1 && !/h/i.test(s)) {
           // day fraction (0.375 = 9:00). Whole numbers are date serials, NOT a time-of-day → 0.
           return (num > 0 && num < 1.5) ? Math.round(num * 1440) % 1440 : 0;
       }
       // French-Canadian 24h clock that a fr-CA sheet returns: "13 h 30", "9 h", "11h00".
       let fr = s.match(/^(\d{1,2})\s*h\s*(\d{2})?$/i);
       if (fr) { let fh = parseInt(fr[1], 10), fm = fr[2] ? parseInt(fr[2], 10) : 0; return (fh > 23 || fm > 59) ? 0 : fh * 60 + fm; }
       let match = s.match(/(\d{1,2})[:\.](\d{2})/);   // HH:MM, ignoring any trailing :SS
       if (!match) return 0;
       let h = parseInt(match[1], 10), m = parseInt(match[2], 10);
       let amp = /p\.?\s*m/i.test(s) ? 'PM' : (/a\.?\s*m/i.test(s) ? 'AM' : null);   // meridian anywhere in the string
       if (amp === 'PM' && h < 12) h += 12;
       if (amp === 'AM' && h === 12) h = 0;
       if (h > 23 || m > 59) return 0;
       return (h * 60) + m;
  },
  
  _parseCSVLine: function(text) {
    if (text.includes('\t')) return text.split('\t').map(s => s.trim());
    let ret = [], inQuote = false, token = "";
    for(let i=0; i<text.length; i++) {
      let char = text[i];
      if(char === '"') { inQuote = !inQuote; continue; }
      if(char === ',' && !inQuote) { ret.push(token.trim()); token = ""; } else token += char;
    }
    ret.push(token.trim()); return ret;
  },
  
  _parseDate: function(s) { return this._formatDate(s); },
  _formatDate: function(d) { 
      if (d == null || d === "") return "";
      if (d instanceof Date) return Utilities.formatDate(d, "America/Toronto", "yyyy-MM-dd");
      let s = String(d).trim();
      let num = Number(s);
      if (!isNaN(num) && num > 30000) {
          let date = new Date((num - 25569) * 86400 * 1000);
          date.setMinutes(date.getMinutes() + date.getTimezoneOffset());
          return Utilities.formatDate(date, "America/Toronto", "yyyy-MM-dd");
      }
      let isoMatch = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
      if (isoMatch) return `${isoMatch[1]}-${isoMatch[2].padStart(2,'0')}-${isoMatch[3].padStart(2,'0')}`;
      let regMatch = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
    
      if (regMatch) { let 
      p1 = parseInt(regMatch[1]), p2 = parseInt(regMatch[2]);
      let m = p1 > 12 ? p2 : p1;
      let day = p1 > 12 ? p1 : p2;
      return `${regMatch[3]}-${String(m).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
      }
      let pDate = new Date(s);
      if (!isNaN(pDate)) return Utilities.formatDate(pDate, "America/Toronto", "yyyy-MM-dd");
      return s.substring(0, 10);
  },
  
  _formatTimeStr: function(t) { let d=new Date(`2000/01/01 ${t}`); return isNaN(d)?t:Utilities.formatDate(d, "America/Toronto", 'HH:mm');
  },
  _cleanActivity: function(s) { return s.replace(/\d{2}\s?[AP]M/gi, '').trim(); },
  _timeToBucket: function(val) { let mins = this._timeToMins(val);
  return mins < 0 ? -1 : Math.floor(mins / 15); },
  _indexToTime: function(i) { let h=Math.floor(i/4), m=(i%4)*15; return `${h<10?'0'+h:h}:${m===0?'00':m}`;
  }
};

/**
 * One-time maintenance: purge mis-classified "TI Mentor" rows from
 * WF_COACHING (they were ingested before the exclusion existed).
 * Run once from the Apps Script editor; future imports exclude them
 * automatically. Safe to run repeatedly.
 */
function cleanupTiMentorCoaching() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('WF_COACHING');
  if (!sheet || sheet.getLastRow() < 2) return 'WF_COACHING empty — nothing to clean.';
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  const keep = data.filter(r => !COACH_EXCLUDE_RGX.test(String(r[2] || '')));
  const removed = data.length - keep.length;
  if (!removed) return 'No TI Mentor rows found in WF_COACHING.';
  sheet.getRange(2, 1, data.length, sheet.getLastColumn()).clearContent();
  if (keep.length) sheet.getRange(2, 1, keep.length, keep[0].length).setValues(keep);
  return removed + ' TI Mentor row(s) removed from WF_COACHING (' + keep.length + ' kept).';
}

function fetchSyncMetadata() {
  try {
    const props = PropertiesService.getDocumentProperties();
    return JSON.stringify({
      sched: props.getProperty('SYNC_SCHED') || "Awaiting WFM Sync...",
      idp: props.getProperty('SYNC_IDP') || "Awaiting IDP Sync...",
      muSet: props.getProperty('MU_SET') || ""
    });
  } catch(e) {
    return JSON.stringify({ sched: "Metadata unavailable", idp: "Metadata unavailable", muSet: "" });
  }
}

function getSmartUnifiedReport(dateStr, cycleFilter) {
    const reqParts = dateStr.split('-');
    const reqY = parseInt(reqParts[0]);
    const reqM = parseInt(reqParts[1]) - 1;
    const cycle = cycleFilter || 'ALL';

    const today = new Date();
    const isCurrentMonth = (reqY === today.getFullYear() && reqM === today.getMonth());

    if (isCurrentMonth) {
        return WorkforceTracker.getUnifiedReport(dateStr, cycle);
    } else {
        const bounds = WorkforceTracker._calculateEpochBoundaries('month', dateStr);
        const archivedData = WorkforceTracker.getArchivedReport(bounds.label, cycle);

        let parsed = JSON.parse(archivedData);
        if (parsed.data && parsed.data.length > 0) {
            return archivedData;
        } else {
            let liveCalc = JSON.parse(WorkforceTracker.getUnifiedReport(dateStr, cycle));
            liveCalc.cycle = "HISTORICAL RECORD";
            return JSON.stringify(liveCalc);
        }
    }
}

function runRetroactiveArchive() {
   const db = WorkforceTracker._getDB('WF_GEM_DATA_V3');
   if (!db || db.getLastRow() < 2) return "No GEM data found to archive.";

   let data = db.getRange(2, 1, db.getLastRow()-1, 1).getDisplayValues().flat();
   let months = new Set();
   data.forEach(d => {
       if (d && d.length >= 7) months.add(d.substring(0,7) + "-01");
   });

   let todayStr = Utilities.formatDate(new Date(), "America/Toronto", "yyyy-MM");
   let archivedList = [];

   months.forEach(mStr => {
       if (!mStr.startsWith(todayStr)) {
           WorkforceTracker.archiveUnifiedReport(mStr, 'ALL');
           WorkforceTracker.archiveUnifiedReport(mStr, 'WEEK A');
           WorkforceTracker.archiveUnifiedReport(mStr, 'WEEK B');
           archivedList.push(mStr.substring(0,7));
       }
   });

   return `Successfully retro-archived: ${archivedList.join(', ')}`;
}

function fetchAbsenceProfiles(year) {
   return WorkforceTracker.getAbsenceProfiles(year);
}
