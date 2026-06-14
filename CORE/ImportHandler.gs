/**
 * MODULE: IMPORT HANDLER (SMART WFM - LOCAL HOST ONLY)
 */

const ImportHandler = {
  run: function(text, date) { return processWFMImport(text, date); }
};

function processWFMImport(rawText, forcedDate) {
  if (!rawText) return "Error: No text provided.";
  Logger.log('[wfm] start, text len=' + rawText.length);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone("America/Toronto");
  let sheet = ss.getSheetByName("Raw Schedule");
  if (!sheet) { sheet = ss.insertSheet("Raw Schedule"); }

  // Batch RegionRegistry writes — pushAgent calls upsert() once per agent,
  // and without batching each call costs 2 sheet API round-trips. With
  // 100+ agents that's the difference between 30-150s of overhead and ~0.
  if (typeof RegionRegistry !== 'undefined') RegionRegistry.beginBatch();

  try {
    // Reset by clearing CONTENT (fast, no row reflow) rather than deleteRows,
    // which physically removes/shifts every row and on a 2-month re-paste
    // (thousands of rows) costs tens of seconds. clearContent blanks the old
    // data in one batch; the new rows overwrite from row 2.
    var t0 = new Date().getTime();
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 12).clearContent();
    var hdrRange = sheet.getRange(1, 1, 1, 12);
    hdrRange.setValues([["Agent Name", "ID", "DateStr", "Shift Start", "Shift End", "Shift Type", "Region", "Breaks JSON", "Role", "AbsentType", "StartEpoch", "EndEpoch"]]);
    hdrRange.setFontWeight("bold").setBackground("#e0e0e0");
    Logger.log('[wfm] sheet reset in ' + (new Date().getTime() - t0) + 'ms (was ' + lastRow + ' rows)');

  const lines = rawText.split(/\r?\n/);
  Logger.log('[wfm] parsing ' + lines.length + ' lines');
  let rosterData = [];
  let currentAgent = null;
  let currentID = null;
  let buffer = resetBuffer();

  if (!forcedDate) forcedDate = Utilities.formatDate(new Date(), "America/Toronto", "yyyy-MM-dd");
  const [defY, defM, defD] = forcedDate.split('-').map(Number);

  const rgxAgent = /Agent:\s*(\d+)\s*(.*)/i;
  const rgxAnyDateLine = /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/;
  const rgxShiftLine = /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})\s+(\d{1,2}:\d{2}\s*[AP]M)\s+(\d{1,2}:\d{2}\s*[AP]M)/i;
  const rgxActivityLine = /(\d{1,2}:\d{2}\s*[AP]M)\s+(\d{1,2}:\d{2}\s*[AP]M)\s*$/i;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (line.length < 2) continue;

    let agentMatch = line.match(rgxAgent);
    if (agentMatch) {
      if (currentAgent) pushAgent(rosterData, currentAgent, currentID, buffer, defY, defM, defD);
      currentID = agentMatch[1]; currentAgent = agentMatch[2].replace(/["',]/g, "").trim(); 
      buffer = resetBuffer(); 
      if (String(currentID).startsWith("3")) buffer.isOffshore = true;
      continue;
    }

    if (currentAgent) {
       let dateMatch = line.match(rgxAnyDateLine);
       if (dateMatch) {
          const foundDate = dateMatch[1];
          if ((buffer.dateStr && buffer.dateStr !== foundDate) || (buffer.isOff && !buffer.dateStr)) {
             pushAgent(rosterData, currentAgent, currentID, buffer, defY, defM, defD);
             let wasOffshore = buffer.isOffshore;
             buffer = resetBuffer();
             buffer.isOffshore = wasOffshore;
          }
       }

       let shiftMatch = line.match(rgxShiftLine);
       if (shiftMatch) { buffer.dateStr = shiftMatch[1]; buffer.start = shiftMatch[2]; buffer.end = shiftMatch[3]; buffer.isOff = false; } 
       else if (line.includes("Off") && !buffer.start) { if (dateMatch) buffer.dateStr = dateMatch[1]; buffer.isOff = true; }

       const upper = line.toUpperCase();
       if (upper.includes("TI ") || upper.includes("OFFSHORE")) buffer.isOffshore = true;

       if (upper.includes("SICK") || upper.includes("MALADIE") || upper.includes("SICU")) {
           if (upper.includes("PLANNED") || upper.includes("STD") || upper.includes("LTD")) buffer.absentType = "Medical Leave";
           else buffer.absentType = "SICK";
       }
       else if (upper.includes("AWOL") || upper.includes("NCNS")) buffer.absentType = "NCNS";
       else if (upper.includes("VACATION") || upper.includes("VACP") || upper.includes("CONGÉ")) buffer.absentType = "VACATION";
       // NEW: Explicitly catch Personal Wellness / ALT
       else if (upper.match(/\bPW\b/) || upper.match(/\bALT\b/) || upper.includes("PERSONAL") || upper.includes("WELLNESS")) buffer.absentType = "PW/ALT";

       let actMatch = line.match(rgxActivityLine);
       if (actMatch) {
           const actStart = actMatch[1], actEnd = actMatch[2];
           
           if ((upper.includes("BREAK") || upper.includes("LUNCH") || upper.includes("REPAS") || upper.includes("PAUSE")) && !upper.includes("PAID LUNCH")) {
               let type = (upper.includes("LUNCH") || upper.includes("REPAS")) ? "Lunch" : "Break";
               buffer.breaks.push({ type: type, start: actStart, end: actEnd });
           } 
           else if (upper.includes("COACH") || upper.includes("TRN") || upper.includes("TRAINING") || upper.includes("HUDDLE") || upper.includes("MEET") || upper.includes("ONE ON ONE") || upper.includes("QUAL")) {
               buffer.breaks.push({ type: "Training", start: actStart, end: actEnd });
           }
           else if (upper.includes("ACSU") || upper.includes("SOLICITED") || upper.includes("VOLUNTARY") || upper.includes("LIBÉRATION")) {
               buffer.breaks.push({ type: "ACSU", start: actStart, end: actEnd });
           }
           else if (upper.includes("SAFE ONQUEUE") || upper.includes("SAFE EN LIGNE")) {
               buffer.breaks.push({ type: "SAFE", start: actStart, end: actEnd });
           }
           else if (upper.includes("ICL") || upper.includes("INCIDENT")) {
               buffer.breaks.push({ type: "ICL", start: actStart, end: actEnd });
           }
           else if (upper.includes("ULC") || upper.includes("FIRE") || upper.includes("FEU")) {
               buffer.breaks.push({ type: "ULC FIRE", start: actStart, end: actEnd });
           }
           // NEW: Prevent PW/ALT from making the engine think the agent is actively working
           else if (!upper.includes("VACATION") && !upper.includes("SICK") && !upper.includes("ABSENT") && !upper.includes("VACP") && !upper.includes("OFF") && !upper.match(/\bPW\b/) && !upper.match(/\bALT\b/) && !upper.includes("PERSONAL") && !upper.includes("WELLNESS")) {
               buffer.hasWork = true;
           }

           if (buffer.end && (upper.includes("VACATION") || upper.includes("VACP") || upper.includes("SICK") || upper.includes("PERSONAL") || upper.match(/\bPW\b/) || upper.match(/\bALT\b/))) {
               if (compareTimeStrings(actEnd, buffer.end)) buffer.end = actStart;
           }
       }
    }
  }

    if (currentAgent) pushAgent(rosterData, currentAgent, currentID, buffer, defY, defM, defD);
    Logger.log('[wfm] parsed ' + rosterData.length + ' entries in ' + (new Date().getTime() - t0) + 'ms');

    if (rosterData.length > 0) {
      var tWrite = new Date().getTime();
      // Force Shift Start/End (cols 4-5) to plain text BEFORE writing so Sheets
      // never coerces "9:00 AM" into a date/serial (which getDisplayValues would
      // return as "12/30/1899", losing the time for every reader).
      try { sheet.getRange(1, 4, rosterData.length + 1, 2).setNumberFormat('@'); } catch (e) {}
      sheet.getRange(2, 1, rosterData.length, 12).setValues(rosterData);
      Logger.log('[wfm] setValues done in ' + (new Date().getTime() - tWrite) + 'ms');

      // Permanent schedule archive: Raw Schedule is wiped every paste (current
      // period only), so going back in time loses all shift/break fidelity.
      // Append each pasted day into Schedule_History, deduped by agent+date so
      // re-pasting the same period self-heals instead of duplicating.
      try {
        var tArch = new Date().getTime();
        var nArch = archiveScheduleHistory(ss, rosterData);
        Logger.log('[wfm] archived ' + nArch + ' rows to Schedule_History in ' + (new Date().getTime() - tArch) + 'ms');
      } catch (eArch) { Logger.log('[wfm] Schedule_History archive failed: ' + eArch); }

      if (typeof AgentMonitor !== 'undefined') {
        var tMon = new Date().getTime();
        AgentMonitor.getPayload();
        Logger.log('[wfm] AgentMonitor.getPayload done in ' + (new Date().getTime() - tMon) + 'ms');
      }
      Logger.log('[wfm] total ' + (new Date().getTime() - t0) + 'ms');
      return `Synced ${rosterData.length} entries successfully to Local Database.`;
    }
    return "No valid data found.";
  } finally {
    if (typeof RegionRegistry !== 'undefined') {
      var tCommit = new Date().getTime();
      RegionRegistry.commitBatch();
      Logger.log('[wfm] RegionRegistry.commitBatch in ' + (new Date().getTime() - tCommit) + 'ms');
    }
  }
}

function resetBuffer() { return { breaks: [], start: null, end: null, dateStr: null, absentType: "", role: "", isOffshore: false, isOff: false, hasWork: false }; }

function pushAgent(roster, name, id, buf, defY, defM, defD) {
  if (!buf) return;
  if (!buf.start && !buf.isOff && !buf.absentType) return;
  if (buf.isOff) { buf.start = ""; buf.end = ""; }

  if (buf.hasWork) {
      if (buf.absentType === "VACATION") buf.absentType = "";
      if (buf.absentType === "SICK") buf.absentType = "Leaving Early (Sick)";
      if (buf.absentType === "PW/ALT") buf.absentType = "Leaving Early (PW/ALT)";
  }

  let finalDateObj = safeParseDateStr(buf.dateStr, defY, defM, defD);
  let finalDateStr = finalDateObj.str;
  
  let startEpoch = "", endEpoch = "";
  if (buf.start && buf.end) {
     const sObj = parseTime(buf.start);
     const eObj = parseTime(buf.end);
     if (sObj && eObj) {
        let sDate = new Date(finalDateObj.y, finalDateObj.m - 1, finalDateObj.d, sObj.h, sObj.m, 0);
        let eDate = new Date(finalDateObj.y, finalDateObj.m - 1, finalDateObj.d, eObj.h, eObj.m, 0);
        if (eDate < sDate) eDate.setDate(eDate.getDate() + 1);
        startEpoch = sDate.getTime(); endEpoch = eDate.getTime();
     }
  }

  let type = "Off";
  if (startEpoch && endEpoch) {
     let midEpoch = (startEpoch + endEpoch) / 2;
     const h = new Date(midEpoch).getHours();
     if (h >= 23 || h < 7) type = "Night";
     else if (h >= 15 && h < 23) type = "Evening";
     else type = "Morning";
  }

  let cleanName = name;
  if (name.includes(",")) {
     const parts = name.split(",");
     if(parts.length === 2) cleanName = `${parts[1].trim()} ${parts[0].trim()}`;
  }

  let finalOffshore = buf.isOffshore;
  // Prefer ID prefix 3 as the strongest positive signal; otherwise the keyword/flag
  // accumulated on the buffer. Consult & upsert the Region Registry so the result
  // is durable across imports and can be manually overridden.
  if (typeof RegionRegistry !== 'undefined') {
      const registered = RegionRegistry.getRegion(cleanName);
      const src = RegionRegistry.getSource(cleanName);
      if (registered && (src === 'manual' || src === 'masterlist')) {
          finalOffshore = registered === 'Offshore';
      } else {
          const source = (id && String(id).startsWith('3')) ? 'auto-wfm-id'
                        : (buf.isOffshore ? 'auto-wfm-keyword' : 'auto-wfm-default');
          RegionRegistry.upsert(cleanName, finalOffshore ? 'Offshore' : 'Onshore', source);
      }
  }

  roster.push([cleanName, id, finalDateStr, buf.start || "", buf.end || "", type, finalOffshore ? "Offshore" : "Onshore", JSON.stringify(buf.breaks), buf.role, buf.absentType, startEpoch, endEpoch]);
}

function safeParseDateStr(ds, defY, defM, defD) {
    if (!ds) return { y: defY, m: defM, d: defD, str: `${defY}-${defM < 10 ? '0'+defM : defM}-${defD < 10 ? '0'+defD : defD}` };
    let y = defY, m = defM, d = defD;
    let parts = ds.includes('-') ? ds.split('-') : ds.split('/');
    if (parts.length === 3) {
        if (parts[0].length === 4) { y = parseInt(parts[0]); m = parseInt(parts[1]); d = parseInt(parts[2]); }
        else { y = parseInt(parts[2]); m = parseInt(parts[0]); d = parseInt(parts[1]); }
    }
    if (y < 100) y += 2000;
    return { y, m, d, str: `${y}-${m < 10 ? '0'+m : m}-${d < 10 ? '0'+d : d}` };
}

// ─────────────────────────── SCHEDULE HISTORY ───────────────────────────
// Keep a permanent, full-fidelity record of every pasted schedule day so the
// SAFE Day Board can show real shifts + breaks for ANY past date — not just
// the current period that lives in (and is wiped from) "Raw Schedule".
//
// Columns mirror Raw Schedule exactly:
//   [0]Agent [1]ID [2]DateStr [3]Shift Start [4]Shift End [5]Shift Type
//   [6]Region [7]Breaks JSON [8]Role [9]AbsentType [10]StartEpoch [11]EndEpoch
//
// Dedup key = lowercased agent name + '|' + date. A re-paste of the same period
// replaces those agent-days rather than stacking duplicates. Rows older than
// HISTORY_KEEP_DAYS are pruned so the sheet can't grow without bound.
var HISTORY_KEEP_DAYS = 550; // ~18 months
function _histDedupKey(name, dateStr) {
  return String(name || '').trim().toLowerCase() + '|' + String(dateStr || '').trim();
}
function archiveScheduleHistory(ss, rosterData) {
  if (!rosterData || !rosterData.length) return 0;
  var HEADERS = ["Agent Name", "ID", "DateStr", "Shift Start", "Shift End", "Shift Type", "Region", "Breaks JSON", "Role", "AbsentType", "StartEpoch", "EndEpoch"];
  var hist = ss.getSheetByName("Schedule_History");
  if (!hist) {
    hist = ss.insertSheet("Schedule_History");
    hist.appendRow(HEADERS);
    hist.getRange(1, 1, 1, 12).setFontWeight("bold").setBackground("#e0e0e0");
  }

  // Agent-days present in this paste — these supersede any archived copy.
  var incoming = {};
  rosterData.forEach(function (r) { incoming[_histDedupKey(r[0], r[2])] = true; });

  // Prune cutoff (yyyy-MM-dd string compare is safe for ISO dates).
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - HISTORY_KEEP_DAYS);
  var cutoffStr = Utilities.formatDate(cutoff, "America/Toronto", "yyyy-MM-dd");

  var kept = [];
  var last = hist.getLastRow();
  if (last > 1) {
    var existing = hist.getRange(2, 1, last - 1, 12).getValues();
    for (var i = 0; i < existing.length; i++) {
      var row = existing[i];
      var dStr = String(row[2] || '').trim();
      if (dStr && dStr < cutoffStr) continue;                 // prune stale
      if (incoming[_histDedupKey(row[0], dStr)]) continue;    // superseded by this paste
      kept.push(row);
    }
  }

  var merged = kept.concat(rosterData);
  // Write in place + clear the tail instead of deleteRows(2,last-1). deleteRows is
  // O(all rows) and physically reflows the whole sheet — on an 18-month
  // Schedule_History (10k+ rows) that was the dominant import cost (minutes) and
  // the main reason a 2-month paste timed out. Overwriting yields the same
  // end-state far faster (one setValues + a tail clearContent).
  if (merged.length) {
    try { hist.getRange(1, 4, merged.length + 1, 2).setNumberFormat('@'); } catch (e) {}  // shift cols stay text
    hist.getRange(2, 1, merged.length, 12).setValues(merged);
  }
  var newLast = merged.length + 1;                       // header + data rows
  if (last > newLast) hist.getRange(newLast + 1, 1, last - newLast, 12).clearContent();
  return merged.length;
}

// parseTime() is shared from CORE/AgentMonitor.gs (one global scope in GAS).
function compareTimeStrings(t1, t2) {
    if (!t1 || !t2) return false;
    const n1 = t1.replace(/\s+/g, '').replace(/^0/, ''), n2 = t2.replace(/\s+/g, '').replace(/^0/, '');
    return n1 === n2;
}
