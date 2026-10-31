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
          // Compare the DAY, not the text. The report repeats the current day's
          // date mid-day in another format ("10/29/26" vs "10/29/2026", e.g. at a
          // page break); a raw string compare split the day there, and the rest
          // of its breaks, lunches and furlough became a separate "Off" row dated
          // the paste day (1,525 such rows in Schedule_History). A real new day
          // still splits.
          const sameDay = buffer.dateStr && safeParseDateStr(buffer.dateStr, defY, defM, defD).str === safeParseDateStr(foundDate, defY, defM, defD).str;
          if ((buffer.dateStr && !sameDay) || (buffer.isOff && !buffer.dateStr)) {
             pushAgent(rosterData, currentAgent, currentID, buffer, defY, defM, defD);
             let wasOffshore = buffer.isOffshore;
             buffer = resetBuffer();
             buffer.isOffshore = wasOffshore;
          }
       }

       let shiftMatch = line.match(rgxShiftLine);
       if (shiftMatch && buffer.start && buffer.dateStr && safeParseDateStr(buffer.dateStr, defY, defM, defD).str === safeParseDateStr(shiftMatch[1], defY, defM, defD).str) shiftMatch = null;  // same day repeated: keep its shift
       if (shiftMatch) { buffer.dateStr = shiftMatch[1]; buffer.start = shiftMatch[2]; buffer.end = shiftMatch[3]; buffer.shiftEnd = shiftMatch[3]; buffer.isOff = false; } 
       else if (line.includes("Off") && !buffer.start) { if (dateMatch) buffer.dateStr = dateMatch[1]; buffer.isOff = true; }

       const upper = line.toUpperCase();
       if (upper.includes("TI ") || upper.includes("OFFSHORE")) buffer.isOffshore = true;

       // One shared taxonomy with the reporting pipeline — see
       // WorkforceTracker.classifyAbsenceActivity. This block used to be a weaker
       // private copy with NO UNAB and NO LOA branch, so 208 UNAB agent-days and 40
       // LOA/TI LOA agent-days reached the Floor with no absence type at all. It also
       // matched CONGÉ before PW/ALT, which made the "catch Personal Wellness" branch
       // unreachable and filed those agents as VACATION.
       var _absType = (typeof classifyAbsenceActivity === 'function') ? classifyAbsenceActivity(line) : '';
       if (_absType) buffer.absentType = _absType;

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
           // Absence blocks are kept with their times (type 'Absence') so the Floor
           // can put a part-day absence in Needs Attention only WHILE it runs, and
           // the SAFE day board can draw it.
           else if (classifyAbsenceActivity(line) || looksLikeAbsenceWording(line)) {
               buffer.breaks.push({ type: "Absence", start: actStart, end: actEnd, abs: classifyAbsenceActivity(line) || 'Absence' });
           }
           // Anything that isn't a recognised absence, isn't worded as an absence, and
           // isn't an "Off" line counts as real work. The old guard tested for
           // "ABSENT", but the real strings say "ABSENCE" ("UNAB Absence non
           // autorisée…", "LOA Absence / Leave of Absence") — and "ABSENT" is not a
           // substring of "ABSENCE", so the importer positively concluded those agents
           // had worked.
           else if (!upper.includes("OFF")) {
               buffer.hasWork = true;
           }

           // An absence block running to the end of the shift means the agent left
           // early — truncate the shift at its start. Driven off the shared classifier
           // so UNAB/LOA are included too.
           if (buffer.end && (classifyAbsenceActivity(line) || looksLikeAbsenceWording(line))) {
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

function resetBuffer() { return { breaks: [], start: null, end: null, shiftEnd: null, dateStr: null, absentType: "", role: "", isOffshore: false, isOff: false, hasWork: false }; }

function pushAgent(roster, name, id, buf, defY, defM, defD) {
  if (!buf) return;
  if (!buf.start && !buf.isOff && !buf.absentType) return;
  if (buf.isOff) { buf.start = ""; buf.end = ""; }
  // Absent for the WHOLE shift (no working activity at all): the leaving-early
  // trim above pulled the end back to the absence start, i.e. end == start, so
  // the agent read as "Shift Ended" and vanished from Needs Attention. Keep the
  // full scheduled shift so they show for as long as it runs.
  if (!buf.hasWork && buf.absentType && buf.shiftEnd && !buf.isOff) buf.end = buf.shiftEnd;

  // An agent can be absent for part of a day and work the rest. Previously a
  // worked day WIPED a VACATION absence outright, and bereavement (which sets
  // hasWork through its "Congé" wording) was silently cleared with it — those
  // agents rendered as plain active. Now the absence is kept and simply marked as
  // partial, so the Floor can say "UNAB (part day)" instead of losing the person.
  // An agent can be absent for part of a day and work the rest. The partial state
  // is encoded in the type string rather than a new Raw Schedule column: every
  // downstream matcher is \b-anchored, so "UNAB (part day)" still reads as UNAB,
  // and the label is already what the Floor should display.
  if (buf.hasWork && buf.absentType) {
      if (buf.absentType === "VACATION") buf.absentType = "";
      else if (buf.absentType.indexOf('(part day)') === -1) buf.absentType = buf.absentType + " (part day)";
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
// Dedup key = lowercased agent name + '|' + day. A re-paste of the same period
// replaces those agent-days rather than stacking duplicates (newest paste wins).
// Rows older than HISTORY_KEEP_DAYS are pruned so the sheet can't grow without
// bound. Rows are kept sorted by day, then agent, so a paste of recent days only
// rewrites the tail of the sheet instead of all of it.
var HISTORY_KEEP_DAYS = 550; // ~18 months

// Day key ("yyyy-MM-dd") for a DateStr cell. Sheets turns the "2026-06-15" we
// write into a real date, so reading it back gives a Date — and String(Date)
// never equalled the incoming "2026-06-15". That is why every re-paste stacked
// instead of replacing (90k rows, 41k of them duplicates, by Oct 2026) and why
// nothing was ever pruned.
var _histDayMemo = {};
function _histDay(v) {
  if (v instanceof Date) {
    var t = v.getTime();
    return _histDayMemo[t] || (_histDayMemo[t] = Utilities.formatDate(v, "America/Toronto", "yyyy-MM-dd"));
  }
  var s = String(v == null ? '' : v).trim(), m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  return s;
}
function _histDedupKey(name, dateVal) {
  return String(name || '').trim().toLowerCase() + '|' + _histDay(dateVal);
}
// "9:30 PM" → 1290; 24-hour "21:30" too. null when unreadable.
function _histMins(t) {
  var m = String(t == null ? '' : t).match(/(\d{1,2}):(\d{2})\s*([AP]M)?/i);
  if (!m) return null;
  var h = +m[1];
  if (m[3]) { h = h % 12; if (m[3].toUpperCase() === 'PM') h += 12; }
  return h * 60 + (+m[2]);
}

// Folds back the days the old importer split in two (see processWFMImport): an
// "Off" row with no shift but with breaks, right after the same agent's real
// day, whose breaks carry on where that day's stopped and sit inside its shift.
// Works on rows in paste order (it relies on the split half following its day).
function _histMergeFragments(rows) {
  var out = [], merged = 0;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i], p = out.length ? out[out.length - 1] : null;
    if (p && !r[3] && r[7] && r[7] !== '[]' && p[3] &&
        String(p[0]).trim().toLowerCase() === String(r[0]).trim().toLowerCase()) {
      var st = _histMins(p[3]), en = _histMins(p[4]);
      if (st != null && en != null) {
        if (en <= st) en += 1440;
        var rel = function (t) { var x = _histMins(t); return x == null ? null : (x < st ? x + 1440 : x); };
        var pb, fb;
        try { pb = JSON.parse(p[7] || '[]'); fb = JSON.parse(r[7]); } catch (e) { pb = null; }
        if (pb && fb && fb.length) {
          var last = st;
          pb.forEach(function (b) { var e = rel(b.end); if (e != null && e > last) last = e; });
          var cont = fb.every(function (b) { var s = rel(b.start); return s != null && s >= last && s <= en; });
          if (cont) { p[7] = JSON.stringify(pb.concat(fb)); merged++; continue; }
        }
      }
    }
    out.push(r);
  }
  return { rows: out, merged: merged };
}

// existing (sheet order) + incoming (paste order) → the sheet's new rows:
// fragments merged, one row per agent-day (incoming wins, then the newest
// existing copy), sorted by day then agent.
function _histRebuild(existing, incomingRows) {
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - HISTORY_KEEP_DAYS);
  var cutoffStr = Utilities.formatDate(cutoff, "America/Toronto", "yyyy-MM-dd");
  // Prune in month-sized batches: dropping the oldest day shifts every row up,
  // which turns the tail-only write into a full rewrite.
  cutoff.setDate(cutoff.getDate() - 30);
  var pruneIfBefore = Utilities.formatDate(cutoff, "America/Toronto", "yyyy-MM-dd");

  var ex = _histMergeFragments(existing), inc = _histMergeFragments(incomingRows || []);
  var byKey = {}, keys = [], dropped = 0, prune = false;
  var put = function (row) {
    var k = _histDedupKey(row[0], row[2]);
    if (k in byKey) dropped++; else keys.push(k);
    byKey[k] = row;
  };
  var incKeys = {};
  inc.rows.forEach(function (r) { incKeys[_histDedupKey(r[0], r[2])] = true; });
  ex.rows.forEach(function (r) { var d = _histDay(r[2]); if (d && d < pruneIfBefore) prune = true; });
  ex.rows.forEach(function (r) {
    var d = _histDay(r[2]);
    if (prune && d && d < cutoffStr) { dropped++; return; }
    if (incKeys[_histDedupKey(r[0], r[2])]) { dropped++; return; }
    put(r);
  });
  inc.rows.forEach(put);
  var rows = keys.map(function (k) { return { k: k.split('|')[1] + '|' + k.split('|')[0], r: byKey[k] }; });
  rows.sort(function (a, b) { return a.k < b.k ? -1 : (a.k > b.k ? 1 : 0); });
  return { rows: rows.map(function (x) { return x.r; }), merged: ex.merged + inc.merged, dropped: dropped };
}

// Writes `rows` over `existing`, starting at the first row that differs, and
// blanks any leftover tail. Returns the index of the first rewritten row.
function _histWrite(sheet, existing, rows) {
  var sig = function (r) {
    var a = [];
    for (var j = 0; j < 12; j++) { var v = r[j]; a.push(j === 2 ? _histDay(v) : (v instanceof Date ? v.getTime() : String(v == null ? '' : v))); }
    return a.join('\u0001');
  };
  var from = 0, n = Math.min(existing.length, rows.length);
  while (from < n && sig(existing[from]) === sig(rows[from])) from++;
  if (from < rows.length) {
    var tail = rows.slice(from);
    try { sheet.getRange(from + 2, 4, tail.length, 2).setNumberFormat('@'); } catch (e) {}  // shift cols stay text
    sheet.getRange(from + 2, 1, tail.length, 12).setValues(tail);
  }
  if (existing.length > rows.length) sheet.getRange(rows.length + 2, 1, existing.length - rows.length, 12).clearContent();
  return from;
}

function _histSheet(ss) {
  var hist = ss.getSheetByName("Schedule_History");
  if (!hist) {
    hist = ss.insertSheet("Schedule_History");
    hist.appendRow(["Agent Name", "ID", "DateStr", "Shift Start", "Shift End", "Shift Type", "Region", "Breaks JSON", "Role", "AbsentType", "StartEpoch", "EndEpoch"]);
    hist.getRange(1, 1, 1, 12).setFontWeight("bold").setBackground("#e0e0e0");
  }
  return hist;
}

function archiveScheduleHistory(ss, rosterData) {
  if (!rosterData || !rosterData.length) return 0;
  var hist = _histSheet(ss);
  var last = hist.getLastRow();
  var existing = last > 1 ? hist.getRange(2, 1, last - 1, 12).getValues() : [];
  var res = _histRebuild(existing, rosterData);
  var from = _histWrite(hist, existing, res.rows);
  Logger.log('[wfm] history: ' + res.rows.length + ' rows, rewrote from row ' + (from + 2) + ' (' + (res.rows.length - from) + ' rows), ' + res.dropped + ' replaced/pruned, ' + res.merged + ' split days rejoined');
  return res.rows.length;
}

// One-off clean-up of Schedule_History: rejoins split days, keeps the newest
// copy of each agent-day, sorts. Safe to run any time; see repairScheduleData().
function compactScheduleHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hist = _histSheet(ss);
  var last = hist.getLastRow();
  var existing = last > 1 ? hist.getRange(2, 1, last - 1, 12).getValues() : [];
  var res = _histRebuild(existing, []);
  _histWrite(hist, existing, res.rows);
  var out = { before: existing.length, after: res.rows.length, removed: res.dropped, rejoined: res.merged };
  Logger.log('[compactScheduleHistory] ' + JSON.stringify(out));
  return out;
}

// parseTime() is shared from CORE/AgentMonitor.gs (one global scope in GAS).
function compareTimeStrings(t1, t2) {
    if (!t1 || !t2) return false;
    const n1 = t1.replace(/\s+/g, '').replace(/^0/, ''), n2 = t2.replace(/\s+/g, '').replace(/^0/, '');
    return n1 === n2;
}
