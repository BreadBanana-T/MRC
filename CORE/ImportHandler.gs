/**
 * MODULE: IMPORT HANDLER (SMART WFM - LOCAL HOST ONLY)
 */

const ImportHandler = {
  run: function(text, date) { return processWFMImport(text, date); }
};

function processWFMImport(rawText, forcedDate) {
  if (!rawText) return "Error: No text provided.";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone("America/Toronto");
  let sheet = ss.getSheetByName("Raw Schedule");
  if (!sheet) { sheet = ss.insertSheet("Raw Schedule"); }
  
  sheet.clear();
  sheet.appendRow(["Agent Name", "ID", "DateStr", "Shift Start", "Shift End", "Shift Type", "Region", "Breaks JSON", "Role", "AbsentType", "StartEpoch", "EndEpoch"]);
  sheet.getRange(1, 1, 1, 12).setFontWeight("bold").setBackground("#e0e0e0");

  const lines = rawText.split(/\r?\n/);
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
           else if (!upper.includes("VACATION") && !upper.includes("SICK") && !upper.includes("ABSENT") && !upper.includes("VACP") && !upper.includes("OFF")) {
               buffer.hasWork = true;
           }

           if (buffer.end && (upper.includes("VACATION") || upper.includes("VACP") || upper.includes("SICK") || upper.includes("PERSONAL"))) {
               if (compareTimeStrings(actEnd, buffer.end)) buffer.end = actStart;
           }
       }
    }
  }

  if (currentAgent) pushAgent(rosterData, currentAgent, currentID, buffer, defY, defM, defD);
  if (rosterData.length > 0) {
    sheet.getRange(2, 1, rosterData.length, 12).setValues(rosterData);
    if (typeof AgentMonitor !== 'undefined') AgentMonitor.getPayload();
    return `Synced ${rosterData.length} entries successfully to Local Database.`;
  }
  return "No valid data found.";
}

function resetBuffer() { return { breaks: [], start: null, end: null, dateStr: null, absentType: "", role: "", isOffshore: false, isOff: false, hasWork: false }; }

function pushAgent(roster, name, id, buf, defY, defM, defD) {
  if (!buf) return;
  if (!buf.start && !buf.isOff && !buf.absentType) return;
  if (buf.isOff) { buf.start = ""; buf.end = ""; }

  if (buf.hasWork) {
      if (buf.absentType === "VACATION") buf.absentType = "";
      if (buf.absentType === "SICK") buf.absentType = "Leaving Early (Sick)";
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

  roster.push([cleanName, id, finalDateStr, buf.start || "", buf.end || "", type, buf.isOffshore ? "Offshore" : "Onshore", JSON.stringify(buf.breaks), buf.role, buf.absentType, startEpoch, endEpoch]);
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

function parseTime(tStr) {
   if(!tStr) return null; const match = tStr.match(/(\d{1,2}):(\d{2})\s*([AP]M)?/i); if(!match) return null;
   let h = parseInt(match[1]), m = parseInt(match[2]), amp = match[3] ? match[3].toUpperCase() : null;
   if (amp === "PM" && h < 12) h += 12;
   if (amp === "AM" && h === 12) h = 0; return { h, m };
}
function compareTimeStrings(t1, t2) {
    if (!t1 || !t2) return false;
    const n1 = t1.replace(/\s+/g, '').replace(/^0/, ''), n2 = t2.replace(/\s+/g, '').replace(/^0/, '');
    return n1 === n2;
}
