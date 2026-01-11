/**
 * MODULE: IMPORT HANDLER (SMART WFM)
 * - Based on your "Old Working Code" for stability.
 * - Adds ID-based Offshore detection (ID starts with 3).
 * - Handles Partial Sick vs Full Sick.
 * - Ignores Planned/LTD sickness for alerts.
 */

const ImportHandler = {
  run: function(text, date) {
    return processWFMImport(text, date);
  }
};

function processWFMImport(rawText, forcedDate) {
  if (!rawText) return "Error: No text provided.";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Ensure we are working with the correct timezone for timestamps
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
  
  if (!forcedDate) {
     forcedDate = Utilities.formatDate(new Date(), "America/Toronto", "yyyy-MM-dd");
  }
  
  const [defY, defM, defD] = forcedDate.split('-').map(Number);

  const rgxAgent = /Agent:\s*(\d+)\s*(.*)/i;
  const rgxAnyDateLine = /(\d{1,2}\/\d{1,2}\/\d{2,4})/;
  const rgxShiftLine = /(\d{1,2}\/\d{1,2}\/\d{2,4})\s+(\d{1,2}:\d{2}\s*[AP]M)\s+(\d{1,2}:\d{2}\s*[AP]M)/i;
  const rgxActivityLine = /(\d{1,2}:\d{2}\s*[AP]M)\s+(\d{1,2}:\d{2}\s*[AP]M)\s*$/i;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (line.length < 2) continue;

    // 1. DETECT NEW AGENT
    let agentMatch = line.match(rgxAgent);
    if (agentMatch) {
      if (currentAgent) pushAgent(rosterData, currentAgent, currentID, buffer, defY, defM, defD);
      currentID = agentMatch[1];
      currentAgent = agentMatch[2].replace(/["',]/g, "").trim(); // Clean extra quotes
      buffer = resetBuffer();
      continue;
    }

    if (currentAgent) {
       // 2. NEW DAY DETECTION
       let dateMatch = line.match(rgxAnyDateLine);
       if (dateMatch) {
          const foundDate = dateMatch[1];
          // If date changes, we save the previous buffer and reset (handle multi-day paste)
          if ((buffer.dateStr && buffer.dateStr !== foundDate) || (buffer.isOff && !buffer.dateStr)) {
             pushAgent(rosterData, currentAgent, currentID, buffer, defY, defM, defD);
             buffer = resetBuffer();
          }
       }

       // 3. SHIFT LINE (Main Shift)
       let shiftMatch = line.match(rgxShiftLine);
       if (shiftMatch) {
          buffer.dateStr = shiftMatch[1];
          buffer.start = shiftMatch[2];
          buffer.end = shiftMatch[3];
          buffer.isOff = false; 
       } 
       else if (line.includes("Off") && !buffer.start) {
          if (dateMatch) buffer.dateStr = dateMatch[1];
          buffer.isOff = true;
       }

       // 4. ACTIVITY PARSING
       const upper = line.toUpperCase();
       
       // Region Detection (Primary: TI code)
       if (upper.includes("TI ") || upper.includes("OFFSHORE")) buffer.isOffshore = true;
       
       // Role Detection (SAFE / Training)
       if (upper.includes("SAFE ONQUEUE") || upper.includes("SAFE EN LIGNE")) buffer.role = "SAFE";
       else if (upper.includes("COACHING") || upper.includes("TRN")) buffer.role = "Training";

       // Absence Detection
       if (upper.includes("SICK") || upper.includes("MALADIE") || upper.includes("SICU")) {
           if (upper.includes("PLANNED") || upper.includes("STD") || upper.includes("LTD")) {
               buffer.absentType = "Medical Leave"; // Treated as Planned
           } else {
               buffer.absentType = "SICK";
           }
       }
       else if (upper.includes("AWOL") || upper.includes("NCNS")) buffer.absentType = "NCNS";
       else if (upper.includes("VACATION") || upper.includes("VACP") || upper.includes("CONGÉ")) buffer.absentType = "VACATION";
       else if (upper.includes("TRAINING") || upper.includes("TRN")) buffer.absentType = "TRAINING";

       // 5. BREAKS & WORK DETECTION
       let actMatch = line.match(rgxActivityLine);
       if (actMatch) {
           const actStart = actMatch[1];
           const actEnd = actMatch[2];

           // A. BREAKS
           if ((upper.includes("BREAK") || upper.includes("LUNCH") || upper.includes("REPAS") || upper.includes("PAUSE")) && !upper.includes("PAID LUNCH")) {
               let type = (upper.includes("LUNCH") || upper.includes("REPAS")) ? "Lunch" : "Break";
               buffer.breaks.push({ type: type, start: actStart, end: actEnd });
            } 
           
           // B. WORK SEGMENT (Keeps them Active if they worked partial shift)
           else if (!upper.includes("VACATION") && !upper.includes("SICK") && !upper.includes("ABSENT") && !upper.includes("VACP") && !upper.includes("OFF")) {
               buffer.hasWork = true;
            }

           // C. PARTIAL ABSENCE CUTOFF
           if (buffer.end && (upper.includes("VACATION") || upper.includes("VACP") || upper.includes("SICK") || upper.includes("PERSONAL"))) {
               if (compareTimeStrings(actEnd, buffer.end)) {
                   buffer.end = actStart; // Snap shift end to start of absence
               }
           }
       }
    }
  }

  if (currentAgent) pushAgent(rosterData, currentAgent, currentID, buffer, defY, defM, defD);
  
  if (rosterData.length > 0) {
    // Write starting at Row 2
    sheet.getRange(2, 1, rosterData.length, 12).setValues(rosterData);
    
    // Sync to AgentMonitor instantly if available
    if (typeof AgentMonitor !== 'undefined') AgentMonitor.getPayload();
    
    return `Synced ${rosterData.length} entries.`;
  }
  return "No valid data found.";
}

function resetBuffer() {
  return { 
      breaks: [], start: null, end: null, dateStr: null, 
      absentType: "", role: "", isOffshore: false, isOff: false, hasWork: false 
  };
}

function pushAgent(roster, name, id, buf, defY, defM, defD) {
  if (!buf) return;
  // If no start time, no Off status, and no absence, ignore (empty line)
  if (!buf.start && !buf.isOff && !buf.absentType) return;
  if (buf.isOff) { buf.start = ""; buf.end = ""; }

  // --- SMART LOGIC ---

  // 1. Offshore ID Check (The Backup Rule)
  // If ID starts with 3, FORCE Offshore
  if (String(id).startsWith("3")) {
      buf.isOffshore = true;
  }

  // 2. Partial Work Fix
  if (buf.hasWork) {
      if (buf.absentType === "VACATION") buf.absentType = ""; // Worked part of vacation? Treat as active.
      if (buf.absentType === "SICK") {
          // Worked part of sick day? Treat as active with note.
          // We clear absentType so they don't show red "SICK", but we can add a sub-label if needed later.
          // For now, we clear it so they show up on the floor.
          buf.absentType = "Leaving Early (Sick)"; 
      }
  }

  // 3. Epoch Calculation
  let startEpoch = "", endEpoch = "";
  let finalDateStr = buf.dateStr;
  if (!finalDateStr) finalDateStr = `${defM}/${defD}/${defY}`; // Use fallback date

  if (buf.start && buf.end) {
     let y, m, d;
     if (buf.dateStr) {
        const parts = buf.dateStr.split('/');
        m = parseInt(parts[0]);
        d = parseInt(parts[1]);
        y = parseInt(parts[2]);
        if (y < 100) y += 2000;
     } else {
        y = defY; m = defM; d = defD;
     }

     const sObj = parseTime(buf.start);
     const eObj = parseTime(buf.end);
     if (sObj && eObj) {
        let sDate = new Date(y, m - 1, d, sObj.h, sObj.m, 0);
        let eDate = new Date(y, m - 1, d, eObj.h, eObj.m, 0);
        
        // Overnight Logic: If End < Start, add 1 day
        if (eDate < sDate) eDate.setDate(eDate.getDate() + 1);

        startEpoch = sDate.getTime();
        endEpoch = eDate.getTime();
     }
  }

  // 4. Shift Name
  let type = "Off";
  if (startEpoch) {
     const h = new Date(startEpoch).getHours();
     if (h >= 21 || h < 5) type = "Night";
     else if (h >= 14) type = "Evening";
     else type = "Day";
  }

  let cleanName = name;
  if (name.includes(",")) {
     const parts = name.split(",");
     if(parts.length === 2) cleanName = `${parts[1].trim()} ${parts[0].trim()}`;
  }

  // OUTPUT ROW
  roster.push([
    cleanName, 
    id, 
    finalDateStr,
    buf.start || "", 
    buf.end || "",
    type, 
    buf.isOffshore ? "Offshore" : "Onshore",
    JSON.stringify(buf.breaks),
    buf.role, 
    buf.absentType,
    startEpoch, 
    endEpoch
  ]);
}

function parseTime(tStr) {
   if(!tStr) return null;
   const match = tStr.match(/(\d{1,2}):(\d{2})\s*([AP]M)/i);
   if(!match) return null;
   let h = parseInt(match[1]);
   let m = parseInt(match[2]);
   let amp = match[3].toUpperCase();
   if (amp === "PM" && h < 12) h += 12;
   if (amp === "AM" && h === 12) h = 0;
   return { h, m };
}

function compareTimeStrings(t1, t2) {
    if (!t1 || !t2) return false;
    const n1 = t1.replace(/\s+/g, '').replace(/^0/, '');
    const n2 = t2.replace(/\s+/g, '').replace(/^0/, '');
    return n1 === n2;
}
