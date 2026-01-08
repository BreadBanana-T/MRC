/**
 * WFM IMPORT ENGINE
 */

// --- LINKER: Connects this file to Code.gs ---
const ImportHandler = {
  run: function(text, date) {
    // Delegates the call to the main function below
    return processWFMImport(text, date);
  }
};
// ---------------------------------------------

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
  
  // --- SAFETY FIX FOR DATE ---
  if (!forcedDate) {
     forcedDate = Utilities.formatDate(new Date(), "America/Toronto", "yyyy-MM-dd");
  }
  
  const [defY, defM, defD] = forcedDate.split('-').map(Number);
  // ---------------------------

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
      currentAgent = agentMatch[2].trim(); 
      buffer = resetBuffer();
      continue;
    }

    if (currentAgent) {
       // 2. NEW DAY DETECTION
       let dateMatch = line.match(rgxAnyDateLine);
       if (dateMatch) {
          const foundDate = dateMatch[1];
          // If buffer has data and date changes, save & reset
          if ((buffer.dateStr && buffer.dateStr !== foundDate) || (buffer.isOff && !buffer.dateStr)) {
             pushAgent(rosterData, currentAgent, currentID, buffer, defY, defM, defD);
             buffer = resetBuffer();
          }
       }

       // 3. SHIFT LINE
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
       if (upper.includes("TI ") || upper.includes("OFFSHORE")) buffer.isOffshore = true;
       
       if (upper.includes("SICK")) buffer.absentType = "SICK";
       else if (upper.includes("AWOL") || upper.includes("NCNS")) buffer.absentType = "NCNS";
       else if (upper.includes("VACATION") || upper.includes("VACP")) buffer.absentType = "VACATION";
       else if (upper.includes("TRAINING") || upper.includes("TRN")) buffer.absentType = "TRAINING";

       // 5. BREAKS
       if ((upper.includes("BREAK") || upper.includes("LUNCH") || upper.includes("REPAS")) && !upper.includes("PAID LUNCH")) {
           let actMatch = line.match(rgxActivityLine);
           if (actMatch) {
               let type = (upper.includes("LUNCH") || upper.includes("REPAS")) ? "Lunch" : "Break";
               buffer.breaks.push({ type: type, start: actMatch[1], end: actMatch[2] });
           }
       }
    }
  }

  if (currentAgent) pushAgent(rosterData, currentAgent, currentID, buffer, defY, defM, defD);

  if (rosterData.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rosterData.length, 12).setValues(rosterData);
    return `Synced ${rosterData.length} entries.`;
  }
  return "No valid data found.";
}

function resetBuffer() {
  return { breaks: [], start: null, end: null, dateStr: null, absentType: "", isOffshore: false, isOff: false };
}

function pushAgent(roster, name, id, buf, defY, defM, defD) {
  if (!buf) return;
  if (!buf.start && !buf.isOff && !buf.absentType) return;
  // Skip empty buffers

  if (buf.isOff) { buf.start = ""; buf.end = ""; }

  let startEpoch = "", endEpoch = "";
  let finalDateStr = buf.dateStr;
  
  // Fallback to default date if none found in text
  if (!finalDateStr) finalDateStr = `${defM}/${defD}/${defY}`;

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
        
        if (eDate <= sDate) eDate.setDate(eDate.getDate() + 1);

        startEpoch = sDate.getTime();
        endEpoch = eDate.getTime();
     }
  }

  let type = "Off";
  if (startEpoch) {
     const h = new Date(startEpoch).getHours();
     type = h >= 14 ? "Evening" : "Day";
  }

  let cleanName = name;
  if (name.includes(",")) {
      const parts = name.split(",");
      if(parts.length === 2) cleanName = `${parts[1].trim()} ${parts[0].trim()}`;
  }

  roster.push([
    cleanName, id, finalDateStr,
    buf.start || "", buf.end || "",
    type, 
    buf.isOffshore ? "Offshore" : "Onshore",
    JSON.stringify(buf.breaks),
    "", 
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
