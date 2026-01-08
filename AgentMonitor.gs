/**
 * MODULE: AGENT MONITOR
 * Consumes Raw Schedule and StatusTracker to build the dashboard payload.
 */

const AgentMonitor = {
  getPayload: function() { return compileFloorData(); },

  // Delegate all writes to StatusTracker
  setStatus: function(n, t, v) { return StatusTracker.updateStatus(n, t, v); },
  updateAgentBreaks: function(n, j) { return StatusTracker.updateBreaks(n, j); },
  logOvertime: function(n, s, e, bs, be) { 
      const res = StatusTracker.logOvertime(n, s, e, bs, be); 
      if(typeof NotificationHandler !== 'undefined') NotificationHandler.add(n, `OVERTIME: ${s}-${e}`);
      return res;
  },
  
  updateBreak: function() { return "Deprecated"; },
  clearFlag: function() { return "Flag Cleared"; }
};

/* --- PUBLIC ENDPOINTS --- */

function getFloorStatus() {
  return compileFloorData();
}

// RESTORED: This was removed previously but might be needed by legacy calls or specific router setups
function getLiveDashboardData() {
  try {
    if (typeof WeatherService !== 'undefined') {
      return JSON.stringify(WeatherService.fetch());
    }
    return JSON.stringify({ weather: {}, alerts: [] });
  } catch (e) {
    return JSON.stringify({ weather: {}, alerts: [] });
  }
}

function updateAgentStatus(name, type, value) {
  return AgentMonitor.setStatus(name, type, value);
}

function updateAgentBreaks(name, jsonBreaks) {
  return AgentMonitor.updateAgentBreaks(name, jsonBreaks);
}

function submitOvertime(name, start, end, bStart, bEnd) {
  return AgentMonitor.logOvertime(name, start, end, bStart, bEnd);
}

/* --- CORE LOGIC --- */

function compileFloorData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Raw Schedule");
  
  const emptyFloor = {
    active: [], safe: [], icl: [], training: [],
    onBreak: [], upcomingBreak: [], startingSoon: [],
    vacation: [], planned: [], unplanned: [], off: []
  };
  
  const overrides = StatusTracker.getConsolidatedData(); 
  const data = (sheet && sheet.getLastRow() > 1) ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
  
  const now = new Date().getTime();
  const agentMap = new Map();
  
  const SCORES = {
    'safe': 60, 'icl': 60,
    'active': 50, 'onBreak': 50, 'upcomingBreak': 50,
    'startingSoon': 40, 'training': 30, 'unplanned': 20,
    'vacation': 10, 'planned': 10, 'off': 0
  };

  const processEntry = (rawName, row) => {
    const cleanName = String(rawName).trim().toLowerCase();
    const persistentData = overrides.get(cleanName) || {};

    let shiftType = row ? row[5] : "Off";
    let region = row ? row[6] : "Offshore";
    let startEpoch = row ? Number(row[10]) : 0;
    let endEpoch = row ? Number(row[11]) : 0;
    let originalBreaksJson = row ? row[7] : "[]";
    
    let role = (persistentData.role !== undefined && persistentData.role !== "") ? persistentData.role : (row ? row[8] : "");
    let absentType = (persistentData.absent !== undefined && persistentData.absent !== "") ? persistentData.absent : (row ? row[9] : "");
    
    const otList = persistentData.ot || [];
    const customBreaks = persistentData.breaks;
    const isOT = otList.length > 0;

    if ((shiftType === "Off" || region === "Off") && !isOT) return;
    if ((!startEpoch || !endEpoch) && !isOT) return;

    if (startEpoch && endEpoch) {
       if (endEpoch < (now - 60000) && !isOT) return; 
       const TWO_HOURS_MS = 7200000;
       if (startEpoch > (now + TWO_HOURS_MS) && !isOT) return; 
    }

    let dateStr = row && row[2] instanceof Date ? Utilities.formatDate(row[2], "America/Toronto", "M/d/yy") : (row ? row[2] : "");
    let shiftStr = "";
    
    if(startEpoch && endEpoch) {
       const sFmt = Utilities.formatDate(new Date(startEpoch), "America/Toronto", "HH:mm");
       const eFmt = Utilities.formatDate(new Date(endEpoch), "America/Toronto", "HH:mm");
       shiftStr = `${sFmt} - ${eFmt}`;
    }

    let activeBreaksJson = customBreaks ? customBreaks : originalBreaksJson;
    const isModified = !!customBreaks;
    const rawBreaks = parseBreaksSafe(activeBreaksJson);
    const originalBreaksRaw = isModified ? parseBreaksSafe(originalBreaksJson) : [];
    
    if (isOT) {
       otList.forEach(o => {
           if(o.bStart && o.bStart !== "-" && o.bEnd && o.bEnd !== "-") {
               rawBreaks.push({ type: "OT Break", start: o.bStart, end: o.bEnd });
           }
       });
    }

    const enrichedBreaks = [];
    let nextBreakStr = null;
    
    if((startEpoch || isOT) && rawBreaks.length > 0) {
       const baseTime = startEpoch ? new Date(startEpoch) : new Date();
       const shiftDateStr = Utilities.formatDate(baseTime, "America/Toronto", "yyyy-MM-dd");
       
       rawBreaks.sort((a, b) => parseTimeValue(a.start) - parseTimeValue(b.start));

       let breakCounter = 0;
       rawBreaks.forEach(b => {
          let displayLabel = b.type;
          if (b.type === 'Break') {
             breakCounter++;
             displayLabel = `${breakCounter}${['st','nd','rd'][breakCounter-1]||'th'} Break`;
          }

          const bStart24 = Utilities.parseDate(`${shiftDateStr} ${b.start}`, "America/Toronto", "yyyy-MM-dd hh:mm a");
          const bEnd24 = Utilities.parseDate(`${shiftDateStr} ${b.end}`, "America/Toronto", "yyyy-MM-dd hh:mm a");
          
          if(bStart24 && bEnd24) {
             let bs = bStart24.getTime();
             let be = bEnd24.getTime();
             if (startEpoch && bs < startEpoch) { bs += 86400000; be += 86400000; }
              
             enrichedBreaks.push({ 
                 type: displayLabel, originalType: b.type,
                 start: b.start, end: b.end, 
                 epochStart: bs, epochEnd: be 
             });

             if (bs > now && !nextBreakStr && (bs - now < 1800000)) { 
                 nextBreakStr = `${b.start} - ${b.end}`;
             }
          }
       });
    }

    if (role === "SAFE") region += " [SAFE]";
    if (role === "ICL") region += " [ICL]";

    let agent = {
      name: rawName, id: row ? row[1] : "", region: region, shift: shiftStr, shiftType: shiftType,
      dateStr: dateStr, subStatus: "",
      rawBreaks: enrichedBreaks, 
      originalBreaks: originalBreaksRaw, 
      isModified: isModified,            
      breakTimeStr: nextBreakStr,
      timer: "--:--", 
      auxLabel: "Remaining",
      onBreakNow: false, 
      startEpoch: startEpoch,
      isOT: isOT 
    };

    let category = "active";
    
    if (absentType) {
        const upper = absentType.toUpperCase();
        if (upper.match(/NCNS|UNAB|AWOL|COMP/) && rawBreaks.length > 1) category = "active"; 
        else if (upper.match(/SICK|NCNS|UNAB|AWOL|COMP/)) { category = "unplanned"; agent.subStatus = absentType; }
        else if (upper.match(/TRAIN|TRN/)) { category = "training"; agent.subStatus = "Training"; }
        else { category = "vacation"; agent.subStatus = absentType; }
    }
    else if (role === "SAFE") category = "safe";
    else if (role === "ICL") category = "icl";
    else if (role === "Training") category = "training";

    if (['active', 'safe', 'icl'].includes(category)) {
        if (startEpoch && startEpoch > now) {
             const diff = startEpoch - now;
             const twoHours = 120 * 60 * 1000; 
             if (diff <= twoHours) {
                 category = "startingSoon";
                 agent.subStatus = `Starts in ${Math.ceil(diff/60000)}m`;
             } 
        } else {
             let breakFound = false;
             for (let b of enrichedBreaks) {
                // ACTIVE BREAK
                if (now >= b.epochStart && now <= b.epochEnd) {
                   category = "onBreak";
                   agent.subStatus = b.type;
                   const rem = Math.ceil((b.epochEnd - now)/60000);
                   agent.timer = `${rem}m`;
                   agent.auxLabel = "Remaining"; 
                   agent.onBreakNow = true;
                   breakFound = true;
                   break;
                }
                
                // UPCOMING BREAK (Next 30 mins)
                const timeToBreak = b.epochStart - now;
                if (timeToBreak > 0 && timeToBreak <= 1800000) {
                   category = "upcomingBreak";
                   agent.timer = `${b.start} - ${b.end}`;
                   agent.auxLabel = `in ${Math.ceil(timeToBreak/60000)}m`;
                   breakFound = true;
                   break;
                }
             }
             if(!breakFound && category === "startingSoon") category = "active";
        }
    }

    const newScore = SCORES[category] || 0;
    if (agentMap.has(rawName)) {
       const existing = agentMap.get(rawName);
       if (newScore > SCORES[existing.category]) agentMap.set(rawName, { agent, category });
    } else {
       agentMap.set(rawName, { agent, category });
    }
  };

  data.forEach(row => processEntry(row[0], row));

  overrides.forEach((val, key) => {
      if (!agentMap.has(key) && !agentMap.has(toTitleCase(key)) && val.ot.length > 0) {
          processEntry(toTitleCase(key), null); 
      }
  });

  agentMap.forEach(item => {
      const targetCat = item.category;
      if (emptyFloor[targetCat]) emptyFloor[targetCat].push(item.agent);
      else emptyFloor.active.push(item.agent);
  });
  return JSON.stringify(emptyFloor);
}

function parseBreaksSafe(j) { try { return JSON.parse(j); } catch(e){return[];} }
function parseTimeValue(tStr) {
   if(!tStr) return 9999;
   const match = tStr.match(/(\d{1,2}):(\d{2})\s*([AP]M)/i);
   if(!match) return 9999;
   let h = parseInt(match[1]);
   let m = parseInt(match[2]);
   let amp = match[3].toUpperCase();
   if (amp === "PM" && h < 12) h += 12;
   if (amp === "AM" && h === 12) h = 0;
   return (h * 60) + m;
}
function toTitleCase(str) {
  return str.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
}
