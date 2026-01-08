/**
 * MODULE: AGENT MONITOR
 * Consumes Raw Schedule and StatusTracker to build the dashboard payload.
 */

const AgentMonitor = {
  getPayload: function() { return compileFloorData();
},
  // UPDATE: Use RoleManager for instant status updates
  setStatus: function(n, t, v) { return RoleManager.setStatus(n, t, v);
},
  updateAgentBreaks: function(n, j) { return StatusTracker.updateBreaks(n, j); },
  logOvertime: function(n, s, e, bs, be) { 
      const res = StatusTracker.logOvertime(n, s, e, bs, be);
      if(typeof NotificationHandler !== 'undefined') NotificationHandler.add(n, `OVERTIME: ${s}-${e}`);
      return res;
  },
  updateBreak: function() { return "Deprecated";
},
  clearFlag: function() { return "Flag Cleared"; }
};

/* --- PUBLIC ENDPOINTS --- */
function getFloorStatus() { return compileFloorData();
}
function updateAgentStatus(name, type, value) { return AgentMonitor.setStatus(name, type, value); }
function updateAgentBreaks(name, jsonBreaks) { return AgentMonitor.updateAgentBreaks(name, jsonBreaks);
}
function submitOvertime(name, start, end, bStart, bEnd) { return AgentMonitor.logOvertime(name, start, end, bStart, bEnd);
}

/* --- CORE LOGIC --- */
function compileFloorData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Raw Schedule");
  const emptyFloor = {
    active: [], startingSoon: [], safe: [], icl: [], training: [],
    upcomingBreak: [], vacation: [], planned: [], unplanned: [], off: []
  };
  // FETCH BOTH SOURCES
  const sheetOverrides = StatusTracker.getConsolidatedData(); 
  const fastOverrides = RoleManager.getFastMap();
  // Get instant cache

  const data = (sheet && sheet.getLastRow() > 1) ?
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
  
  const now = new Date().getTime();
  const agentMap = new Map();
  const SCORES = { 'active': 50, 'startingSoon': 45, 'training': 30, 'unplanned': 20, 'vacation': 10, 'planned': 10, 'off': 0 };
  const processEntry = (rawName, row) => {
    const cleanName = String(rawName).trim().toLowerCase();
    // MERGE LOGIC: Fast Cache > Sheet > Import
    const persistentData = sheetOverrides.get(cleanName) || {};
    const fastData = fastOverrides[cleanName] || {};
    
    // Prioritize Fast Cache if it exists
    let role = (fastData.role !== undefined) ?
    fastData.role : 
               (persistentData.role !== undefined && persistentData.role !== "") ?
    persistentData.role : 
               (row ? row[8] : "");
    let absentType = (fastData.absent !== undefined) ? fastData.absent :
                     (persistentData.absent !== undefined && persistentData.absent !== "") ?
    persistentData.absent :
                     (row ? row[9] : "");
    let shiftType = row ? row[5] : "Off";
    let region = row ? row[6] : "Offshore";
    let startEpoch = row ? Number(row[10]) : 0;
    let endEpoch = row ? Number(row[11]) : 0;
    let originalBreaksJson = row ? row[7] : "[]";
    
    if (role && role !== "") absentType = "";
    const otList = persistentData.ot || [];
    const customBreaks = persistentData.breaks;
    const isOT = otList.length > 0;
    // --- FILTERING ---
    if ((shiftType === "Off" || region === "Off") && !isOT) return;
    if ((!startEpoch || !endEpoch) && !isOT) return;

    if (startEpoch && endEpoch) {
       // FIX: Removed buffer so they disappear instantly when shift ends
       if (endEpoch <= now && !isOT) return;

       const NINETY_MINS = 90 * 60 * 1000;
       if (startEpoch > (now + NINETY_MINS) && !isOT) return;
    }

    let dateStr = row && row[2] instanceof Date ?
    Utilities.formatDate(row[2], "America/Toronto", "M/d/yy") : (row ? row[2] : "");
    let shiftStr = "";
    if(startEpoch && endEpoch) {
       const sFmt = Utilities.formatDate(new Date(startEpoch), "America/Toronto", "HH:mm");
       const eFmt = Utilities.formatDate(new Date(endEpoch), "America/Toronto", "HH:mm");
       shiftStr = `${sFmt} - ${eFmt}`;
    }

    let activeBreaksJson = customBreaks ? customBreaks : originalBreaksJson;
    const isModified = !!customBreaks;
    const rawBreaks = parseBreaksSafe(activeBreaksJson);
    if (isOT) {
       otList.forEach(o => {
           if(o.bStart && o.bStart !== "-" && o.bEnd && o.bEnd !== "-") {
               rawBreaks.push({ type: "OT Break", start: o.bStart, end: o.bEnd });
           }
       });
    }

    const enrichedBreaks = [];
    let nextBreakStr = null;
    let onBreakNow = false;
    let breakTimer = "";
    let breakLabel = "";

    if((startEpoch || isOT) && rawBreaks.length > 0) {
       const baseTime = startEpoch ?
       new Date(startEpoch) : new Date();
       const shiftDateStr = Utilities.formatDate(baseTime, "America/Toronto", "yyyy-MM-dd");
       rawBreaks.sort((a, b) => parseTimeValue(a.start) - parseTimeValue(b.start));
       let breakCounter = 0;
       
       rawBreaks.forEach(b => {
          let displayLabel = b.type;
          if (b.type === 'Break') {
             breakCounter++;
             displayLabel = `${breakCounter}${['st','nd','rd'][breakCounter-1]||'th'} Break`;
          }

          let bs = 0, be = 0;
          try {
 
             const bStart24 = Utilities.parseDate(`${shiftDateStr} ${b.start}`, "America/Toronto", "yyyy-MM-dd hh:mm a");
             const bEnd24 = Utilities.parseDate(`${shiftDateStr} ${b.end}`, "America/Toronto", "yyyy-MM-dd hh:mm a");
             if(bStart24 && bEnd24) {
                 bs = bStart24.getTime();
                 be = bEnd24.getTime();
     
             if (startEpoch && bs < startEpoch) { bs += 86400000; be += 86400000; }
                 if (be < bs) be += 86400000;
             }
          } catch(e) {}

          if(bs > 0) {
             enrichedBreaks.push({ 
                 type: displayLabel, originalType: b.type,
                 start: b.start, end: b.end, 
                 epochStart: bs, epochEnd: be 
             });
             if (now >= bs && now <= be) {
                 onBreakNow = true;
                 const rem = Math.ceil((be - now)/60000);
                 breakTimer = `${rem}m`;
                 breakLabel = displayLabel;
             }
             if (bs > now && !nextBreakStr && (bs - now < 1800000)) { 
                 nextBreakStr = `${b.start} - ${b.end}`;
             }
          }
       });
    }

    // --- CATEGORIZATION ---
    let category = "active";
    let subStatus = role ||
    ""; 

    if (absentType) {
        const upper = absentType.toUpperCase();
        if (upper.match(/NCNS|UNAB|AWOL|COMP/) && rawBreaks.length > 1) { category = "active";
        } 
        else if (upper.match(/SICK|NCNS|UNAB|AWOL|COMP/)) { category = "unplanned"; subStatus = absentType;
        }
        else if (upper.match(/TRAIN|TRN/)) { category = "training"; subStatus = "Training";
        }
        else { category = "vacation"; subStatus = absentType;
        }
    } else if (role === "Training") {
        category = "training";
    }

    // Starting Soon
    if (['active', 'safe', 'icl'].includes(category)) {
        if (startEpoch && startEpoch > now) {
             category = "startingSoon";
             const diff = startEpoch - now;
             subStatus = `In ${Math.ceil(diff/60000)}m`;
        }
    }

    if (category === "active" && onBreakNow) {
        subStatus = breakLabel ||
        "On Break";
    }

    let agent = {
      name: rawName, id: row ?
      row[1] : "", region: region, shift: shiftStr, shiftType: shiftType,
      dateStr: dateStr, subStatus: subStatus, role: role, rawBreaks: enrichedBreaks, 
      isModified: isModified, breakTimeStr: nextBreakStr, timer: breakTimer, auxLabel: "Remaining",
      onBreakNow: onBreakNow, startEpoch: startEpoch, isOT: isOT 
    };
    if (category === 'active' && nextBreakStr) {
       let upAgent = {...agent};
       upAgent.timer = nextBreakStr;
       upAgent.auxLabel = "Starts Soon";
       emptyFloor.upcomingBreak.push(upAgent);
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

  sheetOverrides.forEach((val, key) => {
      // FIX: Changed 'overrides' to 'sheetOverrides' which is the correct variable name in this scope
      if (!agentMap.has(key) && !agentMap.has(toTitleCase(key)) && val.ot.length > 0) {
          processEntry(toTitleCase(key), null); 
      }
  });
  
  agentMap.forEach(item => {
      const targetCat = item.category;
      if (emptyFloor[targetCat]) emptyFloor[targetCat].push(item.agent);
      else emptyFloor.active.push(item.agent);

      if (item.agent.role === 'SAFE') emptyFloor.safe.push(item.agent);
      if (item.agent.role === 'ICL') emptyFloor.icl.push(item.agent);
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
