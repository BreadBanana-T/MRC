/**
 * MODULE: AGENT MONITOR (V4.6)
 * - Rock Solid Date Math
 * - Intraday ACSU Snapping
 */

const AgentMonitor = {
  getPayload: function() { return compileFloorData(); },
  setStatus: function(n, t, v) { return RoleManager.setStatus(n, t, v); },
  updateAgentBreaks: function(n, j) { return StatusTracker.updateBreaks(n, j); },
  logOvertime: function(n, s, e, bs, be) { 
      const res = StatusTracker.logOvertime(n, s, e, bs, be);
      if(typeof NotificationHandler !== 'undefined') NotificationHandler.add(n, `OVERTIME: ${s}-${e}`);
      return res;
  },
  syncFromRaw: function() { return "Synced"; }
};

function getFloorStatus() { return compileFloorData(); }
function updateAgentStatus(name, type, value) { return AgentMonitor.setStatus(name, type, value); }
function updateAgentBreaks(name, jsonBreaks) { return AgentMonitor.updateAgentBreaks(name, jsonBreaks); }
function submitOvertime(name, start, end, bStart, bEnd) { return AgentMonitor.logOvertime(name, start, end, bStart, bEnd); }

function compileFloorData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Raw Schedule");
  const emptyFloor = {
    active: [], startingSoon: [], safe: [], icl: [], ulc: [], training: [],
    upcomingBreak: [], vacation: [], planned: [], unplanned: [], off: []
  };

  const sheetOverrides = (typeof StatusTracker !== 'undefined') ? StatusTracker.getConsolidatedData() : new Map(); 
  const fastOverrides = (typeof RoleManager !== 'undefined') ? RoleManager.getFastMap() : {};
  
  const data = (sheet && sheet.getLastRow() > 1) ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
  const now = new Date().getTime();
  const agentMap = new Map();

  const SCORES = { 'active': 50, 'startingSoon': 45, 'training': 30, 'unplanned': 20, 'vacation': 10, 'planned': 10, 'off': 0 };

  const processEntry = (rawName, row) => {
    const cleanName = String(rawName).trim().toLowerCase();
    const persistentData = sheetOverrides.get(cleanName) || {};
    const fastData = fastOverrides[cleanName] || {};
    
    let role = (fastData.role !== undefined) ? fastData.role : (persistentData.role || (row ? row[8] : ""));
    let absentType = (fastData.absent !== undefined) ? fastData.absent : (persistentData.absent || (row ? row[9] : ""));
    let shiftType = row ? row[5] : "Off";
    let region = row ? row[6] : "Offshore";
    let startEpoch = row ? Number(row[10]) : 0;
    let endEpoch = row ? Number(row[11]) : 0;
    let originalBreaksJson = row ? row[7] : "[]";
    let agentID = row ? row[1] : "";

    if (role && role !== "") absentType = "";
    const otList = persistentData.ot || [];
    const customBreaks = persistentData.breaks;
    const isOT = otList.length > 0;
    let dateStr = row ? row[2] : "";
    let shiftStr = "";

    if (isOT && (!startEpoch || !endEpoch)) {
       let today = new Date();
       let otStart = otList[0].start;
       let otEnd = otList[0].end;
       
       let parseOTTime = (timeStr, baseDate) => {
           let d = new Date(baseDate.getTime());
           let parts = timeStr.match(/(\d{1,2}):(\d{2})/);
           if (parts) d.setHours(parseInt(parts[1]), parseInt(parts[2]), 0, 0);
           return d;
       };
       
       let sDate = parseOTTime(otStart, today);
       let eDate = parseOTTime(otEnd, today);
       if (eDate < sDate) eDate.setDate(eDate.getDate() + 1);
       
       startEpoch = sDate.getTime();
       endEpoch = eDate.getTime();
       shiftType = "OT";
       if (!dateStr) dateStr = Utilities.formatDate(today, "America/Toronto", "yyyy-MM-dd");
    }

    const LOOKAHEAD = 60 * 60 * 1000;
    if ((shiftType === "Off" || region === "Off") && !isOT) return;
    if ((!startEpoch || !endEpoch) && !isOT) return;
    if (startEpoch && endEpoch) {
       if (endEpoch <= now && !isOT) return;
       if (startEpoch > (now + LOOKAHEAD) && !isOT) return;
    }

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
           if(o.bStart && o.bStart !== "-" && o.bEnd && o.bEnd !== "-") rawBreaks.push({ type: "OT Break", start: o.bStart, end: o.bEnd });
       });
    }

    const enrichedBreaks = [];
    let nextBreakStr = null;
    let currentBreakStr = null;
    let onBreakNow = false;
    let inTrainingNow = false;
    let onAcsuNow = false;
    let breakTimer = "";
    let activeIntradayLabel = "";
    let breakStartsIn = "";

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
          let bs = 0, be = 0;
          try {
             let pStart = parseTime(b.start);
             let pEnd = parseTime(b.end);
             if(pStart && pEnd) {
                 let [y, mo, d] = shiftDateStr.split('-').map(Number);
                 let bDateStart = new Date(y, mo - 1, d, pStart.h, pStart.m, 0);
                 let bDateEnd = new Date(y, mo - 1, d, pEnd.h, pEnd.m, 0);
                 bs = bDateStart.getTime();
                 be = bDateEnd.getTime();
                 if (startEpoch && bs < startEpoch) { bs += 86400000; be += 86400000; }
                 if (be < bs) be += 86400000;
             }
          } catch(e) {}

          if(bs > 0) {
             enrichedBreaks.push({ type: displayLabel, start: b.start, end: b.end, epochStart: bs, epochEnd: be });
             if (now >= bs && now <= be) {
                 if (b.type === "Training") {
                     inTrainingNow = true;
                     activeIntradayLabel = displayLabel;
                 } else if (b.type === "ACSU") {
                     onAcsuNow = true;
                     activeIntradayLabel = "Furlough (ACSU)";
                 } else {
                     onBreakNow = true;
                     activeIntradayLabel = displayLabel;
                 }
                 breakTimer = `${Math.ceil((be - now)/60000)}m`;
                 currentBreakStr = `${b.start} - ${b.end}`;
             }
             if (bs > now && !nextBreakStr && (bs - now < 1800000) && b.type !== "Training" && b.type !== "ACSU") { 
                 nextBreakStr = `${b.start} - ${b.end}`;
                 breakStartsIn = Math.ceil((bs - now)/60000) + "m";
             }
          }
       });
    }

    let category = "active";
    let subStatus = role || "";

    if (absentType) {
        const upper = absentType.toUpperCase();
        if (upper.match(/NCNS|UNAB|AWOL|COMP/) && rawBreaks.length > 1) category = "active";
        else if (upper.match(/SICK|NCNS|UNAB|AWOL|COMP/)) { category = "unplanned"; subStatus = absentType; }
        else { category = "vacation"; subStatus = absentType; }
    } 

    if (['active', 'safe', 'icl', 'ulc'].includes(category)) {
        if (startEpoch && startEpoch > now) {
             category = "startingSoon";
             const diff = startEpoch - now;
             subStatus = `In ${Math.ceil(diff/60000)}m`;
        }
    }

    // Live Intraday Overrides (This moves them dynamically out of Active!)
    if (category === "active") {
        if (onAcsuNow) { category = "unplanned"; subStatus = activeIntradayLabel; }
        else if (inTrainingNow) { category = "training"; subStatus = activeIntradayLabel || "Training"; }
        else if (onBreakNow) subStatus = activeIntradayLabel;
    }

    let agent = {
      name: rawName, id: agentID, region: region, shift: shiftStr, shiftType: shiftType,
      dateStr: dateStr, subStatus: subStatus, role: role, rawBreaks: enrichedBreaks, 
      isModified: isModified, breakTimeStr: nextBreakStr, currentBreakStr: currentBreakStr, 
      timer: breakTimer, auxLabel: "Remaining", startsIn: breakStartsIn,
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
      if (!agentMap.has(key) && !agentMap.has(toTitleCase(key)) && val.ot.length > 0) processEntry(toTitleCase(key), null); 
  });

  agentMap.forEach(item => {
      const targetCat = item.category;
      if (emptyFloor[targetCat]) emptyFloor[targetCat].push(item.agent);
      else emptyFloor.active.push(item.agent);
      
      const r = (item.agent.role || "").toUpperCase();
      if (r.includes('SAFE')) emptyFloor.safe.push(item.agent);
      if (r.includes('ICL')) emptyFloor.icl.push(item.agent);
      if (r.includes('ULC') || r.includes('FIRE')) emptyFloor.ulc.push(item.agent);
  });

  return JSON.stringify(emptyFloor);
}

function parseBreaksSafe(j) { try { return JSON.parse(j); } catch(e){return[];} }
function parseTimeValue(tStr) {
   if(!tStr) return 9999;
   const match = tStr.match(/(\d{1,2}):(\d{2})\s*([AP]M)?/i);
   if(!match) return 9999;
   let h = parseInt(match[1]); let m = parseInt(match[2]); let amp = match[3] ? match[3].toUpperCase() : null;
   if (amp === "PM" && h < 12) h += 12;
   if (amp === "AM" && h === 12) h = 0;
   return (h * 60) + m;
}
function parseTime(tStr) {
   if(!tStr) return null; const match = tStr.match(/(\d{1,2}):(\d{2})\s*([AP]M)?/i); if(!match) return null;
   let h = parseInt(match[1]), m = parseInt(match[2]), amp = match[3] ? match[3].toUpperCase() : null;
   if (amp === "PM" && h < 12) h += 12;
   if (amp === "AM" && h === 12) h = 0; return { h, m };
}
function toTitleCase(str) { return str.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();}); }
