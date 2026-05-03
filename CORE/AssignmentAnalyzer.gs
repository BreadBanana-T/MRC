/**
 * MODULE: ASSIGNMENT ANALYZER (RED FLAGS - LOCAL HOST ONLY)
 */

var AssignmentAnalyzer = {

  _getDB: function(sheetName) {
      return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  },

  importMasterList: function(rawText) {
      try {
          if (!rawText || rawText.trim() === "") return "Error: No data pasted.";
          let lines = rawText.split(/\r?\n/).filter(l => l.trim().length > 0);
          
          let headerIdx = lines.findIndex(l => l.toLowerCase().includes('employee_name') || l.toLowerCase().includes('job_title') || l.toLowerCase().includes('name_first'));
          if (headerIdx === -1) return "Error: Could not find MasterList headers.";
          
          let headers = this._parseCSVLine(lines[headerIdx]);
          let colName = headers.findIndex(h => h.toLowerCase().includes('employee_name') || h.toLowerCase() === 'name');
          let colJob = headers.findIndex(h => h.toLowerCase().includes('title') || h.toLowerCase().includes('job'));
          let colStatus = headers.findIndex(h => h.toLowerCase().includes('status'));
          let colLevel = headers.findIndex(h => h.toLowerCase().includes('erc_level') || h.toLowerCase().includes('level'));
          
          // NEW: Grab extra metadata
          let colSup = headers.findIndex(h => h.toLowerCase().includes('supervisor'));
          let colSkills = headers.findIndex(h => h.toLowerCase() === 'skills');
          let colLoc = headers.findIndex(h => h.toLowerCase().includes('location'));
          let colEmailTI = headers.findIndex(h => h.toLowerCase().includes('international'));

          let validAgents = [];
          for (let i = headerIdx + 1; i < lines.length; i++) {
              let cols = this._parseCSVLine(lines[i]);
              let name = cols[colName];
              let job = cols[colJob] || "";
              let status = cols[colStatus] || "";
              
              let level = colLevel > -1 ? parseInt(cols[colLevel]) || 2 : 2;
              let sup = colSup > -1 ? cols[colSup] : "";
              let skills = colSkills > -1 ? cols[colSkills] : "";
              let loc = colLoc > -1 ? cols[colLoc] : "";
              let emailTI = colEmailTI > -1 ? cols[colEmailTI] : "";

              let isManager = job.toLowerCase().includes('manager') || job.toLowerCase().includes('supervisor') || job.toLowerCase().includes('director');
              // Status field is rarely updated in practice — agents who've
              // been here a year may still show "Short term" or "Hired".
              // Switched from a positive "must say active" filter to a
              // negative "must NOT say terminated/etc" filter so outdated
              // status values don't silently drop real agents.
              let s = status.toLowerCase();
              let isInactive = s.includes('terminated') || s.includes('resigned') || s.includes('inactive') || s.includes('separated') || s === 'left';

              if (name && !isInactive && !isManager) {
                  validAgents.push([name.replace(/(^"|"$)/g, '').trim(), level, sup, skills, loc, emailTI]);
              }
          }
          
          if (validAgents.length > 0) {
              const ss = SpreadsheetApp.getActiveSpreadsheet();
              let sheet = ss.getSheetByName('WF_MASTERLIST');
              if (!sheet) sheet = ss.insertSheet('WF_MASTERLIST');
              sheet.clearContents();
              // Save the 6 columns
              sheet.appendRow(["Agent Name", "ERC Level", "Supervisor", "Skills", "Location", "Email_TI"]);
              sheet.getRange(2, 1, validAgents.length, 6).setValues(validAgents);

              // Sync to RegionRegistry with source=masterlist so these entries
              // outrank future WFM auto-detection but remain overridable manually.
              if (typeof RegionRegistry !== 'undefined') {
                  validAgents.forEach(function(row) {
                      var agentName = row[0];
                      var loc = String(row[4] || '').toUpperCase();
                      var emailTi = String(row[5] || '');
                      var isOffshore = loc.includes('EL SALVADOR') || loc.includes('GUATEMALA') || loc.startsWith('TI') || emailTi.includes('@');
                      RegionRegistry.upsert(agentName, isOffshore ? 'Offshore' : 'Onshore', 'masterlist');
                  });
              }
              return `Success: Locked ${validAgents.length} Active Agents into the engine.\nManagers are permanently ignored.`;
          }
          return "Error: No active agents found in the pasted list.";
      } catch (e) {
          return "Error parsing MasterList: " + e.message;
      }
  },

  importGEMData: function(rawText) {
      try {
          let map = this._parseGEM(rawText);
          if (Object.keys(map).length === 0) return JSON.stringify({error: "Invalid data format or missing GEM headers."});

          let repMonth = "";
          for (let k in map) { if (map[k].month) { repMonth = map[k].month; break; } }
          if (!repMonth) repMonth = Utilities.formatDate(new Date(), "America/Toronto", "yyyy-MM-01");
          
          this._saveGEMToDB(map, repMonth);
          return this.getAnalyzerData(repMonth);
      } catch (e) {
          return JSON.stringify({error: "Engine Error: " + e.message});
      }
  },

  excludeAgent: function(agentName) {
      try {
          const ss = SpreadsheetApp.getActiveSpreadsheet();
          let sheet = ss.getSheetByName('WF_EXCLUSIONS');
          if (!sheet) { sheet = ss.insertSheet('WF_EXCLUSIONS'); sheet.appendRow(["Excluded Agents"]); }
          
          let existing = sheet.getDataRange().getValues().flat();
          if (!existing.includes(agentName)) sheet.appendRow([agentName]);
          
          return "Agent Dismissed";
      } catch (e) { return "Error dismissing agent."; }
  },

  getAnalyzerData: function(targetMonth) {
      try {
          let db = this._getDB('WF_GEM_DATA_V3');
          if (!db || db.getLastRow() < 2) {
              db = this._getDB('WF_GEM_DATA_V2');
              if (!db || db.getLastRow() < 2) return JSON.stringify({ error: "No GEM data available. Please import your report.", months: [] });
          }
          
          let validAgents = new Map();
          let mlSheet = this._getDB('WF_MASTERLIST');
          let useMasterList = false;
          
          if (mlSheet && mlSheet.getLastRow() > 1) {
              useMasterList = true;
              mlSheet.getDataRange().getDisplayValues().slice(1).forEach(r => {
                  let lvl = parseInt(r[1]);
                  if (isNaN(lvl)) lvl = 2;
                  let cleanName = _normalizeAgentKey(r[0]);
                  validAgents.set(cleanName, lvl);
              });
          }

          let exclusions = new Set();
          let exclSheet = this._getDB('WF_EXCLUSIONS');
          if (exclSheet && exclSheet.getLastRow() > 1) {
              exclSheet.getDataRange().getValues().slice(1).forEach(r => exclusions.add(String(r[0]).trim()));
          }

          let data = db.getDataRange().getValues().slice(1);
          let validData = [];
          data.forEach(r => {
              let dateStr = r[0] instanceof Date ? Utilities.formatDate(r[0], "America/Toronto", "yyyy-MM-dd") : String(r[0]);
              if (dateStr.match(/^20\d{2}-\d{2}-\d{2}/)) {
                  r[0] = dateStr;
                  validData.push(r);
              }
          });

          if (validData.length === 0) return JSON.stringify({ error: "Database was corrupted. Please re-import your GEM report.", months: [] });
          let monthSet = new Set();
          validData.forEach(r => monthSet.add(r[0]));
          let months = Array.from(monthSet).sort().reverse();
          if (!targetMonth || !monthSet.has(targetMonth)) {
              targetMonth = months.length > 0 ? months[0] : null;
          }

          let prevMonth = null;
          let tIdx = months.indexOf(targetMonth);
          if (tIdx !== -1 && tIdx < months.length - 1) {
              prevMonth = months[tIdx + 1];
          }

          const fixTime = (v) => {
              if (!v || v === '-') return '-';
              let s = String(v).trim();
              if (s.startsWith("'")) s = s.substring(1);
              let m = s.match(/T(\d{2}:\d{2})/);
              if (m) return m[1];
              if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "HH:mm");
              return s;
          };

          let currentMap = {};
          let prevMap = {};
          validData.forEach(r => {
              let agent = String(r[1]).trim();
              let cleanAgent = _normalizeAgentKey(agent);
              let agentLevel = 2;

              if (useMasterList) {
                  let matchedLvl = null;
                  if (validAgents.has(cleanAgent)) {
                      matchedLvl = validAgents.get(cleanAgent);
                  } else {
                      let gnParts = cleanAgent.replace(/,/g, ' ').split(' ').filter(x => x.length > 1);
                      for (let [mlName, lvl] of validAgents.entries()) {
                          let mlParts = mlName.replace(/,/g, ' ').split(' ').filter(x => x.length > 1);
                          if (gnParts.length > 1 && mlParts.length > 1 && gnParts[0] === mlParts[0] && gnParts[1] === mlParts[1]) {
                              matchedLvl = lvl; break;
                          }
                          let minLen = Math.min(cleanAgent.length, mlName.length);
                          if (minLen >= 8 && (cleanAgent.includes(mlName) || mlName.includes(cleanAgent))) {
                              matchedLvl = lvl; break;
                          }
                      }
                  }
                  // No MasterList match — pass through with default level 2
                  // instead of dropping the agent. The MasterList stays useful
                  // as a manager filter (managers never enter it in the first
                  // place via the job-title check at import), but isn't a
                  // hard gate anymore: agents missing because the MasterList
                  // is stale still appear in Red Flags.
                  agentLevel = (matchedLvl === null) ? 2 : matchedLvl;
              }
              
              if (exclusions.has(agent)) return;
              let inP = parseFloat(r[3]) || 0; if (inP > 1) inP = inP / 100;
              let outP = parseFloat(r[4]) || 0; if (outP > 1) outP = outP / 100;
              let cphV = parseFloat(r[2]) || 0;
              
              if (r[0] === targetMonth) {
                  currentMap[agent] = {
                      level: agentLevel,
                      cph: cphV, inPct: inP, outPct: outP, aht: fixTime(r[5]), 
                      tasksXCalls: parseFloat(r[6]) || 0, totalCalls: parseInt(r[7]) || 0, 
                      callsAnsw: parseInt(r[8]) || 0, wDays: parseInt(r[9]) || 0
                  };
              } else if (r[0] === prevMonth) {
                  prevMap[agent] = { cph: cphV, inPct: inP, outPct: outP };
              }
          });
          
          let inboundList = [];
          let outOfScopeList = [];
          let outboundList = [];
          let cphList = [];
          
          Object.keys(currentMap).forEach(agent => {
              let c = currentMap[agent];
              let p = prevMap[agent] || null;
              
              if (c.wDays < 5) return; 

              let trendIn = p ? (c.inPct - p.inPct) : 0;
              let trendOut = p ? (c.outPct - p.outPct) : 0;
              let trendCPH = p ? (c.cph - p.cph) : 0;

              let payload = {
                  name: agent, wDays: c.wDays,
                  inPct: c.inPct, prevIn: p ? p.inPct : null, trendIn: trendIn,
                  outPct: c.outPct, prevOut: p ? p.outPct : null, trendOut: trendOut,
                  cph: c.cph, prevCph: p ? p.cph : null, trendCph: trendCPH,
                  aht: c.aht, tasksXCalls: c.tasksXCalls, totalCalls: c.totalCalls
              };

              // Mirrors the Assignment Tracker 2026 RedFlag dashboard rule:
              //   Inbound Red Flags  = inPct >= 75% AND tasksXCalls >= 50%
              //   Inbound Out of Scope = inPct >= 75% AND tasksXCalls < 50%
              //                         (mutually exclusive — Excel's COUNTIF
              //                          in F6 explicitly excludes Red Flag
              //                          agents from the Out of Scope list)
              // wDays >= 5 already enforced at the top of this loop.
              if (c.inPct >= 0.75) {
                  if (c.tasksXCalls >= 0.50) inboundList.push(payload);
                  else outOfScopeList.push(payload);
              }
              if (c.outPct >= 0.75 && c.level >= 2) outboundList.push(payload);
              if (c.cph > 0 && c.cph <= 7.3) cphList.push(payload);
          });
          
          inboundList.sort((a,b) => b.inPct - a.inPct);
          outOfScopeList.sort((a,b) => b.inPct - a.inPct);
          outboundList.sort((a,b) => b.outPct - a.outPct);
          cphList.sort((a,b) => a.cph - b.cph);
          
          return JSON.stringify({
              selectedMonth: targetMonth, availableMonths: months, usingMasterList: useMasterList,
              inbound: inboundList, outOfScope: outOfScopeList, outbound: outboundList, cph: cphList
          });
      } catch (e) {
          return JSON.stringify({ error: "Data generation failed: " + e.message });
      }
  },

  _parseGEM: function(raw) {
      if (!raw || raw.trim() === "") return {};
      let lines = raw.split(/\r?\n/).filter(l => l.trim().length > 0);
      let headerIdx = lines.findIndex(l => l.toLowerCase().includes('in & out by hr') || l.toLowerCase().includes('cph'));
      if (headerIdx === -1) return {};
      
      let headers = this._parseCSVLine(lines[headerIdx]);
      let colMap = {};
      headers.forEach((h, i) => {
          let low = h.toLowerCase();
          if (low.includes('agent')) colMap.agent = i;
          else if (low.includes('month of ref') || low.includes('date')) colMap.month = i;
          else if (low.includes('in & out by hr') || low.includes('cph')) colMap.cph = i;
          else if (low.includes('total all calls') || low.includes('total tasks')) colMap.totalCalls = i;
          else if (low.includes('inbound time %')) colMap.inPct = i;
          else if (low.includes('outbound time %')) colMap.outPct = i;
          else if (low.includes('nb calls answ') || low.includes('calls answered')) colMap.callsAnsw = i;
          else if (low.includes('worked days') || low.includes('wdays')) colMap.wDays = i;
          else if (low.includes('aht')) colMap.aht = i;
          else if (low.includes('nb transfers')) colMap.transfers = i;
          else if (low.includes('% transf')) colMap.transfPct = i;
          else if (low.includes('avg talk time out')) colMap.outTalk = i;
          else if (low.includes('avg talk time ext')) colMap.extTalk = i;
          else if (low.includes('avg talk time')) colMap.avgTalk = i;
          else if (low.includes('avg hold')) colMap.avgHold = i;
          else if (low.includes('avg acw total') || low.includes('avg acw')) colMap.avgAcw = i;
          else if (low.includes('outbound calls')) colMap.outCalls = i;
          else if (low.includes('extension in calls') || low.includes('ext in calls')) colMap.extCalls = i;
      });
      
      let map = {};
      for (let i = headerIdx + 1; i < lines.length; i++) {
          let cols = this._parseCSVLine(lines[i]);
          let agent = cols[colMap.agent];
          if (!agent || agent.toLowerCase().includes('total')) continue;
          agent = _titleCaseName(agent.replace(/"/g, '').trim());

          let tCalls = parseInt(String(cols[colMap.totalCalls]).replace(/,/g, '')) || 0;
          let cAnsw = parseInt(String(cols[colMap.callsAnsw]).replace(/,/g, '')) || 0;
          let txc = tCalls > 0 ? (cAnsw / tCalls) : 0;
          let wD = parseInt(String(cols[colMap.wDays]).replace(/,/g, '')) || 0;

          let rawMonth = cols[colMap.month] || "";
          let cleanMonth = "";
          let mMatch = String(rawMonth).match(/(\d{4}-\d{2}-\d{2})/);
          if (mMatch) cleanMonth = mMatch[1];

          map[agent] = {
              month: cleanMonth,
              cph: this._parseSafeFloat(cols[colMap.cph]),
              inPct: this._parseSafePercent(cols[colMap.inPct]),
              outPct: this._parseSafePercent(cols[colMap.outPct]),
              totalCalls: tCalls,
              callsAnsw: cAnsw,
              tasksXCalls: txc,
              wDays: wD,
              aht: this._parseSafeAHT(cols[colMap.aht]),
              transfers: parseInt(String(cols[colMap.transfers] || 0).replace(/,/g, '')),
              transfPct: this._parseSafePercent(cols[colMap.transfPct]),
              avgTalk: this._parseSafeAHT(cols[colMap.avgTalk]),
              avgHold: this._parseSafeAHT(cols[colMap.avgHold]),
              avgAcw: this._parseSafeAHT(cols[colMap.avgAcw]),
              outCalls: parseInt(String(cols[colMap.outCalls] || 0).replace(/,/g, '')),
              outTalk: this._parseSafeAHT(cols[colMap.outTalk]),
              extCalls: parseInt(String(cols[colMap.extCalls] || 0).replace(/,/g, '')),
              extTalk: this._parseSafeAHT(cols[colMap.extTalk])
          };
      }
      return map;
  },

  _parseSafePercent: function(val) {
      if (!val) return 0;
      let s = String(val).replace(/\s/g, '').trim().replace(/,/g, '.');
      let f = parseFloat(s.replace('%', ''));
      if (isNaN(f)) return 0;
      if (s.includes('%')) return f / 100;
      if (f > 1) return f / 100; 
      return f;
  },

  _parseSafeFloat: function(val) {
      if (!val) return 0;
      let f = parseFloat(String(val).replace(/\s/g, '').trim().replace(/,/g, '.'));
      return isNaN(f) ? 0 : f;
  },

  _parseSafeAHT: function(val) {
      if (!val) return "00:00";
      let s = String(val).trim();
      if (s.includes(':')) {
          let parts = s.split(':');
          if (parts.length >= 3) return parts[1] + ":" + parts[2]; 
          return s;
      }
      let f = parseFloat(s.replace(/,/g, '.'));
      if (!isNaN(f) && f < 1) { 
          let totalSecs = Math.round(f * 86400);
          let m = Math.floor(totalSecs / 60);
          let sc = totalSecs % 60;
          return (m < 10 ? '0'+m : m) + ":" + (sc < 10 ? '0'+sc : sc);
      }
      return s;
  },

  _saveGEMToDB: function(gemMap, repMonth) {
      let sheetName = 'WF_GEM_DATA_V3';
      const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
      
      const writeToSheet = (targetSS) => {
          let sheet = targetSS.getSheetByName(sheetName);
          let expectedHeaders = ["Month", "Agent Name", "CPH", "Inbound %", "Outbound %", "AHT", "Tasks x Calls", "Total Tasks", "Calls Answered", "WDays", "Transfers", "Transf %", "Avg Talk", "Avg Hold", "Avg ACW", "Out Calls", "Out Talk", "Ext Calls", "Ext Talk"];
          if (!sheet) {
              sheet = targetSS.insertSheet(sheetName);
              sheet.appendRow(expectedHeaders);
          }
          
          let existing = sheet.getDataRange().getValues();
          if (existing.length === 1 && existing[0].join('') === "") existing = [];
          
          let headers = existing.length > 0 ? existing.shift() : expectedHeaders;
          
          let retained = existing.filter(row => {
              let dStr = row[0] instanceof Date ? Utilities.formatDate(row[0], "America/Toronto", "yyyy-MM-dd") : String(row[0]);
              return dStr !== repMonth;
          });

          let safeT = function(t) { return (t && t !== '-' && t !== '00:00') ? "'" + t : "-"; };
          
          let newRows = [];
          Object.keys(gemMap).forEach(agent => {
              let m = gemMap[agent];
              newRows.push([
                  repMonth, agent, m.cph, m.inPct, m.outPct, safeT(m.aht), m.tasksXCalls, m.totalCalls, m.callsAnsw, m.wDays,
                  m.transfers, m.transfPct, safeT(m.avgTalk), safeT(m.avgHold), safeT(m.avgAcw), m.outCalls, safeT(m.outTalk), m.extCalls, safeT(m.extTalk)
              ]);
          });
          
          let combined = retained.concat(newRows);
          sheet.clearContents();
          sheet.appendRow(headers);
          if (combined.length > 0) sheet.getRange(2, 1, combined.length, 19).setValues(combined);
      };

      writeToSheet(ssLocal);
  },

  _parseCSVLine: function(text) {
    if (text.includes('\t')) return text.split('\t').map(s => s.replace(/(^"|"$)/g, '').trim());
    let ret = [], inQuote = false, token = "";
    for(let i=0; i<text.length; i++) {
      let char = text[i];
      if(char === '"') { inQuote = !inQuote; continue; }
      if(char === ',' && !inQuote) { ret.push(token.trim()); token = ""; } else token += char;
    }
    ret.push(token.trim()); return ret;
  }
};

function importNewGEMReport(rawText) { return AssignmentAnalyzer.importGEMData(rawText); }
function fetchAnalyzerData(monthStr) { return AssignmentAnalyzer.getAnalyzerData(monthStr); }
function processMasterListUpload(rawText) { return AssignmentAnalyzer.importMasterList(rawText); }
function ignoreAnalyzerAgent(agentName) { return AssignmentAnalyzer.excludeAgent(agentName); }

/**
 * Coaching cadence red flag.
 * Scans WF_COACHING for each agent's most recent session. Any agent whose
 * latest is older than `thresholdDays` (default 30) OR has zero sessions in
 * the last 90 days is flagged.
 *
 * Roster-agnostic: builds the active-agent list from WF_ROLES + WF_COACHING
 * + WF_ABSENCES presence in the last 60 days (so offshore agents that
 * aren't in MasterList still get flagged). Region is resolved via
 * RegionRegistry when available.
 *
 * Exclusions list (WF_EXCLUSIONS) is honored.
 */
function getCoachingCadenceFlags(thresholdDays) {
  try {
    thresholdDays = thresholdDays || 30;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();
    var nowMs = now.getTime();
    var activityWindowMs = 60 * 86400000;
    var dayMs = 86400000;

    var normKey = function(s) {
      return (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey(s) : String(s || '').trim().toLowerCase();
    };

    // Build active-agent registry from WF_ROLES (recent 60d activity).
    var agents = {};
    var sheets = ['WF_ROLES', 'WF_COACHING', 'WF_ABSENCES'];
    sheets.forEach(function(name) {
      var s = ss.getSheetByName(name);
      if (!s || s.getLastRow() < 2) return;
      var data = s.getRange(2, 1, s.getLastRow() - 1, 6).getDisplayValues();
      data.forEach(function(row) {
        var agent = String(row[0] || '').trim();
        if (!agent) return;
        var dStr = String(row[1] || '');
        if (!dStr.match(/^\d{4}-\d{2}-\d{2}/)) return;
        var ts = new Date(dStr + 'T12:00:00').getTime();
        if (isNaN(ts) || (nowMs - ts) > activityWindowMs) return;
        var k = normKey(agent);
        if (!agents[k]) agents[k] = { name: agent, lastCoaching: null, coachingCount90d: 0, region: null };
        if (!agents[k].region && typeof RegionRegistry !== 'undefined') {
          agents[k].region = RegionRegistry.getRegion(agent);
        }
        if (!agents[k].region) {
          var reg = String(row[5] || '').trim();
          agents[k].region = (reg === 'Offshore') ? 'Offshore' : 'Onshore';
        }
      });
    });

    // Walk WF_COACHING, tracking latest + 90-day count per agent.
    var dbCoach = ss.getSheetByName('WF_COACHING');
    if (dbCoach && dbCoach.getLastRow() > 1) {
      var data = dbCoach.getRange(2, 1, dbCoach.getLastRow() - 1, 6).getDisplayValues();
      data.forEach(function(row) {
        var agent = String(row[0] || '').trim();
        if (!agent) return;
        var dStr = String(row[1] || '');
        if (!dStr.match(/^\d{4}-\d{2}-\d{2}/)) return;
        var ts = new Date(dStr + 'T12:00:00').getTime();
        if (isNaN(ts)) return;
        var k = normKey(agent);
        if (!agents[k]) return; // not in active window; skip
        if (agents[k].lastCoaching === null || ts > agents[k].lastCoaching) {
          agents[k].lastCoaching = ts;
        }
        if ((nowMs - ts) <= 90 * dayMs) agents[k].coachingCount90d++;
      });
    }

    // Exclusions
    var excluded = {};
    var exSheet = ss.getSheetByName('WF_EXCLUSIONS');
    if (exSheet && exSheet.getLastRow() > 1) {
      exSheet.getRange(2, 1, exSheet.getLastRow() - 1, 1).getValues().forEach(function(r) {
        if (r[0]) excluded[normKey(r[0])] = true;
      });
    }

    var flags = [];
    Object.keys(agents).forEach(function(k) {
      if (excluded[k]) return;
      var a = agents[k];
      var daysSince = a.lastCoaching ? Math.floor((nowMs - a.lastCoaching) / dayMs) : null;
      var flaggedReason = null;
      if (a.lastCoaching === null) flaggedReason = 'No session on record (60d activity)';
      else if (daysSince >= thresholdDays) flaggedReason = daysSince + ' days since last session';

      if (flaggedReason) {
        flags.push({
          name: a.name,
          region: a.region || 'Onshore',
          lastCoachingDate: a.lastCoaching ? Utilities.formatDate(new Date(a.lastCoaching), 'America/Toronto', 'yyyy-MM-dd') : null,
          daysSince: daysSince,
          coachingCount90d: a.coachingCount90d,
          reason: flaggedReason
        });
      }
    });
    flags.sort(function(a, b) {
      if (a.daysSince === null) return -1;
      if (b.daysSince === null) return 1;
      return b.daysSince - a.daysSince;
    });

    return JSON.stringify({ threshold: thresholdDays, flags: flags, totalActive: Object.keys(agents).length });
  } catch (e) {
    return JSON.stringify({ flags: [], error: e.message });
  }
}

/**
 * One-shot diagnostic: for a given month, report which agents are present
 * in the GEM analyzer logs but missing from WF_MASTERLIST (and therefore
 * silently dropped by the Red Flag categorizer at line ~194). Run from
 * the editor: select `debugMissingFromMasterList` → Run.
 */
function debugMissingFromMasterList(targetMonth) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Build masterlist key set
  var ml = ss.getSheetByName('WF_MASTERLIST');
  var keys = new Set();
  if (ml && ml.getLastRow() > 1) {
    ml.getDataRange().getDisplayValues().slice(1).forEach(function(r) {
      var k = (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey(r[0]) : String(r[0] || '').trim().toLowerCase();
      if (k) keys.add(k);
    });
  }
  Logger.log('MasterList keys loaded: ' + keys.size);

  // Walk analyzer log sheet (assignment analyzer logs the GEM data here)
  var ag = ss.getSheetByName('WF_GEM_DATA_V3') || ss.getSheetByName('Analyzer_Logs') || ss.getSheetByName('GEM_Logs');
  if (!ag) {
    Logger.log('No analyzer log sheet found. Looked for: WF_GEM_DATA_V3, Analyzer_Logs, GEM_Logs');
    Logger.log('Available sheets: ' + ss.getSheets().map(function(s) { return s.getName(); }).join(', '));
    return;
  }
  Logger.log('Reading from sheet: ' + ag.getName());
  var rows = ag.getDataRange().getDisplayValues();
  var header = rows[0];
  var data = rows.slice(1);
  Logger.log('Total rows: ' + data.length);

  if (!targetMonth) {
    // Use the latest month present
    var months = {};
    data.forEach(function(r) { var m = String(r[0]); if (m) months[m] = (months[m] || 0) + 1; });
    targetMonth = Object.keys(months).sort().pop();
  }
  Logger.log('Target month: ' + targetMonth);

  var presentInGEM = 0, missingFromML = [];
  data.forEach(function(r) {
    if (String(r[0]) !== targetMonth) return;
    presentInGEM++;
    var agent = String(r[1] || '').trim();
    if (!agent) return;
    var key = (typeof _normalizeAgentKey === 'function') ? _normalizeAgentKey(agent) : agent.toLowerCase();
    if (!keys.has(key)) {
      missingFromML.push({ name: agent, key: key });
    }
  });

  Logger.log('Agents in GEM for ' + targetMonth + ': ' + presentInGEM);
  Logger.log('Missing from MasterList: ' + missingFromML.length);
  missingFromML.slice(0, 50).forEach(function(m) {
    Logger.log('  - ' + m.name + '   (key="' + m.key + '")');
  });
  if (missingFromML.length > 50) Logger.log('  ... ' + (missingFromML.length - 50) + ' more');
}
