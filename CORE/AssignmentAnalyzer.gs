/**
 * MODULE: ASSIGNMENT ANALYZER (RED FLAGS) V5.4
 * Features: Exact Excel Math (70%, 75%, < 7.3), Fuzzy Name Matching, Level Failsafes
 */

var AssignmentAnalyzer = {

  _getDB: function(sheetName) {
      const local = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (local && local.getLastRow() > 1) return local;

      if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
          try { 
              const mSheet = SpreadsheetApp.openById(MasterConnector.DB_ID).getSheetByName(sheetName);
              if (mSheet && mSheet.getLastRow() > 1) return mSheet;
          } catch(e) {}
      }
      return local;
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

          let validAgents = [];
          for (let i = headerIdx + 1; i < lines.length; i++) {
              let cols = this._parseCSVLine(lines[i]);
              let name = cols[colName];
              let job = cols[colJob] || "";
              let status = cols[colStatus] || "";
              
              // Safely default to Level 2 if the column is empty/missing so they aren't hidden from Outbound
              let level = colLevel > -1 ? parseInt(cols[colLevel]) || 2 : 2;

              if (name && job.toLowerCase().includes('monitoring') && status.toLowerCase().includes('active')) {
                  validAgents.push([name.replace(/(^"|"$)/g, '').trim(), level]);
              }
          }
          
          if (validAgents.length > 0) {
              const ss = SpreadsheetApp.getActiveSpreadsheet();
              let sheet = ss.getSheetByName('WF_MASTERLIST');
              if (!sheet) sheet = ss.insertSheet('WF_MASTERLIST');
              sheet.clearContents();
              sheet.appendRow(["Agent Name", "ERC Level"]); 
              sheet.getRange(2, 1, validAgents.length, 2).setValues(validAgents);
              
              if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
                  try {
                      let mSheet = SpreadsheetApp.openById(MasterConnector.DB_ID).getSheetByName('WF_MASTERLIST');
                      if (!mSheet) mSheet = SpreadsheetApp.openById(MasterConnector.DB_ID).insertSheet('WF_MASTERLIST');
                      mSheet.clearContents();
                      mSheet.appendRow(["Agent Name", "ERC Level"]); 
                      mSheet.getRange(2, 1, validAgents.length, 2).setValues(validAgents);
                  } catch(e) {}
              }
              return `Success: Locked ${validAgents.length} Monitoring Agents into the engine.\nManagers will now be ignored.`;
          }
          return "Error: No active monitoring agents found in the pasted list.";
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
          
          if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
              try {
                  let mSheet = SpreadsheetApp.openById(MasterConnector.DB_ID).getSheetByName('WF_EXCLUSIONS');
                  if (!mSheet) { mSheet = SpreadsheetApp.openById(MasterConnector.DB_ID).insertSheet('WF_EXCLUSIONS'); mSheet.appendRow(["Excluded Agents"]); }
                  mSheet.appendRow([agentName]);
              } catch(e) {}
          }
          return "Agent Dismissed";
      } catch (e) { return "Error dismissing agent."; }
  },

  getAnalyzerData: function(targetMonth) {
      try {
          let db = this._getDB('WF_GEM_DATA_V2');
          if (!db || db.getLastRow() < 2) return JSON.stringify({ error: "No GEM data available. Please import your report.", months: [] });
          
          let validAgents = new Map();
          let mlSheet = this._getDB('WF_MASTERLIST');
          let useMasterList = false;
          
          if (mlSheet && mlSheet.getLastRow() > 1) {
              useMasterList = true;
              mlSheet.getDataRange().getDisplayValues().slice(1).forEach(r => {
                  let lvl = parseInt(r[1]);
                  if (isNaN(lvl)) lvl = 2; // Ultimate Failsafe
                  let cleanName = String(r[0]).trim().toLowerCase().replace(/\s+/g, ' ');
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

          let currentMap = {};
          let prevMap = {};
          
          validData.forEach(r => {
              let agent = String(r[1]).trim();
              let cleanAgent = agent.toLowerCase().replace(/\s+/g, ' ');
              let agentLevel = 2; // Default if ML missing
              
              if (useMasterList) {
                  let matchedLvl = null;
                  
                  // 1. Try Exact Match
                  if (validAgents.has(cleanAgent)) {
                      matchedLvl = validAgents.get(cleanAgent);
                  } else {
                      // 2. Try Fuzzy Match (e.g. "Dansou, Senami Patrick" vs "Dansou, Senami")
                      let gnParts = cleanAgent.replace(/,/g, ' ').replace(/\s+/g, ' ').trim().split(' ').filter(x => x.length > 1);
                      
                      for (let [mlName, lvl] of validAgents.entries()) {
                          let mlParts = mlName.replace(/,/g, ' ').replace(/\s+/g, ' ').trim().split(' ').filter(x => x.length > 1);
                          
                          // If Last Name and First Name match perfectly, it's the same person
                          if (gnParts.length > 1 && mlParts.length > 1 && gnParts[0] === mlParts[0] && gnParts[1] === mlParts[1]) {
                              matchedLvl = lvl;
                              break;
                          }
                          // Fallback substring check
                          if (cleanAgent.includes(mlName) || mlName.includes(cleanAgent)) {
                              matchedLvl = lvl;
                              break;
                          }
                      }
                  }
                  
                  // If we still can't find them, they are a Manager/TMA -> Skip entirely
                  if (matchedLvl === null) return; 
                  agentLevel = matchedLvl;
              }
              
              if (exclusions.has(agent)) return; 

              let inP = parseFloat(r[3]) || 0; if (inP > 1) inP = inP / 100;
              let outP = parseFloat(r[4]) || 0; if (outP > 1) outP = outP / 100;
              let cphV = parseFloat(r[2]) || 0;
              
              if (r[0] === targetMonth) {
                  currentMap[agent] = {
                      level: agentLevel, // Store the verified level!
                      cph: cphV, inPct: inP, outPct: outP, aht: r[5], 
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

              // INBOUND LOGIC (>= 75% Red)
              if (c.inPct >= 0.75) {
                  if (c.tasksXCalls >= 0.50 || c.totalCalls === 0) inboundList.push(payload);
                  else outOfScopeList.push(payload);
              }
              
              // OUTBOUND LOGIC (>= 75% Red, Level 2+)
              if (c.outPct >= 0.75 && c.level >= 2) {
                  outboundList.push(payload);
              }

              // CP/H LOGIC (< 7.3 Red)
              if (c.cph > 0 && c.cph <= 7.3) {
                  cphList.push(payload);
              }
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
      });

      let map = {};
      for (let i = headerIdx + 1; i < lines.length; i++) {
          let cols = this._parseCSVLine(lines[i]);
          let agent = cols[colMap.agent];
          if (!agent || agent.toLowerCase().includes('total')) continue;
          agent = agent.replace(/"/g, '').replace(/\b\w/g, c => c.toUpperCase()).trim();

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
              aht: this._parseSafeAHT(cols[colMap.aht])
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
      let sheetName = 'WF_GEM_DATA_V2';
      const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
      
      const writeToSheet = (targetSS) => {
          let sheet = targetSS.getSheetByName(sheetName);
          let expectedHeaders = ["Month", "Agent Name", "CPH", "Inbound %", "Outbound %", "AHT", "Tasks x Calls", "Total Tasks", "Calls Answered", "WDays"];
          
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
          
          let newRows = [];
          Object.keys(gemMap).forEach(agent => {
              let m = gemMap[agent];
              newRows.push([repMonth, agent, m.cph, m.inPct, m.outPct, m.aht, m.tasksXCalls, m.totalCalls, m.callsAnsw, m.wDays]);
          });
          
          let combined = retained.concat(newRows);
          sheet.clearContents();
          sheet.appendRow(headers);
          if (combined.length > 0) sheet.getRange(2, 1, combined.length, 10).setValues(combined);
      };

      writeToSheet(ssLocal);
      if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
          try { writeToSheet(SpreadsheetApp.openById(MasterConnector.DB_ID)); } catch(e) {}
      }
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
