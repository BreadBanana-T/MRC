/**
 * MODULE: WORKFORCE TRACKER
 * Features: Smart Upsert, DB Splitting, Master DB Routing, & Quarterly Bounds
 */

var WorkforceTracker = {

  // --- SMART DB ROUTER ---
  _getDB: function(sheetName) {
      const local = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (local && local.getLastRow() > 1) return local;

      if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
          try { 
              const mSheet = SpreadsheetApp.openById(MasterConnector.DB_ID).getSheetByName(sheetName);
              if (mSheet && mSheet.getLastRow() > 1) return mSheet;
          } catch(e) { console.error("Master DB Fallback Failed", e); }
      }
      return local;
  },

  importData: function(schedRaw, idpRaw) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ['WF_COACHING', 'WF_FURLOUGH', 'WF_IDP'].forEach(n => {
       if(!ss.getSheetByName(n)) ss.insertSheet(n);
    });
    if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
        try {
            const masterSS = SpreadsheetApp.openById(MasterConnector.DB_ID);
            ['WF_COACHING', 'WF_FURLOUGH', 'WF_IDP'].forEach(n => {
               if(!masterSS.getSheetByName(n)) masterSS.insertSheet(n);
            });
        } catch(e) {}
    }
    
    let msg = [];
    
    // --- SCHEDULE PARSER (ULTRA-LEAN WITH DB SPLITTING) ---
    if (schedRaw && schedRaw.trim().length > 0) {
      let cleanCoach = [];
      let cleanFurlough = [];
      const lines = schedRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      let currentAgent = "", currentID = "", currentDateStr = "";
      let currentY = 0, currentM = 0, currentD = 0;
      let lastTimeMins = -1, daysAdded = 0;
      let agentBuffer = []; 
      
      const segmentRegex = /([a-zA-ZÀ-ÿ0-9\/\(\)\s\-\.&]+?)\s+(\d{1,2}:\d{2}(?:\s?[AP]M)?)\s+(\d{1,2}:\d{2}(?:\s?[AP]M)?)\s*$/i;
      const dateRegex = /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/;
      const COACHING_CODES = ['ce séance', 'huddle', 'echo', 'mentor', 'hsc', 'health and safety', 'meet', 'roadshow', 'one on one', 'individuelle', 'pulsecheck', 'qual session', 'quality', 'sbys', 'survey', 'sondage en ligne', 'team'];
      const ACSU_CODES = ['acsu', 'solicited', 'libération', 'voluntary'];

      const flushAgentBuffer = () => {
          if (agentBuffer.length === 0) return;
          let isOffshore = false;
          if (currentID && String(currentID).startsWith("3")) isOffshore = true;
          for (let obj of agentBuffer) {
              if (obj.raw.toUpperCase().includes("TI ") || obj.raw.toUpperCase().includes("OFFSHORE")) {
                  isOffshore = true; break;
              }
          }
          let reg = isOffshore ? "Offshore" : "Onshore";
          
          agentBuffer.forEach(obj => { 
              let actLower = obj.act.toLowerCase();
              let isTeamLead = actLower.includes('team lead') || actLower.includes('équipe') || actLower.includes('equipe');
              let isCoach = !isTeamLead && COACHING_CODES.some(c => actLower.includes(c));
              let isFurlough = ACSU_CODES.some(c => actLower.includes(c));
              
              if (isCoach) cleanCoach.push([currentAgent, obj.dateStr, obj.act, obj.start, obj.end, reg]); 
              // OPTIMIZATION: Completely skip importing Furloughs for Offshore agents
              if (isFurlough && !isOffshore) cleanFurlough.push([currentAgent, obj.dateStr, obj.act, obj.start, obj.end, reg]);
          });
          agentBuffer = [];
      };

      lines.forEach(line => {
        let text = line.trim();
        if(text.startsWith('Agent Name') || text.startsWith('"Agent Name"')) return;

        if (text.includes('Agent:')) {
          flushAgentBuffer();
          let parts = text.split(':');
          if (parts.length > 1) {
              let agentData = parts[1].trim();
              let idMatch = agentData.match(/^(\d+)/);
              currentID = idMatch ? idMatch[1] : "";
              currentAgent = agentData.replace(/^\d+\s+/, '').trim();
          }
          return;
        } 
        
        if (text.includes('"') && text.includes(',')) {
          flushAgentBuffer(); 
          let csvParts = this._parseCSVLine(text);
          if (csvParts.length >= 6) { 
              let actLower = this._cleanActivity(csvParts[2]).toLowerCase();
              let isTeamLead = actLower.includes('team lead') || actLower.includes('équipe') || actLower.includes('equipe');
              let isCoach = !isTeamLead && COACHING_CODES.some(c => actLower.includes(c));
              let isFurlough = ACSU_CODES.some(c => actLower.includes(c));
              let isOff = csvParts[5] && csvParts[5].includes("Offshore");
              
              if (isCoach) cleanCoach.push([csvParts[0], this._parseDate(csvParts[1]), this._cleanActivity(csvParts[2]), csvParts[3], csvParts[4], csvParts[5]]);
              if (isFurlough && !isOff) cleanFurlough.push([csvParts[0], this._parseDate(csvParts[1]), this._cleanActivity(csvParts[2]), csvParts[3], csvParts[4], csvParts[5]]);
          }
          return;
        }

        let dMatch = text.match(dateRegex);
        if (dMatch) {
            currentDateStr = this._parseDate(dMatch[1]);
            [currentY, currentM, currentD] = currentDateStr.split('-').map(Number);
            lastTimeMins = -1;
            daysAdded = 0;
        }

        if (currentAgent && currentY) {
          let segMatch = text.match(segmentRegex);
          if (segMatch) {
            let act = this._cleanActivity(segMatch[1].trim());
            let tStartMins = this._timeToMins(segMatch[2].trim());
            if (lastTimeMins > -1 && tStartMins < lastTimeMins) daysAdded++;
            lastTimeMins = tStartMins;
            let actDateStr = Utilities.formatDate(new Date(currentY, currentM - 1, currentD + daysAdded), Session.getScriptTimeZone(), "yyyy-MM-dd");
            if (!act.toLowerCase().match(/^activity|^scheduled/)) agentBuffer.push({ raw: text, dateStr: actDateStr, act: act, start: segMatch[2].trim(), end: segMatch[3].trim() });
          }
        }
      });
      flushAgentBuffer(); 
      
      if (cleanCoach.length > 0) {
        this._upsertData('WF_COACHING', cleanCoach, [0, 1], ['Agent Name', 'Date', 'Activity', 'Start Time', 'End Time', 'Region']);
        msg.push(`✔ Coaching DB: Processed ${cleanCoach.length} records.`);
      }
      if (cleanFurlough.length > 0) {
        this._upsertData('WF_FURLOUGH', cleanFurlough, [0, 1], ['Agent Name', 'Date', 'Activity', 'Start Time', 'End Time', 'Region']);
        msg.push(`✔ Furlough DB: Processed ${cleanFurlough.length} records.`);
      }
    }

    // --- IDP PARSER ---
    if (idpRaw && idpRaw.trim().length > 0) {
      let cleanIDP = [];
      const lines = idpRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      let headerIdx = lines.findIndex(l => { const low = l.toLowerCase(); return (low.includes('req') || low.includes('besoin')) && (low.includes('open') || low.includes('ouvert') || low.includes('dispo')); });
      if (headerIdx > -1) {
        let headers = this._parseCSVLine(lines[headerIdx]);
        let colMap = {};
        headers.forEach((h, i) => {
          let match = h.match(/(\w+\s\d{1,2},?\s\d{4})/);
          if (match) {
            let dStr = this._parseDate(match[1]);
            if (h.toLowerCase().includes('req') || h.toLowerCase().includes('besoin')) colMap[i] = { date: dStr, type: 'req' };
            else if ((h.toLowerCase().includes('open') || h.toLowerCase().includes('ouvert')) && !h.toLowerCase().includes('+/-')) colMap[i] = { date: dStr, type: 'open' };
          }
        });
        let dataByDay = {};
        for (let i = headerIdx + 1; i < lines.length; i++) {
          let cols = this._parseCSVLine(lines[i]);
          if (cols[0] && cols[0].includes(':')) {
            let tNorm = this._formatTimeStr(cols[0]);
            Object.keys(colMap).forEach(idx => {
               if (cols[idx] !== undefined) {
                 if (!dataByDay[colMap[idx].date]) dataByDay[colMap[idx].date] = {};
                 if (!dataByDay[colMap[idx].date][tNorm]) dataByDay[colMap[idx].date][tNorm] = { req:0, open:0 };
                 let val = parseFloat(String(cols[idx]).replace(/,/g, '')) || 0;
                 dataByDay[colMap[idx].date][tNorm][colMap[idx].type] = val;
               }
            });
          }
        }
        Object.keys(dataByDay).forEach(day => Object.keys(dataByDay[day]).forEach(time => cleanIDP.push([day, time, dataByDay[day][time].req, dataByDay[day][time].open])));
        if (cleanIDP.length > 0) {
          this._upsertData('WF_IDP', cleanIDP, [0, 1], ['Day', 'Interval', 'Required', 'Open']);
          msg.push(`✔ IDP: Imported ${cleanIDP.length} intervals.`);
        }
      }
    }
    return msg.length ? msg.join('\n') : "No valid data found to import.";
  },

  _getCycleForEpoch: function(epoch) {
      let targetUTC = Date.UTC(new Date(epoch).getFullYear(), new Date(epoch).getMonth(), new Date(epoch).getDate(), 12, 0, 0);
      if (new Date(epoch).getHours() >= 23) targetUTC += 86400000; 
      let diffWeeks = Math.floor(Math.floor((targetUTC - Date.UTC(2026, 0, 29, 12, 0, 0)) / 86400000) / 7);
      return (Math.abs(diffWeeks) % 2 === 1) ? "WEEK B" : "WEEK A";
  },

  _calculateEpochBoundaries: function(mode, refDateStr) {
      let [tY, tM, tD] = refDateStr.split('-').map(Number);
      let tObj = new Date(tY, tM-1, tD, 12, 0, 0, 0);
      let tStart = 0, tEnd = 0, label = "", cycle = "";
      
      if (mode === 'day') {
          tStart = new Date(tY, tM-1, tD - 1, 23, 0, 0, 0).getTime();
          tEnd = new Date(tY, tM-1, tD, 23, 0, 0, 0).getTime();
          label = Utilities.formatDate(tObj, Session.getScriptTimeZone(), "MMM dd");
          cycle = this._getCycleForEpoch(tStart);
      } 
      else if (mode === 'week') {
          let diff = (tObj.getDay() >= 4) ? (4 - tObj.getDay()) : -(tObj.getDay() + 3);
          let wDate = new Date(tObj); wDate.setDate(tObj.getDate() + diff - 1); wDate.setHours(23, 0, 0, 0);
          tStart = wDate.getTime(); tEnd = tStart + (7 * 86400000);
          cycle = this._getCycleForEpoch(tStart);
          label = `${Utilities.formatDate(new Date(tStart), Session.getScriptTimeZone(), "MMM dd")} to ${Utilities.formatDate(new Date(tEnd), Session.getScriptTimeZone(), "MMM dd")}`;
      } 
      else if (mode === 'month' || mode === 'quarter') {
          let sMonth = (mode === 'month') ? tM - 1 : Math.floor((tM - 1) / 3) * 3;
          let eMonth = (mode === 'month') ? tM : sMonth + 3;

          let sWed = new Date(tY, sMonth, 1, 12, 0, 0);
          let sOff = (sWed.getDay() >= 3) ? (3 - sWed.getDay()) : -(sWed.getDay() + 4);
          if (sOff > 0) sOff -= 7; sWed.setDate(sWed.getDate() + sOff); sWed.setHours(23, 0, 0, 0);
          tStart = sWed.getTime();
          
          let eWed = new Date(tY, eMonth, 1, 12, 0, 0);
          let eOff = (eWed.getDay() >= 3) ? (3 - eWed.getDay()) : -(eWed.getDay() + 4);
          if (eOff > 0) eOff -= 7; eWed.setDate(eWed.getDate() + eOff); eWed.setHours(23, 0, 0, 0);
          tEnd = eWed.getTime();
          
          label = mode === 'month' ? `WFM Month: ${Utilities.formatDate(new Date(tStart), Session.getScriptTimeZone(), "MMM dd")} - ${Utilities.formatDate(new Date(tEnd), Session.getScriptTimeZone(), "MMM dd")}` : `Q${Math.floor((tM - 1) / 3) + 1}: ${Utilities.formatDate(new Date(tStart), Session.getScriptTimeZone(), "MMM dd")} - ${Utilities.formatDate(new Date(tEnd), Session.getScriptTimeZone(), "MMM dd")}`;
          cycle = mode === 'month' ? "MONTH" : "QUARTER"; 
      }
      return { start: tStart, end: tEnd, label: label, cycle: cycle, startStr: Utilities.formatDate(new Date(tStart), Session.getScriptTimeZone(), "yyyy-MM-dd") };
  },

  getAnalytics: function(mode, refDate, trackerType, regionFilter = 'All', cycleFilter = 'ALL') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbIDP = this._getDB('WF_IDP');
    
    let dbSched;
    if (trackerType === 'coaching') dbSched = this._getDB('WF_COACHING');
    else if (trackerType === 'furlough') dbSched = this._getDB('WF_FURLOUGH');
    
    if (!dbIDP || !dbSched) return JSON.stringify({ error: "Databases missing. Please run WFM Import first." });
    
    const bounds = this._calculateEpochBoundaries(mode, refDate);
    let searchStart = new Date(bounds.start); searchStart.setDate(searchStart.getDate() - 1);
    const searchStartStr = Utilities.formatDate(searchStart, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const endStr = Utilities.formatDate(new Date(bounds.end), Session.getScriptTimeZone(), "yyyy-MM-dd");

    const idpData = dbIDP.getLastRow() > 1 ? dbIDP.getDataRange().getValues().slice(1) : [];
    const schedData = dbSched.getLastRow() > 1 ? dbSched.getDataRange().getValues().slice(1) : [];

    let buckets = [];
    let combinedEvents = [];
    let groupedLogs = {}; 

    if (mode === 'day' && trackerType === 'furlough') {
      buckets = Array.from({length: 96}, (_, i) => ({ index: i, label: this._indexToTime(i), supply: 0, demand: 0, net: 0 }));
      idpData.forEach(row => {
        let rowDateStr = this._formatDate(row[0]);
        if (!rowDateStr) return;
        let [rY, rM, rD] = rowDateStr.split('-').map(Number);
        let mins = this._timeToMins(row[1]);
        let blockTime = new Date(rY, rM-1, rD, Math.floor(mins/60), mins%60, 0, 0).getTime();
        
        if (blockTime >= bounds.start && blockTime < bounds.end) { 
          let idx = this._timeToBucket(row[1]);
          if (idx > -1) { buckets[idx].demand += Number(row[2] || 0); buckets[idx].supply += Number(row[3] || 0); }
        }
      });
    }

    schedData.forEach(row => {
        let rowDateStr = this._formatDate(row[1]);
        if (!rowDateStr) return;
        let rowRegion = row[5] ? String(row[5]).trim() : 'Onshore';
        if (regionFilter !== 'All' && rowRegion !== regionFilter) return;

        if (rowDateStr >= searchStartStr && rowDateStr <= endStr) {
            let agent = String(row[0]).trim();
            let [rY, rM, rD] = rowDateStr.split('-').map(Number);
            let startMins = this._timeToMins(row[3]);
            let endMinsRaw = this._timeToMins(row[4]);
            let endMins = endMinsRaw < startMins ? endMinsRaw + 1440 : endMinsRaw; 

            this._getShiftSplits(startMins, endMins).forEach(split => {
                let splitStartEpoch = new Date(rY, rM-1, rD, Math.floor(split.startMins/60), split.startMins%60, 0, 0).getTime();

                if (splitStartEpoch >= bounds.start && splitStartEpoch < bounds.end) {
                    if ((mode === 'month' || mode === 'quarter') && cycleFilter !== 'ALL') {
                        if (this._getCycleForEpoch(splitStartEpoch) !== cycleFilter) return;
                    }
                    
                    let effDateStr = Utilities.formatDate(new Date(splitStartEpoch), Session.getScriptTimeZone(), "yyyy-MM-dd");

                    if (trackerType === 'furlough' && mode === 'day') {
                        for (let min = split.startMins; min < split.endMins; min += 15) {
                            let blockTime = new Date(rY, rM-1, rD, Math.floor(min/60), min%60, 0, 0).getTime();
                            if (blockTime >= bounds.start && blockTime < bounds.end) {
                                let idx = Math.floor((min % 1440) / 15);
                                if (idx >= 0 && idx < 96) buckets[idx].supply = Math.max(0, buckets[idx].supply - 1);
                            }
                        }
                    }

                    // FIX: Make groupKey truly unique by adding the raw start time (row[3])
                    // This prevents multiple distinct furlough blocks from combining into one visually misleading entry.
                    let rawStartStr = String(row[3]).trim();
                    let groupKey = `${agent}_${effDateStr}_${split.shift}_${rawStartStr}`;
                    if (trackerType === 'coaching') groupKey += `_${row[2]}`;

                    if (!groupedLogs[groupKey]) {
                        groupedLogs[groupKey] = { 
                            date: effDateStr, agent: agent, 
                            activityName: trackerType === 'coaching' ? row[2] : "Time Off", 
                            shift: split.shift, hours: split.hours, 
                            timeStart: this._minsToTime(split.startMins), 
                            timeEnd: this._minsToTime(split.endMins) 
                        };
                    } else {
                        groupedLogs[groupKey].hours += split.hours; 
                        groupedLogs[groupKey].timeEnd = this._minsToTime(split.endMins);
                    }
                }
            });
        }
    });

    if (mode === 'day' && trackerType === 'furlough') buckets.forEach(b => { b.net = parseFloat((b.supply - b.demand).toFixed(2)); });

    combinedEvents = Object.values(groupedLogs).map(g => ({ date: g.date, agent: g.agent, activityName: g.activityName, shift: g.shift, hours: parseFloat(g.hours.toFixed(2)), time: `${g.timeStart} - ${g.timeEnd}` }));
    let totals = { all: 0, morning: 0, evening: 0, night: 0, count: combinedEvents.length };
    combinedEvents.forEach(f => { totals.all += f.hours; if (f.shift === 'Morning') totals.morning += f.hours; else if (f.shift === 'Evening') totals.evening += f.hours; else totals.night += f.hours; });

    return JSON.stringify({ mode: mode, trackerType: trackerType, label: bounds.label, cycle: bounds.cycle, grid: buckets, events: combinedEvents, totals: totals });
  },

  getUnifiedWeeklyReport: function(refDateStr) {
        const bounds = this._calculateEpochBoundaries("week", refDateStr);
        let report = { cycle: bounds.cycle, period: bounds.label, agents: {} };
        const getAg = (name) => {
            let n = String(name).replace(/\b\w/g, c => c.toUpperCase()).trim(); 
            if(!report.agents[n]) report.agents[n] = { name: n, acsu: 0, coach: 0, safe: 0, icl: 0, ulc: 0, total: 0 };
            return report.agents[n];
        };

        let searchStart = new Date(bounds.start); searchStart.setDate(searchStart.getDate() - 1);
        const sStr = Utilities.formatDate(searchStart, Session.getScriptTimeZone(), "yyyy-MM-dd");
        const eStr = Utilities.formatDate(new Date(bounds.end), Session.getScriptTimeZone(), "yyyy-MM-dd");

        const parseDB = (sheetName, metricName) => {
            const db = this._getDB(sheetName);
            if (!db || db.getLastRow() < 2) return;
            db.getDataRange().getValues().slice(1).forEach(row => {
                let dStr = this._formatDate(row[1]);
                if (dStr >= sStr && dStr <= eStr) {
                    let [rY, rM, rD] = dStr.split('-').map(Number);
                    let sMins = this._timeToMins(row[3]); let eMinsR = this._timeToMins(row[4]);
                    let eMins = eMinsR < sMins ? eMinsR + 1440 : eMinsR; 
                    this._getShiftSplits(sMins, eMins).forEach(s => {
                        let epoch = new Date(rY, rM-1, rD, Math.floor(s.startMins/60), s.startMins%60, 0, 0).getTime();
                        if (epoch >= bounds.start && epoch < bounds.end) { getAg(row[0])[metricName] += s.hours; getAg(row[0]).total += s.hours; }
                    });
                }
            });
        };

        parseDB('WF_COACHING', 'coach');
        parseDB('WF_FURLOUGH', 'acsu');

        const dbSess = this._getDB('DB_Sessions');
        if (dbSess && dbSess.getLastRow() > 1) {
            dbSess.getDataRange().getValues().slice(1).forEach(row => {
                let sessionEpoch = new Date(row[3]).getTime();
                if (sessionEpoch >= bounds.start && sessionEpoch < bounds.end) {
                    let role = String(row[2]).toUpperCase(); let h = Number(row[6]) || 0; 
                    if (role.includes('SAFE')) { getAg(row[1]).safe += h; getAg(row[1]).total += h; }
                    else if (role.includes('ICL')) { getAg(row[1]).icl += h; getAg(row[1]).total += h; }
                    else if (role.includes('ULC') || role.includes('FIRE')) { getAg(row[1]).ulc += h; getAg(row[1]).total += h; }
                }
            });
        }

        let finalArr = Object.values(report.agents).filter(a => a.total > 0);
        finalArr.forEach(a => { ['acsu','coach','safe','icl','ulc','total'].forEach(k => a[k] = parseFloat(a[k].toFixed(2))); });
        report.data = finalArr.sort((a,b) => b.total - a.total);
        return JSON.stringify(report);
  },

  archiveUnifiedWeek: function(refDateStr) {
      const reportStr = this.getUnifiedWeeklyReport(refDateStr);
      const report = JSON.parse(reportStr);
      
      const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
      
      const writeToArchiveSheet = (spreadsheet) => {
          let archiveSheet = spreadsheet.getSheetByName('Weekly_Archives');
          if (!archiveSheet) { 
              archiveSheet = spreadsheet.insertSheet('Weekly_Archives'); 
              archiveSheet.appendRow(["Timestamp", "Cycle", "Period", "Agent Name", "ACSU", "Coaching", "SAFE", "ICL", "ULC FIRE", "Total Off-Phone"]); 
              archiveSheet.getRange(1,1,1,10).setFontWeight("bold"); 
          }

          const data = archiveSheet.getDataRange().getValues();
          for (let i = data.length - 1; i >= 1; i--) {
              if (data[i][2] === report.period) archiveSheet.deleteRow(i + 1);
          }

          const rows = report.data.map(a => [ new Date(), report.cycle, report.period, a.name, a.acsu, a.coach, a.safe, a.icl, a.ulc, a.total ]);
          if (rows.length > 0) archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
      };

      writeToArchiveSheet(ssLocal);
      
      if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
          try {
             writeToArchiveSheet(SpreadsheetApp.openById(MasterConnector.DB_ID));
          } catch(e) { console.error("Master DB Archive Sync Failed", e); }
      }

      return `Successfully archived ${report.data.length} agents for ${report.cycle} (${report.period}).`;
  },

  getArchiveList: function() {
      let sheet = this._getDB('Weekly_Archives');
      if (!sheet || sheet.getLastRow() < 2) return "[]";
      
      const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getValues(); 
      const unique = [];
      const seen = new Set();
      
      for (let i = data.length - 1; i >= 0; i--) {
          const cycle = data[i][0];
          const period = data[i][1];
          if (!seen.has(period)) { seen.add(period); unique.push({ cycle: cycle, period: period }); }
      }
      return JSON.stringify(unique);
  },

  getArchivedReport: function(targetPeriod) {
      let sheet = this._getDB('Weekly_Archives');
      if (!sheet || sheet.getLastRow() < 2) return "{}";
      
      const data = sheet.getDataRange().getValues();
      let report = { cycle: "ARCHIVED", period: targetPeriod, data: [] };
      
      for (let i = 1; i < data.length; i++) {
          if (data[i][2] === targetPeriod) {
              report.cycle = data[i][1];
              report.data.push({ name: data[i][3], acsu: data[i][4], coach: data[i][5], safe: data[i][6], icl: data[i][7], ulc: data[i][8], total: data[i][9] });
          }
      }
      
      report.data.sort((a,b) => b.total - a.total);
      return JSON.stringify(report);
  },

  _getShiftSplits: function(startMins, endMins) {
      let splits = []; let current = startMins;
      while (current < endMins) {
          let timeOfDay = current % 1440; let shiftType = "Night"; let nextBound = current - timeOfDay + 420;
          if (timeOfDay >= 420 && timeOfDay < 900) { shiftType = "Morning"; nextBound = current - timeOfDay + 900; } 
          else if (timeOfDay >= 900 && timeOfDay < 1380) { shiftType = "Evening"; nextBound = current - timeOfDay + 1380; } 
          else if (timeOfDay >= 1380) nextBound = current - timeOfDay + 1860;

          let chunkEnd = Math.min(endMins, nextBound);
          splits.push({ shift: shiftType, hours: (chunkEnd - current) / 60, startMins: current, endMins: chunkEnd });
          current = chunkEnd;
      }
      return splits;
  },
  
  _upsertData: function(sheetName, newRows, keyIndices, headersArray) {
    const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
    this._executeUpsert(ssLocal, sheetName, newRows, keyIndices, headersArray);
    if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
        try { const ssMaster = SpreadsheetApp.openById(MasterConnector.DB_ID);
        this._executeUpsert(ssMaster, sheetName, newRows, keyIndices, headersArray); } catch(e) {}
    }
  },
  
  _executeUpsert: function(targetSpreadsheet, sheetName, newRows, keyIndices, headersArray) {
    let sheet = targetSpreadsheet.getSheetByName(sheetName);
    if (!sheet) sheet = targetSpreadsheet.insertSheet(sheetName);
    
    // FIX: Safely handles both Strings AND native Google Sheet Date/Time Objects so matching is identical
    const makeKey = (row) => keyIndices.map(i => {
        let val = row[i];
        if (val instanceof Date) {
            // Checks if it is specifically a Time object (often stored with an 1899 year base)
            if (val.getFullYear() < 1950) {
                return Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
            }
            return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        return String(val).trim().toLowerCase();
    }).join('_');

    const incomingKeys = new Set(newRows.map(makeKey));
    
    let existingData = [];
    if (sheet.getLastRow() > 0) existingData = sheet.getDataRange().getValues();
    if (existingData.length === 1 && existingData[0].join('') === "") existingData = [];
    
    const headers = existingData.length > 0 ? existingData.shift() : headersArray;
    
    const retainedRows = existingData.filter(row => {
       if(!row[keyIndices[0]]) return true; 
       return !incomingKeys.has(makeKey(row));
    });
    
    const combined = retainedRows.concat(newRows);
    sheet.clearContents(); sheet.appendRow(headers);
    if (combined.length > 0) sheet.getRange(2, 1, combined.length, combined[0].length).setValues(combined);
  },
  
  _minsToTime: function(mins) { let m = mins % 1440; let h = Math.floor(m / 60); let mm = m % 60; return `${h < 10 ? '0'+h : h}:${mm < 10 ? '0'+mm : mm}`; },
  _timeToMins: function(tStr) {
       if (!tStr) return 0;
       if (tStr instanceof Date) return (tStr.getHours() * 60) + tStr.getMinutes();
       let match = String(tStr).match(/(\d{1,2}):(\d{2})\s*([AP]M)?/i); if (!match) return 0;
       let h = parseInt(match[1]), m = parseInt(match[2]), amp = match[3] ? match[3].toUpperCase() : null;
       if (amp === 'PM' && h < 12) h += 12; if (amp === 'AM' && h === 12) h = 0; return (h * 60) + m;
  },
  _parseCSVLine: function(text) {
    if (text.includes('\t')) return text.split('\t').map(s => s.trim());
    let ret = [], inQuote = false, token = "";
    for(let i=0; i<text.length; i++) {
      let char = text[i];
      if(char === '"') { inQuote = !inQuote; continue; }
      if(char === ',' && !inQuote) { ret.push(token.trim()); token = ""; } else token += char;
    }
    ret.push(token.trim()); return ret;
  },
  _parseDate: function(s) { let d=new Date(s); return isNaN(d)?s:Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); },
  _formatDate: function(d) { 
      if (!d) return "";
      return (d instanceof Date) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(d).substring(0,10); 
  },
  _formatTimeStr: function(t) { let d=new Date(`2000/01/01 ${t}`); return isNaN(d)?t:Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm'); },
  _cleanActivity: function(s) { return s.replace(/\d{2}\s?[AP]M/gi, '').trim(); },
  _timeToBucket: function(val) {
    if (!val) return -1;
    if (val instanceof Date) return (val.getHours()*4) + Math.floor(val.getMinutes()/15);
    let parts = String(val).match(/(\d+):(\d+)\s?([AP]M)?/i);
    if (parts) {
      let h = parseInt(parts[1]), m = parseInt(parts[2]), amp = parts[3] ? parts[3].toUpperCase() : null;
      if (amp === 'PM' && h < 12) h += 12; if (amp === 'AM' && h === 12) h = 0; return (h * 4) + Math.floor(m / 15);
    } return -1;
  },
  _indexToTime: function(i) { let h=Math.floor(i/4), m=(i%4)*15; return `${h<10?'0'+h:h}:${m===0?'00':m}`; }
};
