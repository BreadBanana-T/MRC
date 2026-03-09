/**
 * MODULE: WORKFORCE TRACKER (V7.4 - LOCAL HOST ONLY)
 */

var WorkforceTracker = {

  _getDB: function(sheetName) {
      return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  },

  importData: function(schedRaw, idpRaw) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ['WF_COACHING', 'WF_FURLOUGH', 'WF_ROLES', 'WF_IDP'].forEach(n => {
       if(!ss.getSheetByName(n)) ss.insertSheet(n);
    });
    
    let msg = [];
    let schedDates = [];
    let idpDates = [];
    let muSet = "";

    if (schedRaw && schedRaw.trim().length > 0) {
      let cleanCoach = [];
      let cleanFurlough = [];
      let cleanRoles = [];
      
      const lines = schedRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      let currentAgent = "", currentID = "", currentDateStr = "";
      let currentY = 0, currentM = 0, currentD = 0;
      let lastTimeMins = -1, daysAdded = 0;
      let agentBuffer = []; 
      
      const segmentRegex = /([a-zA-ZÀ-ÿ0-9\/\(\)\s\-\.&]+?)\s+(\d{1,2}:\d{2}(?:\s?[AP]M)?)\s+(\d{1,2}:\d{2}(?:\s?[AP]M)?)\s*$/i;
      const dateRegex = /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/;
      
      const COACHING_CODES = ['ce séance', 'huddle', 'echo', 'mentor', 'hsc', 'health and safety', 'meet', 'roadshow', 'one on one', 'individuelle', 'pulsecheck', 'qual session', 'quality', 'sbys', 'survey', 'sondage en ligne', 'team'];
      const ACSU_CODES = ['acsu', 'solicited', 'libération', 'voluntary'];
      const ROLE_CODES = ['safe onqueue', 'safe en ligne', 'icl', 'incident', 'ulc', 'fire', 'feu', 'wofqt', 'tower'];

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
              let isRole = ROLE_CODES.some(c => actLower.includes(c));
              
              let roleType = "";
              if (isRole) {
                  if (actLower.includes('safe')) roleType = "SAFE";
                  else if (actLower.includes('icl') || actLower.includes('incident')) roleType = "ICL";
                  else if (actLower.includes('ulc') || actLower.includes('fire') || actLower.includes('feu')) roleType = "ULC FIRE";
                  else if (actLower.includes('wofqt') || actLower.includes('tower')) roleType = "TOWER";
              }
              
              if (isCoach) cleanCoach.push([currentAgent, obj.dateStr, obj.act, obj.start, obj.end, reg]); 
              if (isFurlough && !isOffshore) cleanFurlough.push([currentAgent, obj.dateStr, obj.act, obj.start, obj.end, reg]);
              if (isRole) cleanRoles.push([currentAgent, obj.dateStr, roleType, obj.start, obj.end, reg]);
              schedDates.push(obj.dateStr);
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
              let isRole = ROLE_CODES.some(c => actLower.includes(c));
              
              let roleType = "";
              if (isRole) {
                  if (actLower.includes('safe')) roleType = "SAFE";
                  else if (actLower.includes('icl') || actLower.includes('incident')) roleType = "ICL";
                  else if (actLower.includes('ulc') || actLower.includes('fire') || actLower.includes('feu')) roleType = "ULC FIRE";
                  else if (actLower.includes('wofqt') || actLower.includes('tower')) roleType = "TOWER";
              }

              let isOff = csvParts[5] && csvParts[5].includes("Offshore");
              let pDate = this._parseDate(csvParts[1]);
              
              if (isCoach) cleanCoach.push([csvParts[0], pDate, this._cleanActivity(csvParts[2]), csvParts[3], csvParts[4], csvParts[5]]);
              if (isFurlough && !isOff) cleanFurlough.push([csvParts[0], pDate, this._cleanActivity(csvParts[2]), csvParts[3], csvParts[4], csvParts[5]]);
              if (isRole) cleanRoles.push([csvParts[0], pDate, roleType, csvParts[3], csvParts[4], csvParts[5]]);
              schedDates.push(pDate);
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
        this._executeDestructiveUpsert('WF_COACHING', cleanCoach, ['Agent Name', 'Date', 'Activity', 'Start Time', 'End Time', 'Region']);
        msg.push(`Coaching`);
      }
      if (cleanFurlough.length > 0) {
        this._executeDestructiveUpsert('WF_FURLOUGH', cleanFurlough, ['Agent Name', 'Date', 'Activity', 'Start Time', 'End Time', 'Region']);
        msg.push(`Furloughs`);
      }
      if (cleanRoles.length > 0) {
        this._executeDestructiveUpsert('WF_ROLES', cleanRoles, ['Agent Name', 'Date', 'Role', 'Start Time', 'End Time', 'Region']);
        msg.push(`Roles`);
      }
    }

    if (idpRaw && idpRaw.trim().length > 0) {
      let cleanIDP = [];
      const lines = idpRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      
      for (let i = 0; i < Math.min(lines.length, 15); i++) {
          let muMatch = lines[i].match(/MU Set:\s*(\d+)/i);
          if (muMatch) { muSet = muMatch[1]; break; }
      }

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
        Object.keys(dataByDay).forEach(day => {
            idpDates.push(day);
            Object.keys(dataByDay[day]).forEach(time => cleanIDP.push([day, time, dataByDay[day][time].req, dataByDay[day][time].open]))
        });
        if (cleanIDP.length > 0) {
          this._executeDestructiveUpsert('WF_IDP', cleanIDP, ['Day', 'Interval', 'Required', 'Open']);
          msg.push(`IDP`);
        }
      }
    }

    try {
        const props = PropertiesService.getDocumentProperties();
        const curTime = Utilities.formatDate(new Date(), "America/Toronto", "MMM dd, HH:mm");
        if (schedDates.length > 0) {
            schedDates.sort();
            props.setProperty('SYNC_SCHED', `WFM: ${schedDates[0]} to ${schedDates[schedDates.length-1]} (Sync: ${curTime})`);
        }
        if (idpDates.length > 0) {
            idpDates.sort();
            props.setProperty('SYNC_IDP', `IDP: ${idpDates[0]} to ${idpDates[idpDates.length-1]} (Sync: ${curTime})`);
            if (muSet) props.setProperty('MU_SET', muSet);
        }
    } catch(e) {}
    
    if (msg.length === 0) return "Basic Schedule Synced.";
    return `Synced: ${msg.join(' | ')}`;
  },

  _executeDestructiveUpsert: function(sheetName, newRows, headersArray) {
    const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
    this._runDestructiveLogic(ssLocal, sheetName, newRows, headersArray);
  },

  _runDestructiveLogic: function(targetSpreadsheet, sheetName, newRows, headersArray) {
      let sheet = targetSpreadsheet.getSheetByName(sheetName);
      if (!sheet) sheet = targetSpreadsheet.insertSheet(sheetName);

      let existingData = [];
      if (sheet.getLastRow() > 0) existingData = sheet.getDataRange().getDisplayValues();
      if (existingData.length === 1 && existingData[0].join('') === "") existingData = [];
      const headers = existingData.length > 0 ? existingData.shift() : headersArray;

      let isIDP = sheetName === 'WF_IDP';
      let wipeKeys = new Set();
      newRows.forEach(r => {
          let dateStr = this._formatDate(r[isIDP ? 0 : 1]);
          if (isIDP) wipeKeys.add(dateStr);
          else wipeKeys.add(String(r[0]).trim().toLowerCase() + "_" + dateStr);
      });
      
      let cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - 45);
      let cutoffStr = Utilities.formatDate(cutoffDate, "America/Toronto", "yyyy-MM-dd");
      
      const retainedRows = existingData.filter(row => {
          if (!row[0]) return false;
          let rDate = this._formatDate(row[isIDP ? 0 : 1]);
          if (!rDate || rDate < cutoffStr) return false; 
          let k = isIDP ? rDate : String(row[0]).trim().toLowerCase() + "_" + rDate;
          if (wipeKeys.has(k)) return false; 
          return true;
      });

      const combined = retainedRows.concat(newRows);
      sheet.clearContents(); 
      sheet.appendRow(headers);
      if (combined.length > 0) sheet.getRange(2, 1, combined.length, combined[0].length).setValues(combined);
  },

  _getCycleForEpoch: function(epoch) {
      let d = new Date(epoch);
      let targetUTC = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0);
      if (d.getHours() >= 23) targetUTC += 86400000;
      let diffWeeks = Math.floor(Math.floor((targetUTC - Date.UTC(2026, 0, 29, 12, 0, 0)) / 86400000) / 7);
      return (Math.abs(diffWeeks) % 2 === 1) ? "WEEK B" : "WEEK A";
  },

  _calculateEpochBoundaries: function(mode, refDateStr) {
      let rY = parseInt(refDateStr.substring(0,4));
      let rM = parseInt(refDateStr.substring(5,7));
      let rD = parseInt(refDateStr.substring(8,10));
      let tObj = new Date(rY, rM-1, rD, 12, 0, 0, 0);
      
      let tStart = 0, tEnd = 0, label = "", cycle = "";
      
      if (mode === 'day') {
          tStart = new Date(rY, rM-1, rD, 0, 0, 0, 0).getTime();
          tEnd = new Date(rY, rM-1, rD, 23, 59, 59, 999).getTime();
          label = Utilities.formatDate(tObj, "America/Toronto", "yyyy-MM-dd");
          cycle = this._getCycleForEpoch(tStart);
      } 
      else if (mode === 'week') {
          let wStart = new Date(rY, rM-1, rD, 0, 0, 0, 0);
          let dayOfWeek = wStart.getDay(); 
          let offset = (dayOfWeek >= 3) ? (dayOfWeek - 3) : (dayOfWeek + 4);
          wStart.setDate(wStart.getDate() - offset);
          wStart.setHours(23, 0, 0, 0);
          
          let wEnd = new Date(wStart);
          wEnd.setDate(wStart.getDate() + 7);
          wEnd.setHours(22, 59, 59, 999);
          
          tStart = wStart.getTime();
          tEnd = wEnd.getTime();
          
          cycle = this._getCycleForEpoch(tStart);
          label = `${Utilities.formatDate(wStart, "America/Toronto", "MMM dd, HH:mm")} to ${Utilities.formatDate(wEnd, "America/Toronto", "MMM dd, HH:mm")}`;
      } 
      else if (mode === 'month' || mode === 'quarter') {
          let sMonth = (mode === 'month') ? rM - 1 : Math.floor((rM - 1) / 3) * 3;
          let eMonth = (mode === 'month') ? rM : sMonth + 3;

          let sDate = new Date(rY, sMonth, 1, 0, 0, 0, 0);
          tStart = sDate.getTime();
          
          let eDate = new Date(rY, eMonth, 0, 23, 59, 59, 999);
          tEnd = eDate.getTime();
          
          label = mode === 'month' 
              ? `Month: ${Utilities.formatDate(sDate, "America/Toronto", "MMM dd")} to ${Utilities.formatDate(eDate, "America/Toronto", "MMM dd")}` 
              : `Q${Math.floor((rM - 1) / 3) + 1}: ${Utilities.formatDate(sDate, "America/Toronto", "MMM dd")} to ${Utilities.formatDate(eDate, "America/Toronto", "MMM dd")}`;
          
          cycle = mode === 'month' ? "MONTH" : "QUARTER"; 
      }
      return { start: tStart, end: tEnd, label: label, cycle: cycle, startStr: Utilities.formatDate(new Date(tStart), "America/Toronto", "yyyy-MM-dd") };
  },

  getAnalytics: function(mode, refDate, trackerType, regionFilter = 'All', cycleFilter = 'ALL') {
    const dbIDP = this._getDB('WF_IDP');
    let dbSched;
    if (trackerType === 'coaching') dbSched = this._getDB('WF_COACHING');
    else if (trackerType === 'furlough') dbSched = this._getDB('WF_FURLOUGH');
    else if (trackerType === 'roles') dbSched = this._getDB('WF_ROLES');
    
    if (!dbIDP || !dbSched) return JSON.stringify({ error: "Databases missing. Please run WFM Import first." });
    
    const bounds = this._calculateEpochBoundaries(mode, refDate);
    let searchStart = new Date(bounds.start); searchStart.setDate(searchStart.getDate() - 1);
    const searchStartStr = Utilities.formatDate(searchStart, "America/Toronto", "yyyy-MM-dd");
    const endStr = Utilities.formatDate(new Date(bounds.end), "America/Toronto", "yyyy-MM-dd");

    const idpData = dbIDP.getLastRow() > 1 ? dbIDP.getDataRange().getDisplayValues().slice(1) : [];
    const schedData = dbSched.getLastRow() > 1 ? dbSched.getDataRange().getDisplayValues().slice(1) : [];

    let buckets = [];
    let combinedEvents = [];
    let groupedLogs = {}; 

    if (mode === 'day' && trackerType === 'furlough') {
      buckets = Array.from({length: 96}, (_, i) => ({ index: i, label: this._indexToTime(i), supply: 0, demand: 0, net: 0 }));
      
      idpData.forEach(row => {
        let rowDateStr = this._formatDate(row[0]);
        if (!rowDateStr) return;
        
        let rY = parseInt(rowDateStr.substring(0,4));
        let rM = parseInt(rowDateStr.substring(5,7));
        let rD = parseInt(rowDateStr.substring(8,10));
        if (isNaN(rY) || isNaN(rM) || isNaN(rD)) return; 
        
        let mins = this._timeToMins(row[1]);
        let blockTime = new Date(rY, rM-1, rD, Math.floor(mins/60), mins%60, 0, 0).getTime();
        
        if (blockTime >= bounds.start && blockTime <= bounds.end) { 
          let idx = this._timeToBucket(row[1]);
          if (idx > -1) { 
              let dem = parseFloat(String(row[2]).replace(',', '.')) || 0;
              let sup = parseFloat(String(row[3]).replace(',', '.')) || 0;
              buckets[idx].demand += dem; 
              buckets[idx].supply += sup; 
          }
        }
      });
    }

    let processedEvents = new Set();

    schedData.forEach(row => {
        let rowDateStr = this._formatDate(row[1]);
        if (!rowDateStr) return;
        let rowRegion = row[5] ? String(row[5]).trim() : 'Onshore';
        if (regionFilter !== 'All' && rowRegion !== regionFilter) return;

        if (rowDateStr >= searchStartStr && rowDateStr <= endStr) {
            let agent = String(row[0]).trim();
            let rY = parseInt(rowDateStr.substring(0,4));
            let rM = parseInt(rowDateStr.substring(5,7));
            let rD = parseInt(rowDateStr.substring(8,10));
            if (isNaN(rY) || isNaN(rM) || isNaN(rD)) return;

            let startMins = this._timeToMins(row[3]);
            let endMinsRaw = this._timeToMins(row[4]);
            
            let actSlice = String(row[2]).trim().substring(0, 10);
            let eventHash = `${agent}_${rowDateStr}_${startMins}_${endMinsRaw}_${actSlice}`;
            if (processedEvents.has(eventHash)) return; 
            processedEvents.add(eventHash);

            let endMins = endMinsRaw < startMins ? endMinsRaw + 1440 : endMinsRaw; 

            this._getShiftSplits(startMins, endMins).forEach(split => {
                let splitStartEpoch = new Date(rY, rM-1, rD, Math.floor(split.startMins/60), split.startMins%60, 0, 0).getTime();

                if (splitStartEpoch >= bounds.start && splitStartEpoch <= bounds.end) {
                    if ((mode === 'month' || mode === 'quarter') && cycleFilter !== 'ALL') {
                        if (this._getCycleForEpoch(splitStartEpoch) !== cycleFilter) return;
                    }
                    
                    let effDateStr = Utilities.formatDate(new Date(splitStartEpoch), "America/Toronto", "yyyy-MM-dd");

                    if (trackerType === 'furlough' && mode === 'day') {
                        for (let min = split.startMins; min < split.endMins; min += 15) {
                            let blockTime = new Date(rY, rM-1, rD, Math.floor(min/60), min%60, 0, 0).getTime();
                            if (blockTime >= bounds.start && blockTime <= bounds.end) {
                                let idx = Math.floor((min % 1440) / 15);
                                if (idx >= 0 && idx < 96) buckets[idx].supply = Math.max(0, buckets[idx].supply - 1);
                            }
                        }
                    }

                    let rawStartStr = String(row[3]).trim();
                    let groupKey = `${agent}_${effDateStr}_${split.shift}_${rawStartStr}`;
                    
                    let actName = "Time Off";
                    if (trackerType === 'coaching' || trackerType === 'roles') actName = row[2];
                    groupKey += `_${actName}`;

                    if (!groupedLogs[groupKey]) {
                        groupedLogs[groupKey] = { 
                            date: effDateStr, agent: agent, 
                            activityName: actName, 
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
        
        // Setup agent memory with regional tags and tower
        const getAg = (name, reg) => {
            let n = String(name).replace(/\b\w/g, c => c.toUpperCase()).trim();
            if(!report.agents[n]) report.agents[n] = { name: n, region: reg || 'Onshore', acsu: 0, coach: 0, safe: 0, icl: 0, ulc: 0, tower: 0, total: 0 };
            if (reg && String(reg).includes('Offshore')) report.agents[n].region = 'Offshore';
            return report.agents[n];
        };

        let searchStart = new Date(bounds.start); searchStart.setDate(searchStart.getDate() - 1);
        const sStr = Utilities.formatDate(searchStart, "America/Toronto", "yyyy-MM-dd");
        const eStr = Utilities.formatDate(new Date(bounds.end), "America/Toronto", "yyyy-MM-dd");

        const parseDB = (sheetName, metricName) => {
            const db = this._getDB(sheetName);
            if (!db || db.getLastRow() < 2) return;
            
            let processedEvents = new Set();
            db.getDataRange().getDisplayValues().slice(1).forEach(row => {
                let dStr = this._formatDate(row[1]);
                if (dStr >= sStr && dStr <= eStr) {
                    let rY = parseInt(dStr.substring(0,4)); let rM = parseInt(dStr.substring(5,7)); let rD = parseInt(dStr.substring(8,10));
                    if(isNaN(rY) || isNaN(rM) || isNaN(rD)) return;
                    
                    let agent = String(row[0]).trim();
                    let sMins = this._timeToMins(row[3]); let eMinsR = this._timeToMins(row[4]);
                    let region = row[5] ? String(row[5]).trim() : 'Onshore';
                    
                    let actSlice = String(row[2]).trim().substring(0, 10);
                    let eventHash = `${agent}_${dStr}_${sMins}_${eMinsR}_${actSlice}`;
                    if (processedEvents.has(eventHash)) return; processedEvents.add(eventHash);

                    let eMins = eMinsR < sMins ? eMinsR + 1440 : eMinsR; 
                    this._getShiftSplits(sMins, eMins).forEach(s => {
                        let epoch = new Date(rY, rM-1, rD, Math.floor(s.startMins/60), s.startMins%60, 0, 0).getTime();
                        if (epoch >= bounds.start && epoch <= bounds.end) { 
                            getAg(agent, region)[metricName] += s.hours; 
                            getAg(agent, region).total += s.hours;
                        }
                    });
                }
            });
        };

        parseDB('WF_COACHING', 'coach');
        parseDB('WF_FURLOUGH', 'acsu');
        
        // PURE WFM ROLES (SAFE / TOWER)
        const dbRoles = this._getDB('WF_ROLES');
        if (dbRoles && dbRoles.getLastRow() > 1) {
            let processedRoles = new Set();
            dbRoles.getDataRange().getDisplayValues().slice(1).forEach(row => {
                let dStr = this._formatDate(row[1]);
                if (dStr >= sStr && dStr <= eStr) {
                    let rY = parseInt(dStr.substring(0,4)); let rM = parseInt(dStr.substring(5,7)); let rD = parseInt(dStr.substring(8,10));
                    if(isNaN(rY) || isNaN(rM) || isNaN(rD)) return;
                    
                    let agent = String(row[0]).trim();
                    let sMins = this._timeToMins(row[3]); let eMinsR = this._timeToMins(row[4]);
                    let region = row[5] ? String(row[5]).trim() : 'Onshore';
                    
                    let actSlice = String(row[2]).trim().substring(0, 10);
                    let eventHash = `${agent}_${dStr}_${sMins}_${eMinsR}_${actSlice}`;
                    if (processedRoles.has(eventHash)) return; processedRoles.add(eventHash);

                    let eMins = eMinsR < sMins ? eMinsR + 1440 : eMinsR; 
                    let roleType = String(row[2]).toUpperCase();
                    
                    this._getShiftSplits(sMins, eMins).forEach(s => {
                        let epoch = new Date(rY, rM-1, rD, Math.floor(s.startMins/60), s.startMins%60, 0, 0).getTime();
                        if (epoch >= bounds.start && epoch <= bounds.end) { 
                            if (roleType.includes('SAFE')) { getAg(agent, region).safe += s.hours; getAg(agent, region).total += s.hours; }
                            else if (roleType.includes('TOWER') || roleType.includes('WOFQT')) { getAg(agent, region).tower += s.hours; getAg(agent, region).total += s.hours; }
                            // Note: Also captures WFM ICL/ULC if someone schedules them, just in case
                            else if (roleType.includes('ICL')) { getAg(agent, region).icl += s.hours; getAg(agent, region).total += s.hours; }
                            else if (roleType.includes('ULC') || roleType.includes('FIRE')) { getAg(agent, region).ulc += s.hours; getAg(agent, region).total += s.hours; }
                        }
                    });
                }
            });
        }

        // RESTORED MANUAL TRACKER (DB_SESSIONS) FOR ICL & ULC (IGNORES SAFE)
        const dbSess = this._getDB('DB_Sessions');
        if (dbSess && dbSess.getLastRow() > 1) {
            dbSess.getDataRange().getDisplayValues().slice(1).forEach(row => {
                let sessionEpoch = new Date(row[3]).getTime();
                if (sessionEpoch >= bounds.start && sessionEpoch <= bounds.end) {
                    let agentName = String(row[1]).trim();
                    let role = String(row[2]).toUpperCase(); 
                    let h = Number(row[6]) || 0; 
                    
                    // We DO NOT add SAFE here, because SAFE is now fully automated via WFM
                    if (role.includes('ICL')) { 
                        getAg(agentName).icl += h; 
                        getAg(agentName).total += h; 
                    }
                    else if (role.includes('ULC') || role.includes('FIRE')) { 
                        getAg(agentName).ulc += h; 
                        getAg(agentName).total += h; 
                    }
                }
            });
        }

        // DEEP GEM INTEGRATION (V3 DB)
        const dbGEM = this._getDB('WF_GEM_DATA_V3');
        let gemData = {};
        if (dbGEM && dbGEM.getLastRow() > 1) {
            dbGEM.getDataRange().getValues().slice(1).forEach(row => {
                let agName = String(row[1]).replace(/\b\w/g, c => c.toUpperCase()).trim();
                gemData[agName] = { 
                    cph: row[2], inPct: row[3], outPct: row[4], aht: row[5],
                    transfers: row[10] || 0, transfPct: row[11] || 0,
                    avgTalk: row[12] || '-', avgHold: row[13] || '-', avgAcw: row[14] || '-',
                    outCalls: row[15] || 0, outTalk: row[16] || '-',
                    extCalls: row[17] || 0, extTalk: row[18] || '-'
                };
            });
        }

        let finalArr = Object.values(report.agents).filter(a => a.total > 0 || gemData[a.name]);
        
        finalArr.forEach(a => { 
            ['acsu','coach','safe','icl','ulc','tower','total'].forEach(k => a[k] = parseFloat(a[k].toFixed(2))); 
            
            if (gemData[a.name]) {
                a.gem = gemData[a.name]; // Attach entire nested GEM object for the UI Card
                a.cph = parseFloat(gemData[a.name].cph).toFixed(2);
                a.aht = gemData[a.name].aht;
                a.inOut = Math.round(parseFloat(gemData[a.name].inPct) * 100) + "% / " + Math.round(parseFloat(gemData[a.name].outPct) * 100) + "%";
            } else {
                a.gem = null; a.cph = '-'; a.aht = '-'; a.inOut = '-';
            }
        });
        
        report.data = finalArr.sort((a,b) => b.total - a.total);
        return JSON.stringify(report);
  },

  archiveUnifiedWeek: function(refDateStr) {
      const reportStr = this.getUnifiedWeeklyReport(refDateStr);
      const report = JSON.parse(reportStr);
      const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
      
      const writeToArchiveSheet = (spreadsheet) => {
          let archiveSheet = spreadsheet.getSheetByName('Weekly_Archives_V3');
          if (!archiveSheet) { 
              archiveSheet = spreadsheet.insertSheet('Weekly_Archives_V3');
              archiveSheet.appendRow(["Timestamp", "Cycle", "Period", "Agent Name", "Region", "Total Off-Phone", "JSON Payload"]); 
              archiveSheet.getRange(1,1,1,7).setFontWeight("bold");
          }
          const data = archiveSheet.getDataRange().getValues();
          for (let i = data.length - 1; i >= 1; i--) { if (data[i][2] === report.period) archiveSheet.deleteRow(i + 1); }
          
          const rows = report.data.map(a => [ new Date(), report.cycle, report.period, a.name, a.region, a.total, JSON.stringify(a) ]);
          if (rows.length > 0) archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
      };

      writeToArchiveSheet(ssLocal);
      return `Successfully archived ${report.data.length} agents for ${report.cycle} (${report.period}).`;
  },

  getArchiveList: function() {
      let sheet = this._getDB('Weekly_Archives_V3');
      if (!sheet || sheet.getLastRow() < 2) return "[]";
      
      const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getDisplayValues();
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
      let sheet = this._getDB('Weekly_Archives_V3');
      if (!sheet || sheet.getLastRow() < 2) return "{}";
      
      const data = sheet.getDataRange().getDisplayValues();
      let report = { cycle: "ARCHIVED", period: targetPeriod, data: [] };
      
      for (let i = 1; i < data.length; i++) {
          if (data[i][2] === targetPeriod) {
              report.cycle = data[i][1];
              try { report.data.push(JSON.parse(data[i][6])); } catch(e) {}
          }
      }
      report.data.sort((a,b) => b.total - a.total);
      return JSON.stringify(report);
  },

  _getShiftSplits: function(startMins, endMins) {
      let splits = []; let current = startMins;
      while (current < endMins) {
          let timeOfDay = current % 1440;
          let shiftType = "Night"; let nextBound = current - timeOfDay + 420;
          if (timeOfDay >= 420 && timeOfDay < 900) { shiftType = "Morning"; nextBound = current - timeOfDay + 900; } 
          else if (timeOfDay >= 900 && timeOfDay < 1380) { shiftType = "Evening"; nextBound = current - timeOfDay + 1380; } 
          else if (timeOfDay >= 1380) nextBound = current - timeOfDay + 1860;
          
          let chunkEnd = Math.min(endMins, nextBound);
          splits.push({ shift: shiftType, hours: (chunkEnd - current) / 60, startMins: current, endMins: chunkEnd });
          current = chunkEnd;
      }
      return splits;
  },
  
  _minsToTime: function(mins) { let m = mins % 1440; let h = Math.floor(m / 60); let mm = m % 60; return `${h < 10 ? '0'+h : h}:${mm < 10 ? '0'+mm : mm}`; },
  _timeToMins: function(tStr) {
       if (tStr == null || tStr === "") return 0;
       let s = String(tStr).trim();
       let num = Number(s);
       if (!isNaN(num) && s !== "" && !s.includes(":")) { let m = Math.round(num * 1440); return m >= 1440 ? m % 1440 : m; }
       if (tStr instanceof Date) { s = Utilities.formatDate(tStr, "America/Toronto", 'HH:mm'); }
       let match = s.match(/(\d{1,2})[:\.](\d{2})\s*([AP]M)?/i);
       if (!match) return 0;
       let h = parseInt(match[1]), m = parseInt(match[2]), amp = match[3] ? match[3].toUpperCase() : null;
       if (amp === 'PM' && h < 12) h += 12;
       if (amp === 'AM' && h === 12) h = 0; 
       return (h * 60) + m;
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
  
  _parseDate: function(s) { return this._formatDate(s); },
  _formatDate: function(d) { 
      if (d == null || d === "") return "";
      if (d instanceof Date) return Utilities.formatDate(d, "America/Toronto", "yyyy-MM-dd");
      let s = String(d).trim();
      let num = Number(s);
      if (!isNaN(num) && num > 30000) {
          let date = new Date((num - 25569) * 86400 * 1000);
          date.setMinutes(date.getMinutes() + date.getTimezoneOffset());
          return Utilities.formatDate(date, "America/Toronto", "yyyy-MM-dd");
      }
      let isoMatch = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
      if (isoMatch) return `${isoMatch[1]}-${isoMatch[2].padStart(2,'0')}-${isoMatch[3].padStart(2,'0')}`;
      let regMatch = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
      if (regMatch) { let p1 = parseInt(regMatch[1]), p2 = parseInt(regMatch[2]); let m = p1 > 12 ? p2 : p1; let day = p1 > 12 ? p1 : p2; return `${regMatch[3]}-${String(m).padStart(2,'0')}-${String(day).padStart(2,'0')}`; }
      let pDate = new Date(s);
      if (!isNaN(pDate)) return Utilities.formatDate(pDate, "America/Toronto", "yyyy-MM-dd");
      return s.substring(0, 10); 
  },
  
  _formatTimeStr: function(t) { let d=new Date(`2000/01/01 ${t}`); return isNaN(d)?t:Utilities.formatDate(d, "America/Toronto", 'HH:mm'); },
  _cleanActivity: function(s) { return s.replace(/\d{2}\s?[AP]M/gi, '').trim(); },
  _timeToBucket: function(val) { let mins = this._timeToMins(val); return mins < 0 ? -1 : Math.floor(mins / 15); },
  _indexToTime: function(i) { let h=Math.floor(i/4), m=(i%4)*15; return `${h<10?'0'+h:h}:${m===0?'00':m}`; }
};

function fetchSyncMetadata() {
  try {
    const props = PropertiesService.getDocumentProperties();
    return JSON.stringify({
      sched: props.getProperty('SYNC_SCHED') || "Awaiting WFM Sync...",
      idp: props.getProperty('SYNC_IDP') || "Awaiting IDP Sync...",
      muSet: props.getProperty('MU_SET') || ""
    });
  } catch(e) {
    return JSON.stringify({ sched: "Metadata unavailable", idp: "Metadata unavailable", muSet: "" });
  }
}
