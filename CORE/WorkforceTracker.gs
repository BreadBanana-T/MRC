/**
 * MODULE: WORKFORCE TRACKER
 * Features: Ultra-Lean Import Filter, WFM Monthly Calendar, & Epoch Boundaries
 */

const WorkforceTracker = {

  importData: function(schedRaw, idpRaw) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ['WF_SCHEDULE', 'WF_IDP', 'WF_FURLOUGH'].forEach(n => {
       if(!ss.getSheetByName(n)) ss.insertSheet(n);
    });
    if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
        try {
            const masterSS = SpreadsheetApp.openById(MasterConnector.DB_ID);
            ['WF_SCHEDULE', 'WF_IDP', 'WF_FURLOUGH'].forEach(n => {
               if(!masterSS.getSheetByName(n)) masterSS.insertSheet(n);
            });
        } catch(e) {}
    }
    
    let msg = [];
    
    // --- SCHEDULE PARSER (ULTRA-LEAN) ---
    if (schedRaw && schedRaw.trim().length > 0) {
      let cleanSched = [];
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
              
              if (isCoach || isFurlough) {
                  cleanSched.push([currentAgent, obj.dateStr, obj.act, obj.start, obj.end, reg]); 
              }
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
              if (isCoach || isFurlough) {
                  cleanSched.push([csvParts[0], this._parseDate(csvParts[1]), this._cleanActivity(csvParts[2]), csvParts[3], csvParts[4], csvParts[5]]); 
              }
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
            let tStart = segMatch[2].trim();
            let tEnd = segMatch[3].trim();
            
            let tStartMins = this._timeToMins(tStart);
            if (lastTimeMins > -1 && tStartMins < lastTimeMins) daysAdded++;
            lastTimeMins = tStartMins;
            let actDate = new Date(currentY, currentM - 1, currentD + daysAdded);
            let actDateStr = Utilities.formatDate(actDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
            
            if (!act.toLowerCase().match(/^activity|^scheduled/)) {
               agentBuffer.push({ raw: text, dateStr: actDateStr, act: act, start: tStart, end: tEnd });
            }
          }
        }
      });
      flushAgentBuffer(); 
      
      if (cleanSched.length > 0) {
        this._upsertData('WF_SCHEDULE', cleanSched, 1, ['Agent Name', 'Date', 'Activity', 'Start Time', 'End Time', 'Region']);
        msg.push(`✔ Schedule: Imported ${cleanSched.length} tracked events (Junk filtered).`);
      }
    }

    // --- IDP PARSER ---
    if (idpRaw && idpRaw.trim().length > 0) {
      let cleanIDP = [];
      const lines = idpRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      
      let headerIdx = lines.findIndex(l => {
         const lower = l.toLowerCase();
         return (lower.includes('req') || lower.includes('besoin')) && (lower.includes('open') || lower.includes('ouvert') || lower.includes('dispo'));
      });
      if (headerIdx > -1) {
        let headers = this._parseCSVLine(lines[headerIdx]);
        let colMap = {};
        headers.forEach((h, i) => {
          let lower = h.toLowerCase();
          let dateMatch = h.match(/(\w+\s\d{1,2},?\s\d{4})/);
          if (dateMatch) {
            let dateStr = this._parseDate(dateMatch[1]);
            if (lower.includes('req') || lower.includes('besoin')) colMap[i] = { date: dateStr, type: 'req' };
            else if ((lower.includes('open') || lower.includes('ouvert')) && !lower.includes('+/-')) colMap[i] = { date: dateStr, type: 'open' };
          }
        });
        let dataByDay = {};
        for (let i = headerIdx + 1; i < lines.length; i++) {
          let cols = this._parseCSVLine(lines[i]);
          let timeStr = cols[0]; 
          if (timeStr && timeStr.includes(':')) {
            let tNorm = this._formatTimeStr(timeStr);
            Object.keys(colMap).forEach(idx => {
               if (cols[idx] !== undefined) {
                 let info = colMap[idx];
                 if (!dataByDay[info.date]) dataByDay[info.date] = {};
                 if (!dataByDay[info.date][tNorm]) dataByDay[info.date][tNorm] = { req:0, open:0 };
                 let val = parseFloat(String(cols[idx]).replace(/,/g, '')) || 0;
                 if (info.type === 'req') dataByDay[info.date][tNorm].req = val;
                 if (info.type === 'open') dataByDay[info.date][tNorm].open = val;
               }
            });
          }
        }
        
        Object.keys(dataByDay).forEach(day => {
          Object.keys(dataByDay[day]).forEach(time => {
            cleanIDP.push([day, time, dataByDay[day][time].req, dataByDay[day][time].open]);
          });
        });
        if (cleanIDP.length > 0) {
          this._upsertData('WF_IDP', cleanIDP, 0, ['Day', 'Interval', 'Required', 'Open']);
          msg.push(`✔ IDP: Imported ${cleanIDP.length} intervals.`);
        } else msg.push(`❌ IDP found headers but no grid data matched.`);
      } else msg.push(`❌ IDP missing valid headers (Requirements/Open).`);
    }
    return msg.length ? msg.join('\n') : "No valid data found to import.";
  },

  _getCycleForEpoch: function(epoch) {
      // Anchored to Jan 28, 2026 @ 23:00 (Week A)
      let baseA = new Date(2026, 0, 28, 23, 0, 0, 0).getTime();
      let diffWeeks = Math.floor((epoch - baseA) / (7 * 24 * 60 * 60 * 1000));
      return (Math.abs(diffWeeks) % 2 === 1) ? "WEEK B" : "WEEK A";
  },

  // --- MATHEMATICAL EPOCH ENGINE (WFM Calendar) ---
  _calculateEpochBoundaries: function(mode, refDateStr) {
      let [tY, tM, tD] = refDateStr.split('-').map(Number);
      let tObj = new Date(tY, tM-1, tD, 12, 0, 0, 0); 
      
      let targetStartEpoch = 0, targetEndEpoch = 0, label = "", cycleName = "";

      if (mode === 'day') {
          let dObj = new Date(tY, tM-1, tD, 0, 0, 0, 0);
          targetStartEpoch = dObj.getTime();
          targetEndEpoch = targetStartEpoch + 86400000;
          label = this._formatDate(dObj);
          cycleName = this._getCycleForEpoch(targetStartEpoch);
      } 
      else if (mode === 'week') {
          let dayOfWeek = tObj.getDay(); 
          let diffToThu = (dayOfWeek >= 4) ? (4 - dayOfWeek) : -(dayOfWeek + 3);
          let thuDate = new Date(tObj);
          thuDate.setDate(tObj.getDate() + diffToThu);
          
          let wedDate = new Date(thuDate);
          wedDate.setDate(thuDate.getDate() - 1);
          wedDate.setHours(23, 0, 0, 0); // Exact 23:00 Boundary
          
          targetStartEpoch = wedDate.getTime();
          targetEndEpoch = targetStartEpoch + (7 * 24 * 60 * 60 * 1000);
          
          cycleName = this._getCycleForEpoch(targetStartEpoch);
          label = `${Utilities.formatDate(new Date(targetStartEpoch), Session.getScriptTimeZone(), "MMM dd")} to ${Utilities.formatDate(new Date(targetEndEpoch), Session.getScriptTimeZone(), "MMM dd")}`;
      } 
      else if (mode === 'month') {
          // TRUE WFM MONTH: Finds the Wed 23:00 before the 1st of the month
          let firstOfMonth = new Date(tY, tM-1, 1, 12, 0, 0);
          let day = firstOfMonth.getDay();
          let offsetToWed = (day >= 3) ? (3 - day) : -(day + 4);
          if (offsetToWed > 0) offsetToWed -= 7;
          let startWed = new Date(firstOfMonth);
          startWed.setDate(firstOfMonth.getDate() + offsetToWed);
          startWed.setHours(23, 0, 0, 0);
          
          targetStartEpoch = startWed.getTime();
          
          // Finds the next Wed 23:00 boundary for the end of the month
          let firstOfNextMonth = new Date(tY, tM, 1, 12, 0, 0);
          let dayNext = firstOfNextMonth.getDay();
          let offsetToNextWed = (dayNext >= 3) ? (3 - dayNext) : -(dayNext + 4);
          if (offsetToNextWed > 0) offsetToNextWed -= 7;
          let endWed = new Date(firstOfNextMonth);
          endWed.setDate(firstOfNextMonth.getDate() + offsetToNextWed);
          endWed.setHours(23, 0, 0, 0);
          
          targetEndEpoch = endWed.getTime();
          label = `WFM Month: ${Utilities.formatDate(new Date(targetStartEpoch), Session.getScriptTimeZone(), "MMM dd")} to ${Utilities.formatDate(new Date(targetEndEpoch), Session.getScriptTimeZone(), "MMM dd")}`;
          cycleName = "MONTH"; // Month acts as an aggregate
      }

      return { start: targetStartEpoch, end: targetEndEpoch, label: label, cycle: cycleName, startStr: Utilities.formatDate(new Date(targetStartEpoch), Session.getScriptTimeZone(), "yyyy-MM-dd") };
  },

  getAnalytics: function(mode, refDate, trackerType, regionFilter = 'All', cycleFilter = 'ALL') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbIDP = ss.getSheetByName('WF_IDP');
    const dbSched = ss.getSheetByName('WF_SCHEDULE');
    if (!dbIDP || !dbSched) return JSON.stringify({ error: "Database missing. Please run WFM Import first." });

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
        if (this._formatDate(row[0]) === bounds.startStr) { 
          let idx = this._timeToBucket(row[1]);
          if (idx > -1) { buckets[idx].demand += Number(row[2] || 0); buckets[idx].supply += Number(row[3] || 0); }
        }
      });
    }

    const COACHING_CODES = ['ce séance', 'huddle', 'echo', 'mentor', 'hsc', 'health and safety', 'meet', 'roadshow', 'one on one', 'individuelle', 'pulsecheck', 'qual session', 'quality', 'sbys', 'survey', 'sondage en ligne', 'team'];
    const ACSU_CODES = ['acsu', 'solicited', 'libération', 'voluntary']; 

    schedData.forEach(row => {
        let rowDateStr = this._formatDate(row[1]);
        let rowRegion = row[5] ? String(row[5]).trim() : 'Onshore';
        if (regionFilter !== 'All' && rowRegion !== regionFilter) return;

        if (rowDateStr >= searchStartStr && rowDateStr <= endStr) {
            let agent = String(row[0]).trim();
            let act = String(row[2]).toLowerCase();
            
            let isTeamLead = act.includes('team lead') || act.includes('équipe') || act.includes('equipe');
            let isCoach = trackerType === 'coaching' && !isTeamLead && COACHING_CODES.some(c => act.includes(c));
            let isFurlough = trackerType === 'furlough' && ACSU_CODES.some(c => act.includes(c));

            if (isCoach || isFurlough) {
                let [rY, rM, rD] = rowDateStr.split('-').map(Number);
                let startMins = this._timeToMins(row[3]);
                let endMinsRaw = this._timeToMins(row[4]);
                let endMins = endMinsRaw < startMins ? endMinsRaw + 1440 : endMinsRaw; 

                let splits = this._getShiftSplits(startMins, endMins);

                splits.forEach(split => {
                    let sObj = new Date(rY, rM-1, rD);
                    sObj.setHours(Math.floor(split.startMins/60), split.startMins%60, 0, 0);
                    let splitStartEpoch = sObj.getTime();

                    // EXACT MATHEMATICAL BOUNDARY
                    if (splitStartEpoch >= bounds.start && splitStartEpoch < bounds.end) {
                        
                        // MONTHLY SMART FILTER: Skips data that doesn't match the selected Week A/B view
                        if (mode === 'month' && cycleFilter !== 'ALL') {
                            let eventCycle = this._getCycleForEpoch(splitStartEpoch);
                            if (eventCycle !== cycleFilter) return;
                        }
                        
                        let effDateStr = Utilities.formatDate(sObj, Session.getScriptTimeZone(), "yyyy-MM-dd");

                        if (isFurlough && mode === 'day') {
                            for (let min = split.startMins; min < split.endMins; min += 15) {
                                let blockObj = new Date(rY, rM-1, rD);
                                blockObj.setHours(Math.floor(min/60), min%60, 0, 0);
                                if (blockObj.getTime() >= bounds.start && blockObj.getTime() < bounds.end) {
                                    let idx = Math.floor((min % 1440) / 15);
                                    if (idx >= 0 && idx < 96) buckets[idx].supply = Math.max(0, buckets[idx].supply - 1);
                                }
                            }
                        }

                        let groupKey = `${agent}_${effDateStr}_${split.shift}`;
                        if (isCoach) groupKey += `_${row[2]}`;

                        if (!groupedLogs[groupKey]) {
                            groupedLogs[groupKey] = {
                                date: effDateStr, agent: agent,
                                activityName: isCoach ? row[2] : "Time Off",
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
        }
    });

    if (mode === 'day' && trackerType === 'furlough') {
        buckets.forEach(b => { b.net = parseFloat((b.supply - b.demand).toFixed(2)); });
    }

    combinedEvents = Object.values(groupedLogs).map(g => ({
        date: g.date, agent: g.agent, activityName: g.activityName, shift: g.shift,
        hours: parseFloat(g.hours.toFixed(2)), time: `${g.timeStart} - ${g.timeEnd}`
    }));
    
    let totals = { all: 0, morning: 0, evening: 0, night: 0, count: combinedEvents.length };
    combinedEvents.forEach(f => {
        totals.all += f.hours;
        if (f.shift === 'Morning') totals.morning += f.hours;
        else if (f.shift === 'Evening') totals.evening += f.hours;
        else totals.night += f.hours;
    });
    
    return JSON.stringify({ mode: mode, trackerType: trackerType, label: bounds.label, cycle: bounds.cycle, grid: buckets, events: combinedEvents, totals: totals });
  },

  getUnifiedWeeklyReport: function(refDateStr) {
        const bounds = this._calculateEpochBoundaries("week", refDateStr);
        let report = { cycle: bounds.cycle, period: bounds.label, agents: {} };

        const getAg = (name) => {
            let n = String(name).trim();
            n = n.replace(/\b\w/g, c => c.toUpperCase()); 
            if(!report.agents[n]) report.agents[n] = { name: n, acsu: 0, coach: 0, safe: 0, icl: 0, ulc: 0, total: 0 };
            return report.agents[n];
        };

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        
        const dbSched = ss.getSheetByName('WF_SCHEDULE');
        if (dbSched && dbSched.getLastRow() > 1) {
            const schedData = dbSched.getDataRange().getValues().slice(1);
            const COACHING_CODES = ['ce séance', 'huddle', 'echo', 'mentor', 'hsc', 'health and safety', 'meet', 'roadshow', 'one on one', 'individuelle', 'pulsecheck', 'qual session', 'quality', 'sbys', 'survey', 'sondage en ligne', 'team'];
            const ACSU_CODES = ['acsu', 'solicited', 'libération', 'voluntary'];

            let searchStart = new Date(bounds.start); searchStart.setDate(searchStart.getDate() - 1);
            const searchStartStr = Utilities.formatDate(searchStart, Session.getScriptTimeZone(), "yyyy-MM-dd");
            const endStr = Utilities.formatDate(new Date(bounds.end), Session.getScriptTimeZone(), "yyyy-MM-dd");

            schedData.forEach(row => {
                let rowDateStr = this._formatDate(row[1]);
                if (rowDateStr >= searchStartStr && rowDateStr <= endStr) {
                    let [rY, rM, rD] = rowDateStr.split('-').map(Number);
                    let startMins = this._timeToMins(row[3]);
                    let endMinsRaw = this._timeToMins(row[4]);
                    let endMins = endMinsRaw < startMins ? endMinsRaw + 1440 : endMinsRaw; 
                    
                    this._getShiftSplits(startMins, endMins).forEach(split => {
                        let sObj = new Date(rY, rM-1, rD);
                        sObj.setHours(Math.floor(split.startMins/60), split.startMins%60, 0, 0);
                        let splitStartEpoch = sObj.getTime();
                        
                        if (splitStartEpoch >= bounds.start && splitStartEpoch < bounds.end) {
                            let act = String(row[2]).toLowerCase();
                            let isTeamLead = act.includes('team lead') || act.includes('équipe') || act.includes('equipe');
                            
                            if (ACSU_CODES.some(c => act.includes(c))) {
                                getAg(row[0]).acsu += split.hours; getAg(row[0]).total += split.hours;
                            } else if (!isTeamLead && COACHING_CODES.some(c => act.includes(c))) {
                                getAg(row[0]).coach += split.hours; getAg(row[0]).total += split.hours;
                            }
                        }
                    });
                }
            });
        }

        const dbSess = ss.getSheetByName('DB_Sessions');
        if (dbSess && dbSess.getLastRow() > 1) {
            const sessData = dbSess.getDataRange().getValues().slice(1);
            sessData.forEach(row => {
                let sStart = new Date(row[3]);
                let sessionEpoch = sStart.getTime();
                
                if (sessionEpoch >= bounds.start && sessionEpoch < bounds.end) {
                    let role = String(row[2]).toUpperCase();
                    let hours = Number(row[6]) || 0; 
                    
                    if (role.includes('SAFE')) { getAg(row[1]).safe += hours; getAg(row[1]).total += hours; }
                    else if (role.includes('ICL')) { getAg(row[1]).icl += hours; getAg(row[1]).total += hours; }
                    else if (role.includes('ULC') || role.includes('FIRE')) { getAg(row[1]).ulc += hours; getAg(row[1]).total += hours; }
                }
            });
        }

        let finalArr = Object.values(report.agents).filter(a => a.total > 0);
        finalArr.forEach(a => {
            a.acsu = parseFloat(a.acsu.toFixed(2));
            a.coach = parseFloat(a.coach.toFixed(2));
            a.safe = parseFloat(a.safe.toFixed(2));
            a.icl = parseFloat(a.icl.toFixed(2));
            a.ulc = parseFloat(a.ulc.toFixed(2));
            a.total = parseFloat(a.total.toFixed(2));
        });
        
        report.data = finalArr.sort((a,b) => b.total - a.total);
        return JSON.stringify(report);
  },

  archiveUnifiedWeek: function(refDateStr) {
      const reportStr = this.getUnifiedWeeklyReport(refDateStr);
      const report = JSON.parse(reportStr);
      
      const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
      let archiveSheet = ssLocal.getSheetByName('Weekly_Archives');
      if (!archiveSheet) {
          archiveSheet = ssLocal.insertSheet('Weekly_Archives');
          archiveSheet.appendRow(["Timestamp", "Cycle", "Period", "Agent Name", "ACSU", "Coaching", "SAFE", "ICL", "ULC FIRE", "Total Off-Phone"]);
          archiveSheet.getRange(1,1,1,10).setFontWeight("bold");
      }

      const data = archiveSheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
          if (data[i][2] === report.period) archiveSheet.deleteRow(i + 1);
      }

      const now = new Date();
      const rowsToPush = report.data.map(a => [
          now, report.cycle, report.period, a.name, a.acsu, a.coach, a.safe, a.icl, a.ulc, a.total
      ]);
      if (rowsToPush.length > 0) archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToPush.length, 10).setValues(rowsToPush);

      return `Successfully archived ${rowsToPush.length} agents for ${report.cycle} (${report.period}).`;
  },

  _getShiftSplits: function(startMins, endMins) {
      let splits = [];
      let current = startMins;
      while (current < endMins) {
          let shiftType = "";
          let nextBoundary = 0;
          let timeOfDay = current % 1440;
          if (timeOfDay >= 420 && timeOfDay < 900) { shiftType = "Morning"; nextBoundary = current - timeOfDay + 900; } 
          else if (timeOfDay >= 900 && timeOfDay < 1380) { shiftType = "Evening"; nextBoundary = current - timeOfDay + 1380; } 
          else { shiftType = "Night"; if (timeOfDay >= 1380) nextBoundary = current - timeOfDay + 1860; else nextBoundary = current - timeOfDay + 420; }

          let chunkEnd = Math.min(endMins, nextBoundary);
          splits.push({ shift: shiftType, hours: (chunkEnd - current) / 60, startMins: current, endMins: chunkEnd });
          current = chunkEnd;
      }
      return splits;
  },
  _minsToTime: function(mins) {
      let m = mins % 1440; let h = Math.floor(m / 60); let mm = m % 60;
      return `${h < 10 ? '0'+h : h}:${mm < 10 ? '0'+mm : mm}`;
  },
  _upsertData: function(sheetName, newRows, dateColIdx, headersArray) {
    const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
    this._executeUpsert(ssLocal, sheetName, newRows, dateColIdx, headersArray);
    if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
        try { const ssMaster = SpreadsheetApp.openById(MasterConnector.DB_ID);
        this._executeUpsert(ssMaster, sheetName, newRows, dateColIdx, headersArray); } catch(e) {}
    }
  },
  _executeUpsert: function(targetSpreadsheet, sheetName, newRows, dateColIdx, headersArray) {
    let sheet = targetSpreadsheet.getSheetByName(sheetName);
    if (!sheet) sheet = targetSpreadsheet.insertSheet(sheetName);
    const incomingDates = new Set(newRows.map(row => String(row[dateColIdx]).trim()));
    let existingData = [];
    if (sheet.getLastRow() > 0) existingData = sheet.getDataRange().getValues();
    if (existingData.length === 1 && existingData[0].join('') === "") existingData = [];
    const headers = existingData.length > 0 ? existingData.shift() : headersArray;
    const retainedRows = existingData.filter(row => {
       if(!row[dateColIdx]) return true; return !incomingDates.has(String(this._parseDate(row[dateColIdx])).trim());
    });
    const combined = retainedRows.concat(newRows);
    sheet.clearContents(); sheet.appendRow(headers);
    if (combined.length > 0) sheet.getRange(2, 1, combined.length, combined[0].length).setValues(combined);
  },
  _timeToMins: function(tStr) {
       let match = String(tStr).match(/(\d{1,2}):(\d{2})\s*([AP]M)?/i);
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
  _parseDate: function(s) { let d=new Date(s); return isNaN(d)?s:Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); },
  _formatDate: function(d) { return (d instanceof Date) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(d).substring(0,10); },
  _formatTimeStr: function(t) { let d=new Date(`2000/01/01 ${t}`); return isNaN(d)?t:Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm'); },
  _formatTime: function(d) { return (d instanceof Date) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "HH:mm") : String(d); },
  _cleanActivity: function(s) { return s.replace(/\d{2}\s?[AP]M/gi, '').trim(); },
  _timeToBucket: function(val) {
    if (!val) return -1;
    if (val instanceof Date) return (val.getHours()*4) + Math.floor(val.getMinutes()/15);
    let parts = String(val).match(/(\d+):(\d+)\s?([AP]M)?/i);
    if (parts) {
      let h = parseInt(parts[1]), m = parseInt(parts[2]), amp = parts[3] ? parts[3].toUpperCase() : null;
      if (amp === 'PM' && h < 12) h += 12;
      if (amp === 'AM' && h === 12) h = 0; return (h * 4) + Math.floor(m / 15);
    } return -1;
  },
  _indexToTime: function(i) { let h=Math.floor(i/4), m=(i%4)*15; return `${h<10?'0'+h:h}:${m===0?'00':m}`; }
};
