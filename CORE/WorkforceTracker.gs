/**
 * MODULE: WORKFORCE TRACKER
 * Features: Shift Splitter Engine & UI Aggregator
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

    // --- SCHEDULE PARSER ---
    if (schedRaw && schedRaw.trim().length > 0) {
      let cleanSched = [];
      const lines = schedRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      let currentAgent = "", currentDateStr = "";
      let currentY = 0, currentM = 0, currentD = 0;
      let lastTimeMins = -1, daysAdded = 0;
      
      const segmentRegex = /([a-zA-ZÀ-ÿ0-9\/\(\)\s\-\.&]+?)\s+(\d{1,2}:\d{2}(?:\s?[AP]M)?)\s+(\d{1,2}:\d{2}(?:\s?[AP]M)?)\s*$/i;
      const dateRegex = /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/;

      lines.forEach(line => {
        let text = line.trim();
        if(text.startsWith('Agent Name') || text.startsWith('"Agent Name"')) return;

        if (text.includes('Agent:')) {
          let parts = text.split(':');
          if (parts.length > 1) currentAgent = parts[1].replace(/^\s*\d+\s+/, '').trim();
          return;
        } else if (text.includes('"') && text.includes(',')) {
          let csvParts = this._parseCSVLine(text);
          if (csvParts.length >= 5) { cleanSched.push([csvParts[0], this._parseDate(csvParts[1]), this._cleanActivity(csvParts[2]), csvParts[3], csvParts[4]]); return; }
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
               cleanSched.push([currentAgent, actDateStr, act, tStart, tEnd]);
            }
          }
        }
      });

      if (cleanSched.length > 0) {
        this._upsertData('WF_SCHEDULE', cleanSched, 1, ['Agent Name', 'Date', 'Activity', 'Start Time', 'End Time']);
        msg.push(`✔ Schedule: Imported ${cleanSched.length} blocks.`);
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

  getAnalytics: function(mode, refDate, trackerType) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbIDP = ss.getSheetByName('WF_IDP');
    const dbSched = ss.getSheetByName('WF_SCHEDULE');
    
    if (!dbIDP || !dbSched) return JSON.stringify({ error: "Database missing. Please run WFM Import first." });

    const dateObj = new Date(refDate + 'T00:00:00');
    let startDate = new Date(dateObj), endDate = new Date(dateObj), label = "";

    if (mode === 'day') {
      label = this._formatDate(startDate);
    } else if (mode === 'week') {
      let day = startDate.getDay(); let diff = startDate.getDate() - day + (day == 0 ? -6 : 1); 
      startDate.setDate(diff); endDate = new Date(startDate); endDate.setDate(startDate.getDate() + 6);
      label = `Week of ${this._formatDate(startDate)}`;
    } else if (mode === 'month') {
      startDate.setDate(1); endDate = new Date(startDate.getFullYear(), startDate.getMonth() + 1, 0);
      label = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MMMM yyyy");
    }

    const startStr = this._formatDate(startDate);
    const endStr = this._formatDate(endDate);

    const idpData = dbIDP.getLastRow() > 1 ? dbIDP.getDataRange().getValues().slice(1) : [];
    const schedData = dbSched.getLastRow() > 1 ? dbSched.getDataRange().getValues().slice(1) : [];

    let buckets = [];
    let combinedEvents = [];
    let groupedLogs = {}; // Aggregator

    if (mode === 'day' && trackerType === 'furlough') {
      buckets = Array.from({length: 96}, (_, i) => ({ index: i, label: this._indexToTime(i), supply: 0, demand: 0, net: 0 }));
      idpData.forEach(row => {
        if (this._formatDate(row[0]) === startStr) { 
          let idx = this._timeToBucket(row[1]);
          if (idx > -1) { buckets[idx].demand += Number(row[2] || 0); buckets[idx].supply += Number(row[3] || 0); }
        }
      });
    }

    const COACHING_CODES = ['ce séance', 'huddle', 'echo', 'mentor', 'hsc', 'health and safety', 'meet', 'roadshow', 'one on one', 'individuelle', 'pulsecheck', 'qual session', 'quality', 'sbys', 'survey', 'sondage en ligne', 'team'];
    const ACSU_CODES = ['acsu', 'solicited', 'libération', 'voluntary'];

    schedData.forEach(row => {
        let sDateStr = this._formatDate(row[1]);
        if (sDateStr >= startStr && sDateStr <= endStr) {
            let agent = String(row[0]).trim();
            let act = String(row[2]).toLowerCase();
            
            let isCoach = trackerType === 'coaching' && COACHING_CODES.some(c => act.includes(c));
            let isFurlough = trackerType === 'furlough' && ACSU_CODES.some(c => act.includes(c));

            if (isCoach || isFurlough) {
                let startMins = this._timeToMins(row[3]);
                let endMinsRaw = this._timeToMins(row[4]);
                let endMins = endMinsRaw < startMins ? endMinsRaw + 1440 : endMinsRaw; // Carry overnight

                // Split shifts at 07:00, 15:00, 23:00
                let splits = this._getShiftSplits(startMins, endMins);

                splits.forEach(split => {
                    // Grid Deduction
                    if (isFurlough && mode === 'day') {
                        let sStartBuck = Math.floor((split.startMins % 1440) / 15);
                        let sEndBuck = Math.floor((split.endMins % 1440) / 15);
                        let buckEnd = split.endMins > split.startMins && sEndBuck <= sStartBuck ? sEndBuck + 96 : sEndBuck;
                        
                        for (let i = sStartBuck; i < buckEnd; i++) {
                            let idx = i % 96;
                            buckets[idx].supply = Math.max(0, buckets[idx].supply - 1);
                        }
                    }

                    // Aggregation UI Logic
                    let groupKey = `${agent}_${sDateStr}_${split.shift}`;
                    if (isCoach) groupKey += `_${row[2]}`;

                    if (!groupedLogs[groupKey]) {
                        groupedLogs[groupKey] = {
                            date: sDateStr, agent: agent,
                            activityName: isCoach ? row[2] : "Time Off",
                            shift: split.shift, hours: split.hours,
                            timeStart: this._minsToTime(split.startMins),
                            timeEnd: this._minsToTime(split.endMins)
                        };
                    } else {
                        groupedLogs[groupKey].hours += split.hours;
                        groupedLogs[groupKey].timeEnd = this._minsToTime(split.endMins);
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

    return JSON.stringify({ mode: mode, trackerType: trackerType, label: label, grid: buckets, events: combinedEvents, totals: totals });
  },

  // --- ENGINE UTILS ---
  _getShiftSplits: function(startMins, endMins) {
      let splits = [];
      let current = startMins;
      while (current < endMins) {
          let shiftType = "";
          let nextBoundary = 0;
          let timeOfDay = current % 1440;
          
          if (timeOfDay >= 420 && timeOfDay < 900) { shiftType = "Morning"; nextBoundary = current - timeOfDay + 900; } 
          else if (timeOfDay >= 900 && timeOfDay < 1380) { shiftType = "Evening"; nextBoundary = current - timeOfDay + 1380; } 
          else {
              shiftType = "Night";
              if (timeOfDay >= 1380) nextBoundary = current - timeOfDay + 1860;
              else nextBoundary = current - timeOfDay + 420;
          }

          let chunkEnd = Math.min(endMins, nextBoundary);
          splits.push({ shift: shiftType, hours: (chunkEnd - current) / 60, startMins: current, endMins: chunkEnd });
          current = chunkEnd;
      }
      return splits;
  },
  _minsToTime: function(mins) {
      let m = mins % 1440; let h = Math.floor(m / 60); let mm = m % 60;
      let hh = h < 10 ? '0'+h : h; let mmm = mm < 10 ? '0'+mm : mm;
      return `${hh}:${mmm}`;
  },
  _upsertData: function(sheetName, newRows, dateColIdx, headersArray) {
    const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
    this._executeUpsert(ssLocal, sheetName, newRows, dateColIdx, headersArray);
    if (typeof MasterConnector !== 'undefined' && MasterConnector.DB_ID) {
        try { const ssMaster = SpreadsheetApp.openById(MasterConnector.DB_ID); this._executeUpsert(ssMaster, sheetName, newRows, dateColIdx, headersArray); } catch(e) {}
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
       if (amp === 'PM' && h < 12) h += 12; if (amp === 'AM' && h === 12) h = 0;
       return (h * 60) + m;
  },
  _parseCSVLine: function(text) {
    if (text.includes('\t')) return text.split('\t').map(s => s.trim());
    let ret = [], inQuote = false, token = "";
    for(let i=0; i<text.length; i++) {
      let char = text[i]; if(char === '"') { inQuote = !inQuote; continue; }
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
      if (amp === 'PM' && h < 12) h += 12; if (amp === 'AM' && h === 12) h = 0; return (h * 4) + Math.floor(m / 15);
    } return -1;
  },
  _indexToTime: function(i) { let h=Math.floor(i/4), m=(i%4)*15; return `${h<10?'0'+h:h}:${m===0?'00':m}`; }
};
