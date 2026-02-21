/**
 * MODULE: WORKFORCE TRACKER
 * Handles Furlough & Coaching Analytics and Importing
 */

const WorkforceTracker = {

  // --- 1. IMPORT LOGIC ---
  importData: function(schedRaw, idpRaw) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ['WF_SCHEDULE', 'WF_IDP', 'WF_FURLOUGH'].forEach(n => {
       if(!ss.getSheetByName(n)) ss.insertSheet(n);
    });
    let msg = [];

    // Schedule Parser
    if (schedRaw && schedRaw.trim().length > 0) {
      let cleanSched = [];
      const lines = schedRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      let currentAgent = "", currentDate = "";
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
          if (csvParts.length >= 5) {
            cleanSched.push([csvParts[0], this._parseDate(csvParts[1]), this._cleanActivity(csvParts[2]), csvParts[3], csvParts[4]]);
            return; 
          }
        }

        let dMatch = text.match(dateRegex);
        if (dMatch) currentDate = this._parseDate(dMatch[1]);

        if (currentAgent && currentDate) {
          let segMatch = text.match(segmentRegex);
          if (segMatch) {
            let act = this._cleanActivity(segMatch[1].trim());
            if (!act.toLowerCase().match(/^activity|^scheduled/)) {
               cleanSched.push([currentAgent, currentDate, act, segMatch[2].trim(), segMatch[3].trim()]);
            }
          }
        }
      });

      if (cleanSched.length > 0) {
        this._upsertData('WF_SCHEDULE', cleanSched, 1, ['Agent Name', 'Date', 'Activity', 'Start Time', 'End Time']);
        msg.push(`✔ Schedule: Imported ${cleanSched.length} blocks.`);
      }
    }

    // IDP Parser
    if (idpRaw && idpRaw.trim().length > 0) {
      let cleanIDP = [];
      const lines = idpRaw.split(/\r?\n/).filter(l => l.trim().length > 0);
      let headerIdx = lines.findIndex(l => l.toLowerCase().includes('req') && l.toLowerCase().includes('open'));

      if (headerIdx > -1) {
        let headers = this._parseCSVLine(lines[headerIdx]);
        let colMap = {}; 
        headers.forEach((h, i) => {
          let lower = h.toLowerCase();
          let dateMatch = h.match(/(\w+\s\d{1,2},?\s\d{4})/);
          if (dateMatch) {
            let dateStr = this._parseDate(dateMatch[1]);
            if (lower.includes('req')) colMap[i] = { date: dateStr, type: 'req' };
            else if (lower.includes('open') && !lower.includes('+/-')) colMap[i] = { date: dateStr, type: 'open' };
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
        }
      }
    }
    return msg.length ? msg.join('\n') : "No valid data found to import.";
  },

  // --- 2. ANALYTICS API ---
  getAnalytics: function(mode, refDate, trackerType) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbIDP = ss.getSheetByName('WF_IDP');
    const dbSched = ss.getSheetByName('WF_SCHEDULE');
    
    if (!dbIDP || !dbSched) return JSON.stringify({ error: "Database missing. Please run WFM Import first." });

    // Date Bounds
    const dateObj = new Date(refDate + 'T00:00:00');
    let startDate = new Date(dateObj), endDate = new Date(dateObj), label = "";

    if (mode === 'day') {
      label = this._formatDate(startDate);
    } else if (mode === 'week') {
      let day = startDate.getDay();
      let diff = startDate.getDate() - day + (day == 0 ? -6 : 1); 
      startDate.setDate(diff); endDate = new Date(startDate); endDate.setDate(startDate.getDate() + 6);
      label = `Week of ${this._formatDate(startDate)}`;
    } else if (mode === 'month') {
      startDate.setDate(1); endDate = new Date(startDate.getFullYear(), startDate.getMonth() + 1, 0);
      label = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MMMM yyyy");
    }

    const startStr = this._formatDate(startDate);
    const endStr = this._formatDate(endDate);

    const idpData = dbIDP.getDataRange().getValues(); idpData.shift();
    const schedData = dbSched.getDataRange().getValues(); schedData.shift();

    let buckets = [];
    let combinedEvents = [];

    // --- COACHING MODE ---
    if (trackerType === 'coaching') {
        const COACHING_CODES = ['ce séance', 'huddle', 'echo', 'mentor', 'hsc', 'health and safety', 'meet', 'roadshow', 'one on one', 'individuelle', 'pulsecheck', 'qual session', 'quality', 'sbys', 'survey', 'sondage en ligne', 'team'];
        
        schedData.forEach(row => {
            let sDateStr = this._formatDate(row[1]);
            if (sDateStr >= startStr && sDateStr <= endStr) {
                let act = String(row[2]).toLowerCase();
                if (COACHING_CODES.some(c => act.includes(c))) {
                    let sStart = this._timeToBucket(row[3]);
                    combinedEvents.push({
                       date: sDateStr, agent: row[0],
                       time: `${this._formatTime(row[3])} - ${this._formatTime(row[4])}`,
                       hours: this._getDuration(row[3], row[4]),
                       shift: this._getShift(sStart),
                       activityName: row[2]
                    });
                }
            }
        });
    } 
    // --- FURLOUGH MODE ---
    else {
        if (mode === 'day') {
          buckets = Array.from({length: 96}, (_, i) => ({ index: i, label: this._indexToTime(i), supply: 0, demand: 0, net: 0 }));
          idpData.forEach(row => {
            if (this._formatDate(row[0]) === startStr) { 
              let idx = this._timeToBucket(row[1]);
              if (idx > -1) { buckets[idx].demand += Number(row[2] || 0); buckets[idx].supply += Number(row[3] || 0); }
            }
          });
        }

        const ACSU_CODES = ['acsu', 'solicited', 'libération', 'voluntary'];
        let prodMap = {};
        
        schedData.forEach(row => {
          let sDateStr = this._formatDate(row[1]);
          if (sDateStr >= startStr && sDateStr <= endStr) {
              let agent = String(row[0]).trim();
              let act = String(row[2]).toLowerCase();
              let sStart = this._timeToBucket(row[3]);
              let sEnd = this._timeToBucket(row[4]);
              if (sEnd < sStart) sEnd = 96;

              if (ACSU_CODES.some(c => act.includes(c))) {
                   combinedEvents.push({
                     type: 'auto', date: sDateStr, agent: row[0],
                     time: `${this._formatTime(row[3])} - ${this._formatTime(row[4])}`,
                     hours: this._getDuration(row[3], row[4]),
                     shift: this._getShift(sStart), startIdx: sStart, endIdx: sEnd
                   });
              } else if (!['break', 'lunch', 'off', 'sick', 'maladie'].some(ex => act.includes(ex))) {
                   if (!prodMap[sDateStr]) prodMap[sDateStr] = {};
                   if (!prodMap[sDateStr][agent]) prodMap[sDateStr][agent] = [];
                   prodMap[sDateStr][agent].push({ start: sStart, end: sEnd });
              }
          }
        });

        // Deduct Supply
        if (mode === 'day') {
           combinedEvents.forEach(f => {
              for (let i = f.startIdx; i < f.endIdx; i++) if (i >= 0 && i < 96) buckets[i].supply = Math.max(0, buckets[i].supply - 1);
           });
           buckets.forEach(b => { b.net = parseFloat((b.supply - b.demand).toFixed(2)); });
        }
    }

    let totals = { all: 0, morning: 0, evening: 0, night: 0, count: combinedEvents.length };
    combinedEvents.forEach(f => {
        totals.all += f.hours;
        if (f.shift === 'Morning') totals.morning += f.hours;
        else if (f.shift === 'Evening') totals.evening += f.hours;
        else totals.night += f.hours;
    });

    return JSON.stringify({
      mode: mode, trackerType: trackerType, label: label,
      grid: buckets, events: combinedEvents, totals: totals
    });
  },

  // --- UTILS ---
  _upsertData: function(sheetName, newRows, dateColIdx, headersArray) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const incomingDates = new Set(newRows.map(row => String(row[dateColIdx]).trim()));
    const existingData = sheet.getDataRange().getValues();
    const headers = existingData.length > 0 ? existingData.shift() : headersArray;
    const retainedRows = existingData.filter(row => !incomingDates.has(String(this._parseDate(row[dateColIdx])).trim()));
    const combined = retainedRows.concat(newRows);
    sheet.clearContents(); sheet.appendRow(headers);
    if (combined.length > 0) sheet.getRange(2, 1, combined.length, combined[0].length).setValues(combined);
  },
  _parseCSVLine: function(text) {
    if (text.includes('\t')) return text.split('\t').map(s => s.trim());
    let ret = [], inQuote = false, token = "";
    for(let i=0; i<text.length; i++) {
      let char = text[i];
      if(char === '"') { inQuote = !inQuote; continue; }
      if(char === ',' && !inQuote) { ret.push(token.trim()); token = ""; }
      else token += char;
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
      if (amp === 'PM' && h < 12) h += 12; if (amp === 'AM' && h === 12) h = 0;
      return (h * 4) + Math.floor(m / 15);
    } return -1;
  },
  _indexToTime: function(i) { let h=Math.floor(i/4), m=(i%4)*15; return `${h<10?'0'+h:h}:${m===0?'00':m}`; },
  _getShift: function(idx) { if (idx >= 28 && idx < 60) return 'Morning'; if (idx >= 60 && idx < 92) return 'Evening'; return 'Night'; },
  _getDuration: function(t1, t2) {
      let parseTimeObj = str => {
        let parts = String(str).match(/(\d+):(\d+)\s?([AP]M)?/i);
        if (parts) { let d=new Date(), h=parseInt(parts[1]), m=parseInt(parts[2]), amp=parts[3]?parts[3].toUpperCase():null; if(amp==='PM'&&h<12) h+=12; if(amp==='AM'&&h===12) h=0; d.setHours(h,m,0,0); return d;}
        return null;
      };
      let d1 = (t1 instanceof Date)?t1:parseTimeObj(t1), d2 = (t2 instanceof Date)?t2:parseTimeObj(t2);
      if(!d1||!d2) return 0;
      let diff = d2 - d1; if(diff<0) diff+=(24*60*60*1000);
      return parseFloat((diff / 3600000).toFixed(2));
  }
};
