/**
 * MODULE: OVERTIME TRACKER  (WFM / IEX schedule-derived overtime)
 *
 * Deliberately isolated in its own file so the overtime pay logic -- which is
 * liability-sensitive -- can be audited and changed without touching the
 * furlough / coaching / roles classification in WorkforceTracker.gs.
 *
 * WHAT IT DOES
 *   - Re-parses the SAME pasted WFM schedule (its own independent pass) and
 *     extracts only the overtime-coded activity segments.
 *   - Onshore agents only (for now). Offshore (TI) is skipped, mirroring the
 *     furlough tracker's "!isOffshore" gate.
 *   - Writes to its own sheet "WF_OVERTIME" via a destructive upsert keyed by
 *     agent+date, so re-pasting a corrected schedule self-heals.
 *
 * THE CODE TAXONOMY  (3 axes -- see OT_CODES below)
 *   Rate   : OTST... = x1 (straight time) | bare OT... (no ST) = x1.5 (premium)
 *   Bucket : OFQ/ODQ = Off Queue | OTHR = Other | SAFE = SAFE | else On Queue
 *   Break  : trailing CB = a paid OT break
 *
 * PAY RULE (per product owner): every code present in the schedule counts as
 * paid OT -- including the CB break codes. Breaks are NOT broken out; they fold
 * into their rate/bucket total. Anything not coded is unpaid and not counted.
 * Hours are reported split by rate (x1 / x1.5) and by queue bucket; there is
 * no dollar/cost figure -- this purely tallies overtime hours.
 */

var OvertimeTracker = {

  SHEET: 'WF_OVERTIME',
  HEADERS: ['Agent Name', 'Date', 'Code', 'Rate', 'Bucket', 'IsBreak', 'Start Time', 'End Time', 'Region'],

  // Full code table. "c" is matched as a whole token (case-insensitive). The
  // list is sorted longest-first at call time so OTSTOTHRCB wins over OTST,
  // and OTST wins over OT -- critical because every code shares the "OT" stem.
  OT_CODES: [
    { c: 'OTSTOTHRCB', rate: 1.0, bucket: 'Other',    brk: true  },
    { c: 'OTSTOFQCB',  rate: 1.0, bucket: 'OffQueue', brk: true  },
    { c: 'OTOTHRCB',   rate: 1.5, bucket: 'Other',    brk: true  },
    { c: 'OTOFQCB',    rate: 1.5, bucket: 'OffQueue', brk: true  },
    { c: 'OTSTOTHR',   rate: 1.0, bucket: 'Other',    brk: false },
    { c: 'OTSTOFQ',    rate: 1.0, bucket: 'OffQueue', brk: false },
    { c: 'OTSTODQ',    rate: 1.0, bucket: 'OffQueue', brk: false },
    { c: 'OTST SAFE',  rate: 1.0, bucket: 'SAFE',     brk: false },
    { c: 'OTSTSAFE',   rate: 1.0, bucket: 'SAFE',     brk: false },
    { c: 'OTSTCB',     rate: 1.0, bucket: 'OnQueue',  brk: true  },
    { c: 'OTOTHR',     rate: 1.5, bucket: 'Other',    brk: false },
    { c: 'OTOFQ',      rate: 1.5, bucket: 'OffQueue', brk: false },
    { c: 'OTST',       rate: 1.0, bucket: 'OnQueue',  brk: false },
    { c: 'OTCB',       rate: 1.5, bucket: 'OnQueue',  brk: true  },
    { c: 'OT',         rate: 1.5, bucket: 'OnQueue',  brk: false }
  ],

  /**
   * Decode one activity string into overtime dimensions.
   * Handles BOTH the short IEX codes (e.g. "OTSTOFQ", "OTST SAFE") AND the
   * descriptive French/English phrasing (e.g. "Hors file Heures supp x1.5").
   * Returns { code, rate, bucket, brk } or null when the activity is not OT.
   */
  decode: function (raw) {
    if (raw == null) return null;
    var sDot = String(raw).toUpperCase();
    // Spaced/normalized form for whole-word matching ("OTST SAFE" survives).
    var s = ' ' + sDot.replace(/[^A-Z0-9]+/g, ' ').replace(/\s+/g, ' ').trim() + ' ';
    // Compact alpha-run for the "starts with OT code" guard against false hits
    // like "OTHER MEETING" (which is not in the table, so it won't match).
    var compact = sDot.replace(/[^A-Z]/g, '');

    // 1) Short-code path -- try each known code as a standalone token, longest
    //    first. This is the most reliable signal when present.
    var codes = this.OT_CODES.slice().sort(function (a, b) {
      return b.c.replace(/\s/g, '').length - a.c.replace(/\s/g, '').length;
    });
    for (var i = 0; i < codes.length; i++) {
      var token = ' ' + codes[i].c + ' ';
      if (s.indexOf(token) !== -1) {
        return { code: codes[i].c, rate: codes[i].rate, bucket: codes[i].bucket, brk: codes[i].brk };
      }
    }

    // 2) Description path -- only engage when the French/English overtime
    //    wording is present, so non-OT activities never leak in.
    var hasPhrase = /HEURES?\s*SUPP/.test(sDot) || /\bOVERTIME\b/.test(s) || /^OT(\b|\d)/.test(compact);
    if (!hasPhrase) return null;

    // Rate: explicit 1.5 wins; explicit x1 / 1x = straight; bare OT = premium.
    var rate;
    if (/1\s*[.,]\s*5/.test(sDot)) rate = 1.5;
    else if (/\bX?\s*1\s*X?\b/.test(s) || /OTST/.test(compact)) rate = 1.0;
    else rate = 1.5;

    // Bucket: Off Queue (Hors file) checked before File so "Hors file" can't
    // fall through to the On Queue default.
    var bucket;
    if (/HORS\s*FIL/.test(sDot) || /\bOFQ\b/.test(s) || /\bODQ\b/.test(s) || /OFF\s*QUEUE/.test(sDot) || /\bOFFQ\b/.test(s)) bucket = 'OffQueue';
    else if (/\bAUTRE\b/.test(s) || /\bOTHR\b/.test(s) || /\bOTHER\b/.test(s)) bucket = 'Other';
    else if (/\bSAFE\b/.test(s)) bucket = 'SAFE';
    else bucket = 'OnQueue';

    var brk = /\bCB\b/.test(s) || /\bPAUSE\b/.test(s) || /\bBREAK\b/.test(s);

    return { code: 'OT?', rate: rate, bucket: bucket, brk: brk };
  },

  // Human-readable label used in the UI log + chart.
  label: function (rate, bucket) {
    var b = bucket === 'OffQueue' ? 'Off Queue' : (bucket === 'OnQueue' ? 'On Queue' : bucket);
    return b + ' x' + (rate === 1.5 ? '1.5' : '1');
  },

  /**
   * Independent parse pass over the pasted schedule. Extracts OT rows only and
   * writes them to WF_OVERTIME. Wrapped by the caller in try/catch so a parse
   * problem here can never break the main furlough/coaching import.
   */
  importFromSchedule: function (schedRaw) {
    if (!schedRaw || !schedRaw.trim()) return 'No schedule text.';
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null;
    if (!WT) return 'WorkforceTracker helpers unavailable.';

    var self = this;
    var cleanOT = [];

    var segmentRegex = /([a-zA-Z\u00C0-\u00FF0-9\/\(\)\s\-\.&']+?)\s+(\d{1,2}:\d{2}(?:\s?[AP]M)?)\s+(\d{1,2}:\d{2}(?:\s?[AP]M)?)\s*$/i;
    var dateRegex = /(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/;

    var lines = schedRaw.split(/\r?\n/).filter(function (l) { return l.trim().length > 0; });

    var currentAgent = '', currentID = '';
    var currentY = 0, currentM = 0, currentD = 0;
    var lastTimeMins = -1, daysAdded = 0;
    var agentBuffer = [];

    // Resolve onshore/offshore using the same hierarchy as the main import:
    // manual/masterlist registry wins; else ID-prefix-3 / TI keyword.
    var isOffshoreFor = function (name, id, rawTextFlag) {
      var off = false;
      if (id && String(id).charAt(0) === '3') off = true;
      if (rawTextFlag) off = true;
      if (typeof RegionRegistry !== 'undefined' && name) {
        var reg = RegionRegistry.getRegion(name);
        var src = RegionRegistry.getSource(name);
        if (reg && (src === 'manual' || src === 'masterlist')) off = (reg === 'Offshore');
        else if (reg) off = (reg === 'Offshore');
      }
      return off;
    };

    var flushAgentBuffer = function () {
      if (agentBuffer.length === 0) return;
      var keywordOffshore = false;
      for (var b = 0; b < agentBuffer.length; b++) {
        var up = agentBuffer[b].raw.toUpperCase();
        if (up.indexOf('TI ') !== -1 || up.indexOf('OFFSHORE') !== -1) { keywordOffshore = true; break; }
      }
      var offshore = isOffshoreFor(currentAgent, currentID, keywordOffshore);
      if (!offshore) {
        agentBuffer.forEach(function (obj) {
          var ot = self.decode(obj.act);
          if (ot) {
            cleanOT.push([currentAgent, obj.dateStr, ot.code, ot.rate, ot.bucket,
                          ot.brk ? 'Y' : 'N', obj.start, obj.end, 'Onshore']);
          }
        });
      }
      agentBuffer = [];
    };

    lines.forEach(function (line) {
      var text = line.trim();
      if (text.indexOf('Agent Name') === 0 || text.indexOf('"Agent Name"') === 0) return;

      // Agent header line.
      if (text.indexOf('Agent:') !== -1) {
        flushAgentBuffer();
        var parts = text.split(':');
        if (parts.length > 1) {
          var agentData = parts[1].trim();
          var idMatch = agentData.match(/^(\d+)/);
          currentID = idMatch ? idMatch[1] : '';
          currentAgent = agentData.replace(/^\d+\s+/, '').trim();
        }
        return;
      }

      // CSV / TSV row:  Agent, Date, Activity, Start, End, Region
      if (text.indexOf('"') !== -1 && text.indexOf(',') !== -1) {
        flushAgentBuffer();
        var csv = WT._parseCSVLine(text);
        if (csv.length >= 6) {
          var act = WT._cleanActivity(csv[2]);
          var ot = self.decode(act);
          if (ot) {
            var rowName = csv[0];
            var rowOff = csv[5] && csv[5].indexOf('Offshore') !== -1;
            if (typeof RegionRegistry !== 'undefined' && rowName) {
              var reg = RegionRegistry.getRegion(rowName);
              var src = RegionRegistry.getSource(rowName);
              if (reg && (src === 'manual' || src === 'masterlist')) rowOff = (reg === 'Offshore');
              else if (reg) rowOff = (reg === 'Offshore');
            }
            if (!rowOff) {
              cleanOT.push([rowName, WT._parseDate(csv[1]), ot.code, ot.rate, ot.bucket,
                            ot.brk ? 'Y' : 'N', csv[3], csv[4], 'Onshore']);
            }
          }
        }
        return;
      }

      // Date marker line (block format).
      var dMatch = text.match(dateRegex);
      if (dMatch) {
        var ds = WT._parseDate(dMatch[1]);
        var p = ds.split('-').map(Number);
        currentY = p[0]; currentM = p[1]; currentD = p[2];
        lastTimeMins = -1; daysAdded = 0;
      }

      // Activity segment line (block format).
      if (currentAgent && currentY) {
        var segMatch = text.match(segmentRegex);
        if (segMatch) {
          var actName = WT._cleanActivity(segMatch[1].trim());
          var tStartMins = WT._timeToMins(segMatch[2].trim());
          if (lastTimeMins > -1 && tStartMins < lastTimeMins) daysAdded++;
          lastTimeMins = tStartMins;
          var actDateStr = Utilities.formatDate(
            new Date(currentY, currentM - 1, currentD + daysAdded),
            Session.getScriptTimeZone(), 'yyyy-MM-dd');
          if (!actName.toLowerCase().match(/^activity|^scheduled/)) {
            agentBuffer.push({ raw: text, dateStr: actDateStr, act: actName,
                               start: segMatch[2].trim(), end: segMatch[3].trim() });
          }
        }
      }
    });
    flushAgentBuffer();

    if (cleanOT.length > 0) {
      WT._executeDestructiveUpsert(this.SHEET, cleanOT, this.HEADERS);
      return 'Overtime: ' + cleanOT.length + ' segment(s).';
    }
    return 'Overtime: none found.';
  },

  /**
   * Analytics engine for the Overtime tracker UI. Shaped to match
   * WorkforceTracker.getAnalytics so the existing renderTracker can reuse it,
   * plus an "otTotals" block with the rate (x1/x1.5) and bucket breakdown.
   */
  getAnalytics: function (mode, refDate, regionFilter, cycleFilter) {
    regionFilter = regionFilter || 'Onshore';
    cycleFilter = cycleFilter || 'ALL';
    var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null;
    if (!WT) return JSON.stringify({ error: 'Engine unavailable.' });

    var db = WT._getDB(this.SHEET);
    if (!db || db.getLastRow() < 2) {
      return JSON.stringify({ error: 'No overtime data yet. Paste a WFM schedule containing OT codes.' });
    }

    var bounds = WT._calculateEpochBoundaries(mode, refDate);
    var searchStart = new Date(bounds.start); searchStart.setDate(searchStart.getDate() - 1);
    var sStr = Utilities.formatDate(searchStart, 'America/Toronto', 'yyyy-MM-dd');
    var eStr = Utilities.formatDate(new Date(bounds.end), 'America/Toronto', 'yyyy-MM-dd');

    var rows = db.getDataRange().getDisplayValues().slice(1);
    var grouped = {};
    var seen = {};

    rows.forEach(function (row) {
      var agent = String(row[0]).trim();
      var dStr = WT._formatDate(row[1]);
      if (!dStr || dStr < sStr || dStr > eStr) return;

      var region = row[8] ? String(row[8]).trim() : 'Onshore';
      if (regionFilter !== 'All' && region !== regionFilter) return;

      var rY = parseInt(dStr.substring(0, 4), 10);
      var rM = parseInt(dStr.substring(5, 7), 10);
      var rD = parseInt(dStr.substring(8, 10), 10);
      if (isNaN(rY) || isNaN(rM) || isNaN(rD)) return;

      var rate = parseFloat(row[3]) || 1.0;
      var bucket = String(row[4]).trim() || 'OnQueue';
      var startMins = WT._timeToMins(row[6]);
      var endRaw = WT._timeToMins(row[7]);

      var dedup = agent + '|' + dStr + '|' + row[2] + '|' + startMins + '|' + endRaw;
      if (seen[dedup]) return; seen[dedup] = true;

      var endMins = endRaw < startMins ? endRaw + 1440 : endRaw;

      WT._getShiftSplits(startMins, endMins).forEach(function (split) {
        var epoch = new Date(rY, rM - 1, rD, Math.floor(split.startMins / 60), split.startMins % 60, 0, 0).getTime();
        if (epoch < bounds.start || epoch > bounds.end) return;
        if ((mode === 'month' || mode === 'quarter') && cycleFilter !== 'ALL') {
          if (WT._getCycleForEpoch(epoch) !== cycleFilter) return;
        }
        var effDate = Utilities.formatDate(new Date(epoch), 'America/Toronto', 'yyyy-MM-dd');
        var key = agent + '|' + effDate + '|' + rate + '|' + bucket + '|' + split.shift + '|' + String(row[6]).trim();
        if (!grouped[key]) {
          grouped[key] = {
            date: effDate, agent: agent, rate: rate, bucket: bucket,
            shift: split.shift, hours: split.hours,
            timeStart: WT._minsToTime(split.startMins), timeEnd: WT._minsToTime(split.endMins)
          };
        } else {
          grouped[key].hours += split.hours;
          grouped[key].timeEnd = WT._minsToTime(split.endMins);
        }
      });
    });

    var self = this;
    var events = Object.keys(grouped).map(function (k) {
      var g = grouped[k];
      return {
        date: g.date, agent: g.agent, shift: g.shift,
        rate: g.rate, bucket: g.bucket,
        activityName: self.label(g.rate, g.bucket),
        hours: parseFloat(g.hours.toFixed(2)),
        time: g.timeStart + ' - ' + g.timeEnd
      };
    });

    var totals = { all: 0, morning: 0, evening: 0, night: 0, count: events.length };
    var otTotals = { x1: 0, x15: 0, OnQueue: 0, OffQueue: 0, Other: 0, SAFE: 0 };
    events.forEach(function (e) {
      totals.all += e.hours;
      if (e.shift === 'Morning') totals.morning += e.hours;
      else if (e.shift === 'Evening') totals.evening += e.hours;
      else totals.night += e.hours;

      if (e.rate === 1.5) otTotals.x15 += e.hours; else otTotals.x1 += e.hours;
      if (otTotals[e.bucket] !== undefined) otTotals[e.bucket] += e.hours;
    });
    otTotals.x1 = parseFloat(otTotals.x1.toFixed(2));
    otTotals.x15 = parseFloat(otTotals.x15.toFixed(2));
    ['OnQueue', 'OffQueue', 'Other', 'SAFE'].forEach(function (b) { otTotals[b] = parseFloat(otTotals[b].toFixed(2)); });
    ['all', 'morning', 'evening', 'night'].forEach(function (k) { totals[k] = parseFloat(totals[k].toFixed(2)); });

    events.sort(function (a, b) { return (a.date.localeCompare(b.date)) || (b.hours - a.hours); });

    return JSON.stringify({
      mode: mode, trackerType: 'overtime', label: bounds.label, cycle: bounds.cycle,
      grid: [], events: events, totals: totals, otTotals: otTotals
    });
  }
};

// Router exports (called from Code.gs / google.script.run)
function importOvertimeData(schedRaw) {
  return (typeof OvertimeTracker !== 'undefined') ? OvertimeTracker.importFromSchedule(schedRaw) : 'Error';
}
function getOvertimeAnalytics(mode, refDate, region, cycle) {
  return (typeof OvertimeTracker !== 'undefined')
    ? OvertimeTracker.getAnalytics(mode, refDate, region, cycle)
    : JSON.stringify({ error: 'OvertimeTracker not loaded.' });
}
