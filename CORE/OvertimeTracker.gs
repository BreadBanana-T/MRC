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

  // ───────────────────── OPEN SLOTS (WFM export: JSON or table) ─────────────────────
  // "Open" = what WFM posted. The export holds BOTH overtime offers (OTST/OT)
  // and ACSU "Solicited Time Off" (released time — the opposite of OT). Each
  // row is classified as OT / ACSU / OTHER; only OT rows count toward open-OT
  // totals, ACSU is tracked as its own released-time series.
  //
  // The WF_OT_OPEN sheet is human-readable, in the same column layout as the
  // flattened WFM table: Date, Start, End, Slots, Activity, Min Length,
  // Agent Data Group, Agent Data Values — followed by the computed columns.
  //
  // Accepted paste formats (auto-sniffed):
  //   1. Raw WFM JSON  — {"slots":[...]} (truncated pastes are salvaged)
  //   2. Tab-separated table — Date⇥Start Time⇥End Time⇥Slots⇥Activity⇥
  //      Min Length⇥Agent Data Group⇥Agent Data Value Required[⇥Skills]

  OPEN_SHEET: 'WF_OT_OPEN',
  OPEN_HEADERS: ['Date', 'Start Time', 'End Time', 'Slots', 'Activity', 'Min Length',
                 'Agent Data Group', 'Agent Data Values',
                 'Type', 'Rate', 'Req Group', 'Window Hours', 'Open Hours', 'Visible', 'OID'],

  // Requirement group display name. These are WFM "Agent Data Groups"
  // (who may take the slot), NOT skills: e.g. "ADT Monit Knowledge Level"
  // with values Junior/Senior/Regular/Expert, "ADT SAFE: Yes",
  // "Language: French Only". Empty group = any agent may take it.
  _skillName: function(adgName) {
    var n = String(adgName || '').trim();
    if (!n) return 'Any agent';
    if (/safe/i.test(n)) return 'SAFE';
    if (/knowledge/i.test(n)) return 'Knowledge Level';
    if (/language/i.test(n)) return 'Language';
    return n.replace(/^ADT\s+/i, '');
  },

  // "12:00 AM" / "4:00 PM" / "06:00" / "16:00:00" → minutes since midnight.
  _dispToMins: function(t) {
    var m = String(t || '').trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*([AP]\.?M\.?)?$/i);
    if (!m) return -1;
    var h = parseInt(m[1], 10), mm = parseInt(m[2], 10);
    var ap = m[3] ? m[3].toUpperCase().replace(/\./g, '') : null;
    if (ap === 'PM' && h < 12) h += 12;
    if (ap === 'AM' && h === 12) h = 0;
    return h * 60 + mm;
  },

  // "2026-06-30" stays; "6/30/26" → "2026-06-30".
  _parseDateLoose: function(d) {
    d = String(d || '').trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(d)) return d.substring(0, 10);
    var m = d.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (!m) return '';
    var y = parseInt(m[3], 10); if (y < 100) y += 2000;
    var mo = parseInt(m[1], 10), dy = parseInt(m[2], 10);
    return y + '-' + (mo < 10 ? '0' : '') + mo + '-' + (dy < 10 ? '0' : '') + dy;
  },

  // One classified sheet row, or null when unusable.
  _openRow: function(dateStr, startDisp, endDisp, slots, activity, minLen, adgName, advsStr, visible, oid) {
    var d = this._parseDateLoose(dateStr);
    var sMin = this._dispToMins(startDisp), eMin = this._dispToMins(endDisp);
    if (!d || sMin < 0 || eMin < 0) return null;
    if (eMin <= sMin) eMin += 1440; // 4:00 PM → 12:00 AM wraps midnight
    var winH = Math.round(((eMin - sMin) / 60) * 100) / 100;
    var n = parseInt(slots, 10) || 0;
    if (!n) return null;
    var act = String(activity || '').trim();
    var ot = this.decode(act);
    var isAcsu = /acsu|solicited|lib[ée]ration|voluntary/i.test(act);
    var type = ot ? 'OT' : (isAcsu ? 'ACSU' : 'OTHER');
    return [
      d, String(startDisp).trim(), String(endDisp).trim(), n, act, String(minLen || '').trim(),
      String(adgName || '').trim(), String(advsStr || '').trim(),
      type, ot ? ot.rate : '', this._skillName(adgName),
      winH, Math.round(winH * n * 100) / 100,
      visible === false ? 'N' : 'Y', String(oid)
    ];
  },

  // Pastes are routinely truncated mid-array. Salvage everything up to the
  // last complete slot object instead of failing.
  _tolerantParse: function(raw) {
    var txt = String(raw).trim().replace(/^```[a-z]*\s*/i, '').replace(/```\s*$/, '').trim();
    var tryParse = function(t) { try { return JSON.parse(t); } catch (e) { return null; } };
    var obj = tryParse(txt);
    if (!obj) {
      var idx = txt.lastIndexOf('},{"date"');
      if (idx > -1) obj = tryParse(txt.substring(0, idx + 1) + ']}');
    }
    if (!obj) {
      var arrIdx = txt.indexOf('[');
      if (arrIdx > -1) {
        var t2 = txt.substring(arrIdx);
        var o2 = tryParse(t2);
        if (!o2) {
          var idx2 = t2.lastIndexOf('},{"date"');
          if (idx2 > -1) o2 = tryParse(t2.substring(0, idx2 + 1) + ']');
        }
        if (o2) obj = { slots: o2 };
      }
    }
    if (!obj) return null;
    if (Array.isArray(obj)) return obj;
    if (obj.slots && Array.isArray(obj.slots)) return obj.slots;
    return null;
  },

  importOpenSlots: function(raw) {
    if (!raw || !String(raw).trim()) return 'OT Open: empty paste.';
    var self = this;
    var txt = String(raw).trim();
    var rows = [], dates = {}, seen = {};
    var push = function(row) {
      if (!row) return;
      if (seen[row[14]]) return; seen[row[14]] = true;
      rows.push(row); dates[row[0]] = true;
    };

    if (txt.indexOf('"slotCount"') !== -1 || /^[\[{]/.test(txt)) {
      // ── JSON path ──
      var slots = this._tolerantParse(txt);
      if (!slots || !slots.length) return 'OT Open: no slot objects found in the JSON.';
      slots.forEach(function(sl, i) {
        try {
          push(self._openRow(
            sl.date && sl.date.value,
            (sl.startTime && (sl.startTime.description || sl.startTime.value)) || '',
            (sl.endTime && (sl.endTime.description || sl.endTime.value)) || '',
            sl.slotCount,
            sl.activity && sl.activity.name,
            sl.minLength && sl.minLength.description,
            sl.adg && sl.adg.name,
            (sl.advs || []).map(function(a) { return a && a.value; }).filter(Boolean).join(' or '),
            sl.visible !== false,
            sl.oid || ('json-' + i)
          ));
        } catch (e) {}
      });
    } else {
      // ── Tab-separated table path (the flattened export) ──
      var lines = txt.split(/\r?\n/);
      var idx = 0;
      lines.forEach(function(line) {
        if (!line.trim()) return;
        if (/^date\s*\t/i.test(line)) return; // header row
        var f = line.split('\t');
        if (f.length < 5) return;
        push(self._openRow(f[0], f[1], f[2], f[3], f[4], f[5], f[6], f[7],
                           true, 'tsv-' + self._parseDateLoose(f[0]) + '-' + (idx++)));
      });
      if (!rows.length) return 'OT Open: table detected but no usable rows (need Date⇥Start⇥End⇥Slots⇥Activity columns).';
    }
    if (!rows.length) return 'OT Open: nothing usable in the paste.';

    // Replace-by-date: a re-export is the current truth for those days.
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(this.OPEN_SHEET);
    var W = this.OPEN_HEADERS.length;
    if (!sheet) {
      sheet = ss.insertSheet(this.OPEN_SHEET);
      sheet.appendRow(this.OPEN_HEADERS);
      sheet.getRange(1, 1, 1, W).setFontWeight('bold');
      try { sheet.setFrozenRows(1); } catch (e) {}
    } else if (String(sheet.getRange(1, 9).getValue()) !== 'Type' || String(sheet.getRange(1, 11).getValue()) !== 'Req Group') {
      // Old schema — wipe and rebuild (data is re-pastable by design).
      sheet.clear();
      sheet.appendRow(this.OPEN_HEADERS);
      sheet.getRange(1, 1, 1, W).setFontWeight('bold');
      try { sheet.setFrozenRows(1); } catch (e) {}
    }
    var keep = [];
    if (sheet.getLastRow() > 1) {
      var WT = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker : null;
      keep = sheet.getRange(2, 1, sheet.getLastRow() - 1, W).getDisplayValues()
        .filter(function(r) { return !dates[self._parseDateLoose(WT ? WT._formatDate(r[0]) : r[0])]; });
      sheet.getRange(2, 1, sheet.getLastRow() - 1, W).clearContent();
    }
    var all = keep.concat(rows);
    if (all.length) sheet.getRange(2, 1, all.length, W).setValues(all);

    var nOt = rows.filter(function(r) { return r[8] === 'OT'; }).length;
    var nAcsu = rows.filter(function(r) { return r[8] === 'ACSU'; }).length;
    return 'Open slots: ' + rows.length + ' window(s) across ' + Object.keys(dates).length + ' day(s) — ' +
           nOt + ' OT, ' + nAcsu + ' ACSU released-time' + (rows.length - nOt - nAcsu ? ', ' + (rows.length - nOt - nAcsu) + ' other' : '') + '.';
  },

  // Open-slot aggregation for the analytics window. Only Type=OT rows count
  // toward open-OT; ACSU released-time is reported separately.
  _openBlock: function(WT, bounds, mode, cycleFilter, givenByDay) {
    var open = { hours: 0, slots: 0, hidden: 0, skills: {}, acsuHours: 0, acsuSlots: 0, days: [], windows: [] };
    var db = WT._getDB(this.OPEN_SHEET);
    if (!db || db.getLastRow() < 2) return null;
    var self = this;
    var sStr = Utilities.formatDate(new Date(bounds.start), 'America/Toronto', 'yyyy-MM-dd');
    var eStr = Utilities.formatDate(new Date(bounds.end), 'America/Toronto', 'yyyy-MM-dd');
    var byDay = {};
    var any = false;
    db.getDataRange().getDisplayValues().slice(1).forEach(function(r) {
      var d = self._parseDateLoose(WT._formatDate(r[0]));
      if (!d || d < sStr || d > eStr) return;
      if ((mode === 'month' || mode === 'quarter') && cycleFilter !== 'ALL') {
        var p = d.split('-');
        var epoch = new Date(parseInt(p[0],10), parseInt(p[1],10)-1, parseInt(p[2],10), 12, 0, 0).getTime();
        if (WT._getCycleForEpoch(epoch) !== cycleFilter) return;
      }
      var type = String(r[8]);
      var slots = parseInt(r[3], 10) || 0;
      var oh = parseFloat(r[12]) || 0;
      any = true;
      if (String(r[13]) === 'N') { open.hidden += slots; return; }
      if (type === 'ACSU') {
        open.acsuHours = Math.round((open.acsuHours + oh) * 100) / 100;
        open.acsuSlots += slots;
      } else if (type === 'OT') {
        open.hours += oh;
        open.slots += slots;
        var skill = String(r[10]) || 'Any agent';
        open.skills[skill] = Math.round(((open.skills[skill] || 0) + oh) * 100) / 100;
        if (!byDay[d]) byDay[d] = { open: 0, slots: 0 };
        byDay[d].open += oh;
        byDay[d].slots += slots;
      }
      if (open.windows.length < 100) {
        open.windows.push({ date: d, start: String(r[1]), end: String(r[2]), slots: slots, hours: oh,
                            group: String(r[10]) || 'Any agent', values: String(r[7] || ''), type: type });
      }
    });
    if (!any) return null;
    var allDays = {};
    Object.keys(byDay).forEach(function(d) { allDays[d] = true; });
    Object.keys(givenByDay).forEach(function(d) { allDays[d] = true; });
    open.days = Object.keys(allDays).sort().map(function(d) {
      return {
        date: d,
        open: byDay[d] ? Math.round(byDay[d].open * 100) / 100 : 0,
        slots: byDay[d] ? byDay[d].slots : 0,
        given: Math.round((givenByDay[d] || 0) * 100) / 100
      };
    });
    open.hours = Math.round(open.hours * 100) / 100;
    open.windows.sort(function(a, b) { return a.date.localeCompare(b.date) || a.start.localeCompare(b.start); });
    return open;
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

    var givenByDay = {};
    events.forEach(function (e) { givenByDay[e.date] = (givenByDay[e.date] || 0) + e.hours; });
    var open = null;
    try { open = this._openBlock(WT, bounds, mode, cycleFilter, givenByDay); } catch (e) {}

    return JSON.stringify({
      mode: mode, trackerType: 'overtime', label: bounds.label, cycle: bounds.cycle,
      grid: [], events: events, totals: totals, otTotals: otTotals, open: open
    });
  }
};

/**
 * Editor self-test: ingests two synthetic slot windows and reports.
 * Proves the server half (parser + WF_OT_OPEN sheet write) independently
 * of the web UI. Run from the Apps Script editor's function dropdown;
 * inspect the WF_OT_OPEN tab afterwards. The test rows use date
 * 2000-01-01 so they never pollute real reporting windows.
 */
function otOpenSelfTest() {
  var sample = JSON.stringify({ slots: [
    { date: { value: '2000-01-01' }, startTime: { value: '06:00:00.000' }, endTime: { value: '12:00:00.000' },
      slotCount: 2, activity: { name: 'OTST File Heures supp x1 / OT x1 On Queue' }, visible: true,
      adg: { name: 'ADT Monit Knowledge Level' }, advs: [{ value: 'Junior' }], oid: 'selftest-1' },
    { date: { value: '2000-01-01' }, startTime: { value: '16:00:00.000' }, endTime: { value: '00:00:00.000' },
      slotCount: 1, activity: { name: 'OTST File Heures supp x1 / OT x1 On Queue' }, visible: true,
      adg: { name: 'ADT SAFE' }, advs: [{ value: 'Yes' }], oid: 'selftest-2' }
  ] });
  var res = OvertimeTracker.importOpenSlots(sample);
  Logger.log(res);
  return res + ' — check the WF_OT_OPEN tab for two 2000-01-01 rows.';
}

// Router exports (called from Code.gs / google.script.run)
function importOvertimeData(schedRaw) {
  return (typeof OvertimeTracker !== 'undefined') ? OvertimeTracker.importFromSchedule(schedRaw) : 'Error';
}
function getOvertimeAnalytics(mode, refDate, region, cycle) {
  return (typeof OvertimeTracker !== 'undefined')
    ? OvertimeTracker.getAnalytics(mode, refDate, region, cycle)
    : JSON.stringify({ error: 'OvertimeTracker not loaded.' });
}
