/**
 * MODULE: BREAK PLANNER  (server side of the standalone "MRC Workforce" app)
 *
 * The app lives at <web app URL>?app=workforce — its own page, outside the
 * dashboard, but in this same Apps Script project so it reads the exact sheets
 * MRC OPS already maintains. Nothing new has to be imported: the WFM/IEX paste
 * keeps going through MRC's Import box.
 *
 *   Reads   Raw Schedule / Schedule_History  shifts + IEX blocks (Breaks JSON)
 *           Break_Overrides                  floor edits (they ARE the current plan)
 *           Status_Overrides + fast map      absences marked on the floor
 *           WF_LANG_MAP, WF_MASTERLIST       language, ERC level, Question Line, site
 *           Genesys_Assignments              Escalation / Questions queues
 *           WF_IDP                           IDP Required per 15 min
 *   Writes  WF_BREAK_PLANS                   drafts + approved plans (one row per agent)
 *           Break_Overrides (on approve)     so the floor timers show the new times
 *
 * The optimizer itself runs in the browser (WorkforceEngine.html). This module
 * only assembles the day and persists plans.
 */
var BreakPlanner = {
  TZ: 'America/Toronto',
  RULES_PROP: 'BP_RULES_V1',
  PLAN_SHEET: 'WF_BREAK_PLANS',
  PLAN_HEADERS: ['DateStr', 'Agent Key', 'Agent Name', 'Plan JSON', 'Status', 'By', 'Timestamp'],
  PLAN_KEEP_DAYS: 45,

  // Mirrors WFEngine.DEFAULT_RULES — the client merges saved rules over its
  // own copy, so the server only stores what the admin changed.
  getRules: function () {
    try {
      var raw = PropertiesService.getScriptProperties().getProperty(this.RULES_PROP);
      return raw ? JSON.parse(raw) : null;
    } catch (e) { return null; }
  },

  setRules: function (json) {
    if (!this._isAdmin()) return { ok: false, error: 'Only MRC admins can change the planner rules.' };
    var r;
    try { r = (typeof json === 'string') ? JSON.parse(json) : json; } catch (e) { return { ok: false, error: 'Bad rules JSON' }; }
    if (!r || !Array.isArray(r.grid)) return { ok: false, error: 'Rules need a grid.' };
    var s = JSON.stringify(r);
    if (s.length > 8500) return { ok: false, error: 'Rules too large.' };
    PropertiesService.getScriptProperties().setProperty(this.RULES_PROP, s);
    return { ok: true };
  },

  resetRules: function () {
    if (!this._isAdmin()) return { ok: false, error: 'Only MRC admins can change the planner rules.' };
    PropertiesService.getScriptProperties().deleteProperty(this.RULES_PROP);
    return { ok: true };
  },

  // ── Day ─────────────────────────────────────────────────────────────────────
  getDay: function (dateStr) {
    var today = this._today();
    dateStr = this._normDate(dateStr) || today;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var src = this._rowsForDates(ss, [dateStr]);
    var rows = src.byDate[dateStr] || [];

    var ctx = this._context(ss);
    var overrides = this._overridesFor(ss, dateStr);
    var absences = this._absencesFor(ss, dateStr, dateStr === today);

    var agents = [], seen = {};
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var name = String(row[0] || '').trim();
      var key = _normalizeAgentKey(name);
      if (!key || seen[key]) continue;
      var sh = (typeof SafeTracker !== 'undefined') ? SafeTracker._shiftFromRow(row) : null;
      if (!sh) continue;                        // Off / no real shift
      seen[key] = true;

      var ov = overrides[key];
      var blocksJson = ov != null ? ov : row[7];
      var acts = this._blocks(blocksJson, sh.start);

      var absent = (absences[key] != null) ? absences[key] : String(row[9] || '');
      var partDay = /\(part day\)/i.test(absent);
      var fullAbsent = !!absent && !partDay;   // part-day absences are fixed blocks instead

      var ml = ctx.ml[key] || {};
      var lang = ctx.lang[key] || this._langFromText(ml.lang) || 'EN';
      var q = ctx.queues[key] || [];
      var qs = q.join(' ');
      var level = ml.level;
      var region = '';
      try { if (typeof RegionRegistry !== 'undefined') region = RegionRegistry.getRegion(name) || ''; } catch (e) {}
      if (!region) region = String(row[6] || '');

      agents.push({
        key: key, name: name, id: String(row[1] || ''), region: region, lang: lang,
        level: (level === 'Trainee') ? 'Trainee' : (parseInt(level, 10) || null),
        trainee: /trainee/i.test(String(level || '')),
        isQL: /question\s*line/i.test(ml.skills || '') || /question/i.test(qs),
        isESC: /escalat/i.test(qs),
        supervisor: ml.supervisor || '', site: ml.site || '', calgary: /calg/i.test(ml.site || ''),
        absent: absent, fullAbsent: fullAbsent, hasOverride: ov != null,
        s: sh.start, e: sh.end, acts: acts, queues: q
      });
    }
    agents.sort(function (a, b) { return (a.s - b.s) || a.name.localeCompare(b.name); });

    var nowMin = null;
    if (dateStr === today) {
      var hm = Utilities.formatDate(new Date(), this.TZ, 'HH:mm').split(':');
      nowMin = parseInt(hm[0], 10) * 60 + parseInt(hm[1], 10);
    }
    return {
      ok: true, date: dateStr, today: today, nowMin: nowMin, source: src.source[dateStr] || '',
      agents: agents, idp: this._idpFor(ss, dateStr), idpNext: this._idpFor(ss, this._addDays(dateStr, 1)), rules: this.getRules(),
      demand: this._demand(dateStr, dateStr === today),
      plan: this._planFor(ss, dateStr), user: this._me(), isAdmin: this._isAdmin(),
      range: src.range
    };
  },

  // ── Week (read-only shift grid) ────────────────────────────────────────────
  getWeek: function (startStr) {
    var start = this._normDate(startStr) || this._today();
    var p = start.split('-').map(Number), dates = [];
    for (var d = 0; d < 7; d++) dates.push(Utilities.formatDate(new Date(p[0], p[1] - 1, p[2] + d, 12), this.TZ, 'yyyy-MM-dd'));
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var src = this._rowsForDates(ss, dates);
    var ctx = this._context(ss);
    var byKey = {};
    var self = this;
    dates.forEach(function (ds) {
      (src.byDate[ds] || []).forEach(function (row) {
        var name = String(row[0] || '').trim(), key = _normalizeAgentKey(name);
        if (!key) return;
        var a = byKey[key] || (byKey[key] = { key: key, name: name, region: String(row[6] || ''), lang: ctx.lang[key] || self._langFromText((ctx.ml[key] || {}).lang) || 'EN', level: (ctx.ml[key] || {}).level || '', supervisor: (ctx.ml[key] || {}).supervisor || '', days: {} });
        if (a.days[ds]) return;
        var sh = (typeof SafeTracker !== 'undefined') ? SafeTracker._shiftFromRow(row) : null;
        var blocks = sh ? self._blocks(row[7], sh.start) : [];
        a.days[ds] = sh
          ? { s: sh.start, e: sh.end, absent: String(row[9] || ''), b: blocks.filter(function (x) { return x.type === 'Break'; }).length, l: blocks.filter(function (x) { return x.type === 'Lunch'; }).length }
          : { off: true, absent: String(row[9] || '') };
      });
    });
    var agents = Object.keys(byKey).map(function (k) { return byKey[k]; })
      .sort(function (a, b) { return a.name.localeCompare(b.name); });
    return { ok: true, dates: dates, agents: agents, range: src.range };
  },

  // ── Plans ──────────────────────────────────────────────────────────────────
  saveDraft: function (dateStr, planJson) { return this._savePlan(dateStr, planJson, 'draft'); },

  // Approve = save the plan AND push it to the floor through Break_Overrides
  // for every agent whose blocks differ from what the floor shows now (plus any
  // agent who already has an override, so re-approving can also undo).
  approve: function (dateStr, planJson) {
    var res = this._savePlan(dateStr, planJson, 'approved');
    if (!res.ok) return res;
    var plan = res.plan, self = this, rows = [];
    plan.agents.forEach(function (a) {
      if (!a.changed && !a.hasOverride) return;
      rows.push({ name: a.name, json: JSON.stringify(self._toBreaksJson(a)) });
    });
    var written = rows.length ? StatusTracker.updateBreaksBatch(res.date, rows) : 0;
    try { if (typeof AgentMonitor !== 'undefined') AgentMonitor._bust(); } catch (e) {}
    return { ok: true, date: res.date, saved: plan.agents.length, written: written, by: res.by, at: res.at };
  },

  _savePlan: function (dateStr, planJson, status) {
    var date = this._normDate(dateStr);
    if (!date) return { ok: false, error: 'Bad date' };
    var plan;
    try { plan = (typeof planJson === 'string') ? JSON.parse(planJson) : planJson; } catch (e) { return { ok: false, error: 'Bad plan JSON' }; }
    if (!plan || !Array.isArray(plan.agents)) return { ok: false, error: 'Empty plan' };

    var lock = LockService.getScriptLock();
    try { lock.waitLock(20000); } catch (e) { return { ok: false, error: 'Busy, try again' }; }
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName(this.PLAN_SHEET);
      if (!sh) {
        sh = ss.insertSheet(this.PLAN_SHEET);
        sh.getRange(1, 1, 1, this.PLAN_HEADERS.length).setValues([this.PLAN_HEADERS]).setFontWeight('bold');
        sh.setFrozenRows(1);
      }
      var cutoff = Utilities.formatDate(new Date(Date.now() - this.PLAN_KEEP_DAYS * 86400000), this.TZ, 'yyyy-MM-dd');
      var last = sh.getLastRow();
      var keep = [];
      if (last > 1) {
        var self = this;
        keep = sh.getRange(2, 1, last - 1, this.PLAN_HEADERS.length).getValues().filter(function (r) {
          var d = self._normDate(r[0]);
          return d && d !== date && d >= cutoff;
        }).map(function (r) { r[0] = self._normDate(r[0]); return r; });
      }
      var by = this._me(), at = new Date();
      var mine = plan.agents.map(function (a) {
        return [date, String(a.key || ''), String(a.name || ''), JSON.stringify({ acts: a.acts || [], fixed: a.fixed || [], extra: a.extra || [], changed: !!a.changed }), status, by, at];
      });
      var all = keep.concat(mine);
      if (last > 1) sh.getRange(2, 1, last - 1, this.PLAN_HEADERS.length).clearContent();
      if (all.length) {
        sh.getRange(2, 1, all.length, 1).setNumberFormat('@');   // keep DateStr as text
        sh.getRange(2, 1, all.length, this.PLAN_HEADERS.length).setValues(all);
      }
      return { ok: true, date: date, plan: plan, by: by, at: Utilities.formatDate(at, this.TZ, 'yyyy-MM-dd HH:mm') };
    } finally { lock.releaseLock(); }
  },

  _planFor: function (ss, date) {
    var sh = ss.getSheetByName(this.PLAN_SHEET);
    if (!sh || sh.getLastRow() < 2) return null;
    var out = null, self = this;
    sh.getRange(2, 1, sh.getLastRow() - 1, this.PLAN_HEADERS.length).getValues().forEach(function (r) {
      if (self._normDate(r[0]) !== date) return;
      if (!out) out = { status: String(r[4] || ''), by: String(r[5] || ''), at: r[6] instanceof Date ? Utilities.formatDate(r[6], self.TZ, 'yyyy-MM-dd HH:mm') : String(r[6] || ''), agents: {} };
      try { out.agents[String(r[1])] = JSON.parse(r[3]); } catch (e) {}
    });
    return out;
  },

  // Floor format: the same [{type,start,end}] the IEX import writes, clock
  // strings like "10:30 AM". Fixed blocks go back exactly as they came in.
  _toBreaksJson: function (a) {
    var self = this, out = [];
    (a.fixed || []).forEach(function (f) {
      var o = { type: f.type, start: f.rs || self._clock(f.s), end: f.re || self._clock(f.e) };
      if (f.abs) o.abs = f.abs;
      out.push(o);
    });
    (a.extra || []).forEach(function (b) { out.push({ type: b.type, start: self._clock(b.s), end: self._clock(b.e) }); });
    (a.acts || []).forEach(function (b) { out.push({ type: b.type === 'Lunch' ? 'Lunch' : 'Break', start: self._clock(b.s), end: self._clock(b.e) }); });
    out.sort(function (x, y) { return self._mins(x.start) - self._mins(y.start); });
    return out;
  },

  // ── Readers ────────────────────────────────────────────────────────────────

  // Schedule rows per date. The current paste (Raw Schedule) wins; dates it
  // doesn't cover come from Schedule_History. History can be large, so only its
  // DateStr column is scanned first and then just the span of matching rows read.
  _rowsForDates: function (ss, dates) {
    var want = {}, byDate = {}, source = {}, self = this;
    dates.forEach(function (d) { want[d] = true; byDate[d] = []; });
    var range = { min: '', max: '' };
    var note = function (d) { if (!range.min || d < range.min) range.min = d; if (!range.max || d > range.max) range.max = d; };

    var raw = ss.getSheetByName('Raw Schedule');
    if (raw && raw.getLastRow() > 1) {
      raw.getRange(2, 1, raw.getLastRow() - 1, 12).getValues().forEach(function (r) {
        var d = self._normDate(r[2]); if (!d) return; note(d);
        if (want[d] && r[0]) { byDate[d].push(r); source[d] = 'Raw Schedule'; }
      });
    }
    var missing = dates.filter(function (d) { return !byDate[d].length; });
    var hist = ss.getSheetByName('Schedule_History');
    if (hist && hist.getLastRow() > 1) {
      var n = hist.getLastRow() - 1;
      var col = hist.getRange(2, 3, n, 1).getValues();
      var lo = -1, hi = -1, miss = {};
      missing.forEach(function (d) { miss[d] = true; });
      for (var i = 0; i < n; i++) {
        var d = self._normDate(col[i][0]); if (!d) continue; note(d);
        if (miss[d]) { if (lo < 0) lo = i; hi = i; }
      }
      if (lo >= 0) {
        hist.getRange(2 + lo, 1, hi - lo + 1, 12).getValues().forEach(function (r) {
          var d = self._normDate(r[2]);
          if (d && miss[d] && r[0]) { byDate[d].push(r); source[d] = 'Schedule_History'; }
        });
      }
    }
    return { byDate: byDate, source: source, range: range };
  },

  _context: function (ss) {
    var ml = {}, lang = {}, queues = {};
    var mlSh = ss.getSheetByName('WF_MASTERLIST');
    if (mlSh && mlSh.getLastRow() > 1) {
      mlSh.getRange(1, 1, mlSh.getLastRow(), Math.max(7, mlSh.getLastColumn())).getDisplayValues().slice(1).forEach(function (r) {
        var k = _normalizeAgentKey(r[0]); if (!k) return;
        ml[k] = { level: String(r[1] || '').trim(), supervisor: String(r[2] || '').trim(), skills: String(r[3] || ''), site: String(r[4] || '').trim(), lang: String(r[6] || '') };
      });
    }
    var lm = ss.getSheetByName('WF_LANG_MAP');
    if (lm && lm.getLastRow() > 1) {
      lm.getRange(2, 1, lm.getLastRow() - 1, 3).getDisplayValues().forEach(function (r) {
        var k = String(r[0] || '').trim(), v = String(r[2] || '').trim().toUpperCase();
        if (k && (v === 'EN' || v === 'FR' || v === 'BL')) lang[k] = v;
      });
    }
    var ga = ss.getSheetByName('Genesys_Assignments');
    if (ga && ga.getLastRow() > 1) {
      ga.getRange(2, 1, ga.getLastRow() - 1, 4).getValues().forEach(function (r) {
        var k = _normalizeAgentKey(r[1]); if (!k) return;
        var q = [];
        if (r[2]) { try { q = JSON.parse(r[3] || '[]'); } catch (e) { q = []; } }
        queues[k] = Array.isArray(q) ? q.map(String) : [];
      });
    }
    return { ml: ml, lang: lang, queues: queues };
  },

  _overridesFor: function (ss, date) {
    var out = {}, self = this;
    var sh = ss.getSheetByName('Break_Overrides');
    if (!sh || sh.getLastRow() < 2) return out;
    sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function (r) {
      if (self._normDate(r[2]) !== date) return;
      var k = _normalizeAgentKey(r[1]); if (k) out[k] = r[3];
    });
    return out;
  },

  // Absence marked on the floor for the date ('' = explicitly cleared).
  _absencesFor: function (ss, date, isToday) {
    var out = {}, self = this;
    var sh = ss.getSheetByName('Status_Overrides');
    if (sh && sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(function (r) {
        if (self._normDate(r[4]) !== date) return;
        // Same precedence as the floor (AgentMonitor): a blank sheet value is a
        // role-only row and falls through to the schedule; the fast map's
        // explicit '' is a real "cleared".
        var k = _normalizeAgentKey(r[1]); if (k && r[3]) out[k] = String(r[3]);
      });
    }
    if (isToday && typeof RoleManager !== 'undefined') {
      var fm = RoleManager.getFastMap();
      for (var k in fm) if (fm[k] && fm[k].absent !== undefined) out[k] = String(fm[k].absent || '');
    }
    return out;
  },

  _idpFor: function (ss, date) {
    var req = [], open = [], self = this;
    for (var i = 0; i < 96; i++) { req.push(null); open.push(null); }
    var sh = ss.getSheetByName('WF_IDP');
    if (!sh || sh.getLastRow() < 2) return { req: req, open: open, has: false };
    var n = sh.getLastRow() - 1;
    var vals = sh.getRange(2, 1, n, 4).getValues(), disp = sh.getRange(2, 1, n, 4).getDisplayValues(), has = false;
    for (var r = 0; r < n; r++) {
      if (self._normDate(vals[r][0]) !== date) continue;
      var m = (typeof SafeTracker !== 'undefined') ? SafeTracker._safeTime(disp[r][1]) : self._mins(disp[r][1]);
      if (m < 0) continue;
      var b = Math.floor(m / 15);
      var rq = parseFloat(String(disp[r][2]).replace(',', '.')), op = parseFloat(String(disp[r][3]).replace(',', '.'));
      if (!isNaN(rq)) { req[b] = rq; has = true; }
      if (!isNaN(op)) open[b] = op;
    }
    return { req: req, open: open, has: has };
  },

  // Breaks JSON → [{type, s, e, fixed, rs, re, abs}] in day-minutes aligned to
  // the shift (a block before the shift start belongs to the next day).
  _blocks: function (json, shiftStart) {
    var arr = [];
    try { arr = (typeof json === 'string') ? JSON.parse(json || '[]') : (json || []); } catch (e) { arr = []; }
    if (!Array.isArray(arr)) return [];
    var self = this, out = [];
    arr.forEach(function (b) {
      if (!b || !b.start || !b.end) return;
      var s = (typeof SafeTracker !== 'undefined') ? SafeTracker._safeTime(b.start) : self._mins(b.start);
      var e = (typeof SafeTracker !== 'undefined') ? SafeTracker._safeTime(b.end) : self._mins(b.end);
      if (s < 0 || e < 0) return;
      if (s < shiftStart - 120) { s += 1440; e += 1440; }
      if (e <= s) e += 1440;
      var t = String(b.type || '');
      var type = /lunch|repas/i.test(t) ? 'Lunch' : (/break|pause/i.test(t) && !/^OT /i.test(t)) ? 'Break' : t;
      var o = { type: type, s: s, e: e, fixed: !(type === 'Break' || type === 'Lunch') };
      if (o.fixed) { o.rs = b.start; o.re = b.end; if (b.abs) o.abs = b.abs; }
      out.push(o);
    });
    return out.sort(function (a, b) { return a.s - b.s; });
  },

  // Uplifts on top of the IDP (forecast bias, live stats, learned conditions) —
  // CORE/DemandSignals.gs. Never lets a feed problem break the planner.
  _demand: function (dateStr, isToday) {
    try { return (typeof DemandSignals !== 'undefined') ? DemandSignals.getFactors(dateStr, isToday) : null; }
    catch (e) { return { factors: [], error: String(e && e.message || e) }; }
  },
  _addDays: function (ds, n) {
    var p = ds.split('-').map(Number);
    return Utilities.formatDate(new Date(p[0], p[1] - 1, p[2] + n, 12), this.TZ, 'yyyy-MM-dd');
  },

  // ── Small helpers ──────────────────────────────────────────────────────────
  _langFromText: function (v) {
    var s = String(v || '').toLowerCase();
    if (!s) return '';
    if (s.indexOf('biling') !== -1 || s === 'bl') return 'BL';
    if (s.indexOf('french') !== -1 || s.indexOf('fran') !== -1 || s === 'fr') return 'FR';
    if (s.indexOf('english') !== -1 || s === 'en') return 'EN';
    return '';
  },
  _normDate: function (v) {
    if (v == null || v === '') return '';
    if (v instanceof Date) return Utilities.formatDate(v, this.TZ, 'yyyy-MM-dd');
    var s = String(v).trim();
    var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
    return '';
  },
  _today: function () { return Utilities.formatDate(new Date(), this.TZ, 'yyyy-MM-dd'); },
  _mins: function (s) {
    var m = String(s || '').match(/(\d{1,2}):(\d{2})\s*([AP]M)?/i); if (!m) return -1;
    var h = parseInt(m[1], 10), mi = parseInt(m[2], 10), ap = (m[3] || '').toUpperCase();
    if (ap === 'PM' && h < 12) h += 12; if (ap === 'AM' && h === 12) h = 0;
    return h * 60 + mi;
  },
  _clock: function (m) {
    var x = ((Math.round(Number(m)) % 1440) + 1440) % 1440, h = Math.floor(x / 60), mi = x % 60;
    var ap = h >= 12 ? 'PM' : 'AM', h12 = h % 12 === 0 ? 12 : h % 12;
    return h12 + ':' + (mi < 10 ? '0' : '') + mi + ' ' + ap;
  },
  _me: function () {
    try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
  },
  _isAdmin: function () {
    try { return (typeof FeedbackTracker !== 'undefined') ? !!FeedbackTracker.isAdmin() : false; } catch (e) { return false; }
  }
};
