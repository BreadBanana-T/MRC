/**
 * MODULE: ASSIGNMENT ANALYZER (RED FLAGS - LOCAL HOST ONLY)
 */

var AssignmentAnalyzer = {

  // WF_MASTERLIST layout. Columns 1-6 are the original SharePoint shape and
  // must not move — AgentMonitor, SafeTracker, OutageTracker, TrainingTracker
  // and WorkforceTracker all read them positionally. Columns 7-11 are the
  // Genesys additions, appended so those readers keep working untouched.
  ML_HEADERS: ["Agent Name", "ERC Level", "Supervisor", "Skills", "Location", "Email_TI",
               "Agent Id", "Language", "Level Name", "Skill List", "Source"],

  _getDB: function(sheetName) {
      return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  },

  /**
   * Roster sync. Accepts EITHER source and auto-detects which one was pasted:
   *
   *   GENESYS CLOUD (preferred)  "Agent Id / Agent Name / Skills / …"
   *       Level, skills and language all come out of the Skills cell — see
   *       GenesysRoster. Juniors and profile-less rows are still stored (the
   *       floor wants their names) but are marked non-countable so Red Flags
   *       drops them.
   *   SHAREPOINT MASTERLIST (legacy)  "employee_name / job_title / erc_level / …"
   *       Unchanged behaviour: managers filtered by job title, region inferred
   *       from Location / TI email.
   */
  importMasterList: function(rawText) {
      try {
          if (!rawText || rawText.trim() === "") return "Error: No data pasted.";
          let lines = rawText.split(/\r?\n/).filter(l => l.trim().length > 0);

          let rows = lines.map(l => this._parseCSVLine(l));

          // ── Genesys Cloud export, with its header row? ── Scan every line,
          // not just the top: a paste often carries a title/filter preamble
          // above the real header. Data rows can't false-positive — the check
          // needs literal "Agent Id" / "Agent Name" / "Skills" cells.
          let gHeaderIdx = -1;
          for (let i = 0; i < rows.length; i++) {
              if (GenesysRoster.looksLikeGenesysHeader(rows[i])) { gHeaderIdx = i; break; }
          }
          if (gHeaderIdx !== -1) return this._importGenesysRoster(rows.slice(gHeaderIdx), null);

          let headerIdx = lines.findIndex(l => l.toLowerCase().includes('employee_name') || l.toLowerCase().includes('job_title') || l.toLowerCase().includes('name_first'));

          // ── Genesys rows WITHOUT the header ── Selecting rows in the Genesys
          // grid and copying gives data only. Infer the layout from the
          // SAM_Monit tags rather than rejecting a perfectly good paste.
          if (headerIdx === -1) {
              let layout = GenesysRoster.sniffLayout(rows);
              if (layout) return this._importGenesysRoster(rows, layout);
              return "Error: Not a recognized roster. Paste the Genesys agent export (Agent Id / Agent Name / Skills) or the SharePoint MasterList.";
          }

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
                  let cleanName = name.replace(/(^"|"$)/g, '').trim();
                  // Cols 7-11 stay blank for the SharePoint source: no Genesys
                  // id/level-tag, and the language is inferred at read time.
                  validAgents.push([cleanName, level, sup, skills, loc, emailTI, "", "", "", "", "sharepoint"]);
              }
          }

          if (validAgents.length > 0) {
              const ss = SpreadsheetApp.getActiveSpreadsheet();
              let sheet = ss.getSheetByName('WF_MASTERLIST');
              if (!sheet) sheet = ss.insertSheet('WF_MASTERLIST');
              sheet.clearContents();
              sheet.appendRow(this.ML_HEADERS);
              sheet.getRange(2, 1, validAgents.length, this.ML_HEADERS.length).setValues(validAgents);

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
              return `Agent roster synced from the SharePoint MasterList — ${validAgents.length} active agent${validAgents.length === 1 ? '' : 's'} locked in.\nManagers are permanently ignored. Paste the Genesys export instead to get levels, skills and languages.`;
          }
          return "Error: No active agents found in the pasted list.";
      } catch (e) {
          return "Error parsing MasterList: " + e.message;
      }
  },

  /**
   * Genesys Cloud roster writer.
   * @param {string[][]} rows    parsed cells; starts at the header row unless
   *                             `layout` is given, in which case it's all data
   * @param {?Object} layout     column indices from GenesysRoster.sniffLayout
   *
   * Region is NOT touched: the export carries no location or TI email, so
   * upserting everyone as "Onshore" would poison RegionRegistry at masterlist
   * rank and silently move offshore agents. Region keeps coming from the
   * registry / WFM schedule.
   */
  _importGenesysRoster: function(rows, layout) {
      var built = GenesysRoster.buildRoster(rows, layout);
      if (!built.agents.length) return "Error: No agents found in the Genesys export.";

      var out = built.agents.map(function(a) {
          var p = a.profile;
          return [
              a.name,
              p.level === null ? "" : p.level,   // blank level = no Lv tag = not a monitoring specialist
              "",                                // Supervisor: not in the Genesys export
              p.raw,                              // raw Skills cell — SafeTracker/OutageTracker read this
              "", "",                             // Location / Email_TI: not in the Genesys export
              a.id,
              p.lang,
              p.levelName,
              p.skillText,
              "genesys"
          ];
      });

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('WF_MASTERLIST');
      if (!sheet) sheet = ss.insertSheet('WF_MASTERLIST');
      sheet.clearContents();
      sheet.appendRow(this.ML_HEADERS);
      sheet.getRange(2, 1, out.length, this.ML_HEADERS.length).setValues(out);

      var c = built.counts;
      return "Agent roster synced from Genesys — " + c.countable + " monitoring specialist" + (c.countable === 1 ? "" : "s") + " counted"
           + " (" + c.senior + " Senior / " + c.regular + " Regular · "
           + c.bilingual + " BL, " + c.english + " EN, " + c.french + " FR).\n"
           + "Excluded from Red Flags: " + c.junior + " Junior, " + c.noProfile + " with no skills attached"
           + (c.noLevel ? ", " + c.noLevel + " with a language but no level" : "") + "."
           + (layout ? "\n(no header row in the paste — columns inferred from the SAM_Monit tags)" : "");
  },

  /**
   * Read WF_MASTERLIST into a normalized-key map. Single reader shared by the
   * analyzer so the Genesys/SharePoint difference is resolved in one place.
   *
   * @return {{byKey:Object, isGenesys:boolean, counts:Object}}
   *   byKey[key] = { name, level, levelName, lang, langLabel, skills:[],
   *                  skillText, countable, reason, agentId, source }
   */
  readRoster: function() {
      var out = { byKey: {}, isGenesys: false, counts: { total: 0, countable: 0, junior: 0, noProfile: 0, noLevel: 0 } };
      var sheet = this._getDB('WF_MASTERLIST');
      if (!sheet || sheet.getLastRow() < 2) return out;

      var rows = sheet.getDataRange().getDisplayValues().slice(1);
      rows.forEach(function(r) {
          var name = String(r[0] || '').trim();
          if (!name) return;
          var key = _normalizeAgentKey(name);
          if (!key) return;

          var source = String(r[10] || '').toLowerCase();
          var rawSkills = String(r[3] || '');
          // Rows written before this column existed have no Source: fall back
          // to sniffing the Skills cell for SAM_Monit tags.
          var genesys = (source === 'genesys') || (!source && GenesysRoster.looksLikeGenesys(rawSkills));

          var entry;
          if (genesys) {
              out.isGenesys = true;
              var p = GenesysRoster.parseSkills(rawSkills);
              // Prefer the stored decoded columns when present (cheap + lets a
              // human correct a row by hand in the sheet).
              var storedLvl = parseInt(r[1], 10);
              if (!isNaN(storedLvl)) p.level = storedLvl;
              if (String(r[7] || '').trim()) p.lang = String(r[7]).trim().toUpperCase();
              if (String(r[8] || '').trim()) p.levelName = String(r[8]).trim();
              if (String(r[9] || '').trim()) { p.skillText = String(r[9]).trim(); p.skills = p.skillText.split(/\s*,\s*/).filter(Boolean); }
              p.langLabel = GenesysRoster.LANG_LABELS[p.lang] || '';
              // Recompute countability off the (possibly corrected) values.
              if (p.level === null || isNaN(p.level)) { p.countable = false; p.reason = p.lang ? 'no-level' : 'no-profile'; }
              else if (p.level < 2) { p.countable = false; p.reason = 'junior'; }
              else { p.countable = true; p.reason = ''; }

              entry = {
                  name: name, level: p.level, levelName: p.levelName, lang: p.lang, langLabel: p.langLabel,
                  skills: p.skills, skillText: p.skillText, countable: p.countable, reason: p.reason,
                  agentId: String(r[6] || '').trim(), source: 'genesys'
              };
          } else {
              var lvl = parseInt(r[1], 10);
              if (isNaN(lvl)) lvl = 2;
              // Legacy rows carry no level tags — everyone in the SharePoint
              // MasterList is a monitoring specialist by construction (managers
              // were dropped at import), so they all stay countable.
              entry = {
                  name: name, level: lvl, levelName: '', lang: '', langLabel: '',
                  skills: [], skillText: '', countable: true, reason: '',
                  agentId: '', source: 'sharepoint'
              };
          }

          out.byKey[key] = entry;
          out.counts.total++;
          if (entry.countable) out.counts.countable++;
          else if (entry.reason === 'junior') out.counts.junior++;
          else if (entry.reason === 'no-level') out.counts.noLevel++;
          else out.counts.noProfile++;
      });
      return out;
  },

  importGEMData: function(rawText) {
      try {
          let parsed = this._parseGEM(rawText);
          let map = parsed.agents;
          if (Object.keys(map).length === 0) return JSON.stringify({error: "Invalid data format or missing GEM headers."});

          let repMonth = "";
          for (let k in map) { if (map[k].month) { repMonth = map[k].month; break; } }
          if (!repMonth) repMonth = Utilities.formatDate(new Date(), "America/Toronto", "yyyy-MM-01");

          // "Inbound Time %" / "Outbound Time %" are the ONLY source for three of
          // the four flag rules. They are a *time-staffed-on-skill* measure, not
          // anything derivable from the call columns — in the Assignment Tracker
          // an agent shows 100% inbound while placing 931 outbound calls. So when
          // an export arrives without them we must not silently write zeros:
          // that reads as "nobody is unbalanced" instead of "we can't tell".
          // Keep whatever percentages the month already had.
          this._saveGEMToDB(map, repMonth, !parsed.hasPct);

          let res = JSON.parse(this.getAnalyzerData(repMonth));
          res.importWarning = parsed.hasPct ? null : {
              kind: 'missing-pct',
              month: repMonth,
              missing: parsed.missingCols,
              text: 'This export has no "Inbound Time %" / "Outbound Time %" columns, so Inbound, Out of Scope and Outbound flags cannot be computed for ' + repMonth +
                    '. CP/H flags are unaffected. Any percentages already stored for this month were kept, not overwritten. Re-export the GEM report with both columns to restore the other three.'
          };
          return JSON.stringify(res);
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
          
          // Roster gate. With a Genesys roster the roster is authoritative and
          // acts as a hard filter (see _classifyAgent); with the legacy
          // SharePoint list it stays advisory — a stale list must not silently
          // erase real agents from the flags.
          const roster = this.readRoster();
          const useMasterList = roster.counts.total > 0;
          const genesysMode = roster.isGenesys;

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

          // Why each agent was dropped, so the dashboard can prove the filter
          // ran instead of just showing a shorter list. `notOnRoster` is NOT a
          // drop list — those agents are counted at L2; it's a data-health
          // signal for name mismatches between GEM and Genesys.
          let filtered = { junior: [], noProfile: [], noLevel: [], notOnRoster: [] };

          let currentMap = {};
          let prevMap = {};
          validData.forEach(r => {
              let agent = String(r[1]).trim();
              let verdict = this._classifyAgent(agent, roster, genesysMode);

              if (r[0] === targetMonth) {
                  if (verdict.dropReason) {
                      let bucket = { 'junior': 'junior', 'no-profile': 'noProfile', 'no-level': 'noLevel' }[verdict.dropReason];
                      if (bucket && filtered[bucket].indexOf(agent) === -1) filtered[bucket].push(agent);
                  } else if (verdict.unmatched && filtered.notOnRoster.indexOf(agent) === -1) {
                      filtered.notOnRoster.push(agent);
                  }
              }
              if (verdict.dropReason) return;

              if (exclusions.has(agent)) return;
              let inP = parseFloat(r[3]) || 0; if (inP > 1) inP = inP / 100;
              let outP = parseFloat(r[4]) || 0; if (outP > 1) outP = outP / 100;
              let cphV = parseFloat(r[2]) || 0;

              if (r[0] === targetMonth) {
                  currentMap[agent] = {
                      level: verdict.level, levelName: verdict.levelName,
                      lang: verdict.lang, langLabel: verdict.langLabel, skills: verdict.skills,
                      cph: cphV, inPct: inP, outPct: outP, aht: fixTime(r[5]),
                      tasksXCalls: parseFloat(r[6]) || 0, totalCalls: parseInt(r[7]) || 0,
                      callsAnsw: parseInt(r[8]) || 0, wDays: parseInt(r[9]) || 0
                  };
              } else if (r[0] === prevMonth) {
                  prevMap[agent] = { cph: cphV, inPct: inP, outPct: outP };
              }
          });

          // Does this month have any In/Out % at all? If every agent reads 0,
          // the import lacked the two columns — the Inbound / Out of Scope /
          // Outbound cards are then "not measurable", not "nobody is flagged".
          let pctPresent = Object.keys(currentMap).some(a => currentMap[a].inPct > 0 || currentMap[a].outPct > 0);

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
                  aht: c.aht, tasksXCalls: c.tasksXCalls, totalCalls: c.totalCalls,
                  level: c.level, levelName: c.levelName, lang: c.lang, langLabel: c.langLabel, skills: c.skills
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
              rosterSource: genesysMode ? 'genesys' : (useMasterList ? 'sharepoint' : 'none'),
              rosterCounts: roster.counts,
              pctPresent: pctPresent, agentsInMonth: Object.keys(currentMap).length,
              filtered: {
                  junior: filtered.junior.sort(), noProfile: filtered.noProfile.sort(),
                  noLevel: filtered.noLevel.sort(), notOnRoster: filtered.notOnRoster.sort()
              },
              inbound: inboundList, outOfScope: outOfScopeList, outbound: outboundList, cph: cphList
          });
      } catch (e) {
          return JSON.stringify({ error: "Data generation failed: " + e.message });
      }
  },

  /**
   * Decide whether a GEM row belongs in Red Flags, and with what profile.
   *
   * The roster only removes people it EXPLICITLY classifies:
   *   - juniors (not counted, per product owner)
   *   - agents with no SAM_Monit tags at all ("nothing attached")
   *   - agents with a language but no level tag
   * An agent the roster has never heard of is KEPT at level 2 — in every month.
   * They're someone who has since left, or a GEM/Genesys name mismatch, and
   * neither is a reason to silently erase a real agent from the flags or to
   * rewrite an old month. They're reported separately as `unmatched` so a
   * spelling mismatch is visible instead of invisible.
   *
   * LEGACY (SharePoint) MODE — no level tags exist, so nothing is dropped.
   *
   * @return {{dropReason:string, unmatched:boolean, level:number,
   *           levelName:string, lang:string, langLabel:string,
   *           skills:string[]}}  dropReason '' = keep
   */
  _classifyAgent: function(agentName, roster, genesysMode) {
      var keep = function(e, unmatched) {
          return { dropReason: '', unmatched: !!unmatched,
                   level: e ? e.level : 2, levelName: e ? e.levelName : '',
                   lang: e ? e.lang : '', langLabel: e ? e.langLabel : '', skills: (e && e.skills) ? e.skills : [] };
      };
      if (!roster || roster.counts.total === 0) return keep(null);

      // Memoized: the GEM table repeats every agent once per month, and the
      // fuzzy fallback below is an O(roster) scan.
      if (!roster._verdicts) roster._verdicts = {};
      var ck = _normalizeAgentKey(agentName);
      if (roster._verdicts[ck]) return roster._verdicts[ck];

      var entry = this._matchRoster(agentName, roster.byKey);
      var verdict;
      if (!entry) {
          verdict = keep(null, genesysMode);   // unmatched only means anything in Genesys mode
      } else if (genesysMode && !entry.countable) {
          verdict = { dropReason: entry.reason || 'no-profile', unmatched: false,
                      level: entry.level, levelName: entry.levelName,
                      lang: entry.lang, langLabel: entry.langLabel, skills: entry.skills };
      } else {
          verdict = keep(entry, false);
      }
      if (ck) roster._verdicts[ck] = verdict;
      return verdict;
  },

  /**
   * Roster lookup: exact normalized key first, then the historical fallbacks
   * for the same person spelled differently across GEM and the roster
   * (first+last token pair, or one name contained in the other).
   */
  _matchRoster: function(agentName, byKey) {
      var key = _normalizeAgentKey(agentName);
      if (!key) return null;
      if (byKey[key]) return byKey[key];

      var parts = key.split(' ').filter(function(x) { return x.length > 1; });
      var keys = Object.keys(byKey);
      for (var i = 0; i < keys.length; i++) {
          var mlKey = keys[i];
          var mlParts = mlKey.split(' ').filter(function(x) { return x.length > 1; });
          if (parts.length > 1 && mlParts.length > 1 && parts[0] === mlParts[0] && parts[1] === mlParts[1]) return byKey[mlKey];
          var minLen = Math.min(key.length, mlKey.length);
          if (minLen >= 8 && (key.indexOf(mlKey) !== -1 || mlKey.indexOf(key) !== -1)) return byKey[mlKey];
      }
      return null;
  },

  /**
   * @return {{agents:Object, hasPct:boolean, missingCols:string[]}}
   *   hasPct is false when the export omits Inbound/Outbound Time % — three of
   *   the four flag rules are then uncomputable and the caller must say so
   *   rather than treat 0% as a measurement.
   */
  _parseGEM: function(raw) {
      let empty = { agents: {}, hasPct: false, missingCols: [] };
      if (!raw || raw.trim() === "") return empty;
      let lines = raw.split(/\r?\n/).filter(l => l.trim().length > 0);
      let headerIdx = lines.findIndex(l => l.toLowerCase().includes('in & out by hr') || l.toLowerCase().includes('cph'));
      if (headerIdx === -1) return empty;

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

      // Three of the four flag rules read nothing but these two columns.
      let missingCols = [];
      if (colMap.inPct === undefined) missingCols.push('Inbound Time %');
      if (colMap.outPct === undefined) missingCols.push('Outbound Time %');
      let hasPct = missingCols.length === 0;

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
      return { agents: map, hasPct: hasPct, missingCols: missingCols };
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

  /**
   * @param {boolean} keepExistingPct  the export had no Inbound/Outbound Time %
   *   columns, so every parsed percentage is a structural 0, not a measurement.
   *   Carry forward whatever the month already held instead of destroying it.
   */
  _saveGEMToDB: function(gemMap, repMonth, keepExistingPct) {
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
          
          let monthOf = function(row) {
              return row[0] instanceof Date ? Utilities.formatDate(row[0], "America/Toronto", "yyyy-MM-dd") : String(row[0]);
          };
          let retained = existing.filter(row => monthOf(row) !== repMonth);

          // Percentages already on file for this month, keyed the same way the
          // analyzer keys agents, so a re-import with a column-short export
          // doesn't wipe them.
          let priorPct = {};
          if (keepExistingPct) {
              existing.forEach(row => {
                  if (monthOf(row) !== repMonth) return;
                  let k = _normalizeAgentKey(row[1]);
                  if (!k) return;
                  let pin = parseFloat(row[3]) || 0, pout = parseFloat(row[4]) || 0;
                  if (pin > 0 || pout > 0) priorPct[k] = { inPct: pin, outPct: pout };
              });
          }

          let safeT = function(t) { return (t && t !== '-' && t !== '00:00') ? "'" + t : "-"; };

          let newRows = [];
          Object.keys(gemMap).forEach(agent => {
              let m = gemMap[agent];
              let inP = m.inPct, outP = m.outPct;
              if (keepExistingPct) {
                  let p = priorPct[_normalizeAgentKey(agent)];
                  if (p) { inP = p.inPct; outP = p.outPct; }
              }
              newRows.push([
                  repMonth, agent, m.cph, inP, outP, safeT(m.aht), m.tasksXCalls, m.totalCalls, m.callsAnsw, m.wDays,
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
// Legacy name kept for any cached client bundle still calling it; both doors
// (Import Data box and Red Flags → Sync Roster) go through importAgentRoster
// in Code.gs so the analytics cache is always busted after a roster change.
function processMasterListUpload(rawText) { return importAgentRoster(rawText); }
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
