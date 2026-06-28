/**
 * MAIN ROUTER
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.appUrl = ScriptApp.getService().getUrl();
  template.popoutParam = (e && e.parameter && e.parameter.popout) ? "true" : "false";
  template.modeParam = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : "";
  
  return template.evaluate()
      .setTitle('MRC Operations Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

// --- DATA FETCHERS ---
function getFloorStatus() { return (typeof AgentMonitor !== 'undefined') ? AgentMonitor.getPayload() : "{}"; }
function getStatsHistory() { return (typeof StatsTracker !== 'undefined') ? StatsTracker.getHistory() : "[]"; }
function getIdpHistory() { return (typeof StatsTracker !== 'undefined') ? StatsTracker.getIdpHistory() : "[]"; }
function getLiveDashboardData() {
  try { return (typeof WeatherService !== 'undefined') ? JSON.stringify(WeatherService.fetch()) : "{}"; } 
  catch (e) { return "{}"; }
}
function getSystemNotifications() { return (typeof NotificationHandler !== 'undefined') ? NotificationHandler.getPending() : "[]"; }

// --- DELEGATED HANDLERS ---
function getDailyCalendarEvents() { return (typeof CalendarHandler !== 'undefined') ? CalendarHandler.getEvents() : "[]"; }
function calculateIdpFromText(text) { return (typeof IdpCalculator !== 'undefined') ? calculateAndLogIdp(text) : JSON.stringify({success:false}); }

// --- LOG RETRIEVAL ---
function getDailyRoleLogs(dateStr) { return fetchLogs(rowDate => rowDate === dateStr); }
function getWeeklyRoleLogs(dateStr) {
  const target = new Date(dateStr + "T12:00:00");
  const day = target.getDay();
  const diff = target.getDate() - day; 
  const start = new Date(target); start.setDate(diff); 
  const end = new Date(start); end.setDate(start.getDate() + 6);
  const startStr = Utilities.formatDate(start, "America/Toronto", "yyyy-MM-dd");
  const endStr = Utilities.formatDate(end, "America/Toronto", "yyyy-MM-dd");
  return fetchLogs(rowDate => rowDate >= startStr && rowDate <= endStr);
}
function getDailyStatsLog(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Stats History"); 
  if (!sheet) return "[]";
  const data = sheet.getDataRange().getValues();
  const stats = data.slice(1).filter(row => {
    const rowDate = Utilities.formatDate(new Date(row[0]), "America/Toronto", "yyyy-MM-dd");
    return rowDate === dateStr;
  }).map(row => ({
    time: Utilities.formatDate(new Date(row[0]), "America/Toronto", "HH:mm"),
    svl: row[1],
    ack: row[2]
  }));
  return JSON.stringify(stats);
}
function fetchLogs(filterFn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("DB_Sessions");
  if (!sheet) return "[]";
  const data = sheet.getDataRange().getValues();
  const logs = data.slice(1).filter(row => {
    if (!row[3]) return false; 
    const rowDate = Utilities.formatDate(new Date(row[3]), "America/Toronto", "yyyy-MM-dd");
    return filterFn(rowDate);
  }).map(row => ({
    date: Utilities.formatDate(new Date(row[3]), "America/Toronto", "yyyy-MM-dd"),
    agent: row[1],
    role: row[2],
    start: Utilities.formatDate(new Date(row[3]), "America/Toronto", "HH:mm"),
    end: Utilities.formatDate(new Date(row[4]), "America/Toronto", "HH:mm"),
    duration: row[5]
  }));
  return JSON.stringify(logs.reverse());
}

// --- ACTIONS ---
function updateAgentStatus(n, t, v) { if(typeof AgentMonitor!=='undefined') return AgentMonitor.setStatus(n, t, v); }
function updateAgentBreaks(n, j) { if(typeof AgentMonitor!=='undefined') return AgentMonitor.updateAgentBreaks(n, j); }
function submitOvertime(n, s, e, bs, be) { if(typeof AgentMonitor!=='undefined') return AgentMonitor.logOvertime(n, s, e, bs, be); }
function runCalculator(i, o) { if(typeof calculateMetrics!=='undefined') return calculateMetrics(i, o); return "{}"; }
function fetchScripts() { return (typeof ScriptHandler!=='undefined') ? JSON.stringify(ScriptHandler.getAll()) : "[]"; }
// saveTeamScript / deleteTeamScript / getTeamScripts live in CORE/ScriptHandler.gs.
// Do NOT redeclare them here: duplicate top-level names across .gs files
// shadow each other based on file load order.
function fillWindsToSheet() { if(typeof WeatherService!=='undefined') return LogSync.fillWinds(WeatherService.fetch()); }
function runImport(t) { if(typeof ImportHandler!=='undefined') return ImportHandler.run(t); }
function submitIdpValue(v) { if(typeof StatsTracker!=='undefined') return StatsTracker.logIdp(v); }

// --- SAFE TRACKER (forensic per-agent SAFE hours) ---
function getSafeAnalytics(mode, refDate, region, cycle) {
  return (typeof SafeTracker !== 'undefined')
    ? SafeTracker.getAnalytics(mode, refDate, region, cycle)
    : JSON.stringify({ error: 'SafeTracker not loaded.' });
}
function getSafeScheduleRoster() {
  return (typeof SafeTracker !== 'undefined') ? SafeTracker.getScheduleRoster() : '[]';
}
function getSafeScheduleBoard(dateStr, agentsPipe) {
  return (typeof SafeTracker !== 'undefined')
    ? SafeTracker.getScheduleBoard(dateStr, agentsPipe)
    : JSON.stringify({ error: 'SafeTracker not loaded.' });
}
function getSafeScheduleRange(startStr, endStr, agentsPipe) {
  return (typeof SafeTracker !== 'undefined')
    ? SafeTracker.getScheduleRange(startStr, endStr, agentsPipe)
    : JSON.stringify({ error: 'SafeTracker not loaded.' });
}
function getSafeScheduleDates() {
  return (typeof SafeTracker !== 'undefined') ? SafeTracker.getScheduleDates() : '[]';
}
function getSafeCapableScheduledForDate(dateStr, hStart, hEnd) {
  return (typeof SafeTracker !== 'undefined') ? SafeTracker.getCapableScheduledForDate(dateStr, hStart, hEnd) : '[]';
}
function setSafeAgentLang(name, lang) {
  return (typeof SafeTracker !== 'undefined') ? SafeTracker.setAgentLang(name, lang) : 'Error';
}

// --- OT OPEN SLOTS (WFM JSON export) ---
function importOtOpenSlots(t) {
  return (typeof OvertimeTracker !== 'undefined') ? OvertimeTracker.importOpenSlots(t) : 'Error';
}

// --- WORKFORCE TRACKER EXPORTS ---
function getWorkforceAnalytics(mode, date, type, region, cycleFilter) { 
  return (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker.getAnalytics(mode, date, type, region, cycleFilter) : "{}"; 
}
function importWorkforceData(sched, idp) {
  return (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker.importData(sched, idp) : "Error";
}
// Drains the deferred per-month unified-report archive queue (called by the client
// right after an import). Its own execution budget, so it can't time out the import.
function flushPendingArchives() {
  return (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker.flushPendingArchives() : '{}';
}

// ADMIN-ONLY — intentionally NOT wired to any UI button (too destructive for an
// everyday user). Run it manually from the Apps Script editor: rebuildDataSheets()
// or rebuildDataSheets(true). Clears ONLY the schedule + activity sheets (and, if
// asked, the monthly report archive) for a clean re-upload with no duplicates/
// orphans. EXPLICIT allow-list — never touches WF_MASTERLIST, WF_REGION_MAP,
// WF_LANG_MAP, WF_IDP, WF_OT_OPEN, GEM, Stats, DB_Sessions, Overtime_Tracking,
// Training_* or Outage History. Clears CONTENT (keeps the header row) so there is
// no slow row-reflow.
function rebuildDataSheets(alsoArchives) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targets = ['Raw Schedule', 'WF_ROLES', 'WF_COACHING', 'WF_FURLOUGH', 'WF_ABSENCES', 'WF_OVERTIME', 'Schedule_History'];
  if (alsoArchives) targets.push('Weekly_Archives_V3');
  var cleared = [];
  targets.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, Math.max(1, sh.getLastColumn())).clearContent();
    cleared.push(name);
  });
  try { PropertiesService.getScriptProperties().setProperty('WF_CACHE_VER', String(Date.now())); } catch (e) {}  // bust analytics cache
  Logger.log('[rebuildDataSheets] cleared: ' + cleared.join(', ') + (alsoArchives ? ' (incl. archives)' : ''));
  return JSON.stringify({ ok: true, cleared: cleared });
}

// ── Chunked upload transport ──────────────────────────────────────────────
// google.script.run reliably carries ~1 MB; large WFM/IDP pastes blow past
// that and get dropped silently with no error and no response. Chunk the
// text into ~80 KB pieces (safely under CacheService's 100 KB per-key cap),
// stash each in cache, then assemble + process server-side in one go.
function uploadChunk(token, index, total, chunk) {
  var cache = CacheService.getScriptCache();
  cache.put('uplk_' + token + '_' + index, chunk, 600); // 10 min TTL
  return index + 1;
}

function processChunkedImport(token, total, kind) {
  Logger.log('[chunkedImport] start: token=' + token + ' total=' + total + ' kind=' + kind);
  var cache = CacheService.getScriptCache();
  var keys = [];
  for (var i = 0; i < total; i++) keys.push('uplk_' + token + '_' + i);
  var bag = cache.getAll(keys);
  var parts = [];
  var missing = 0;
  for (var j = 0; j < total; j++) {
    var p = bag['uplk_' + token + '_' + j];
    if (p == null) { missing++; continue; }
    parts.push(p);
  }
  if (missing) throw new Error(missing + ' of ' + total + ' chunks missing from cache (TTL expired?)');
  cache.removeAll(keys);
  var fullText = parts.join('');
  Logger.log('[chunkedImport] assembled ' + fullText.length + ' chars');
  var t0 = new Date().getTime();
  var result;
  if (kind === 'idp') {
    result = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker.importData('', fullText) : 'Error';
  } else if (kind === 'otopen') {
    result = (typeof OvertimeTracker !== 'undefined') ? OvertimeTracker.importOpenSlots(fullText) : 'Error';
  } else {
    if (typeof ImportHandler !== 'undefined') {
      Logger.log('[chunkedImport] running ImportHandler...');
      ImportHandler.run(fullText);
      Logger.log('[chunkedImport] ImportHandler done at +' + (new Date().getTime() - t0) + 'ms');
    }
    Logger.log('[chunkedImport] running WorkforceTracker.importData...');
    result = (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker.importData(fullText, '') : 'Error';
  }
  Logger.log('[chunkedImport] done at +' + (new Date().getTime() - t0) + 'ms; result=' + result);
  return result;
}

// --- GRAND UNIFIED TRACKER & ARCHIVES ---
function getUnifiedReport(dateStr, cycleFilter) { 
  return (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker.getUnifiedReport(dateStr, cycleFilter) : "{}"; 
}
function archiveUnifiedReport(dateStr, cycleFilter) { 
  return (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker.archiveUnifiedReport(dateStr, cycleFilter) : "Error";
}
function getArchiveList() { 
  return (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker.getArchiveList() : "[]";
}
function getArchivedReport(period, cycleFilter) {
  return (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker.getArchivedReport(period, cycleFilter) : "{}";
}

// --- POWER OUTAGES (BC Hydro + Hydro One + Hydro-Québec) ---
function getPowerOutages() {
  return (typeof OutageTracker !== 'undefined') ? OutageTracker.fetchAll() : "{}";
}
function getOutageHistory() {
  return (typeof OutageTracker !== 'undefined') ? OutageTracker.getHistorySeries(26) : "[]";
}

// --- STAFFING BALANCE (current 15-min IDP bucket) ---
function getStaffingBalance() {
  return (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker.getStaffingBalance() : "{}";
}

// --- MANAGEMENT VIEW (corporate calendar: day / Sun-Sat week / month / quarter) ---
function getManagementDashboard(grain, refDate) {
  return (typeof ManagementView !== 'undefined') ? ManagementView.getDashboard(grain, refDate) : "{}";
}

// getOutageAgentCorrelation() defined in CORE/OutageTracker.gs
// getCoachingCadenceFlags() defined in CORE/AssignmentAnalyzer.gs

function fetchCoachingCadence(thresholdDays) {
  return (typeof getCoachingCadenceFlags === 'function') ? getCoachingCadenceFlags(thresholdDays) : "{\"flags\":[]}";
}

// --- MANAGER FEEDBACK (suggestions / ideas / bug reports) ---
function getCurrentUser() {
  return (typeof FeedbackTracker !== 'undefined') ? FeedbackTracker.whoAmI() : JSON.stringify({ email: '', name: '' });
}
function submitFeedback(type, message, page) {
  return (typeof FeedbackTracker !== 'undefined') ? FeedbackTracker.submit(type, message, page) : JSON.stringify({ ok: false, error: 'FeedbackTracker not loaded.' });
}
function getFeedbackList(limit) {
  return (typeof FeedbackTracker !== 'undefined') ? FeedbackTracker.getList(limit) : "[]";
}
