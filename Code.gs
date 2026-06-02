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
function fetchScripts() { if(typeof getTeamScripts!=='undefined') return getTeamScripts(); return "[]"; }
function saveTeamScript(i, t, b, c) { if(typeof saveTeamScript!=='undefined') return ScriptHandler.save(i, t, b, c); }
function deleteTeamScript(i) { if(typeof deleteTeamScript!=='undefined') return ScriptHandler.delete(i); }
function fillWindsToSheet() { if(typeof WeatherService!=='undefined') return LogSync.fillWinds(WeatherService.fetch()); }
function runImport(t) { if(typeof ImportHandler!=='undefined') return ImportHandler.run(t); }
function submitIdpValue(v) { if(typeof StatsTracker!=='undefined') return StatsTracker.logIdp(v); }

// --- WORKFORCE TRACKER EXPORTS ---
function getWorkforceAnalytics(mode, date, type, region, cycleFilter) { 
  return (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker.getAnalytics(mode, date, type, region, cycleFilter) : "{}"; 
}
function importWorkforceData(sched, idp) {
  return (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker.importData(sched, idp) : "Error";
}

// --- OVERTIME TRACKER (WFM-derived OT codes, onshore only) ---
function getOvertimeAnalytics(mode, date, region, cycle) {
  return (typeof OvertimeTracker !== 'undefined')
    ? OvertimeTracker.getAnalytics(mode, date, region, cycle)
    : JSON.stringify({ error: "OvertimeTracker not loaded." });
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

// --- STAFFING BALANCE (current 15-min IDP bucket) ---
function getStaffingBalance() {
  return (typeof WorkforceTracker !== 'undefined') ? WorkforceTracker.getStaffingBalance() : "{}";
}

// getOutageAgentCorrelation() defined in CORE/OutageTracker.gs
// getCoachingCadenceFlags() defined in CORE/AssignmentAnalyzer.gs

function fetchCoachingCadence(thresholdDays) {
  return (typeof getCoachingCadenceFlags === 'function') ? getCoachingCadenceFlags(thresholdDays) : "{\"flags\":[]}";
}
