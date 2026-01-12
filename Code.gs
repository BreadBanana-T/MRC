/**
 * MAIN ROUTER
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.appUrl = ScriptApp.getService().getUrl();
  template.popoutParam = (e && e.parameter && e.parameter.popout) ? "true" : "false";
  template.modeParam = (e && e.parameter && e.parameter.mode) ?
      e.parameter.mode : "";
  return template.evaluate()
      .setTitle('MRC Command Center')
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
