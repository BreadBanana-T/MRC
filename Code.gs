/**
 * MAIN ROUTER
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.appUrl = ScriptApp.getService().getUrl(); 
  template.popoutParam = (e && e.parameter && e.parameter.popout) ? "true" : "false";
  template.modeParam = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : "";
  return template.evaluate()
      .setTitle('MRC Command Center')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

// --- DATA FETCHERS ---
function getFloorStatus() { return (typeof AgentMonitor !== 'undefined') ? AgentMonitor.getPayload() : "{}"; }
function getStatsHistory() { return (typeof StatsTracker !== 'undefined') ? StatsTracker.getHistory() : "[]"; }
function getLiveDashboardData() {
  try { return (typeof WeatherService !== 'undefined') ? JSON.stringify(WeatherService.fetch()) : "{}"; } 
  catch (e) { return "{}"; }
}
function getSystemNotifications() { return (typeof NotificationHandler !== 'undefined') ? NotificationHandler.getPending() : "[]"; }

// --- LOG RETRIEVAL FOR INSIGHTS ---
function getDailyRoleLogs(dateStr) {
  return fetchLogs(rowDate => rowDate === dateStr);
}

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
  // Filter by date (Col 0)
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
  const sheet = ss.getSheetByName("DB_Sessions"); // Reads local history
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
function updateAgentStatus(name, type, val) { if(typeof AgentMonitor!=='undefined') return AgentMonitor.setStatus(name, type, val); }
function updateAgentBreaks(name, json) { if(typeof AgentMonitor!=='undefined') return AgentMonitor.updateAgentBreaks(name, json); }
function submitOvertime(name, s, e, bs, be) { if(typeof AgentMonitor!=='undefined') return AgentMonitor.logOvertime(name, s, e, bs, be); }
function runCalculator(i, o) { if(typeof calculateMetrics!=='undefined') return calculateMetrics(i, o); return "{}"; }
function fetchScripts() { if(typeof getTeamScripts!=='undefined') return getTeamScripts(); return "[]"; }
function saveTeamScript(i, t, b, c) { if(typeof saveTeamScript!=='undefined') return ScriptHandler.save(i, t, b, c); }
function deleteTeamScript(i) { if(typeof deleteTeamScript!=='undefined') return ScriptHandler.delete(i); }
function saveJournalEntry(c, n) { if(typeof LogSync!=='undefined') return LogSync.writeToJournal(c, n, "User"); }
function commitShiftAction(n) { if(typeof LogSync!=='undefined') return LogSync.commitShift(n); }
function fillWindsToSheet() { if(typeof WeatherService!=='undefined') return LogSync.fillWinds(WeatherService.fetch()); }
function runImport(t) { if(typeof ImportHandler!=='undefined') return ImportHandler.run(t); }
