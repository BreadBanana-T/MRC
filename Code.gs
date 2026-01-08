/**
 * MRC CORE
 * Main Router for Command Center
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  
  // 1. Pass Real App URL (Prevents "White Page" on popout)
  template.appUrl = ScriptApp.getService().getUrl(); 

  // 2. Pass URL Parameters to bypass Sandbox restrictions
  template.popoutParam = (e && e.parameter && e.parameter.popout) ? "true" : "false";
  template.modeParam = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : "";

  return template.evaluate()
      .setTitle('MRC Command Center')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- CORE: INCLUDE FUNCTION ---
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* --- ROUTER: CALCULATOR --- */
function runCalculator(inText, outText) {
  if (typeof calculateMetrics !== 'undefined') {
    return calculateMetrics(inText, outText);
  }
  return JSON.stringify({ report: "Error: Calculator.gs file not found." });
}

/* --- ROUTER: STATS --- */
function getStatsHistory() {
  if (typeof StatsTracker !== 'undefined') return StatsTracker.getHistory();
  return JSON.stringify([]);
}

function logCurrentStats(svl) {
  if (typeof StatsTracker !== 'undefined') return StatsTracker.logHourlyStats(svl);
}

/* --- ROUTER: DASHBOARD --- */
function getFloorStatus() {
  if (typeof AgentMonitor !== 'undefined') return AgentMonitor.getPayload();
  return JSON.stringify({ error: "AgentMonitor missing" }); 
}

function getLiveDashboardData() {
  try {
    if (typeof WeatherService !== 'undefined') {
      return JSON.stringify(WeatherService.fetch());
    }
    return JSON.stringify({ weather: {}, alerts: [] });
  } catch (e) {
    return JSON.stringify({ weather: {}, alerts: [] });
  }
}

/* --- ROUTER: ACTIONS --- */
function updateAgentStatus(name, type, val) {
  if (typeof AgentMonitor !== 'undefined') return AgentMonitor.setStatus(name, type, val);
  return "Backend Error: AgentMonitor missing";
}

function updateScheduleBreak(name, index, start, end) {
  if (typeof AgentMonitor !== 'undefined') return AgentMonitor.updateBreak(name, index, start, end);
  return "Backend Error";
}

function updateAgentBreaks(name, json) {
  if (typeof AgentMonitor !== 'undefined') return AgentMonitor.updateAgentBreaks(name, json);
  return "Backend Error";
}

function confirmIEXUpdate(name) {
  if (typeof AgentMonitor !== 'undefined') return AgentMonitor.clearFlag(name);
}

function runImport(text, date) {
  if (typeof ImportHandler !== 'undefined') return ImportHandler.run(text, date);
  return "ImportHandler missing";
}

function submitOvertime(name, start, end, bStart, bEnd) {
   if (typeof AgentMonitor !== 'undefined') return AgentMonitor.logOvertime(name, start, end, bStart, bEnd);
   return "Backend Error";
}

/* --- SYSTEM NOTIFICATIONS --- */
function getSystemNotifications() {
    if (typeof NotificationHandler !== 'undefined') return NotificationHandler.getPending();
    return JSON.stringify([]);
}

function addSystemNotification(text) {
    if (typeof NotificationHandler !== 'undefined') return NotificationHandler.add("System", text);
}

function dismissSystemNotification(id) {
    if (typeof NotificationHandler !== 'undefined') return NotificationHandler.dismiss(id);
}

/* --- SCRIPT ROUTERS --- */
function fetchScripts() { 
    if (typeof getTeamScripts !== 'undefined') return getTeamScripts();
    return JSON.stringify([]);
}
function saveTeamScript(i, t, b, c) { 
    if (typeof saveTeamScript !== 'undefined') return ScriptHandler.save(i, t, b, c);
}
function deleteTeamScript(i) { 
    if (typeof deleteTeamScript !== 'undefined') return ScriptHandler.delete(i); 
}

/* --- LOGGING & SHEET FILLING ROUTERS --- */
function commitShiftAction(note) {
   if (typeof LogSync !== 'undefined') return LogSync.commitShift(note);
   return "LogSync Module Missing";
}

function saveJournalEntry(cat, note) {
   if (typeof LogSync !== 'undefined') return LogSync.writeToJournal(cat, note, "User");
   return "LogSync Module Missing";
}

function fillWindsToSheet() { 
  if (typeof WeatherService === 'undefined') return "Weather Module Missing";
  if (typeof LogSync === 'undefined') return "LogSync Module Missing";
  const data = WeatherService.fetch();
  return LogSync.fillWinds(data);
}
