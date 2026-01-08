/**
 * MRC CORE
 * Main Router for Command Center
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  
  // --- CRITICAL FIX: Define appUrl here so Index.html doesn't crash ---
  template.appUrl = ScriptApp.getService().getUrl(); 

  return template.evaluate()
      .setTitle('MRC Command Center')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* --- ROUTER: CALCULATOR --- */
function runCalculator(inText, outText) {
  if (typeof calculateMetrics !== 'undefined') {
    const resultJson = calculateMetrics(inText, outText);
    return resultJson;
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
    // Fallback if WeatherService is missing (should not happen if Weather.gs is present)
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
