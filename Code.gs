/**
 * MRC CORE
 * Main Router & Server Logic
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  
  // Pass App Data to Frontend
  template.appUrl = ScriptApp.getService().getUrl(); 
  template.popoutParam = (e && e.parameter && e.parameter.popout) ? "true" : "false";
  template.modeParam = (e && e.parameter && e.parameter.mode) ? e.parameter.mode : "";

  return template.evaluate()
      .setTitle('MRC Command Center')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- CORE: INCLUDE FUNCTION (THIS IS WHAT WAS MISSING/BROKEN) ---
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* --- DATA ROUTERS --- */
function getFloorStatus() {
  if (typeof AgentMonitor !== 'undefined') return AgentMonitor.getPayload();
  return JSON.stringify({ error: "AgentMonitor missing" }); 
}

function getLiveDashboardData() {
  try {
    if (typeof WeatherService !== 'undefined') return JSON.stringify(WeatherService.fetch());
    return JSON.stringify({ weather: {}, alerts: [] });
  } catch (e) { return JSON.stringify({ weather: {}, alerts: [] }); }
}

function getStatsHistory() {
  if (typeof StatsTracker !== 'undefined') return StatsTracker.getHistory();
  return JSON.stringify([]);
}

/* --- ACTION ROUTERS --- */
function updateAgentStatus(name, type, val) {
  if (typeof AgentMonitor !== 'undefined') return AgentMonitor.setStatus(name, type, val);
  return "Error";
}

function updateAgentBreaks(name, json) {
  if (typeof AgentMonitor !== 'undefined') return AgentMonitor.updateAgentBreaks(name, json);
  return "Error";
}

function submitOvertime(name, s, e, bs, be) {
   if (typeof AgentMonitor !== 'undefined') return AgentMonitor.logOvertime(name, s, e, bs, be);
   return "Error";
}

/* --- TOOL ROUTERS --- */
function runCalculator(inText, outText) {
  if (typeof calculateMetrics !== 'undefined') return calculateMetrics(inText, outText);
  return JSON.stringify({ report: "Error" });
}

function fetchScripts() { 
    if (typeof getTeamScripts !== 'undefined') return getTeamScripts();
    return "[]";
}
function saveTeamScript(i, t, b, c) { 
    if (typeof saveTeamScript !== 'undefined') return ScriptHandler.save(i, t, b, c);
}
function deleteTeamScript(i) { 
    if (typeof deleteTeamScript !== 'undefined') return ScriptHandler.delete(i); 
}

/* --- LOGGING ROUTERS --- */
function saveJournalEntry(cat, note) {
   if (typeof LogSync !== 'undefined') return LogSync.writeToJournal(cat, note, "User");
   return "LogSync Missing";
}

function commitShiftAction(note) {
   if (typeof LogSync !== 'undefined') return LogSync.commitShift(note);
   return "LogSync Missing";
}

function fillWindsToSheet() { 
  if (typeof WeatherService === 'undefined') return "Weather Module Missing";
  if (typeof LogSync === 'undefined') return "LogSync Module Missing";
  const data = WeatherService.fetch();
  return LogSync.fillWinds(data);
}
