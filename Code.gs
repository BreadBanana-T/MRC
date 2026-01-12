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
function getFloorStatus() { return (typeof AgentMonitor !== 'undefined') ?
    AgentMonitor.getPayload() : "{}"; }
function getStatsHistory() { return (typeof StatsTracker !== 'undefined') ? StatsTracker.getHistory() : "[]";
}
function getIdpHistory() { return (typeof StatsTracker !== 'undefined') ? StatsTracker.getIdpHistory() : "[]"; 
}
function getLiveDashboardData() {
  try { return (typeof WeatherService !== 'undefined') ? JSON.stringify(WeatherService.fetch()) : "{}";
  } 
  catch (e) { return "{}"; }
}
function getSystemNotifications() { return (typeof NotificationHandler !== 'undefined') ? NotificationHandler.getPending() : "[]";
}

// --- CALENDAR SYNC (STRICT FILTER + 45 DAYS + DESCRIPTION) ---
function getDailyCalendarEvents() {
  try {
    const cal = CalendarApp.getDefaultCalendar();
    if (!cal) return "[]";
    
    const now = new Date();
    const future = new Date(now);
    future.setDate(now.getDate() + 45); // Look 45 days ahead
    
    // Fetch events
    const events = cal.getEvents(now, future);
    const validTitles = ["Weekly Operational Meetings", "Scheduled Maintenance"];

    const mapped = events.filter(e => {
      const t = e.getTitle();
      // Strict check: Must contain one of the valid phrases
      return validTitles.some(vt => t.includes(vt));
    }).map(e => ({
      title: e.getTitle(),
      description: e.getDescription() || "", // Get Description
      date: Utilities.formatDate(e.getStartTime(), "America/Toronto", "MMM dd"),
      startTime: Utilities.formatDate(e.getStartTime(), "America/Toronto", "HH:mm"),
      endTime: Utilities.formatDate(e.getEndTime(), "America/Toronto", "HH:mm"),
      isAllDay: e.isAllDayEvent(),
      type: (e.getTitle().toLowerCase().includes("maintenance")) ? "maintenance" : "meeting"
    }));
    
    return JSON.stringify(mapped);
  } catch (e) { return "[]"; }
}

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
  const sheet = ss.getSheetByName("DB_Sessions");
  // Reads local history
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
function updateAgentStatus(name, type, val) { if(typeof AgentMonitor!=='undefined') return AgentMonitor.setStatus(name, type, val);
}
function updateAgentBreaks(name, json) { if(typeof AgentMonitor!=='undefined') return AgentMonitor.updateAgentBreaks(name, json); }
function submitOvertime(name, s, e, bs, be) { if(typeof AgentMonitor!=='undefined') return AgentMonitor.logOvertime(name, s, e, bs, be);
}
function runCalculator(i, o) { if(typeof calculateMetrics!=='undefined') return calculateMetrics(i, o); return "{}"; }
function fetchScripts() { if(typeof getTeamScripts!=='undefined') return getTeamScripts(); return "[]";
}
function saveTeamScript(i, t, b, c) { if(typeof saveTeamScript!=='undefined') return ScriptHandler.save(i, t, b, c); }
function deleteTeamScript(i) { if(typeof deleteTeamScript!=='undefined') return ScriptHandler.delete(i);
}
function fillWindsToSheet() { if(typeof WeatherService!=='undefined') return LogSync.fillWinds(WeatherService.fetch()); }
function runImport(t) { if(typeof ImportHandler!=='undefined') return ImportHandler.run(t); }

// NEW: IDP Logging
function submitIdpValue(val) { if(typeof StatsTracker!=='undefined') return StatsTracker.logIdp(val); }
