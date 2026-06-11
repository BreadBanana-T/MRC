/**
 * MODULE: CALENDAR HANDLER
 * Fetches and filters events 45 days out.
 * FILTERS:
 * - INCLUDE: "Weekly Operational Meetings", "Scheduled Maintenance", "DEO", "Monthly MRC Meeting"
 * - EXCLUDE: "Prep", "Out of Office", "OOO", "Canceled", "Declined"
 */

const CalendarHandler = {

  // When the webapp is embedded in sites.google.com the iframe context can
  // interfere with CalendarApp permissions; try the OAuth+UrlFetch path as a
  // fallback so events still render on an embedded dashboard.
  getEvents: function() {
    const errors = [];
    let raw = null;

    // Primary: CalendarApp (works when deployed "Execute as: Me")
    try {
      const cal = CalendarApp.getDefaultCalendar();
      if (!cal) throw new Error("getDefaultCalendar() returned null");

      const now = new Date();
      const future = new Date(now);
      future.setDate(now.getDate() + 45);
      raw = cal.getEvents(now, future).map(e => ({
        title: e.getTitle(),
        description: e.getDescription() || "",
        start: e.getStartTime(),
        end: e.getEndTime(),
        isAllDay: e.isAllDayEvent()
      }));
    } catch (e) {
      errors.push("CalendarApp: " + e.message);
    }

    // Fallback: Calendar API v3 via UrlFetch + script OAuth token
    if (raw === null) {
      try {
        const token = ScriptApp.getOAuthToken();
        const now = new Date();
        const future = new Date(now); future.setDate(now.getDate() + 45);
        const url = "https://www.googleapis.com/calendar/v3/calendars/primary/events" +
          "?timeMin=" + encodeURIComponent(now.toISOString()) +
          "&timeMax=" + encodeURIComponent(future.toISOString()) +
          "&singleEvents=true&orderBy=startTime&maxResults=250";
        const res = UrlFetchApp.fetch(url, {
          headers: { Authorization: "Bearer " + token },
          muteHttpExceptions: true
        });
        const code = res.getResponseCode();
        if (code !== 200) throw new Error("HTTP " + code + ": " + res.getContentText().substring(0, 200));
        const body = JSON.parse(res.getContentText());
        raw = (body.items || []).map(ev => ({
          title: ev.summary || "(No title)",
          description: ev.description || "",
          start: new Date(ev.start.dateTime || ev.start.date),
          end: new Date(ev.end.dateTime || ev.end.date),
          isAllDay: !!ev.start.date
        }));
      } catch (e) {
        errors.push("API v3: " + e.message);
      }
    }

    if (raw === null) {
      console.error("Calendar fetch failed:", errors.join(" | "));
      // Surface to UI so a misconfigured deployment is visible, not silent.
      return JSON.stringify({ __error: errors.join(" | ") || "Unknown calendar error" });
    }

    const validTitles = ["Weekly Operational Meetings", "Scheduled Maintenance", "DEO", "Monthly MRC Meeting"];
    const blockedTerms = ["Prep", "Out of Office", "OOO", "Canceled", "Declined", "Tentative"];

    const filtered = raw.filter(e => {
      const t = e.title || "";
      const hasValid = validTitles.some(vt => t.includes(vt));
      const hasBlocked = blockedTerms.some(bt => t.toLowerCase().includes(bt.toLowerCase()));
      return hasValid && !hasBlocked;
    }).map(e => ({
      title: e.title,
      description: e.description,
      date: Utilities.formatDate(e.start, "America/Toronto", "MMM dd"),
      startTime: Utilities.formatDate(e.start, "America/Toronto", "HH:mm"),
      endTime: Utilities.formatDate(e.end, "America/Toronto", "HH:mm"),
      isAllDay: e.isAllDay,
      type: (e.title || "").toLowerCase().includes("maintenance") ? "maintenance" : "meeting"
    }));

    return JSON.stringify(filtered);
  }
};

// Router export getDailyCalendarEvents() lives in Code.gs.
