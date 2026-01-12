/**
 * MODULE: CALENDAR HANDLER
 * Fetches and filters events 45 days out.
 * FILTERS:
 * - INCLUDE: "Weekly Operational Meetings", "Scheduled Maintenance", "DEO", "Monthly MRC Meeting"
 * - EXCLUDE: "Prep", "Out of Office", "OOO", "Canceled", "Declined"
 */

const CalendarHandler = {
  
  getEvents: function() {
    try {
      const cal = CalendarApp.getDefaultCalendar();
      if (!cal) return [];
      
      const now = new Date();
      const future = new Date(now);
      future.setDate(now.getDate() + 45); // 45 Day Lookahead
      
      const events = cal.getEvents(now, future);
      
      // CONFIGURATION
      const validTitles = [
          "Weekly Operational Meetings", 
          "Scheduled Maintenance",
          "DEO",
          "Monthly MRC Meeting"
      ];
      
      const blockedTerms = ["Prep", "Out of Office", "OOO", "Canceled", "Declined", "Tentative"];
  
      const mapped = events.filter(e => {
        const t = e.getTitle();
        // 1. Check Include List
        const hasValid = validTitles.some(vt => t.includes(vt));
        // 2. Check Exclude List (Overrides Include)
        const hasBlocked = blockedTerms.some(bt => t.toLowerCase().includes(bt.toLowerCase()));
        
        return hasValid && !hasBlocked;
      }).map(e => ({
        title: e.getTitle(),
        description: e.getDescription() || "",
        date: Utilities.formatDate(e.getStartTime(), "America/Toronto", "MMM dd"),
        startTime: Utilities.formatDate(e.getStartTime(), "America/Toronto", "HH:mm"),
        endTime: Utilities.formatDate(e.getEndTime(), "America/Toronto", "HH:mm"),
        isAllDay: e.isAllDayEvent(),
        type: (e.getTitle().toLowerCase().includes("maintenance")) ? "maintenance" : "meeting"
      }));
      
      return JSON.stringify(mapped);
    } catch (e) {
      console.error("Calendar Error", e);
      return "[]";
    }
  }
};

// Global Export
function getDailyCalendarEvents() { return CalendarHandler.getEvents(); }
