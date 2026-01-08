/**
 * MODULE: LOG SYNC
 * "Feeds the Beast" - Writes Dashboard data into the Master Spreadsheet.
 */

const LogSync = {
  
  // --- MASTER ACTION: COMMIT SHIFT ---
  commitShift: function(handoverNote) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = this._getCurrentDaySheet(ss);
    if (!sheet) return "Error: Could not find today's sheet.";

    let log = [];

    // 1. FILL WINDS
    if (typeof WeatherService !== 'undefined') {
        const weather = WeatherService.fetch();
        this._fillWeatherSection(sheet, weather);
        log.push("Winds Filled");
    }

    // 2. COMPILE ABSENCES
    if (typeof AgentMonitor !== 'undefined') {
        // Parse the current floor state directly
        const floorData = JSON.parse(AgentMonitor.getPayload()); 
        this._fillAbsenceSection(sheet, floorData);
        log.push("Absences Synced");
    }

    // 3. LOG HANDOVER NOTE (If provided)
    if (handoverNote) {
        this.writeToJournal("MRC and Bckup", handoverNote, "End of Shift");
        log.push("Handover Saved");
    }

    return log.join(", ");
  },

  // --- A. SMART JOURNAL (Writes to specific rows) ---
  writeToJournal: function(category, text, user) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = this._getCurrentDaySheet(ss);
    if (!sheet) return "Sheet Not Found";

    // Map Dashboard Categories to Spreadsheet Headers
    const headerMap = {
      "Critical": "Critical Event(s)",
      "Weather": "Weather Alert",
      "Maintenance": "Maintenance",
      "General": "MRC and Bckup"
    };

    const targetHeader = headerMap[category] || "MRC and Bckup";
    
    // Find the row
    const row = this._findRowByContent(sheet, 1, targetHeader); // Search Col A
    if (row > 0) {
       // Append to Column B (next to header)
       const currentVal = sheet.getRange(row, 2).getValue();
       const newVal = currentVal ? currentVal + "\n• " + text : "• " + text;
       sheet.getRange(row, 2).setValue(newVal);
       
       // Also save to DB_Journal for history
       this._appendToDb("DB_Journal", [new Date(), category, user, text]);
       return "Logged to " + targetHeader;
    }
    return "Header Not Found";
  },

  // --- B. ABSENCE COMPILER ---
  _fillAbsenceSection: function(sheet, floorData) {
     // Gather lists
     let planned = [];
     let unplanned = [];
     
     // Helper to check lists
     const checkList = (list) => {
         if(!list) return;
         list.forEach(a => {
             const name = a.name;
             // Check Status_Overrides or Import Data
             // Note: AgentMonitor categorization logic puts SICK in 'unplanned' array
             if (a.subStatus === "SICK" || a.subStatus === "NCNS" || a.subStatus === "AWOL") {
                 unplanned.push(`${name} (${a.subStatus})`);
             } else if (a.subStatus === "VACATION" || a.subStatus === "PERSONAL") {
                 planned.push(name);
             }
         });
     };

     // Scan all categories from AgentMonitor
     checkList(floorData.unplanned);
     checkList(floorData.vacation);
     checkList(floorData.planned);
     
     // Find "Absences" Section
     const startRow = this._findRowByContent(sheet, 1, "Absences");
     if (startRow > 0) {
         // Look in the next few rows for "Planned" and "Unplanned"
         // Based on CSV, they are usually right below "Absences"
         const pRow = this._findRowByContent(sheet, 1, "Planned", startRow, 5);
         const uRow = this._findRowByContent(sheet, 1, "Unplanned", startRow, 5);

         if (pRow > 0 && planned.length > 0) sheet.getRange(pRow, 2).setValue(planned.join(", "));
         if (uRow > 0 && unplanned.length > 0) sheet.getRange(uRow, 2).setValue(unplanned.join(", "));
     }
  },

  // --- C. WEATHER FILLER ---
  _fillWeatherSection: function(sheet, weatherData) {
    const startRow = this._findRowByContent(sheet, 6, "Key Locations"); // Search Col F
    if (startRow < 1) return;

    // Determine Column (Night/Day/Even)
    const hour = new Date().getHours();
    let colOffset = 2; // Default to Column H (Night)
    if (hour >= 6 && hour < 14) colOffset = 3; // Day
    else if (hour >= 14 && hour < 22) colOffset = 4; // Evening

    const targetCol = 6 + colOffset; // F is 6, so 6+2=8 (H)

    const cities = ["Toronto", "Ottawa", "Calgary", "Edmonton", "Vancouver", "Prince George", "Montreal", "Quebec City"];
    
    // Flatten data
    let flat = [];
    Object.values(weatherData.weather).forEach(arr => flat.push(...arr));

    cities.forEach((city, idx) => {
        const found = flat.find(c => c.name === city);
        if (found) {
            let val = found.wind.toString().replace("Light", "5");
            // Write to relative row (StartRow + 1 + index)
            sheet.getRange(startRow + 1 + idx, targetCol).setValue(val);
        }
    });
  },

  // --- D. SESSION LOGGER (Called by StatusTracker) ---
  logSession: function(agent, role, start, end) {
      const durMins = Math.round((end - start) / 60000);
      const hours = (durMins / 60).toFixed(2);
      this._appendToDb("DB_Sessions", [new Date(), agent, role, new Date(start), new Date(end), durMins, hours]);
  },

  // --- HELPERS ---
  _getCurrentDaySheet: function(ss) {
    const dayName = Utilities.formatDate(new Date(), "America/Toronto", "EEEE");
    return ss.getSheetByName(dayName);
  },

  _findRowByContent: function(sheet, colIndex, searchText, startRow = 1, limit = 100) {
     const data = sheet.getRange(startRow, colIndex, limit, 1).getValues();
     for (let i = 0; i < data.length; i++) {
         if (String(data[i][0]).includes(searchText)) {
             return startRow + i;
         }
     }
     return -1;
  },

  _appendToDb: function(tabName, rowData) {
     const ss = SpreadsheetApp.getActiveSpreadsheet();
     let sheet = ss.getSheetByName(tabName);
     if (!sheet) {
         sheet = ss.insertSheet(tabName);
         // Auto-header if new
         if (tabName === "DB_Journal") sheet.appendRow(["Timestamp", "Category", "User", "Note"]);
         if (tabName === "DB_Sessions") sheet.appendRow(["Timestamp", "Agent", "Role", "Start", "End", "Mins", "Hours"]);
     }
     sheet.appendRow(rowData);
  }
};

// --- EXPORTS ---
function commitShiftAction(note) { return LogSync.commitShift(note); }
function saveJournalEntry(cat, note) { return LogSync.writeToJournal(cat, note, "User"); }
