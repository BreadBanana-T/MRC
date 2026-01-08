// ... (Top of file remains the same) ...

  // --- 1. ROLES & ABSENCE (Status_Overrides) ---
  updateStatus: function(name, type, value) {
    const sheet = this._getSheet("Status_Overrides", ["Timestamp", "Agent Name", "Role", "Absent Type", "DateStr"]);
    const dateStr = this._getTodayStr(); 
    const cleanName = String(name).trim().toLowerCase();
    
    const data = sheet.getDataRange().getValues();
    let foundRow = -1;
    let oldData = null;

    for (let i = 1; i < data.length; i++) {
      if (this._matchRow(data[i], cleanName, dateStr, 1, 4)) {
        foundRow = i + 1;
        oldData = data[i]; // Capture previous state
        break;
      }
    }

    const now = new Date();

    // --- SESSION TRACKING LOGIC ---
    if (foundRow > -1 && oldData && typeof LogSync !== 'undefined') {
        const oldRole = oldData[2]; // Role Column
        const oldTime = oldData[0]; // Timestamp Column

        // If Role changed (e.g., SAFE -> Active) OR Role removed
        if (type === 'role' && oldRole && oldRole !== "" && oldRole !== value) {
             const startEpoch = new Date(oldTime).getTime();
             const endEpoch = now.getTime();
             // Log to DB_Sessions
             LogSync.logSession(name, oldRole, startEpoch, endEpoch);
        }
    }
    // -----------------------------

    if (foundRow > -1) {
// ... (Rest of file remains exactly the same as before) ...
