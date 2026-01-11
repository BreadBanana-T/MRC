/**
 * MODULE: IMPORT HANDLER (WFM PARSER)
 * Parses "dirty" text from WFM, extracts Schedule, Region, Breaks, and Status.
 * Handles "Partial Sick" vs "Full Sick".
 */
const ImportHandler = {
  run: function(text) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Raw Schedule");
    if (!sheet) return "Error: 'Raw Schedule' sheet missing";
    
    // 1. CLEAN THE INPUT
    // Remove the "UnknownActive" garbage and split into lines
    const cleanText = text.replace(/UnknownActive/g, "\n");
    const lines = cleanText.split(/\r?\n/);
    
    const outputRows = [];
    
    // STATE VARIABLES
    let currentRegion = "Onshore"; // Default
    let currentAgentName = null;
    let currentAgentID = "";
    let currentShiftDate = "";
    let currentShiftStart = "";
    let currentShiftEnd = "";
    let currentBreaks = [];
    let isOff = false;
    
    // Flags for status detection
    let hasWorkedHours = false;
    let sickType = ""; // SICK, SICU, MED, etc.
    let isNCNS = false;
    let isAWOL = false;

    // --- HELPER TO SAVE CURRENT AGENT ---
    const saveAgent = () => {
        if (currentAgentName) {
            // Determine Status based on flags
            let status = "Active";
            let subStatus = "";

            if (isOff) {
                status = "Off";
            } else if (isNCNS) {
                status = "Absent";
                subStatus = "NCNS";
            } else if (isAWOL) {
                status = "Absent";
                subStatus = "AWOL";
            } else if (sickType) {
                // THE CRITICAL FIX:
                // If they have a start/end time AND a sick code, it's Partial.
                // If the shift appears to be empty/00:00, it's Full Sick.
                if (hasWorkedHours) {
                    status = "Active"; // Keep them on the board
                    subStatus = "Leaving Early (Sick)";
                } else {
                    status = "Absent";
                    subStatus = sickType; // e.g., "SICK"
                }
            } else {
                status = "Active"; // Default
            }

            // Calculate Shift Name (Day/Night/Evening)
            let shiftName = "Day";
            if (currentShiftStart.includes("PM")) {
                const hour = parseInt(currentShiftStart);
                if (hour >= 9 || hour === 12) shiftName = "Night"; // 12 PM is noon, 12 AM is midnight... logic check needed
                else if (hour >= 1 && hour < 9) shiftName = "Evening";
            } else if (currentShiftStart.includes("AM")) {
                const hour = parseInt(currentShiftStart);
                if (hour < 5 || hour === 12) shiftName = "Night"; // 12 AM is night
            }

            // Format Breaks as JSON
            const breaksJson = JSON.stringify(currentBreaks);

            // PUSH ROW: [Name, ID, Date, Start, End, ShiftName, Region, Breaks, Status, SubStatus]
            // Only add if they are NOT Off (unless you want to track Offs too)
            if (status !== "Off") {
                outputRows.push([
                    currentAgentName,
                    currentAgentID,
                    currentShiftDate,
                    currentShiftStart,
                    currentShiftEnd,
                    shiftName,
                    currentRegion,
                    breaksJson,
                    status,
                    subStatus
                ]);
            }
        }
    };

    // --- MAIN PARSE LOOP ---
    for (let i = 0; i < lines.length; i++) {
        let line = lines[i].trim();
        if (line.length < 2) continue;

        // A. DETECT REGION (MU Set)
        if (line.includes("MU Set:")) {
            // Heuristic: "MONIT ALL" usually implies Onshore. Adjust keywords as needed.
            if (line.toUpperCase().includes("MONIT")) currentRegion = "Onshore";
            else if (line.toUpperCase().includes("OFFSHORE")) currentRegion = "Offshore";
            else currentRegion = "Onshore"; // Fallback
        }

        // B. DETECT NEW AGENT
        else if (line.startsWith("Agent:")) {
            saveAgent(); // Save previous agent before starting new one
            
            // Reset State
            currentAgentName = null;
            currentAgentID = "";
            currentShiftDate = "";
            currentShiftStart = "";
            currentShiftEnd = "";
            currentBreaks = [];
            isOff = false;
            hasWorkedHours = false;
            sickType = "";
            isNCNS = false;
            isAWOL = false;

            // Parse Name & ID: "Agent: 326333 Aazzaz, Hamza"
            const parts = line.split(":");
            if (parts.length > 1) {
                const info = parts[1].trim().split(" ");
                currentAgentID = info[0]; // 326333
                // Remainder is name (Aazzaz, Hamza)
                currentAgentName = info.slice(1).join(" ");
            }
        }

        // C. DETECT DATE & MAIN SHIFT
        // Matches: "1/10/26 12:30 AM 9:30 AM Open/Ouvert..." or "1/11/26 Off"
        else if (line.match(/^\d{1,2}\/\d{1,2}\/\d{2}/)) {
            // Check if this is the "Current" date or "Next" date? 
            // The paste often contains 2 days. We usually only want the first one or Today's.
            // For now, we take the first date found for the agent.
            if (currentShiftDate === "") {
                const dateParts = line.split(" ");
                currentShiftDate = dateParts[0];

                if (line.includes("Off")) {
                    isOff = true;
                } else {
                    // Extract Start/End times
                    // Regex for times: 12:30 AM
                    const times = line.match(/\d{1,2}:\d{2}\s(?:AM|PM)/g);
                    if (times && times.length >= 2) {
                        currentShiftStart = times[0];
                        currentShiftEnd = times[1];
                        hasWorkedHours = true;
                    }
                }
            }
        }

        // D. DETECT BREAKS / LUNCH
        else if (line.includes("Break") || line.includes("Lunch") || line.includes("Pause") || line.includes("Repas")) {
            // "TI 1st Break 2:15 AM 2:30 AM"
            const times = line.match(/\d{1,2}:\d{2}\s(?:AM|PM)/g);
            if (times && times.length >= 2) {
                let type = "Break";
                if (line.toLowerCase().includes("lunch") || line.toLowerCase().includes("repas")) type = "Lunch";
                
                currentBreaks.push({
                    type: type,
                    start: times[0],
                    end: times[1]
                });
            }
        }

        // E. DETECT SICKNESS / EXCEPTIONS
        else if (line.includes("SICK") || line.includes("SICU") || line.includes("MALADIE")) {
            // "SICU Maladie imprévue... 3:45 AM 8:00 AM"
            sickType = "SICK";
            
            // Check if this line itself has hours.
            // Sometimes the main line is "Open", but then a sub-line says "SICU 3am 8am".
            // We already marked `hasWorkedHours = true` from the main line. 
            // If the MAIN line was blank, `hasWorkedHours` would be false.
        }
        else if (line.includes("NCNS")) isNCNS = true;
        else if (line.includes("AWOL")) isAWOL = true;
    }

    // Save the very last agent
    saveAgent();

    // 4. WRITE TO SHEET
    sheet.clearContents();
    // Headers (Optional, but good for debugging)
    // sheet.appendRow(["Name", "ID", "Date", "Start", "End", "Shift", "Region", "Breaks", "Status", "SubStatus"]);
    
    if (outputRows.length > 0) {
        sheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
        
        // SYNC TRIGGER
        if (typeof AgentMonitor !== 'undefined') AgentMonitor.syncFromRaw();
        
        return `Imported ${outputRows.length} agents successfully.`;
    }
    return "No valid agent data found.";
  }
};
