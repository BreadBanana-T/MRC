/**
 * MODULE: IDP CALCULATOR
 * Parses WFM Clipboard Data to calculate IDP %.
 * Logic: Sum(Occupied Seats) / Sum(Requirements) * 100
 * Supports: CSV (comma) and Excel (tab) formats.
 */

const IdpCalculator = {
  
  process: function(rawText) {
    if (!rawText) return { success: false, msg: "No text provided" };

    try {
      const lines = rawText.split(/\r?\n/);
      
      // 1. Determine "Today's" Date String to find the right column
      // Matches format in your CSV: "Monday, January 12, 2026"
      const now = new Date();
      const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
      const todayStr = now.toLocaleDateString('en-US', options); 
      // Note: This matches "Monday, January 12, 2026"
      
      let reqIdx = -1;
      let occIdx = -1;
      let startRow = -1;
      
      // 2. Find Headers
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        // Handle CSV quotes and split by comma or tab
        // Simple regex to split by comma but ignore commas inside quotes
        const parts = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(p => p.replace(/"/g, "").trim());
        
        // Strategy: Look for "Requirements" AND the Date
        // If not found (wrong date format), fallback to just "Requirements" (first found)
        
        // Find indices
        for (let c = 0; c < parts.length; c++) {
           const header = parts[c];
           if (header.includes("Requirements")) {
               // If we haven't found a specific date match yet, or this one matches today
               if (reqIdx === -1 || header.includes(todayStr)) {
                   reqIdx = c;
               }
           }
           if (header.includes("Occupied Seats")) {
               if (occIdx === -1 || header.includes(todayStr)) {
                   occIdx = c;
               }
           }
        }
        
        if (reqIdx > -1 && occIdx > -1) {
           startRow = i + 1;
           break;
        }
      }

      if (startRow === -1 || reqIdx === -1 || occIdx === -1) {
          return { success: false, msg: "Could not find 'Requirements' and 'Occupied Seats' columns." };
      }

      // 3. Sum Data
      let totalReq = 0;
      let totalOcc = 0;
      
      for (let i = startRow; i < lines.length; i++) {
         const line = lines[i].trim();
         if (!line) continue;
         
         // Split logic matching header detection
         let parts;
         if (line.includes("\t")) {
             parts = line.split("\t"); // Excel copy
         } else {
             parts = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/); // CSV
         }
         
         if (parts.length > Math.max(reqIdx, occIdx)) {
             // Clean strings (remove commas in numbers if present)
             const rVal = parseFloat(parts[reqIdx].replace(/,/g, "").replace(/"/g, "")) || 0;
             const oVal = parseFloat(parts[occIdx].replace(/,/g, "").replace(/"/g, "")) || 0;
             
             totalReq += rVal;
             totalOcc += oVal;
         }
      }

      // 4. Calculate
      if (totalReq === 0) return { success: false, msg: "Total Requirements is 0" };
      
      const idpPercent = (totalOcc / totalReq) * 100;
      const finalVal = idpPercent.toFixed(1);
      
      // 5. Log
      if (typeof StatsTracker !== 'undefined') {
         StatsTracker.logIdp(finalVal);
      }
      
      return { success: true, val: finalVal, msg: `IDP Calculated: ${finalVal}% (Req:${totalReq.toFixed(1)} / Occ:${totalOcc.toFixed(1)})` };
      
    } catch (e) {
      console.error("IDP Parse Error", e);
      return { success: false, msg: "Parsing Error: " + e.message };
    }
  }
};

// Global Export
function calculateIdpFromText(text) { return JSON.stringify(IdpCalculator.process(text)); }
