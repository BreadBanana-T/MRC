/**
 * MODULE: IDP CALCULATOR
 * Parses WFM Clipboard Data to calculate IDP % and extract Forecast vs Actual series.
 * Logic: Sum(Occupied Seats) / Sum(Requirements) * 100
 */

const IdpCalculator = {
  
  process: function(rawText) {
    if (!rawText) return { success: false, msg: "No text provided" };

    try {
      const lines = rawText.split(/\r?\n/);
      const now = new Date();
      const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
      const todayStr = now.toLocaleDateString('en-US', options); // e.g., "Monday, January 12, 2026"
      
      let reqIdx = -1;
      let occIdx = -1;
      let timeIdx = -1;
      let startRow = -1;
      
      // --- HEADER DISCOVERY ---
      // We scan the first 20 lines to find the headers
      for (let i = 0; i < Math.min(lines.length, 20); i++) {
        const line = lines[i];
        // Split by comma (CSV) or Tab (Excel)
        const delim = line.includes("\t") ? "\t" : ",";
        const parts = line.split(delim).map(p => p.replace(/"/g, "").trim());

        for (let c = 0; c < parts.length; c++) {
           const header = parts[c].toLowerCase();
           
           // Detect Time Column
           if (header.includes("time") || header.match(/^\d{2}:\d{2}/)) timeIdx = c;

           // Detect Requirements (Allow "Req" or "Requirements")
           if (header.includes("requirement") || header === "req") {
               // If we haven't found one yet, or if this specific column is under today's date (logic simplified)
               if (reqIdx === -1 || parts.join("").includes(todayStr)) reqIdx = c;
           }
           
           // Detect Actual/Occupied (Allow "Occupied", "Open", or "Act")
           if (header.includes("occupied") || header === "open" || header.includes("actual")) {
               if (occIdx === -1 || parts.join("").includes(todayStr)) occIdx = c;
           }
        }
        
        // If we found the headers, the data usually starts on the NEXT row
        if (reqIdx > -1 && occIdx > -1) {
           startRow = i + 1;
           // If we didn't find a time column in headers, assume column 0
           if (timeIdx === -1) timeIdx = 0; 
           break;
        }
      }

      if (startRow === -1) return { success: false, msg: "Could not find 'Req/Requirements' and 'Open/Occupied' columns." };

      // --- DATA EXTRACTION ---
      let totalReq = 0;
      let totalOcc = 0;
      
      let seriesTime = [];
      let seriesReq = [];
      let seriesOcc = [];

      for (let i = startRow; i < lines.length; i++) {
         const line = lines[i].trim();
         if (!line) continue;
         
         const delim = line.includes("\t") ? "\t" : ",";
         const parts = line.split(delim);
         
         if (parts.length > Math.max(reqIdx, occIdx)) {
             // Extract and Clean Numbers
             const rStr = (parts[reqIdx] || "0").replace(/,/g, "").replace(/"/g, "");
             const oStr = (parts[occIdx] || "0").replace(/,/g, "").replace(/"/g, "");
             const timeStr = (parts[timeIdx] || "").replace(/"/g, "");

             // Skip lines that don't look like data (e.g. footers or empty cells)
             if (isNaN(parseFloat(rStr))) continue;

             const rVal = parseFloat(rStr) || 0;
             const oVal = parseFloat(oStr) || 0;
             
             totalReq += rVal;
             totalOcc += oVal;

             // Add to arrays for the graph
             // Only add if we have a valid time string (e.g., "00:00") to avoid clutter
             if (timeStr.includes(":")) {
                seriesTime.push(timeStr);
                seriesReq.push(rVal);
                seriesOcc.push(oVal);
             }
         }
      }

      // --- CALCULATION ---
      if (totalReq === 0) return { success: false, msg: "Total Requirements is 0" };
      
      const idpPercent = (totalOcc / totalReq) * 100;
      const finalVal = idpPercent.toFixed(1);
      
      // Log to history (optional, keeps your old logic working)
      if (typeof StatsTracker !== 'undefined') StatsTracker.logIdp(finalVal);
      
      return { 
        success: true, 
        val: finalVal, 
        msg: `IDP: ${finalVal}%`,
        graphData: {
          labels: seriesTime,
          forecast: seriesReq,
          actual: seriesOcc
        }
      };
      
    } catch (e) {
      console.error("IDP Parse Error", e);
      return { success: false, msg: "Parsing Error: " + e.message };
    }
  }
};

function calculateIdpFromText(text) { return JSON.stringify(IdpCalculator.process(text)); }
