/**
 * MODULE: IDP CALCULATOR
 * Parses WFM Clipboard Data to calculate IDP % and extract Forecast vs Actual series.
 * Logic: Sum(Open/Occupied) / Sum(Requirements) * 100
 */

const IdpCalculator = {
  
  process: function(rawText) {
    if (!rawText) return { success: false, msg: "No text provided" };

    try {
      const lines = rawText.split(/\r?\n/);
      const now = new Date();
      const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
      // Example: "Monday, January 12, 2026" matches your Excel header format
      const todayStr = now.toLocaleDateString('en-US', options); 
      
      let reqIdx = -1;
      let occIdx = -1;
      let timeIdx = -1;
      let startRow = -1;
      
      // --- HEADER DISCOVERY ---
      // Scan the first 20 lines to find the headers
      for (let i = 0; i < Math.min(lines.length, 20); i++) {
        const line = lines[i];
        // Handle Tab (Excel copy) or Comma (CSV)
        const delim = line.includes("\t") ? "\t" : ",";
        const parts = line.split(delim).map(p => p.replace(/"/g, "").trim());

        for (let c = 0; c < parts.length; c++) {
           const header = parts[c].toLowerCase(); // Normalize to lowercase for checking
           const originalHeader = parts[c]; // Keep original for date checking
           
           // 1. Detect Time Column
           if (header === "time" || header.match(/^\d{2}:\d{2}/)) {
               if(timeIdx === -1) timeIdx = c;
           }

           // 2. Detect Requirements (Forecast)
           // Must include "req" AND the date string (e.g. "Requirements Monday, January 12, 2026")
           // OR if no specific date is found in headers, just take the first "Requirements"
           if (header.includes("req")) {
               if (originalHeader.includes(todayStr) || reqIdx === -1) {
                   reqIdx = c;
               }
           }
           
           // 3. Detect Actual (Open / Occupied)
           // Your file has "Open" and "Open +/-". We want "Open". 
           // We explicitly exclude "+/-" to avoid the variance column.
           if ((header.includes("open") || header.includes("occupied") || header.includes("actual")) && !header.includes("+/-")) {
               if (originalHeader.includes(todayStr) || occIdx === -1) {
                   occIdx = c;
               }
           }
        }
        
        // If we found both columns, the data starts on the NEXT row
        if (reqIdx > -1 && occIdx > -1) {
           startRow = i + 1;
           if (timeIdx === -1) timeIdx = 0; // Default to first column if Time not found
           break;
        }
      }

      if (startRow === -1) return { success: false, msg: `Could not find 'Requirements' and 'Open' columns for ${todayStr}.` };

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
         
         // Ensure row has enough columns
         if (parts.length > Math.max(reqIdx, occIdx)) {
             // Extract and Clean Numbers
             const rStr = (parts[reqIdx] || "0").replace(/,/g, "").replace(/"/g, "");
             const oStr = (parts[occIdx] || "0").replace(/,/g, "").replace(/"/g, "");
             const timeStr = (parts[timeIdx] || "").replace(/"/g, "");

             // Skip invalid lines
             if (isNaN(parseFloat(rStr)) && isNaN(parseFloat(oStr))) continue;

             const rVal = parseFloat(rStr) || 0;
             const oVal = parseFloat(oStr) || 0;
             
             totalReq += rVal;
             totalOcc += oVal;

             // Prepare Graph Data
             // Only add if we have a valid time-looking string to keep the X-axis clean
             if (timeStr.includes(":") || timeStr.match(/^\d+$/)) { 
                seriesTime.push(timeStr);
                seriesReq.push(rVal);
                seriesOcc.push(oVal);
             }
         }
      }

      // --- CALCULATION ---
      if (totalReq === 0) return { success: false, msg: "Total Requirements is 0" };
      
      // Standard IDP % Calculation: (Actual / Requirements) * 100
      const idpPercent = (totalOcc / totalReq) * 100;
      const finalVal = idpPercent.toFixed(1);
      
      // Log for history
      if (typeof StatsTracker !== 'undefined') StatsTracker.logIdp(finalVal);
      
      return { 
        success: true, 
        val: finalVal, 
        msg: `IDP Calculated: ${finalVal}%`,
        graphData: {
          labels: seriesTime,
          forecast: seriesReq, // Requirements
          actual: seriesOcc    // Open (Actual Staff)
        }
      };
      
    } catch (e) {
      console.error("IDP Parse Error", e);
      return { success: false, msg: "Parsing Error: " + e.message };
    }
  }
};

function calculateIdpFromText(text) { return JSON.stringify(IdpCalculator.process(text)); }
