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
      const todayStr = now.toLocaleDateString('en-US', options); 
      
      let reqIdx = -1;
      let occIdx = -1;
      let timeIdx = -1;
      let startRow = -1;
      
      // --- HEADER DISCOVERY ---
      // Scan the first 20 lines to find the headers
      for (let i = 0; i < Math.min(lines.length, 20); i++) {
        const line = lines[i];
        const delim = line.includes("\t") ? "\t" : ",";
        const parts = line.split(delim).map(p => p.replace(/"/g, "").trim());

        for (let c = 0; c < parts.length; c++) {
           const header = parts[c].toLowerCase(); 
           const originalHeader = parts[c];
           
           // 1. Detect Time Column
           if (header === "time" || header.match(/^\d{2}:\d{2}/)) {
               if(timeIdx === -1) timeIdx = c;
           }

           // 2. Detect Requirements (Forecast)
           if (header.includes("req")) {
               if (originalHeader.includes(todayStr) || reqIdx === -1) {
                   reqIdx = c;
               }
           }
           
           // 3. Detect Actual (Open)
           // Exclude "+/-" to avoid variance columns
           if ((header.includes("open") || header.includes("occupied") || header.includes("actual")) && !header.includes("+/-")) {
               if (originalHeader.includes(todayStr) || occIdx === -1) {
                   occIdx = c;
               }
           }
        }
        
        if (reqIdx > -1 && occIdx > -1) {
           startRow = i + 1;
           if (timeIdx === -1) timeIdx = 0; 
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
         
         if (parts.length > Math.max(reqIdx, occIdx)) {
             // Clean strings
             const rStr = (parts[reqIdx] || "0").replace(/,/g, "").replace(/"/g, "");
             const oStr = (parts[occIdx] || "0").replace(/,/g, "").replace(/"/g, "");
             const timeStr = (parts[timeIdx] || "").replace(/"/g, "").trim();

             // --- STRICT VALIDATION TO FIX "BOTTOM PART" ISSUE ---
             
             // 1. Stop if we hit a Footer row (Total, Average, etc.)
             if (timeStr.match(/^(total|grand|average|notes|count)/i)) break;

             // 2. Skip if Time doesn't look like a Time (Must contain ":")
             if (!timeStr.includes(":")) continue;

             // 3. Skip if values are not numbers
             if (isNaN(parseFloat(rStr)) && isNaN(parseFloat(oStr))) continue;

             const rVal = parseFloat(rStr) || 0;
             const oVal = parseFloat(oStr) || 0;
             
             totalReq += rVal;
             totalOcc += oVal;

             seriesTime.push(timeStr);
             seriesReq.push(rVal);
             seriesOcc.push(oVal);
         }
      }

      // --- CALCULATION ---
      if (totalReq === 0) return { success: false, msg: "Total Requirements is 0" };
      
      const idpPercent = (totalOcc / totalReq) * 100;
      const finalVal = idpPercent.toFixed(1);
      
      if (typeof StatsTracker !== 'undefined') StatsTracker.logIdp(finalVal);
      
      return { 
        success: true, 
        val: finalVal, 
        msg: `IDP Calculated: ${finalVal}%`,
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
