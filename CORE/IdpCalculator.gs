/**
 * MODULE: IDP CALCULATOR
 * Parses WFM Clipboard Data to calculate IDP %.
 * Logic: Sum(Occupied Seats) / Sum(Requirements) * 100
 */

const IdpCalculator = {
  
  process: function(rawText) {
    if (!rawText) return { success: false, msg: "No text provided" };

    try {
      const lines = rawText.split(/\r?\n/);
      let headers = [];
      let reqIdx = -1;
      let occIdx = -1;
      
      let totalReq = 0;
      let totalOcc = 0;
      
      // 1. Find Headers & Indices
      // We look for a line containing both "Requirements" and "Occupied"
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        // Normalize: remove quotes, handle CSV or Tab
        const parts = line.replace(/"/g, "").split(/,|\t/);
        
        // Dynamic Search for "Today" columns if possible, otherwise first pair found
        // Simple logic: Find first index of "Requirements" and "Occupied"
        const rI = parts.findIndex(p => p.trim().startsWith("Requirements"));
        const oI = parts.findIndex(p => p.trim().startsWith("Occupied Seats"));
        
        if (rI > -1 && oI > -1) {
           headers = parts;
           reqIdx = rI;
           occIdx = oI;
           // Start processing data from next line
           this._sumData(lines, i + 1, reqIdx, occIdx, (r, o) => { totalReq += r; totalOcc += o; });
           break;
        }
      }

      // 2. Calculate
      if (totalReq === 0) return { success: false, msg: "Could not find valid data" };
      
      const idpPercent = (totalOcc / totalReq) * 100;
      const finalVal = idpPercent.toFixed(1); // 1 decimal place
      
      // 3. Log it
      if (typeof StatsTracker !== 'undefined') {
         StatsTracker.logIdp(finalVal);
      }
      
      return { success: true, val: finalVal, msg: `Calculated: ${finalVal}%` };
      
    } catch (e) {
      console.error("IDP Parse Error", e);
      return { success: false, msg: "Parsing Error" };
    }
  },

  _sumData: function(lines, startRow, rIdx, oIdx, callback) {
     for (let i = startRow; i < lines.length; i++) {
        const parts = lines[i].replace(/"/g, "").split(/,|\t/);
        if (parts.length > Math.max(rIdx, oIdx)) {
           const req = parseFloat(parts[rIdx]) || 0;
           const occ = parseFloat(parts[oIdx]) || 0;
           callback(req, occ);
        }
     }
  }
};

// Global Export
function calculateAndLogIdp(text) { return JSON.stringify(IdpCalculator.process(text)); }
