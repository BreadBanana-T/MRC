/**
 * MODULE: CALCULATOR
 * Status: FIXED Parsing & ASA.
 * - ASA: Uses global regex scan (no longer relies on column position).
 * - Sections: Flexible header detection.
 * - Stats: Matches "Real Stats" filtering logic.
 */

function runCalculator(inText, outText) {
  try {
    return calculateMetrics(inText, outText);
  } catch (e) {
    return JSON.stringify({
      svl: "ERR", ack: "ERR", 
      report: "CALC ERROR:\n" + e.toString(),
      trendIn: "0%", trendOut: "0%"
    });
  }
}

function calculateMetrics(inText, outText) {
  let stats = {
    svl: "0%", ack: "0s", 
    trendData: { labels: [], actual: [], trend: [] }, 
    trendIn: "0%", trendOut: "0%",
    asa: "0s", inSVL: "0%",
    safeSL: "N/A", 
    report: ""
  };

  // 1. STRICT LIST (Trend Volume & SVL)
  const LIST_TREND = [
      "1-FIRE", "1-GAS", "1-H/U", "1-MED", 
      "2-FARM", 
      "3-VID", 
      "4-BURG", "4-COMM", "4-TAMP", 
      "6-O/C"
  ];

  // 2. RESPONSE LIST (ACK - Excludes Admin Signals)
  const LIST_ACK = [
      "1-FIRE", "1-GAS", "1-H/U", "1-MED", 
      "2-FARM", 
      "3-VID", "3-LWK", 
      "4-BURG", "4-TAMP" 
  ];
  
  // ----------------------------------------------------
  // A. INBOUND PARSING (Global Search)
  // ----------------------------------------------------
  if (inText) {
    // 1. Find ASA (Time format like 0:31, 00:31, or 31s) anywhere in text
    // Looks for "ASA" followed strictly by time, OR just the time in the Monit row
    const asaMatch = inText.match(/ASA[\s\t]*(\d{1,2}:\d{2}|\d+s)/i) || 
                     inText.match(/Average Speed[\s\w]*(\d{1,2}:\d{2})/i);
    
    if (asaMatch) {
        stats.asa = fmt(asaMatch[1]);
    } else {
        // Fallback: Look for the Monit line and grab the time-like string in 3rd/4th position
        const monitLine = inText.split(/\r?\n/).find(l => l.includes("Monit"));
        if (monitLine) {
            const timePart = monitLine.match(/(\d{1,2}:\d{2})/);
            if (timePart) stats.asa = fmt(timePart[1]);
        }
    }

    // 2. Find Inbound SVL (Number > 50 near the end of Monit line)
    const monitLine = inText.split(/\r?\n/).find(l => l.includes("Monit") && l.includes("Intraday"));
    if (monitLine) {
        const parts = monitLine.match(/(\d{2,3})%?/g); // Grab all numbers > 10
        if (parts) {
            // Usually the last high number is SL
            const slCandidates = parts.filter(n => parseInt(n) > 50 && parseInt(n) <= 100);
            if (slCandidates.length > 0) stats.inSVL = slCandidates[slCandidates.length - 1] + "%";
        }
    }

    // 3. Find Trend Inbound (Signed percentage)
    const trendLine = inText.split(/\r?\n/).find(l => l.includes("Monit") && l.includes("Last 60"));
    if (trendLine) {
        const diffMatch = trendLine.match(/([+-]?\d+%)/);
        if (diffMatch) stats.trendIn = diffMatch[1];
    }
  }

  // ----------------------------------------------------
  // B. OUTBOUND PARSING (Robust Line Scanner)
  // ----------------------------------------------------
  if (outText) {
    // Clean artifacts
    const cleanOut = outText
        .replace(/–/g, '-').replace(/—/g, '-').replace(/\u00A0/g, ' ')
        .replace(/(\d+-[A-Z\/]+)([A-Z])/g, '$1 $2')
        .replace(/(\d)(\d{2}:\d{2}:\d{2})/g, '$1 $2');

    const lines = cleanOut.split(/\r?\n/);
    let currentMode = "NONE"; 
    
    let svlVol = 0, svlW = 0; 
    let ackVol = 0, ackW = 0;
    let trendDiff = 0, trendRef = 0;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (!line) continue;

        // MODE SWITCHING (Keyword search)
        if (line.match(/Intraday/i)) { currentMode = "INTRA"; continue; }
        if (line.match(/Last 60/i)) { currentMode = "LAST60"; continue; }
        if (line.match(/Last 15/i) || line.includes("Pending")) { currentMode = "NONE"; continue; }

        // ROW DETECTION
        const codeMatch = line.match(/^(\d+-[A-Z\/]+)/);
        
        // SAFE STATS (Any line with SAFE/LIF)
        if (currentMode === "INTRA" && line.match(/(SAFE|LIF)/i)) {
             // Priority: Look for percentage sign
             const pct = line.match(/(\d{2,3})%/);
             if (pct) {
                 stats.safeSL = pct[1] + "%";
             } else {
                 // Fallback: Look for last number in the line (usually SL)
                 const nums = line.match(/(\d{2,3})\b/g);
                 if (nums) {
                     const lastNum = parseInt(nums[nums.length - 1]);
                     if (lastNum > 50 && lastNum <= 100) stats.safeSL = lastNum + "%";
                 }
             }
        }

        if (codeMatch) {
            const code = codeMatch[1];
            const parts = line.split(/\s+/);
            // Anchor: Time Column (H:MM:SS)
            const timeIdx = parts.findIndex(p => p.match(/^\d{1,2}:\d{2}:\d{2}$/));

            // Must see Vol (index -3) and SL (index +1)
            if (timeIdx > -1 && timeIdx >= 3) {
                const colVol = parseInt(parts[timeIdx - 3].replace(/,/g,'')) || 0;
                const colHandled = parseInt(parts[timeIdx - 2].replace(/,/g,'')) || 0; // Trend column
                const colDiff = parseInt(parts[timeIdx - 1].replace(/,/g,'')) || 0;
                
                // INTRADAY CALC
                if (currentMode === "INTRA") {
                    const timeSec = dur(parts[timeIdx]);
                    // SL is usually right after time
                    let slVal = 0;
                    if (parts[timeIdx + 1]) slVal = parseFloat(parts[timeIdx + 1]);

                    if (checkList(code, LIST_ACK)) {
                        ackVol += colVol;
                        ackW += (colVol * timeSec);
                        svlVol += colVol;
                        svlW += (colVol * slVal);
                    }
                }
                
                // LAST 60 CALC
                else if (currentMode === "LAST60") {
                    if (checkList(code, LIST_TREND)) {
                        trendDiff += colDiff;
                        trendRef += colHandled;
                        
                        if (colVol > 0 || colHandled > 0) {
                            stats.trendData.labels.push(code);
                            stats.trendData.actual.push(colVol);
                            stats.trendData.trend.push(colHandled);
                        }
                    }
                }
            }
        }
    }

    // FINAL MATH
    stats.svl = svlVol > 0 ? Math.round(svlW / svlVol) + "%" : "0%";
    stats.ack = ackVol > 0 ? Math.floor(ackW / ackVol) + "s" : "0s";
    
    if (trendRef > 0) {
       const growth = (trendDiff / trendRef) * 100;
       stats.trendOut = (growth > 0 ? "+" : "") + growth.toFixed(2) + "%";
    }
  }

  // LOGGING
  try {
      if (typeof MasterConnector !== 'undefined' && stats.svl !== "0%") {
          MasterConnector.logStats(stats.svl, stats.ack, "", "");
      }
      if (typeof StatsTracker !== 'undefined' && stats.svl !== "0%") {
          StatsTracker.logHourlyStats(stats.svl, stats.ack);
      }
  } catch (e) {}

  stats.report = `STATS UPDATE:\nSVL OUT: ${stats.svl}\nSVL IN: ${stats.inSVL}\nACK: ${stats.ack}\nASA: ${stats.asa}\nSAFE: ${stats.safeSL}\n\nTRENDS:\nInbound: ${stats.trendIn}\nOutbound: ${stats.trendOut}\n\nDELAYS: None\n\nNOTES:\n %% Coachings Open%%`;
  return JSON.stringify(stats);
}

// HELPERS
function checkList(id, list) { return list.some(key => id.startsWith(key)); }
function dur(t) { 
    if (!t) return 0;
    const parts = t.split(":");
    if (parts.length === 3) return (parseInt(parts[0]) * 3600) + (parseInt(parts[1]) * 60) + parseInt(parts[2]);
    if (parts.length === 4) return (parseInt(parts[0]) * 86400) + (parseInt(parts[1]) * 3600) + (parseInt(parts[2]) * 60) + parseInt(parts[3]);
    return 0;
}
function fmt(t) { 
  if(!t) return "0s";
  if(t.includes(":")) {
      const p = t.split(":");
      const s = parseInt(p[p.length-1]);
      const m = parseInt(p[p.length-2] || 0);
      const h = parseInt(p[p.length-3] || 0);
      if(h>0) return `${h}h ${m}m`;
      if(m>0) return `${m}m ${s}s`;
      return `${s}s`;
  }
  return t;
}
