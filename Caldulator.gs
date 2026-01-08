/**
 * MODULE: CALCULATOR
 * Consumes raw text dumps to calculate Service Level (SVL) and Average Response Time (ACK).
 */

/* --- PUBLIC ENDPOINT --- */
function runCalculator(inText, outText) {
  return calculateMetrics(inText, outText);
}

/* --- CORE LOGIC --- */
function calculateMetrics(inText, outText) {
  let stats = {
    svl: "0%", ack: "0s", trendOut: "0%",
    asa: "0s", inSVL: "0%", trendIn: "0%",
    report: ""
  };

  // --- CONFIGURATION ---
  
  // 1. ACK CALCULATION LIST
  // STRICT List matching user requirement for ~6s average.
  // Excludes 5-SUPF (Supervision Fire) which often has 0s but high volume, or high seconds.
  const LIST_ACK = [
    "1-FIRE", "1-GAS", "1-H/U", "1-MED", "2-CCM", "2-FARM", 
    "3-LWK", "3-VID", "4-BURG", "4-TAMP", "6-O/C",
    "7-TRB", "5-SUPV" 
  ];

  // 2. SVL CALCULATION LIST
  const LIST_SVL = [
    "1-FIRE", "1-GAS", "1-H/U", "1-MED", "2-CCM", "2-FARM", 
    "3-LWK", "3-VID", "4-BURG", "4-TAMP", "6-O/C",
    "7-TRB", "5-SUPV" 
  ];

  // 3. TREND LIST
  const LIST_TREND = [
    "1-FIRE", "1-GAS", "1-H/U", "1-MED", "2-CCM", "2-FARM", 
    "3-LWK", "3-VID", "4-BURG", "4-TAMP", "6-O/C", 
    "5-SUPV", "7-TRB" 
  ];

  // --- 1. INBOUND PARSING ---
  if (inText) {
    const lines = inText.split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (line.match(/^\d+-[A-Z\/]+/)) continue; 
      
      // Look for "Monit - Intraday" or just "Intraday"
      if (line.includes("Intraday")) {
          // Extract ASA (e.g. 00:00:20)
          const timeMatch = line.match(/(\d{2}:\d{2}:\d{2})/);
          if (timeMatch) stats.asa = fmt(timeMatch[1]);
          
          // Extract SVL (85)
          const parts = line.trim().split(/\s+/);
          const pctIdx = parts.findIndex(p => p.includes("%"));
          if (pctIdx > -1 && parts[pctIdx+1]) {
              stats.inSVL = parts[pctIdx+1] + "%";
          }
      }
      
      if (line.includes("Last 60 minutes")) {
          const parts = line.split(/\s+/);
          const pctIdx = parts.findIndex(p => p.includes("%"));
          if (pctIdx > -1) {
              stats.trendIn = parts[pctIdx];
          }
      }
    }
  }

  // --- 2. OUTBOUND PARSING ---
  if (outText) {
    const secIntra = extractSection(outText, "Alarm Resp Time - Intraday");

    let svlVol=0, svlW=0;
    let ackVol=0, ackW=0;
    
    // Trend Variables
    let totalTrendRef = 0;
    let totalDiff = 0; 
    
    if (secIntra) {
      parseTable(secIntra, (id, vol, sl, sec, diff, ref) => {
        if (vol > 0) {
             // 1. ACK Calculation
             if (checkList(id, LIST_ACK)) {
                 ackVol += vol; 
                 ackW += (vol * sec);
             }

             // 2. SVL Calculation (Distinct List)
             if (checkList(id, LIST_SVL)) {
                 svlVol += vol;
                 svlW += (vol * sl);
             }

             // 3. Trend Calculation
             // Logic: Growth = (Sum of Diffs / Sum of Trend Reference) * 100
             if (checkList(id, LIST_TREND)) {
                 totalDiff += diff;
                 totalTrendRef += ref;
             }
        }
      });
    }

    // --- FINALIZE ---
    // SVL
    const outSVL = svlVol > 0 ? Math.round(svlW / svlVol) + "%" : "0%";
    
    // ACK
    const avgAck = ackVol > 0 ? Math.round(ackW / ackVol) + "s" : "0s";
    
    // Trend Outbound
    // Formula: (Total Diff / Total Trend Reference) * 100
    let trendOut = "0%";
    if (totalTrendRef > 0) {
      const growth = (totalDiff / totalTrendRef) * 100;
      trendOut = (growth > 0 ? "+" : "") + growth.toFixed(1) + "%";
    } else if (totalDiff > 0) {
      trendOut = "+100%";
    }

    stats.svl = outSVL;
    stats.ack = avgAck;
    stats.trendOut = trendOut;
  }

  // --- REPORT ---
  // Explicitly constructing with newlines to ensure correct formatting
  stats.report = "STATS UPDATE:\n" +
                 "SVL OUT: " + stats.svl + "\n" +
                 "SVL IN: " + stats.inSVL + "\n" +
                 "ACK: " + stats.ack + "\n" +
                 "ASA: " + stats.asa + "\n" +
                 "SAFE: 100%\n\n" +
                 "TRENDS:\n" +
                 "Inbound: " + stats.trendIn + "\n" +
                 "Outbound: " + stats.trendOut + "\n\n" +
                 "DELAYS: None\n\n" +
                 "NOTES:\n\n" +
                 " %% Coachings Open%%";
  
  // Optional logging if StatusTracker supports it
  if (typeof StatusTracker !== 'undefined' && typeof StatusTracker.logHourlyStats === 'function' && stats.svl !== "0%") {
      StatusTracker.logHourlyStats(stats.svl, stats.ack);
  }

  return JSON.stringify(stats);
}

// --- HELPERS ---
function extractSection(text, header) {
  const idx = text.indexOf(header);
  if (idx === -1) return "";
  const remainder = text.substring(idx + header.length);
  const nextSection = remainder.search(/(Alarm Resp Time|Pending Alarm|Logged-in Users)/);
  return nextSection === -1 ? remainder : remainder.substring(0, nextSection);
}

function parseTable(text, callback) {
  const lines = text.split(/\r?\n/);
  lines.forEach(line => {
    const match = line.match(/^(\d+-[A-Z\/]+)/);
    if (match) {
      const id = match[1];
      let cols = line.trim().split(/\t/);
      if (cols.length < 3) cols = line.trim().split(/\s{2,}/);
      
      // Flexible parsing for lines that might merge columns
      if (cols.length >= 4) {
        // Correct Column Mapping based on User Input:
        // Service Type | Description | Actual | Trend | Diff | ACK | SL
        // Index: 0     | 1           | 2      | 3     | 4    | 5   | 6
        
        const vol = parseInt(cols[2]) || 0; // Actual
        const ref = parseInt(cols[3]) || 0; // Trend (Reference)
        const diff = parseInt(cols[4]) || 0; // Diff
        
        // Find Time column (xx:xx:xx) usually near end
        const timeIdx = cols.findIndex(c => c.match(/^\d{2}:\d{2}:\d{2}$/));
        const sec = timeIdx > -1 ? dur(cols[timeIdx]) : 0;
        
        // Find SL column (last one usually, 0-100)
        const slIdx = cols.length - 1;
        const sl = parseFloat(cols[slIdx]) || 0;
        
        callback(id, vol, sl, sec, diff, ref);
      } 
    }
  });
}

function checkList(id, list) {
  return list.some(key => id.startsWith(key));
}

function fmt(t) { 
  if(!t || !t.includes(":")) return t || "0s"; 
  const p=t.split(":"); 
  const h=parseInt(p[0]);
  const m=parseInt(p[1]);
  const s=parseInt(p[2]);
  
  if (h>0) return `${h}h ${m}m`;
  if (m>0) return `${m}m ${s}s`;
  return `${s}s`;
}

function dur(t) { 
  if(!t) return 0; 
  const p=t.split(":"); 
  if (p.length !== 3) return 0;
  return (parseInt(p[0]||0)*3600)+(parseInt(p[1]||0)*60)+parseInt(p[2]||0); 
}
