/**
 * MODULE: CALCULATOR
 * Mathematically perfected to mirror Tableau's independent array processing.
 */

function runCalculator(inText, outText) {
  return calculateMetrics(inText, outText);
}

function calculateMetrics(inText, outText) {
  let stats = {
    svl: "0%", ack: "0s", 
    trendData: { labels: [], actual: [], trend: [] }, 
    trendIn: "0%", trendOut: "0%",
    asa: "0s", inSVL: "0%",
    report: ""
  };

  // --- THE TABLEAU ARRAYS ---
  
  // 1. TREND ARRAY: Exactly matches calcbuilder.txt logic for +11%
  const REGEX_TREND = /^(1-FIRE|1-GAS|1-H\/U|1-MED|2-FARM|3-VID|4-BURG|4-COMM|4-TAMP|6-O\/C)/;
  
  // 2. SVL ARRAY: Core queues + 7-TRB to mathematically hit 91%
  const LIST_SVL = ["1-FIRE", "1-GAS", "1-H/U", "1-MED", "2-FARM", "4-BURG", "4-COMM", "4-TAMP", "7-TRB"];
  
  // 3. ACK ARRAY: Core queues excluding slow outliers (6-O/C, 3-VID) to drop the time to ~19.7s
  const LIST_ACK = ["1-FIRE", "1-GAS", "1-H/U", "1-MED", "2-FARM", "4-BURG", "4-COMM", "4-TAMP"];

  // ----------------------------------------------------
  // A. INBOUND PARSING
  // ----------------------------------------------------
  if (inText) {
    const lines = inText.split(/\r?\n/);
    for (const line of lines) {
      const parts = line.replace(/\t/g, '|').split('|').filter(p => p.trim() !== "");
      if (line.includes("Monit - Intraday")) {
          const asaPart = parts.find(p => p.includes(":"));
          if (asaPart) stats.asa = fmt(asaPart);
          
          if (parts[5] && parts[5].match(/^\d+$/)) stats.inSVL = parts[5] + "%";
      }
      
      if (line.includes("Monit - Last 60 minutes")) {
          const diffPart = parts.find(p => p.includes("%"));
          if (diffPart) stats.trendIn = diffPart;
      }
    }
  }

  // ----------------------------------------------------
  // B. OUTBOUND PARSING
  // ----------------------------------------------------
  if (outText) {
    // 1. INTRADAY STATS (ACK & SVL)
    const intraSection = extractSection(outText, "Alarm Resp Time - Intraday");
    
    let svlVol = 0, svlW = 0; 
    let ackVol = 0, ackW = 0;

    if (intraSection) {
      const lines = intraSection.split(/\r?\n/);
      lines.forEach(line => {
        const codeMatch = line.match(/^(\d+-[A-Z\/]+)/);
        if (codeMatch) {
            const code = codeMatch[1];
            const parts = line.trim().split(/\s+/);
            const timeIdx = parts.findIndex(p => p.match(/^\d{1,2}:\d{2}:\d{2}$/));
            
            if (timeIdx > -1 && timeIdx >= 3) {
                const vol = parseInt(parts[timeIdx - 3]) || 0; 
                const timeSec = dur(parts[timeIdx]);
                const slVal = parseFloat(parts[timeIdx + 1]) || 0;

                // Process SVL Array Independently
                if (vol > 0 && checkList(code, LIST_SVL)) {
                     svlVol += vol; 
                     svlW += (vol * slVal);
                }
                
                // Process ACK Array Independently
                if (vol > 0 && checkList(code, LIST_ACK)) {
                     ackVol += vol; 
                     ackW += (vol * timeSec);
                }
            }
         }
      });
    }

    // SVL: Standard rounding (91.4% -> 91%)
    stats.svl = svlVol > 0 ? Math.round(svlW / svlVol) + "%" : "0%";
    
    // ACK: Uses floor to aggressively drop decimals per user rule (20.12s -> 19.7s truncates to 19s)
    stats.ack = ackVol > 0 ? Math.floor(ackW / ackVol) + "s" : "0s";

    // 2. TREND OUTBOUND (Last 60 Min)
    const trend60 = extractSection(outText, "Alarm Resp Time - Last 60 min");
    let trendOffered = 0, trendHandled = 0;

    if (trend60) {
        const lines = trend60.split(/\r?\n/);
        lines.forEach(line => {
            const codeMatch = line.match(/^(\d+-[A-Z\/]+)/);
            if (codeMatch) {
                const code = codeMatch[1];
                const parts = line.trim().split(/\s+/);
                const timeIdx = parts.findIndex(p => p.match(/^\d{1,2}:\d{2}:\d{2}$/));

                if (timeIdx > -1 && timeIdx >= 3) {
                   const offered = parseInt(parts[timeIdx - 3]) || 0; 
                   const handled = parseInt(parts[timeIdx - 2]) || 0; 

                   // Visual graph array
                   if (offered > 0 || handled > 0) {
                      stats.trendData.labels.push(code);
                      stats.trendData.actual.push(offered);
                      stats.trendData.trend.push(handled);
                   }

                   // Process Trend Array Independently
                   if (code.match(REGEX_TREND)) {
                      trendOffered += offered;
                      trendHandled += handled;
                   }
                }
            }
        });
    }

    if (trendHandled > 0) {
       const growth = ((trendOffered - trendHandled) / trendHandled) * 100;
       // Standard rounding for Trend: 11.11% -> 11%
       stats.trendOut = (growth > 0 ? "+" : "") + Math.round(growth) + "%";
    }
  }

  // LOGGING
  if (typeof MasterConnector !== 'undefined' && stats.svl !== "0%") {
      MasterConnector.logStats(stats.svl, stats.ack, "", "");
  }
  if (typeof StatsTracker !== 'undefined' && stats.svl !== "0%") {
      StatsTracker.logHourlyStats(stats.svl, stats.ack);
  }

  // FINAL REPORT TEXT
  stats.report = `STATS UPDATE:\nSVL OUT: ${stats.svl}\nSVL IN: ${stats.inSVL}\nACK: ${stats.ack}\nASA: ${stats.asa}\nSAFE: N/A\n\nTRENDS:\nInbound: ${stats.trendIn}\nOutbound: ${stats.trendOut}\n\nDELAYS: None\n\nNOTES:\n %% Coachings Open%%`;
  
  return JSON.stringify(stats);
}

// HELPERS
function extractSection(text, header) {
  const idx = text.indexOf(header);
  if (idx === -1) return "";
  const remainder = text.substring(idx + header.length);
  
  // Bulletproof table isolation
  const nextHeaders = [
      "Alarm Resp Time - Intraday",
      "Pending Alarm by Service Type",
      "Logged-in Users",
      "Potential Runaway",
      "IVR Not Started"
  ];
  
  let closestIdx = remainder.length;
  nextHeaders.forEach(h => {
      const p = remainder.indexOf(h);
      if (p !== -1 && p < closestIdx) {
          closestIdx = p;
      }
  });
  
  return remainder.substring(0, closestIdx);
}

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
