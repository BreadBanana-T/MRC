/**
 * MODULE: CALCULATOR (LOCAL HOST ONLY)
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

  const LIST_SVL_ACK = [
      "1-FIRE", "1-GAS", "1-H/U", "1-MED", 
      "2-CCM", "2-FARM", 
      "3-LWK", "3-VID", 
      "4-BURG", "4-TAMP"
  ];

  const LIST_TREND = [
      "1-FIRE", "1-GAS", "1-H/U", "1-MED", 
      "2-CCM", "2-FARM", 
      "3-LWK", "3-VID", 
      "4-BURG", "4-COMM", "4-TAMP", 
      "6-O/C"
  ];

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

  if (outText) {
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

                if (vol > 0 && checkList(code, LIST_SVL_ACK)) {
                     ackVol += vol; ackW += (vol * timeSec);
                     svlVol += vol; svlW += (vol * slVal);
                }
            }
         }
      });
    }

    stats.svl = svlVol > 0 ? Math.round(svlW / svlVol) + "%" : "0%";
    stats.ack = ackVol > 0 ? Math.round(ackW / ackVol) + "s" : "0s";

    const trend60 = extractSection(outText, "Alarm Resp Time - Last 60 min");
    let trendDiff = 0, trendRef = 0;

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
                   const diff = parseInt(parts[timeIdx - 1]) || 0; 

                   if (offered > 0 || handled > 0) {
                      stats.trendData.labels.push(code);
                      stats.trendData.actual.push(offered);
                      stats.trendData.trend.push(handled);
                   }

                   if (checkList(code, LIST_TREND)) {
                      trendDiff += diff;
                      trendRef += handled;
                   }
                }
            }
        });
    }

    if (trendRef > 0) {
       const growth = (trendDiff / trendRef) * 100;
       stats.trendOut = formatTrend(growth);
    }
  }

  // LOGGING (Only Local Now)
  if (typeof StatsTracker !== 'undefined' && stats.svl !== "0%") {
      StatsTracker.logHourlyStats(stats.svl, stats.ack);
  }

  stats.report = `STATS UPDATE:\nSVL OUT: ${stats.svl}\nSVL IN: ${stats.inSVL}\nACK: ${stats.ack}\nASA: ${stats.asa}\nSAFE: N/A\n\nTRENDS:\nInbound: ${stats.trendIn}\nOutbound: ${stats.trendOut}\n\nDELAYS: None\n\nNOTES:\n %% Coachings Open%%`;
  return JSON.stringify(stats);
}

function extractSection(text, header) {
  const idx = text.indexOf(header);
  if (idx === -1) return "";
  const remainder = text.substring(idx + header.length);
  const nextIdx = remainder.search(/(Alarm Resp Time|Pending Alarm|Logged-in Users|Potential Runaway|IVR Not Started)/);
  return nextIdx === -1 ? remainder : remainder.substring(0, nextIdx);
}

function checkList(id, list) { 
  return list.some(key => id.startsWith(key));
}

function formatTrend(val) {
  if (val === 0) return "0%";
  const isNeg = val < 0;
  const absVal = Math.abs(val);
  const intPart = Math.floor(absVal);
  const tenths = Math.floor((absVal * 10) % 10);
  
  let finalString;
  if (tenths >= 5) {
     finalString = intPart + "." + tenths; 
  } else {
     finalString = intPart;
  }
  return (isNeg ? "-" : "+") + finalString + "%";
}

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
