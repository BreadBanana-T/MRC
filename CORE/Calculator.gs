/**
 * MODULE: CALCULATOR
 * Parses Metrics, Trends, and sends data to Master DB.
 */

function runCalculator(inText, outText) {
  return calculateMetrics(inText, outText);
}

function calculateMetrics(inText, outText) {
  let stats = {
    svl: "0%", ack: "0s", 
    trendData: { labels: [], actual: [], trend: [] }, 
    idp: "0", law: "0s",
    trendIn: "0%", trendOut: "0%",
    report: ""
  };

  const LIST_ACK = ["1-FIRE", "1-GAS", "1-H/U", "1-MED", "2-CCM", "2-FARM", "3-LWK", "3-VID", "4-BURG", "4-TAMP", "6-O/C", "7-TRB", "5-SUPV"];
  const LIST_SVL = ["1-FIRE", "1-GAS", "1-H/U", "1-MED", "2-CCM", "2-FARM", "3-LWK", "3-VID", "4-BURG", "4-TAMP", "6-O/C", "7-TRB", "5-SUPV"];
  
  // 1. INBOUND PARSING
  if (inText) {
    const lines = inText.split(/\r?\n/);
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].includes("Intraday")) {
          const parts = lines[i].trim().split(/\s+/);
          const pctIdx = parts.findIndex(p => p.includes("%"));
          if (pctIdx > -1 && parts[pctIdx+1]) stats.inSVL = parts[pctIdx+1] + "%";
      }
      if (lines[i].includes("Last 60 minutes")) {
          const parts = lines[i].split(/\s+/);
          const pctIdx = parts.findIndex(p => p.includes("%"));
          if (pctIdx > -1) stats.trendIn = parts[pctIdx];
      }
    }
  }

  // 2. OUTBOUND PARSING
  if (outText) {
    // A. LAW
    const lawMatch = outText.match(/Longest.*?(\d+s?)/i);
    if (lawMatch) stats.law = lawMatch[1]; 

    // B. IDP
    const idpMatch = outText.match(/IDP.*?(\d+)/i);
    if (idpMatch) stats.idp = idpMatch[1];

    // C. TREND GRAPH (Last 60 min)
    const trendSection = extractSection(outText, "Alarm Resp Time - Last 60 min");
    if (trendSection) {
       const lines = trendSection.split(/\r?\n/);
       lines.forEach(line => {
           // Heuristic: Look for lines starting with digit-Letter (e.g. 1-FIRE)
           const parts = line.trim().split(/\t/);
           // If tab split fails, try space split but be careful of descriptions with spaces
           const cleanLine = line.trim();
           if (cleanLine.match(/^\d-[A-Z]+/)) {
               // Simple space split might break description, but usually Actual/Trend are at the end
               // Let's assume columns: Code | Desc | Actual | Trend | Diff | ACK | SL
               // We need Actual (col 3) and Trend (col 4)
               // Regex to find the numbers
               const nums = cleanLine.match(/(\d+)\s+(\d+)\s+(-?\d+)\s+\d{2}:\d{2}:\d{2}/);
               if (nums) {
                   const code = cleanLine.split(" ")[0];
                   stats.trendData.labels.push(code);
                   stats.trendData.actual.push(parseInt(nums[1]));
                   stats.trendData.trend.push(parseInt(nums[2]));
               }
           }
       });
    }

    // D. GLOBAL SVL & ACK (Intraday)
    const secIntra = extractSection(outText, "Alarm Resp Time - Intraday");
    let svlVol=0, svlW=0, ackVol=0, ackW=0, totalDiff=0, totalRef=0;

    if (secIntra) {
      parseTable(secIntra, (id, vol, sl, sec, diff, ref) => {
        if (vol > 0) {
             if (checkList(id, LIST_ACK)) { ackVol += vol; ackW += (vol * sec); }
             if (checkList(id, LIST_SVL)) { svlVol += vol; svlW += (vol * sl); }
             totalDiff += diff; 
             totalRef += ref;
        }
      });
    }

    stats.svl = svlVol > 0 ? Math.round(svlW / svlVol) + "%" : "0%";
    stats.ack = ackVol > 0 ? Math.round(ackW / ackVol) + "s" : "0s";
    
    if (totalRef > 0) {
        const growth = (totalDiff / totalRef) * 100;
        stats.trendOut = (growth > 0 ? "+" : "") + growth.toFixed(1) + "%";
    }
  }

  // --- LOGGING ---
  if (typeof MasterConnector !== 'undefined' && stats.svl !== "0%") {
      MasterConnector.logStats(stats.svl, stats.ack, stats.law, stats.idp);
  }
  
  if (typeof StatsTracker !== 'undefined' && stats.svl !== "0%") {
      StatsTracker.logHourlyStats(stats.svl, stats.ack);
  }

  stats.report = `STATS UPDATE:\nSVL OUT: ${stats.svl}\nSVL IN: ${stats.inSVL}\nACK: ${stats.ack}\nLAW: ${stats.law}\nIDP: ${stats.idp}\n\nTRENDS:\nInbound: ${stats.trendIn}\nOutbound: ${stats.trendOut}\n\nNOTES:\n %% Coachings Open%%`;
  return JSON.stringify(stats);
}

function extractSection(text, header) {
  const idx = text.indexOf(header);
  if (idx === -1) return "";
  const remainder = text.substring(idx + header.length);
  const nextSection = remainder.search(/(Alarm Resp Time|Pending Alarm|Logged-in Users|« Prev)/);
  return nextSection === -1 ? remainder : remainder.substring(0, nextSection);
}

function parseTable(text, callback) {
  const lines = text.split(/\r?\n/);
  lines.forEach(line => {
    const match = line.match(/^(\d+-[A-Z\/]+)/);
    if (match) {
      let cols = line.trim().split(/\t/);
      if (cols.length < 3) cols = line.trim().split(/\s{2,}/);
      if (cols.length >= 4) {
        const vol = parseInt(cols[2]) || 0;
        const ref = parseInt(cols[3]) || 0; 
        const diff = parseInt(cols[4]) || 0; 
        const timeIdx = cols.findIndex(c => c.match(/^\d{2}:\d{2}:\d{2}$/));
        const sec = timeIdx > -1 ? dur(cols[timeIdx]) : 0;
        const slIdx = cols.length - 1;
        const sl = parseFloat(cols[slIdx]) || 0;
        callback(match[1], vol, sl, sec, diff, ref);
      } 
    }
  });
}

function checkList(id, list) { return list.some(key => id.startsWith(key)); }
function dur(t) { if(!t) return 0; const p=t.split(":"); return (parseInt(p[0]||0)*3600)+(parseInt(p[1]||0)*60)+parseInt(p[2]||0); }
