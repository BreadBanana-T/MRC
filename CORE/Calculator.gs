/**
 * MODULE: CALCULATOR (LOCAL HOST ONLY)
 */

// Router export runCalculator() lives in Code.gs.

function calculateMetrics(inText, outText) {
  let stats = {
    svl: "0%", ack: "0s", safe: "N/A",
    trendData: { labels: [], actual: [], trend: [] },
    trendIn: "N/A", trendOut: "0%",
    asa: "N/A", inSVL: "N/A", lcw: "N/A",
    burgActual: 0, burgAckSec: 0, updateHour: null,
    report: ""
  };

  // SVL OUT = Actual-weighted SL over the EXACT service types on the MAS
  // heatmap the reported SVL/ACK are read from:
  //   1-FIRE 1-GAS 1-H/U 1-MED 3-VID 4-BURG 4-TAMP
  // NOT 4-COMM / 2-CCM / 2-FARM / 3-LWK — those are not on that heatmap, and
  // an earlier note here wrongly added 4-COMM (dropping it changes nothing:
  // its SL sits on the blend). Verified across 5 paired captures.
  // NOTE: the heatmap measures these on a different basis than this
  // bureau-wide alarm export — its 4-BURG cell can sit ~2 pts above the raw's
  // 4-BURG in ANY window — so this from-raw figure tracks the reported SVL to
  // ~1-3 pts but is not digit-exact. That residual is a data-source
  // difference, not a queue-list bug; it cannot be closed from this raw.
  const LIST_SVL = [
      "1-FIRE", "1-GAS", "1-H/U", "1-MED",
      "3-VID",
      "4-BURG", "4-TAMP"
  ];

  // Trend % uses only the "priority" queues the tableau tracks for trending.
  // Excludes 2-CCM (not a trend queue) and 3-LWK (Lone Worker, tracked
  // separately).
  const LIST_TREND = [
      "1-FIRE", "1-GAS", "1-H/U", "1-MED",
      "2-FARM",
      "3-VID",
      "4-BURG", "4-COMM", "4-TAMP",
      "6-O/C"
  ];

  if (inText) {
    // Inbound now comes from three Genesys Cloud widgets pasted as one block:
    // External ASA / External SL / External LCW, each with a Today column and
    // a Current interval column. The team reports ASA from Today, SVL IN and
    // LCW from the Current interval. SAFE lives on a separate tableau and the
    // Genesys paste has no inbound trend, so both stay "N/A" for manual fill.
    // Times show as "10s" (or "2m", "1:05"); SL shows as "61%".
    const TIME_RE = /^(\d+(\.\d+)?(s|m|min|h)|\d{1,2}:\d{2}(:\d{2})?)$/i;
    const PCT_RE = /^\d+(\.\d+)?%$/;

    const asaP = genesysPair(inText, "External ASA", TIME_RE);
    if (asaP) stats.asa = asaP.today;

    const slP = genesysPair(inText, "External SL", PCT_RE);
    if (slP) stats.inSVL = slP.interval;

    const lcwP = genesysPair(inText, "External LCW", TIME_RE);
    if (lcwP) stats.lcw = lcwP.interval;
  }

  if (outText) {
    // Time of the snapshot, for the automated Point-of-No-Return check.
    const updMatch = outText.match(/Last update:\s*\d{1,2}-\d{1,2}-\d{4}\s+(\d{1,2}):(\d{2})/);
    if (updMatch) stats.updateHour = Math.round((parseInt(updMatch[1]) + parseInt(updMatch[2]) / 60) * 100) / 100;

    const intraSection = extractSection(outText, "Alarm Resp Time - Intraday");
    let svlVol = 0, svlW = 0;

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

                if (vol > 0 && checkList(code, LIST_SVL)) {
                     svlVol += vol; svlW += (vol * slVal);
                }
                // ACK = the 4-BURG Intraday response time, full stop. The old
                // weighted average let high-volume slow queues (5-SUPV, 7-TRB)
                // poison the number (81s reported vs 24s real). This matches
                // both the reported stats and the Point-of-No-Return playbook
                // ("Take the ACK" → the BURG row).
                if (code === "4-BURG") {
                     stats.ack = timeSec + "s";
                     stats.burgActual = vol;
                     stats.burgAckSec = timeSec;
                }
            }
         }
      });
    }

    stats.svl = svlVol > 0 ? Math.round(svlW / svlVol) + "%" : "0%";

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

  // SAFE is read off a separate tableau and Inbound trend isn't available
  // from Genesys yet — both print as N/A for the sender to fill in.
  stats.report = `STATS UPDATE:\nSVL OUT: ${stats.svl}\nSVL IN: ${stats.inSVL} (Genesys)\nACK: ${stats.ack}\nASA: ${stats.asa} (Genesys)\nLCW Interval: ${stats.lcw}\nSAFE: N/A\n\nTRENDS:\nInbound: ${stats.trendIn}\nOutbound: ${stats.trendOut}\n\nDELAYS: None\n\nNOTES:\n%% Coaching Closed %%`;
  return JSON.stringify(stats);
}

// Genesys widget paste: widget title line ("External ASA"…), then the
// "Today" / "Current interval" column labels and the metric name, then the
// values — first value-shaped token = Today, second = Current interval (the
// copy duplicates the pair; only the first two count). The scan is cut at
// the next "External " title so widgets can be pasted in any order.
function genesysPair(text, header, valueRe) {
  const idx = text.indexOf(header);
  if (idx === -1) return null;
  let block = text.substring(idx + header.length);
  const nextIdx = block.indexOf("External ");
  if (nextIdx !== -1) block = block.substring(0, nextIdx);
  const tokens = block.split(/\s+/).filter(t => valueRe.test(t));
  return tokens.length >= 2 ? { today: tokens[0], interval: tokens[1] } : null;
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

  // User rule: a trend is shown with 1-decimal precision when the tenths
  // digit is >= 5 — otherwise integer only. Decision is made on the
  // UNROUNDED tenths so 1.4999 stays "1%", but the displayed number is
  // rounded so 1.7964 shows as "1.8%" (matching the tableau), not "1.7%".
  const origTenths = Math.floor((absVal * 10) % 10);
  if (origTenths < 5) {
    return (isNeg ? "-" : "+") + Math.floor(absVal) + "%";
  }
  const rounded = Math.round(absVal * 10) / 10;
  return (isNeg ? "-" : "+") + rounded.toFixed(1) + "%";
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
