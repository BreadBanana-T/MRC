/**
 * MODULE: CALCULATOR (LOCAL HOST ONLY)
 */

// Router export runCalculator() lives in Code.gs.

// The report layout is editable at runtime (no redeploy) via the Calculator
// view — it's stored in a Script Property and read through CalcSettings
// (CORE/CalcSettings.gs). This constant is the built-in fallback used when no
// custom template has been saved. Placeholders {…} are filled from the
// computed stats by buildCalcReport(); any literal text (labels, DELAYS,
// NOTES, "(Genesys)" suffixes, etc.) is edited straight in the template.
const DEFAULT_CALC_TEMPLATE =
`STATS UPDATE:
SVL OUT: {svlOut}
SVL IN: {svlIn}
ACK: {ack}
ASA: {asa}
LCW Interval: {lcw}
SAFE: {safe}

TRENDS:
Inbound: {trendIn}
Outbound: {trendOut}

DELAYS: None

NOTES:
%% Coaching Open %%`;

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
    // a Current interval column. The team reports ASA and SVL IN from Today,
    // LCW from the Current interval. SAFE lives on a separate tableau and the
    // Genesys paste has no inbound trend, so both stay "N/A" for manual fill.
    // A time value can be multi-part with spaces ("2m 29s", "1 min 5 sec"), a
    // single unit ("10s", "2m"), or a clock value ("1:05"). SL shows as "61%".
    // These match a whole value (one line / one tab-cell), never a fragment.
    const TIME_RE = /^(?:\d{1,2}:\d{2}(?::\d{2})?|(?:\d+(?:\.\d+)?\s*(?:hours?|hrs?|h|minutes?|mins?|m|seconds?|secs?|s)\b\s*)+)$/i;
    const PCT_RE = /^\d+(?:\.\d+)?\s*%$/;

    const asaC = genesysCells(inText, "ASA", TIME_RE);
    const slC = genesysCells(inText, "SL", PCT_RE);
    const lcwC = genesysCells(inText, "LCW", TIME_RE);

    // When "Current interval" has no data yet (start of an interval, quiet
    // period) Genesys leaves that cell blank, and a blank cell vanishes from
    // a one-cell-per-line copy — so the duplicated pair "10s / 13s / 10s /
    // 13s" collapses to "10s / 10s". Reading that positionally used to put
    // TODAY's value in the interval slot (LCW showed the day's max wait).
    // Detect it: the copy is known to be duplicated if any widget shows 4+
    // cells, or if every pasted widget collapsed to the same [x, x] pattern.
    const found = [asaC, slC, lcwC].filter(Boolean);
    const isPairDup = c => !c.hadGap && c.cells.length === 2 && c.cells[0] === c.cells[1];
    const dupCopy = found.some(c => c.cells.length >= 4) ||
                    (found.length > 1 && found.every(isPairDup));
    const pair = c => {
      if (!c || !c.cells.length) return null;
      if (dupCopy && isPairDup(c)) return { today: c.cells[0], interval: null };
      return { today: c.cells[0], interval: c.cells.length > 1 ? c.cells[1] : null };
    };

    const asaP = pair(asaC), slP = pair(slC), lcwP = pair(lcwC);
    if (asaP && asaP.today) stats.asa = asaP.today;
    if (slP && slP.today) stats.inSVL = slP.today;
    // No current-interval value -> leave LCW "N/A" for manual fill rather
    // than reporting a wrong number.
    if (lcwP && lcwP.interval) stats.lcw = lcwP.interval;
  }

  if (outText) {
    // Time of the snapshot, for the automated Point-of-No-Return check.
    const updMatch = outText.match(/Last update:\s*\d{1,2}-\d{1,2}-\d{4}\s+(\d{1,2}):(\d{2})/);
    if (updMatch) stats.updateHour = Math.round((parseInt(updMatch[1]) + parseInt(updMatch[2]) / 60) * 100) / 100;

    // Only the first table under each header is read, and only the first row
    // per service type (see alarmRows) — so a later table that also lists
    // 4-BURG (e.g. a per-alarm-queue breakdown) can't overwrite the ACK or be
    // double-counted into SVL.
    const intraRows = alarmRows(extractSection(outText, "Alarm Resp Time - Intraday"));
    let svlVol = 0, svlW = 0;

    intraRows.forEach(r => {
      if (r.actual > 0 && checkList(r.code, LIST_SVL)) {
        svlVol += r.actual; svlW += (r.actual * r.sl);
      }
      // ACK = the 4-BURG Intraday response time, full stop. The old
      // weighted average let high-volume slow queues (5-SUPV, 7-TRB)
      // poison the number (81s reported vs 24s real). This matches
      // both the reported stats and the Point-of-No-Return playbook
      // ("Take the ACK" → the BURG row).
      if (r.code === "4-BURG") {
        stats.ack = r.ackSec + "s";
        stats.burgActual = r.actual;
        stats.burgAckSec = r.ackSec;
      }
    });
    // No BURG row in the paste: say so instead of reporting a fake "0s".
    if (!intraRows.some(r => r.code === "4-BURG")) stats.ack = "N/A";

    stats.svl = svlVol > 0 ? Math.round(svlW / svlVol) + "%" : "0%";

    let trendDiff = 0, trendRef = 0;
    alarmRows(extractSection(outText, "Alarm Resp Time - Last 60 min")).forEach(r => {
      // Last-60 columns: Actual (offered), Trend (handled ref), Diff.
      if (r.actual > 0 || r.trend > 0) {
        stats.trendData.labels.push(r.code);
        stats.trendData.actual.push(r.actual);
        stats.trendData.trend.push(r.trend);
      }
      if (checkList(r.code, LIST_TREND)) {
        trendDiff += r.diff;
        trendRef += r.trend;
      }
    });

    if (trendRef > 0) {
       const growth = (trendDiff / trendRef) * 100;
       stats.trendOut = formatTrend(growth);
    }
  }

  // LOGGING (Only Local Now)
  if (typeof StatsTracker !== 'undefined' && stats.svl !== "0%" && stats.ack !== "N/A") {
      StatsTracker.logHourlyStats(stats.svl, stats.ack);
  }

  // Render the report from the (possibly customized) template. SAFE and the
  // inbound trend default to "N/A" for the sender to fill in manually.
  stats.report = buildCalcReport(stats);
  return JSON.stringify(stats);
}

// Fill the report template's {tokens} from the computed stats. Unknown tokens
// are left untouched so a typo in the template is visible rather than silently
// dropping text. Uses CalcSettings.template() when available, else the
// built-in DEFAULT_CALC_TEMPLATE.
function buildCalcReport(stats) {
  const tpl = (typeof CalcSettings !== 'undefined') ? CalcSettings.template() : DEFAULT_CALC_TEMPLATE;
  const map = {
    svlOut: stats.svl, svlIn: stats.inSVL, ack: stats.ack, asa: stats.asa,
    lcw: stats.lcw, safe: stats.safe, trendIn: stats.trendIn, trendOut: stats.trendOut
  };
  return tpl.replace(/\{(\w+)\}/g, (m, key) => (map[key] !== undefined ? map[key] : m));
}

// Genesys widget paste: widget title line ("External ASA"…), then the
// "Today" / "Current interval" column labels and the metric name, then the
// values in column order: Today, Current interval (the copy usually repeats
// the pair). Returns the value cells in order, with null for a cell that is
// present but empty ("-", "--", "N/A", or an empty tab-cell), so a missing
// Current interval can never shift Today's value into its slot. hadGap says
// whether any such empty cell was seen. The scan is cut at the next
// "External …" title so widgets can be pasted in any order.
function genesysCells(text, metric, valueRe) {
  const m = new RegExp("External\\s+" + metric + "\\b", "i").exec(text);
  if (!m) return null;
  let block = text.substring(m.index + m[0].length);
  const next = block.search(/External\s+\w/i);
  if (next !== -1) block = block.substring(0, next);

  const EMPTY_RE = /^(?:-+|\u2013|\u2014|n\/?a|null|no data)?$/i;
  const cells = [];
  let hadGap = false;
  block.split(/\r?\n/).forEach(line => {
    // Tab-separated rows ("ASA<TAB>10s<TAB>") keep empty cells positionally;
    // a bare blank line (one-cell-per-line copy) carries no position, skip it.
    // Split on tabs/newlines only — NOT spaces — so "2m 29s" stays one value.
    let parts = line.split("\t").map(x => x.trim());
    if (parts.length === 1 && parts[0] === "") return;
    // Drop leading labels/blank cells ("ASA", "Service Level %", "Today"…).
    while (parts.length && !valueRe.test(parts[0]) && (parts[0] === "" || !EMPTY_RE.test(parts[0]))) parts.shift();
    // A label-only row ("Today<TAB>Current interval") has nothing left.
    parts.forEach(x => {
      if (valueRe.test(x)) cells.push(x.replace(/\s+/g, " "));
      else if (EMPTY_RE.test(x)) { cells.push(null); hadGap = true; }
    });
  });
  return { cells: cells, hadGap: hadGap };
}

// Rows of the first "Service Type | Description | Actual | Trend | Diff | ACK
// | SL" table in a section. Columns are anchored on the ACK time cell
// (hh:mm:ss) because the Description has spaces. Stops at the first line that
// isn't a row, a pager ("« Prev | 1 | 2 …"), a column header or blank once the
// table has started, and keeps only the first row per service type.
function alarmRows(section) {
  const rows = [], seen = {};
  if (!section) return rows;
  let started = false;
  const lines = section.split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    const codeMatch = line.match(/^(\d+-[A-Z\/]+)\s/);
    if (!codeMatch) {
      if (started && line !== "" && !/Prev\s*\|/.test(line) && !/^Service Type\b/.test(line)) break;
      continue;
    }
    started = true;
    const code = codeMatch[1];
    if (seen[code]) continue;
    const parts = line.split(/\s+/);
    const timeIdx = parts.findIndex(p => /^\d{1,2}:\d{2}:\d{2}$/.test(p));
    if (timeIdx < 3) continue;
    seen[code] = true;
    rows.push({
      code: code,
      actual: parseInt(parts[timeIdx - 3]) || 0,
      trend: parseInt(parts[timeIdx - 2]) || 0,
      diff: parseInt(parts[timeIdx - 1]) || 0,
      ackSec: dur(parts[timeIdx]),
      sl: parseFloat(parts[timeIdx + 1]) || 0
    });
  }
  return rows;
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
