/**
 * ============================================================
 * 17. BID ANALYSIS (S025)
 * Reads Proposal_Generator and produces two reports:
 *   1. Competition level breakdown (by Bid_1st proposal count)
 *      with send rates and flagged "sent into extreme competition" rows.
 *   2. Boost pattern report -- flags excessive boosts and
 *      boosts applied to already-skipped jobs (wasted planning).
 *
 * Thresholds (based on S022-S033 data, adjustable here):
 *   EXTREME_THRESHOLD  -- Bid_1st >= this -> extreme competition
 *   HIGH_THRESHOLD     -- Bid_1st >= this -> high competition
 *   MEDIUM_THRESHOLD   -- Bid_1st >= this -> medium competition
 *   BOOST_HEAVY_RATIO  -- Boost/Connects_Required >= this -> heavy
 *   BOOST_EXCESSIVE_RATIO -- Boost/Connects_Required >= this -> excessive
 * ============================================================
 */
var EXTREME_THRESHOLD_      = 50;
var HIGH_THRESHOLD_         = 30;
var MEDIUM_THRESHOLD_       = 10;
var BOOST_HEAVY_RATIO_      = 0.75;
var BOOST_EXCESSIVE_RATIO_  = 1.0;

function ANALYZE_BID_PATTERNS() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName("Proposal_Generator");

  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert("No data found in Proposal_Generator.");
    return;
  }

  var map        = getHeaderMap_(sheet);
  var titleCol   = getCol_(map, ["Job_Title"]);
  var bid1Col    = getCol_(map, ["Bid_1st"]);
  var bid2Col    = getCol_(map, ["Bid_2nd"]);
  var bid3Col    = getCol_(map, ["Bid_3rd"]);
  var boostCol   = getCol_(map, ["Boost_Connects"]);
  var connCol    = getCol_(map, ["Connects_Required"]);
  var totalCol   = getCol_(map, ["Total_Connects_Spent"]);
  var statusCol  = getCol_(map, ["Proposal_Status"]);

  if (!bid1Col || !statusCol) {
    ui.alert("Required columns not found. Confirm Bid_1st and Proposal_Status exist.");
    return;
  }

  var lastRow = sheet.getLastRow();
  var data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  // buckets: { count, sent, skip, other }
  var buckets = {
    low:     { label: "Low     (0-9)",   count: 0, sent: 0, skip: 0, other: 0 },
    medium:  { label: "Medium  (10-29)", count: 0, sent: 0, skip: 0, other: 0 },
    high:    { label: "High    (30-49)", count: 0, sent: 0, skip: 0, other: 0 },
    extreme: { label: "Extreme (50+)",   count: 0, sent: 0, skip: 0, other: 0 }
  };

  var bidRowCount      = 0;
  var extremeSentFlags = [];  // rows that were sent despite extreme competition
  var boostRows        = [];  // rows with Boost_Connects > 0

  for (var i = 0; i < data.length; i++) {
    var b1     = bid1Col   ? Number(data[i][bid1Col   - 1]) || 0 : 0;
    var b2     = bid2Col   ? Number(data[i][bid2Col   - 1]) || 0 : 0;
    var b3     = bid3Col   ? Number(data[i][bid3Col   - 1]) || 0 : 0;
    var boost  = boostCol  ? Number(data[i][boostCol  - 1]) || 0 : 0;
    var conn   = connCol   ? Number(data[i][connCol   - 1]) || 0 : 0;
    var total  = totalCol  ? Number(data[i][totalCol  - 1]) || 0 : 0;
    var status = statusCol ? String(data[i][statusCol - 1]).trim() : "";
    var title  = titleCol  ? String(data[i][titleCol  - 1]).trim() : "";

    var hasBid   = (b1 > 0 || b2 > 0 || b3 > 0);
    var hasBoost = (boost > 0);

    if (!hasBid && !hasBoost) continue;

    bidRowCount++;

    // competition bucket
    var bucket;
    if      (b1 >= EXTREME_THRESHOLD_) bucket = buckets.extreme;
    else if (b1 >= HIGH_THRESHOLD_)    bucket = buckets.high;
    else if (b1 >= MEDIUM_THRESHOLD_)  bucket = buckets.medium;
    else                               bucket = buckets.low;

    bucket.count++;
    if      (status === "Sent") bucket.sent++;
    else if (status === "Skip") bucket.skip++;
    else                        bucket.other++;

    // flag extreme + sent
    if (b1 >= EXTREME_THRESHOLD_ && status === "Sent") {
      extremeSentFlags.push({
        title:  title.substring(0, 44),
        b1:     b1,
        conn:   conn,
        boost:  boost,
        total:  total
      });
    }

    // boost tracking
    if (hasBoost) {
      var ratio = conn > 0 ? boost / conn : 0;
      var level = ratio >= BOOST_EXCESSIVE_RATIO_ ? "EXCESSIVE"
                : ratio >= BOOST_HEAVY_RATIO_      ? "HEAVY"
                : "";
      boostRows.push({
        title:  title.substring(0, 44),
        boost:  boost,
        conn:   conn,
        total:  total,
        ratio:  ratio,
        level:  level,
        status: status,
        b1:     b1
      });
    }
  }

  if (bidRowCount === 0) {
    ui.alert("No rows with bid data found. Enter Bid_1st values in Proposal_Generator to enable analysis.");
    return;
  }

  // ---- build report ----------------------------------------
  var totalSent    = buckets.low.sent + buckets.medium.sent + buckets.high.sent + buckets.extreme.sent;
  var totalSkip    = buckets.low.skip + buckets.medium.skip + buckets.high.skip + buckets.extreme.skip;
  var sendAccuracy = totalSent > 0
    ? Math.round(((buckets.low.sent + buckets.medium.sent) / totalSent) * 100) + "% of sends were low/medium competition"
    : "N/A";

  var competitionSection =
    "COMPETITION LEVELS (by Bid_1st)\n" +
    "  " + fmtBucketLine_(buckets.low)     + "\n" +
    "  " + fmtBucketLine_(buckets.medium)  + "\n" +
    "  " + fmtBucketLine_(buckets.high)    + "\n" +
    "  " + fmtBucketLine_(buckets.extreme) + "\n" +
    "Send accuracy: " + sendAccuracy + "\n";

  var flagSection = "";
  if (extremeSentFlags.length > 0) {
    flagSection =
      "\nEXTREME COMPETITION -- SENT (" + extremeSentFlags.length + " flags)\n" +
      "These jobs had 50+ proposals. Connects spent with near-zero win odds.\n";
    for (var f = 0; f < extremeSentFlags.length; f++) {
      var ef = extremeSentFlags[f];
      flagSection += "  Bid1=" + ef.b1 + " | " + ef.conn + " conn" +
        (ef.boost > 0 ? " +" + ef.boost + " boost" : "") +
        " | " + ef.title + "\n";
    }
  } else {
    flagSection = "\nEXTREME COMPETITION: No sends into extreme competition. Good discipline.\n";
  }

  var boostSection = "\nBOOST SUMMARY (" + boostRows.length + " jobs boosted)\n";
  var excessiveBoosts = [];
  var heavyBoosts     = [];
  var normalBoosts    = [];

  for (var b = 0; b < boostRows.length; b++) {
    var br = boostRows[b];
    var line = "  +" + br.boost + "/" + br.conn + " (" + Math.round(br.ratio * 100) + "%) " +
               "| " + br.status + " | B1=" + br.b1 + " | " + br.title;
    if (br.level === "EXCESSIVE") excessiveBoosts.push(line);
    else if (br.level === "HEAVY") heavyBoosts.push(line);
    else normalBoosts.push(line);
  }

  if (excessiveBoosts.length > 0) {
    boostSection += "EXCESSIVE (boost >= base connects):\n" + excessiveBoosts.join("\n") + "\n";
  }
  if (heavyBoosts.length > 0) {
    boostSection += "HEAVY (boost >= 75% of base):\n" + heavyBoosts.join("\n") + "\n";
  }
  if (normalBoosts.length > 0) {
    boostSection += "Normal:\n" + normalBoosts.join("\n") + "\n";
  }
  if (boostRows.length === 0) {
    boostSection += "No boosts recorded.\n";
  }

  var report =
    "BID ANALYSIS -- " + bidRowCount + " rows with bid data\n" +
    "Total: " + totalSent + " sent, " + totalSkip + " skipped\n" +
    "════════════════════════════════\n\n" +
    competitionSection +
    flagSection +
    boostSection;

  ui.alert("Bid Analysis", report, ui.ButtonSet.OK);
}


function fmtBucketLine_(b) {
  if (b.count === 0) return b.label + ": 0 jobs";
  var sendRate = Math.round((b.sent / b.count) * 100);
  return b.label + ": " + b.count + " jobs | " +
    "Sent=" + b.sent + " Skip=" + b.skip +
    (b.other > 0 ? " Other=" + b.other : "") +
    " | Send rate " + sendRate + "%";
}
