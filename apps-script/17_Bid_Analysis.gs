/**
 * ============================================================
 * 17. BID ANALYSIS (S025)
 * Reads Proposal_Generator and produces two reports:
 *   1. Competition level breakdown (by Bid_1st proposal count)
 *      with send rates and flagged "sent into extreme competition" rows.
 *   2. Boost pattern report -- flags excessive boosts and
 *      boosts applied to already-skipped jobs (wasted planning).
 *
 * The bucketing/flagging algorithm and its 5 tuning thresholds
 * moved to the Apps Script Library (FFLib.bucketBidPatterns,
 * FFLib.fmtBucketLine). This file is the sheet-read + ui.alert
 * wrapper.
 * ============================================================
 */
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

  var rows = data.map(function (d) {
    return {
      title:              titleCol  ? d[titleCol  - 1] : "",
      bid1:               bid1Col   ? d[bid1Col   - 1] : 0,
      bid2:               bid2Col   ? d[bid2Col   - 1] : 0,
      bid3:               bid3Col   ? d[bid3Col   - 1] : 0,
      boost:              boostCol  ? d[boostCol  - 1] : 0,
      connectsRequired:   connCol   ? d[connCol   - 1] : 0,
      totalConnectsSpent: totalCol  ? d[totalCol  - 1] : 0,
      status:             statusCol ? d[statusCol - 1] : ""
    };
  });

  var result = FFLib.bucketBidPatterns(rows);

  if (result.bidRowCount === 0) {
    ui.alert("No rows with bid data found. Enter Bid_1st values in Proposal_Generator to enable analysis.");
    return;
  }

  var buckets = result.buckets;
  var competitionSection =
    "COMPETITION LEVELS (by Bid_1st)\n" +
    "  " + FFLib.fmtBucketLine(buckets.low)     + "\n" +
    "  " + FFLib.fmtBucketLine(buckets.medium)  + "\n" +
    "  " + FFLib.fmtBucketLine(buckets.high)    + "\n" +
    "  " + FFLib.fmtBucketLine(buckets.extreme) + "\n" +
    "Send accuracy: " + result.sendAccuracy + "\n";

  var flagSection = "";
  if (result.extremeSentFlags.length > 0) {
    flagSection =
      "\nEXTREME COMPETITION -- SENT (" + result.extremeSentFlags.length + " flags)\n" +
      "These jobs had 50+ proposals. Connects spent with near-zero win odds.\n";
    result.extremeSentFlags.forEach(function (ef) {
      flagSection += "  Bid1=" + ef.b1 + " | " + ef.conn + " conn" +
        (ef.boost > 0 ? " +" + ef.boost + " boost" : "") +
        " | " + ef.title + "\n";
    });
  } else {
    flagSection = "\nEXTREME COMPETITION: No sends into extreme competition. Good discipline.\n";
  }

  var boostSection = "\nBOOST SUMMARY (" + result.boostRows.length + " jobs boosted)\n";
  var excessiveBoosts = [];
  var heavyBoosts     = [];
  var normalBoosts    = [];

  result.boostRows.forEach(function (br) {
    var line = "  +" + br.boost + "/" + br.conn + " (" + Math.round(br.ratio * 100) + "%) " +
               "| " + br.status + " | B1=" + br.b1 + " | " + br.title;
    if (br.level === "EXCESSIVE") excessiveBoosts.push(line);
    else if (br.level === "HEAVY") heavyBoosts.push(line);
    else normalBoosts.push(line);
  });

  if (excessiveBoosts.length > 0) {
    boostSection += "EXCESSIVE (boost >= base connects):\n" + excessiveBoosts.join("\n") + "\n";
  }
  if (heavyBoosts.length > 0) {
    boostSection += "HEAVY (boost >= 75% of base):\n" + heavyBoosts.join("\n") + "\n";
  }
  if (normalBoosts.length > 0) {
    boostSection += "Normal:\n" + normalBoosts.join("\n") + "\n";
  }
  if (result.boostRows.length === 0) {
    boostSection += "No boosts recorded.\n";
  }

  var report =
    "BID ANALYSIS -- " + result.bidRowCount + " rows with bid data\n" +
    "Total: " + result.totalSent + " sent, " + result.totalSkip + " skipped\n" +
    "════════════════════════════════\n\n" +
    competitionSection +
    flagSection +
    boostSection;

  ui.alert("Bid Analysis", report, ui.ButtonSet.OK);
}
