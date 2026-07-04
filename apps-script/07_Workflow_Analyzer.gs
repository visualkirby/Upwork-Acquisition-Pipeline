/**
 * ============================================================
 * 7. WORKFLOW ANALYZER
 *
 * ANALYZE_JOB_WORKFLOW: four-stage pipeline funnel report.
 *   Stage 1 -- Job_Discovery: Discovery_Action distribution
 *   Stage 2 -- Job_Scoring: Final_Decision distribution
 *   Stage 3 -- Proposal_Generator: Proposal_Status distribution
 *   Stage 4 -- Proposal_Tracker: Hired / Interview / Viewed outcomes
 *
 * getWorkflowAnalysis_: AI-powered per-job breakdown (used by
 *   other functions; kept here for future wiring).
 * ============================================================
 */
function ANALYZE_JOB_WORKFLOW() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var discSheet    = ss.getSheetByName("Job_Discovery");
  var scoringSheet = ss.getSheetByName("Job_Scoring");
  var pgSheet      = ss.getSheetByName("Proposal_Generator");
  var ptSheet      = ss.getSheetByName("Proposal_Tracker");

  var missing = [];
  if (!discSheet)    missing.push("Job_Discovery");
  if (!scoringSheet) missing.push("Job_Scoring");
  if (!pgSheet)      missing.push("Proposal_Generator");
  if (!ptSheet)      missing.push("Proposal_Tracker");
  if (missing.length > 0) {
    ui.alert("Sheets not found: " + missing.join(", "));
    return;
  }

  // ---- Stage 1: Job_Discovery --------------------------------
  var discMap       = getHeaderMap_(discSheet);
  var discActionCol = getCol_(discMap, ["Discovery_Action"]);

  var discTotal     = 0;
  var discToScoring = 0;
  var discReview    = 0;
  var discOther     = 0;

  if (discActionCol && discSheet.getLastRow() > 1) {
    var discData = discSheet
      .getRange(2, discActionCol, discSheet.getLastRow() - 1, 1)
      .getValues();
    for (var i = 0; i < discData.length; i++) {
      var action = String(discData[i][0]).trim();
      if (action === "") continue;
      discTotal++;
      if (action === "Move to Scoring") discToScoring++;
      else if (action === "Review Later") discReview++;
      else discOther++;
    }
  }

  // ---- Stage 2: Job_Scoring ----------------------------------
  var jsMap       = getHeaderMap_(scoringSheet);
  var jsDecCol    = getCol_(jsMap, ["Final_Decision"]);

  var jsTotal  = 0;
  var jsApply  = 0;
  var jsHold   = 0;
  var jsSkip   = 0;
  var jsOther  = 0;

  if (jsDecCol && scoringSheet.getLastRow() > 1) {
    var jsData = scoringSheet
      .getRange(2, jsDecCol, scoringSheet.getLastRow() - 1, 1)
      .getValues();
    for (var i = 0; i < jsData.length; i++) {
      var dec = String(jsData[i][0]).trim();
      if (dec === "") continue;
      jsTotal++;
      if      (dec === "APPLY") jsApply++;
      else if (dec === "HOLD")  jsHold++;
      else if (dec === "SKIP")  jsSkip++;
      else                      jsOther++;
    }
  }

  // ---- Stage 3: Proposal_Generator ---------------------------
  var pgMap       = getHeaderMap_(pgSheet);
  var pgStatusCol = getCol_(pgMap, ["Proposal_Status"]);

  var pgTotal = 0;
  var pgSent  = 0;
  var pgSkip  = 0;
  var pgReady = 0;
  var pgOther = 0;

  if (pgStatusCol && pgSheet.getLastRow() > 1) {
    var pgData = pgSheet
      .getRange(2, pgStatusCol, pgSheet.getLastRow() - 1, 1)
      .getValues();
    for (var i = 0; i < pgData.length; i++) {
      var pgStatus = String(pgData[i][0]).trim();
      if (pgStatus === "") continue;
      pgTotal++;
      if      (pgStatus === "Sent")  pgSent++;
      else if (pgStatus === "Skip")  pgSkip++;
      else if (pgStatus === "Ready") pgReady++;
      else                           pgOther++;
    }
  }

  // ---- Stage 4: Proposal_Tracker -----------------------------
  var ptMap      = getHeaderMap_(ptSheet);
  var ptHiredCol = getCol_(ptMap, ["Hired"]);
  var ptReplyCol = getCol_(ptMap, ["Viewed"]);
  var ptIntCol   = getCol_(ptMap, ["Interview"]);

  var ptTotal    = 0;
  var ptHiredY   = 0;
  var ptReplyY   = 0;
  var ptIntY     = 0;

  if (ptSheet.getLastRow() > 1) {
    var ptLastCol = ptSheet.getLastColumn();
    var ptData    = ptSheet
      .getRange(2, 1, ptSheet.getLastRow() - 1, ptLastCol)
      .getValues();
    for (var i = 0; i < ptData.length; i++) {
      var hiredVal = ptHiredCol ? String(ptData[i][ptHiredCol - 1]).trim() : "";
      if (hiredVal === "" && !ptHiredCol) continue;
      ptTotal++;
      if (ptHiredCol && hiredVal === "Y") ptHiredY++;
      if (ptReplyCol && String(ptData[i][ptReplyCol - 1]).trim() === "Y") ptReplyY++;
      if (ptIntCol   && String(ptData[i][ptIntCol   - 1]).trim() === "Y") ptIntY++;
    }
  }

  // ---- conversion rates --------------------------------------
  function pct(num, den) {
    if (!den || den === 0) return "N/A";
    return Math.round((num / den) * 1000) / 10 + "%";
  }

  function line(label, count, total) {
    var p    = total > 0 ? pct(count, total) : "--";
    var pad  = "                    ".substring(label.length);
    var cPad = "     ".substring(String(count).length);
    return "  " + label + pad + count + cPad + "(" + p + ")";
  }

  var unreviewedNote = discReview > 0
    ? "\n  Note: " + discReview + " Review Later jobs are an untapped pool not yet scored."
    : "";

  var discrepancyNote = "";
  if (pgSent > 0 && ptTotal > 0 && Math.abs(pgSent - ptTotal) > 2) {
    discrepancyNote =
      "\n  Note: Proposal_Generator shows " + pgSent + " Sent; " +
      "Proposal_Tracker has " + ptTotal + " rows. " +
      "Difference of " + Math.abs(pgSent - ptTotal) + " may be early manual entries.";
  }

  var report =
    "PIPELINE FUNNEL ANALYSIS\n" +
    "════════════════════════════════\n\n" +

    "STAGE 1 -- Discovery (" + discTotal + " jobs logged)\n" +
    line("Move to Scoring:", discToScoring, discTotal) + "\n" +
    line("Review Later:   ", discReview,    discTotal) + "\n" +
    (discOther > 0 ? line("Other:          ", discOther, discTotal) + "\n" : "") +
    unreviewedNote + "\n\n" +

    "STAGE 2 -- Scoring (" + jsTotal + " jobs scored)\n" +
    line("APPLY:", jsApply, jsTotal) + "\n" +
    line("HOLD: ", jsHold,  jsTotal) + "\n" +
    line("SKIP: ", jsSkip,  jsTotal) + "\n" +
    (jsOther > 0 ? line("Other:", jsOther, jsTotal) + "\n" : "") + "\n" +

    "STAGE 3 -- Proposals (" + pgTotal + " APPLY jobs in queue)\n" +
    line("Sent:          ", pgSent,  pgTotal) + "\n" +
    line("Skip:          ", pgSkip,  pgTotal) + "\n" +
    line("Ready (unsent):", pgReady, pgTotal) + "\n" +
    (pgOther > 0 ? line("Other:         ", pgOther, pgTotal) + "\n" : "") +
    discrepancyNote + "\n\n" +

    "STAGE 4 -- Outcomes (" + ptTotal + " proposals tracked)\n" +
    line("Hired (Y):      ", ptHiredY, ptTotal) + "\n" +
    line("Interview (Y):  ", ptIntY,   ptTotal) + "\n" +
    line("Viewed (Y):     ", ptReplyY, ptTotal) + "\n" +
    line("Not viewed (N): ", ptTotal - ptReplyY, ptTotal) + "\n\n" +

    "END-TO-END CONVERSION\n" +
    "  Discovery -> Scoring:    " + pct(discToScoring, discTotal)  + "  (" + discToScoring + " / " + discTotal  + ")\n" +
    "  Scoring -> APPLY:        " + pct(jsApply, jsTotal)          + "  (" + jsApply       + " / " + jsTotal    + ")\n" +
    "  APPLY -> Sent:           " + pct(pgSent, jsApply)           + "  (" + pgSent         + " / " + jsApply   + ")\n" +
    "  Sent -> Hired:           " + pct(ptHiredY, ptTotal)         + "  (" + ptHiredY        + " / " + ptTotal   + ")\n" +
    "  Overall (logged -> hire): " + pct(ptHiredY, discTotal)      + "  (" + ptHiredY        + " / " + discTotal + ")";

  ui.alert("Pipeline Funnel", report, ui.ButtonSet.OK);
}


// ---- per-job AI analysis (kept for future wiring) -----------
// Moved to the Apps Script Library: FFLib.getWorkflowAnalysis(jobTitle, description, apiKey).
