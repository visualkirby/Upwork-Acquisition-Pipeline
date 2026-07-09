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
 * The actual stage-counting happens in FFLib.computeFunnelStages (Library),
 * via the shared readFunnelStagesFromSheets_ (03_Helpers.gs) -- this stays
 * the same computation the Dashboard sheet's funnel section uses
 * (25_Dashboard.gs), so the two never disagree. This function just formats
 * the result as a text report.
 *
 * getWorkflowAnalysis_: AI-powered per-job breakdown (used by
 *   other functions; kept here for future wiring).
 * ============================================================
 */
function ANALYZE_JOB_WORKFLOW() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var missing = [];
  if (!ss.getSheetByName("Job_Discovery"))      missing.push("Job_Discovery");
  if (!ss.getSheetByName("Job_Scoring"))        missing.push("Job_Scoring");
  if (!ss.getSheetByName("Proposal_Generator")) missing.push("Proposal_Generator");
  if (!ss.getSheetByName("Proposal_Tracker"))   missing.push("Proposal_Tracker");
  if (missing.length > 0) {
    ui.alert("Sheets not found: " + missing.join(", "));
    return;
  }

  var funnel    = readFunnelStagesFromSheets_(ss);
  var discovery = funnel.discovery;
  var scoring   = funnel.scoring;
  var proposals = funnel.proposals;
  var outcomes  = funnel.outcomes;

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

  var unreviewedNote = discovery.reviewLater > 0
    ? "\n  Note: " + discovery.reviewLater + " Review Later jobs are an untapped pool not yet scored."
    : "";

  var discrepancyNote = "";
  if (proposals.sent > 0 && outcomes.total > 0 && Math.abs(proposals.sent - outcomes.total) > 2) {
    discrepancyNote =
      "\n  Note: Proposal_Generator shows " + proposals.sent + " Sent; " +
      "Proposal_Tracker has " + outcomes.total + " rows. " +
      "Difference of " + Math.abs(proposals.sent - outcomes.total) + " may be early manual entries.";
  }

  var report =
    "PIPELINE FUNNEL ANALYSIS\n" +
    "════════════════════════════════\n\n" +

    "STAGE 1 -- Discovery (" + discovery.total + " jobs logged)\n" +
    line("Move to Scoring:", discovery.toScoring,  discovery.total) + "\n" +
    line("Review Later:   ", discovery.reviewLater, discovery.total) + "\n" +
    (discovery.other > 0 ? line("Other:          ", discovery.other, discovery.total) + "\n" : "") +
    unreviewedNote + "\n\n" +

    "STAGE 2 -- Scoring (" + scoring.total + " jobs scored)\n" +
    line("APPLY:", scoring.apply, scoring.total) + "\n" +
    line("HOLD: ", scoring.hold,  scoring.total) + "\n" +
    line("SKIP: ", scoring.skip,  scoring.total) + "\n" +
    (scoring.other > 0 ? line("Other:", scoring.other, scoring.total) + "\n" : "") + "\n" +

    "STAGE 3 -- Proposals (" + proposals.total + " APPLY jobs in queue)\n" +
    line("Sent:          ", proposals.sent,  proposals.total) + "\n" +
    line("Skip:          ", proposals.skip,  proposals.total) + "\n" +
    line("Ready (unsent):", proposals.ready, proposals.total) + "\n" +
    (proposals.other > 0 ? line("Other:         ", proposals.other, proposals.total) + "\n" : "") +
    discrepancyNote + "\n\n" +

    "STAGE 4 -- Outcomes (" + outcomes.total + " proposals tracked)\n" +
    line("Hired (Y):      ", outcomes.hired,     outcomes.total) + "\n" +
    line("Interview (Y):  ", outcomes.interview, outcomes.total) + "\n" +
    line("Viewed (Y):     ", outcomes.viewed,    outcomes.total) + "\n" +
    line("Not viewed (N): ", outcomes.notViewed, outcomes.total) + "\n\n" +

    "END-TO-END CONVERSION\n" +
    "  Discovery -> Scoring:    " + pct(discovery.toScoring, discovery.total) + "  (" + discovery.toScoring + " / " + discovery.total + ")\n" +
    "  Scoring -> APPLY:        " + pct(scoring.apply, scoring.total)         + "  (" + scoring.apply       + " / " + scoring.total   + ")\n" +
    "  APPLY -> Sent:           " + pct(proposals.sent, scoring.apply)        + "  (" + proposals.sent      + " / " + scoring.apply   + ")\n" +
    "  Sent -> Hired:           " + pct(outcomes.hired, outcomes.total)       + "  (" + outcomes.hired      + " / " + outcomes.total  + ")\n" +
    "  Overall (logged -> hire): " + pct(outcomes.hired, discovery.total)     + "  (" + outcomes.hired      + " / " + discovery.total + ")";

  ui.alert("Pipeline Funnel", report, ui.ButtonSet.OK);
}


// ---- per-job AI analysis (kept for future wiring) -----------
// Moved to the Apps Script Library: FFLib.getWorkflowAnalysis(jobTitle, description, apiKey).
