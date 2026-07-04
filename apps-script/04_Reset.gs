/**
 * ============================================================
 * 4. SMART RESET
 * Clears only true input columns. Preserves formulas / auto columns.
 * ============================================================
 */
function RESET_SYSTEM() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function clearByHeaders(sheetName, headerNames) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    var lastRow = sh.getLastRow();
    if (lastRow <= 1) return;
    var map = getHeaderMap_(sh);
    headerNames.forEach(function (name) {
      var col = getCol_(map, [name]);
      if (col) {
        sh.getRange(2, col, lastRow - 1, 1).clearContent();
      }
    });
  }

  // Every column in these three is transactional/derived (chat messages,
  // contract/milestone status) -- no identity column worth preserving the
  // way Proposal_Tracker's Job_Title/Client_Name is, so a full clear (not
  // clearByHeaders) is the right shape here.
  function clearAllRows(sheetName) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    var lastRow = sh.getLastRow();
    if (lastRow <= 1) return;
    sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
  }

  clearByHeaders("Job_Discovery", [
    "Job_Title", "Description", "Additional_Questions", "Client_Name", "Client Name",
    "Keyword_Search", "Experience_Level", "Hours_Since_Posted",
    "Days_Since_Posted", "Proposal_Count", "Payment_Verified",
    "Client_Hires", "Budget_Type", "Budget", "Hourly_Rate",
    "Job_Link", "Connects_Required", "AI_Fit_Notes"
  ]);

  // Job_Scoring is no longer a manual-entry sheet -- its raw-data columns
  // are a FILTER pull from Job_Discovery (see applyJobScoringPullFormula_ in
  // 00_Setup_Wizard.gs). Clearing them here would delete the FILTER formula
  // itself, not just its spilled values. Clearing Job_Discovery above already
  // empties Job_Scoring's pull automatically since the FILTER has nothing
  // left to match.

  clearByHeaders("Proposal_Generator", [
    "Notes", "Proposal_Status"
  ]);

  clearByHeaders("Proposal_Tracker", [
    "Viewed", "Interview", "Hired", "Revenue", "Notes"
  ]);

  clearAllRows("Client_Chat_Log");
  clearAllRows("Contract_Tracker");
  clearAllRows("Milestone_Tracker");
  clearAllRows("Hourly_Log");
}
